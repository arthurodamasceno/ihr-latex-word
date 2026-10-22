/* Painel do suplemento: liga o editor LaTeX ao documento aberto no Word
 *
 * Modelo de uso:
 *  - "Modo inserção": o cursor está num ponto vazio; Ctrl+Enter insere o LaTeX ali.
 *  - "Vinculado": um trecho do documento está ligado ao editor (marcador oculto _LtxEdit).
 *    Ctrl+Enter atualiza esse mesmo trecho. Após inserir/atualizar, o trecho continua vinculado.
 *  - Leitura automática: ao selecionar um trecho no Word, ele é convertido e vinculado,
 *    a menos que haja edições não aplicadas no editor.
 */
(function () {
  'use strict';
  var L = window.IHRLX.latex2ooxml;
  var BIB = window.IHRLX.bib;
  var O = window.IHRLX.ooxml2latex;
  var BM = '_LtxEdit';
  var VERSION = window.IHRLX_VERSION || 'dev';

  var S = {
    images: {},
    lib: null,
    bound: false,        // há um trecho vinculado
    paras: 0,
    clean: '',           // texto do editor que corresponde ao trecho (sem edições pendentes)
    base: '',            // LaTeX do trecho no documento no momento do vínculo (detecta mudanças no Word)
    unsafe: [],
    busy: false,
    auto: true,
    suppressUntil: 0,
    lastRefresh: 0,
    selTimer: null,
    selPending: false
  };

  var $ = function (id) { return document.getElementById(id); };
  var ed = $('editor');
  function norm(s) { return (s || '').replace(/\s+/g, ' ').trim(); }
  function dirty() { return S.bound && norm(ed.value) !== norm(S.clean); }

  // ------------------------------------------------------------ interface
  function status(msg, kind, list) {
    var el = $('status');
    el.className = 'status ' + (kind || '');
    el.textContent = msg || '';
    if (list && list.length) {
      var ul = document.createElement('ul');
      list.forEach(function (w) { var li = document.createElement('li'); li.textContent = w; ul.appendChild(li); });
      el.appendChild(ul);
    }
  }
  function notice(text, actions) {
    var el = $('notice');
    el.innerHTML = '';
    if (!text) { el.classList.add('hidden'); return; }
    el.classList.remove('hidden');
    var p = document.createElement('div'); p.textContent = text; el.appendChild(p);
    (actions || []).forEach(function (a) {
      if (a.checkbox) {
        var lab = document.createElement('label');
        var cb = document.createElement('input'); cb.type = 'checkbox'; cb.id = a.id;
        lab.appendChild(cb); lab.appendChild(document.createTextNode(' ' + a.label));
        el.appendChild(lab);
        return;
      }
      var b = document.createElement('button'); b.textContent = a.label; b.className = 'small' + (a.primary ? ' primary' : '');
      b.onclick = a.run; el.appendChild(b);
    });
  }
  function renderMode() {
    var bar = $('modeBar');
    bar.className = 'modebar ' + (S.bound ? 'bound' : 'insert');
    $('modeText').textContent = S.bound
      ? 'Vinculado a ' + S.paras + ' parágrafo(s) do documento' + (dirty() ? ' · edições não aplicadas' : '')
      : 'Modo inserção: o conteúdo entra no cursor';
    $('btnUnbind').classList.toggle('hidden', !S.bound);
    $('btnPrimary').textContent = S.bound ? 'Atualizar trecho' : 'Inserir no cursor';
    $('btnPrimary').title = (S.bound ? 'Substitui o trecho vinculado pelo LaTeX do editor' : 'Insere o LaTeX no ponto do cursor') + ' (Ctrl+Enter)';
  }
  function setBusy(on) {
    S.busy = on;
    $('btnPrimary').disabled = on; $('btnRead').disabled = on;
    if (!on && S.selPending) { S.selPending = false; scheduleSelCheck(); }
  }
  function errMsg(e) {
    var m = (e && (e.message || e.toString())) || 'erro desconhecido';
    if (e && e.debugInfo && e.debugInfo.errorLocation) m += ' (' + e.debugInfo.errorLocation + ')';
    return m;
  }
  function api(v) { return typeof Office !== 'undefined' && Office.context && Office.context.requirements && Office.context.requirements.isSetSupported('WordApi', v); }
  function needWord() {
    if (typeof Word === 'undefined') { status('Abra este painel dentro do Word para usar os botões.', 'err'); return true; }
    return false;
  }

  // --------------------------------------------------------- rascunho/opções
  try {
    var d = localStorage.getItem('ihrlx.draft'); if (d) ed.value = d;
    var a = localStorage.getItem('ihrlx.auto'); if (a !== null) S.auto = a === '1';
  } catch (e) { /* sem armazenamento */ }
  $('autoRead').checked = S.auto;
  $('autoRead').addEventListener('change', function () {
    S.auto = this.checked;
    try { localStorage.setItem('ihrlx.auto', S.auto ? '1' : '0'); } catch (e) { /* ignora */ }
    if (S.auto) scheduleSelCheck();
  });
  var saveT = null;
  ed.addEventListener('input', function () {
    renderMode();
    clearTimeout(saveT);
    saveT = setTimeout(function () { try { localStorage.setItem('ihrlx.draft', ed.value); } catch (e) { /* ignora */ } }, 400);
  });
  ed.addEventListener('keydown', function (e) {
    if (e.key === 'Tab' && !e.shiftKey) {
      e.preventDefault();
      ed.setRangeText('  ', ed.selectionStart, ed.selectionEnd, 'end');
    } else if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
      e.preventDefault(); primary();
    }
  });
  function setEditor(text) {
    ed.value = text;
    try { localStorage.setItem('ihrlx.draft', text); } catch (e) { /* ignora */ }
    renderMode();
  }

  // ---------------------------------------------------------- imagens
  function renderImages() {
    var ul = $('imgList'); ul.innerHTML = '';
    var names = Object.keys(S.images);
    $('imgCount').textContent = names.length;
    names.forEach(function (n) {
      var im = S.images[n];
      var li = document.createElement('li');
      var th = document.createElement('img'); th.src = 'data:' + im.contentType + ';base64,' + im.base64;
      var nm = document.createElement('span'); nm.className = 'nm'; nm.textContent = n; nm.title = n + ' (' + im.width + '×' + im.height + ' px)';
      var bt = document.createElement('button'); bt.textContent = '×'; bt.title = 'Remover';
      bt.onclick = function () { delete S.images[n]; renderImages(); };
      li.appendChild(th); li.appendChild(nm); li.appendChild(bt); ul.appendChild(li);
    });
  }
  $('imgInput').addEventListener('change', function (ev) {
    Array.prototype.slice.call(ev.target.files || []).forEach(function (f) {
      var rd = new FileReader();
      rd.onload = function () {
        var url = rd.result, b64 = url.slice(url.indexOf(',') + 1);
        var img = new Image();
        img.onload = function () {
          S.images[f.name] = { base64: b64, contentType: f.type || 'image/png', width: img.naturalWidth, height: img.naturalHeight };
          renderImages();
        };
        img.src = url;
      };
      rd.readAsDataURL(f);
    });
    ev.target.value = '';
    $('imgBox').open = true;
  });

  // --------------------------------------------------------- modelos
  var Y = new Date().getFullYear();
  var SNIP = {
    secao: '\\chapter{Título 1}\n\nTexto do capítulo.\n\n\\section{Título 2}\n\nTexto da seção.\n\n\\subsection{Título 3}\n\nTexto da subseção.\n',
    figura: '\\begin{figure}\n  \\caption{Título da figura}\\label{fig:nome}\n  \\includegraphics[width=12cm]{arquivo.png}\n  \\fonte{IHR, ' + Y + '.}\n\\end{figure}\n',
    tabela: '\\begin{table}\n  \\caption{Título da tabela}\\label{tab:nome}\n  \\begin{tabular}{Nc}\n    \\cab{Nome} & \\cab{Valor} \\\\\n    Item A & 10 \\\\\n    Item B & 20 \\\\\n  \\end{tabular}\n  \\fonte{IHR, ' + Y + '.}\n\\end{table}\n',
    quadro: '\\begin{quadro}\n  \\caption{Título do quadro}\\label{qua:nome}\n  \\begin{tabular}{|c|c|c|}\n    \\hline\n    Coluna 1 & Coluna 2 & Coluna 3 \\\\ \\hline\n    a & b & c \\\\ \\hline\n  \\end{tabular}\n  \\fonte{IHR, ' + Y + '.}\n\\end{quadro}\n',
    equacao: 'Conforme \\eqref{eq:nome}:\n\\begin{equation}\n  y = a x^2 + b x + c\n  \\label{eq:nome}\n\\end{equation}\n',
    listas: '\\begin{enumerate}\n  \\item Primeiro passo.\n  \\item Segundo passo.\n\\end{enumerate}\n\n\\begin{itemize}\n  \\item Um ponto.\n  \\item Outro ponto.\n\\end{itemize}\n',
    texto: 'Texto com \\textbf{negrito}, \\textit{itálico}, \\underline{sublinhado}, H\\textsubscript{2}O, m\\textsuperscript{2}, uma fórmula $\\alpha = \\frac{a}{b}$ e uma nota\\footnote{Texto da nota.}.\n',
    sumario: '\\tableofcontents\n'
  };
  $('snippets').addEventListener('change', function (e) {
    var k = e.target.value; e.target.value = '';
    if (!SNIP[k]) return;
    var s = ed.selectionStart, before = ed.value.slice(0, s);
    var pre = !before || /\n\n$/.test(before) ? '' : (/\n$/.test(before) ? '\n' : '\n\n');
    ed.setRangeText(pre + SNIP[k], s, ed.selectionEnd, 'end');
    ed.focus(); ed.dispatchEvent(new Event('input'));
  });

  // ------------------------------------------------------- Word: apoio
  async function labelKinds(ctx) {
    var map = {};
    if (!api('1.4')) return map;
    try {
      var names = ctx.document.body.getRange('Whole').getBookmarks(true, true);
      await ctx.sync();
      names.value.forEach(function (n) { var d = O.decLabel(n); if (d.kind) map[d.label] = d.kind; });
    } catch (e) { /* sem marcadores */ }
    return map;
  }
  async function updateFields(ctx, res) {
    if (!api('1.5')) return false;
    try {
      var fields = ctx.document.body.fields;
      fields.load('items/code');
      await ctx.sync();
      var seq = [], ref = [], toc = [];
      fields.items.forEach(function (f) {
        var c = (f.code || '').trim();
        if (/^SEQ\b/i.test(c)) seq.push(f);
        else if (/^(REF|PAGEREF)\b/i.test(c)) ref.push(f);
        else if (/^TOC\b/i.test(c)) toc.push(f);
      });
      seq.forEach(function (f) { f.updateResult(); });
      await ctx.sync();
      ref.forEach(function (f) { f.updateResult(); });
      if (res && res.blocks && res.blocks.some(function (b) { return b.type === 'toc'; })) toc.forEach(function (f) { f.updateResult(); });
      await ctx.sync();
      return true;
    } catch (e) { return false; }
  }
  async function readRange(ctx, range) {
    var ox = range.getOoxml();
    await ctx.sync();
    return O.convert(ox.value);
  }
  async function boundRange(ctx) {
    if (!api('1.4')) return null;
    var b = ctx.document.getBookmarkRangeOrNullObject(BM);
    b.load('isNullObject');
    await ctx.sync();
    return b.isNullObject ? null : b;
  }
  // vincula o intervalo, seleciona-o e registra o LaTeX atual dele como "base"
  async function bindTo(ctx, range, editorText) {
    range.insertBookmark(BM);
    var ps = range.paragraphs; ps.load('items');
    range.select();
    S.suppressUntil = Date.now() + 1500;
    await ctx.sync();
    var r = await readRange(ctx, range);
    await fillImages(ctx, range, r);
    Object.keys(r.images).forEach(function (k) { if (!S.images[k]) S.images[k] = r.images[k]; });
    S.bound = true;
    S.paras = ps.items.length;
    S.base = r.latex;
    S.clean = editorText !== undefined ? editorText : r.latex;
    S.unsafe = r.unsafe;
    return r;
  }
  function unbindState(clear) {
    S.bound = false; S.paras = 0; S.base = ''; S.unsafe = [];
    if (clear) { S.clean = ''; setEditor(''); }
    notice('');
    renderMode();
  }
  function convertEditor(kinds, cited) {
    var res = L.convert(ed.value, { images: S.images, labelKinds: kinds, bib: S.lib, citedKeys: cited || [] });
    return { res: res, pkg: L.buildPackage(res) };
  }

  // ------------------------------------------------------- referências
  var BIBNS = 'urn:ihr-latex-word:bib';
  function oc(fn) { return new Promise(function (ok) { fn(function (r) { ok(r); }); }); }
  async function loadLibrary() {
    S.lib = new BIB.Library('');
    try {
      var r = await oc(function (cb) { Office.context.document.customXmlParts.getByNamespaceAsync(BIBNS, cb); });
      if (r.status === 'succeeded' && r.value.length) {
        var x = await oc(function (cb) { r.value[0].getXmlAsync(cb); });
        if (x.status === 'succeeded') S.lib = new BIB.Library(window.IHRLX.xml.text(window.IHRLX.xml.parse(x.value)));
      }
    } catch (e) { /* sem partes XML: biblioteca vazia */ }
    renderBib();
  }
  async function saveLibrary() {
    var xml = '<ihrbib xmlns="' + BIBNS + '">' + window.IHRLX.xml.esc(S.lib.serialize()) + '</ihrbib>';
    var r = await oc(function (cb) { Office.context.document.customXmlParts.getByNamespaceAsync(BIBNS, cb); });
    if (r.status === 'succeeded') for (var i = 0; i < r.value.length; i++) await oc(function (cb) { r.value[i].deleteAsync(cb); });
    var a = await oc(function (cb) { Office.context.document.customXmlParts.addAsync(xml, cb); });
    if (a.status !== 'succeeded') throw new Error('não foi possível guardar o .bib no documento');
  }
  function renderBib() {
    var n = S.lib ? S.lib.count() : 0;
    $('bibCount').textContent = n;
    var q = BIB.sortKey($('bibFilter').value || '');
    var ul = $('bibList'); ul.innerHTML = '';
    if (!n) return;
    var keys = S.lib.order.filter(function (k) { return !q || BIB.sortKey(k + ' ' + S.lib.label(k)).indexOf(q) >= 0; }).slice(0, 60);
    keys.forEach(function (k) {
      var li = document.createElement('li');
      li.innerHTML = '<code></code><span></span>';
      li.firstChild.textContent = k; li.lastChild.textContent = S.lib.label(k);
      li.title = 'Clique para inserir \\cite{' + k + '} no editor';
      li.onclick = function () {
        ed.setRangeText('\\cite{' + k + '}', ed.selectionStart, ed.selectionEnd, 'end');
        ed.focus(); ed.dispatchEvent(new Event('input'));
      };
      ul.appendChild(li);
    });
  }
  async function citedKeys(ctx) {
    var keys = [], seen = {};
    if (!api('1.5')) return keys;
    try {
      var fields = ctx.document.body.fields; fields.load('items/code');
      await ctx.sync();
      fields.items.forEach(function (f) {
        var m = /^\s*ADDIN\s+IHRCITE\s+(\{.*\})\s*$/i.exec(f.code || '');
        if (!m) return;
        try { JSON.parse(m[1]).k.forEach(function (k) { if (!seen[k]) { seen[k] = 1; keys.push(k); } }); } catch (e) { /* ignora */ }
      });
    } catch (e) { /* sem campos */ }
    return keys;
  }
  // Regenera a lista REFERÊNCIAS: entradas do .bib citadas no documento + entradas manuais (de quem não usa o
  // suplemento), em ordem alfabética. Entradas manuais nunca são apagadas.
  async function regenBib(ctx, createIfMissing) {
    if (!S.lib || !S.lib.count()) return 'Nenhum .bib carregado.';
    var keys = (await citedKeys(ctx)).filter(function (k) { return S.lib.has(k); });
    var paras = ctx.document.body.paragraphs; paras.load('items/text,styleBuiltIn');
    await ctx.sync();
    var items = paras.items, h = -1;
    for (var i = items.length - 1; i >= 0; i--) {
      if (items[i].styleBuiltIn === 'Heading1' && /^REFER[EÊ]NCIAS$/i.test((items[i].text || '').trim())) { h = i; break; }
    }
    if (h < 0) {
      if (!createIfMissing || !keys.length) return keys.length ? 'Não encontrei o capítulo "Referências" no documento.' : 'Nenhuma citação do .bib no documento.';
      var res0 = L.convertBibList({ bib: S.lib, keys: keys, manual: [], withHeading: true });
      ctx.document.body.insertOoxml(L.buildPackage(res0), 'End');
      await ctx.sync();
      return 'Capítulo "Referências" criado no fim do documento com ' + keys.length + ' referência(s).';
    }
    var list = [];
    for (var j = h + 1; j < items.length && items[j].styleBuiltIn === 'Bibliography'; j++) list.push(items[j]);
    var manual = [];
    if (list.length) {
      var bms = list.map(function (p) { return p.getRange('Whole').getBookmarks(true, true); });
      await ctx.sync();
      var oxs = [];
      list.forEach(function (p, k) {
        var managed = bms[k].value.some(function (n) { return /^_RefB_/.test(n); });
        if (!managed && (p.text || '').trim()) oxs.push(p.getRange('Whole').getOoxml());
      });
      if (oxs.length) await ctx.sync();
      oxs.forEach(function (ox) {
        var t = O.convert(ox.value).latex.trim();
        var m = /^\\refmanual\{([\s\S]*)\}$/.exec(t);
        manual.push(m ? m[1] : t);
      });
    }
    var res = L.convertBibList({ bib: S.lib, keys: keys, manual: manual });
    var pkg = L.buildPackage(res);
    if (list.length) list[0].getRange('Whole').expandTo(list[list.length - 1].getRange('Whole')).insertOoxml(pkg, 'Replace');
    else if (res.body) items[h].insertParagraph('', 'After').insertOoxml(pkg, 'Replace');
    await ctx.sync();
    return 'Referências atualizadas: ' + keys.length + ' do .bib e ' + manual.length + ' manual(is).';
  }
  function finish(res, okMsg, fieldsOk) {
    var w = res.warnings.slice();
    if (!fieldsOk && res.fields) w.push('Não consegui atualizar os campos automaticamente: selecione tudo (Ctrl+A) e pressione F9 para atualizar números e referências.');
    status(okMsg, w.length ? 'warn' : 'ok', w);
  }

  // ------------------------------------------------------- ações
  function primary() { return S.bound ? updateBound(false) : insertAtCursor(); }

  async function insertAtCursor() {
    if (needWord()) return;
    if (!ed.value.trim()) { status('O editor está vazio.', 'warn'); return; }
    setBusy(true); status('Convertendo…');
    try {
      await Word.run(async function (ctx) {
        var kinds = await labelKinds(ctx);
        var c = convertEditor(kinds, await citedKeys(ctx));
        var rng = ctx.document.getSelection().insertOoxml(c.pkg, 'Replace');
        await ctx.sync();
        var ok = c.res.fields ? await updateFields(ctx, c.res) : true;
        var bibMsg = c.res.cites && !c.res.hasBib && S.lib && S.lib.count() ? await regenBib(ctx, true) : '';
        if (api('1.4')) await bindTo(ctx, rng, ed.value);
        notice(''); renderMode();
        finish(c.res, (S.bound ? 'Inserido. O trecho continua vinculado: Ctrl+Enter atualiza ele.' : 'Inserido no documento.') + (bibMsg ? ' ' + bibMsg : ''), ok);
      });
    } catch (e) {
      status('Não foi possível inserir: ' + errMsg(e), 'err');
    } finally { setBusy(false); }
  }

  async function updateBound(force) {
    if (needWord()) return;
    if (S.unsafe.length && !force && !($('unsafeOk') && $('unsafeOk').checked)) {
      notice('O trecho vinculado contém ' + S.unsafe.join(', ') + ', que não têm equivalente em LaTeX e serão perdidos ao atualizar.',
        [{ checkbox: true, id: 'unsafeOk', label: 'Atualizar mesmo assim' }]);
      status('Confirme no aviso acima para continuar.', 'warn');
      return;
    }
    setBusy(true); status('Atualizando…');
    try {
      await Word.run(async function (ctx) {
        var b = await boundRange(ctx);
        if (!b) { unbindState(false); status('O trecho vinculado não existe mais no documento. Selecione de novo ou use "Inserir no cursor".', 'err'); return; }
        // o trecho mudou no Word desde a leitura? (você ou um colega editou)
        var cur = await readRange(ctx, b);
        if (!force && norm(cur.latex) !== norm(S.base)) {
          notice('Este trecho foi alterado no Word depois que foi lido no painel (por você ou por outra pessoa). Atualizar agora apagaria essas alterações.', [
            { label: 'Recarregar do Word', primary: true, run: function () { reloadBound(); } },
            { label: 'Sobrescrever', run: function () { notice(''); updateBound(true); } }
          ]);
          status('');
          return;
        }
        var kinds = await labelKinds(ctx);
        var c = convertEditor(kinds, await citedKeys(ctx));
        if (c.res.missingImages.length && !force) {
          notice('Faltam no painel as imagens ' + c.res.missingImages.join(', ') + '. Atualizar agora trocaria essas imagens do documento por marcadores de texto. Carregue-as em "Imagens" ou confirme.', [
            { label: 'Atualizar mesmo assim', run: function () { notice(''); updateBound(true); } }
          ]);
          $('imgBox').open = true; status('');
          return;
        }
        var rng = b.insertOoxml(c.pkg, 'Replace');
        await ctx.sync();
        var ok = c.res.fields ? await updateFields(ctx, c.res) : true;
        var bibMsg = c.res.cites && !c.res.hasBib && S.lib && S.lib.count() ? await regenBib(ctx, true) : '';
        await bindTo(ctx, rng, ed.value);
        notice(''); renderMode();
        finish(c.res, 'Trecho atualizado. Ele continua vinculado ao painel.' + (bibMsg ? ' ' + bibMsg : ''), ok);
      });
    } catch (e) {
      status('Não foi possível atualizar: ' + errMsg(e), 'err');
    } finally { setBusy(false); }
  }

  async function reloadBound() {
    setBusy(true);
    try {
      await Word.run(async function (ctx) {
        var b = await boundRange(ctx);
        if (!b) { unbindState(false); return; }
        var r = await bindTo(ctx, b);
        Object.keys(r.images).forEach(function (k) { S.images[k] = r.images[k]; }); renderImages();
        setEditor(r.latex); notice('');
        status('Trecho recarregado do Word.', 'ok');
      });
    } catch (e) { status('Não foi possível recarregar: ' + errMsg(e), 'err'); }
    finally { setBusy(false); }
  }

  // Limita o trecho a uma seção e nunca inclui o parágrafo que guarda a quebra de seção
  // (apagá-la mudaria cabeçalhos, rodapés e o fundo da capa).
  async function limitToSection(ctx, range) {
    var secs = ctx.document.sections; secs.load('items');
    await ctx.sync();
    if (secs.items.length < 2) return null;
    var info = secs.items.map(function (sec, i) {
      var x = range.intersectWithOrNullObject(sec.body.getRange('Whole'));
      x.load('isNullObject');
      return { i: i, x: x, last: sec.body.paragraphs.getLast() };
    });
    await ctx.sync();
    var hit = info.filter(function (h) { return !h.x.isNullObject; });
    if (!hit.length) return null;
    var h = hit[0];
    var ps = h.x.paragraphs; ps.load('items');
    await ctx.sync();
    var items = ps.items.slice();
    if (h.i < secs.items.length - 1 && items.length) {
      var cmp = items[items.length - 1].getRange('Whole').compareLocationWith(h.last.getRange('Whole'));
      await ctx.sync();
      if (cmp.value === 'Equal' || cmp.value === 'InsideStart' || cmp.value === 'InsideEnd' || cmp.value === 'Contains') items.pop();
    }
    if (hit.length === 1 && items.length === ps.items.length) return null;     // nada a cortar
    if (!items.length) return { empty: true };
    return { range: items[0].getRange('Whole').expandTo(items[items.length - 1].getRange('Whole')), n: items.length };
  }

  // completa imagens cujos bytes não vieram no OOXML, lendo-as pela API do Word
  async function fillImages(ctx, range, r) {
    var order = r.imageOrder || [];
    if (!order.some(function (x) { return x.missing; })) return;
    var pics = range.inlinePictures; pics.load('items');
    await ctx.sync();
    var srcs = pics.items.map(function (p) { return p.getBase64ImageSrc(); });
    await ctx.sync();
    order.filter(function (x) { return x.inline; }).forEach(function (im, i) {
      if (!im.missing || !srcs[i] || !srcs[i].value) return;
      var b64 = srcs[i].value.replace(/^data:[^,]*,/, '');
      var ct = /^\/9j\//.test(b64) ? 'image/jpeg' : /^R0lG/.test(b64) ? 'image/gif' : 'image/png';
      var sz = O.imageSize(b64);
      r.images[im.name] = { base64: b64, contentType: ct, width: sz.w, height: sz.h };
      im.missing = false;
    });
  }

  async function readSelection(fromAuto, diag) {
    if (needWord()) return;
    setBusy(true);
    if (!fromAuto) status('Lendo a seleção…');
    try {
      await Word.run(async function (ctx) {
        var sel = ctx.document.getSelection();
        var paras = sel.paragraphs; paras.load('items');
        await ctx.sync();
        var n = paras.items.length;
        if (!n) { status('Selecione um ou mais parágrafos no documento.', 'warn'); return; }
        var range = paras.items[0].getRange('Whole').expandTo(paras.items[n - 1].getRange('Whole'));
        var trimmed = false, nParas = n;
        var lim = await limitToSection(ctx, range);
        if (lim && lim.empty) { status('A seleção contém só a quebra de seção, que não é editável pelo painel.', 'warn'); return; }
        if (lim) { range = lim.range; nParas = lim.n; trimmed = true; }
        var r = await readRange(ctx, range);
        await fillImages(ctx, range, r);
        if (diag) {
          setEditor(['% Diagnóstico da leitura', '% parágrafos na seleção: ' + n + ' | no trecho: ' + nParas + ' | no OOXML: ' + r.paraCount + (trimmed ? ' | cortado na seção' : ''),
            '% quebras de seção no OOXML: ' + JSON.stringify(r.sectBreaks), '% partes: ' + r.partNames.join(', '),
            '% imagens: ' + JSON.stringify(r.imageOrder), '% avisos: ' + JSON.stringify(r.warnings)].join('\n') + '\n\n' + r.latex);
          status('Diagnóstico gerado (nada foi vinculado).', 'ok');
          return;
        }
        if (r.sectBreaks && r.sectBreaks.length) {
          S.bound = false; setEditor(r.latex);
          notice('O trecho contém uma quebra de seção e não pode ser vinculado. Selecione só a parte antes ou depois da quebra.');
          status(''); renderMode();
          return;
        }
        if (!api('1.4')) {
          setEditor(r.latex);
          status('Esta versão do Word não permite vincular o trecho: use "Inserir no cursor" com o trecho selecionado.', 'warn');
          return;
        }
        range.insertBookmark(BM);
        if (trimmed) { range.select(); S.suppressUntil = Date.now() + 1500; }
        await ctx.sync();
        S.bound = true; S.paras = nParas; S.base = r.latex; S.clean = r.latex; S.unsafe = r.unsafe;
        Object.keys(r.images).forEach(function (k2) { S.images[k2] = r.images[k2]; }); renderImages();
        setEditor(r.latex);
        notice(r.unsafe.length ? 'Atenção: o trecho contém ' + r.unsafe.join(', ') + ', que não têm equivalente em LaTeX e serão perdidos se você atualizar.' : '');
        var w = r.warnings.slice();
        var miss = (r.imageOrder || []).filter(function (x) { return x.missing; }).map(function (x) { return x.name; });
        if (miss.length) w.unshift('Não consegui ler ' + miss.length + ' imagem(ns) do documento (' + miss.join(', ') + '). Não atualize o trecho sem carregá-las em "Imagens", ou elas serão trocadas por marcadores.');
        if (trimmed) w.unshift('O trecho foi limitado até a quebra de seção (ela fica intacta no documento).');
        status(w.length ? 'Trecho lido e vinculado:' : 'Trecho lido e vinculado.', w.length ? 'warn' : 'ok', w);
      });
    } catch (e) {
      status('Não foi possível ler a seleção: ' + errMsg(e), 'err');
    } finally { setBusy(false); }
  }

  async function unbind(clear) {
    if (typeof Word !== 'undefined' && api('1.4')) {
      try { await Word.run(async function (ctx) { ctx.document.deleteBookmark(BM); await ctx.sync(); }); } catch (e) { /* ignora */ }
    }
    unbindState(clear);
  }

  // --------------------------------------------- leitura automática da seleção
  function scheduleSelCheck() {
    if (!S.auto) return;
    clearTimeout(S.selTimer);
    S.selTimer = setTimeout(function () {
      if (S.busy) { S.selPending = true; return; }
      checkSelection();
    }, 450);
  }
  var INSIDE = { Equal: 1, Inside: 1, InsideStart: 1, InsideEnd: 1 };
  async function checkSelection() {
    if (!S.auto || S.busy || Date.now() < S.suppressUntil || typeof Word === 'undefined' || !api('1.4')) return;
    try {
      var action = null;
      await Word.run(async function (ctx) {
        var sel = ctx.document.getSelection();
        sel.load('isEmpty');
        var paras = sel.paragraphs; paras.load('items');
        await ctx.sync();
        var b = S.bound ? await boundRange(ctx) : null;
        if (S.bound && !b) { unbindState(false); }
        var inBound = false;
        if (b) {
          var cmp = sel.compareLocationWith(b);
          await ctx.sync();
          inBound = !!INSIDE[cmp.value];
        }
        if (inBound) {
          // dentro do trecho vinculado: se não há edições pendentes, acompanha mudanças feitas no Word
          if (!dirty() && Date.now() - S.lastRefresh > 1200) {
            S.lastRefresh = Date.now();
            var r = await readRange(ctx, b);
            if (norm(r.latex) !== norm(S.base)) { S.base = r.latex; S.clean = r.latex; S.unsafe = r.unsafe; setEditor(r.latex); }
          }
          return;
        }
        if (sel.isEmpty) {
          if (S.bound) {
            if (dirty()) {
              notice('Há edições no painel ainda não aplicadas ao trecho vinculado.', [
                { label: 'Aplicar (Ctrl+Enter)', primary: true, run: function () { updateBound(false); } },
                { label: 'Descartar', run: function () { unbind(true); } }
              ]);
            } else {
              action = 'unbind';
            }
          }
          return;
        }
        if (!paras.items.length) return;
        if (!S.bound && norm(ed.value)) {
          notice('Há um rascunho no painel. O que fazer com o trecho selecionado?', [
            { label: 'Substituir pelo rascunho', primary: true, run: function () { notice(''); insertAtCursor(); } },
            { label: 'Ler a seleção (descarta o rascunho)', run: function () { readSelection(false); } }
          ]);
          return;
        }
        if (dirty()) {
          notice('Você selecionou outro trecho, mas há edições no painel ainda não aplicadas.', [
            { label: 'Aplicar no trecho anterior', primary: true, run: function () { updateBound(false); } },
            { label: 'Descartar e ler a seleção', run: function () { readSelection(false); } }
          ]);
          return;
        }
        action = 'read';
      });
      if (action === 'unbind') { await unbind(true); status('Cursor posicionado: o que você escrever será inserido aqui.', ''); }
      else if (action === 'read') await readSelection(true);
    } catch (e) { /* silencioso: a leitura automática não deve incomodar */ }
  }

  // ------------------------------------------------------- botões
  $('bibInput').addEventListener('change', function (ev) {
    var f = (ev.target.files || [])[0]; ev.target.value = '';
    if (!f) return;
    var rd = new FileReader();
    rd.onload = async function () {
      if (!S.lib) S.lib = new BIB.Library('');
      var m = S.lib.merge(String(rd.result));
      renderBib();
      try {
        if (typeof Office !== 'undefined' && Office.context && Office.context.document) await saveLibrary();
        status('.bib carregado: ' + m.added + ' nova(s), ' + m.updated + ' atualizada(s). As referências ficam guardadas dentro deste documento.', 'ok');
      } catch (e) { status('.bib carregado no painel, mas não consegui guardá-lo no documento: ' + errMsg(e), 'warn'); }
    };
    rd.readAsText(f, 'utf-8');
  });
  $('bibFilter').addEventListener('input', renderBib);
  $('btnBibUpdate').addEventListener('click', async function () {
    if (needWord()) return;
    setBusy(true); status('Atualizando referências…');
    try { await Word.run(async function (ctx) { status(await regenBib(ctx, true), 'ok'); }); }
    catch (e) { status('Não foi possível atualizar as referências: ' + errMsg(e), 'err'); }
    finally { setBusy(false); }
  });

  $('btnPrimary').addEventListener('click', primary);
  $('btnRead').addEventListener('click', function (ev) { readSelection(false, ev.shiftKey); });
  $('btnUnbind').addEventListener('click', function () { unbind(!dirty()); status('Trecho desvinculado.', ''); });

  if (typeof Office !== 'undefined') {
    Office.onReady(function (info) {
      if (info.host && info.host !== Office.HostType.Word) { status('Este suplemento funciona só no Word.', 'err'); return; }
      if (info.host && !api('1.3')) { status('Esta versão do Word é antiga demais para o suplemento (precisa do WordApi 1.3).', 'err'); return; }
      if (info.host) {
        try {
          Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, scheduleSelCheck);
        } catch (e) { /* sem eventos: segue manual */ }
        loadLibrary();
        // remove um vínculo esquecido de uma sessão anterior
        Word.run(async function (ctx) { ctx.document.deleteBookmark(BM); await ctx.sync(); }).catch(function () {});
      }
    });
  }
  $('ver').textContent = 'v' + VERSION;
  renderImages();
  renderMode();
})();
