/* Painel do suplemento: liga o editor LaTeX ao documento aberto no Word */
(function () {
  'use strict';
  var L = window.IHRLX.latex2ooxml;
  var O = window.IHRLX.ooxml2latex;
  var BM = '_LtxEdit';                 // marcador oculto do trecho lido
  var S = { images: {}, unsafe: [], hasRead: false };

  var $ = function (id) { return document.getElementById(id); };
  var ed = $('editor');

  // ------------------------------------------------------------ status
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
  function busy(on) {
    ['btnInsert', 'btnRead', 'btnApply'].forEach(function (id) { $(id).dataset.busy = on ? '1' : ''; });
    $('btnInsert').disabled = on; $('btnRead').disabled = on;
    $('btnApply').disabled = on || !S.hasRead;
  }
  function errMsg(e) {
    var m = (e && (e.message || e.toString())) || 'erro desconhecido';
    if (e && e.debugInfo && e.debugInfo.errorLocation) m += ' (' + e.debugInfo.errorLocation + ')';
    return m;
  }

  // --------------------------------------------------------- rascunho
  try { var d = localStorage.getItem('ihrlx.draft'); if (d) ed.value = d; } catch (e) { /* sem armazenamento */ }
  var saveT = null;
  ed.addEventListener('input', function () {
    clearTimeout(saveT);
    saveT = setTimeout(function () { try { localStorage.setItem('ihrlx.draft', ed.value); } catch (e) { /* ignora */ } }, 400);
  });
  ed.addEventListener('keydown', function (e) {
    if (e.key === 'Tab' && !e.shiftKey) {
      e.preventDefault();
      var s = ed.selectionStart, t = ed.selectionEnd;
      ed.setRangeText('  ', s, t, 'end');
    } else if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
      e.preventDefault(); insertAtCursor();
    }
  });

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
    var files = Array.prototype.slice.call(ev.target.files || []);
    files.forEach(function (f) {
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
  var SNIP = {
    secao: '\\chapter{Título 1}\n\nTexto do capítulo.\n\n\\section{Título 2}\n\nTexto da seção.\n\n\\subsection{Título 3}\n\nTexto da subseção.\n',
    figura: '\\begin{figure}\n  \\caption{Título da figura}\\label{fig:nome}\n  \\includegraphics[width=12cm]{arquivo.png}\n  \\fonte{IHR, ' + new Date().getFullYear() + '.}\n\\end{figure}\n',
    tabela: '\\begin{table}\n  \\caption{Título da tabela}\\label{tab:nome}\n  \\begin{tabular}{Nc}\n    \\cab{Nome} & \\cab{Valor} \\\\\n    Item A & 10 \\\\\n    Item B & 20 \\\\\n  \\end{tabular}\n  \\fonte{IHR, ' + new Date().getFullYear() + '.}\n\\end{table}\n',
    quadro: '\\begin{quadro}\n  \\caption{Título do quadro}\\label{qua:nome}\n  \\begin{tabular}{|c|c|c|}\n    \\hline\n    Coluna 1 & Coluna 2 & Coluna 3 \\\\ \\hline\n    a & b & c \\\\ \\hline\n  \\end{tabular}\n  \\fonte{IHR, ' + new Date().getFullYear() + '.}\n\\end{quadro}\n',
    equacao: 'Conforme \\eqref{eq:nome}:\n\\begin{equation}\n  y = a x^2 + b x + c\n  \\label{eq:nome}\n\\end{equation}\n',
    listas: '\\begin{enumerate}\n  \\item Primeiro passo.\n  \\item Segundo passo.\n\\end{enumerate}\n\n\\begin{itemize}\n  \\item Um ponto.\n  \\item Outro ponto.\n\\end{itemize}\n',
    texto: 'Texto com \\textbf{negrito}, \\textit{itálico}, \\underline{sublinhado}, H\\textsubscript{2}O, m\\textsuperscript{2}, uma fórmula $\\alpha = \\frac{a}{b}$ e uma nota\\footnote{Texto da nota.}.\n',
    sumario: '\\tableofcontents\n'
  };
  $('snippets').addEventListener('change', function (e) {
    var k = e.target.value; e.target.value = '';
    if (!SNIP[k]) return;
    var s = ed.selectionStart;
    var before = ed.value.slice(0, s);
    var pre = !before || /\n\n$/.test(before) ? '' : (/\n$/.test(before) ? '\n' : '\n\n');
    ed.setRangeText(pre + SNIP[k], s, ed.selectionEnd, 'end');
    ed.focus(); ed.dispatchEvent(new Event('input'));
  });

  // ------------------------------------------------------- Word: apoio
  function api(v) { return typeof Office !== 'undefined' && Office.context && Office.context.requirements && Office.context.requirements.isSetSupported('WordApi', v); }

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

  // atualiza a numeração (SEQ), depois as referências (REF) e o sumário
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
    } catch (e) {
      return false;
    }
  }

  function convertEditor(kinds) {
    var res = L.convert(ed.value, { images: S.images, labelKinds: kinds });
    return { res: res, pkg: L.buildPackage(res) };
  }
  function finish(res, okMsg, fieldsOk) {
    var w = res.warnings.slice();
    if (!fieldsOk && res.fields) w.push('Não consegui atualizar os campos automaticamente: selecione tudo (Ctrl+A) e pressione F9 para atualizar números e referências.');
    status(okMsg, w.length ? 'warn' : 'ok', w);
  }
  function needWord() {
    if (typeof Word === 'undefined') { status('Abra este painel dentro do Word para usar os botões.', 'err'); return true; }
    return false;
  }

  // ------------------------------------------------------- ações
  async function insertAtCursor() {
    if (needWord()) return;
    if (!ed.value.trim()) { status('O editor está vazio.', 'warn'); return; }
    busy(true); status('Convertendo…');
    try {
      await Word.run(async function (ctx) {
        var kinds = await labelKinds(ctx);
        var c = convertEditor(kinds);
        var sel = ctx.document.getSelection();
        sel.insertOoxml(c.pkg, 'Replace');
        await ctx.sync();
        var ok = c.res.fields ? await updateFields(ctx, c.res) : true;
        finish(c.res, 'Inserido no documento.', ok);
      });
    } catch (e) {
      status('Não foi possível inserir: ' + errMsg(e), 'err');
    } finally { busy(false); }
  }

  async function readSelection() {
    if (needWord()) return;
    busy(true); status('Lendo a seleção…');
    try {
      await Word.run(async function (ctx) {
        var sel = ctx.document.getSelection();
        var paras = sel.paragraphs;
        paras.load('items');
        await ctx.sync();
        if (!paras.items.length) { status('Selecione um ou mais parágrafos no documento.', 'warn'); return; }
        var range = paras.getFirst().getRange('Whole').expandTo(paras.getLast().getRange('Whole'));
        var ox = range.getOoxml();
        await ctx.sync();
        var r = O.convert(ox.value);
        Object.keys(r.images).forEach(function (k) { S.images[k] = r.images[k]; });
        renderImages();
        if (api('1.4')) {
          range.insertBookmark(BM);   // substitui um marcador anterior de mesmo nome
          await ctx.sync();
          S.hasRead = true;
        } else {
          S.hasRead = false;
          r.warnings.push('Esta versão do Word não permite marcar o trecho: use "Inserir no cursor" com o trecho selecionado.');
        }
        S.unsafe = r.unsafe;
        ed.value = r.latex; ed.dispatchEvent(new Event('input'));
        var info = $('readInfo');
        info.classList.remove('hidden');
        info.innerHTML = '';
        info.appendChild(document.createTextNode('Trecho lido: ' + paras.items.length + ' parágrafo(s). Edite e clique em "Aplicar no trecho lido".'));
        if (r.unsafe.length) {
          var lab = document.createElement('label');
          var cb = document.createElement('input'); cb.type = 'checkbox'; cb.id = 'unsafeOk';
          lab.appendChild(cb);
          lab.appendChild(document.createTextNode(' A seleção contém ' + r.unsafe.join(', ') + ', que não têm equivalente em LaTeX e serão perdidos ao aplicar. Aplicar mesmo assim'));
          info.appendChild(lab);
        }
        status(r.warnings.length ? 'Seleção convertida, com avisos:' : 'Seleção convertida.', r.warnings.length ? 'warn' : 'ok', r.warnings);
      });
    } catch (e) {
      status('Não foi possível ler a seleção: ' + errMsg(e), 'err');
    } finally { busy(false); }
  }

  async function applyEdit() {
    if (needWord()) return;
    if (S.unsafe.length && !($('unsafeOk') && $('unsafeOk').checked)) {
      status('Marque a confirmação abaixo dos botões para substituir um trecho com elementos que serão perdidos.', 'warn'); return;
    }
    busy(true); status('Aplicando…');
    try {
      await Word.run(async function (ctx) {
        var r = ctx.document.getBookmarkRangeOrNullObject(BM);
        r.load('isNullObject');
        await ctx.sync();
        if (r.isNullObject) { S.hasRead = false; status('Não encontrei o trecho lido (talvez tenha sido apagado). Selecione de novo e clique em "Ler seleção".', 'err'); return; }
        var kinds = await labelKinds(ctx);
        var c = convertEditor(kinds);
        r.insertOoxml(c.pkg, 'Replace');
        try { ctx.document.deleteBookmark(BM); } catch (e) { /* ignora */ }
        await ctx.sync();
        S.hasRead = false; S.unsafe = [];
        $('readInfo').classList.add('hidden');
        var ok = c.res.fields ? await updateFields(ctx, c.res) : true;
        finish(c.res, 'Trecho atualizado no documento.', ok);
      });
    } catch (e) {
      status('Não foi possível aplicar: ' + errMsg(e), 'err');
    } finally { busy(false); }
  }

  $('btnInsert').addEventListener('click', insertAtCursor);
  $('btnRead').addEventListener('click', readSelection);
  $('btnApply').addEventListener('click', applyEdit);

  if (typeof Office !== 'undefined') {
    Office.onReady(function (info) {
      if (info.host && info.host !== Office.HostType.Word) status('Este suplemento funciona só no Word.', 'err');
      else if (info.host && !api('1.3')) status('Esta versão do Word é antiga demais para o suplemento (precisa do WordApi 1.3).', 'err');
    });
  }
  renderImages();
})();
