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
  function isUnnum(p) { return /sem numera/i.test(p.style || ''); }
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
    var d = localStorage.getItem('ihrlx.draft'); if (d) { ed.value = d; }
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
      e.preventDefault(); primary(e);
    }
  });
  // ---- realce: camada <pre> atrás do textarea transparente (comandos, comentários e $)
  var hlEl = $('edHL'), hlPending = false;
  function escH(t) { return t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
  function highlight(t) {
    var re = /\\(?:begin|end)\{[^}\n]*\}|\\(?:[A-Za-z@]+\*?|[^A-Za-z\n])|%[^\n]*|\$\$?/g, out = '', last = 0, m;
    while ((m = re.exec(t))) {
      out += escH(t.slice(last, m.index));
      var tok = m[0], cls = tok[0] === '%' ? 'k' : tok[0] === '$' ? 'm' : 'c';
      out += '<span class="' + cls + '">' + escH(tok) + '</span>';
      last = m.index + tok.length;
    }
    return out + escH(t.slice(last)) + '\n ';
  }
  function syncHL() {
    if (hlPending) return; hlPending = true;
    requestAnimationFrame(function () { hlPending = false; hlEl.innerHTML = highlight(ed.value); hlEl.scrollTop = ed.scrollTop; hlEl.scrollLeft = ed.scrollLeft; });
  }
  ed.addEventListener('input', syncHL);
  ed.addEventListener('scroll', function () { hlEl.scrollTop = ed.scrollTop; hlEl.scrollLeft = ed.scrollLeft; });
  var wrapOn = true;
  try { wrapOn = localStorage.getItem('ihrlx.wrap') !== '0'; } catch (e) { /* padrão: quebra */ }
  function applyWrap() {
    $('edWrap').classList.toggle('wrap', wrapOn); ed.wrap = wrapOn ? 'soft' : 'off';
    $('btnWrap').classList.toggle('on', wrapOn); syncHL();
  }
  $('btnWrap').addEventListener('click', function () {
    wrapOn = !wrapOn; try { localStorage.setItem('ihrlx.wrap', wrapOn ? '1' : '0'); } catch (e) { /* ignora */ }
    applyWrap(); ed.focus();
  });
  function setMax(on) {
    document.body.classList.toggle('max', on); $('btnMax').classList.toggle('on', on);
    $('btnMax').textContent = on ? '⤡' : '⤢'; $('btnMax').title = on ? 'Voltar ao painel completo (Esc)' : 'Maximizar o editor (Esc volta)';
    syncHL();
  }
  $('btnMax').addEventListener('click', function () { setMax(!document.body.classList.contains('max')); ed.focus(); });
  document.addEventListener('keydown', function (e) { if (e.key === 'Escape' && document.body.classList.contains('max')) setMax(false); });
  if (window.ResizeObserver) new ResizeObserver(syncHL).observe($('edWrap'));
  applyWrap();
  // ---- autocompletar: "\" + letras sugere comandos; dentro de \ref{ / \cite{ sugere rótulos e referências
  // modelo: | marca onde o cursor fica depois de inserir
  var CMDS = [
    ['chapter', '\\chapter{|}', 'Título 1'], ['chapter*', '\\chapter*{|}', 'título sem número'], ['section', '\\section{|}', 'Título 2'],
    ['subsection', '\\subsection{|}', 'Título 3'], ['subsubsection', '\\subsubsection{|}', 'Título 4'], ['paragraph', '\\paragraph{|}', 'Título 5'],
    ['subparagraph', '\\subparagraph{|}', 'Título 6'], ['textbf', '\\textbf{|}', 'negrito'], ['textit', '\\textit{|}', 'itálico'], ['emph', '\\emph{|}', 'ênfase'],
    ['underline', '\\underline{|}', 'sublinhado'], ['texttt', '\\texttt{|}', 'monoespaçado'], ['textsuperscript', '\\textsuperscript{|}', 'sobrescrito'],
    ['textsubscript', '\\textsubscript{|}', 'subscrito'], ['footnote', '\\footnote{|}', 'nota de rodapé'], ['url', '\\url{|}', 'link'], ['href', '\\href{|}{}', 'link com texto'],
    ['cite', '\\cite{|}', 'citação (AUTOR, ano)'], ['textcite', '\\textcite{|}', 'citação Autor (ano)'], ['citeauthor', '\\citeauthor{|}', 'só o autor'], ['citeyear', '\\citeyear{|}', 'só o ano'],
    ['ref', '\\ref{|}', 'número do rótulo'], ['eqref', '\\eqref{|}', 'número da equação (n)'], ['pageref', '\\pageref{|}', 'página do rótulo'], ['label', '\\label{|}', 'rótulo'],
    ['caption', '\\caption{|}', 'legenda'], ['fonte', '\\fonte{|}', 'fonte da figura/tabela'], ['includegraphics', '\\includegraphics[width=14cm]{|}', 'imagem'],
    ['item', '\\item |', 'item de lista'], ['newpage', '\\newpage', 'quebra de página'], ['tableofcontents', '\\tableofcontents', 'sumário'],
    ['printbibliography', '\\printbibliography', 'capítulo Referências'], ['refmanual', '\\refmanual{|}', 'referência digitada à mão'],
    ['hline', '\\hline', 'linha da tabela'], ['multicolumn', '\\multicolumn{|}{c}{}', 'mesclar colunas'], ['cab', '\\cab{|}', 'cabeçalho da tabela IHR'],
    ['linhaihr', '\\linhaihr', 'linha do cabeçalho IHR'], ['newcommand', '\\newcommand{\\|}{}', 'novo comando'], ['DeclareMathOperator', '\\DeclareMathOperator{\\|}{}', 'novo operador'],
    ['begin', '\\begin{|}', 'ambiente'], ['end', '\\end{|}', 'fim de ambiente'],
    ['begin{equation}', '\\begin{equation}\n  |\n  \\label{eq:}\n\\end{equation}', 'equação numerada'],
    ['begin{align}', '\\begin{align}\n  | &= \\\\\n  &= \n\\end{align}', 'equações alinhadas'],
    ['begin{figure}', '\\begin{figure}\n  \\caption{|}\\label{fig:}\n  \\includegraphics[width=14cm]{}\n  \\fonte{IHR, ' + new Date().getFullYear() + '.}\n\\end{figure}', 'figura'],
    ['begin{table}', '\\begin{table}\n  \\caption{|}\\label{tab:}\n  \\rowcolors{2}{white}{ihrcinza}\n  \\begin{tabular}{Nc}\n    \\multicolumn{1}{l}{\\cab{Coluna 1}} & \\cab{Coluna 2} \\\\ \\linhaihr\n    a & b \\\\\n  \\end{tabular}\n  \\fonte{IHR, ' + new Date().getFullYear() + '.}\n\\end{table}', 'tabela IHR'],
    ['begin{quadro}', '\\begin{quadro}\n  \\caption{|}\\label{qua:}\n  \\begin{tabular}{|c|c|}\n    \\hline\n    a & b \\\\ \\hline\n  \\end{tabular}\n  \\fonte{IHR, ' + new Date().getFullYear() + '.}\n\\end{quadro}', 'quadro com grade'],
    ['begin{itemize}', '\\begin{itemize}\n  \\item |\n\\end{itemize}', 'lista com marcadores'], ['begin{enumerate}', '\\begin{enumerate}\n  \\item |\n\\end{enumerate}', 'lista numerada'],
    ['begin{tabular}', '\\begin{tabular}{|}\n  \n\\end{tabular}', 'tabela simples'], ['begin{cases}', '\\begin{cases}\n  | & \\text{se } \\\\\n  & \\text{caso contrário}\n\\end{cases}', 'casos (em fórmula)'],
    ['begin{bmatrix}', '\\begin{bmatrix}\n  | & \\\\\n  & \n\\end{bmatrix}', 'matriz [ ]'], ['begin{pmatrix}', '\\begin{pmatrix}\n  | & \\\\\n  & \n\\end{pmatrix}', 'matriz ( )'],
    ['begin{aligned}', '\\begin{aligned}\n  | &= \\\\\n  &= \n\\end{aligned}', 'alinhado (dentro de equação)'], ['begin{center}', '\\begin{center}\n  |\n\\end{center}', 'centralizado'],
    ['frac', '\\frac{|}{}', 'fração'], ['dfrac', '\\dfrac{|}{}', 'fração grande'], ['sqrt', '\\sqrt{|}', 'raiz'], ['sum', '\\sum_{|}^{}', 'somatório'], ['prod', '\\prod_{|}^{}', 'produtório'],
    ['int', '\\int_{|}^{}', 'integral'], ['iint', '\\iint', 'integral dupla'], ['oint', '\\oint', 'integral fechada'], ['lim', '\\lim_{|}', 'limite'],
    ['left(', '\\left( | \\right)', 'parênteses ajustáveis'], ['left[', '\\left[ | \\right]', 'colchetes ajustáveis'], ['left\\{', '\\left\\{ | \\right\\}', 'chaves ajustáveis'],
    ['mathbf', '\\mathbf{|}', 'negrito (fórmula)'], ['boldsymbol', '\\boldsymbol{|}', 'negrito itálico'], ['mathrm', '\\mathrm{|}', 'reto'], ['mathbb', '\\mathbb{|}', 'duplo (ℝ)'],
    ['mathcal', '\\mathcal{|}', 'caligráfico'], ['mathfrak', '\\mathfrak{|}', 'fraktur'], ['operatorname', '\\operatorname{|}', 'operador'], ['text', '\\text{|}', 'texto na fórmula'],
    ['hat', '\\hat{|}', 'chapéu'], ['bar', '\\bar{|}', 'barra'], ['vec', '\\vec{|}', 'vetor'], ['dot', '\\dot{|}', 'derivada'], ['ddot', '\\ddot{|}', '2ª derivada'],
    ['tilde', '\\tilde{|}', 'til'], ['overline', '\\overline{|}', 'sobrelinha'], ['partial', '\\partial', '∂'], ['nabla', '\\nabla', '∇'], ['infty', '\\infty', '∞'],
    ['cdot', '\\cdot', '·'], ['times', '\\times', '×'], ['div', '\\div', '÷'], ['pm', '\\pm', '±'], ['mp', '\\mp', '∓'], ['leq', '\\leq', '≤'], ['geq', '\\geq', '≥'],
    ['neq', '\\neq', '≠'], ['approx', '\\approx', '≈'], ['equiv', '\\equiv', '≡'], ['sim', '\\sim', '∼'], ['propto', '\\propto', '∝'], ['in', '\\in', '∈'], ['notin', '\\notin', '∉'],
    ['subset', '\\subset', '⊂'], ['subseteq', '\\subseteq', '⊆'], ['cup', '\\cup', '∪'], ['cap', '\\cap', '∩'], ['forall', '\\forall', '∀'], ['exists', '\\exists', '∃'],
    ['rightarrow', '\\rightarrow', '→'], ['leftarrow', '\\leftarrow', '←'], ['Rightarrow', '\\Rightarrow', '⇒'], ['Leftrightarrow', '\\Leftrightarrow', '⇔'], ['to', '\\to', '→'],
    ['mapsto', '\\mapsto', '↦'], ['ldots', '\\ldots', '…'], ['cdots', '\\cdots', '⋯'], ['vdots', '\\vdots', '⋮'], ['ddots', '\\ddots', '⋱'], ['circ', '\\circ', '∘'],
    ['sin', '\\sin', 'seno'], ['cos', '\\cos', 'cosseno'], ['tan', '\\tan', 'tangente'], ['log', '\\log', 'log'], ['ln', '\\ln', 'ln'], ['exp', '\\exp', 'exp'],
    ['max', '\\max', 'máx'], ['min', '\\min', 'mín'], ['det', '\\det', 'det'], ['quad', '\\quad', 'espaço'], ['qquad', '\\qquad', 'espaço duplo']
  ];
  'alpha beta gamma delta epsilon varepsilon zeta eta theta vartheta iota kappa lambda mu nu xi pi rho sigma tau upsilon phi varphi chi psi omega Gamma Delta Theta Lambda Xi Pi Sigma Phi Psi Omega'
    .split(' ').forEach(function (g) { CMDS.push([g, '\\' + g, 'letra grega']); });
  var acBox = document.createElement('ul'); acBox.id = 'acBox'; acBox.className = 'hidden'; $('edWrap').appendChild(acBox);
  var AC = { items: [], sel: 0, from: 0, open: false };
  function acClose() { AC.open = false; acBox.classList.add('hidden'); }
  function acContext() {
    var pos = ed.selectionStart, before = ed.value.slice(Math.max(0, pos - 200), pos);
    var m = /\\(ref|eqref|pageref|autoref)\{([^}\n]*)$/.exec(before);
    if (m) return { kind: 'label', q: m[2], from: pos - m[2].length };
    m = /\\(cite|textcite|citeonline|citeauthor|citeyear|nocite)(?:\[[^\]\n]*\])*\{(?:[^}\n]*,)?\s*([^},\n]*)$/.exec(before);
    if (m) return { kind: 'cite', q: m[2], from: pos - m[2].length };
    m = /\\([A-Za-z@]*(?:\{[A-Za-z*]*)?)$/.exec(before);
    if (m && ed.selectionStart === ed.selectionEnd) return { kind: 'cmd', q: m[1], from: pos - m[1].length - 1 };
    return null;
  }
  function acCandidates(c) {
    var q = c.q, ql = q.toLowerCase(), out = [];
    function rank(name) { return name.indexOf(q) === 0 ? 0 : name.toLowerCase().indexOf(ql) === 0 ? 1 : name.toLowerCase().indexOf(ql) > 0 ? 2 : 9; }
    if (c.kind === 'cmd') {
      if (!q) return [];
      var user = [];
      try { user = Object.keys(editorDefs()).map(function (k) { return [k, '\\' + k, 'comando seu / pacote']; }); } catch (e) { /* sem definições */ }
      CMDS.concat(user).forEach(function (it) { var r = rank(it[0]); if (r < 9 && it[0] !== q) out.push({ r: r, label: '\\' + it[0], tpl: it[1], hint: it[2] }); });
    } else if (c.kind === 'label') {
      var seen = {};
      (ed.value.match(/\\label\{([^}]+)\}/g) || []).map(function (x) { return x.slice(7, -1); }).concat(Object.keys(S.docLabels || {})).forEach(function (l) {
        if (seen[l]) return; seen[l] = 1; var r = q ? rank(l) : 0; if (r < 9) out.push({ r: r, label: l, tpl: l + '|', hint: 'rótulo' });
      });
    } else if (c.kind === 'cite' && S.lib) {
      S.lib.order.forEach(function (k) {
        var r = q ? Math.min(rank(k), BIB.sortKey(S.lib.label(k)).indexOf(BIB.sortKey(q)) >= 0 ? 2 : 9) : 0;
        if (r < 9) out.push({ r: r, label: k, tpl: k + '|', hint: S.lib.label(k) });
      });
    }
    out.sort(function (a, b) { return a.r - b.r || a.label.length - b.label.length; });
    return out.slice(0, 8);
  }
  function caretXY() {
    // mede a posição do cursor pela camada de realce (mesmo layout do textarea)
    var pos = ed.selectionStart;
    hlEl.innerHTML = escH(ed.value.slice(0, pos)) + '<span id="acMark">\u200b</span>' + escH(ed.value.slice(pos)) + '\n ';
    var mk = $('acMark'), xy = { x: mk.offsetLeft - ed.scrollLeft, y: mk.offsetTop - ed.scrollTop, h: mk.offsetHeight || 18 };
    syncHL();
    return xy;
  }
  function acRender() {
    acBox.innerHTML = '';
    AC.items.forEach(function (it, i) {
      var li = document.createElement('li'); if (i === AC.sel) li.className = 'sel';
      var a = document.createElement('code'); a.textContent = it.label; var b = document.createElement('span'); b.textContent = it.hint;
      li.appendChild(a); li.appendChild(b);
      li.onmousedown = function (ev) { ev.preventDefault(); AC.sel = i; acAccept(); };
      acBox.appendChild(li);
    });
  }
  function acUpdate() {
    var c = acContext();
    if (!c) return acClose();
    AC.items = acCandidates(c); AC.sel = 0; AC.from = c.from; AC.kind = c.kind;
    if (!AC.items.length) return acClose();
    acRender(); AC.open = true; acBox.classList.remove('hidden');
    var xy = caretXY(), wrapW = $('edWrap').clientWidth, wrapH = $('edWrap').clientHeight;
    var w = Math.min(300, wrapW - 16), left = Math.max(4, Math.min(xy.x, wrapW - w - 18));
    var top = xy.y + xy.h + 2; if (top + acBox.offsetHeight > wrapH && xy.y - acBox.offsetHeight > 0) top = xy.y - acBox.offsetHeight - 2;
    acBox.style.width = w + 'px'; acBox.style.left = left + 'px'; acBox.style.top = top + 'px';
  }
  function acAccept() {
    var it = AC.items[AC.sel]; if (!it) return acClose();
    var pos = ed.selectionStart, line0 = ed.value.lastIndexOf('\n', AC.from - 1) + 1;
    var indent = (/^[ \t]*/.exec(ed.value.slice(line0, AC.from)) || [''])[0];
    var tpl = it.tpl.replace(/\n/g, '\n' + indent), k = tpl.indexOf('|'); tpl = tpl.replace('|', '');
    // completa só o que falta: se o texto seguinte já fecha a chave, não duplica
    var end = pos, close = false;
    if (AC.kind !== 'cmd') {
      var rest = /^[^},\s]*/.exec(ed.value.slice(pos))[0]; end = pos + rest.length;
      close = !/^[},]/.test(ed.value.slice(end));          // \ref{ sem fechar: fecha a chave
      if (close) tpl += '}';
    }
    ed.setRangeText(tpl, AC.from, end, 'end');
    var caret = AC.from + (close ? tpl.length : (k >= 0 ? k : tpl.length)); ed.setSelectionRange(caret, caret);
    acClose(); ed.dispatchEvent(new Event('input'));
  }
  ed.addEventListener('keydown', function (e) {
    if (!AC.open) return;
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); e.stopImmediatePropagation(); AC.sel = (AC.sel + (e.key === 'ArrowDown' ? 1 : AC.items.length - 1)) % AC.items.length; acRender(); }
    else if ((e.key === 'Enter' && !e.ctrlKey && !e.metaKey) || e.key === 'Tab') { e.preventDefault(); e.stopImmediatePropagation(); acAccept(); }
    else if (e.key === 'Escape') { e.preventDefault(); e.stopImmediatePropagation(); acClose(); }
    else if (e.key === 'Enter') acClose();
  }, true);
  ed.addEventListener('input', function (e) { if (e.isTrusted) acUpdate(); });
  ed.addEventListener('blur', function () { setTimeout(acClose, 150); });
  ed.addEventListener('scroll', function () { if (AC.open) acClose(); });
  ed.addEventListener('click', acClose);

  function setEditor(text) {
    ed.value = text; syncHL();
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
      th.title = nm.title + ' · clique para inserir no editor';
      th.onclick = nm.onclick = function () { insertFigure(n); };
      li.appendChild(th); li.appendChild(nm); li.appendChild(bt); ul.appendChild(li);
    });
  }
  function insertFigure(n) {
    var im = S.images[n], base = n.replace(/\.[a-z0-9]+$/i, '');
    var wcm = Math.min(14, Math.round(((im.width || 600) / 96 * 2.54) * 2) / 2);
    var inc = '\\includegraphics[width=' + String(wcm).replace('.', ',').replace(',', '.') + 'cm]{' + base + '}';
    var s0 = ed.selectionStart, before = ed.value.slice(0, s0);
    var insideFig = before.lastIndexOf('\\begin{figure}') > before.lastIndexOf('\\end{figure}');
    if (insideFig) { ed.setRangeText(inc, s0, ed.selectionEnd, 'end'); }
    else {
      var lbl = 'fig:' + base.normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^A-Za-z0-9]+/g, '-').replace(/^-|-$/g, '').toLowerCase();
      var pre = !before || /\n\n$/.test(before) ? '' : (/\n$/.test(before) ? '\n' : '\n\n');
      var head = pre + '\\begin{figure}\n  \\caption{';
      var tail = '}\\label{' + lbl + '}\n  ' + inc + '\n  \\fonte{IHR, ' + new Date().getFullYear() + '.}\n\\end{figure}\n';
      ed.setRangeText(head + tail, s0, ed.selectionEnd, 'end');
      var cap = s0 + head.length; ed.setSelectionRange(cap, cap);       // cursor dentro do \caption{}
    }
    ed.focus(); ed.dispatchEvent(new Event('input'));
    status('Figura "' + base + '" inserida no editor. Escreva a legenda e aplique.', 'ok');
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
    S.docLabels = map;
    return map;
  }
  // ---- numeração provisória: sem a API de campos (WordApi 1.5) o Word não renumera sozinho,
  // então o painel conta as legendas/equações que já existem antes do ponto de inserção
  var SEQKINDS = ['Figura', 'Tabela', 'Quadro', 'Equação'];
  function countSeq(ooxml) {
    var instr = (ooxml.match(/<w:instrText[^>]*>[^<]*<\/w:instrText>/g) || []).map(function (x) { return x.replace(/<[^>]+>/g, ''); }).join(' ');
    var c = {};
    SEQKINDS.forEach(function (k) { c[k] = (instr.match(new RegExp('SEQ\\s+' + k + '\\b', 'g')) || []).length; });
    return c;
  }
  function countLatex(tex) {
    var c = { Figura: 0, Tabela: 0, Quadro: 0, 'Equação': 0 };
    var ENV = { figure: 'Figura', table: 'Tabela', quadro: 'Quadro' };
    (tex || '').replace(/\\begin\{(figure|table|quadro)\}([\s\S]*?)\\end\{\1\}/g, function (m, e, body) { if (/\\caption\b/.test(body)) c[ENV[e]]++; return ''; });
    c['Equação'] = (tex.match(/\\begin\{(equation|align|gather|multline|eqnarray|flalign)\}/g) || []).length;
    return c;
  }
  async function numbersBefore(ctx, point) {
    try {
      var r = ctx.document.body.getRange('Start').expandTo(point.getRange('Start'));
      var ox = r.getOoxml();
      await ctx.sync();
      return countSeq(ox.value);
    } catch (e) { return null; }       // ponto no início do documento
  }
  function addCounts(a, b) { var o = {}; SEQKINDS.forEach(function (k) { o[k] = (a && a[k] || 0) + (b && b[k] || 0); }); return o; }
  // texto atual dos rótulos citados no trecho e definidos fora dele (ex.: "3", "(2)")
  async function refTexts(ctx, tex, kinds) {
    var out = {};
    if (!api('1.4')) return out;
    var defined = {}; (tex.match(/\\label\{([^}]+)\}/g) || []).forEach(function (m) { defined[m.slice(7, -1)] = 1; });
    var want = {}; tex.replace(/\\(?:ref|eqref|autoref)\{([^}]+)\}/g, function (m, l) { if (!defined[l]) want[l] = 1; return ''; });
    var got = Object.keys(want).map(function (l) {
      var k = kinds[l] || (/^_Ref\d+$/.test(l) ? 'W' : null);
      if (!k) return null;
      var r = ctx.document.getBookmarkRangeOrNullObject(k === 'W' ? l : L.bookmarkName(l, k)); r.load('text,isNullObject');
      return { l: l, r: r };
    }).filter(Boolean);
    if (!got.length) return out;
    try { await ctx.sync(); } catch (e) { return out; }
    got.forEach(function (g) { if (!g.r.isNullObject && (g.r.text || '').trim()) out[g.l] = g.r.text.trim(); });
    return out;
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
  function convertEditor(kinds, cited, extra) {
    var res = L.convert(MAC.prepare(ed.value, S.preamble).text, convOpts(Object.assign({ images: S.images, labelKinds: kinds, bib: S.lib, citedKeys: cited || [] }, extra || {})));
    return { res: res, pkg: L.buildPackage(res) };
  }

  // ------------------------------------------------------- perfil do template
  // Um .dotx pode trazer um perfil (parte XML personalizada) que ajusta a conversão ao template.
  var PERFNS = 'urn:ihr-latex-word:perfil';
  var PERFIL_IHR = { nome: 'template IHR v4', legenda: ' - ', equacaoPorCapitulo: false,
    referencias: { iniciais: false, destaque: 'negrito' }, capaComQuadros: true };
  S.profile = PERFIL_IHR;
  async function loadProfile() {
    S.profile = PERFIL_IHR;
    try {
      var r = await oc(function (cb) { Office.context.document.customXmlParts.getByNamespaceAsync(PERFNS, cb); });
      if (r.status === 'succeeded' && r.value.length) {
        var x = await oc(function (cb) { r.value[0].getXmlAsync(cb); });
        if (x.status === 'succeeded') {
          var j = JSON.parse(window.IHRLX.xml.text(window.IHRLX.xml.parse(x.value)));
          S.profile = Object.assign({}, PERFIL_IHR, j, { referencias: Object.assign({}, PERFIL_IHR.referencias, j.referencias || {}) });
        }
      }
    } catch (e) { /* sem perfil: IHR */ }
    $('perfil').textContent = 'perfil: ' + S.profile.nome;
    $('perfil').title = S.profile === PERFIL_IHR ? 'Documento sem perfil próprio: usa o padrão do template IHR.' : 'Perfil lido do template deste documento.';
  }
  function convOpts(extra) {
    var P = S.profile;
    if (S.lib) S.lib.setStyle({ initials: !!P.referencias.iniciais });
    return Object.assign({ captionSep: P.legenda, eqChapter: !!P.equacaoPorCapitulo,
      bibEmph: P.referencias.destaque === 'italico' ? 'i' : 'b', coverFrames: P.capaComQuadros !== false }, extra);
  }

  // ------------------------------------------------------- preparar modelo
  // Acrescenta ao documento aberto os estilos que o LaTeX IHR usa e que ainda não existem nele.
  // Os que já existem (pelo nome interno: "heading 1", "caption"...) ficam como estão.
  var NOME_PT = { 'heading 1': 'Título 1', 'heading 2': 'Título 2', 'heading 3': 'Título 3', 'heading 4': 'Título 4', 'heading 5': 'Título 5',
    'heading 6': 'Título 6', 'caption': 'Legenda', 'toc heading': 'Cabeçalho do Sumário', 'bibliography': 'Bibliografia',
    'list paragraph': 'Parágrafo da Lista', 'table grid': 'Tabela com grade', 'plain table 3': 'Tabela Simples 3',
    'footnote text': 'Texto de nota de rodapé', 'footnote reference': 'Ref. de nota de rodapé', 'hyperlink': 'Hiperlink',
    'normal': 'Normal', 'normal table': 'Tabela normal', 'default paragraph font': 'Fonte parág. padrão', 'no list': 'Sem lista' };
  function nomePt(n) { var m = /^toc (\d)$/.exec(n); return m ? 'Sumário ' + m[1] : (NOME_PT[n] || n); }
  function templateStyles() {
    var out = [];
    (window.IHRLX.template.styles.match(/<w:style [\s\S]*?<\/w:style>/g) || []).forEach(function (x) {
      var id = (/w:styleId="([^"]+)"/.exec(x) || [])[1], nm = ((/<w:name w:val="([^"]+)"/.exec(x) || [])[1] || '').toLowerCase();
      var ty = (/w:type="([^"]+)"/.exec(x) || [])[1];
      var dep = []; x.replace(/<w:(?:basedOn|next|link) w:val="([^"]+)"/g, function (m, d) { dep.push(d); });
      out.push({ id: id, name: nm, type: ty, xml: x, dep: dep });
    });
    return out;
  }
  async function prepareTemplate() {
    if (needWord()) return;
    status('Verificando os estilos do documento…', '');
    try {
      var added = [], kept = [];
      await Word.run(async function (ctx) {
        var ox = ctx.document.body.paragraphs.getFirst().getRange('Whole').getOoxml();
        await ctx.sync();
        var have = {};
        var tree = window.IHRLX.xml.parse(ox.value);
        window.IHRLX.xml.find(tree, 'w:style').forEach(function (st) {
          var nm = window.IHRLX.xml.kid(st, 'w:name');
          if (nm) have[(window.IHRLX.xml.val(nm) || '').toLowerCase()] = true;
        });
        var all = templateStyles(), byId = {};
        all.forEach(function (t) { byId[t.id] = t; });
        var need = {};
        all.forEach(function (t) {
          if (/^(normal|normal table|default paragraph font|no list)$/.test(t.name) || t.id === 'Referncias') return;
          if (have[t.name]) { kept.push(nomePt(t.name)); return; }
          need[t.id] = t;
        });
        // dependências (basedOn/next) que também faltam
        Object.keys(need).forEach(function (id) {
          need[id].dep.forEach(function (d) { if (byId[d] && !have[byId[d].name] && !need[d] && !/^(normal|normal table|default paragraph font|no list)$/.test(byId[d].name)) need[d] = byId[d]; });
        });
        var list = Object.keys(need).map(function (k) { return need[k]; });
        if (!list.length) return;
        var defs = list.map(function (t) { return t.xml; }).join('');
        var body = '';
        list.filter(function (t) { return t.type === 'table'; }).forEach(function (t) {
          body += '<w:tbl><w:tblPr><w:tblStyle w:val="' + t.id + '"/><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
            '<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>';
        });
        list.filter(function (t) { return t.type === 'paragraph'; }).forEach(function (t) {
          body += '<w:p><w:pPr><w:pStyle w:val="' + t.id + '"/></w:pPr><w:r><w:t>.</w:t></w:r></w:p>';
        });
        var chars = list.filter(function (t) { return t.type === 'character'; });
        if (chars.length) body += '<w:p>' + chars.map(function (t) { return '<w:r><w:rPr><w:rStyle w:val="' + t.id + '"/></w:rPr><w:t>.</w:t></w:r>'; }).join('') + '</w:p>';
        var styles = window.IHRLX.template.styles.replace(/(<w:styles[^>]*>)[\s\S]*$/, '$1') +
          (window.IHRLX.template.styles.match(/<w:docDefaults>[\s\S]*?<\/w:docDefaults>/) || [''])[0] + defs + '</w:styles>';
        var pkg = L.buildPackage({ body: body, footnotes: [], rels: [], media: [], nums: [] }).replace(/(<pkg:part pkg:name="\/word\/styles\.xml"[^>]*><pkg:xmlData>)[\s\S]*?(<\/pkg:xmlData>)/, function (m, a, b) { return a + styles + b; });
        var paras = ctx.document.body.paragraphs; paras.load('items');
        await ctx.sync();
        var before = paras.items.length;
        var ins = ctx.document.body.insertOoxml(pkg, 'End');
        await ctx.sync();
        ins.delete();
        await ctx.sync();
        // o último parágrafo do documento não pode ser apagado: devolve o estilo Normal se sobrou um vazio
        var ps2 = ctx.document.body.paragraphs; ps2.load('items/text');
        await ctx.sync();
        var last = ps2.items[ps2.items.length - 1];
        if (ps2.items.length > before && !(last.text || '').trim()) { last.styleBuiltIn = Word.Style.normal; await ctx.sync(); }
        added = list.map(function (t) { return nomePt(t.name); });
      });
      var ul = $('tplList'); ul.innerHTML = '';
      function li(t, cls) { var e = document.createElement('li'); e.textContent = t; e.className = cls; ul.appendChild(e); }
      added.forEach(function (n) { li('+ ' + n, 'add'); });
      kept.forEach(function (n) { li('= ' + n, 'keep'); });
      status(added.length ? added.length + ' estilo(s) do LaTeX IHR acrescentado(s); ' + kept.length + ' já existia(m) e ficaram como estavam. Ajuste-os no Word (Página Inicial → Estilos) e salve como Modelo do Word (.dotx).'
        : 'Este documento já tem todos os estilos do LaTeX IHR (' + kept.length + ').', 'ok');
    } catch (e) { status('Não foi possível preparar o documento: ' + errMsg(e), 'err'); }
  }
  $('btnPrep').addEventListener('click', prepareTemplate);

  // ------------------------------------------------------- pacotes e comandos
  var MAC = window.IHRLX.macros, PRENS = 'urn:ihr-latex-word:preambulo';
  S.preamble = '';
  var PKGMSG = { nativo: 'já suportado', emulado: 'emulado', layout: 'ignorado: a formatação vem do modelo do Word', nao: 'não suportado: mande o nome do pacote para incluirmos' };
  function renderPreamble() {
    var src = $('preText').value, p = MAC.parseDefs(src), ul = $('preList'); ul.innerHTML = '';
    var nDefs = Object.keys(p.defs).length;
    $('preCount').textContent = p.packages.length + nDefs;
    p.packages.forEach(function (pk) {
      var k = MAC.classify(pk.name), li = document.createElement('li');
      li.className = k;
      var extra = k === 'emulado' ? ' (' + Object.keys(MAC.EMULADO[pk.name]).slice(0, 6).map(function (c) { return '\\' + c; }).join(', ') + '…)' : '';
      li.textContent = (k === 'nao' ? '✗ ' : k === 'layout' ? '– ' : '✓ ') + pk.name + ': ' + PKGMSG[k] + extra;
      ul.appendChild(li);
    });
    if (nDefs) { var li2 = document.createElement('li'); li2.className = 'nativo'; li2.textContent = '✓ ' + nDefs + ' comando(s) definido(s): ' + Object.keys(p.defs).map(function (c) { return '\\' + c; }).join(', '); ul.appendChild(li2); }
  }
  async function loadPreamble() {
    try {
      var r = await oc(function (cb) { Office.context.document.customXmlParts.getByNamespaceAsync(PRENS, cb); });
      if (r.status === 'succeeded' && r.value.length) {
        var x = await oc(function (cb) { r.value[0].getXmlAsync(cb); });
        if (x.status === 'succeeded') S.preamble = window.IHRLX.xml.text(window.IHRLX.xml.parse(x.value));
      }
    } catch (e) { /* sem preâmbulo */ }
    $('preText').value = S.preamble; renderPreamble();
  }
  async function savePreamble() {
    S.preamble = $('preText').value;
    var xml = '<preambulo xmlns="' + PRENS + '">' + window.IHRLX.xml.esc(S.preamble) + '</preambulo>';
    var r = await oc(function (cb) { Office.context.document.customXmlParts.getByNamespaceAsync(PRENS, cb); });
    if (r.status === 'succeeded') for (var i = 0; i < r.value.length; i++) await oc(function (cb) { r.value[i].deleteAsync(cb); });
    if (S.preamble.trim()) {
      var a = await oc(function (cb) { Office.context.document.customXmlParts.addAsync(xml, cb); });
      if (a.status !== 'succeeded') throw new Error('não foi possível guardar no documento');
    }
  }
  var preT = null;
  $('preText').addEventListener('input', function () { clearTimeout(preT); preT = setTimeout(renderPreamble, 300); });
  $('btnPreSave').addEventListener('click', async function () {
    S.preamble = $('preText').value; renderPreamble();
    try {
      if (typeof Office !== 'undefined' && Office.context && Office.context.document) await savePreamble();
      status('Pacotes e comandos salvos no documento: valem para todo o texto do painel.', 'ok');
    } catch (e) { status('Salvo só no painel (não consegui gravar no documento): ' + errMsg(e), 'warn'); }
  });
  // texto do editor (ou de um bloco) com os comandos expandidos
  function editorDefs() { return MAC.prepare(ed.value, S.preamble).defs; }
  function expandTex(tex, defs) { return MAC.expand(MAC.parseDefs(tex).rest, defs || editorDefs()); }

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
    var paras = ctx.document.body.paragraphs; paras.load('items/text,styleBuiltIn,style');
    await ctx.sync();
    var items = paras.items, h = -1;
    for (var i = items.length - 1; i >= 0; i--) {
      if ((items[i].styleBuiltIn === 'Heading1' || isUnnum(items[i])) && /^REFER[EÊ]NCIAS$/i.test((items[i].text || '').trim())) { h = i; break; }
    }
    if (h < 0) {
      if (!createIfMissing || !keys.length) return keys.length ? 'Não encontrei o capítulo "Referências" no documento.' : 'Nenhuma citação do .bib no documento.';
      var res0 = L.convertBibList(convOpts({ bib: S.lib, keys: keys, manual: [], withHeading: true }));
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
    var res = L.convertBibList(convOpts({ bib: S.lib, keys: keys, manual: manual }));
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
  function primary(ev) { if (S.bound && ev && ev.shiftKey) S.forceAll = true; return S.bound ? updateBound(false) : insertAtCursor(); }

  async function insertAtCursor(force) {
    if (needWord()) return;
    if (!ed.value.trim()) { status('O editor está vazio.', 'warn'); return; }
    setBusy(true); status('Convertendo…');
    try {
      await Word.run(async function (ctx) {
        var kinds = await labelKinds(ctx);
        var sel0 = ctx.document.getSelection();
        var nums = await numbersBefore(ctx, sel0), rt = await refTexts(ctx, ed.value, kinds);
        if (nums && S.profile.equacaoPorCapitulo) delete nums['Equação'];      // numeradas por capítulo: o Word recalcula
        var c = convertEditor(kinds, await citedKeys(ctx), { startNumbers: nums || undefined, labelText: rt });
        if (!force && c.res.missingImages.length) {
          var miss = c.res.missingImages.filter(function (v, i, a) { return a.indexOf(v) === i; });
          notice('Faltam no painel as imagens ' + miss.join(', ') + '. Carregue-as em "Imagens" e insira de novo; ' +
            'se inserir assim, entra um marcador [imagem: …] no lugar (ele volta a ser \\includegraphics quando o trecho for lido).',
            [{ label: 'Inserir mesmo assim', run: function () { notice(''); insertAtCursor(true); } }]);
          $('imgBox').open = true; status(''); return;
        }
        var rng = ctx.document.getSelection().insertOoxml(c.pkg, 'Replace');
        await ctx.sync();
        await restyleLast(ctx, rng, c.res.body);
        var ok = c.res.fields ? await updateFields(ctx, c.res) : true;
        var bibMsg = c.res.cites && !c.res.hasBib && S.lib && S.lib.count() ? await regenBib(ctx, true) : '';
        if (api('1.4')) await bindTo(ctx, rng, ed.value);
        notice(''); renderMode(); refreshOutline();
        finish(c.res, (S.bound ? 'Inserido. O trecho continua vinculado: Ctrl+Enter atualiza ele.' : 'Inserido no documento.') + (bibMsg ? ' ' + bibMsg : ''), ok);
      });
    } catch (e) {
      status('Não foi possível inserir: ' + errMsg(e), 'err');
    } finally { setBusy(false); }
  }

  // O último parágrafo inserido se funde com o parágrafo de destino e fica com o estilo dele
  // (ex.: a \fonte de uma tabela inserida numa linha vazia virava Normal). Reaplica o estilo
  // e o alinhamento que o conversor gerou para esse parágrafo.
  var BUILTIN = { 'normal': 'Normal', 'heading 1': 'Heading1', 'heading 2': 'Heading2', 'heading 3': 'Heading3', 'heading 4': 'Heading4',
    'heading 5': 'Heading5', 'heading 6': 'Heading6', 'caption': 'Caption', 'bibliography': 'Bibliography', 'list paragraph': 'ListParagraph',
    'toc heading': 'TocHeading', 'footnote text': 'FootnoteText' };
  var JC = { center: 'Centered', both: 'Justified', left: 'Left', right: 'Right' };
  function lastParaSpec(body) {
    if (!body || !/<\/w:p>\s*$/.test(body)) return null;
    var i = body.lastIndexOf('<w:p>');
    if (i < 0 || body.indexOf('</w:tbl>', i) >= 0) return null;
    var p = body.slice(i), ppr = (/^<w:p><w:pPr>([\s\S]*?)<\/w:pPr>/.exec(p) || [])[1] || '';
    var id = (/<w:pStyle w:val="([^"]+)"/.exec(ppr) || [])[1] || 'Normal';
    var st = (window.IHRLX.template.styles.match(new RegExp('<w:style [^>]*w:styleId="' + id + '"[\\s\\S]*?<\\/w:style>')) || [''])[0];
    var name = ((/<w:name w:val="([^"]+)"/.exec(st) || [])[1] || id);
    var jc = (/<w:jc w:val="([^"]+)"/.exec(ppr) || [])[1];
    var text = p.replace(/<w:instrText[\s\S]*?<\/w:instrText>/g, '').replace(/<[^>]+>/g, '');
    return { name: name, builtin: BUILTIN[name.toLowerCase()], jc: jc, text: text };
  }
  async function restyleLast(ctx, rng, body) {
    var sp = lastParaSpec(body);
    if (!sp || !rng) return;
    try {
      var lp = rng.paragraphs.getLast(); lp.load('text,style,styleBuiltIn');
      await ctx.sync();
      if (paraKey(lp.text) !== paraKey(sp.text)) return;          // não é o parágrafo esperado: não mexe
      if (sp.builtin) { if (lp.styleBuiltIn !== sp.builtin) lp.styleBuiltIn = sp.builtin; }
      else if (lp.style !== sp.name) lp.style = sp.name;
      if (sp.jc && JC[sp.jc]) lp.alignment = JC[sp.jc];
      await ctx.sync();
    } catch (e) { /* estilo inexistente ou API antiga: segue */ }
  }

  // ---- diferença por blocos: só os blocos alterados são reescritos no documento
  function lcsPairs(a, b) {
    var n = a.length, m = b.length, i, j;
    var dp = []; for (i = 0; i <= n; i++) { dp.push(new Uint16Array(m + 1)); }
    for (i = n - 1; i >= 0; i--) for (j = m - 1; j >= 0; j--) dp[i][j] = a[i] === b[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
    var pairs = []; i = 0; j = 0;
    while (i < n && j < m) {
      if (a[i] === b[j]) { pairs.push([i, j]); i++; j++; }
      else if (dp[i + 1][j] >= dp[i][j + 1]) i++; else j++;
    }
    return pairs;
  }
  function paraKey(s) { return (s || '').replace(/[^\p{L}\p{N}]/gu, '').toLowerCase(); }
  // mapa índice-no-pacote → índice-na-coleção de parágrafos do Word (-1 = sem par)
  function mapParas(A, B) {
    var a = A.map(function (s, i) { return paraKey(s) || '\u0001' + i; });
    var b = B.map(function (s, j) { return paraKey(s) || '\u0002' + j; });
    var map = A.map(function () { return -1; });
    var anchors = [[-1, -1]].concat(lcsPairs(a, b)).concat([[A.length, B.length]]);
    for (var k = 1; k < anchors.length; k++) {
      var p = anchors[k - 1], c = anchors[k];
      if (c[0] < A.length) map[c[0]] = c[1];
      var ga = c[0] - p[0] - 1, gb = c[1] - p[1] - 1;
      if (ga === gb) for (var t = 1; t <= ga; t++) map[p[0] + t] = p[1] + t;     // trechos sem texto em comum, mesmo tamanho
    }
    return map;
  }
  function makeHunks(oldN, newN, pairs) {
    var h = [], pi = 0, pj = 0;
    pairs.concat([[oldN, newN]]).forEach(function (p) {
      if (p[0] > pi || p[1] > pj) h.push({ i0: pi, i1: p[0], k0: pj, k1: p[1] });
      pi = p[0] + 1; pj = p[1] + 1;
    });
    return h;
  }
  async function applyDiff(ctx, region, cur, force) {
    var paras = region.paragraphs; paras.load('items/tableNestingLevel,text');
    await ctx.sync();
    var P = paras.items;
    var newB = O.splitBlocks(ed.value);
    // o pacote OOXML do Word pode ter parágrafos a mais/a menos que a coleção (ex.: parágrafo vazio no fim);
    // relaciona os dois pelo texto e posiciona cada bloco nos parágrafos reais do documento
    var map = mapParas(cur.paraTexts || [], P.map(function (p) { return p.text || ''; }));
    var oldB = [], lost = [];
    cur.blocks.forEach(function (b) {
      var a = -1, z = -1;
      for (var q = b.p0; q <= b.p1; q++) if (map[q] >= 0) { if (a < 0) a = map[q]; z = map[q]; }
      if (a < 0) lost.push(b.tex.split('\n')[0]);
      oldB.push(Object.assign({}, b, { p0: a, p1: z }));
    });
    if (lost.length) {
      notice('Não consegui localizar no documento estes blocos do trecho: ' + lost.slice(0, 3).join(' | ') + '. Clique em "Recarregar do Word" e tente de novo.',
        [{ label: 'Recarregar do Word', run: function () { reloadBound(); } }]);
      status(''); return null;
    }
    // blocos "stale" (formato antigo) nunca casam, para serem reescritos com o estilo atual
    // Shift ao aplicar: reescreve todos os blocos (útil depois de uma atualização do painel)
    var all = S.forceAll; S.forceAll = false;
    var pairs = all ? [] : lcsPairs(oldB.map(function (b, i) { return b.stale ? '\u0001' + i : norm(b.tex); }), newB.map(norm));
    if (all) { var lostF = oldB.filter(function (b) { return b.fixed; }); if (lostF.length) { notice('Este trecho tem sumário ou quebra de seção, que não podem ser reescritos. Selecione só o trecho com as equações e use Shift + Atualizar de novo.'); status(''); return null; } }
    var matched = {}; pairs.forEach(function (p) { matched[p[0]] = true; });
    var lostFixed = oldB.filter(function (b, i) { return b.fixed && !matched[i]; });
    if (lostFixed.length) {
      notice('Estes blocos não podem ser alterados pelo painel (sumário e quebras de seção do Word). Deixe as linhas exatamente como estavam: ' +
        lostFixed.map(function (b) { return b.tex.split('\n')[0]; }).join(' | '), [{ label: 'Recarregar do Word', run: function () { reloadBound(); } }]);
      status(''); return null;
    }
    var hunks = makeHunks(oldB.length, newB.length, pairs);
    if (!hunks.length) return { changed: 0 };
    var kinds = await labelKinds(ctx), cited = await citedKeys(ctx);
    var base = S.profile.equacaoPorCapitulo ? null : await numbersBefore(ctx, region), rt = await refTexts(ctx, ed.value, kinds);
    var defsAll = editorDefs();
    var miss = [], unsafeHit = false, warnings = [], fields = 0, cites = 0, hasBib = false, blocks = [];
    hunks.forEach(function (h) {
      var tex = newB.slice(h.k0, h.k1).join('\n\n');
      var start = base ? addCounts(base, countLatex(newB.slice(0, h.k0).join('\n\n'))) : undefined;
      var res = L.convert(expandTex(tex, defsAll), convOpts({ images: S.images, labelKinds: kinds, bib: S.lib, citedKeys: cited, startNumbers: start, labelText: rt }));
      h.pkg = res.body ? L.buildPackage(res) : null;
      h.body = res.body;
      miss = miss.concat(res.missingImages);
      res.warnings.forEach(function (w) { if (warnings.indexOf(w) < 0) warnings.push(w); });
      fields += res.fields; cites += res.cites; hasBib = hasBib || res.hasBib; blocks = blocks.concat(res.blocks);
      for (var q = h.i0; q < h.i1; q++) if (oldB[q].unsafe) unsafeHit = true;
    });
    if (!force && (miss.length || unsafeHit)) {
      var why = [];
      if (miss.length) why.push('faltam no painel as imagens ' + miss.join(', ') + ' (elas virariam marcadores de texto)');
      if (unsafeHit) why.push('um bloco alterado contém ' + S.unsafe.join(', ') + ', que seriam perdidos');
      notice('Atenção: ' + why.join('; ') + '.', [{ label: 'Atualizar mesmo assim', run: function () { notice(''); updateBound(true); } }]);
      if (miss.length) $('imgBox').open = true;
      status(''); return null;
    }
    function edge(i, which) {      // parágrafo i (ou a tabela inteira que o contém)
      if (P[i].tableNestingLevel > 0) return P[i].parentTableOrNullObject.getRange('Whole');
      return P[i].getRange('Whole');
    }
    // aplica de trás para frente, para os índices anteriores continuarem válidos
    var inserted = {};
    for (var x = hunks.length - 1; x >= 0; x--) {
      var h = hunks[x];
      if (h.i1 > h.i0) {
        var rg = edge(oldB[h.i0].p0).expandTo(edge(oldB[h.i1 - 1].p1));
        if (h.pkg) inserted[x] = rg.insertOoxml(h.pkg, 'Replace'); else rg.delete();
      } else if (h.pkg) {
        if (h.i0 > 0) inserted[x] = edge(oldB[h.i0 - 1].p1).insertOoxml(h.pkg, 'After');
        else if (oldB.length) inserted[x] = edge(oldB[0].p0).insertOoxml(h.pkg, 'Before');
        else inserted[x] = region.insertOoxml(h.pkg, 'Replace');
      }
    }
    await ctx.sync();
    for (var y = 0; y < hunks.length; y++) if (inserted[y]) await restyleLast(ctx, inserted[y], hunks[y].body);
    // novo intervalo do trecho: primeira e última peça que sobraram/entraram
    var pieces = [], used = {};
    var replacedAt = {}, afterIdx = {};
    hunks.forEach(function (h, x) {
      if (h.i1 > h.i0) { for (var q = oldB[h.i0].p0; q <= oldB[h.i1 - 1].p1; q++) replacedAt[q] = x; }
      else { var a0 = h.i0 > 0 ? oldB[h.i0 - 1].p1 : -1; (afterIdx[a0] = afterIdx[a0] || []).push(x); }
    });
    (afterIdx[-1] || []).forEach(function (x) { if (inserted[x]) pieces.push(inserted[x]); });
    for (var q2 = 0; q2 < P.length; q2++) {
      if (replacedAt[q2] !== undefined) {
        var hx = replacedAt[q2];
        if (!used[hx]) { used[hx] = 1; if (inserted[hx]) pieces.push(inserted[hx]); }
      } else pieces.push(P[q2]);
      (afterIdx[q2] || []).forEach(function (x) { if (inserted[x]) pieces.push(inserted[x]); });
    }
    var reg = pieces.length ? pieces[0].getRange('Whole').expandTo(pieces[pieces.length - 1].getRange('Whole')) : region;
    return { changed: hunks.length, region: reg, warnings: warnings, fields: fields, cites: cites, hasBib: hasBib, blocks: blocks };
  }

  async function updateBound(force) {
    if (needWord()) return;
    setBusy(true); status('Atualizando…');
    try {
      await Word.run(async function (ctx) {
        var b = await boundRange(ctx);
        if (!b) { unbindState(false); status('O trecho vinculado não existe mais no documento. Selecione de novo ou use "Inserir no cursor".', 'err'); return; }
        // o trecho mudou no Word desde a leitura? (você ou um colega editou)
        var cur = await readRange(ctx, b);
        if (!force && norm(cur.latex) !== norm(S.base)) {
          S.forceAll = false;
          notice('Este trecho foi alterado no Word depois que foi lido no painel (por você ou por outra pessoa). Atualizar agora apagaria essas alterações.', [
            { label: 'Recarregar do Word', primary: true, run: function () { reloadBound(); } },
            { label: 'Sobrescrever', run: function () { notice(''); updateBound(true); } }
          ]);
          status('');
          return;
        }
        var d = await applyDiff(ctx, b, cur, force);
        if (!d) return;                                   // aviso já mostrado
        if (!d.changed) { status('Nada mudou em relação ao documento.', 'ok'); return; }
        var ok = d.fields ? await updateFields(ctx, { blocks: d.blocks }) : true;
        var bibMsg = d.cites && !d.hasBib && S.lib && S.lib.count() ? await regenBib(ctx, true) : '';
        await bindTo(ctx, d.region, ed.value);
        notice(''); renderMode(); refreshOutline();
        finish({ warnings: d.warnings, fields: d.fields }, 'Trecho atualizado: ' + d.changed + ' bloco(s) reescrito(s), o resto ficou intocado.' + (bibMsg ? ' ' + bibMsg : ''), ok);
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
        var r = await readRange(ctx, range);
        await fillImages(ctx, range, r);
        if (diag) {
          setEditor(['% Diagnóstico da leitura', '% parágrafos na seleção: ' + n + ' | no trecho: ' + nParas + ' | no OOXML: ' + r.paraCount + (trimmed ? ' | cortado na seção' : ''),
            '% quebras de seção no OOXML: ' + JSON.stringify(r.sectBreaks), '% partes: ' + r.partNames.join(', '),
            '% imagens: ' + JSON.stringify(r.imageOrder), '% avisos: ' + JSON.stringify(r.warnings)].join('\n') + '\n\n' + r.latex);
          status('Diagnóstico gerado (nada foi vinculado).', 'ok');
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
        notice('');
        var w = r.warnings.slice();
        var miss = (r.imageOrder || []).filter(function (x) { return x.missing; }).map(function (x) { return x.name; });
        if (miss.length) w.unshift('Não consegui ler ' + miss.length + ' imagem(ns) do documento (' + miss.join(', ') + '). Não atualize o trecho sem carregá-las em "Imagens", ou elas serão trocadas por marcadores.');
        if (r.sectBreaks && r.sectBreaks.length) w.unshift('O trecho tem quebra(s) de seção: elas aparecem como uma linha de comentário (%) e ficam intactas no documento.');
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

  // ------------------------------------------------------- estrutura (árvore de títulos)
  var OUT = [];
  async function scanOutline(ctx) {
    var ps = ctx.document.body.paragraphs; ps.load('items/text,styleBuiltIn,style');
    await ctx.sync();
    var heads = [];
    ps.items.forEach(function (p, i) {
      var m = /^Heading([1-6])$/.exec(p.styleBuiltIn || '') || (isUnnum(p) ? [0, '1'] : null);
      if (m && (p.text || '').trim()) heads.push({ idx: i, level: parseInt(m[1], 10), text: p.text.trim() });
    });
    var lis = heads.map(function (h) { var li = ps.items[h.idx].listItemOrNullObject; li.load('listString'); return li; });
    if (lis.length) await ctx.sync();
    heads.forEach(function (h, k) { h.num = lis[k] && !lis[k].isNullObject ? (lis[k].listString || '') : ''; });
    return { ps: ps.items, heads: heads };
  }
  function renderOutline() {
    var ul = $('outList'); ul.innerHTML = '';
    function add(label, cls, k, title) {
      var li = document.createElement('li'); li.className = cls; li.textContent = label; li.title = title || label;
      li.onclick = function () { openSection(k); };
      ul.appendChild(li);
    }
    add('Documento inteiro', 'lv0 all', -1, 'Vincula o documento todo (a capa, o sumário e as quebras de seção ficam protegidos)');
    if (OUT.pre) add('Capa e pré-textuais', 'lv0', -2);
    OUT.forEach(function (h, k) { add((h.num ? h.num + '  ' : '') + h.text, 'lv' + h.level, k); });
  }
  async function refreshOutline() {
    if (typeof Word === 'undefined') return;
    try {
      await Word.run(async function (ctx) {
        var o = await scanOutline(ctx);
        OUT = o.heads; OUT.pre = o.heads.length ? o.heads[0].idx > 0 : o.ps.length > 0;
      });
      renderOutline();
    } catch (e) { /* sem estrutura */ }
  }
  async function openSection(k) {
    if (dirty()) {
      notice('Há edições no painel ainda não aplicadas ao trecho vinculado.', [
        { label: 'Aplicar (Ctrl+Enter)', primary: true, run: function () { updateBound(false); } },
        { label: 'Descartar e abrir a seção', run: function () { S.clean = ed.value; openSection(k); } }
      ]);
      return;
    }
    try {
      await Word.run(async function (ctx) {
        var o = await scanOutline(ctx);
        OUT = o.heads; OUT.pre = o.heads.length ? o.heads[0].idx > 0 : o.ps.length > 0;
        var ps = o.ps, a = 0, z = ps.length - 1;
        if (k === -2) z = o.heads.length ? o.heads[0].idx - 1 : z;
        else if (k >= 0) {
          var h = o.heads[k]; if (!h) return;
          a = h.idx;
          for (var q = k + 1; q < o.heads.length; q++) if (o.heads[q].level <= h.level) { z = o.heads[q].idx - 1; break; }
        }
        if (z < a) return;
        S.suppressUntil = Date.now() + 1500;
        ps[a].getRange('Whole').expandTo(ps[z].getRange('Whole')).select();
        await ctx.sync();
      });
      renderOutline();
      await readSelection(false);
    } catch (e) { status('Não foi possível abrir a seção: ' + errMsg(e), 'err'); }
  }
  $('btnOutline').addEventListener('click', refreshOutline);

  // ------------------------------------------------------- botões
  $('btnBibAdd').addEventListener('click', async function () {
    var txt = $('bibPaste').value.trim();
    if (!txt) { status('Cole uma entrada BibTeX na caixa (começando com @article{…, @book{… etc.).', 'warn'); return; }
    if (!S.lib) S.lib = new BIB.Library('');
    var m = S.lib.merge(txt);
    if (!m.added && !m.updated) { status('Não reconheci nenhuma entrada BibTeX no texto colado. Confira se começa com @tipo{chave, e se as chaves { } fecham.', 'warn'); return; }
    $('bibPaste').value = ''; renderBib(); $('bibBox').open = true;
    try {
      if (typeof Office !== 'undefined' && Office.context && Office.context.document) await saveLibrary();
      status(m.added + ' entrada(s) nova(s), ' + m.updated + ' atualizada(s). Clique numa entrada da lista para inserir a citação.', 'ok');
    } catch (e) { status('Entradas adicionadas no painel, mas não consegui guardá-las no documento: ' + errMsg(e), 'warn'); }
  });
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
        loadProfile();
        loadLibrary();
        loadPreamble();
        Word.run(function (ctx) { return labelKinds(ctx); }).catch(function () {});
        refreshOutline();
        // remove um vínculo esquecido de uma sessão anterior
        Word.run(async function (ctx) { ctx.document.deleteBookmark(BM); await ctx.sync(); }).catch(function () {});
      }
    });
  }
  $('ver').textContent = 'v' + VERSION + (/\/stable\//.test(location.pathname) ? '' : ' · canal de teste');
  renderImages();
  renderMode();
  syncHL();
})();
