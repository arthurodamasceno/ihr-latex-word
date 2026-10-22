/* LaTeX (subconjunto) -> OOXML com os estilos do template IHR */
(function (root) {
  'use strict';
  var R = root.IHRLX || {};
  var X = R.xml || require('./xml.js');
  var M = R.math || require('./math.js');
  var T = R.template || require('./template-parts.js');
  var esc = X.esc;

  var NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ' +
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
  var REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
  var TEXTWIDTH_TW = 9061;          // largura útil do template (16 cm)
  var CM_TW = 566.93, CM_EMU = 360000;

  // ======================================================== utilidades de texto
  function stripComments(s) {
    return s.split('\n').map(function (line) {
      var out = '', i = 0;
      for (; i < line.length; i++) {
        if (line[i] === '\\') { out += line[i] + (line[i + 1] || ''); i++; continue; }
        if (line[i] === '%') return { t: out, c: true };
        out += line[i];
      }
      return { t: out, c: false };
    }).reduce(function (acc, l, idx, arr) {
      // um comentário "come" a quebra de linha (como no TeX)
      return acc + l.t + (l.c ? '' : (idx < arr.length - 1 ? '\n' : ''));
    }, '');
  }

  // lê {grupo} a partir de s[i] === '{'; retorna {content, end}
  function readGroup(s, i) {
    if (s[i] !== '{') return null;
    var depth = 0, j = i;
    for (; j < s.length; j++) {
      var c = s[j];
      if (c === '\\') { j++; continue; }
      if (c === '{') depth++;
      else if (c === '}') { depth--; if (depth === 0) return { content: s.slice(i + 1, j), end: j + 1 }; }
    }
    return { content: s.slice(i + 1), end: s.length };
  }
  function readOpt(s, i) {
    if (s[i] !== '[') return null;
    var depth = 0, j = i;
    for (; j < s.length; j++) {
      var c = s[j];
      if (c === '\\') { j++; continue; }
      if (c === '{') depth++;
      else if (c === '}') depth--;
      else if (c === ']' && depth === 0) return { content: s.slice(i + 1, j), end: j + 1 };
    }
    return null;
  }
  function skipWs(s, i) { while (i < s.length && /[ \t\r\n]/.test(s[i])) i++; return i; }
  function skipSp(s, i) { while (i < s.length && /[ \t]/.test(s[i])) i++; return i; }
  // lê argumento obrigatório (grupo ou token único)
  function readArg(s, i) {
    i = skipWs(s, i);
    if (s[i] === '{') return readGroup(s, i);
    if (s[i] === '\\') { var m = /^\\([a-zA-Z]+|.)/.exec(s.slice(i)); return { content: m[0], end: i + m[0].length }; }
    return { content: s[i] || '', end: i + 1 };
  }
  function cmdAt(s, i) {
    if (s[i] !== '\\') return null;
    var m = /^\\([a-zA-Z@]+\*?|.)/.exec(s.slice(i));
    return m ? m[1] : null;
  }
  // encontra \end{env} correspondente a partir de i (logo após \begin{env})
  function findEnd(s, i, env) {
    var re = new RegExp('\\\\(begin|end)\\{' + env.replace(/\*/g, '\\*') + '\\}', 'g');
    re.lastIndex = i;
    var depth = 1, m;
    while ((m = re.exec(s))) {
      if (m[1] === 'begin') depth++; else depth--;
      if (depth === 0) return { contentEnd: m.index, end: re.lastIndex };
    }
    return { contentEnd: s.length, end: s.length };
  }
  // divide no nível superior (fora de chaves e ambientes) por um separador regex
  function splitTop(s, sepRe) {
    var parts = [], depth = 0, envDepth = 0, last = 0, i = 0;
    while (i < s.length) {
      var c = s[i];
      if (c === '\\') {
        if (s.startsWith('\\begin{', i)) envDepth++;
        else if (s.startsWith('\\end{', i)) envDepth--;
        if (depth === 0 && envDepth === 0) {
          sepRe.lastIndex = i;
          var m = sepRe.exec(s);
          if (m && m.index === i) { parts.push(s.slice(last, i)); i += m[0].length; last = i; continue; }
        }
        i += 2; continue;
      }
      if (c === '{') depth++;
      else if (c === '}') depth--;
      else if (depth === 0 && envDepth === 0) {
        sepRe.lastIndex = i;
        var m2 = sepRe.exec(s);
        if (m2 && m2.index === i) { parts.push(s.slice(last, i)); i += m2[0].length; last = i; continue; }
      }
      i++;
    }
    parts.push(s.slice(last));
    return parts;
  }

  // ========================================================= rótulos/bookmarks
  function encLabel(l) {
    return l.replace(/[^A-Za-z0-9]/g, function (ch) { return '_' + ch.charCodeAt(0).toString(16).toUpperCase().padStart(2, '0'); });
  }
  function bookmarkName(label, kind) {
    if (/^_Ref\d+$/.test(label) || /^_Ref[SNE]_/.test(label)) return label;
    var n = '_Ref' + kind + '_' + encLabel(label);
    if (n.length > 40) {
      var h = 0; for (var i = 0; i < label.length; i++) h = (h * 31 + label.charCodeAt(i)) >>> 0;
      n = n.slice(0, 31) + h.toString(36).slice(0, 8);
    }
    return n;
  }

  // ================================================================ inline
  var ACC = { "'": '́', '`': '̀', '^': '̂', '~': '̃', '"': '̈', 'c': '̧', '=': '̄',
    '.': '̇', 'u': '̆', 'v': '̌', 'H': '̋', 'r': '̊', 'k': '̨' };
  var SIMPLE = { '%': '%', '&': '&', '$': '$', '#': '#', '_': '_', '{': '{', '}': '}', ' ': ' ', ',': ' ',
    ';': ' ', ':': ' ', '!': '', '-': '', '/': '', '@': '', '|': '‖',
    'textbackslash': '\\', 'textasciitilde': '~', 'textasciicircum': '^', 'ldots': '…', 'dots': '…', 'textellipsis': '…',
    'LaTeX': 'LaTeX', 'TeX': 'TeX', 'LaTeXe': 'LaTeX2e', 'S': '§', 'P': '¶', 'copyright': '©', 'textregistered': '®',
    'texttrademark': '™', 'textdegree': '°', 'degree': '°', 'textbullet': '•', 'textendash': '–', 'textemdash': '—',
    'quad': ' ', 'qquad': '  ', 'enspace': ' ', 'thinspace': ' ', 'i': 'ı', 'j': 'ȷ',
    'o': 'ø', 'O': 'Ø', 'ss': 'ß', 'ae': 'æ', 'AE': 'Æ', 'oe': 'œ', 'OE': 'Œ', 'aa': 'å', 'AA': 'Å', 'l': 'ł', 'L': 'Ł',
    'euro': '€', 'textmu': 'µ', 'textpm': '±', 'texttimes': '×', 'guillemotleft': '«', 'guillemotright': '»',
    'noindent': '', 'centering': '', 'raggedright': '', 'raggedleft': '', 'par': '', 'indent': '', 'protect': '',
    'small': '', 'footnotesize': '', 'scriptsize': '', 'tiny': '', 'large': '', 'Large': '', 'LARGE': '', 'huge': '',
    'Huge': '', 'normalsize': '', 'selectfont': '', 'hfill': '', 'vfill': '', 'null': '', 'relax': '', 'xspace': '',
    'smallskip': '', 'medskip': '', 'bigskip': '', 'clearpage': '', 'linebreak': '' };
  var DECL = { bfseries: ['b', true], mdseries: ['b', false], itshape: ['i', true], slshape: ['i', true], upshape: ['i', false],
    em: ['i', 'toggle'], ttfamily: ['tt', true], rmfamily: ['tt', false], sffamily: ['tt', false], scshape: ['sc', true] };
  var WRAP = { textbf: ['b', true], textit: ['i', true], textsl: ['i', true], emph: ['i', 'toggle'], underline: ['u', true],
    uline: ['u', true], texttt: ['tt', true], textsc: ['sc', true], textsuperscript: ['sup', true], textsubscript: ['sub', true],
    sout: ['strike', true], textup: ['i', false], textmd: ['b', false], textrm: [null], textsf: [null], textnormal: ['reset'],
    mbox: [null], text: [null], hbox: [null], makebox: [null], fbox: [null], cab: ['cab', true], textcolor: ['color2'],
    MakeUppercase: ['upper', true], uppercase: ['upper', true] };
  var IGNORE_1ARG = { label: 1, index: 1, vspace: 1, 'vspace*': 1, hspace: 1, 'hspace*': 1, phantom: 1, hyphenation: 1,
    pagestyle: 1, thispagestyle: 1, setlength: 2, addtolength: 2, setcounter: 2, addtocounter: 2, color: 1, rule: 2,
    addcontentsline: 3, includeonly: 1, bibliographystyle: 1, nocite: 1, glsresetall: 0 };

  function Ctx(opts) {
    this.opts = opts || {};
    this.images = this.opts.images || {};
    this.labelKinds = Object.assign({}, this.opts.labelKinds || {});
    this.labelText = {};
    this.footnotes = [];
    this.rels = [];
    this.media = [];
    this.warnings = [];
    this.nums = [];              // instâncias de numeração de listas
    this.bmId = 7000 + Math.floor(Math.random() * 20000);
    this.docPrId = 3000 + Math.floor(Math.random() * 20000);
    this.counters = Object.assign({ Figura: 0, Tabela: 0, Quadro: 0, 'Equação': 0 }, this.opts.startNumbers || {});
    this.fields = 0;
  }
  Ctx.prototype.warn = function (m) { if (this.warnings.indexOf(m) < 0) this.warnings.push(m); };
  Ctx.prototype.rel = function (type, target, external) {
    var id = 'rIdL' + (this.rels.length + 1);
    this.rels.push({ id: id, type: type, target: target, external: !!external });
    return id;
  };

  function fmtCopy(f) { return Object.assign({}, f); }

  // retorna lista de "itens" inline
  function parseInline(s, ctx, fmt) {
    fmt = fmt || {};
    var out = [];
    var buf = '';
    function flush() { if (buf) { out.push({ t: 'text', text: buf, fmt: fmtCopy(fmt) }); buf = ''; } }
    var i = 0;
    while (i < s.length) {
      var c = s[i];
      if (c === '\\') {
        var name = cmdAt(s, i);
        if (name === null) { i++; continue; }
        var j = i + 1 + name.length;
        var bare = name.replace(/\*$/, '');
        // escapes simples
        if (name.length === 1 && SIMPLE.hasOwnProperty(name)) { buf += SIMPLE[name]; i = j; continue; }
        if (name === '\\') { // quebra de linha
          flush(); out.push({ t: 'br' });
          var jj = skipSp(s, j); if (s[jj] === '*') jj++; var o0 = readOpt(s, jj); i = o0 ? o0.end : j; continue;
        }
        if (name.length === 1 && ACC[name] !== undefined) {
          var a = readArg(s, j);
          var base = a.content.replace(/^\\i$/, 'i').replace(/^\\j$/, 'j');
          buf += (base + ACC[name]).normalize('NFC'); i = a.end; continue;
        }
        if (ACC[bare] !== undefined && /^[a-zA-Z]$/.test(bare) && s[skipSp(s, j)] === '{') { // \c{c}, \u{a}
          var a2 = readGroup(s, skipSp(s, j)); buf += (a2.content + ACC[bare]).normalize('NFC'); i = a2.end; continue;
        }
        if (SIMPLE.hasOwnProperty(bare)) {
          buf += SIMPLE[bare];
          i = j;
          if (/^[a-zA-Z]+$/.test(bare)) i = skipWs(s, i);      // o TeX "come" o espaço após uma palavra de controle
          if (s.startsWith('{}', i)) i += 2;
          continue;
        }
        if (bare === 'newline') { flush(); out.push({ t: 'br' }); i = j; continue; }
        if (bare === 'newpage' || bare === 'pagebreak') { flush(); out.push({ t: 'pagebreak' }); i = j; continue; }
        if (bare === 'today') {
          var d = new Date(); buf += d.toLocaleDateString('pt-BR', { day: 'numeric', month: 'long', year: 'numeric' }); i = j; continue;
        }
        if (bare === '(') { // \( math \)
          var e1 = s.indexOf('\\)', j); if (e1 < 0) e1 = s.length;
          flush(); out.push({ t: 'math', tex: s.slice(j, e1) }); i = e1 + 2; continue;
        }
        if (DECL[bare]) {
          var dd = DECL[bare];
          fmt = fmtCopy(fmt);
          flush();
          fmt[dd[0]] = dd[1] === 'toggle' ? !fmt[dd[0]] : dd[1];
          i = skipSp(s, j); continue;
        }
        if (bare === 'normalfont') { flush(); fmt = {}; i = skipSp(s, j); continue; }
        if (WRAP[bare]) {
          var w = WRAP[bare];
          var jx = j;
          if (bare === 'textcolor') { var cc = readArg(s, jx); jx = cc.end; }
          if (bare === 'makebox' || bare === 'fbox') { var o1 = readOpt(s, skipWs(s, jx)); while (o1) { jx = o1.end; o1 = readOpt(s, skipWs(s, jx)); } }
          var g = readArg(s, jx);
          var f2 = fmtCopy(fmt);
          if (w[0] === 'reset') f2 = {};
          else if (w[0] === 'cab') { f2.b = true; f2.upper = true; }
          else if (w[0] && w[0] !== 'color2') f2[w[0]] = w[1] === 'toggle' ? !fmt[w[0]] : w[1];
          flush();
          out = out.concat(parseInline(g.content, ctx, f2));
          i = g.end; continue;
        }
        if (bare === 'ref' || bare === 'eqref' || bare === 'pageref' || bare === 'autoref' || bare === 'cref' || bare === 'Cref' || bare === 'nameref') {
          var g2 = readArg(s, j); flush();
          out.push({ t: 'ref', label: g2.content.trim(), kind: bare === 'eqref' ? 'eqref' : bare === 'pageref' ? 'pageref' : 'ref', fmt: fmtCopy(fmt) });
          i = g2.end; continue;
        }
        if (bare === 'footnote') {
          var o2 = readOpt(s, j); var g3 = readArg(s, o2 ? o2.end : j); flush();
          out.push({ t: 'footnote', src: g3.content }); i = g3.end; continue;
        }
        if (bare === 'url') { var g4 = readArg(s, j); flush(); out.push({ t: 'link', url: g4.content, text: g4.content }); i = g4.end; continue; }
        if (bare === 'href') {
          var g5 = readArg(s, j); var g6 = readArg(s, g5.end); flush();
          out.push({ t: 'link', url: g5.content, inner: parseInline(g6.content, ctx, fmt) }); i = g6.end; continue;
        }
        if (bare === 'cite' || bare === 'citep' || bare === 'citet' || bare === 'parencite' || bare === 'textcite') {
          var o3 = readOpt(s, j); var g7 = readArg(s, o3 ? o3.end : j);
          buf += '[' + g7.content + ']'; ctx.warn('\\cite ainda não é convertido em citação do Word; ficou como texto [' + g7.content + '].');
          i = g7.end; continue;
        }
        if (bare === 'verb') {
          var dl = s[j]; var e2 = s.indexOf(dl, j + 1); if (e2 < 0) e2 = s.length; flush();
          out.push({ t: 'text', text: s.slice(j + 1, e2), fmt: Object.assign(fmtCopy(fmt), { tt: true }) }); i = e2 + 1; continue;
        }
        if (bare === 'includegraphics') {
          var o4 = readOpt(s, skipWs(s, j)); var g8 = readArg(s, o4 ? o4.end : j); flush();
          out.push({ t: 'image', opts: o4 ? o4.content : '', file: g8.content.trim() }); i = g8.end; continue;
        }
        if (IGNORE_1ARG.hasOwnProperty(bare)) {
          var jj2 = j, nargs = IGNORE_1ARG[bare];
          var o5 = readOpt(s, skipSp(s, jj2)); if (o5) jj2 = o5.end;
          for (var q = 0; q < nargs; q++) jj2 = readArg(s, jj2).end;
          i = jj2; continue;
        }
        // comando desconhecido: mantém os argumentos como texto
        ctx.warn('Comando \\' + bare + ' não reconhecido: mantido só o texto dos argumentos.');
        i = j;
        var o6 = readOpt(s, skipSp(s, i)); if (o6) i = o6.end;
        continue;
      }
      if (c === '$') {
        var dbl = s[i + 1] === '$';
        var st = i + (dbl ? 2 : 1);
        var en = st;
        while (en < s.length) { if (s[en] === '\\') { en += 2; continue; } if (s[en] === '$') break; en++; }
        flush(); out.push({ t: 'math', tex: s.slice(st, en), display: dbl });
        i = en + (dbl ? 2 : 1); continue;
      }
      if (c === '{') {
        var g9 = readGroup(s, i); flush();
        out = out.concat(parseInline(g9.content, ctx, fmtCopy(fmt)));
        i = g9.end; continue;
      }
      if (c === '}') { i++; continue; }
      if (c === '~') { buf += ' '; i++; continue; }
      if (c === '-' && s.startsWith('---', i)) { buf += '—'; i += 3; continue; }
      if (c === '-' && s.startsWith('--', i)) { buf += '–'; i += 2; continue; }
      if (c === '`' && s[i + 1] === '`') { buf += '“'; i += 2; continue; }
      if (c === "'" && s[i + 1] === "'") { buf += '”'; i += 2; continue; }
      if (c === '`') { buf += '‘'; i++; continue; }
      if (c === '<' && s[i + 1] === '<') { buf += '«'; i += 2; continue; }
      if (c === '>' && s[i + 1] === '>') { buf += '»'; i += 2; continue; }
      if (/\s/.test(c)) {
        if (!/[  ]$/.test(buf) && !(buf === '' && out.length && out[out.length - 1].t === 'br')) buf += ' ';
        i++; while (i < s.length && /\s/.test(s[i])) i++;
        continue;
      }
      buf += c; i++;
    }
    flush();
    return out;
  }

  // ============================================================ renderização
  function rPr(f, extra) {
    f = f || {};
    var x = '';
    if (extra && extra.rStyle) x += '<w:rStyle w:val="' + extra.rStyle + '"/>';
    if (f.tt) x += '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New"/>';
    if (f.b === true) x += '<w:b/><w:bCs/>'; else if (f.b === false || (extra && extra.noBold)) x += '<w:b w:val="0"/><w:bCs w:val="0"/>';
    if (f.i) x += '<w:i/><w:iCs/>';
    if (f.upper) x += '<w:caps/>'; else if (extra && extra.noCaps) x += '<w:caps w:val="0"/>';
    if (f.sc) x += '<w:smallCaps/>';
    if (f.strike) x += '<w:strike/>';
    if (f.u) x += '<w:u w:val="single"/>';
    if (f.sup) x += '<w:vertAlign w:val="superscript"/>'; else if (f.sub) x += '<w:vertAlign w:val="subscript"/>';
    return x ? '<w:rPr>' + x + '</w:rPr>' : '';
  }
  function textRun(t, f, extra) {
    return '<w:r>' + rPr(f, extra) + '<w:t xml:space="preserve">' + esc(t) + '</w:t></w:r>';
  }
  function field(instr, result, f, extra) {
    return '<w:r>' + rPr(f, extra) + '<w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r>' + rPr(f, extra) + '<w:instrText xml:space="preserve"> ' + esc(instr) + ' </w:instrText></w:r>' +
      '<w:r>' + rPr(f, extra) + '<w:fldChar w:fldCharType="separate"/></w:r>' +
      textRun(result, f, Object.assign({}, extra || {}, { noProof: true })) +
      '<w:r>' + rPr(f, extra) + '<w:fldChar w:fldCharType="end"/></w:r>';
  }

  function renderInline(items, ctx, extra) {
    var x = '';
    items.forEach(function (it) {
      switch (it.t) {
        case 'text': x += textRun(it.text, it.fmt, extra); break;
        case 'br': x += '<w:r><w:br/></w:r>'; break;
        case 'pagebreak': x += '<w:r><w:br w:type="page"/></w:r>'; break;
        case 'math':
          var r = M.latexToOmml(it.tex.trim(), false);
          if (r.ok) x += r.omml;
          else { ctx.warn('Erro na fórmula $' + it.tex + '$: ' + r.error); x += textRun('$' + it.tex + '$', { b: false }, extra); }
          break;
        case 'ref':
          var kind = ctx.labelKinds[it.label] || (/^_Ref\d+$/.test(it.label) ? 'W' : null);
          if (!kind) { kind = it.kind === 'eqref' ? 'E' : 'N'; ctx.warn('Rótulo "' + it.label + '" não encontrado no trecho nem no documento.'); }
          var bm = kind === 'W' ? it.label : bookmarkName(it.label, kind);
          var res = ctx.labelText[it.label] || (kind === 'E' ? '(?)' : '?');
          if (it.kind === 'pageref') x += field('PAGEREF ' + bm + ' \\h', '?', it.fmt, extra);
          else if (kind === 'S') x += field('REF ' + bm + ' \\r \\h', res, it.fmt, extra);
          else x += field('REF ' + bm + ' \\h', res, it.fmt, extra);
          ctx.fields++;
          break;
        case 'footnote':
          var id = ctx.footnotes.length + 101;   // ids altos: o Word renumera ao inserir
          var paras = splitParas(it.src).map(function (p, k) {
            var inner = renderInline(parseInline(p, ctx, {}), ctx);
            return '<w:p><w:pPr><w:pStyle w:val="Textodenotaderodap"/></w:pPr>' +
              (k === 0 ? '<w:r><w:rPr><w:rStyle w:val="Refdenotaderodap"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r>' : '') +
              inner + '</w:p>';
          }).join('');
          ctx.footnotes.push('<w:footnote w:id="' + id + '">' + paras + '</w:footnote>');
          x += '<w:r><w:rPr><w:rStyle w:val="Refdenotaderodap"/></w:rPr><w:footnoteReference w:id="' + id + '"/></w:r>';
          break;
        case 'link':
          var rid = ctx.rel(REL + 'hyperlink', it.url, true);
          var inner = it.inner ? renderInline(it.inner, ctx, { rStyle: 'Hyperlink' }) : textRun(it.text, {}, { rStyle: 'Hyperlink' });
          x += '<w:hyperlink r:id="' + rid + '" w:history="1">' + inner + '</w:hyperlink>';
          break;
        case 'image': x += imageRun(it, ctx); break;
      }
    });
    return x;
  }
  function splitParas(s) { return s.split(/\n[ \t]*\n/).map(function (p) { return p.trim(); }).filter(Boolean); }

  function bmWrap(name, inner, ctx) {
    var id = ctx.bmId++;
    return '<w:bookmarkStart w:id="' + id + '" w:name="' + esc(name) + '"/>' + inner + '<w:bookmarkEnd w:id="' + id + '"/>';
  }
  function para(inner, ppr) {
    return '<w:p>' + (ppr ? '<w:pPr>' + ppr + '</w:pPr>' : '') + inner + '</w:p>';
  }
  function pStyle(s) { return s ? '<w:pStyle w:val="' + s + '"/>' : ''; }

  // ---------------------------------------------------------------- imagens
  function parseLen(v, ctx) {
    v = (v || '').trim();
    var m = /^([\d.]*)\s*\\(textwidth|linewidth|columnwidth|hsize)$/.exec(v);
    if (m) return (parseFloat(m[1] || '1')) * 16;              // cm
    m = /^([\d.]+)\s*(cm|mm|in|pt|bp|em|px)?$/.exec(v);
    if (!m) return null;
    var n = parseFloat(m[1]);
    switch (m[2] || 'cm') { case 'mm': return n / 10; case 'in': return n * 2.54; case 'pt': return n * 2.54 / 72.27;
      case 'bp': return n * 2.54 / 72; case 'em': return n * 0.42; case 'px': return n * 2.54 / 96; default: return n; }
  }
  function optMap(s) {
    var o = {};
    (s || '').split(',').forEach(function (kv) { var p = kv.split('='); if (p[0].trim()) o[p[0].trim()] = (p[1] || '').trim(); });
    return o;
  }
  function findImage(ctx, file) {
    var imgs = ctx.images;
    if (imgs[file]) return imgs[file];
    var baseName = file.split('/').pop();
    var noExt = baseName.replace(/\.[a-z0-9]+$/i, '');
    var keys = Object.keys(imgs);
    for (var i = 0; i < keys.length; i++) {
      var k = keys[i], kb = k.split('/').pop(), kn = kb.replace(/\.[a-z0-9]+$/i, '');
      if (kb === baseName || kn === noExt) return imgs[k];
    }
    return null;
  }
  function imageRun(it, ctx) {
    var img = findImage(ctx, it.file);
    if (!img) {
      ctx.warn('Imagem "' + it.file + '" não carregada no painel: inserido um marcador no lugar.');
      return textRun('[imagem: ' + it.file + ']', { i: true });
    }
    var o = optMap(it.opts);
    var natW = (img.width || 600) / 96 * 2.54, natH = (img.height || 400) / 96 * 2.54;   // cm @96dpi
    var wcm = null, hcm = null;
    if (o.width) wcm = parseLen(o.width, ctx);
    if (o.height) hcm = parseLen(o.height, ctx);
    if (o.scale) { wcm = natW * parseFloat(o.scale); }
    if (wcm && !hcm) hcm = wcm * natH / natW;
    else if (hcm && !wcm) wcm = hcm * natW / natH;
    else if (!wcm && !hcm) { wcm = Math.min(natW, 16); hcm = wcm * natH / natW; }
    var ext = (img.contentType || 'image/png').split('/')[1].replace('jpeg', 'jpg');
    var name = 'ltximg' + (ctx.media.length + 1) + '.' + ext;
    ctx.media.push({ name: name, base64: img.base64, contentType: img.contentType || 'image/png' });
    var rid = ctx.rel(REL + 'image', 'media/' + name, false);
    var cx = Math.round(wcm * CM_EMU), cy = Math.round(hcm * CM_EMU);
    var id = ctx.docPrId++;
    return '<w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
      '<wp:extent cx="' + cx + '" cy="' + cy + '"/><wp:effectExtent l="0" t="0" r="0" b="0"/>' +
      '<wp:docPr id="' + id + '" name="' + esc(it.file) + '" descr="' + esc(it.file) + '"/>' +
      '<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:nvPicPr><pic:cNvPr id="0" name="' + esc(it.file) + '"/><pic:cNvPicPr/></pic:nvPicPr>' +
      '<pic:blipFill><a:blip r:embed="' + rid + '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
      '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
      '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
  }

  // ================================================================ blocos
  var HEAD = { chapter: 1, section: 2, subsection: 3, subsubsection: 4, paragraph: 5, subparagraph: 6, part: 1 };
  var FLOATS = { figure: 'figure', 'figure*': 'figure', table: 'table', 'table*': 'table', quadro: 'quadro' };
  var MATHENV = { equation: 1, 'equation*': 0, align: 1, 'align*': 0, gather: 1, 'gather*': 0, multline: 1, 'multline*': 0,
    displaymath: 0, eqnarray: 1, 'eqnarray*': 0, flalign: 1, 'flalign*': 0 };
  var IGNORE_BLOCK = { maketitle: 0, title: 1, author: 1, date: 1, subtitulo: 1, local: 1, documentclass: 1, usepackage: 1,
    listoffigures: 0, listoftables: 0, listofquadros: 0, bibliography: 1, printbibliography: 0, appendix: 0,
    frontmatter: 0, mainmatter: 0, backmatter: 0, graphicspath: 1, newcommand: 2, renewcommand: 2, input: 1, include: 1 };

  function parseBlocks(s, ctx) {
    var blocks = [];
    var p = '';
    function flushP() {
      var t = p.trim();
      if (t) blocks.push({ type: 'para', src: t });
      p = '';
    }
    var i = 0;
    while (i < s.length) {
      // linha em branco
      if (s[i] === '\n') {
        var k = i + 1; while (k < s.length && /[ \t\r]/.test(s[k])) k++;
        if (s[k] === '\n') { flushP(); i = k + 1; continue; }
        p += ' '; i++; continue;
      }
      if (s[i] === '$') {
        var dbl = s[i + 1] === '$';
        if (dbl) {
          var e0 = s.indexOf('$$', i + 2); if (e0 < 0) e0 = s.length;
          flushP(); blocks.push({ type: 'math', tex: s.slice(i + 2, e0), numbered: false }); i = e0 + 2; continue;
        }
        var en = i + 1; while (en < s.length) { if (s[en] === '\\') { en += 2; continue; } if (s[en] === '$') break; en++; }
        p += s.slice(i, en + 1); i = en + 1; continue;
      }
      if (s[i] === '{') { var g = readGroup(s, i); p += s.slice(i, g.end); i = g.end; continue; }
      if (s[i] !== '\\') { p += s[i]; i++; continue; }
      var name = cmdAt(s, i) || '';
      var j = i + 1 + name.length;
      if (name === '[') {
        var e1 = s.indexOf('\\]', j); if (e1 < 0) e1 = s.length;
        flushP(); blocks.push({ type: 'math', tex: s.slice(j, e1), numbered: false }); i = e1 + 2; continue;
      }
      if (name === '(') { var e2 = s.indexOf('\\)', j); if (e2 < 0) e2 = s.length; p += s.slice(i, e2 + 2); i = e2 + 2; continue; }
      if (name === 'verb' || name === 'verb*') { var dl = s[j]; var e3 = s.indexOf(dl, j + 1); if (e3 < 0) e3 = s.length - 1; p += s.slice(i, e3 + 1); i = e3 + 1; continue; }
      if (name === 'begin') {
        var g2 = readGroup(s, skipWs(s, j));
        var env = g2.content.trim();
        var fe = findEnd(s, g2.end, env);
        var body = s.slice(g2.end, fe.contentEnd);
        i = fe.end;
        if (env === 'document') { flushP(); blocks = blocks.concat(parseBlocks(body, ctx)); continue; }
        flushP();
        blocks = blocks.concat(envBlock(env, body, ctx));
        continue;
      }
      var bare = name.replace(/\*$/, '');
      if (HEAD[bare]) {
        flushP();
        var o = readOpt(s, skipWs(s, j)); var jj = o ? o.end : j;
        var tg = readArg(s, jj);
        var hb = { type: 'heading', level: HEAD[bare], star: /\*$/.test(name), src: tg.content };
        i = tg.end;
        var after = skipWs(s, i);
        if (s.startsWith('\\label', after)) { var lg = readArg(s, after + 6); hb.label = lg.content.trim(); i = lg.end; }
        blocks.push(hb); continue;
      }
      if (bare === 'label') {
        var lg2 = readArg(s, j);
        var lastB = blocks[blocks.length - 1];
        if (!p.trim() && lastB && (lastB.type === 'heading' || lastB.type === 'math') && !lastB.label) lastB.label = lg2.content.trim();
        i = lg2.end; continue;
      }
      if (bare === 'tableofcontents') { flushP(); blocks.push({ type: 'toc' }); i = j; continue; }
      if (bare === 'newpage' || bare === 'clearpage' || bare === 'pagebreak' || bare === 'cleardoublepage') {
        flushP(); blocks.push({ type: 'pagebreak' }); i = j; continue;
      }
      if (bare === 'item') { p += s.slice(i, j); i = j; continue; }
      if (IGNORE_BLOCK.hasOwnProperty(bare)) {
        flushP();
        var j2 = j, na = IGNORE_BLOCK[bare];
        var o2 = readOpt(s, skipWs(s, j2)); if (o2) j2 = o2.end;
        for (var q = 0; q < na; q++) { j2 = readArg(s, j2).end; var o3 = readOpt(s, skipWs(s, j2)); if (o3) j2 = o3.end; }
        i = j2; continue;
      }
      // comando inline: copia comando + argumentos para o parágrafo
      p += s.slice(i, j); i = j;
    }
    flushP();
    return blocks;
  }

  function envBlock(env, body, ctx) {
    if (FLOATS[env]) return [floatBlock(FLOATS[env], body, ctx)];
    if (MATHENV.hasOwnProperty(env)) {
      var label = null;
      var tex = body.replace(/\\label\s*\{([^}]*)\}/g, function (m, l) { if (!label) label = l.trim(); return ''; })
        .replace(/\\(nonumber|notag)\b/g, '');
      var base = env.replace('*', '');
      if (base === 'align' || base === 'eqnarray' || base === 'flalign') tex = '\\begin{aligned}' + tex.replace(/&\s*=\s*&/g, '&=') + '\\end{aligned}';
      else if (base === 'gather') tex = '\\begin{gathered}' + tex + '\\end{gathered}';
      else if (base === 'multline') tex = tex.replace(/\\\\/g, ' ');
      return [{ type: 'math', tex: tex, numbered: !!MATHENV[env], label: label }];
    }
    if (env === 'itemize' || env === 'enumerate' || env === 'description') return [listBlock(env, body, ctx)];
    if (env === 'center' || env === 'flushleft' || env === 'flushright') {
      var jc = env === 'center' ? 'center' : env === 'flushleft' ? 'left' : 'right';
      return parseBlocks(body, ctx).map(function (b) { if (b.type === 'para') b.jc = jc; return b; });
    }
    if (env === 'tabular' || env === 'tabularx' || env === 'longtable') return [{ type: 'table', kind: 'grid', src: body, env: env }];
    if (env === 'verbatim' || env === 'lstlisting' || env === 'minted') {
      return body.replace(/^\n/, '').replace(/\n$/, '').split('\n').map(function (l) { return { type: 'verb', text: l }; });
    }
    if (env === 'abstract' || env === 'quote' || env === 'quotation' || env === 'minipage' || env === 'small') {
      return parseBlocks(body.replace(/^\s*(\[[^\]]*\])?\s*\{[^}]*\}/, function (m) { return env === 'minipage' ? '' : m; }), ctx);
    }
    ctx.warn('Ambiente "' + env + '" não reconhecido: conteúdo inserido como texto comum.');
    return parseBlocks(body, ctx);
  }

  function floatBlock(kind, body, ctx) {
    var fb = { type: 'float', kind: kind, caption: null, shortCaption: null, label: null, fonte: null, content: [] };
    var s = body, i = 0, rest = '';
    while (i < s.length) {
      if (s[i] === '\\') {
        var name = cmdAt(s, i) || '', j = i + 1 + name.length;
        if (name === 'caption' || name === 'caption*') {
          var o = readOpt(s, skipWs(s, j)); var g = readArg(s, o ? o.end : j);
          fb.caption = g.content; fb.shortCaption = o ? o.content : null; i = g.end;
          var a = skipWs(s, i);
          if (s.startsWith('\\label', a)) { var lg = readArg(s, a + 6); fb.label = lg.content.trim(); i = lg.end; }
          continue;
        }
        if (name === 'label') { var lg2 = readArg(s, j); if (!fb.label) fb.label = lg2.content.trim(); i = lg2.end; continue; }
        if (name === 'fonte') { var g2 = readArg(s, j); fb.fonte = g2.content; i = g2.end; continue; }
        if (name === 'begin') {
          var g3 = readGroup(s, skipWs(s, j)); var env = g3.content.trim();
          var fe = findEnd(s, g3.end, env); var inner = s.slice(g3.end, fe.contentEnd);
          if (rest.trim()) { fb.content.push({ type: 'para', src: rest.trim(), style: 'Imagem' }); rest = ''; }
          if (env === 'tabular' || env === 'tabularx' || env === 'longtable') fb.content.push({ type: 'table', kind: kind === 'figure' ? 'grid' : kind, src: inner, env: env });
          else if (env === 'center' || env === 'minipage') { var sub = floatBlock(kind, inner, ctx); fb.content = fb.content.concat(sub.content); if (sub.caption && !fb.caption) { fb.caption = sub.caption; fb.label = sub.label; } if (sub.fonte) fb.fonte = sub.fonte; }
          else fb.content = fb.content.concat(envBlock(env, inner, ctx));
          i = fe.end; continue;
        }
        if (name === 'includegraphics') {
          var o2 = readOpt(s, skipWs(s, j)); var g4 = readArg(s, o2 ? o2.end : j);
          rest += s.slice(i, g4.end); i = g4.end; continue;
        }
        if (name === 'centering' || name === 'small' || name === 'footnotesize' || name === 'raggedright') { i = j; continue; }
        if (name === 'rowcolors') { var r1 = readArg(s, j); var r2 = readArg(s, r1.end); var r3 = readArg(s, r2.end); i = r3.end; continue; }
        if (name === 'renewcommand' || name === 'setlength') { var a1 = readArg(s, j); var a2 = readArg(s, a1.end); i = a2.end; continue; }
        if (name === 'par' || name === '\\') { if (rest.trim()) { fb.content.push({ type: 'para', src: rest.trim(), style: 'Imagem' }); rest = ''; } i = j; continue; }
      }
      if (s[i] === '\n' && /^\n[ \t]*\n/.test(s.slice(i))) { if (rest.trim()) { fb.content.push({ type: 'para', src: rest.trim(), style: 'Imagem' }); rest = ''; } }
      rest += s[i]; i++;
    }
    if (rest.trim()) fb.content.push({ type: 'para', src: rest.trim(), style: 'Imagem' });
    return fb;
  }

  function listBlock(env, body, ctx) {
    var lb = { type: 'list', ordered: env === 'enumerate', desc: env === 'description', items: [] };
    var parts = splitTop(body, /\\item\b/y);
    parts.shift(); // texto antes do 1º \item
    parts.forEach(function (it) {
      var lab = null;
      var t = it.replace(/^\s*/, '');
      if (t[0] === '[') { var o = readOpt(t, 0); if (o) { lab = o.content; t = t.slice(o.end); } }
      lb.items.push({ label: lab, blocks: parseBlocks(t, ctx) });
    });
    return lb;
  }

  // ------------------------------------------------------------- tabelas
  function parseColSpec(spec) {
    var cols = [], i = 0, pend = null;
    function push(c) { if (pend) { c.align = pend; pend = null; } cols.push(c); }
    while (i < spec.length) {
      var c = spec[i];
      if (c === '|' || c === ' ' || c === '\n') { i++; continue; }
      if (c === '>') {
        var gg = readGroup(spec, skipWs(spec, i + 1));
        if (gg) { if (/\\centering/.test(gg.content)) pend = 'center'; else if (/\\raggedleft/.test(gg.content)) pend = 'right'; else if (/\\raggedright/.test(gg.content)) pend = 'left'; }
        i = gg ? gg.end : i + 1; continue;
      }
      if (c === '@' || c === '!' || c === '<') { var g = readGroup(spec, skipWs(spec, i + 1)); i = g ? g.end : i + 1; continue; }
      if (c === '*') {
        var n = readGroup(spec, i + 1), sp = readGroup(spec, n.end);
        var rep = parseColSpec(sp.content);
        for (var r = 0; r < parseInt(n.content, 10); r++) cols = cols.concat(rep.map(function (x) { return Object.assign({}, x); }));
        i = sp.end; continue;
      }
      if (c === 'p' || c === 'm' || c === 'b') {
        var w = readGroup(spec, skipWs(spec, i + 1));
        push({ align: 'left', widthCm: parseLen(w.content) }); i = w.end; continue;
      }
      if (c === 'X') { push({ align: 'left', widthCm: null, x: true }); i++; continue; }
      if (c === 'N') { push({ align: 'left', caps: true }); i++; continue; }
      push({ align: c === 'c' ? 'center' : c === 'r' ? 'right' : 'left' }); i++;
    }
    return cols;
  }
  function parseTabular(src, env) {
    var s = src.trim();
    if (env === 'tabularx') { var wdt = readArg(s, 0); s = s.slice(wdt.end).trim(); }
    var o = readOpt(s, 0); if (o) s = s.slice(o.end).trim();
    var sg = readGroup(s, 0);
    var cols = parseColSpec(sg.content);
    var rowsSrc = splitTop(s.slice(sg.end), /\\\\(\[[^\]]*\])?/y);
    var rows = [];
    rowsSrc.forEach(function (r) {
      var t = r.replace(/\\(hline|toprule|midrule|bottomrule|linhaihr|endhead|endfirsthead|endfoot|endlastfoot)\b/g, '')
        .replace(/\\(cline|cmidrule)(\([^)]*\))?\{[^}]*\}/g, '').replace(/\\rowcolor(\[[^\]]*\])?\{[^}]*\}/g, '')
        .replace(/\\arrayrulecolor\{[^}]*\}/g, '').trim();
      if (!t) return;
      var cells = splitTop(t, /&/y).map(function (c) {
        c = c.trim();
        var mc = /^\\multicolumn\s*\{(\d+)\}\s*\{([^}]*)\}\s*/.exec(c);
        if (mc) {
          var g = readGroup(c, mc[0].length);
          return { span: parseInt(mc[1], 10), align: (parseColSpec(mc[2])[0] || {}).align, src: g ? g.content : '' };
        }
        return { span: 1, src: c };
      });
      rows.push(cells);
    });
    return { cols: cols, rows: rows };
  }

  function renderTable(tb, ctx) {
    var t = parseTabular(tb.src, tb.env);
    var ncol = Math.max(t.cols.length, Math.max.apply(null, t.rows.map(function (r) { return r.reduce(function (a, c) { return a + c.span; }, 0); }).concat([1])));
    while (t.cols.length < ncol) t.cols.push({ align: 'left' });
    var isTab = tb.kind === 'table';
    var widths;
    var anyW = t.cols.some(function (c) { return c.widthCm; });
    if (anyW) widths = t.cols.map(function (c) { return Math.round((c.widthCm || 3) * CM_TW + 216); });
    else if (isTab) {
      widths = t.cols.map(function (c, ci) {
        var mx = 4;
        t.rows.forEach(function (r) { var k = 0; r.forEach(function (cell) { if (k === ci && cell.span === 1) mx = Math.max(mx, cell.src.replace(/\\[a-zA-Z]+|[{}$]/g, '').length); k += cell.span; }); });
        return Math.round(Math.min(mx * 130 + 300, 5000));
      });
    } else {
      widths = t.cols.map(function () { return Math.floor(TEXTWIDTH_TW / ncol); });
    }
    var total = widths.reduce(function (a, b) { return a + b; }, 0);
    var tblPr = isTab
      ? '<w:tblStyle w:val="SimplesTabela3"/><w:tblW w:w="0" w:type="auto"/><w:jc w:val="center"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
      : '<w:tblStyle w:val="Tabelacomgrade"/><w:tblW w:w="' + total + '" w:type="dxa"/><w:jc w:val="center"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>';
    var x = '<w:tbl><w:tblPr>' + tblPr + '</w:tblPr><w:tblGrid>' + widths.map(function (w) { return '<w:gridCol w:w="' + w + '"/>'; }).join('') + '</w:tblGrid>';
    t.rows.forEach(function (r, ri) {
      x += '<w:tr>' + (isTab ? '<w:trPr><w:trHeight w:val="393"/></w:trPr>' : '');
      var ci = 0;
      r.forEach(function (cell) {
        var w = 0; for (var q = 0; q < cell.span; q++) w += widths[ci + q] || 0;
        var col = t.cols[ci] || {};
        var align = cell.align || col.align || 'left';
        var tcPr = '<w:tcW w:w="' + w + '" w:type="dxa"/>' + (cell.span > 1 ? '<w:gridSpan w:val="' + cell.span + '"/>' : '') + '<w:vAlign w:val="center"/>';
        var style = isTab ? 'Tabelas' : 'Quadro';
        var extra = {};
        if (isTab && ri > 0 && ci > 0) { extra.noBold = true; extra.noCaps = true; }
        var fmt = col.caps && !isTab ? { b: true, upper: true } : {};
        var inner = renderInline(parseInline(cell.src, ctx, fmt), ctx, extra);
        var jc = isTab ? (align === 'left' ? '' : '<w:jc w:val="' + align + '"/>') : (align === 'center' ? '' : '<w:jc w:val="' + (align === 'right' ? 'right' : 'left') + '"/>');
        x += '<w:tc><w:tcPr>' + tcPr + '</w:tcPr>' + para(inner, pStyle(style) + jc + (extra.noBold ? '<w:rPr><w:b w:val="0"/><w:bCs w:val="0"/></w:rPr>' : '')) + '</w:tc>';
        ci += cell.span;
      });
      while (ci < ncol) { x += '<w:tc><w:tcPr><w:tcW w:w="' + widths[ci] + '" w:type="dxa"/></w:tcPr><w:p><w:pPr>' + pStyle(isTab ? 'Tabelas' : 'Quadro') + '</w:pPr></w:p></w:tc>'; ci++; }
      x += '</w:tr>';
    });
    return x + '</w:tbl>';
  }

  // ---------------------------------------------------------- equações
  function renderEquation(b, ctx) {
    var r = M.latexToOmml(b.tex.trim(), true);
    var om;
    if (r.ok) om = r.omml;
    else { ctx.warn('Erro na equação: ' + r.error); om = textRun(b.tex, {}); }
    if (!b.numbered) return para(om, pStyle('Equaes'));
    var n = ++ctx.counters['Equação'];
    var label = '(' + n + ')';
    var num = textRun('(', {}) + field('SEQ Equação \\* ARABIC', String(n), {}) + textRun(')', {});
    if (b.label) { ctx.labelText[b.label] = label; num = bmWrap(bookmarkName(b.label, 'E'), num, ctx); }
    var nob = '<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>';
    return '<w:tbl><w:tblPr><w:tblStyle w:val="Tabelacomgrade"/><w:tblW w:w="0" w:type="auto"/>' +
      '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="704"/><w:gridCol w:w="7655"/><w:gridCol w:w="702"/></w:tblGrid><w:tr>' +
      '<w:tc><w:tcPr><w:tcW w:w="704" w:type="dxa"/>' + nob + '</w:tcPr>' + para('', pStyle('Equaes')) + '</w:tc>' +
      '<w:tc><w:tcPr><w:tcW w:w="7655" w:type="dxa"/>' + nob + '</w:tcPr>' + para(om, pStyle('Equaes')) + '</w:tc>' +
      '<w:tc><w:tcPr><w:tcW w:w="702" w:type="dxa"/>' + nob + '<w:vAlign w:val="center"/></w:tcPr>' + para(num, pStyle('Legenda')) + '</w:tc>' +
      '</w:tr></w:tbl>';
  }

  // ---------------------------------------------------------- legenda/float
  var SEQNAME = { figure: 'Figura', table: 'Tabela', quadro: 'Quadro', grid: 'Quadro' };
  function renderFloat(b, ctx) {
    var x = '';
    var seq = SEQNAME[b.kind];
    if (b.caption !== null) {
      var n = ++ctx.counters[seq];
      var num = field('SEQ ' + seq + ' \\* ARABIC', String(n), {});
      if (b.label) { ctx.labelText[b.label] = String(n); num = bmWrap(bookmarkName(b.label, 'N'), num, ctx); }
      x += para(textRun(seq + ' ', {}) + num + textRun(' - ', {}) + renderInline(parseInline(b.caption, ctx, {}), ctx), pStyle('Legenda'));
    }
    b.content.forEach(function (c) {
      if (c.type === 'table') { c.kind = b.kind === 'figure' ? 'grid' : b.kind; x += renderTable(c, ctx); }
      else if (c.type === 'para') x += para(renderInline(parseInline(c.src, ctx, {}), ctx), pStyle('Imagem'));
      else x += renderBlocks([c], ctx);
    });
    if (b.fonte !== null) x += para(textRun('Fonte: ', {}) + renderInline(parseInline(b.fonte, ctx, {}), ctx), pStyle('Fonte'));
    return x;
  }

  // -------------------------------------------------------------- listas
  function renderList(b, ctx, depth) {
    var numId = ctx.nums.length + 1;
    ctx.nums.push({ id: numId, abs: b.ordered ? 9001 : 9002 });
    var x = '';
    var lvl = Math.min(depth, 8);
    b.items.forEach(function (item) {
      var first = true;
      item.blocks.forEach(function (ib) {
        if (ib.type === 'para') {
          var inner = renderInline(parseInline(ib.src, ctx, {}), ctx);
          if (b.desc && item.label && first) inner = textRun(item.label + ' ', { b: true }) + inner;
          var ppr = pStyle('PargrafodaLista') + (first && !b.desc ? '<w:numPr><w:ilvl w:val="' + lvl + '"/><w:numId w:val="' + numId + '"/></w:numPr>' : '') +
            (first ? '' : '<w:ind w:left="' + (1287 + lvl * 720) + '" w:firstLine="0"/>');
          x += para(inner, ppr);
          first = false;
        } else if (ib.type === 'list') {
          x += renderList(ib, ctx, depth + 1);
          first = false;
        } else {
          x += renderBlocks([ib], ctx);
          first = false;
        }
      });
      if (first) x += para('', pStyle('PargrafodaLista') + '<w:numPr><w:ilvl w:val="' + lvl + '"/><w:numId w:val="' + numId + '"/></w:numPr>');
    });
    return x;
  }

  function renderBlocks(blocks, ctx) {
    var x = '';
    blocks.forEach(function (b) {
      switch (b.type) {
        case 'para':
          var items = parseInline(b.src, ctx, {});
          // parágrafo só com \newpage etc.
          if (!items.length) return;
          var ppr = b.style ? pStyle(b.style) : '';
          if (b.jc) ppr += '<w:ind w:firstLine="0"/><w:jc w:val="' + (b.jc === 'left' ? 'left' : b.jc) + '"/>';
          // imagens soltas viram parágrafo "Imagem"
          if (!b.style && items.every(function (it) { return it.t === 'image' || (it.t === 'text' && !it.text.trim()); }) && items.some(function (it) { return it.t === 'image'; })) ppr = pStyle('Imagem');
          x += para(renderInline(items, ctx), ppr);
          break;
        case 'heading':
          var inner = renderInline(parseInline(b.src, ctx, {}), ctx);
          if (b.label) { ctx.labelText[b.label] = '?'; inner = bmWrap(bookmarkName(b.label, 'S'), inner, ctx); }
          x += para(inner, pStyle('Ttulo' + b.level) + (b.star ? '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>' : ''));
          break;
        case 'math': x += renderEquation(b, ctx); break;
        case 'float': x += renderFloat(b, ctx); break;
        case 'table': x += renderTable(b, ctx) + para('', ''); break;
        case 'list': x += renderList(b, ctx, 0); break;
        case 'pagebreak': x += para('<w:r><w:br w:type="page"/></w:r>', ''); break;
        case 'verb': x += para(textRun(b.text, { tt: true }), '<w:ind w:firstLine="0"/><w:jc w:val="left"/>'); break;
        case 'toc':
          x += para('<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t>SUMÁRIO</w:t></w:r>', '<w:spacing w:line="480" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="center"/>');
          x += para(field('TOC \\o "1-6" \\h \\z \\u', 'Clique com o botão direito aqui e escolha "Atualizar campo" para gerar o sumário.', {}), pStyle('Sumrio1'));
          ctx.fields++;
          break;
      }
    });
    return x;
  }

  // pré-registra os rótulos definidos no próprio trecho (para \ref antes do \label)
  function collectLabels(blocks, ctx) {
    blocks.forEach(function (b) {
      if (b.type === 'heading' && b.label) ctx.labelKinds[b.label] = 'S';
      if (b.type === 'math' && b.label && b.numbered) ctx.labelKinds[b.label] = 'E';
      if (b.type === 'float' && b.label) ctx.labelKinds[b.label] = 'N';
      if (b.type === 'float') collectLabels(b.content, ctx);
      if (b.type === 'list') b.items.forEach(function (it) { collectLabels(it.blocks, ctx); });
    });
  }
  function preNumber(blocks, ctx) {
    // calcula os números (para o texto provisório dos \ref)
    var c = Object.assign({}, ctx.counters);
    (function walk(bs) {
      bs.forEach(function (b) {
        if (b.type === 'float' && b.caption !== null) { var s = SEQNAME[b.kind]; c[s]++; if (b.label) ctx.labelText[b.label] = String(c[s]); }
        if (b.type === 'math' && b.numbered) { c['Equação']++; if (b.label) ctx.labelText[b.label] = '(' + c['Equação'] + ')'; }
        if (b.type === 'list') b.items.forEach(function (it) { walk(it.blocks); });
      });
    })(blocks);
  }

  // ================================================================ API
  function convert(src, opts) {
    var ctx = new Ctx(opts);
    var s = stripComments(src.replace(/\r\n?/g, '\n'));
    var bd = /\\begin\{document\}/.exec(s);
    if (bd) { var ed = s.indexOf('\\end{document}'); s = s.slice(bd.index + bd[0].length, ed < 0 ? s.length : ed); }
    var blocks = parseBlocks(s, ctx);
    collectLabels(blocks, ctx);
    preNumber(blocks, ctx);
    var body = renderBlocks(blocks, ctx);
    return {
      body: body, footnotes: ctx.footnotes, rels: ctx.rels, media: ctx.media, nums: ctx.nums,
      warnings: ctx.warnings, fields: ctx.fields, blocks: blocks
    };
  }

  function numberingXml(nums) {
    return '<w:numbering ' + NS + '>' + T.absDecimal + T.absBullet +
      nums.map(function (n) {
        return '<w:num w:numId="' + n.id + '"><w:abstractNumId w:val="' + n.abs + '"/>' +
          '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>';
      }).join('') + '</w:numbering>';
  }
  function footnotesXml(fns) {
    return '<w:footnotes ' + NS + '>' +
      '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>' +
      '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
      fns.join('') + '</w:footnotes>';
  }

  // Pacote OPC "plano" (formato aceito por Word.Range.insertOoxml)
  function buildPackage(res) {
    var CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.';
    var rels = [{ id: 'rIdS', type: REL + 'styles', target: 'styles.xml' }];
    if (res.nums.length) rels.push({ id: 'rIdN', type: REL + 'numbering', target: 'numbering.xml' });
    if (res.footnotes.length) rels.push({ id: 'rIdF', type: REL + 'footnotes', target: 'footnotes.xml' });
    rels = rels.concat(res.rels);
    function part(name, ct, xml) {
      return '<pkg:part pkg:name="' + name + '" pkg:contentType="' + ct + '"><pkg:xmlData>' + xml + '</pkg:xmlData></pkg:part>';
    }
    var out = '<?xml version="1.0" standalone="yes"?><?mso-application progid="Word.Document"?>' +
      '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">';
    out += part('/_rels/.rels', 'application/vnd.openxmlformats-package.relationships+xml',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="' + REL + 'officeDocument" Target="word/document.xml"/></Relationships>');
    out += part('/word/_rels/document.xml.rels', 'application/vnd.openxmlformats-package.relationships+xml',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      rels.map(function (r) {
        return '<Relationship Id="' + r.id + '" Type="' + r.type + '" Target="' + esc(r.target) + '"' + (r.external ? ' TargetMode="External"' : '') + '/>';
      }).join('') + '</Relationships>');
    out += part('/word/document.xml', CT + 'document.main+xml', '<w:document ' + NS + '><w:body>' + res.body + '</w:body></w:document>');
    out += part('/word/styles.xml', CT + 'styles+xml', T.styles);
    if (res.nums.length) out += part('/word/numbering.xml', CT + 'numbering+xml', numberingXml(res.nums));
    if (res.footnotes.length) out += part('/word/footnotes.xml', CT + 'footnotes+xml', footnotesXml(res.footnotes));
    res.media.forEach(function (m) {
      out += '<pkg:part pkg:name="/word/media/' + m.name + '" pkg:contentType="' + m.contentType + '" pkg:compression="store"><pkg:binaryData>' +
        m.base64 + '</pkg:binaryData></pkg:part>';
    });
    return out + '</pkg:package>';
  }

  var API = { convert: convert, buildPackage: buildPackage, numberingXml: numberingXml, footnotesXml: footnotesXml,
    bookmarkName: bookmarkName, encLabel: encLabel, parseInline: parseInline, NS: NS };
  R.latex2ooxml = API;
  root.IHRLX = R;
  if (typeof module !== 'undefined') module.exports = API;
})(typeof globalThis !== 'undefined' ? globalThis : this);
