/* Comandos personalizados e pacotes: \newcommand, \renewcommand, \DeclareMathOperator, \def
   e emulação de alguns pacotes por meio dos mesmos comandos. A expansão é feita no texto LaTeX
   antes da conversão, então vale tanto no texto quanto nas fórmulas. */
(function (root) {
  'use strict';

  // ------------------------------------------------------------- leitura
  function readGroup(s, i) {            // s[i] === '{' -> {content, end}
    var d = 0;
    for (var j = i; j < s.length; j++) {
      var c = s[j];
      if (c === '\\') { j++; continue; }
      if (c === '{') d++;
      else if (c === '}') { d--; if (d === 0) return { content: s.slice(i + 1, j), end: j + 1 }; }
    }
    return null;
  }
  function skipWs(s, i) { while (i < s.length && /[ \t\r\n]/.test(s[i])) i++; return i; }
  function readName(s, i) {             // \nome ou {\nome}
    i = skipWs(s, i);
    if (s[i] === '{') { var g = readGroup(s, i); if (!g) return null; var m = /^\s*\\([A-Za-z@]+|.)\s*$/.exec(g.content); return m ? { name: m[1], end: g.end } : null; }
    var m2 = /^\\([A-Za-z@]+|.)/.exec(s.slice(i));
    return m2 ? { name: m2[1], end: i + m2[0].length } : null;
  }
  function readOpt(s, i) {
    i = skipWs(s, i);
    if (s[i] !== '[') return null;
    var d = 0;
    for (var j = i; j < s.length; j++) {
      if (s[j] === '{') d++; else if (s[j] === '}') d--;
      else if (s[j] === ']' && d === 0) return { content: s.slice(i + 1, j), end: j + 1 };
    }
    return null;
  }

  // lê definições de um texto; devolve {defs, packages, rest (texto sem as definições)}
  function parseDefs(text) {
    var defs = {}, packages = [], out = '', i = 0, s = text || '';
    var re = /\\(newcommand|renewcommand|providecommand|DeclareMathOperator|def|usepackage|RequirePackage)\b(\*?)/g, m;
    while ((m = re.exec(s))) {
      var st = m.index, j = m.index + m[0].length, ok = false;
      if (m[1] === 'usepackage' || m[1] === 'RequirePackage') {
        var o = readOpt(s, j); if (o) j = o.end;
        j = skipWs(s, j);
        if (s[j] === '{') { var g = readGroup(s, j); if (g) { g.content.split(',').forEach(function (p) { p = p.trim(); if (p) packages.push({ name: p, opts: o ? o.content : '' }); }); j = g.end; ok = true; } }
      } else if (m[1] === 'DeclareMathOperator') {
        var nm = readName(s, j);
        if (nm) { j = skipWs(s, nm.end); var b = s[j] === '{' ? readGroup(s, j) : null;
          if (b) { defs[nm.name] = { n: 0, body: (m[2] ? '\\operatorname*{' : '\\operatorname{') + b.content + '}' }; j = b.end; ok = true; } }
      } else if (m[1] === 'def') {
        var dn = /^\\([A-Za-z@]+)((?:#\d)*)\s*/.exec(s.slice(j).replace(/^\s*/, ''));
        if (dn) { j = skipWs(s, j) + dn[0].length; var db = s[j] === '{' ? readGroup(s, j) : null;
          if (db) { defs[dn[1]] = { n: (dn[2].match(/#/g) || []).length, body: db.content }; j = db.end; ok = true; } }
      } else {
        var nm2 = readName(s, j);
        if (nm2) {
          j = nm2.end;
          var n = 0, dflt = null, o1 = readOpt(s, j);
          if (o1) { n = parseInt(o1.content, 10) || 0; j = o1.end; var o2 = readOpt(s, j); if (o2) { dflt = o2.content; j = o2.end; } }
          j = skipWs(s, j);
          var bd = s[j] === '{' ? readGroup(s, j) : null;
          if (bd) { if (!(m[1] === 'providecommand' && defs[nm2.name])) defs[nm2.name] = { n: n, dflt: dflt, body: bd.content }; j = bd.end; ok = true; }
        }
      }
      if (ok) { out += s.slice(i, st); i = j; re.lastIndex = j; }
    }
    out += s.slice(i);
    return { defs: defs, packages: packages, rest: out };
  }

  // ------------------------------------------------------------- pacotes
  // nativo: o conversor já entende; layout: a formatação vem do modelo do Word; emulado: comandos abaixo
  var NATIVO = 'amsmath amssymb amsfonts mathtools graphicx float xcolor color hyperref url csquotes textcomp array caption multirow booktabs tabularx longtable enumitem footmisc'.split(' ');
  var LAYOUT = 'babel inputenc fontenc lmodern mathptmx times newtxtext newtxmath fontspec geometry setspace indentfirst microtype titlesec fancyhdr parskip ragged2e lipsum'.split(' ');
  var EMULADO = {
    bm: { bm: [1, '\\boldsymbol{#1}'] },
    gensymb: { degree: [0, '{^{\\circ}}'], celsius: [0, '{^{\\circ}\\mathrm{C}}'], ohm: [0, '\\Omega'], micro: [0, '\\mu'], perthousand: [0, '‰'] },
    physics: {
      abs: [1, '\\left|#1\\right|'], norm: [1, '\\left\\|#1\\right\\|'], vb: [1, '\\mathbf{#1}'], va: [1, '\\vec{#1}'], vu: [1, '\\hat{\\mathbf{#1}}'],
      dv: [2, '\\frac{\\mathrm{d}#1}{\\mathrm{d}#2}'], pdv: [2, '\\frac{\\partial #1}{\\partial #2}'], dd: [0, '\\mathrm{d}'],
      grad: [0, '\\nabla'], curl: [0, '\\nabla\\times'], divergence: [0, '\\nabla\\cdot'], laplacian: [0, '\\nabla^2'],
      bra: [1, '\\left\\langle #1\\right|'], ket: [1, '\\left|#1\\right\\rangle'], braket: [2, '\\left\\langle #1\\middle|#2\\right\\rangle'],
      tr: [0, '\\operatorname{tr}'], Tr: [0, '\\operatorname{Tr}'], rank: [0, '\\operatorname{rank}'], qty: [1, '\\left(#1\\right)']
    }
  };
  function packageDefs(packages) {
    var defs = {};
    packages.forEach(function (p) {
      var set = EMULADO[p.name];
      if (set) Object.keys(set).forEach(function (k) { defs[k] = { n: set[k][0], body: set[k][1] }; });
    });
    return defs;
  }
  function classify(name) {
    if (EMULADO[name]) return 'emulado';
    if (NATIVO.indexOf(name) >= 0) return 'nativo';
    if (LAYOUT.indexOf(name) >= 0) return 'layout';
    return 'nao';
  }

  // ------------------------------------------------------------- expansão
  function expand(src, defs) {
    var names = Object.keys(defs || {});
    if (!names.length || !src) return src;
    var s = src;
    for (var pass = 0; pass < 20; pass++) {
      var out = '', i = 0, changed = false;
      while (i < s.length) {
        var k = s.indexOf('\\', i);
        if (k < 0) { out += s.slice(i); break; }
        out += s.slice(i, k);
        var m = /^\\([A-Za-z@]+)/.exec(s.slice(k));
        if (!m) { out += s.substr(k, 2); i = k + 2; continue; }
        var d = defs[m[1]];
        if (!d) { out += m[0]; i = k + m[0].length; continue; }
        var j = k + m[0].length, args = [], okArgs = true;
        for (var a = 0; a < d.n; a++) {
          if (a === 0 && d.dflt !== null && d.dflt !== undefined) {
            var o = readOpt(s, j);
            if (o) { args.push(o.content); j = o.end; } else args.push(d.dflt);
            continue;
          }
          var jj = skipWs(s, j);
          if (s[jj] === '{') { var g = readGroup(s, jj); if (!g) { okArgs = false; break; } args.push(g.content); j = g.end; }
          else if (s[jj] === '\\') { var t = /^\\([A-Za-z@]+|.)/.exec(s.slice(jj)); args.push(t[0]); j = jj + t[0].length; }
          else if (jj < s.length) { args.push(s[jj]); j = jj + 1; }
          else { okArgs = false; break; }
        }
        if (!okArgs) { out += m[0]; i = k + m[0].length; continue; }
        var body = d.body.replace(/#(\d)/g, function (x, n) { return args[parseInt(n, 10) - 1] !== undefined ? args[parseInt(n, 10) - 1] : ''; });
        // comando sem argumentos seguido de letra: protege com {} para não grudar no texto
        if (d.n === 0 && /^[A-Za-z]/.test(s.slice(j))) body = '{' + body + '}';
        out += body; i = j; changed = true;
      }
      s = out;
      if (!changed) break;
    }
    return s;
  }

  // prepara o texto do editor: definições do preâmbulo do documento + as do próprio editor
  function prepare(editorText, preamble) {
    var p = parseDefs(preamble || ''), e = parseDefs(editorText || '');
    var pk = p.packages.concat(e.packages);
    var defs = Object.assign({}, packageDefs(pk), p.defs, e.defs);
    return { text: expand(e.rest, defs), defs: defs, packages: pk };
  }

  var API = { parseDefs: parseDefs, expand: expand, prepare: prepare, classify: classify, packageDefs: packageDefs, EMULADO: EMULADO };
  root.IHRLX = root.IHRLX || {};
  root.IHRLX.macros = API;
  if (typeof module !== 'undefined') module.exports = API;
})(typeof globalThis !== 'undefined' ? globalThis : this);
