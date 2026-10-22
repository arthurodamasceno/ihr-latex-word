/* Matemática: LaTeX -> OMML (via KaTeX/MathML) e OMML -> LaTeX */
(function (root) {
  'use strict';
  var X = root.IHRLX && root.IHRLX.xml ? root.IHRLX.xml : (typeof require !== 'undefined' ? require('./xml.js') : null);
  var katex = root.katex || (typeof require !== 'undefined' ? require('../vendor/katex.min.js') : null);
  var esc = X.esc;

  // ------------------------------------------------------------------ símbolos
  var SYM = {
    'α': '\\alpha', 'β': '\\beta', 'γ': '\\gamma', 'δ': '\\delta', 'ϵ': '\\epsilon', 'ε': '\\varepsilon', 'ζ': '\\zeta',
    'η': '\\eta', 'θ': '\\theta', 'ϑ': '\\vartheta', 'ι': '\\iota', 'κ': '\\kappa', 'λ': '\\lambda', 'μ': '\\mu', 'ν': '\\nu',
    'ξ': '\\xi', 'π': '\\pi', 'ϖ': '\\varpi', 'ρ': '\\rho', 'ϱ': '\\varrho', 'σ': '\\sigma', 'ς': '\\varsigma', 'τ': '\\tau',
    'υ': '\\upsilon', 'ϕ': '\\phi', 'φ': '\\varphi', 'χ': '\\chi', 'ψ': '\\psi', 'ω': '\\omega',
    'Γ': '\\Gamma', 'Δ': '\\Delta', 'Θ': '\\Theta', 'Λ': '\\Lambda', 'Ξ': '\\Xi', 'Π': '\\Pi', 'Σ': '\\Sigma',
    'Υ': '\\Upsilon', 'Φ': '\\Phi', 'Ψ': '\\Psi', 'Ω': '\\Omega',
    '×': '\\times', '·': '\\cdot', '⋅': '\\cdot', '÷': '\\div', '±': '\\pm', '∓': '\\mp', '≤': '\\le', '≥': '\\ge',
    '≠': '\\neq', '≈': '\\approx', '≡': '\\equiv', '∼': '\\sim', '≃': '\\simeq', '∝': '\\propto', '∞': '\\infty',
    '∂': '\\partial', '∇': '\\nabla', '→': '\\to', '←': '\\leftarrow', '↔': '\\leftrightarrow', '⇒': '\\Rightarrow',
    '⇐': '\\Leftarrow', '⇔': '\\Leftrightarrow', '↦': '\\mapsto', '∈': '\\in', '∉': '\\notin', '⊂': '\\subset',
    '⊆': '\\subseteq', '⊃': '\\supset', '∪': '\\cup', '∩': '\\cap', '∅': '\\emptyset', '∀': '\\forall', '∃': '\\exists',
    '¬': '\\neg', '∧': '\\wedge', '∨': '\\vee', '…': '\\ldots', '⋯': '\\cdots', '⋮': '\\vdots', '⋱': '\\ddots',
    '°': '^{\\circ}', '∘': '\\circ', '′': "'", '″': "''", '−': '-', 'ℏ': '\\hbar', 'ℓ': '\\ell', '∠': '\\angle',
    '⊥': '\\perp', '∥': '\\parallel', '⊗': '\\otimes', '⊕': '\\oplus', '≪': '\\ll', '≫': '\\gg', '⟨': '\\langle',
    '⟩': '\\rangle', '∑': '\\sum', '∏': '\\prod', '∫': '\\int', '∬': '\\iint', '∭': '\\iiint', '∮': '\\oint',
    '⋃': '\\bigcup', '⋂': '\\bigcap', '√': '\\surd', '|': '|', '‖': '\\|', '%': '\\%', '#': '\\#', '&': '\\&', '$': '\\$',
    '{': '\\{', '}': '\\}', '⁡': '', '⁢': '', ' ': '\\,', ' ': '\\:', ' ': '\\;', ' ': '~'
  };
  var NARY = { '∑': 1, '∏': 1, '∐': 1, '∫': 1, '∬': 1, '∭': 1, '∮': 1, '⋃': 1, '⋂': 1, '⋁': 1, '⋀': 1, '⨁': 1, '⨂': 1 };
  var FUNCS = ['sin', 'cos', 'tan', 'cot', 'sec', 'csc', 'arcsin', 'arccos', 'arctan', 'sinh', 'cosh', 'tanh', 'coth',
    'log', 'ln', 'lg', 'exp', 'lim', 'max', 'min', 'sup', 'inf', 'det', 'dim', 'ker', 'arg', 'deg', 'gcd', 'Pr', 'mod'];
  var ACC_TO_OMML = { '^': '̂', 'ˆ': '̂', '̂': '̂', '~': '̃', '˜': '̃', '̃': '̃',
    '⃗': '⃗', '→': '⃗', 'ˉ': '̅', '¯': '̅', '‾': '̅', '̅': '̅', '̄': '̅',
    '˙': '̇', '̇': '̇', '¨': '̈', '̈': '̈', 'ˇ': '̌', '̌': '̌',
    '´': '́', '`': '̀', '˘': '̆' };
  var ACC_TO_TEX = { '̂': '\\hat', '̃': '\\tilde', '⃗': '\\vec', '̅': '\\bar', '̄': '\\bar',
    '̇': '\\dot', '̈': '\\ddot', '̌': '\\check', '́': '\\acute', '̀': '\\grave', '̆': '\\breve' };

  // ------------------------------------------------------------- LaTeX -> OMML
  var MR = '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>';
  function run(t, kind) {
    if (t === '') return '';
    var rpr = kind === 'text' ? '<m:rPr><m:nor/></m:rPr>' : kind === 'normal' ? '<m:rPr><m:sty m:val="p"/></m:rPr>' : '';
    return '<m:r>' + rpr + MR + '<m:t xml:space="preserve">' + esc(t) + '</m:t></m:r>';
  }
  function e(inner) { return '<m:e>' + inner + '</m:e>'; }
  var CTRL = '<m:ctrlPr>' + MR + '</m:ctrlPr>';

  function txt(el) { return X.text(el); }
  function isNaryOp(el) { return el && el.name === 'mo' && NARY[txt(el).trim()]; }

  function seq(children) {
    var out = '';
    for (var i = 0; i < children.length; i++) {
      var c = children[i];
      if (!c.name) continue;
      // função: <mi>sin</mi><mo>&#x2061;</mo> arg
      var next = children[i + 1];
      var isFn = next && next.name === 'mo' && txt(next) === '⁡';
      if (isFn) {
        var arg = '';
        var j = i + 2;
        while (j < children.length && !children[j].name) j++;
        if (j < children.length) { arg = conv(children[j]); i = j; } else { i = i + 1; }
        out += '<m:func><m:funcPr>' + CTRL + '</m:funcPr><m:fName>' + fname(c) + '</m:fName>' + e(arg) + '</m:func>';
        continue;
      }
      if (c.name === 'mi' && X.attr(c, 'mathvariant') === 'normal' && /^[A-Za-z]$/.test(txt(c))) {
        var word = txt(c), j3 = i + 1;
        while (j3 < children.length && children[j3].name === 'mi' && X.attr(children[j3], 'mathvariant') === 'normal' && /^[A-Za-z]$/.test(txt(children[j3]))) { word += txt(children[j3]); j3++; }
        if (word.length > 1) { out += run(word, 'normal'); i = j3 - 1; continue; }
      }
      var nb = naryInfo(c);
      if (nb) {
        // operando: elementos seguintes ate um operador/relacao de nivel superior
        var opnd = [];
        var j2 = i + 1;
        for (; j2 < children.length; j2++) {
          var d = children[j2];
          if (!d.name) continue;
          if (d.name === 'mo' && STOP[txt(d).trim()]) break;
          if (naryInfo(d)) break;
          opnd.push(d);
        }
        i = j2 - 1;
        out += nary(nb.chr, nb.sub, nb.sup, nb.loc, seq(opnd));
        continue;
      }
      out += conv(c);
    }
    return out;
  }
  var STOP = { '=': 1, '+': 1, '−': 1, '-': 1, '<': 1, '>': 1, '≤': 1, '≥': 1, '≠': 1, '≈': 1, ',': 1, ';': 1, '≡': 1, '→': 1, '±': 1, '∓': 1 };
  function naryInfo(el) {
    if (!el || !el.name) return null;
    var k = X.kids(el);
    if (el.name === 'mo' && NARY[txt(el).trim()]) return { chr: txt(el).trim(), sub: null, sup: null, loc: 'undOvr' };
    if (!isNaryOp(k[0])) return null;
    var chr = txt(k[0]).trim();
    switch (el.name) {
      case 'msub': return { chr: chr, sub: k[1], sup: null, loc: 'subSup' };
      case 'msup': return { chr: chr, sub: null, sup: k[1], loc: 'subSup' };
      case 'msubsup': return { chr: chr, sub: k[1], sup: k[2], loc: 'subSup' };
      case 'munder': return { chr: chr, sub: k[1], sup: null, loc: 'undOvr' };
      case 'mover': return { chr: chr, sub: null, sup: k[1], loc: 'undOvr' };
      case 'munderover': return { chr: chr, sub: k[1], sup: k[2], loc: 'undOvr' };
    }
    return null;
  }
  function fname(el) {
    if (el.name === 'mi') return run(txt(el), 'normal');
    if (el.name === 'munder') {
      var k = X.kids(el);
      return '<m:limLow><m:limLowPr>' + CTRL + '</m:limLowPr>' + e(fname(k[0])) + '<m:lim>' + conv(k[1]) + '</m:lim></m:limLow>';
    }
    if (el.name === 'mrow') {
      var ks = X.kids(el).filter(function (k) { return !(k.name === 'mo' && txt(k) === '⁡'); });
      return ks.map(fname).join('');
    }
    return conv(el);
  }

  function conv(el) {
    if (!el) return '';
    if (!el.name) return '';
    var k = X.kids(el);
    var t;
    switch (el.name) {
      case 'math': case 'semantics': case 'mstyle': case 'mpadded': case 'mtd':
        return seq(el.children.filter(function (c) { return c.name !== 'annotation' && c.name !== 'annotation-xml'; }));
      case 'mrow':
        if (k.length >= 2 && k[0].name === 'mo' && k[k.length - 1].name === 'mo' &&
          (X.attr(k[0], 'fence') === 'true' || X.attr(k[k.length - 1], 'fence') === 'true')) {
          var beg = txt(k[0]), end = txt(k[k.length - 1]);
          return '<m:d><m:dPr><m:begChr m:val="' + esc(beg) + '"/><m:endChr m:val="' + esc(end) + '"/>' + CTRL + '</m:dPr>' +
            e(seq(k.slice(1, -1))) + '</m:d>';
        }
        return seq(el.children);
      case 'mi':
        t = txt(el);
        if (t.length > 1 || X.attr(el, 'mathvariant') === 'normal') return run(t, 'normal');
        return run(t);
      case 'mn': return run(txt(el));
      case 'mo': t = txt(el); return t === '⁡' ? '' : run(t);
      case 'mtext': return run(txt(el), 'text');
      case 'mspace': return run(' ', 'text');
      case 'msub': case 'msup': case 'msubsup':
        if (isNaryOp(k[0])) {
          return nary(txt(k[0]).trim(), el.name !== 'msup' ? k[1] : null,
            el.name === 'msup' ? k[1] : (el.name === 'msubsup' ? k[2] : null), 'subSup');
        }
        if (el.name === 'msub') return '<m:sSub><m:sSubPr>' + CTRL + '</m:sSubPr>' + e(conv(k[0])) + '<m:sub>' + conv(k[1]) + '</m:sub></m:sSub>';
        if (el.name === 'msup') return '<m:sSup><m:sSupPr>' + CTRL + '</m:sSupPr>' + e(conv(k[0])) + '<m:sup>' + conv(k[1]) + '</m:sup></m:sSup>';
        return '<m:sSubSup><m:sSubSupPr>' + CTRL + '</m:sSubSupPr>' + e(conv(k[0])) + '<m:sub>' + conv(k[1]) + '</m:sub><m:sup>' + conv(k[2]) + '</m:sup></m:sSubSup>';
      case 'munder': case 'mover': case 'munderover':
        if (isNaryOp(k[0])) {
          return nary(txt(k[0]).trim(), el.name !== 'mover' ? k[1] : null,
            el.name === 'mover' ? k[1] : (el.name === 'munderover' ? k[2] : null), 'undOvr');
        }
        if (el.name === 'mover') {
          t = txt(k[1]).trim();
          if (X.attr(el, 'accent') === 'true' && X.attr(k[1], 'stretchy') !== 'true' && ACC_TO_OMML[t]) {
            return '<m:acc><m:accPr><m:chr m:val="' + ACC_TO_OMML[t] + '"/>' + CTRL + '</m:accPr>' + e(conv(k[0])) + '</m:acc>';
          }
          if (t === '‾' || t === '¯' || t === 'ˉ' || t === '̅') return '<m:bar><m:barPr><m:pos m:val="top"/>' + CTRL + '</m:barPr>' + e(conv(k[0])) + '</m:bar>';
          if (t === '⏞' || t === '︷') return '<m:groupChr><m:groupChrPr><m:chr m:val="⏞"/><m:pos m:val="top"/><m:vertJc m:val="bot"/>' + CTRL + '</m:groupChrPr>' + e(conv(k[0])) + '</m:groupChr>';
          return '<m:limUpp><m:limUppPr>' + CTRL + '</m:limUppPr>' + e(conv(k[0])) + '<m:lim>' + conv(k[1]) + '</m:lim></m:limUpp>';
        }
        if (el.name === 'munder') {
          t = txt(k[1]).trim();
          if (t === '‾' || t === '_' || t === '̲') return '<m:bar><m:barPr><m:pos m:val="bot"/>' + CTRL + '</m:barPr>' + e(conv(k[0])) + '</m:bar>';
          if (t === '⏟' || t === '︸') return '<m:groupChr><m:groupChrPr>' + CTRL + '</m:groupChrPr>' + e(conv(k[0])) + '</m:groupChr>';
          var base = txt(k[0]).replace('⁡', '').trim();
          if (FUNCS.indexOf(base) >= 0) {
            return '<m:limLow><m:limLowPr>' + CTRL + '</m:limLowPr>' + e(run(base, 'normal')) + '<m:lim>' + conv(k[1]) + '</m:lim></m:limLow>';
          }
          return '<m:limLow><m:limLowPr>' + CTRL + '</m:limLowPr>' + e(conv(k[0])) + '<m:lim>' + conv(k[1]) + '</m:lim></m:limLow>';
        }
        return '<m:limUpp><m:limUppPr>' + CTRL + '</m:limUppPr>' + e('<m:limLow><m:limLowPr>' + CTRL + '</m:limLowPr>' + e(conv(k[0])) +
          '<m:lim>' + conv(k[1]) + '</m:lim></m:limLow>') + '<m:lim>' + conv(k[2]) + '</m:lim></m:limUpp>';
      case 'mfrac':
        var nobar = /^0/.test(X.attr(el, 'linethickness') || '');
        return '<m:f><m:fPr>' + (nobar ? '<m:type m:val="noBar"/>' : '') + CTRL + '</m:fPr><m:num>' + conv(k[0]) + '</m:num><m:den>' + conv(k[1]) + '</m:den></m:f>';
      case 'msqrt':
        return '<m:rad><m:radPr><m:degHide m:val="1"/>' + CTRL + '</m:radPr><m:deg/>' + e(seq(el.children)) + '</m:rad>';
      case 'mroot':
        return '<m:rad><m:radPr>' + CTRL + '</m:radPr><m:deg>' + conv(k[1]) + '</m:deg>' + e(conv(k[0])) + '</m:rad>';
      case 'mtable':
        var rows = X.kids(el, 'mtr');
        var aligned = /right left/.test(X.attr(el, 'columnalign') || '');
        if (aligned) {
          return '<m:eqArr><m:eqArrPr>' + CTRL + '</m:eqArrPr>' + rows.map(function (r) {
            return e(X.kids(r, 'mtd').map(conv).join('<m:r><m:rPr><m:aln/></m:rPr>' + MR + '<m:t></m:t></m:r>'));
          }).join('') + '</m:eqArr>';
        }
        var ncol = Math.max.apply(null, rows.map(function (r) { return X.kids(r, 'mtd').length; }));
        return '<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="' + ncol + '"/><m:mcJc m:val="center"/></m:mcPr></m:mc></m:mcs>' + CTRL + '</m:mPr>' +
          rows.map(function (r) { return '<m:mr>' + X.kids(r, 'mtd').map(function (c) { return e(conv(c)); }).join('') + '</m:mr>'; }).join('') + '</m:m>';
      case 'menclose':
        return '<m:borderBox><m:borderBoxPr>' + CTRL + '</m:borderBoxPr>' + e(seq(el.children)) + '</m:borderBox>';
      case 'mphantom':
        return '<m:phant><m:phantPr>' + CTRL + '</m:phantPr>' + e(seq(el.children)) + '</m:phant>';
      default:
        return seq(el.children);
    }
  }
  function nary(chr, sub, sup, loc, operand) {
    var pr = '<m:naryPr><m:chr m:val="' + esc(chr) + '"/><m:limLoc m:val="' + loc + '"/>' +
      (sub ? '' : '<m:subHide m:val="1"/>') + (sup ? '' : '<m:supHide m:val="1"/>') + CTRL + '</m:naryPr>';
    // o operando vem a seguir no mrow; o Word aceita m:e vazio
    return '<m:nary>' + pr + '<m:sub>' + (sub ? conv(sub) : '') + '</m:sub><m:sup>' + (sup ? conv(sup) : '') + '</m:sup>' + e(operand || '') + '</m:nary>';
  }

  // Converte LaTeX em OMML. display=true -> m:oMathPara
  function latexToOmml(tex, display) {
    var html;
    try {
      html = katex.renderToString(tex, { output: 'mathml', throwOnError: true, displayMode: !!display, strict: 'ignore' });
    } catch (err) {
      return { ok: false, error: err.message, omml: null };
    }
    var s = html.indexOf('<math'), f = html.lastIndexOf('</math>');
    var tree = X.parse(html.slice(s, f + 7));
    var body = conv(tree);
    var om = '<m:oMath>' + body + '</m:oMath>';
    if (display) om = '<m:oMathPara><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>' + om + '</m:oMathPara>';
    return { ok: true, omml: om };
  }

  // ------------------------------------------------------------- OMML -> LaTeX
  function symTex(s) {
    var out = '';
    for (var ch of s) {
      if (SYM.hasOwnProperty(ch)) {
        var v = SYM[ch];
        out += v;
        if (/^\\[a-zA-Z]+$/.test(v)) out += ' ';
      } else out += ch;
    }
    return out;
  }
  function grp(s) { s = s.trim(); return (s.length === 1 || /^\\[a-zA-Z]+\s*$/.test(s)) ? s : '{' + s + '}'; }
  function brace(s) { return '{' + s.trim() + '}'; }
  function prop(el, prName, child) {
    var pr = X.kid(el, prName);
    var c = pr ? X.kid(pr, child) : null;
    return c ? X.val(c) : null;
  }
  function ommlKids(el) {
    return X.kids(el).map(om2tex).join('');
  }
  function sub(el, name) { return ommlKids(X.kid(el, name)); }
  var DELIM = { '(': '(', ')': ')', '[': '[', ']': ']', '{': '\\{', '}': '\\}', '|': '|', '‖': '\\|', '⟨': '\\langle', '⟩': '\\rangle',
    '⌊': '\\lfloor', '⌋': '\\rfloor', '⌈': '\\lceil', '⌉': '\\rceil', '': '.' };

  function om2tex(el) {
    if (!el || !el.name) return '';
    var n = el.name;
    switch (n) {
      case 'm:oMathPara': return X.kids(el, 'm:oMath').map(om2tex).join(' \\\\ ');
      case 'm:oMath': case 'm:e': case 'm:num': case 'm:den': case 'm:sub': case 'm:sup': case 'm:deg':
      case 'm:lim': case 'm:fName': case 'w:ins': case 'm:box': case 'm:borderBox':
        return ommlKids(el);
      case 'w:del': return '';
      case 'm:r':
        var t = X.find(el, 'm:t').map(X.text).join('');
        var rpr = X.kid(el, 'm:rPr');
        var nor = rpr && (X.kid(rpr, 'm:nor') || (X.kid(rpr, 'm:sty') && X.val(X.kid(rpr, 'm:sty')) === 'p'));
        var alnMark = rpr && X.kid(rpr, 'm:aln') ? '&' : '';
        if (alnMark) return alnMark + symTex(t);
        if (rpr && X.kid(rpr, 'm:nor')) {
          if (/^\s+$/.test(t)) return t.replace(/\u2009/g, '\\,').replace(/\u2005/g, '\\:').replace(/[ \u00a0]/g, '\\ ');
          return '\\text{' + t.replace(/[{}\\]/g, '') + '}';
        }
        if (nor && /^[A-Za-z]{2,}$/.test(t)) return FUNCS.indexOf(t) >= 0 ? '\\' + t + ' ' : '\\mathrm{' + t + '}';
        if (nor && /^[A-Za-z]$/.test(t)) return '\\mathrm{' + t + '}';
        if (/^[A-Za-z]{2,}$/.test(t) && FUNCS.indexOf(t) >= 0) return '\\' + t + ' ';
        return symTex(t);
      case 'm:f':
        var ft = prop(el, 'm:fPr', 'm:type');
        if (ft === 'lin') return sub(el, 'm:num') + '/' + sub(el, 'm:den');
        if (ft === 'noBar') return '\\genfrac{}{}{0pt}{}' + brace(sub(el, 'm:num')) + brace(sub(el, 'm:den'));
        return '\\frac' + brace(sub(el, 'm:num')) + brace(sub(el, 'm:den'));
      case 'm:sSub': return grp(sub(el, 'm:e')) + '_' + grp(sub(el, 'm:sub'));
      case 'm:sSup': return grp(sub(el, 'm:e')) + '^' + grp(sub(el, 'm:sup'));
      case 'm:sSubSup': return grp(sub(el, 'm:e')) + '_' + grp(sub(el, 'm:sub')) + '^' + grp(sub(el, 'm:sup'));
      case 'm:sPre': return '{}_' + grp(sub(el, 'm:sub')) + '^' + grp(sub(el, 'm:sup')) + grp(sub(el, 'm:e'));
      case 'm:rad':
        var hide = prop(el, 'm:radPr', 'm:degHide');
        var deg = sub(el, 'm:deg').trim();
        return '\\sqrt' + ((hide === '1' || hide === 'on' || !deg) ? '' : '[' + deg + ']') + brace(sub(el, 'm:e'));
      case 'm:d':
        var b = prop(el, 'm:dPr', 'm:begChr'), en = prop(el, 'm:dPr', 'm:endChr'), sep = prop(el, 'm:dPr', 'm:sepChr');
        if (b === null) b = '('; if (en === null) en = ')'; if (sep === null) sep = '|';
        var es = X.kids(el, 'm:e');
        if (es.length === 1) {
          var only = X.kids(es[0]).filter(function (q) { return !/Pr$/.test(q.name); });
          if (only.length === 1 && only[0].name === 'm:f' && prop(only[0], 'm:fPr', 'm:type') === 'noBar' && b === '(' && en === ')')
            return '\\binom' + brace(sub(only[0], 'm:num')) + brace(sub(only[0], 'm:den'));
          if (only.length === 1 && only[0].name === 'm:m') {
            var env = b === '[' ? 'bmatrix' : b === '(' ? 'pmatrix' : b === '|' ? 'vmatrix' : b === '{' ? 'Bmatrix' : null;
            if (env) return om2tex(only[0]).replace(/matrix\}/g, env + '}');
          }
        }
        var parts = es.map(ommlKids);
        var inner = parts.join(' ' + (DELIM[sep] || sep) + ' ');
        if ((b === '(' && en === ')') || (b === '[' && en === ']')) {
          var tall = X.find(el, 'm:f').length || X.find(el, 'm:nary').length || X.find(el, 'm:m').length;
          if (!tall) return b + inner + en;
        }
        return '\\left' + (DELIM[b] !== undefined ? DELIM[b] : b) + ' ' + inner + ' \\right' + (DELIM[en] !== undefined ? DELIM[en] : en);
      case 'm:nary':
        var chr = prop(el, 'm:naryPr', 'm:chr') || '∫';
        var op = SYM[chr] || chr;
        var s1 = sub(el, 'm:sub').trim(), s2 = sub(el, 'm:sup').trim();
        return op + (s1 ? '_' + grp(s1) : '') + (s2 ? '^' + grp(s2) : '') + ' ' + sub(el, 'm:e');
      case 'm:func':
        var fn = sub(el, 'm:fName').trim();
        return fn + (/[a-zA-Z]$/.test(fn) ? ' ' : '') + sub(el, 'm:e');
      case 'm:limLow':
        var base = sub(el, 'm:e').trim();
        if (/^\\(lim|max|min|sup|inf)\s*$/.test(base) || /^\\mathrm\{(lim|max|min)\}$/.test(base)) return base.replace(/\\mathrm\{(\w+)\}/, '\\$1').trim() + '_' + brace(sub(el, 'm:lim'));
        return '\\underset' + brace(sub(el, 'm:lim')) + brace(base);
      case 'm:limUpp': return '\\overset' + brace(sub(el, 'm:lim')) + brace(sub(el, 'm:e'));
      case 'm:acc':
        var ac = prop(el, 'm:accPr', 'm:chr') || '̂';
        return (ACC_TO_TEX[ac] || '\\hat') + brace(sub(el, 'm:e'));
      case 'm:bar':
        return (prop(el, 'm:barPr', 'm:pos') === 'top' ? '\\overline' : '\\underline') + brace(sub(el, 'm:e'));
      case 'm:groupChr':
        var gp = prop(el, 'm:groupChrPr', 'm:pos');
        return (gp === 'top' ? '\\overbrace' : '\\underbrace') + brace(sub(el, 'm:e'));
      case 'm:m':
        return '\\begin{matrix}' + X.kids(el, 'm:mr').map(function (r) {
          return X.kids(r, 'm:e').map(ommlKids).join(' & ');
        }).join(' \\\\ ') + '\\end{matrix}';
      case 'm:eqArr':
        return '\\begin{aligned}' + X.kids(el, 'm:e').map(function (r) { var q = ommlKids(r); return q.indexOf('&') >= 0 ? q : q.replace(/^([^=]*)=/, '$1&='); }).join(' \\\\ ') + '\\end{aligned}';
      case 'm:phant': return '\\phantom' + brace(sub(el, 'm:e'));
      case 'm:t': return symTex(X.text(el));
      default:
        if (/Pr$/.test(n) || n === 'm:ctrlPr' || n === 'w:rPr') return '';
        return ommlKids(el);
    }
  }
  function ommlToLatex(el) {
    return om2tex(el).replace(/\s+/g, ' ').replace(/\s*([_^])\s*/g, '$1').trim();
  }

  var M = { latexToOmml: latexToOmml, ommlToLatex: ommlToLatex };
  root.IHRLX = root.IHRLX || {};
  root.IHRLX.math = M;
  if (typeof module !== 'undefined') module.exports = M;
})(typeof globalThis !== 'undefined' ? globalThis : this);
