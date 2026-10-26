/* OOXML (seleção do Word) -> LaTeX no padrão da classe ihr.cls */
(function (root) {
  'use strict';
  var R = root.IHRLX || {};
  var X = R.xml || require('./xml.js');
  var M = R.math || require('./math.js');

  // ------------------------------------------------------------ pacote
  function parsePackage(pkg) {
    var tree = X.parse(pkg);
    var parts = {};
    X.kids(tree, 'pkg:part').forEach(function (p) {
      var name = X.attr(p, 'pkg:name');
      var xd = X.kid(p, 'pkg:xmlData');
      var bd = X.kid(p, 'pkg:binaryData');
      parts[name] = xd ? { xml: X.kids(xd)[0], ct: X.attr(p, 'pkg:contentType') } : { base64: bd ? X.text(bd).replace(/\s+/g, '') : '', ct: X.attr(p, 'pkg:contentType') };
    });
    function rels(of) {
      var p = parts[of]; var map = {};
      if (p && p.xml) X.kids(p.xml, 'Relationship').forEach(function (r) { map[X.attr(r, 'Id')] = { target: X.attr(r, 'Target'), type: X.attr(r, 'Type'), mode: X.attr(r, 'TargetMode') }; });
      return map;
    }
    var names = Object.keys(parts);
    var docName = names.filter(function (n) { return /document\.main\+xml/.test(parts[n].ct || ''); })[0] ||
      names.filter(function (n) { return /\/word\/document\d*\.xml$/.test(n); })[0] || '/word/document.xml';
    var dir = docName.replace(/[^/]*$/, ''), leaf = docName.split('/').pop();
    // relações do documento (e, por garantia, de todas as partes .rels)
    var docRels = rels(dir + '_rels/' + leaf + '.rels');
    names.filter(function (n) { return /\.rels$/.test(n); }).forEach(function (n) {
      var m = rels(n); Object.keys(m).forEach(function (k) { if (!docRels[k]) docRels[k] = m[k]; });
    });
    function byKind(re) { var n = names.filter(function (q) { return re.test(parts[q].ct || ''); })[0]; return n && parts[n].xml; }
    return {
      document: parts[docName] && parts[docName].xml,
      styles: byKind(/styles\+xml/) || (parts['/word/styles.xml'] && parts['/word/styles.xml'].xml),
      numbering: byKind(/numbering\+xml/) || (parts['/word/numbering.xml'] && parts['/word/numbering.xml'].xml),
      footnotes: byKind(/footnotes\+xml/) || (parts['/word/footnotes.xml'] && parts['/word/footnotes.xml'].xml),
      rels: docRels,
      partNames: names,
      part: function (target) {
        if (!target) return null;
        var t = target.replace(/^\.?\//, '');
        var p = parts[dir + t] || parts['/' + t] || parts['/word/' + t];
        if (p) return p;
        var base = t.split('/').pop();
        var hit = names.filter(function (n) { return n.split('/').pop() === base && parts[n].base64; })[0];
        return hit ? parts[hit] : null;
      }
    };
  }

  // ------------------------------------------------------------ utilidades
  function decLabel(name) {
    var m = /^_Ref([SNE])_(.*)$/.exec(name);
    if (!m) return { label: name, kind: /^_Ref\d+$/.test(name) ? 'W' : null };
    return { label: m[2].replace(/_([0-9A-F]{2})/g, function (q, h) { return String.fromCharCode(parseInt(h, 16)); }), kind: m[1] };
  }
  function texEsc(s) {
    return s.replace(/\\/g, '\u0000').replace(/([{}%$&#_])/g, '\\$1').replace(/\^/g, '\\textasciicircum{}')
      .replace(/~/g, '\\textasciitilde{}').replace(/\u0000/g, '\\textbackslash{}').replace(/\u00a0/g, '~').replace(/\u202f/g, '\\,')
      .replace(/\t/g, '\\quad ');
  }
  function imageSize(b64) {
    try {
      var bin = typeof atob !== 'undefined' ? atob(b64.slice(0, 40000)) : Buffer.from(b64.slice(0, 40000), 'base64').toString('binary');
      var u = function (i) { return bin.charCodeAt(i) & 255; };
      if (u(0) === 0x89 && bin.slice(1, 4) === 'PNG') return { w: (u(16) << 24 | u(17) << 16 | u(18) << 8 | u(19)) >>> 0, h: (u(20) << 24 | u(21) << 16 | u(22) << 8 | u(23)) >>> 0 };
      if (bin.slice(0, 3) === 'GIF') return { w: u(6) | u(7) << 8, h: u(8) | u(9) << 8 };
      if (u(0) === 0xFF && u(1) === 0xD8) {
        var i = 2;
        while (i < bin.length) {
          if (u(i) !== 0xFF) { i++; continue; }
          var mk = u(i + 1), len = u(i + 2) << 8 | u(i + 3);
          if (mk >= 0xC0 && mk <= 0xCF && mk !== 0xC4 && mk !== 0xC8 && mk !== 0xCC) return { h: u(i + 5) << 8 | u(i + 6), w: u(i + 7) << 8 | u(i + 8) };
          i += 2 + len;
        }
      }
    } catch (e) { /* ignora */ }
    return { w: 600, h: 400 };
  }

  function Ctx(P) {
    this.P = P;
    this.styles = {};
    this.images = {};
    this.imgCount = 0;
    this.warnings = [];
    this.unsafe = [];
    this.imageOrder = [];
    var self = this;
    if (P.styles) X.kids(P.styles, 'w:style').forEach(function (s) {
      var nm = X.kid(s, 'w:name'), bo = X.kid(s, 'w:basedOn'), ppr = X.kid(s, 'w:pPr');
      var ol = ppr ? X.kid(ppr, 'w:outlineLvl') : null;
      self.styles[X.attr(s, 'w:styleId')] = { name: nm ? X.val(nm) : '', basedOn: bo ? X.val(bo) : null, outline: ol ? parseInt(X.val(ol), 10) : null };
    });
    this.numFmt = {};
    if (P.numbering) {
      var abs = {};
      X.kids(P.numbering, 'w:abstractNum').forEach(function (a) {
        var fm = {};
        X.kids(a, 'w:lvl').forEach(function (l) { var f = X.kid(l, 'w:numFmt'); fm[X.attr(l, 'w:ilvl')] = f ? X.val(f) : 'decimal'; });
        abs[X.attr(a, 'w:abstractNumId')] = fm;
      });
      X.kids(P.numbering, 'w:num').forEach(function (n) {
        var an = X.kid(n, 'w:abstractNumId');
        self.numFmt[X.attr(n, 'w:numId')] = abs[an ? X.val(an) : ''] || {};
      });
    }
  }
  Ctx.prototype.warn = function (m) { if (this.warnings.indexOf(m) < 0) this.warnings.push(m); };
  Ctx.prototype.flagUnsafe = function (m) { if (this.unsafe.indexOf(m) < 0) this.unsafe.push(m); };
  Ctx.prototype.role = function (styleId) {
    var seen = 0, id = styleId;
    while (id && seen++ < 10) {
      var st = this.styles[id];
      var nm = (st ? st.name : id) || '';
      var low = nm.toLowerCase();
      if (low === 'título sem numeração' || id === 'TtuloSemNumerao') return { kind: 'heading', level: 1, unnum: true };
      var hm = /^heading (\d)$/.exec(low) || /^t[ií]tulo (\d)$/.exec(low) || /^Ttulo(\d)$/.exec(id);
      if (hm) return { kind: 'heading', level: parseInt(hm[1], 10) };
      if (low === 'capa - título' || id === 'CapaTtulo') return { kind: 'cover', field: 'title' };
      if (low === 'capa - subtítulo' || id === 'CapaSubttulo') return { kind: 'cover', field: 'subtitulo' };
      if (low === 'capa - local e data' || id === 'CapaLocaleData') return { kind: 'cover', field: 'localdata' };
      if (low === 'toc heading' || id === 'CabealhodoSumrio') return { kind: 'tochead' };
      if (low === 'bibliography' || id === 'Bibliografia') return { kind: 'bibitem' };
      if (low === 'caption' || id === 'Legenda') return { kind: 'caption' };
      if (low === 'fonte') return { kind: 'fonte' };
      if (low === 'imagem') return { kind: 'imagem' };
      if (low === 'equações' || id === 'Equaes') return { kind: 'equacao' };
      if (/^toc \d/.test(low) || /^Sumrio\d/.test(id)) return { kind: 'toc' };
      if (low === 'list paragraph' || id === 'PargrafodaLista') return { kind: 'list' };
      if (low === 'quadro' || low === 'tabelas') return { kind: 'cell' };
      if (low === 'title' || low === 'subtitle') return { kind: 'normal' };
      id = st ? st.basedOn : null;
    }
    return { kind: 'normal' };
  };

  // ------------------------------------------------------- inline -> tokens
  // tokens: {t:'text', s, f}, {t:'math', tex}, {t:'field', instr, res:[tokens]}, {t:'bm', name},
  //         {t:'br'}, {t:'page'}, {t:'img', tex}, {t:'fn', tex}, {t:'link', url, res:[tokens]}
  function runFmt(r) {
    var p = X.kid(r, 'w:rPr'); var f = {};
    if (!p) return f;
    function on(n) { var e = X.kid(p, n); if (!e) return false; var v = X.val(e); return v === null || (v !== '0' && v !== 'false' && v !== 'none'); }
    if (on('w:b')) f.b = 1;
    if (on('w:i')) f.i = 1;
    if (on('w:u')) f.u = 1;
    if (on('w:strike')) f.s = 1;
    if (on('w:smallCaps')) f.sc = 1;
    var va = X.kid(p, 'w:vertAlign'); if (va) { if (X.val(va) === 'superscript') f.sup = 1; else if (X.val(va) === 'subscript') f.sub = 1; }
    var rf = X.kid(p, 'w:rFonts'); if (rf && /Courier|Consolas|Mono/i.test(X.attr(rf, 'w:ascii') || '')) f.tt = 1;
    var rs = X.kid(p, 'w:rStyle'); if (rs && /Hyperlink/i.test(X.val(rs))) f.link = 1;
    return f;
  }

  function inlineTokens(nodes, ctx) {
    var out = [];
    var stack = [];   // campos complexos abertos
    function emit(tok) {
      if (stack.length) {
        var top = stack[stack.length - 1];
        if (top.state === 'res') top.res.push(tok);
      } else out.push(tok);
    }
    function walk(list) {
      list.forEach(function (n) {
        if (!n.name) return;
        switch (n.name) {
          case 'w:r': runNode(n); break;
          case 'w:hyperlink':
            var rid = X.attr(n, 'r:id'), anchor = X.attr(n, 'w:anchor');
            var url = rid && ctx.P.rels[rid] ? ctx.P.rels[rid].target : null;
            var sub = inlineTokens(n.children, ctx);
            if (url) emit({ t: 'link', url: url, res: sub }); else sub.forEach(emit);
            break;
          case 'w:fldSimple':
            emit({ t: 'field', instr: X.attr(n, 'w:instr') || '', res: inlineTokens(n.children, ctx) });
            break;
          case 'm:oMath': emit({ t: 'math', tex: M.ommlToLatex(n) }); break;
          case 'm:oMathPara': emit({ t: 'dmath', tex: M.ommlToLatex(n) }); break;
          case 'w:bookmarkStart': emit({ t: 'bm', name: X.attr(n, 'w:name') }); break;
          case 'w:ins': case 'w:smartTag': case 'w:customXml': case 'w:dir': case 'w:bdo': walk(n.children); if (n.name === 'w:ins') ctx.warn('A seleção tem alterações controladas: o LaTeX mostra o texto como se elas fossem aceitas.'); break;
          case 'w:sdt': var sc = X.kid(n, 'w:sdtContent'); if (sc) walk(sc.children); break;
          case 'w:del': ctx.warn('A seleção tem alterações controladas: o LaTeX mostra o texto como se elas fossem aceitas.'); break;
          case 'w:commentRangeStart': ctx.warn('A seleção tem comentários: eles serão removidos se você aplicar a edição.'); break;
          default: break;
        }
      });
    }
    function runNode(r) {
      var f = runFmt(r);
      r.children.forEach(function (c) {
        if (!c.name) return;
        switch (c.name) {
          case 'w:t': emit({ t: 'text', s: X.text(c), f: f }); break;
          case 'w:delText': break;
          case 'w:tab': emit({ t: 'text', s: '\t', f: f }); break;
          case 'w:br': emit(X.attr(c, 'w:type') === 'page' ? { t: 'page' } : { t: 'br' }); break;
          case 'w:cr': emit({ t: 'br' }); break;
          case 'w:noBreakHyphen': emit({ t: 'text', s: '-', f: f }); break;
          case 'w:sym':
            var code = (X.attr(c, 'w:char') || '').toUpperCase();
            var SYMS = { F0B7: '•', F0E0: '→', F0DE: '⇒', F0AE: '→', F0B1: '±', F0B4: '×', F0B0: '°', F0A3: '≤', F0B3: '≥', F0B9: '≠', F0BB: '≈', F0D7: '·' };
            emit({ t: 'text', s: SYMS[code] || '?', f: f }); break;
          case 'w:fldChar':
            var ty = X.attr(c, 'w:fldCharType');
            if (ty === 'begin') stack.push({ instr: '', res: [], state: 'instr' });
            else if (ty === 'separate' && stack.length) stack[stack.length - 1].state = 'res';
            else if (ty === 'end' && stack.length) { var fld = stack.pop(); emit({ t: 'field', instr: fld.instr, res: fld.res }); }
            break;
          case 'w:instrText': if (stack.length) stack[stack.length - 1].instr += X.text(c); break;
          case 'w:drawing': case 'w:pict': case 'mc:AlternateContent': emit(drawingTok(c, ctx)); break;
          case 'w:footnoteReference': emit({ t: 'fn', tex: footnoteTex(X.attr(c, 'w:id'), ctx) }); break;
          case 'w:endnoteReference': ctx.warn('Notas de fim não são suportadas.'); break;
          default: break;
        }
      });
    }
    walk(nodes);
    // campos não fechados (seleção cortou no meio)
    while (stack.length) { var f2 = stack.pop(); f2.res.forEach(function (t) { out.push(t); }); }
    return out;
  }

  function drawingTok(c, ctx) {
    if (X.find(c, 'wps:txbx').length || X.find(c, 'v:textbox').length) {
      ctx.flagUnsafe('caixa de texto'); return { t: 'text', s: '', f: {} };
    }
    if (X.find(c, 'wp:anchor').length) ctx.flagUnsafe('imagem flutuante (ancorada)');
    var blip = X.find(c, 'a:blip')[0];
    var ext = X.find(c, 'wp:extent')[0];
    var docPr = X.find(c, 'wp:docPr')[0];
    if (!blip) { ctx.flagUnsafe('objeto gráfico'); return { t: 'text', s: '', f: {} }; }
    var rid = X.attr(blip, 'r:embed') || X.attr(blip, 'r:link');
    var rel = ctx.P.rels[rid];
    var part = rel ? ctx.P.part(rel.target) : null;
    var nm = docPr ? (X.attr(docPr, 'descr') || X.attr(docPr, 'name') || '') : '';
    var extn = rel ? (rel.target.split('.').pop() || 'png').toLowerCase() : 'png';
    var leaf = nm.split(/[\\/]/).pop().trim();
    var base;
    if (/^[\w\-. ]+\.(png|jpe?g|gif|bmp|svg|emf|wmf)$/i.test(leaf)) base = leaf.replace(/\s+/g, '-');
    else if (/^[\w\-]+$/.test(leaf) && !/^(Imagem|Picture|Image|Figura)[ _-]?\d*$/i.test(leaf)) base = leaf + '.' + extn;
    else base = 'imagem-' + (++ctx.imgCount) + '.' + extn;
    while (ctx.images[base] && part && ctx.images[base].base64 !== part.base64) base = base.replace(/(\.\w+)$/, '-b$1');
    var inline = !X.find(c, 'wp:anchor').length;
    if (part && part.base64) {
      var sz = imageSize(part.base64);
      ctx.images[base] = { base64: part.base64, contentType: part.ct || 'image/png', width: sz.w, height: sz.h };
    }
    ctx.imageOrder.push({ name: base, inline: inline, missing: !(part && part.base64), rid: rid, target: rel ? rel.target : null });
    var wcm = ext ? (parseInt(X.attr(ext, 'cx'), 10) / 360000) : null;
    return { t: 'img', tex: '\\includegraphics' + (wcm ? '[width=' + (Math.round(wcm * 100) / 100) + 'cm]' : '') + '{' + base + '}' };
  }

  function footnoteTex(id, ctx) {
    if (!ctx.P.footnotes) { ctx.warn('Não consegui ler o texto das notas de rodapé.'); return ''; }
    var fn = X.kids(ctx.P.footnotes, 'w:footnote').filter(function (f) { return X.attr(f, 'w:id') === id; })[0];
    if (!fn) return '';
    return X.kids(fn, 'w:p').map(function (p) { return tokensToTex(inlineTokens(p.children, ctx), ctx).trim(); }).filter(Boolean).join('\n\n');
  }

  // ------------------------------------------------------- tokens -> LaTeX
  function plainToks(toks) { return (toks || []).map(function (t) { return t.t === 'text' ? t.s : t.res ? plainToks(t.res) : ''; }).join(''); }
  function fieldTex(tok, ctx) {
    var ins = tok.instr.trim();
    var m = /^REF\s+(\S+)(.*)$/i.exec(ins);
    if (m) {
      var d = decLabel(m[1]);
      if (d.kind === 'E') return '\\eqref{' + d.label + '}';
      // referência cruzada do Word com "rótulo e número" (ex.: "Tabela 1"): \ref só dá o número,
      // então o rótulo vira texto antes dele, para não sumir ao reescrever o trecho
      var shown = plainToks(tok.res).trim(), pm = /^([A-Za-zÀ-ÿ]+\.?)\s+\S/.exec(shown);
      return (pm ? pm[1] + '~' : '') + '\\ref{' + d.label + '}';
    }
    m = /^ADDIN\s+IHRCITE\s+(\{.*\})\s*$/i.exec(ins);
    if (m) {
      try {
        var ci = JSON.parse(m[1]);
        var cmd = { cite: '\\cite', textcite: '\\textcite', citeauthor: '\\citeauthor', citeyear: '\\citeyear' }[ci.m] || '\\cite';
        var opt = ci.pre ? '[' + ci.pre + '][' + (ci.post || '') + ']' : ci.post ? '[' + ci.post + ']' : '';
        return cmd + opt + '{' + (ci.k || []).join(',') + '}';
      } catch (e) { return tokensToTex(tok.res, ctx); }
    }
    m = /^PAGEREF\s+(\S+)/i.exec(ins);
    if (m) return '\\pageref{' + decLabel(m[1]).label + '}';
    m = /^HYPERLINK\s+"([^"]+)"/i.exec(ins);
    if (m) return '\\href{' + m[1] + '}{' + tokensToTex(tok.res, ctx) + '}';
    if (/^NOTEREF/i.test(ins)) return tokensToTex(tok.res, ctx);
    return tokensToTex(tok.res, ctx);
  }
  var WR = [['b', '\\textbf{'], ['i', '\\textit{'], ['u', '\\underline{'], ['s', '\\sout{'], ['sc', '\\textsc{'], ['tt', '\\texttt{'], ['sup', '\\textsuperscript{'], ['sub', '\\textsubscript{']];
  function wrapFmt(s, f) {
    if (!s) return s;
    var lead = /^\s*/.exec(s)[0], trail = /\s*$/.exec(s)[0];
    var core = s.slice(lead.length, s.length - trail.length);
    if (!core) return s;
    WR.slice().reverse().forEach(function (w) { if (f[w[0]]) core = w[1] + core + '}'; });
    return lead + core + trail;
  }
  function sameFmt(a, b) { return WR.every(function (w) { return !!a[w[0]] === !!b[w[0]]; }); }
  function tokensToTex(toks, ctx) {
    var out = '', run = '', rf = null;
    function flush() { if (run) { out += wrapFmt(run, rf || {}); run = ''; } rf = null; }
    toks.forEach(function (t) {
      if (t.t === 'text') {
        if (rf && !sameFmt(rf, t.f)) flush();
        if (!rf) rf = t.f;
        run += texEsc(t.s);
        return;
      }
      flush();
      switch (t.t) {
        case 'math': out += '$' + t.tex + '$'; break;
        case 'dmath': out += '$' + t.tex + '$'; break;
        case 'field': out += fieldTex(t, ctx); break;
        case 'br': out += '\\\\\n'; break;
        case 'page': out += '\n\\newpage\n'; break;
        case 'img': out += t.tex; break;
        case 'fn': out += '\\footnote{' + t.tex + '}'; break;
        case 'link':
          var inner = tokensToTex(t.res, ctx);
          out += (inner === texEsc(t.url) ? '\\url{' + t.url + '}' : '\\href{' + t.url + '}{' + inner + '}');
          break;
        default: break;
      }
    });
    flush();
    return out.replace(/[ \t]+/g, function (m) { return m; });
  }
  function labelsOf(toks) {
    var ls = [];
    toks.forEach(function (t) {
      if (t.t === 'bm' && !/^_(Toc|GoBack|Hlk|LtxEdit|RefB_)/.test(t.name)) ls.push(decLabel(t.name).label);
      if (t.res) labelsOf(t.res).forEach(function (l) { ls.push(l); });
    });
    return ls;
  }

  // ------------------------------------------------------- blocos
  function pInfo(p, ctx) {
    var ppr = X.kid(p, 'w:pPr');
    var st = ppr ? X.kid(ppr, 'w:pStyle') : null;
    var role = ctx.role(st ? X.val(st) : 'Normal');
    var np = ppr ? X.kid(ppr, 'w:numPr') : null;
    var numId = np && X.kid(np, 'w:numId') ? X.val(X.kid(np, 'w:numId')) : null;
    var ilvl = np && X.kid(np, 'w:ilvl') ? parseInt(X.val(X.kid(np, 'w:ilvl')), 10) : 0;
    var jc = ppr && X.kid(ppr, 'w:jc') ? X.val(X.kid(ppr, 'w:jc')) : null;
    if (ppr && X.kid(ppr, 'w:sectPr')) ctx.flagUnsafe('quebra de seção');
    // capa do template v4 também reconhecida pelo quadro (caso o nome do estilo não venha)
    var fr = ppr ? X.kid(ppr, 'w:framePr') : null;
    if (fr && role.kind === 'normal' && X.attr(fr, 'w:vAnchor') === 'page') {
      var y = parseInt(X.attr(fr, 'w:y') || '0', 10);
      role = { kind: 'cover', field: y < 8000 ? 'titlesub' : 'localdata' };
    }
    return { role: role, numId: numId, ilvl: ilvl, jc: jc };
  }
  function plainText(el) { return X.find(el, 'w:t').map(X.text).join(''); }
  function hasField(toks, re) {
    return toks.some(function (t) { return (t.t === 'field' && re.test(t.instr.trim())) || (t.res && hasField(t.res, re)); });
  }
  function seqName(toks) {
    for (var i = 0; i < toks.length; i++) {
      var t = toks[i];
      if (t.t === 'field') { var m = /^SEQ\s+(\S+)/i.exec(t.instr.trim()); if (m) return { name: m[1], idx: i }; }
    }
    return null;
  }
  function isEquationTable(tbl) {
    var rows = X.kids(tbl, 'w:tr');
    if (rows.length !== 1) return false;
    var cells = X.kids(rows[0], 'w:tc');
    if (cells.length < 2 || cells.length > 3) return false;
    return X.find(tbl, 'm:oMath').length > 0 && cells.every(function (c) {
      var tx = plainText(c).replace(/\s/g, '');
      return tx === '' || /^\(?\d*\)?$/.test(tx) || X.find(c, 'm:oMath').length;
    });
  }
  function equationFromTable(tbl, ctx) {
    var om = X.find(tbl, 'm:oMathPara')[0] || X.find(tbl, 'm:oMath')[0];
    var tex = M.ommlToLatex(om);
    var allToks = inlineTokens(X.find(tbl, 'w:p').reduce(function (a, p) { return a.concat(p.children); }, []), ctx);
    var numbered = hasField(allToks, /^SEQ/i) || /\(\d+\)/.test(plainText(tbl));
    var lbl = labelsOf(allToks)[0];
    return numbered
      ? '\\begin{equation}\n  ' + tex + (lbl ? '\n  \\label{' + lbl + '}' : '') + '\n\\end{equation}'
      : '\\[\n  ' + tex + '\n\\]';
  }

  function tableTex(tbl, kind, ctx) {
    var grid = X.find(X.kid(tbl, 'w:tblGrid') || { children: [] }, 'w:gridCol').map(function (g) { return parseInt(X.attr(g, 'w:w'), 10) || 0; });
    var rows = X.kids(tbl, 'w:tr');
    var ts = X.kid(X.kid(tbl, 'w:tblPr') || { children: [] }, 'w:tblStyle');
    var styleId = ts ? X.val(ts) : '';
    if (!kind) kind = /SimplesTabela3|PlainTable3/i.test(styleId) ? 'table' : 'quadro';
    var ncol = grid.length || Math.max.apply(null, rows.map(function (r) { return X.kids(r, 'w:tc').length; }));
    // alinhamento por coluna (1ª linha de dados)
    var aligns = [];
    var sample = rows[1] || rows[0];
    var ci = 0;
    X.kids(sample, 'w:tc').forEach(function (tc) {
      var p = X.kid(tc, 'w:p'); var info = p ? pInfo(p, ctx) : {};
      var a = info.jc === 'center' ? 'c' : (info.jc === 'right' || info.jc === 'end') ? 'r' : (info.jc === 'left' || info.jc === 'start' || info.jc === 'both') ? 'l' : (kind === 'table' ? 'l' : 'c');
      var span = X.kid(X.kid(tc, 'w:tcPr') || { children: [] }, 'w:gridSpan');
      var n = span ? parseInt(X.val(span), 10) : 1;
      for (var q = 0; q < n; q++) aligns[ci++] = a;
    });
    while (aligns.length < ncol) aligns.push(kind === 'table' ? 'c' : 'c');
    var spec;
    if (kind === 'table') spec = 'N' + aligns.slice(1).join('');
    else {
      var total = grid.reduce(function (a, b) { return a + b; }, 0);
      var uneven = grid.length && grid.some(function (g) { return Math.abs(g - total / grid.length) > 0.15 * total / grid.length; });
      spec = '|' + aligns.map(function (a, k) {
        if (uneven) return '>{' + (a === 'c' ? '\\centering' : a === 'r' ? '\\raggedleft' : '\\raggedright') + '\\arraybackslash}p{' + (Math.round((grid[k] - 216) / 566.93 * 100) / 100) + 'cm}';
        return a;
      }).join('|') + '|';
    }
    var lines = [];
    rows.forEach(function (r, ri) {
      var cells = X.kids(r, 'w:tc').map(function (tc, k) {
        var tex = X.kids(tc, 'w:p').map(function (p) { return tokensToTex(inlineTokens(p.children, ctx), ctx).trim(); }).filter(Boolean).join(' \\newline ');
        X.kids(tc, 'w:tbl').forEach(function () { ctx.warn('Tabela dentro de tabela não é suportada.'); });
        if (kind === 'table' && ri === 0 && tex) tex = '\\cab{' + tex.replace(/^\\textbf\{([^{}]*)\}$/, '$1') + '}';
        var span = X.kid(X.kid(tc, 'w:tcPr') || { children: [] }, 'w:gridSpan');
        var n = span ? parseInt(X.val(span), 10) : 1;
        if (n > 1) tex = '\\multicolumn{' + n + '}{' + (kind === 'table' ? 'c' : '|c|') + '}{' + tex + '}';
        else if (kind === 'table' && ri === 0 && k === 0) tex = '\\multicolumn{1}{l}{' + tex + '}';
        return tex;
      });
      lines.push('    ' + cells.join(' & ') + ' \\\\' + (kind === 'table' ? (ri === 0 ? ' \\linhaihr' : '') : ' \\hline'));
    });
    var head = kind === 'table' ? '  \\rowcolors{2}{white}{ihrcinza}\n  \\begin{tabular}{' + spec + '}\n' : '  \\begin{tabular}{' + spec + '}\n    \\hline\n';
    return head + lines.join('\n') + '\n  \\end{tabular}';
  }

  var KIND = { Figura: 'figure', Figure: 'figure', Tabela: 'table', Table: 'table', Quadro: 'quadro' };
  var ENV = { 1: 'chapter', 2: 'section', 3: 'subsection', 4: 'subsubsection', 5: 'paragraph', 6: 'subparagraph' };

  function countParas(el) {
    if (!el || !el.name || el.name === 'w:txbxContent' || el.name === 'mc:Fallback' || el.name === 'w:del') return 0;
    if (el.name === 'w:p') return 1;
    var c = 0; (el.children || []).forEach(function (k) { c += countParas(k); }); return c;
  }
  function flatten(nodes) {
    var out = [];
    nodes.forEach(function (n) {
      if (n.name === 'w:sdt') {
        var gal = X.find(n, 'w:docPartGallery')[0];
        if (gal && /Table of Contents/i.test(X.val(gal))) { out.push(n); return; }
        var sc = X.kid(n, 'w:sdtContent');
        out = out.concat(flatten(sc ? sc.children : []));
        return;
      }
      if (n.name) out.push(n);
    });
    return out;
  }
  // texto de cada parágrafo, na mesma ordem usada para contar (p0/p1 dos blocos)
  function paraTexts(nodes) {
    var out = [];
    function txt(el) {
      if (!el || !el.name || el.name === 'w:txbxContent' || el.name === 'mc:Fallback' || el.name === 'w:del') return '';
      if (el.name === 'w:t') return X.text(el);
      var s = ''; (el.children || []).forEach(function (k) { s += txt(k); }); return s;
    }
    function walk(el) {
      if (!el || !el.name || el.name === 'w:txbxContent' || el.name === 'mc:Fallback' || el.name === 'w:del') return;
      if (el.name === 'w:p') { out.push(txt(el)); return; }
      (el.children || []).forEach(walk);
    }
    flatten(nodes).forEach(walk);
    return out;
  }
  function isSectPara(n) { var ppr = n && n.name === 'w:p' ? X.kid(n, 'w:pPr') : null; return !!(ppr && X.kid(ppr, 'w:sectPr')); }
  var SECT_MARK = '% --- quebra de seção do Word (não remova esta linha) ---';

  function bodyToLatex(nodes, ctx) {
    nodes = flatten(nodes);
    var out = [], meta = [];
    var i = 0;
    var list = null;   // {items: [...], stack}
    var lastCover = null;
    var tocHead = null;
    var pStart = [], pCnt = [], acc = 0;
    nodes.forEach(function (n, k) { pStart[k] = acc; pCnt[k] = countParas(n); acc += pCnt[k]; });
    ctx.paraTotal = acc;
    function closeList() {
      if (!list) return;
      out.push(renderList(list.root));
      meta[out.length - 1] = { n0: list.n0, n1: list.n1 };
      list = null;
    }
    function renderList(node, indent) {
      indent = indent || '';
      var env = node.ordered ? 'enumerate' : 'itemize';
      var s = indent + '\\begin{' + env + '}\n';
      node.items.forEach(function (it) {
        s += indent + '  \\item ' + it.text.join('\n\n' + indent + '    ') + '\n';
        it.children.forEach(function (ch) { s += renderList(ch, indent + '    ') + '\n'; });
      });
      return s + indent + '\\end{' + env + '}';
    }
    function addItem(info, text) {
      var fmt = (ctx.numFmt[info.numId] || {})[String(info.ilvl)] || (info.numId ? 'decimal' : 'bullet');
      var ordered = fmt !== 'bullet' && fmt !== 'none';
      if (!list) list = { root: { ordered: ordered, items: [], lvl: 0 }, path: [], n0: i };
      list.n1 = i;
      var node = list.root, lvl = 0;
      while (lvl < info.ilvl) {
        var last = node.items[node.items.length - 1];
        if (!last) { last = { text: [''], children: [] }; node.items.push(last); }
        var child = last.children[last.children.length - 1];
        if (!child || lvl + 1 > info.ilvl) { child = { ordered: ordered, items: [] }; last.children.push(child); }
        node = child; lvl++;
      }
      node.items.push({ text: [text], children: [] });
    }
    function addContinuation(text) {
      list.n1 = i;
      var node = list.root;
      while (true) {
        var last = node.items[node.items.length - 1];
        if (last && last.children.length) { node = last.children[last.children.length - 1]; continue; }
        if (last) last.text.push(text);
        break;
      }
    }

    var iterStart = 0, outLen = 0, unsafeLen = 0, extendPrev = false, stale = false;
    function endIter() {
      // metadados dos blocos criados nesta iteração (nós iterStart .. i-1)
      var hasSect = false;
      for (var q = iterStart; q < i; q++) if (isSectPara(nodes[q])) hasSect = true;
      // entradas geradas do .bib pertencem ao bloco anterior (\printbibliography / \refmanual)
      var ext = extendPrev; extendPrev = false;
      if (ext && !hasSect && out.length === outLen && outLen > 0 && meta[outLen - 1]) { meta[outLen - 1].n1 = i - 1; return; }
      var newUnsafe = ctx.unsafe.length > unsafeLen;
      if (hasSect && out.length === outLen && !list) { out.push(SECT_MARK); }
      for (var z = outLen; z < out.length; z++) {
        if (!meta[z]) meta[z] = { n0: iterStart, n1: i - 1 };
        if (hasSect) meta[z].fixed = true;
        if (newUnsafe) meta[z].unsafe = true;
        if (stale) meta[z].stale = true;
      }
      stale = false;
    }
    while (i < nodes.length) {
      if (i > iterStart || out.length > outLen) endIter();
      iterStart = i; outLen = out.length; unsafeLen = ctx.unsafe.length;
      var n = nodes[i];
      if (!n.name) { i++; continue; }
      if (n.name === 'w:sectPr') { i++; continue; }
      if (n.name === 'w:sdt') {
        closeList(); outLen = out.length; out.push('\\tableofcontents'); i++; continue;
      }
      if (n.name === 'w:tbl') {
        closeList();
        if (isEquationTable(n)) out.push(equationFromTable(n, ctx));
        else out.push('\\begin{center}\n' + tableTex(n, null, ctx) + '\n\\end{center}');
        i++; continue;
      }
      if (n.name !== 'w:p') { i++; continue; }

      var info = pInfo(n, ctx);
      var toks = inlineTokens(n.children, ctx);
      // sumário (campo TOC que pode atravessar vários parágrafos)
      if (hasField(toks, /^TOC/i) || X.find(n, 'w:instrText').some(function (it) { return /^\s*TOC/i.test(X.text(it)); })) {
        closeList(); outLen = out.length;
        var tn0 = tocHead !== null ? tocHead : i;
        if (out.length && /^SUM[AÁ]RIO$/i.test(out[out.length - 1].trim())) { tn0 = meta[out.length - 1].n0; out.pop(); meta.pop(); outLen = out.length; }
        out.push('\\tableofcontents');
        var depth = 0, k = i;
        for (; k < nodes.length; k++) {
          X.find(nodes[k], 'w:fldChar').forEach(function (fc) { var ty = X.attr(fc, 'w:fldCharType'); if (ty === 'begin') depth++; else if (ty === 'end') depth--; });
          if (depth <= 0) break;
        }
        i = k + 1;
        meta[out.length - 1] = { n0: tn0, n1: Math.min(k, nodes.length - 1) };
        iterStart = tn0; tocHead = null;
        continue;
      }
      if (info.role.kind === 'tochead') { tocHead = i; i++; continue; }
      if (info.role.kind === 'toc') { i++; continue; }
      if (info.role.kind === 'cover') {
        var ctex = tokensToTex(toks, ctx).trim().replace(/\\\\\n?/g, ' ');
        var fld = info.role.field;
        if (fld === 'localdata') fld = lastCover === 'local' ? 'date' : 'local';
        if (fld === 'titlesub') fld = lastCover === 'title' ? 'subtitulo' : 'title';
        out.push('\\' + fld + '{' + ctex + '}');
        lastCover = fld; i++; continue;
      }
      if (lastCover && !plainText(n).trim() && !toks.some(function (t) { return t.t !== 'text' && t.t !== 'bm'; })) {
        out.push('\\maketitle'); lastCover = null; i++; continue;
      }
      lastCover = null;
      var tex = tokensToTex(toks, ctx).trim();
      var onlyDMath = toks.some(function (t) { return t.t === 'dmath'; }) && toks.every(function (t) { return t.t === 'dmath' || t.t === 'bm' || (t.t === 'text' && !t.s.trim()); });

      // listas
      if (info.numId && info.numId !== '0' && info.role.kind !== 'heading') { addItem(info, tex); i++; continue; }
      if (list && info.role.kind === 'list') { addContinuation(tex); i++; continue; }
      closeList();

      if (info.role.kind === 'bibitem') {
        // entradas geradas do .bib (marcador _RefB_) são recriadas pela lista; as manuais viram \refmanual
        var managed = toks.some(function (t) { return t.t === 'bm' && /^_RefB_/.test(t.name); });
        if (!managed && tex) out.push('\\refmanual{' + tex.replace(/\\\\\n?/g, ' ') + '}');
        else if (managed && /^\\(printbibliography|refmanual)/.test(out[out.length - 1] || '')) extendPrev = true;
        i++; continue;
      }
      if (info.role.kind === 'heading' && info.role.level === 1 && /^REFER[EÊ]NCIAS$/i.test(plainText(n).trim())) {
        if (!info.role.unnum) stale = true;     // título antigo (Título 1): reescreve com o estilo novo
        out.push('\\printbibliography'); i++; continue;
      }
      if (info.role.kind === 'heading') {
        var lbls = labelsOf(toks);
        var star = info.numId === '0' || !!info.role.unnum;
        var cmd = ENV[Math.min(info.role.level, 6)];
        out.push('\\' + cmd + (star ? '*' : '') + '{' + tex.replace(/\\\\\n?/g, ' ') + '}' + (lbls.length ? '\\label{' + lbls[0] + '}' : ''));
        i++; continue;
      }
      if (info.role.kind === 'caption') {
        var sq = seqName(toks);
        var kind = sq ? (KIND[sq.name] || null) : null;
        if (!kind) { var mm = /^(Figura|Tabela|Quadro)\b/.exec(plainText(n)); if (mm) kind = KIND[mm[1]]; }
        if (kind) {
          var fl = floatFrom(nodes, i, toks, sq, kind, ctx);
          out.push(fl.tex); i = fl.next; continue;
        }
      }
      if (info.role.kind === 'fonte') { out.push('\\fonte{' + tex.replace(/^Fonte:\s*/i, '') + '}'); i++; continue; }
      if (onlyDMath) {
        var lb = labelsOf(toks);
        out.push('\\[\n  ' + toks.filter(function (t) { return t.t === 'dmath'; }).map(function (t) { return t.tex; }).join(' \\\\ ') + '\n\\]');
        i++; continue;
      }
      if (!tex) { i++; continue; }
      if (info.jc === 'center' && info.role.kind === 'normal' && !/\\includegraphics/.test(tex)) { out.push('\\begin{center}\n' + tex + '\n\\end{center}'); i++; continue; }
      out.push(tex);
      i++;
    }
    if (i > iterStart || out.length > outLen) endIter();
    closeList();
    return out.map(function (tex, k) {
      var m = meta[k] || { n0: nodes.length - 1, n1: nodes.length - 1 };
      var n1 = Math.min(m.n1, nodes.length - 1);
      return { tex: tex, p0: pStart[m.n0], p1: pStart[n1] + Math.max(pCnt[n1], 1) - 1, fixed: !!m.fixed || /^\\tableofcontents$/.test(tex), unsafe: !!m.unsafe, stale: !!m.stale };
    });
  }

  function floatFrom(nodes, i, toks, sq, kind, ctx) {
    var lbls = labelsOf(toks);
    var capToks = sq ? toks.slice(sq.idx + 1) : toks;
    var cap = tokensToTex(capToks, ctx).replace(/^\s*(-|–|—|:|\.)\s*/, '').trim();
    if (!sq) cap = cap.replace(/^(Figura|Tabela|Quadro)\s*\d+\s*(-|–|—|:)?\s*/, '');
    var env = kind;
    var lines = ['\\begin{' + env + '}', '  \\caption{' + cap + '}' + (lbls.length ? '\\label{' + lbls[0] + '}' : '')];
    var j = i + 1;
    var fonte = null;
    for (; j < nodes.length; j++) {
      var n = nodes[j];
      if (!n.name) continue;
      if (n.name === 'w:tbl') {
        if (isEquationTable(n)) break;
        lines.push(tableTex(n, kind === 'figure' ? 'quadro' : kind, ctx)); continue;
      }
      if (n.name !== 'w:p') break;
      var info = pInfo(n, ctx);
      var t = inlineTokens(n.children, ctx);
      if (info.role.kind === 'fonte') { fonte = tokensToTex(t, ctx).trim().replace(/^Fonte:\s*/i, ''); j++; break; }
      var hasImg = t.some(function (q) { return q.t === 'img'; });
      if (info.role.kind === 'imagem' || hasImg) {
        var tx = tokensToTex(t, ctx).trim();
        if (tx) lines.push('  ' + tx);
        continue;
      }
      break;
    }
    if (fonte !== null) lines.push('  \\fonte{' + fonte + '}');
    lines.push('\\end{' + env + '}');
    return { tex: lines.join('\n'), next: j };
  }

  // ------------------------------------------------------------ API
  function convert(pkgOrParts) {
    var P = typeof pkgOrParts === 'string' ? parsePackage(pkgOrParts) : pkgOrParts;
    var ctx = new Ctx(P);
    if (!P.document) return { latex: '', images: {}, warnings: ['Não encontrei o conteúdo do documento na seleção.'], unsafe: [] };
    var body = X.kid(P.document, 'w:body');
    var blocks = bodyToLatex(body ? body.children : [], ctx);
    var latex = blocks.map(function (b) { return b.tex; }).join('\n\n').replace(/\n{3,}/g, '\n\n').trim() + '\n';
    var pi = paraIndex(body);
    return { latex: latex, blocks: blocks, images: ctx.images, imageOrder: ctx.imageOrder, warnings: ctx.warnings, unsafe: ctx.unsafe,
      paraCount: pi.count, paraTexts: paraTexts(body ? body.children : []), sectBreaks: pi.sectBreaks, partNames: P.partNames || [] };
  }
  // divide LaTeX em blocos: linhas em branco fora de ambientes
  function splitBlocks(tex) {
    var lines = (tex || '').replace(/\r\n?/g, '\n').split('\n');
    var blocks = [], cur = [], depth = 0;
    lines.forEach(function (ln) {
      var code = ln.replace(/(^|[^\\])%.*$/, '$1');
      if (!ln.trim() && depth <= 0) { if (cur.length) { blocks.push(cur.join('\n')); cur = []; } return; }
      cur.push(ln);
      var b = (code.match(/\\begin\{/g) || []).length, e = (code.match(/\\end\{/g) || []).length;
      depth += b - e;
      if (/^\s*\\\[/.test(code)) depth++;
      if (/\\\]\s*$/.test(code)) depth--;
    });
    if (cur.length) blocks.push(cur.join('\n'));
    return blocks.map(function (x) { return x.trim(); }).filter(Boolean);
  }
  // posição (na ordem de Word.Range.paragraphs) dos parágrafos que carregam quebra de seção
  function paraIndex(body) {
    var count = 0, breaks = [];
    (function walk(el) {
      (el && el.children || []).forEach(function (c) {
        if (!c.name || c.name === 'w:txbxContent' || c.name === 'mc:Fallback' || c.name === 'w:del') return;
        if (c.name === 'w:p') {
          var ppr = X.kid(c, 'w:pPr');
          if (ppr && X.kid(ppr, 'w:sectPr')) breaks.push(count);
          count++;
          return;
        }
        walk(c);
      });
    })(body);
    return { count: count, sectBreaks: breaks };
  }

  var API = { convert: convert, parsePackage: parsePackage, decLabel: decLabel, imageSize: imageSize, splitBlocks: splitBlocks, SECT_MARK: SECT_MARK };
  R.ooxml2latex = API;
  root.IHRLX = R;
  if (typeof module !== 'undefined') module.exports = API;
})(typeof globalThis !== 'undefined' ? globalThis : this);
