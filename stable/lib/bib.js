/* BibTeX -> referências e citações no padrão ABNT (NBR 6023 / NBR 10520, sistema autor-data) */
(function (root) {
  'use strict';
  var R = root.IHRLX || {};

  // ------------------------------------------------------------------ parser
  function parse(src) {
    var entries = {}, order = [], strings = { jan: 'jan', feb: 'fev', mar: 'mar', apr: 'abr', may: 'maio', jun: 'jun', jul: 'jul', aug: 'ago', sep: 'set', oct: 'out', nov: 'nov', dec: 'dez' };
    var i = 0, n = src.length;
    function ws() { while (i < n && /\s/.test(src[i])) i++; }
    function braced(open, close) {           // src[i] === open
      var depth = 0, s = i;
      for (; i < n; i++) {
        if (src[i] === '\\') { i++; continue; }
        if (src[i] === open) depth++;
        else if (src[i] === close) { depth--; if (depth === 0) { i++; return src.slice(s + 1, i - 1); } }
      }
      return src.slice(s + 1);
    }
    function value() {
      var parts = [];
      while (i < n) {
        ws();
        if (src[i] === '{') parts.push(braced('{', '}'));
        else if (src[i] === '"') {
          var s = ++i, depth = 0;
          for (; i < n; i++) { if (src[i] === '{') depth++; else if (src[i] === '}') depth--; else if (src[i] === '"' && depth === 0) break; }
          parts.push(src.slice(s, i)); i++;
        } else {
          var m = /^[A-Za-z0-9_:.\-+/]+/.exec(src.slice(i));
          if (!m) break;
          i += m[0].length;
          parts.push(strings.hasOwnProperty(m[0].toLowerCase()) ? strings[m[0].toLowerCase()] : m[0]);
        }
        ws();
        if (src[i] === '#') { i++; continue; }
        break;
      }
      return parts.join('');
    }
    while (i < n) {
      var at = src.indexOf('@', i);
      if (at < 0) break;
      i = at + 1;
      var tm = /^\s*([A-Za-z]+)\s*/.exec(src.slice(i));
      if (!tm) continue;
      var type = tm[1].toLowerCase();
      i += tm[0].length;
      var open = src[i];
      if (open !== '{' && open !== '(') continue;
      var close = open === '{' ? '}' : ')';
      if (type === 'comment' || type === 'preamble') { braced(open, close); continue; }
      i++;
      if (type === 'string') {
        ws(); var sm = /^[A-Za-z0-9_\-]+/.exec(src.slice(i)); if (!sm) continue;
        i += sm[0].length; ws(); if (src[i] === '=') i++;
        strings[sm[0].toLowerCase()] = value(); ws(); if (src[i] === close) i++;
        continue;
      }
      ws();
      var km = /^[^,\s}]+/.exec(src.slice(i));
      if (!km) continue;
      var key = km[0]; i += key.length; ws();
      var fields = {};
      while (i < n) {
        ws();
        if (src[i] === ',') { i++; continue; }
        if (src[i] === close) { i++; break; }
        var fm = /^[A-Za-z0-9_\-]+/.exec(src.slice(i));
        if (!fm) { i++; continue; }
        i += fm[0].length; ws();
        if (src[i] !== '=') continue;
        i++;
        fields[fm[0].toLowerCase()] = value().replace(/\s+/g, ' ').trim();
      }
      if (!entries[key]) order.push(key);
      entries[key] = { key: key, type: type, fields: fields };
    }
    return { entries: entries, order: order };
  }

  // ------------------------------------------------------- LaTeX -> texto
  var ACC = { "'": '́', '`': '̀', '^': '̂', '~': '̃', '"': '̈', 'c': '̧', '=': '̄', '.': '̇', 'u': '̆', 'v': '̌', 'H': '̋', 'r': '̊' };
  function text(s) {
    if (!s) return '';
    s = s.replace(/\{\\([`'^"~=.uvcHr])\s*\{?(\\?[A-Za-z])\}?\}/g, function (m, a, c) { return (c.replace('\\', '') + ACC[a]).normalize('NFC'); })
      .replace(/\\([`'^"~=.])\s*\{?(\\?[A-Za-z])\}?/g, function (m, a, c) { return (c.replace('\\', '') + ACC[a]).normalize('NFC'); })
      .replace(/\\([uvcHr])\{(\\?[A-Za-z])\}/g, function (m, a, c) { return (c.replace('\\', '') + ACC[a]).normalize('NFC'); })
      .replace(/\\(ss|ae|AE|oe|OE|o|O|l|L|i|aa|AA)\b\s*/g, function (m, c) { return { ss: 'ß', ae: 'æ', AE: 'Æ', oe: 'œ', OE: 'Œ', o: 'ø', O: 'Ø', l: 'ł', L: 'Ł', i: 'ı', aa: 'å', AA: 'Å' }[c]; })
      .replace(/\\&/g, '&').replace(/\\%/g, '%').replace(/\\_/g, '_').replace(/\\\$/g, '$').replace(/\\#/g, '#')
      .replace(/---/g, '—').replace(/--/g, '–').replace(/~/g, ' ')
      .replace(/\\(textit|textbf|emph|textsc|textrm|texttt|mathrm|text|url|mbox)\s*/g, '')
      .replace(/\\[a-zA-Z]+\s*/g, '')
      .replace(/[{}]/g, '');
    return s.replace(/\s+/g, ' ').trim();
  }

  // ----------------------------------------------------------- autores
  function splitAnd(s) {
    var out = [], depth = 0, last = 0, re = /\s+and\s+/gi, m;
    var masks = s.replace(/\{[^{}]*\}/g, function (q) { return q.replace(/./g, '_'); });
    while ((m = re.exec(masks))) { out.push(s.slice(last, m.index)); last = m.index + m[0].length; }
    out.push(s.slice(last));
    return out.map(function (x) { return x.trim(); }).filter(Boolean);
  }
  function names(s) {
    if (!s) return [];
    var others = false;
    var list = splitAnd(s).filter(function (p) { if (/^others$/i.test(p)) { others = true; return false; } return true; });
    var res = list.map(function (p) {
      if (/^\{.*\}$/.test(p) && p.indexOf(',') < 0) return { last: text(p), first: '', corp: true };
      var parts = p.split(/,(?![^{]*\})/);
      if (parts.length >= 2) return { last: text(parts[0]), first: text(parts.slice(1).join(',')) };
      var toks = p.replace(/\s+/g, ' ').split(' ');
      // partículas (da, de, do, dos, van, von) ficam com o prenome no padrão brasileiro
      var last = toks.pop();
      return { last: text(last), first: text(toks.join(' ')) };
    });
    if (others) res.others = true;
    return res;
  }
  function up(s) { return s.toLocaleUpperCase('pt-BR'); }

  // -------------------------------------------------------- formatação
  var MES = { 1: 'jan.', 2: 'fev.', 3: 'mar.', 4: 'abr.', 5: 'maio', 6: 'jun.', 7: 'jul.', 8: 'ago.', 9: 'set.', 10: 'out.', 11: 'nov.', 12: 'dez.' };
  function month(m) {
    if (!m) return '';
    var k = parseInt(m, 10);
    if (k >= 1 && k <= 12) return MES[k];
    var t = text(m).toLowerCase().slice(0, 3);
    var map = { jan: 1, fev: 2, feb: 2, mar: 3, abr: 4, apr: 4, mai: 5, may: 5, jun: 6, jul: 7, ago: 8, aug: 8, set: 9, sep: 9, out: 10, oct: 10, nov: 11, dez: 12, dec: 12 };
    return map[t] ? MES[map[t]] : text(m);
  }
  function dateLong(d) {     // 2024-05-10 -> 10 maio 2024
    var m = /^(\d{4})-(\d{1,2})(?:-(\d{1,2}))?/.exec(d || '');
    if (!m) return text(d);
    return (m[3] ? parseInt(m[3], 10) + ' ' : '') + MES[parseInt(m[2], 10)] + ' ' + m[1];
  }
  function authorBlock(list) {
    if (!list.length) return '';
    var fmt = function (a) {
      if (a.corp) return up(a.last);
      return up(a.last) + (a.first ? ', ' + a.first : '');
    };
    var s = list.length > 3 ? fmt(list[0]) + ' et al' : list.map(fmt).join('; ') + (list.others ? ' et al' : '');
    return s.replace(/\.$/, '') + '. ';
  }
  function yearOf(e, suffix) {
    var y = text(e.fields.year || (e.fields.date || '').slice(0, 4)) || '[s. d.]';
    return y + (suffix || '');
  }
  function titleParts(t) {
    var tt = text(t);
    var m = /^(.*?):\s+(.*)$/.exec(tt);
    return m ? { main: m[1], sub: m[2] } : { main: tt, sub: '' };
  }
  // segmentos: [{t:'texto', b:true?}]
  function reference(e, suffix) {
    var f = e.fields, seg = [];
    function T(s, b) { if (s) seg.push({ t: s, b: !!b }); }
    var au = names(f.author);
    var ed = names(f.editor);
    var y = yearOf(e, suffix);
    var tp = titleParts(f.title || '');
    var addr = text(f.address || f.location || '');
    var pub = text(f.publisher || f.organization || f.institution || f.school || '');
    var edition = f.edition ? (/^\d+$/.test(f.edition.trim()) ? f.edition.trim() + '. ed. ' : text(f.edition) + ' ') : '';
    var pages = f.pages ? text(f.pages).replace(/\s*[-–—]+\s*/, '-') : '';
    var imprint = (addr || '[S. l.]') + ': ' + (pub || '[s. n.]') + ', ' + y + '.';
    var noAuthor = !au.length;
    if (!noAuthor) T(authorBlock(au));
    else if (ed.length && e.type !== 'inproceedings') { T(authorBlock(ed).replace(/\. $/, ' (ed.). ')); }

    function boldTitle() {
      if (noAuthor && !ed.length) {
        var w = tp.main.split(' '); T(up(w[0]) + (w.length > 1 ? ' ' + w.slice(1).join(' ') : '') + (tp.sub ? ': ' + tp.sub : '') + '. ');
      } else { T(tp.main, true); T((tp.sub ? ': ' + tp.sub : '') + '. '); }
    }
    function plainTitle() {
      if (noAuthor) { var w = tp.main.split(' '); T(up(w[0]) + (w.length > 1 ? ' ' + w.slice(1).join(' ') : '') + (tp.sub ? ': ' + tp.sub : '') + '. '); }
      else T(tp.main + (tp.sub ? ': ' + tp.sub : '') + '. ');
    }
    switch (e.type) {
      case 'article':
        plainTitle();
        T(text(f.journal || f.journaltitle), true);
        var art = [];
        if (addr) art.push(addr);
        if (f.volume) art.push('v. ' + text(f.volume));
        if (f.number || f.issue) art.push('n. ' + text(f.number || f.issue));
        if (pages) art.push('p. ' + pages);
        art.push((month(f.month) ? month(f.month) + ' ' : '') + y);
        T(', ' + art.join(', ') + '.');
        break;
      case 'inproceedings': case 'conference':
        plainTitle();
        T('In: ' + up(text(f.booktitle || f.eventtitle || '')) + ', ' + y + (f.venue ? ', ' + text(f.venue) : '') + '. ');
        T('Anais', true); T(' [...]. ' + imprint.replace(/\.$/, '') + (pages ? '. p. ' + pages : '') + '.');
        break;
      case 'incollection': case 'inbook':
        plainTitle();
        T('In: ' + (ed.length ? authorBlock(ed).replace(/\. $/, ' (ed.). ') : ''));
        T(text(f.booktitle || f.title), true); T('. ' + edition + imprint.replace(/\.$/, '') + (pages ? '. p. ' + pages : '') + '.');
        break;
      case 'phdthesis': case 'mastersthesis': case 'thesis':
        boldTitle();
        var kind = text(f.type) || (e.type === 'phdthesis' ? 'Tese (Doutorado)' : e.type === 'mastersthesis' ? 'Dissertação (Mestrado)' : 'Trabalho acadêmico');
        T(y + '. ' + kind + ' – ' + (pub || '[s. n.]') + (addr ? ', ' + addr : '') + ', ' + y + '.');
        break;
      case 'techreport': case 'report':
        boldTitle();
        T(imprint + (f.number ? ' (' + (text(f.type) || 'Relatório técnico') + ', n. ' + text(f.number) + ').' : ''));
        break;
      case 'misc': case 'online': case 'electronic': case 'www': case 'webpage': case 'unpublished':
        boldTitle();
        if (f.howpublished && !/\\url|http/.test(f.howpublished)) T(text(f.howpublished) + ', ');
        T(y + '.');
        break;
      default:          // book, manual, standard, booklet, proceedings, collection
        boldTitle();
        T(edition + imprint);
    }
    var url = f.url || (/\\url\{([^}]*)\}/.exec(f.howpublished || '') || [])[1];
    if (url) T(' Disponível em: ' + url.trim() + '.' + (f.urldate ? ' Acesso em: ' + dateLong(f.urldate) + '.' : ''));
    if (f.doi) T(' DOI: ' + text(f.doi) + '.');
    if (f.note) T(' ' + text(f.note).replace(/\.?$/, '.'));
    // junta segmentos de mesmo estilo e limpa pontuação duplicada
    var out = [];
    seg.forEach(function (s) {
      var last = out[out.length - 1];
      if (last && last.b === s.b) last.t += s.t; else out.push({ t: s.t, b: s.b });
    });
    out.forEach(function (s) { s.t = s.t.replace(/\[\.\.\.\]/g, '\u0001').replace(/\.\s*\./g, '.').replace(/\u0001/g, '[...]').replace(/\s+/g, ' ').replace(/ ,/g, ','); });
    return out;
  }

  // --------------------------------------------------------- citações
  function citeNames(e) {
    var au = names(e.fields.author);
    if (!au.length) au = names(e.fields.editor);
    if (!au.length) { var w = titleParts(e.fields.title || '').main.split(' ')[0] || e.key; return { paren: up(w) + (titleParts(e.fields.title || '').main.split(' ').length > 1 ? ' [...]' : ''), text: w }; }
    var L = au.map(function (a) { return a.last; });
    var et = au.length > 3 || au.others;
    var paren = et ? up(L[0]) + ' et al.' : L.map(up).join('; ');
    var txt = et ? L[0] + ' et al.' : L.length === 1 ? L[0] : L.slice(0, -1).join(', ') + ' e ' + L[L.length - 1];
    return { paren: paren, text: txt };
  }

  function Library(src) {
    this.src = src || '';
    var p = parse(this.src);
    this.entries = p.entries; this.order = p.order;
    this._suffix();
  }
  Library.prototype._suffix = function () {
    var groups = {}, self = this;
    this.suffix = {};
    this.order.forEach(function (k) {
      var e = self.entries[k];
      var g = citeNames(e).paren + '|' + yearOf(e);
      (groups[g] = groups[g] || []).push(k);
    });
    Object.keys(groups).forEach(function (g) {
      var ks = groups[g];
      if (ks.length < 2) return;
      ks.sort(function (a, b) { return text(self.entries[a].fields.title || '').localeCompare(text(self.entries[b].fields.title || ''), 'pt-BR'); });
      ks.forEach(function (k, i) { self.suffix[k] = String.fromCharCode(97 + i); });
    });
  };
  Library.prototype.merge = function (src) {
    var p = parse(src), self = this, added = 0, updated = 0;
    p.order.forEach(function (k) { if (self.entries[k]) updated++; else { added++; self.order.push(k); } self.entries[k] = p.entries[k]; });
    this.src = this.serialize();
    this._suffix();
    return { added: added, updated: updated };
  };
  Library.prototype.has = function (k) { return !!this.entries[k]; };
  Library.prototype.count = function () { return this.order.length; };
  Library.prototype.reference = function (k) { return this.entries[k] ? reference(this.entries[k], this.suffix[k]) : null; };
  Library.prototype.referenceText = function (k) { var r = this.reference(k); return r ? r.map(function (s) { return s.t; }).join('') : ''; };
  Library.prototype.label = function (k) {
    var e = this.entries[k];
    if (!e) return k;
    var cn = citeNames(e);
    return cn.text + ' (' + yearOf(e, this.suffix[k]) + ') — ' + text(e.fields.title || '');
  };
  // mode: cite | textcite | citeauthor | citeyear
  Library.prototype.cite = function (keys, mode, pre, post) {
    var self = this;
    var items = keys.map(function (k) {
      var e = self.entries[k];
      if (!e) return { miss: k };
      return { n: citeNames(e), y: yearOf(e, self.suffix[k]) };
    });
    var miss = items.filter(function (x) { return x.miss; }).map(function (x) { return x.miss; });
    var ok = items.filter(function (x) { return !x.miss; });
    var post2 = post ? ', ' + post : '';
    var s;
    if (!ok.length) s = '(' + keys.join('; ') + '?)';
    else if (mode === 'citeauthor') s = ok.map(function (x) { return x.n.text; }).join('; ');
    else if (mode === 'citeyear') s = ok.map(function (x) { return x.y; }).join('; ');
    else if (mode === 'textcite') {
      var parts = ok.map(function (x, i) { return x.n.text + ' (' + x.y + (i === ok.length - 1 ? post2 : '') + ')'; });
      s = (pre ? pre + ' ' : '') + (parts.length > 1 ? parts.slice(0, -1).join(', ') + ' e ' + parts[parts.length - 1] : parts[0]);
    } else {
      s = '(' + (pre ? pre + ' ' : '') + ok.map(function (x) { return x.n.paren + ', ' + x.y; }).join('; ') + post2 + ')';
    }
    return { text: s, missing: miss };
  };
  Library.prototype.serialize = function () {
    var self = this;
    return this.order.map(function (k) {
      var e = self.entries[k];
      return '@' + e.type + '{' + k + ',\n' + Object.keys(e.fields).map(function (f) { return '  ' + f + ' = {' + e.fields[f] + '}'; }).join(',\n') + '\n}';
    }).join('\n\n') + '\n';
  };
  Library.prototype.sortKey = function (k) { return this.referenceText(k).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase(); };

  var B = { parse: parse, names: names, text: text, reference: reference, Library: Library,
    sortKey: function (s) { return s.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase(); } };
  R.bib = B;
  root.IHRLX = R;
  if (typeof module !== 'undefined') module.exports = B;
})(typeof globalThis !== 'undefined' ? globalThis : this);
