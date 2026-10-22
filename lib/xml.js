/* Mini parser XML (suficiente para OOXML/MathML) — funciona no navegador e no Node */
(function (root) {
  'use strict';
  var ENT = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'", nbsp: ' ' };

  function decode(s) {
    if (s.indexOf('&') < 0) return s;
    return s.replace(/&(#x[0-9a-fA-F]+|#\d+|\w+);/g, function (m, e) {
      if (e[0] === '#') {
        var cp = e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
        return String.fromCodePoint(cp);
      }
      return ENT.hasOwnProperty(e) ? ENT[e] : m;
    });
  }

  function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  // Retorna o elemento raiz: {name, attrs, children:[el|{text}]}
  function parse(src) {
    var i = 0, n = src.length;
    var rootEl = { name: '#root', attrs: {}, children: [] };
    var stack = [rootEl];
    while (i < n) {
      var lt = src.indexOf('<', i);
      if (lt < 0) lt = n;
      if (lt > i) {
        var t = src.slice(i, lt);
        if (stack.length > 1) stack[stack.length - 1].children.push({ text: decode(t) });
        i = lt;
        continue;
      }
      if (src.startsWith('<!--', i)) { i = src.indexOf('-->', i) + 3; continue; }
      if (src.startsWith('<![CDATA[', i)) {
        var e = src.indexOf(']]>', i);
        stack[stack.length - 1].children.push({ text: src.slice(i + 9, e) });
        i = e + 3; continue;
      }
      if (src[i + 1] === '?' || src[i + 1] === '!') { i = src.indexOf('>', i) + 1; continue; }
      if (src[i + 1] === '/') {
        var gt = src.indexOf('>', i);
        stack.pop();
        i = gt + 1; continue;
      }
      // tag de abertura
      var j = i + 1;
      while (j < n && !/[\s/>]/.test(src[j])) j++;
      var name = src.slice(i + 1, j);
      var attrs = {};
      var re = /\s*([^\s=/>]+)\s*=\s*("([^"]*)"|'([^']*)')/y;
      re.lastIndex = j;
      var m;
      while ((m = re.exec(src))) { attrs[m[1]] = decode(m[3] !== undefined ? m[3] : m[4]); j = re.lastIndex; }
      while (j < n && src[j] !== '>' && src[j] !== '/') j++;
      var selfClose = src[j] === '/';
      var el = { name: name, attrs: attrs, children: [] };
      stack[stack.length - 1].children.push(el);
      i = src.indexOf('>', j) + 1;
      if (!selfClose) stack.push(el);
    }
    return rootEl.children.filter(function (c) { return c.name; })[0];
  }

  function kids(el, name) {
    if (!el) return [];
    return el.children.filter(function (c) { return c.name && (!name || c.name === name); });
  }
  function kid(el, name) { return kids(el, name)[0] || null; }
  function attr(el, a) { return el && el.attrs ? (el.attrs[a] !== undefined ? el.attrs[a] : null) : null; }
  function val(el) { return attr(el, 'w:val') !== null ? attr(el, 'w:val') : attr(el, 'm:val'); }
  function text(el) {
    if (!el) return '';
    if (el.text !== undefined) return el.text;
    return el.children.map(text).join('');
  }
  function find(el, name, out) {
    out = out || [];
    if (!el || !el.children) return out;
    el.children.forEach(function (c) {
      if (!c.name) return;
      if (c.name === name) out.push(c);
      find(c, name, out);
    });
    return out;
  }
  function serialize(el) {
    if (el.text !== undefined) return esc(el.text);
    var a = Object.keys(el.attrs).map(function (k) { return ' ' + k + '="' + esc(el.attrs[k]) + '"'; }).join('');
    if (!el.children.length) return '<' + el.name + a + '/>';
    return '<' + el.name + a + '>' + el.children.map(serialize).join('') + '</' + el.name + '>';
  }

  var X = { parse: parse, esc: esc, decode: decode, kids: kids, kid: kid, attr: attr, val: val, text: text, find: find, serialize: serialize };
  root.IHRLX = root.IHRLX || {};
  root.IHRLX.xml = X;
  if (typeof module !== 'undefined') module.exports = X;
})(typeof globalThis !== 'undefined' ? globalThis : this);
