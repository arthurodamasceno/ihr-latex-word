/* Carregador do painel: define a versão e carrega o resto com ?v=VERSÃO */
(function () {
  'use strict';
  var V = '1.8.6';
  window.IHRLX_VERSION = V;
  function css(h) { var l = document.createElement('link'); l.rel = 'stylesheet'; l.href = h; document.head.appendChild(l); }
  function js(src) {
    return new Promise(function (ok, fail) {
      var s = document.createElement('script'); s.src = src; s.onload = ok;
      s.onerror = function () { fail(new Error('falha ao carregar ' + src)); };
      document.head.appendChild(s);
    });
  }
  css('taskpane.css?v=' + V);
  fetch('panel.html?v=' + V).then(function (r) { if (!r.ok) throw new Error('panel.html ' + r.status); return r.text(); })
    .then(function (h) {
      document.body.innerHTML = h;
      var list = ['vendor/katex.min.js?v=0.16.47', 'lib/xml.js', 'lib/math.js', 'lib/template-parts.js', 'lib/latex2ooxml.js', 'lib/ooxml2latex.js', 'lib/bib.js', 'taskpane.js'];
      return list.reduce(function (p, src) { return p.then(function () { return js(src.indexOf('?') > 0 ? src : src + '?v=' + V); }); }, Promise.resolve());
    })
    .catch(function (e) {
      document.body.innerHTML = '<p style="font:13px Segoe UI,sans-serif;color:#a4262c">Não foi possível carregar o painel (' + e.message + '). Verifique a conexão e reabra o painel.</p>';
    });
})();
