// Converte os capítulos do template LaTeX da UFRGS com o próprio conversor do suplemento
const fs = require('fs'), path = require('path');
const A = '/home/claude/ihr-latex-word/lib/';
require(A + 'xml.js'); global.katex = require('/home/claude/ihr-latex-word/vendor/katex.min.js');
require(A + 'math.js'); require(A + 'template-parts.js');
const L = require(A + 'latex2ooxml.js'), B = require(A + 'bib.js');
const SRC = '/home/claude/ufrgs/src/';
// lipsum
const ltd = fs.readFileSync('/usr/share/texlive/texmf-dist/tex/latex/lipsum/lipsum.ltd.tex', 'utf8');
const LP = [null]; ltd.replace(/\\NewLipsumPar\{% \{\d+\}\n([\s\S]*?)\}%/g, (m, t) => { LP.push(t.replace(/\s+/g, ' ').trim()); });
function lipsum(spec) { const [a, b] = spec.split('-').map(Number); const out = []; for (let i = a; i <= (b || a); i++) out.push(LP[i]); return out.join('\n\n'); }
function pre(t) {
  return t.replace(/\\lipsum\[([\d-]+)\]/g, (m, s) => '\n\n' + lipsum(s) + '\n\n')
    .replace(/\\cites\{([^}]*)\}\{([^}]*)\}\{([^}]*)\}/g, '\\cite{$1,$2,$3}')
    .replace(/\\tabela\b\s*/g, 'Tabela~').replace(/\\figuras\b\s*/g, 'Figuras~').replace(/\\figura\b\s*/g, 'Figura~')
    .replace(/\\begin\{(table|figure)\}\[H\]/g, '\\begin{$1}').replace(/^\s*\\centering\s*$/mg, '');
}
const ch = ['Introdução', 'FundamentacaoTeorica', 'Metodologia', 'Resultados', 'Conclusão'].map(f => pre(fs.readFileSync(SRC + 'TCC/Paginas/' + f + '.tex', 'utf8')));
let apx = pre(fs.readFileSync(SRC + 'TCC/Paginas/Apendice.tex', 'utf8'))
  .replace(/\\addcontentsline\{[^}]*\}\{[^}]*\}\{[^}]*\}/, '').replace(/\\apendices/, '')
  .replace(/\\chapter\{Pseudocódigos\}/, '\\chapter*{A. PSEUDOCÓDIGOS}');
apx = apx.replace(/\\input\{[^}]*\}/g, '');
const src = ch.join('\n\n') + '\n\n\\printbibliography\n\n' + apx;
fs.writeFileSync(path.join(__dirname, 'conteudo.tex'), src);
const lib = new B.Library(fs.readFileSync(SRC + 'references.bib', 'utf8'));
const png = fs.readFileSync(SRC + 'assets/fig.png');
const images = { 'assets/fig.png': { base64: png.toString('base64'), contentType: 'image/png', width: png.readUInt32BE(16), height: png.readUInt32BE(20) } };
const r = L.convert(src, { bib: lib, images, eqChapter: true, captionSep: ' – ' });
fs.writeFileSync(path.join(__dirname, 'conteudo.json'), JSON.stringify({
  body: r.body, rels: r.rels, media: r.media, numbering: r.nums.length ? L.numberingXml(r.nums) : '',
  footnotes: r.footnotes.length ? L.footnotesXml(r.footnotes) : '', warnings: r.warnings }));
console.log('ok; warnings:', r.warnings, 'rels', r.rels.length, 'media', r.media.map(m => m.name));
// ---- nomenclatura (Simbolos.tex) -> parágrafos "símbolo<tab>descrição<tab>unidade"
const M = global.IHRLX.math;
const simb = fs.readFileSync(SRC + 'TCC/Paginas/Simbolos.tex', 'utf8');
const UN = { '\\mega \\pascal': 'MPa', '\\kelvin^{-1}': 'K^-1', '\\joule': 'J', '\\joule \\cdot\\kelvin^{-1}': 'J·K^-1', '\\watt \\cdot\\meter^{-1} \\cdot \\kelvin^{-1}': 'W·m^-1·K^-1' };
const ents = [];
simb.replace(/^\\nomenclature\[(\w)\]\{(.*?)\}\{(.*)\}\s*$/mg, (m, g, sym, rest) => {
  let unit = '';
  rest = rest.replace(/\\nomunit\{\\SI\{\}\{(.*)\}\}\s*$/, (mm, u) => { unit = UN[u.trim()] || u; return ''; })
    .replace(/\\MPa\s*$/, () => { unit = 'MPa'; return ''; }).replace(/\\K\s*$/, () => { unit = 'K'; return ''; });
  ents.push({ g, sym, desc: rest.trim(), unit });
});
const esc = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const run = (t, extra) => '<w:r>' + (extra ? '<w:rPr>' + extra + '</w:rPr>' : '') + '<w:t xml:space="preserve">' + esc(t) + '</w:t></w:r>';
function unitRuns(u) { return u.split(/(\^-1)/).map(s => s === '^-1' ? run('−1', '<w:vertAlign w:val="superscript"/>') : (s ? run(s) : '')).join(''); }
function symRuns(sym) {
  if (/^\$.*\$$/.test(sym)) { const r = M.latexToOmml(sym.slice(1, -1), false); if (r.ok) return r.omml; }
  return run(sym);
}
const key = e => [e.sym.startsWith('$\\') ? 0 : 1, e.sym.toLowerCase()];
const cmp = (a, b) => { const x = key(a), y = key(b); return x[0] - y[0] || (x[1] < y[1] ? -1 : x[1] > y[1] ? 1 : 0); };
let nom = '';
const P = (inner, st) => '<w:p><w:pPr><w:pStyle w:val="' + st + '"/></w:pPr>' + inner + '</w:p>';
[['A', 'Abreviaturas e acrônimos'], ['S', 'Símbolos']].forEach(([g, nome]) => {
  nom += P(run(nome), 'NomenclaturaGrupo');
  ents.filter(e => e.g === g).sort(cmp).forEach(e => {
    nom += P(symRuns(e.sym) + '<w:r><w:tab/></w:r>' + run(e.desc) + (e.unit ? '<w:r><w:tab/></w:r>' + unitRuns(e.unit) : ''), 'Nomenclatura');
  });
});
const out = JSON.parse(fs.readFileSync(path.join(__dirname, 'conteudo.json'), 'utf8'));
out.nomencl = nom; out.lipsum1 = LP[1];
fs.writeFileSync(path.join(__dirname, 'conteudo.json'), JSON.stringify(out));
console.log('nomenclatura:', ents.length, 'entradas');
