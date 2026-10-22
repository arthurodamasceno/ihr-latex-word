// Monta um .docx de teste: capa+sumário do template + corpo gerado pelo conversor
global.katex = require('../vendor/katex.min.js');
const fs = require('fs');
const JSZip = require('/opt/npm-tools/node_modules/jszip');
const L = require('../lib/latex2ooxml.js');
const B = require('../lib/bib.js');
(async () => {
  const [,, texPath, outPath, tplPath, bibPath] = process.argv;
  const src = fs.readFileSync(texPath, 'utf8').replace(/\\tableofcontents/, '');
  const png = fs.readFileSync('/home/claude/ihr-latex/figura-exemplo.png');
  const res = L.convert(src, { bib: bibPath ? new B.Library(fs.readFileSync(bibPath, 'utf8')) : null, images: { 'figura-exemplo': { base64: png.toString('base64'), contentType: 'image/png', width: 520, height: 275 } } });
  if (res.warnings.length) console.log('avisos:', res.warnings);
  const zip = await JSZip.loadAsync(fs.readFileSync(tplPath));
  let doc = await zip.file('word/document.xml').async('string');
  // fim do parágrafo com a 1a quebra de seção
  const k = doc.indexOf('<w:sectPr', doc.indexOf('<w:body>'));
  const cut = doc.indexOf('</w:p>', k) + 6;
  const last = doc.lastIndexOf('<w:sectPr');
  let body = res.body.replace(/<w:numId w:val="(\d+)"\/>/g, (m, n) => n === '0' ? m : `<w:numId w:val="${900 + +n}"/>`);
  doc = doc.slice(0, cut) + body + doc.slice(last);
  // garante namespaces usados
  zip.file('word/document.xml', doc);
  // numeração
  let num = await zip.file('word/numbering.xml').async('string');
  const abs = L.numberingXml([]).replace(/^<w:numbering[^>]*>/, '').replace('</w:numbering>', '');
  const nums = res.nums.map(n => `<w:num w:numId="${900 + n.id}"><w:abstractNumId w:val="${n.abs}"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>`).join('');
  const firstNum = num.indexOf('<w:num ');
  num = num.slice(0, firstNum) + abs + num.slice(firstNum).replace('</w:numbering>', nums + '</w:numbering>');
  zip.file('word/numbering.xml', num);
  // notas
  if (res.footnotes.length) {
    let fn = await zip.file('word/footnotes.xml').async('string');
    fn = fn.replace('</w:footnotes>', res.footnotes.join('') + '</w:footnotes>');
    zip.file('word/footnotes.xml', fn);
  }
  // rels + mídia
  let rels = await zip.file('word/_rels/document.xml.rels').async('string');
  rels = rels.replace('</Relationships>', res.rels.map(r => `<Relationship Id="${r.id}" Type="${r.type}" Target="${r.target}"${r.external ? ' TargetMode="External"' : ''}/>`).join('') + '</Relationships>');
  zip.file('word/_rels/document.xml.rels', rels);
  res.media.forEach(m => zip.file('word/media/' + m.name, Buffer.from(m.base64, 'base64')));
  let ct = await zip.file('[Content_Types].xml').async('string');
  if (!/Extension="png"/.test(ct)) ct = ct.replace('<Default ', '<Default Extension="png" ContentType="image/png"/><Default ');
  zip.file('[Content_Types].xml', ct);
  fs.writeFileSync(outPath, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
  console.log('ok', outPath);
})();
