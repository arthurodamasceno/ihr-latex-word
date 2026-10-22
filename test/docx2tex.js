// Lê um .docx inteiro como se fosse a seleção e converte para LaTeX
global.katex = require('../vendor/katex.min.js');
const fs = require('fs');
const JSZip = require('/opt/npm-tools/node_modules/jszip');
const X = require('../lib/xml.js');
const O = require('../lib/ooxml2latex.js');
(async () => {
  const zip = await JSZip.loadAsync(fs.readFileSync(process.argv[2]));
  const get = async n => zip.file(n) ? X.parse(await zip.file(n).async('string')) : null;
  const relsT = await get('word/_rels/document.xml.rels');
  const rels = {}; X.kids(relsT, 'Relationship').forEach(r => rels[X.attr(r, 'Id')] = { target: X.attr(r, 'Target'), type: X.attr(r, 'Type') });
  const media = {};
  for (const n of Object.keys(zip.files)) if (n.startsWith('word/media/') && !zip.files[n].dir) media[n.slice(5)] = { base64: await zip.file(n).async('base64'), ct: 'image/' + n.split('.').pop().replace('jpg','jpeg') };
  const P = { document: await get('word/document.xml'), styles: await get('word/styles.xml'), numbering: await get('word/numbering.xml'),
    footnotes: await get('word/footnotes.xml'), rels, part: t => media[t] || null };
  const r = O.convert(P);
  fs.writeFileSync(process.argv[3], r.latex);
  console.log('warnings:', r.warnings); console.log('unsafe:', r.unsafe); console.log('images:', Object.keys(r.images).map(k => k + ' ' + r.images[k].width + 'x' + r.images[k].height));
})();
