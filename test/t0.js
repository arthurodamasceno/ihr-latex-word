global.katex = require('../vendor/katex.min.js');
const L = require('../lib/latex2ooxml.js');
const s = process.argv[2];
const r = L.convert(s, {});
console.log(r.blocks.map(b=>b.type).join(','), r.warnings, r.body.slice(0,300));
