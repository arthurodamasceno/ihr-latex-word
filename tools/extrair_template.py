"""Gera lib/template-parts.js a partir de um template .dotx/.docx do IHR.
Uso: python tools/extrair_template.py "Template Relatório v3.dotx"
"""
import re, json, sys, zipfile

z = zipfile.ZipFile(sys.argv[1])
st = z.read('word/styles.xml').decode('utf8')
num = z.read('word/numbering.xml').decode('utf8')

def style(sid):
    return re.search(r'<w:style [^>]*w:styleId="%s".*?</w:style>' % re.escape(sid), st, re.S).group(0)

need = ['Normal', 'Ttulo1', 'Ttulo2', 'Ttulo3', 'Ttulo4', 'Ttulo5', 'Ttulo6', 'Fontepargpadro', 'Tabelanormal', 'Semlista',
        'Hyperlink', 'PargrafodaLista', 'Legenda', 'Tabelacomgrade', 'SimplesTabela3', 'Textodenotaderodap', 'Refdenotaderodap',
        'Fonte', 'Imagem', 'Equaes', 'Tabelas', 'Quadro', 'Sumrio1', 'Sumrio2', 'Sumrio3', 'Sumrio4', 'Sumrio5', 'Sumrio6', 'Referncias',
        'CapaTtulo', 'CapaSubttulo', 'CapaLocaleData', 'CabealhodoSumrio']
out = []
for sid in need:
    s = style(sid)
    s = re.sub(r' w:rsid\w*="[^"]*"', '', s)
    s = re.sub(r'<w:rsid [^>]*/>', '', s)
    s = re.sub(r'<w:link [^>]*/>', '', s)
    if sid.startswith('Ttulo'):
        s = re.sub(r'<w:numPr>.*?</w:numPr>', '', s, flags=re.S)
    out.append(s)
dd = re.search(r'<w:docDefaults>.*?</w:docDefaults>', st, re.S).group(0)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
styles = '<w:styles %s>%s%s</w:styles>' % (W, dd, ''.join(out))

def absnum(numid):
    m = re.search(r'<w:num w:numId="%s"[^>]*>.*?</w:num>' % numid, num, re.S).group(0)
    aid = re.search(r'abstractNumId w:val="(\d+)"', m).group(1)
    a = re.search(r'<w:abstractNum [^>]*w:abstractNumId="%s".*?</w:abstractNum>' % aid, num, re.S).group(0)
    for pat in [r' w15:restartNumberingAfterBreak="\d"', r' w:tplc="[^"]*"', r' w:tentative="1"', r'<w:nsid [^>]*/>', r'<w:tmpl [^>]*/>']:
        a = re.sub(pat, '', a)
    return a

dec = re.sub(r'w:abstractNumId="\d+"', 'w:abstractNumId="9001"', absnum('49'), count=1)
bul = re.sub(r'w:abstractNumId="\d+"', 'w:abstractNumId="9002"', absnum('50'), count=1)
js = """/* Gerado a partir do template IHR (tools/extrair_template.py) */
(function (root) {
  var T = {};
  T.styles = %s;
  T.absDecimal = %s;
  T.absBullet = %s;
  root.IHRLX = root.IHRLX || {};
  root.IHRLX.template = T;
  if (typeof module !== 'undefined') module.exports = T;
})(typeof globalThis !== 'undefined' ? globalThis : this);
""" % (json.dumps(styles, ensure_ascii=False), json.dumps(dec, ensure_ascii=False), json.dumps(bul, ensure_ascii=False))
open('lib/template-parts.js', 'w', encoding='utf8').write(js)
print('ok')
