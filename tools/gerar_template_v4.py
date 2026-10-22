"""Gera o Template Relatório v4 a partir do v3:
 - fundo da capa passa para o cabeçalho de primeira página da seção 1 (fica fixo, fora do texto)
 - título/subtítulo/local/data da capa viram parágrafos com estilos próprios, posicionados por quadro (framePr)
 - "SUMÁRIO" usa o estilo Título do Sumário (TOC Heading) com quebra de página antes
"""
import re, sys, zipfile, shutil
src, dst = sys.argv[1], sys.argv[2]
zin = zipfile.ZipFile(src)
files = {n: zin.read(n) for n in zin.namelist()}
doc = files['word/document.xml'].decode('utf8')
sty = files['word/styles.xml'].decode('utf8')

FR1 = '<w:framePr w:w="7012" w:hSpace="0" w:wrap="around" w:vAnchor="page" w:hAnchor="page" w:x="3830" w:y="3408"/>'
FR2 = '<w:framePr w:w="9070" w:hSpace="0" w:wrap="around" w:vAnchor="page" w:hAnchor="page" w:x="1986" w:y="14812"/>'
# ---------------------------------------------------------------- corpo
body_start = doc.index('<w:body>') + len('<w:body>')
# parágrafos de topo do corpo
def top_paras(s, start):
    out = []; i = start; depth = 0
    pat = re.compile(r'<(/?)w:(p|tbl|sdt)\b[^>]*?(/?)>')
    cur = None
    for m in pat.finditer(s, start):
        close, tag, selfc = m.group(1), m.group(2), m.group(3)
        if not close and not selfc:
            if depth == 0: cur = m.start()
            depth += 1
        elif close:
            depth -= 1
            if depth == 0: out.append((cur, m.end()))
        elif selfc and depth == 0:
            out.append((m.start(), m.end()))
        if s.startswith('<w:sectPr', m.end()) and depth == 0: break
    return out
paras = top_paras(doc, body_start)
texts = [re.sub(r'<[^>]+>', '', doc[a:b]) for a, b in paras]
i_sum = next(k for k, t in enumerate(texts) if t.strip().endswith('SUMÁRIO'))
cover_end = paras[i_sum][0]
drawing = re.search(r'<w:drawing>.*?</w:drawing>', doc[paras[0][0]:paras[0][1]], re.S).group(0)

FRS = {'CapaTtulo': None, 'CapaSubttulo': None, 'CapaLocaleData': None}
def P(style, text):
    fr = FR1 if style != 'CapaLocaleData' else FR2     # quadro também direto no parágrafo (Word Online/LibreOffice)
    return ('<w:p><w:pPr><w:pStyle w:val="%s"/>%s</w:pPr><w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (style, fr, text))
cover = (P('CapaTtulo', 'PROJETO Xxxxx: Xxxxxxxxxxxxxxxxxxxxx') +
         P('CapaSubttulo', 'Vertical de Xxxxxxx Xxxxxxxxxx') +
         P('CapaLocaleData', 'Farroupilha - RS') +
         P('CapaLocaleData', '2025') +
         '<w:p><w:pPr><w:ind w:firstLine="0"/></w:pPr></w:p>')         # âncora dos quadros na página 1
sumario = '<w:p><w:pPr><w:pStyle w:val="CabealhodoSumrio"/></w:pPr><w:r><w:t>SUMÁRIO</w:t></w:r></w:p>'
doc = doc[:body_start] + cover + sumario + doc[paras[i_sum][1]:]

# página de referências no fim (antes da última quebra de seção do corpo)
def R(parts):
    return '<w:p><w:pPr><w:pStyle w:val="Bibliografia"/></w:pPr>' + ''.join(
        '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % ('<w:rPr><w:b/><w:bCs/></w:rPr>' if b else '', t) for t, b in parts) + '</w:p>'
refs = ('<w:p><w:pPr><w:pStyle w:val="TtuloSemNumerao"/></w:pPr><w:r><w:t>Referências</w:t></w:r></w:p>' +
        R([('ASSOCIAÇÃO BRASILEIRA DE NORMAS TÉCNICAS. ', 0), ('NBR 6023', 1), (': informação e documentação: referências: elaboração. Rio de Janeiro: ABNT, 2018.', 0)]) +
        R([('ASSOCIAÇÃO BRASILEIRA DE NORMAS TÉCNICAS. ', 0), ('NBR 10520', 1), (': informação e documentação: citações em documentos: apresentação. Rio de Janeiro: ABNT, 2023.', 0)]))
last = doc.rfind('<w:sectPr')
doc = doc[:last] + refs + doc[last:]

# seção 1: cabeçalho/rodapé de primeira página
first_sect = re.search(r'<w:sectPr [^>]*>', doc[body_start:])
k = body_start + first_sect.end()
doc = doc[:k] + '<w:headerReference w:type="first" r:id="rIdCapaH"/><w:footerReference w:type="first" r:id="rIdCapaF"/>' + doc[k:]
end1 = doc.index('</w:sectPr>', k)
seg = doc[k:end1]
seg2 = seg.replace('<w:cols ', '<w:titlePg/><w:cols ', 1) if '<w:titlePg' not in seg else seg
# titlePg deve vir depois de pgNumType e antes de... (ordem: ... pgNumType, cols, formProt, vAlign, noEndnote, titlePg, textDirection, bidi, rtlGutter, docGrid)
seg2 = seg.replace('<w:docGrid', '<w:titlePg/><w:docGrid', 1)
doc = doc[:k] + seg2 + doc[end1:]
files['word/document.xml'] = doc.encode('utf8')

# ---------------------------------------------------------------- cabeçalho da capa
hdr_root = re.match(r'(<\?xml[^>]*\?>\s*<w:hdr [^>]*>)', files['word/header5.xml'].decode('utf8')).group(1)
dr = drawing.replace('r:embed="rId11"', 'r:embed="rIdBg"')
dr = re.sub(r'<wp:positionH relativeFrom="page">.*?</wp:positionH>', '<wp:positionH relativeFrom="page"><wp:align>center</wp:align></wp:positionH>', dr, flags=re.S)
dr = re.sub(r'<wp:positionV relativeFrom="paragraph">.*?</wp:positionV>', '<wp:positionV relativeFrom="page"><wp:posOffset>6223</wp:posOffset></wp:positionV>', dr, flags=re.S)
dr = re.sub(r'<wp:docPr id="\d+"', '<wp:docPr id="9101"', dr)
dr = dr.replace('locked="0"', 'locked="1"')
hdr = hdr_root + '<w:p><w:pPr><w:pStyle w:val="Cabealho"/></w:pPr><w:r><w:rPr><w:noProof/></w:rPr>' + dr + '</w:r></w:p></w:hdr>'
files['word/headerCapa.xml'] = hdr.encode('utf8')
files['word/_rels/headerCapa.xml.rels'] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rIdBg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>').encode('utf8')
ftr_root = hdr_root.replace('<w:hdr ', '<w:ftr ')
files['word/footerCapa.xml'] = (ftr_root + '<w:p><w:pPr><w:pStyle w:val="Rodap"/></w:pPr></w:p></w:ftr>').encode('utf8')

rels = files['word/_rels/document.xml.rels'].decode('utf8')
rels = re.sub(r'<Relationship Id="rId11" [^>]*/>', '', rels)     # imagem saiu do corpo
rels = rels.replace('</Relationships>',
    '<Relationship Id="rIdCapaH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="headerCapa.xml"/>'
    '<Relationship Id="rIdCapaF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footerCapa.xml"/></Relationships>')
files['word/_rels/document.xml.rels'] = rels.encode('utf8')
ct = files['[Content_Types].xml'].decode('utf8')
ct = ct.replace('</Types>',
    '<Override PartName="/word/headerCapa.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
    '<Override PartName="/word/footerCapa.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>')
files['[Content_Types].xml'] = ct.encode('utf8')

# ---------------------------------------------------------------- estilos
WHITE = '<w:b/><w:bCs/><w:color w:val="FFFFFF" w:themeColor="background1"/>'
new = (
 '<w:style w:type="paragraph" w:customStyle="1" w:styleId="CapaTtulo"><w:name w:val="Capa - Título"/><w:basedOn w:val="Normal"/><w:next w:val="CapaSubttulo"/><w:uiPriority w:val="1"/><w:qFormat/>'
 '<w:pPr>' + FR1 + '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="center"/></w:pPr>'
 '<w:rPr>' + WHITE + '<w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>'
 '<w:style w:type="paragraph" w:customStyle="1" w:styleId="CapaSubttulo"><w:name w:val="Capa - Subtítulo"/><w:basedOn w:val="CapaTtulo"/><w:next w:val="CapaLocaleData"/><w:uiPriority w:val="1"/><w:qFormat/>'
 '<w:pPr><w:spacing w:before="598"/></w:pPr><w:rPr><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>'
 '<w:style w:type="paragraph" w:customStyle="1" w:styleId="CapaLocaleData"><w:name w:val="Capa - Local e Data"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="1"/><w:qFormat/>'
 '<w:pPr>' + FR2 + '<w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="center"/></w:pPr>'
 '<w:rPr>' + WHITE + '<w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>')
toc_head = ('<w:style w:type="paragraph" w:styleId="CabealhodoSumrio"><w:name w:val="TOC Heading"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:qFormat/>'
 '<w:pPr><w:pageBreakBefore/><w:spacing w:line="480" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="center"/><w:outlineLvl w:val="9"/></w:pPr>'
 '<w:rPr><w:b/><w:bCs/><w:caps/></w:rPr></w:style>')
sty = re.sub(r'<w:style [^>]*w:styleId="CabealhodoSumrio".*?</w:style>', toc_head, sty, flags=re.S)
bib_style = ('<w:style w:type="paragraph" w:styleId="Bibliografia"><w:name w:val="Bibliography"/><w:basedOn w:val="Normal"/><w:next w:val="Bibliografia"/><w:uiPriority w:val="37"/><w:unhideWhenUsed/><w:qFormat/>'
 '<w:pPr><w:spacing w:after="240" w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="left"/></w:pPr></w:style>')
sty = re.sub(r'<w:style [^>]*w:styleId="Bibliografia".*?</w:style>', bib_style, sty, flags=re.S)
# títulos sem número (Referências, Resumo, Apêndice...): como o Título 1, centralizado, entra no sumário
unnum = ('<w:style w:type="paragraph" w:customStyle="1" w:styleId="TtuloSemNumerao"><w:name w:val="Título sem numeração"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/>'
 '<w:pPr><w:keepNext/><w:keepLines/><w:pageBreakBefore/><w:spacing w:line="480" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="center"/><w:outlineLvl w:val="0"/></w:pPr>'
 '<w:rPr><w:rFonts w:eastAsiaTheme="majorEastAsia" w:cstheme="majorBidi"/><w:b/><w:caps/><w:szCs w:val="32"/></w:rPr></w:style>')
sty = sty.replace('</w:styles>', new + unnum + '</w:styles>')
files['word/styles.xml'] = sty.encode('utf8')

zout = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED)
for n in zin.namelist():
    zout.writestr(zin.getinfo(n), files[n])
for n in ['word/headerCapa.xml', 'word/_rels/headerCapa.xml.rels', 'word/footerCapa.xml']:
    zout.writestr(n, files[n])
zout.close()
print('ok', dst)
