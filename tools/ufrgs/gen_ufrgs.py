"""Gera o template Word (.dotx e .docx) equivalente ao template LaTeX do TCC de Engenharia Mecânica da UFRGS.

Medidas tiradas do PDF compilado (pdfplumber) e dos parâmetros do abnTeX2/memoir:
  A4, margens 3/2/3/2 cm; Times 12; entrelinha 14,5 pt; recuo 17,5 pt; títulos em negrito com
  numeração "1." / "2.1." seguida de espaço; capítulo 14 pt em caixa alta; 12 pt antes/depois.
Os estilos usam os mesmos IDs do template IHR, para o suplemento LaTeX IHR funcionar sem mudanças.
O conteúdo dos capítulos vem de conteudo.json (gerado por conv.js com o conversor do suplemento).
"""
import json, zipfile, sys, struct, os
from xml.sax.saxutils import escape as E

HERE = os.path.dirname(os.path.abspath(__file__))
D = json.load(open(os.path.join(HERE, 'conteudo.json'), encoding='utf8'))
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
NS = ('xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
      'xmlns:o="urn:schemas-microsoft-com:office:office" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
      'xmlns:v="urn:schemas-microsoft-com:vml" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:w10="urn:schemas-microsoft-com:office:word" '
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
      'mc:Ignorable="w14"')
TW = 9070          # largura do texto (twips): 21 - 3 - 2 cm
LINE = 'w:line="252" w:lineRule="auto"'   # 14,5 pt para Times 12

# ------------------------------------------------------------------ estilos
def style(sid, name, ppr='', rpr='', based='Normal', nxt='Normal', typ='paragraph', custom=False, extra='<w:qFormat/>'):
    b = '<w:basedOn w:val="%s"/>' % based if based else ''
    n = '<w:next w:val="%s"/>' % nxt if nxt else ''
    return ('<w:style w:type="%s"%s w:styleId="%s"><w:name w:val="%s"/>%s%s%s%s%s</w:style>' %
            (typ, ' w:customStyle="1"' if custom else '', sid, name, b, n, extra,
             '<w:pPr>%s</w:pPr>' % ppr if ppr else '', '<w:rPr>%s</w:rPr>' % rpr if rpr else ''))

H = '<w:keepNext/><w:keepLines/>'
def head(lvl, sz, caps=False, num=True, extra_r=''):
    p = H + ('<w:numPr><w:ilvl w:val="%d"/><w:numId w:val="1"/></w:numPr>' % (lvl - 1) if num else '') + \
        '<w:spacing w:before="240" w:after="240" w:line="240" w:lineRule="auto"/><w:ind w:left="0" w:firstLine="0"/><w:jc w:val="left"/><w:outlineLvl w:val="%d"/>' % (lvl - 1)
    r = '<w:b/><w:bCs/>' + ('<w:caps/>' if caps else '') + extra_r + '<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (sz, sz)
    return p, r

styles = []
styles.append('<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/>'
              '<w:pPr><w:spacing w:after="0" %s/><w:ind w:firstLine="350"/><w:jc w:val="both"/></w:pPr>'
              '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>' % LINE)
for lvl, (nm, sz, caps, num, er) in enumerate([('heading 1', 28, True, True, ''), ('heading 2', 24, False, True, ''),
                                                ('heading 3', 24, False, True, ''), ('heading 4', 24, False, True, ''),
                                                ('heading 5', 24, False, False, ''), ('heading 6', 24, False, False, '<w:i/><w:iCs/>')], 1):
    p, r = head(lvl, sz, caps, num, er)
    styles.append(style('Ttulo%d' % lvl, nm, p, r, extra='<w:uiPriority w:val="9"/><w:qFormat/>'))
p, r = head(1, 28, True, False)
styles.append(style('TtuloSemNumerao', 'Título sem numeração', p, r, custom=True, extra='<w:uiPriority w:val="9"/><w:qFormat/>'))
styles.append(style('CabealhodoSumrio', 'TOC Heading',
                    '<w:pageBreakBefore/><w:spacing w:before="0" w:after="240" w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="left"/><w:outlineLvl w:val="9"/>',
                    '<w:b/><w:bCs/><w:caps/><w:sz w:val="28"/><w:szCs w:val="28"/>', extra='<w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:qFormat/>'))
TOCT = '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="%d"/></w:tabs>' % (TW - 4)
for i, (ind, r) in enumerate([(0, '<w:b/><w:bCs/><w:caps/>'), (350, '<w:caps/>'), (888, ''), (1426, '')], 1):
    styles.append(style('Sumrio%d' % i, 'toc %d' % i, TOCT + '<w:spacing w:after="50" %s/><w:ind w:left="%d" w:firstLine="0"/><w:jc w:val="left"/>' % (LINE, ind), r,
                        extra='<w:uiPriority w:val="39"/><w:unhideWhenUsed/>'))
C0 = '<w:ind w:firstLine="0"/><w:jc w:val="center"/>'
styles.append(style('Legenda', 'caption', '<w:keepNext/><w:spacing w:before="280" w:after="120" w:line="240" w:lineRule="auto"/>' + C0, '',
                    extra='<w:uiPriority w:val="35"/><w:unhideWhenUsed/><w:qFormat/>'))
styles.append(style('Fonte', 'Fonte', '<w:spacing w:before="60" w:after="240" w:line="240" w:lineRule="auto"/>' + C0, '<w:sz w:val="20"/><w:szCs w:val="20"/>', custom=True))
styles.append(style('Imagem', 'Imagem', '<w:spacing w:after="240" w:line="240" w:lineRule="auto"/>' + C0, '<w:noProof/>', custom=True))
styles.append(style('Equaes', 'Equações', '<w:spacing w:before="120" w:after="120" w:line="240" w:lineRule="auto"/>' + C0, '', custom=True))
styles.append(style('Tabelas', 'Tabelas', '<w:spacing w:line="240" w:lineRule="auto"/>' + C0, '', custom=True))
styles.append(style('Quadro', 'Quadro', '<w:spacing w:line="240" w:lineRule="auto"/>' + C0, '', custom=True))
styles.append(style('Bibliografia', 'Bibliography', '<w:spacing w:after="190" w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="both"/>', '',
                    nxt='Bibliografia', extra='<w:uiPriority w:val="37"/><w:unhideWhenUsed/><w:qFormat/>'))
styles.append(style('PargrafodaLista', 'List Paragraph', '<w:ind w:left="720"/><w:contextualSpacing/>', '', extra='<w:uiPriority w:val="34"/><w:qFormat/>'))
styles.append(style('Textodenotaderodap', 'footnote text', '<w:spacing w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0"/>', '<w:sz w:val="20"/><w:szCs w:val="20"/>', extra='<w:semiHidden/>'))
styles.append(style('Cabealho', 'header', '<w:tabs><w:tab w:val="center" w:pos="4535"/><w:tab w:val="right" w:pos="%d"/></w:tabs><w:spacing w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="right"/>' % TW,
                    '<w:sz w:val="20"/><w:szCs w:val="20"/>', extra='<w:uiPriority w:val="99"/><w:unhideWhenUsed/>'))
styles.append(style('Rodap', 'footer', '<w:spacing w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0"/>', '<w:sz w:val="20"/><w:szCs w:val="20"/>', extra='<w:uiPriority w:val="99"/><w:unhideWhenUsed/>'))
# capa e pré-textuais
styles.append(style('CapaInstituio', 'Capa - Instituição', C0, '', custom=True))
styles.append(style('CapaTtulo', 'Capa - Título', C0, '<w:caps/>', custom=True, extra='<w:uiPriority w:val="1"/><w:qFormat/>'))
styles.append(style('CapaAutor', 'Capa - Autor', C0, '', custom=True))
styles.append(style('CapaNatureza', 'Capa - Natureza do trabalho', '<w:ind w:left="4532" w:firstLine="0"/><w:jc w:val="both"/>', '', custom=True))
styles.append(style('CapaLocaleData', 'Capa - Local e Data',
                    '<w:framePr w:w="%d" w:hSpace="0" w:wrap="around" w:vAnchor="page" w:hAnchor="page" w:x="1701" w:y="15440"/>' % TW + C0, '', custom=True,
                    extra='<w:uiPriority w:val="1"/><w:qFormat/>'))
styles.append(style('NomenclaturaGrupo', 'Nomenclatura - Grupo', '<w:spacing w:before="0" w:after="190" %s/><w:ind w:left="58" w:firstLine="0"/><w:jc w:val="left"/>' % LINE, '', custom=True, nxt='Nomenclatura'))
styles.append(style('Nomenclatura', 'Nomenclatura', '<w:tabs><w:tab w:val="left" w:pos="684"/><w:tab w:val="right" w:pos="%d"/></w:tabs><w:spacing w:after="190" %s/><w:ind w:left="684" w:hanging="684"/><w:jc w:val="left"/>' % (TW - 4, LINE),
                    '', custom=True, nxt='Nomenclatura'))
BOX = '<w:pBdr><w:left w:val="single" w:sz="24" w:space="4" w:color="000000"/></w:pBdr>'
styles.append(style('Resumo', 'Resumo (caixa)', BOX + '<w:ind w:left="140" w:firstLine="0"/><w:jc w:val="both"/>', '', custom=True, nxt='Resumo'))
# caracteres e tabelas
styles.append('<w:style w:type="character" w:default="1" w:styleId="Fontepargpadro"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>')
styles.append('<w:style w:type="character" w:styleId="Refdenotaderodap"><w:name w:val="footnote reference"/><w:basedOn w:val="Fontepargpadro"/><w:semiHidden/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>')
styles.append('<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="Fontepargpadro"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>')
styles.append('<w:style w:type="table" w:default="1" w:styleId="Tabelanormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>'
              '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>')
B4 = lambda s: '<w:%s w:val="single" w:sz="4" w:space="0" w:color="auto"/>' % s
styles.append('<w:style w:type="table" w:styleId="Tabelacomgrade"><w:name w:val="Table Grid"/><w:basedOn w:val="Tabelanormal"/><w:uiPriority w:val="39"/>'
              '<w:pPr><w:spacing w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0"/></w:pPr><w:tblPr><w:tblBorders>%s</w:tblBorders></w:tblPr></w:style>' % ''.join(B4(s) for s in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']))
# tabela "table" (ABNT/abnTeX2): sem bordas externas laterais, centralizada
styles.append('<w:style w:type="table" w:styleId="SimplesTabela3"><w:name w:val="Plain Table 3"/><w:basedOn w:val="Tabelanormal"/><w:uiPriority w:val="43"/>'
              '<w:pPr><w:spacing w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="center"/></w:pPr>'
              '<w:tblPr><w:jc w:val="center"/><w:tblBorders>%s</w:tblBorders></w:tblPr></w:style>' % B4('insideV'))
styles.append('<w:style w:type="numbering" w:default="1" w:styleId="Semlista"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>')

STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles %s><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Times New Roman" w:eastAsia="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
          '<w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="pt-BR" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>%s</w:styles>') % (W, ''.join(styles))

# ------------------------------------------------------------------ numeração dos títulos
lv = ''
for i, txt in enumerate(['%1.', '%1.%2.', '%1.%2.%3.', '%1.%2.%3.%4.']):
    lv += ('<w:lvl w:ilvl="%d"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Ttulo%d"/><w:suff w:val="space"/>'
           '<w:lvlText w:val="%s"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr></w:lvl>') % (i, i + 1, txt)
for i in range(4, 9):
    lv += '<w:lvl w:ilvl="%d"><w:start w:val="1"/><w:numFmt w:val="none"/><w:suff w:val="nothing"/><w:lvlText w:val=""/><w:lvlJc w:val="left"/></w:lvl>' % i
num_extra = D.get('numbering', '')
abs_extra = ''.join(__import__('re').findall(r'<w:abstractNum .*?</w:abstractNum>', num_extra, __import__('re').S))
num_list = ''.join(__import__('re').findall(r'<w:num .*?</w:num>', num_extra, __import__('re').S))
NUMBERING = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering %s><w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="multilevel"/>%s</w:abstractNum>%s'
             '<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>%s</w:numbering>') % (W, lv, abs_extra, num_list)

# ------------------------------------------------------------------ helpers de corpo
def r(t, b=False, i=False, sz=None, caps=False, sup=False):
    pr = ('<w:b/><w:bCs/>' if b else '') + ('<w:i/><w:iCs/>' if i else '') + ('<w:caps/>' if caps else '') + \
         ('<w:vertAlign w:val="superscript"/>' if sup else '') + ('<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (sz, sz) if sz else '')
    return '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % ('<w:rPr>%s</w:rPr>' % pr if pr else '', E(t))
def p(inner='', st=None, ppr=''):
    s = '<w:pStyle w:val="%s"/>' % st if st else ''
    return '<w:p><w:pPr>%s%s</w:pPr>%s</w:p>' % (s, ppr, inner)
def blank(n=1, ppr='<w:ind w:firstLine="0"/>'):
    return ''.join(p('', None, ppr) for _ in range(n))
PB = '<w:r><w:br w:type="page"/></w:r>'
def fld(code, res, rpr=''):
    rp = '<w:rPr>%s</w:rPr>' % rpr if rpr else ''
    return ('<w:r>%s<w:fldChar w:fldCharType="begin"/></w:r><w:r>%s<w:instrText xml:space="preserve"> %s </w:instrText></w:r>'
            '<w:r>%s<w:fldChar w:fldCharType="separate"/></w:r><w:r>%s<w:t>%s</w:t></w:r><w:r>%s<w:fldChar w:fldCharType="end"/></w:r>') % (rp, rp, code, rp, rp, E(res), rp)
MAR = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1701" w:right="1134" w:bottom="1134" w:left="1701" w:header="1080" w:footer="709" w:gutter="0"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/>'
def sect(valign='top', hdr=False, last=False, start=None):
    hr = '<w:headerReference w:type="default" r:id="rIdH1"/>' if hdr else ''
    s = '<w:sectPr>%s<w:type w:val="nextPage"/>%s%s</w:sectPr>' % (hr, MAR.replace('<w:cols', '<w:cols'), '<w:vAlign w:val="%s"/>' % valign if valign != 'top' else '')
    # ordem: headerReference, type, pgSz, pgMar, cols, vAlign, docGrid
    s = s.replace('<w:docGrid w:linePitch="360"/>', '').replace('</w:sectPr>', '<w:docGrid w:linePitch="360"/></w:sectPr>')
    if start is not None: s = s.replace('<w:cols ', '<w:pgNumType w:start="%d"/><w:cols ' % start)
    return s if last else p('', None, '<w:ind w:firstLine="0"/>' + s)

# imagens
media = {}
def png_size(path):
    b = open(path, 'rb').read(32); return struct.unpack('>II', b[16:24])
def drawing(rid, cx_pt, cy_pt, pid, name):
    cx, cy = int(cx_pt * 12700), int(cy_pt * 12700)
    return ('<w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="%d" cy="%d"/>'
            '<wp:docPr id="%d" name="%s"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="%d" name="%s"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>') % (
        cx, cy, pid, name, pid, name, rid, cx, cy)

# ------------------------------------------------------------------ pré-textuais
NC = '<w:ind w:firstLine="0"/>'
CTR = '<w:ind w:firstLine="0"/><w:jc w:val="center"/>'
body = ''
# capa
for t in ['UNIVERSIDADE FEDERAL DO RIO GRANDE DO SUL', 'ESCOLA DE ENGENHARIA - CURSO DE ENGENHARIA MECÂNICA', 'TRABALHO DE CONCLUSÃO DE CURSO']:
    body += p(r(t), 'CapaInstituio')
body += blank(4, CTR) + p(r('Template TCC Engenharia Mecânica UFRGS'), 'CapaTtulo') + blank(1, CTR) + p(r('por'), 'CapaAutor') + blank(2, CTR)
body += p(r('Gabriel Rodrigo Hilgert'), 'CapaAutor') + blank(5, CTR)
body += p(r('Monografia apresentada ao Departamento de Engenharia Mecânica da Escola de Engenharia da Universidade Federal do Rio Grande do Sul, como parte dos requisitos para obtenção do diploma de Engenheiro Mecânico.'), 'CapaNatureza')
body += p(r('Porto Alegre, Março de 2023'), 'CapaLocaleData', '<w:framePr w:w="%d" w:hSpace="0" w:wrap="around" w:vAnchor="page" w:hAnchor="page" w:x="1701" w:y="15440"/>' % TW)
body += p('', None, NC)   # âncora do quadro na página 1
# ficha catalográfica (imagem do PDF da biblioteca)
media['ficha.png'] = open(os.path.join(HERE, 'ficha.png'), 'rb').read()
fw, fh = png_size(os.path.join(HERE, 'ficha.png'))
body += p('', None, '<w:pageBreakBefore/><w:spacing w:line="8400" w:lineRule="exact"/>' + NC)   # empurra a ficha para a metade de baixo
body += p(drawing('rIdFicha', fw * 72 / 200, fh * 72 / 200, 9001, 'Ficha catalográfica'), None, CTR)
# folha de aprovação
body += p(PB + r('Gabriel Rodrigo Hilgert'), None, CTR) + blank(5, CTR) + p(r('TEMPLATE TCC ENGENHARIA MECÂNICA UFRGS'), None, CTR) + blank(3, CTR)
BX = '<w:pBdr><w:top w:val="single" w:sz="8" w:space="4" w:color="000000"/><w:left w:val="single" w:sz="8" w:space="4" w:color="000000"/><w:bottom w:val="single" w:sz="8" w:space="4" w:color="000000"/><w:right w:val="single" w:sz="8" w:space="4" w:color="000000"/></w:pBdr><w:ind w:left="567" w:right="0" w:firstLine="0"/>'
body += p(r('ESTA MONOGRAFIA FOI JULGADA ADEQUADA COMO PARTE DOS REQUISITOS PARA A OBTENÇÃO DO TÍTULO DE'), None, BX + '<w:jc w:val="center"/>')
body += p(r('ENGENHEIRO MECÂNICO', b=True), None, BX + '<w:jc w:val="center"/>')
body += p(r('APROVADA EM SUA FORMA FINAL PELA BANCA EXAMINADORA DO CURSO DE ENGENHARIA MECÂNICA'), None, BX + '<w:jc w:val="center"/>')
body += p('', None, BX)
TABC = '<w:tabs><w:tab w:val="left" w:pos="3654"/></w:tabs>'
body += p('<w:r><w:tab/></w:r>' + r('Prof. Cirilo Seppi Bresolin'), None, BX + TABC + '<w:jc w:val="left"/>')
body += p('<w:r><w:tab/></w:r>' + r('Coordenador do Curso de Engenharia Mecânica'), None, BX + TABC + '<w:jc w:val="left"/>')
body += blank(1) + p(r('Área de Concentração: Mecânica dos Sólidos'), None, NC + '<w:jc w:val="left"/>') + blank(1)
body += p(r('Orientador: Nome do Orientador'), None, NC + '<w:jc w:val="left"/>') + blank(1)
body += p(r('Comissão de Avaliação:'), None, NC + '<w:jc w:val="left"/>') + blank(1)
for f in ['A', 'B', 'C']:
    body += p(r('Prof. Fulano ' + f), None, '<w:ind w:left="1701" w:firstLine="0"/><w:jc w:val="left"/>') + blank(1)
body += sect('top', start=0)
# agradecimentos (alinhado embaixo)
body += p(r('AGRADECIMENTOS'), None, NC + '<w:jc w:val="left"/>') + blank(1)
body += p(r(D['lipsum1']), None, '<w:ind w:firstLine="567"/>')
body += sect('bottom')
# epígrafe
body += p(r('EPÍGRAFE'), None, NC + '<w:jc w:val="left"/>') + blank(1)
body += p(r('"' + D['lipsum1'] + '"', i=True), None, NC + '<w:jc w:val="right"/>') + blank(2)
body += p(r('Nome do autor da frase', i=True), None, NC + '<w:jc w:val="right"/>')
body += sect('bottom')
# resumo
for t in ['UNIVERSIDADE FEDERAL DO RIO GRANDE DO SUL', 'ESCOLA DE ENGENHARIA - CURSO DE ENGENHARIA MECÂNICA', 'TRABALHO DE CONCLUSÃO DE CURSO – 2023']:
    body += p(r(t, sz=28), None, CTR + '<w:spacing w:line="240" w:lineRule="auto"/>')
body += blank(1) + p(r('TEMPLATE TCC ENGENHARIA MECÂNICA UFRGS', b=True, sz=28), None, NC + '<w:jc w:val="left"/>') + blank(2)
LIP = D['lipsum1']
body += p(r('Hilgert, Gabriel Rodrigo', b=True), 'Resumo') + p(r('gh_1770@hotmail.com'), 'Resumo') + p('', 'Resumo')
body += p(r('Resumo. ', b=True, i=True) + r('<' + LIP + '>'), 'Resumo') + p('', 'Resumo')
body += p(r('Palavras-Chave: ', b=True, i=True) + r('<Polímeros com Memória de Forma, Transição de Fase, Modelo Constitutivo, Elementos Finitos>'), 'Resumo') + p('', 'Resumo')
body += p(r('Abstract. ', b=True, i=True) + r('<' + LIP + '>', i=True), 'Resumo') + p('', 'Resumo')
body += p(r('Keywords: ', b=True, i=True) + r('<Shape-Memory Polymers, Phase Transition, Constitutive Behaviour, Finite Strain>'), 'Resumo')
# nomenclatura
body += p(PB, None, NC) if False else ''
body += p(r('Nomenclatura'), 'TtuloSemNumerao', '<w:pageBreakBefore/>') + D['nomencl']
# sumário (campo TOC com resultado em cache)
body += p(r('Sumário'), 'CabealhodoSumrio')
toc = [(1, '1. INTRODUÇÃO', 9), (1, '2. FUNDAMENTAÇÃO TEÓRICA', 9), (2, '2.1. Equações', 9), (3, '2.1.1. Tensor Cauchy-Green à direita', 9),
       (2, '2.2. Elementos Gráficos', 9), (3, '2.2.1. Tabelas', 9), (3, '2.2.2. Imagens', 10), (2, '2.3. Citações', 10), (1, '3. METODOLOGIA', 10),
       (2, '3.1. Seção', 10), (3, '3.1.1. Subseção', 10), (1, '4. RESULTADOS', 12), (1, '5. CONCLUSÃO', 13), (1, 'Referências', 13),
       (1, 'Apêndice', 14), (1, 'A. PSEUDOCÓDIGOS', 14)]
for k, (lvl, t, pg) in enumerate(toc):
    pre = ('<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-4" \\h \\z \\u </w:instrText></w:r>'
           '<w:r><w:fldChar w:fldCharType="separate"/></w:r>') if k == 0 else ''
    post = '<w:r><w:fldChar w:fldCharType="end"/></w:r>' if k == len(toc) - 1 else ''
    body += p(pre + r(t) + '<w:r><w:tab/></w:r>' + r(str(pg)) + post, 'Sumrio%d' % lvl)
body += sect('top')
# ------------------------------------------------------------------ textuais (conteúdo convertido pelo suplemento)
content = D['body']
first = content.find('<w:pStyle w:val="Ttulo1"/>')
body += content
rels_extra = D['rels']
for m in D['media']:
    import base64
    media[m['name']] = base64.b64decode(m['base64'])
body += sect('top', hdr=True, last=True)

DOC = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s</w:body></w:document>' % (NS, body)

# cabeçalho: número da página à direita (abnTeX2 conta a partir da folha de rosto: página física - 1)
num = fld('PAGE', '9')
HDR = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr %s><w:p><w:pPr><w:pStyle w:val="Cabealho"/></w:pPr>%s</w:p></w:hdr>' % (NS, num)

R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
rels = [('rIdS', R + 'styles', 'styles.xml'), ('rIdN', R + 'numbering', 'numbering.xml'), ('rIdSet', R + 'settings', 'settings.xml'),
        ('rIdH1', R + 'header', 'header1.xml'), ('rIdFicha', R + 'image', 'media/ficha.png')]
for x in rels_extra:
    rels.append((x['id'], x['type'], x['target']))
if D.get('footnotes'):
    rels.append(('rIdF', R + 'footnotes', 'footnotes.xml'))
RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">%s</Relationships>' % \
       ''.join('<Relationship Id="%s" Type="%s" Target="%s"/>' % (E(a), E(b), E(c)) for a, b, c in rels)
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s><w:defaultTabStop w:val="708"/>'
            '<w:hyphenationZone w:val="425"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>') % W

def write(path, main_ct):
    CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.'
    types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
             '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
             '<Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/>'
             '<Override PartName="/word/document.xml" ContentType="%s"/><Override PartName="/word/styles.xml" ContentType="%sstyles+xml"/>'
             '<Override PartName="/word/numbering.xml" ContentType="%snumbering+xml"/><Override PartName="/word/settings.xml" ContentType="%ssettings+xml"/>'
             '<Override PartName="/word/header1.xml" ContentType="%sheader+xml"/>%s'
             '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>') % (
        main_ct, CT, CT, CT, CT, ('<Override PartName="/word/footnotes.xml" ContentType="%sfootnotes+xml"/>' % CT) if D.get('footnotes') else '')
    z = zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED)
    z.writestr('[Content_Types].xml', types)
    z.writestr('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
               '<Relationship Id="rId1" Type="%sofficeDocument" Target="word/document.xml"/>'
               '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>' % R)
    z.writestr('docProps/core.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
               'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
               '<dc:title>Template TCC Engenharia Mecânica UFRGS</dc:title><dc:creator>Template TCC UFRGS (convertido do LaTeX)</dc:creator></cp:coreProperties>')
    z.writestr('word/document.xml', DOC)
    z.writestr('word/styles.xml', STYLES)
    z.writestr('word/numbering.xml', NUMBERING)
    z.writestr('word/settings.xml', SETTINGS)
    z.writestr('word/header1.xml', HDR)
    z.writestr('word/_rels/document.xml.rels', RELS)
    if D.get('footnotes'):
        z.writestr('word/footnotes.xml', D['footnotes'])
    for n, data in media.items():
        z.writestr('word/media/' + n, data)
    z.close()

out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, 'Template TCC Eng Mecanica UFRGS')
write(out + '.docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml')
write(out + '.dotx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml')
print('ok', out)
