# -*- coding: utf-8 -*-
"""Planilhas .xlsx no padrão OCA, em XML puro (sem Excel, COM ou openpyxl).

O módulo só monta as partes do pacote; cada comando continua gravando o ZIP
do seu jeito (o Exportar Coordenadas usa .NET, o Acabamentos usa zipfile).

Integridade dos dados: os valores são gravados exatamente como o comando
os entrega. A regra de cada aba diz só COMO gravar, nunca o QUE gravar:
- numbers=True: float/int (não bool) viram número; numbers="float": só float
  (int continua texto, como em gravadores que já faziam assim). Formato
  num_format = o mesmo que o comando já usava; nada é arredondado aqui.
- none_text: texto gravado no lugar de None (ex.: "-"); None = célula vazia.
- skip_empty=True: célula sem texto fica vazia, como hoje.
- Todo o resto vira texto (preserva zeros à esquerda, IDs e códigos).
Fórmulas não são criadas. Os dados começam na linha 1 de cada aba; a aba
Resumo é sempre a última, e a identificação do projeto vai no cabeçalho e
rodapé de impressão.

Padrão visual: cabeçalho verde-água com texto branco em negrito, painel
congelado na linha 1, tabela estruturada com filtro e linhas alternadas
(estilo de tabela "OCA"), bordas finas, Segoe UI 10, códigos em Consolas,
impressão em A4 paisagem ajustada à largura com a linha 1 repetida.

Engine: IronPython 2.7 (pyRevit). Sintaxe neutra py2/py3.
"""

import os
import re

from oca_ds import tokens

try:
    unicode
except NameError:           # IronPython 3 / CPython
    unicode = str

NS_MAIN = u"http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = u"http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG = u"http://schemas.openxmlformats.org/package/2006/relationships"
NS_XDR = u"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_A = u"http://schemas.openxmlformats.org/drawingml/2006/main"
CT = u"application/vnd.openxmlformats-officedocument."
HEAD = u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "logo", "logo.png")
BRAND = u"OCA"
SUMMARY_SHEET = u"Resumo"
TABLE_STYLE = u"OCA"

_INVALID_XML = re.compile(u"[\x00-\x08\x0b\x0c\x0e-\x1f]")

# índices de cellXfs (styles_xml monta nesta ordem)
S_DEFAULT, S_HEADER, S_TEXT, S_MONO, S_WRAP, S_TITLE, S_SUBTITLE, S_LABEL, S_VALUE, S_VALUE_NUM = range(10)
S_NUMBER_FIRST = 10          # um estilo por formato numérico, a partir daqui


def _text(value):
    if isinstance(value, unicode):
        return value
    try:
        return unicode(value)
    except Exception:
        try:
            return value.decode("utf-8", "replace")
        except Exception:
            return u""


def xml_escape(value):
    text = _INVALID_XML.sub(u"", _text(value))
    return (text.replace(u"&", u"&amp;").replace(u"<", u"&lt;")
            .replace(u">", u"&gt;").replace(u'"', u"&quot;"))


def col_letter(idx):
    """0 -> A, 25 -> Z, 26 -> AA."""
    letters = u""
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        letters = unicode(chr(65 + rem)) + letters
    return letters


def sheet_name(name):
    """Nome de aba válido no Excel (sem []:*?/\\, até 31 caracteres)."""
    return re.sub(u"[\\[\\]:*?/\\\\]", u" ", _text(name))[:31] or u"Planilha"


def _is_number(value, rule=True):
    if rule == "float":
        return isinstance(value, float)
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _hf_text(value):
    """Texto livre dentro do cabeçalho/rodapé de impressão (& vira &&)."""
    return _text(value).replace(u"&", u"&&")


# ------------------------------------------------------------------
# estilos
# ------------------------------------------------------------------
def styles_xml(number_formats=()):
    """styles.xml do padrão. number_formats: códigos de formato ("0.000"...);
    o estilo do formato i é S_NUMBER_FIRST + i."""
    x = tokens.EXCEL
    font, mono, size = xml_escape(x["font"]), xml_escape(x["font_mono"]), x["size"]
    fonts = [
        u'<font><sz val="{0}"/><color rgb="{1}"/><name val="{2}"/><family val="2"/></font>'.format(size, x["ink"], font),
        u'<font><b/><sz val="{0}"/><color rgb="{1}"/><name val="{2}"/><family val="2"/></font>'.format(size, x["header_ink"], font),
        u'<font><sz val="{0}"/><color rgb="{1}"/><name val="{2}"/><family val="3"/></font>'.format(size, x["ink"], mono),
        u'<font><b/><sz val="14"/><color rgb="{0}"/><name val="{1}"/><family val="2"/></font>'.format(x["ink"], font),
        u'<font><sz val="{0}"/><color rgb="{1}"/><name val="{2}"/><family val="2"/></font>'.format(size, x["ink2"], font),
    ]
    fills = [
        u'<fill><patternFill patternType="none"/></fill>',
        u'<fill><patternFill patternType="gray125"/></fill>',
        u'<fill><patternFill patternType="solid"><fgColor rgb="{}"/><bgColor indexed="64"/></patternFill></fill>'.format(x["header_fill"]),
    ]
    line = u'<{0} style="thin"><color rgb="{1}"/></{0}>'
    borders = [
        u'<border><left/><right/><top/><bottom/><diagonal/></border>',
        u'<border><left/><right/><top/>' + line.format(u"bottom", x["line"]) + u'<diagonal/></border>',
        u'<border><left/><right/><top/>' + line.format(u"bottom", x["header_fill"]) + u'<diagonal/></border>',
    ]
    top = u'<alignment vertical="top"/>'
    xfs = [
        u'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
        u'<xf numFmtId="0" fontId="1" fillId="2" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">'
        u'<alignment vertical="center"/></xf>',
        u'<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">' + top + u'</xf>',
        u'<xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1">' + top + u'</xf>',
        u'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">'
        u'<alignment vertical="top" wrapText="1"/></xf>',
        u'<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
        u'<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
        u'<xf numFmtId="0" fontId="4" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/>',
        u'<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">'
        u'<alignment horizontal="left"/></xf>',
        u'<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">'
        u'<alignment horizontal="left"/></xf>',
    ]
    fmts = []
    for i, code in enumerate(number_formats):
        fmts.append(u'<numFmt numFmtId="{}" formatCode="{}"/>'.format(164 + i, xml_escape(code)))
        xfs.append(u'<xf numFmtId="{}" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" '
                   u'applyBorder="1" applyAlignment="1">'.format(164 + i) + top + u'</xf>')
    solid = u'<fill><patternFill patternType="solid"><bgColor rgb="{}"/></patternFill></fill>'
    dxfs = [
        # 0 tabela inteira: filete inferior; 1 cabeçalho; 2 faixa; 3 linha sinalizada
        u'<dxf><border><bottom style="thin"><color rgb="{}"/></bottom></border></dxf>'.format(x["line"]),
        u'<dxf><font><b/><color rgb="{}"/></font>'.format(x["header_ink"]) + solid.format(x["header_fill"]) + u'</dxf>',
        u'<dxf>' + solid.format(x["band"]) + u'</dxf>',
        u'<dxf><font><color rgb="{}"/></font>'.format(x["warn"]) + solid.format(x["warn_soft"]) + u'</dxf>',
    ]
    return (HEAD + u'<styleSheet xmlns="{}">'.format(NS_MAIN)
            + (u'<numFmts count="{}">{}</numFmts>'.format(len(fmts), u"".join(fmts)) if fmts else u"")
            + u'<fonts count="{}">{}</fonts>'.format(len(fonts), u"".join(fonts))
            + u'<fills count="{}">{}</fills>'.format(len(fills), u"".join(fills))
            + u'<borders count="{}">{}</borders>'.format(len(borders), u"".join(borders))
            + u'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            + u'<cellXfs count="{}">{}</cellXfs>'.format(len(xfs), u"".join(xfs))
            + u'<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
            + u'<dxfs count="{}">{}</dxfs>'.format(len(dxfs), u"".join(dxfs))
            + u'<tableStyles count="1" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16">'
              u'<tableStyle name="{}" pivot="0" count="3">'.format(TABLE_STYLE)
            + u'<tableStyleElement type="wholeTable" dxfId="0"/>'
              u'<tableStyleElement type="headerRow" dxfId="1"/>'
              u'<tableStyleElement type="firstRowStripe" dxfId="2"/>'
              u'</tableStyle></tableStyles></styleSheet>')


# ------------------------------------------------------------------
# abas
# ------------------------------------------------------------------
def _table_name(name, used):
    base = re.sub(u"[^A-Za-z0-9_]", u"", _text(name)) or u"Dados"
    out, n = u"tbl" + base, 2
    while out.lower() in used:
        out, n = u"tbl{}{}".format(base, n), n + 1
    used.add(out.lower())
    return out


def _table_ok(header, nrows):
    """Tabela estruturada só quando o Excel aceita sem reparo: 1 linha de dados
    ou mais, títulos de texto não vazios, sem repetição e sem quebra de linha."""
    if nrows < 1:
        return False
    names = [_text(h) for h in header]
    if any((not n.strip()) or (u"\n" in n) or (u"\r" in n) for n in names):
        return False
    lower = [n.lower() for n in names]
    return len(set(lower)) == len(lower)


def _print_setup(title, project, landscape=True):
    head = u'&L&B{} · {}&B&R{}'.format(BRAND, _hf_text(title), _hf_text(project))
    foot = u'&L&A&RPágina &P de &N'
    return (u'<pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
            u'<pageSetup paperSize="9" orientation="{}" fitToWidth="1" fitToHeight="0"/>'
            u'<headerFooter><oddHeader>{}</oddHeader><oddFooter>{}</oddFooter></headerFooter>'.format(
                u"landscape" if landscape else u"portrait", xml_escape(head), xml_escape(foot)))


def _data_sheet(sheet, idx, sst, numfmt_index, table_names, table_id, title, project):
    """-> (xml da aba, xml da tabela ou None, nome da tabela, intervalo dos dados)."""
    rows = sheet["rows"]
    header, data = rows[0], rows[1:]
    ncol = max([1] + [len(r) for r in rows])
    widths = sheet.get("widths") or [12] * ncol
    numbers = sheet.get("numbers", False)
    none_text = sheet.get("none_text")
    skip_empty = sheet.get("skip_empty", False)
    mono = set(sheet.get("mono_cols", ()))
    wrap = set(sheet.get("wrap_cols", ()))
    num_style = S_NUMBER_FIRST + numfmt_index[sheet.get("num_format") or u"General"] if numbers else S_TEXT
    last_col = col_letter(ncol - 1)
    ref = u"A1:{}{}".format(last_col, len(rows))
    use_table = sheet.get("table", True) and _table_ok(header, len(data))
    filtered = sheet.get("filter", True) and len(data) > 0

    out = [HEAD, u'<worksheet xmlns="{}" xmlns:r="{}">'.format(NS_MAIN, NS_REL),
           u'<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>',
           u'<dimension ref="{}"/>'.format(ref),
           u'<sheetViews><sheetView {}workbookViewId="0"><pane ySplit="1" topLeftCell="A2" '
           u'activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/>'
           u'</sheetView></sheetViews>'.format(u'tabSelected="1" ' if idx == 0 else u""),
           u'<sheetFormatPr defaultRowHeight="15"/><cols>']
    for i, w in enumerate(widths):
        out.append(u'<col min="{0}" max="{0}" width="{1}" customWidth="1"/>'.format(i + 1, w))
    out.append(u'</cols><sheetData>')
    # cabeçalho: texto compartilhado (exigência da tabela estruturada)
    out.append(u'<row r="1" ht="20" customHeight="1">')
    for c, value in enumerate(header):
        out.append(u'<c r="{}1" s="{}" t="s"><v>{}</v></c>'.format(col_letter(c), S_HEADER, sst.index(value)))
    out.append(u'</row>')
    for r, row in enumerate(data):
        rn = r + 2
        out.append(u'<row r="{}">'.format(rn))
        for c, value in enumerate(row):
            cref = col_letter(c) + unicode(rn)
            if numbers and _is_number(value, numbers):
                out.append(u'<c r="{}" s="{}"><v>{}</v></c>'.format(cref, num_style, repr(value)))
                continue
            if value is None:
                if none_text is None:
                    continue
                value = none_text
            text = _text(value)
            if skip_empty and text == u"":
                continue
            style = S_WRAP if c in wrap else (S_MONO if c in mono else S_TEXT)
            out.append(u'<c r="{}" s="{}" t="inlineStr"><is><t xml:space="preserve">{}</t></is></c>'.format(
                cref, style, xml_escape(text)))
        out.append(u'</row>')
    out.append(u'</sheetData>')
    if filtered and not use_table:
        out.append(u'<autoFilter ref="{}"/>'.format(ref))
    flag = sheet.get("flag_col")
    if flag is not None and data:
        col = col_letter(flag)
        ignore = xml_escape(_text(sheet.get("flag_ignore", u"")).replace(u'"', u'""'))
        out.append(u'<conditionalFormatting sqref="A2:{0}{1}"><cfRule type="expression" dxfId="3" priority="1">'
                   u'<formula>AND(LEN(${2}2)&gt;0,${2}2&lt;&gt;"{3}")</formula></cfRule></conditionalFormatting>'.format(
                       last_col, len(rows), col, ignore))
    out.append(_print_setup(title, project, sheet.get("landscape", True)))
    table_xml, tname = None, None
    if use_table:
        tname = _table_name(sheet["name"], table_names)
        out.append(u'<tableParts count="1"><tablePart r:id="rId1"/></tableParts>')
        table_xml = (HEAD + u'<table xmlns="{}" id="{}" name="{}" displayName="{}" ref="{}" totalsRowShown="0">'
                     u'<autoFilter ref="{}"/><tableColumns count="{}">'.format(
                         NS_MAIN, table_id, tname, tname, ref, ref, ncol)
                     + u"".join(u'<tableColumn id="{}" name="{}"/>'.format(i + 1, xml_escape(h))
                                for i, h in enumerate(header))
                     + u'</tableColumns><tableStyleInfo name="{}" showFirstColumn="0" showLastColumn="0" '
                       u'showRowStripes="1" showColumnStripes="0"/></table>'.format(TABLE_STYLE))
    out.append(u'</worksheet>')
    return u"".join(out), table_xml, filtered and not use_table, ref


def _logo_size():
    """Logo com 34 px de altura, proporção do arquivo (PNG: largura/altura no cabeçalho IHDR)."""
    w, h = 294, 274
    try:
        with open(LOGO_PATH, "rb") as f:
            head = bytearray(f.read(24))
        w = (head[16] << 24) | (head[17] << 16) | (head[18] << 8) | head[19]
        h = (head[20] << 24) | (head[21] << 16) | (head[22] << 8) | head[23]
    except Exception:
        pass
    px_h = 34
    return int(round(px_h * float(w) / max(1, h))) * 9525, px_h * 9525


def logo_bytes():
    """Logo como bytearray: o único valor binário das partes. Os gravadores
    distinguem por isinstance(valor, bytearray) (no IronPython 2.7, str e
    unicode são o mesmo tipo, então bytes "puros" não dariam para separar)."""
    try:
        with open(LOGO_PATH, "rb") as f:
            return bytearray(f.read())
    except (IOError, OSError):
        return None


def _summary_sheet(title, subtitle, items, project, has_logo):
    out = [HEAD, u'<worksheet xmlns="{}" xmlns:r="{}">'.format(NS_MAIN, NS_REL),
           u'<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>',
           u'<sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>',
           u'<sheetFormatPr defaultRowHeight="15"/>',
           u'<cols><col min="1" max="1" width="7" customWidth="1"/><col min="2" max="2" width="30" customWidth="1"/>'
           u'<col min="3" max="3" width="60" customWidth="1"/></cols><sheetData>',
           u'<row r="1" ht="24" customHeight="1"><c r="B1" s="{}" t="inlineStr"><is><t xml:space="preserve">{}</t></is></c></row>'.format(
               S_TITLE, xml_escape(title)),
           u'<row r="2"><c r="B2" s="{}" t="inlineStr"><is><t xml:space="preserve">{}</t></is></c></row>'.format(
               S_SUBTITLE, xml_escape(subtitle))]
    for i, (label, value) in enumerate(items):
        rn = i + 4
        if _is_number(value):
            cell = u'<c r="C{}" s="{}"><v>{}</v></c>'.format(rn, S_VALUE_NUM, repr(value))
        else:
            cell = u'<c r="C{}" s="{}" t="inlineStr"><is><t xml:space="preserve">{}</t></is></c>'.format(
                rn, S_VALUE, xml_escape(u"" if value is None else value))
        out.append(u'<row r="{0}"><c r="B{0}" s="{1}" t="inlineStr"><is><t xml:space="preserve">{2}</t></is></c>{3}</row>'.format(
            rn, S_LABEL, xml_escape(label), cell))
    out.append(u'</sheetData>')
    out.append(_print_setup(title, project, landscape=False))
    if has_logo:
        out.append(u'<drawing r:id="rId1"/>')
    out.append(u'</worksheet>')
    return u"".join(out)


def _drawing_xml():
    cx, cy = _logo_size()
    return (HEAD + u'<xdr:wsDr xmlns:xdr="{}" xmlns:a="{}" xmlns:r="{}">'.format(NS_XDR, NS_A, NS_REL)
            + u'<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>57150</xdr:colOff>'
              u'<xdr:row>0</xdr:row><xdr:rowOff>38100</xdr:rowOff></xdr:from>'
            + u'<xdr:ext cx="{}" cy="{}"/>'.format(cx, cy)
            + u'<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Logo {0}" descr="Logo {0}"/>'
              u'<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>'.format(BRAND)
            + u'<xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
            + u'<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{}" cy="{}"/></a:xfrm>'.format(cx, cy)
            + u'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>'
              u'<xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>')


class _SharedStrings(object):
    def __init__(self):
        self.items, self.pos = [], {}

    def index(self, value):
        text = _text(value)
        if text not in self.pos:
            self.pos[text] = len(self.items)
            self.items.append(text)
        return self.pos[text]

    def xml(self):
        return (HEAD + u'<sst xmlns="{}" count="{}" uniqueCount="{}">'.format(NS_MAIN, len(self.items), len(self.items))
                + u"".join(u'<si><t xml:space="preserve">{}</t></si>'.format(xml_escape(t)) for t in self.items)
                + u'</sst>')


def workbook_parts(sheets, title, project=u"", summary=None, subtitle=u""):
    """Partes do pacote .xlsx -> {caminho: texto unicode ou bytearray (logo)}.

    sheets   [{"name", "rows" (1ª linha = cabeçalho), "widths", "numbers",
              "num_format", "none_text", "skip_empty", "mono_cols",
              "wrap_cols", "table", "filter", "flag_col", "flag_ignore",
              "landscape"}] na ordem das abas.
    title    nome da exportação ("Exportação de Coordenadas"): título do
             Resumo e cabeçalho de impressão "OCA · <title>".
    project  identificação do projeto (nome do arquivo do Revit) no
             cabeçalho de impressão e no Resumo.
    summary  [(rótulo, valor)] -> aba Resumo como ÚLTIMA aba; None = sem Resumo.
    """
    sst = _SharedStrings()
    formats = []
    for s in sheets:
        code = s.get("num_format") or u"General"
        if s.get("numbers") and code not in formats:
            formats.append(code)
    numfmt_index = dict((f, i) for i, f in enumerate(formats))

    parts, ct, wb_rels, wb_sheets, names = {}, [], [], [], []
    table_names = set()
    table_no = 0
    for i, s in enumerate(sheets):
        xml, table, autofilter, ref = _data_sheet(s, i, sst, numfmt_index, table_names, table_no + 1, title, project)
        n = i + 1
        sname = sheet_name(s["name"])
        parts[u"xl/worksheets/sheet{}.xml".format(n)] = xml
        ct.append(u'<Override PartName="/xl/worksheets/sheet{}.xml" ContentType="{}spreadsheetml.worksheet+xml"/>'.format(n, CT))
        wb_rels.append(u'<Relationship Id="rId{0}" Type="{1}/worksheet" Target="worksheets/sheet{0}.xml"/>'.format(n, NS_REL))
        wb_sheets.append(u'<sheet name="{}" sheetId="{}" r:id="rId{}"/>'.format(xml_escape(sname), n, n))
        quoted = u"'" + xml_escape(sname).replace(u"'", u"''") + u"'"
        names.append(u'<definedName name="_xlnm.Print_Titles" localSheetId="{}">{}!$1:$1</definedName>'.format(i, quoted))
        if autofilter:
            a, b = ref.split(u":")
            col_b = re.match(u"[A-Z]+", b).group(0)
            names.append(u'<definedName name="_xlnm._FilterDatabase" localSheetId="{}" hidden="1">{}!$A$1:${}${}</definedName>'.format(
                i, quoted, col_b, b[len(col_b):]))
        if table:
            table_no += 1
            parts[u"xl/tables/table{}.xml".format(table_no)] = table
            parts[u"xl/worksheets/_rels/sheet{}.xml.rels".format(n)] = (
                HEAD + u'<Relationships xmlns="{}"><Relationship Id="rId1" Type="{}/table" '
                u'Target="../tables/table{}.xml"/></Relationships>'.format(NS_PKG, NS_REL, table_no))
            ct.append(u'<Override PartName="/xl/tables/table{}.xml" ContentType="{}spreadsheetml.table+xml"/>'.format(table_no, CT))

    logo = logo_bytes() if summary is not None else None
    if summary is not None:
        n = len(sheets) + 1
        parts[u"xl/worksheets/sheet{}.xml".format(n)] = _summary_sheet(
            title, subtitle or u"Plugin {} · pyRevit".format(BRAND), summary, project, logo is not None)
        ct.append(u'<Override PartName="/xl/worksheets/sheet{}.xml" ContentType="{}spreadsheetml.worksheet+xml"/>'.format(n, CT))
        wb_rels.append(u'<Relationship Id="rId{0}" Type="{1}/worksheet" Target="worksheets/sheet{0}.xml"/>'.format(n, NS_REL))
        wb_sheets.append(u'<sheet name="{}" sheetId="{}" r:id="rId{}"/>'.format(SUMMARY_SHEET, n, n))
        if logo is not None:
            parts[u"xl/worksheets/_rels/sheet{}.xml.rels".format(n)] = (
                HEAD + u'<Relationships xmlns="{}"><Relationship Id="rId1" Type="{}/drawing" '
                u'Target="../drawings/drawing1.xml"/></Relationships>'.format(NS_PKG, NS_REL))
            parts[u"xl/drawings/drawing1.xml"] = _drawing_xml()
            parts[u"xl/drawings/_rels/drawing1.xml.rels"] = (
                HEAD + u'<Relationships xmlns="{}"><Relationship Id="rId1" Type="{}/image" '
                u'Target="../media/image1.png"/></Relationships>'.format(NS_PKG, NS_REL))
            parts[u"xl/media/image1.png"] = logo
            ct.append(u'<Override PartName="/xl/drawings/drawing1.xml" ContentType="{}drawing+xml"/>'.format(CT))

    count = len(wb_sheets)
    parts[u"xl/sharedStrings.xml"] = sst.xml()
    parts[u"xl/styles.xml"] = styles_xml(formats)
    parts[u"xl/workbook.xml"] = (
        HEAD + u'<workbook xmlns="{}" xmlns:r="{}"><bookViews><workbookView activeTab="0"/></bookViews><sheets>'.format(NS_MAIN, NS_REL)
        + u"".join(wb_sheets) + u'</sheets>'
        + (u'<definedNames>{}</definedNames>'.format(u"".join(names)) if names else u"")
        + u'</workbook>')
    parts[u"xl/_rels/workbook.xml.rels"] = (
        HEAD + u'<Relationships xmlns="{}">'.format(NS_PKG) + u"".join(wb_rels)
        + u'<Relationship Id="rId{}" Type="{}/styles" Target="styles.xml"/>'.format(count + 1, NS_REL)
        + u'<Relationship Id="rId{}" Type="{}/sharedStrings" Target="sharedStrings.xml"/>'.format(count + 2, NS_REL)
        + u'</Relationships>')
    parts[u"_rels/.rels"] = (
        HEAD + u'<Relationships xmlns="{}"><Relationship Id="rId1" Type="{}/officeDocument" '
        u'Target="xl/workbook.xml"/></Relationships>'.format(NS_PKG, NS_REL))
    parts[u"[Content_Types].xml"] = (
        HEAD + u'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        u'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        u'<Default Extension="xml" ContentType="application/xml"/>'
        + (u'<Default Extension="png" ContentType="image/png"/>' if logo is not None else u"")
        + u'<Override PartName="/xl/workbook.xml" ContentType="{}spreadsheetml.sheet.main+xml"/>'.format(CT)
        + u'<Override PartName="/xl/styles.xml" ContentType="{}spreadsheetml.styles+xml"/>'.format(CT)
        + u'<Override PartName="/xl/sharedStrings.xml" ContentType="{}spreadsheetml.sharedStrings+xml"/>'.format(CT)
        + u"".join(ct) + u'</Types>')
    return parts


# ------------------------------------------------------------------
# planilhas geradas no navegador (relatórios HTML)
# ------------------------------------------------------------------
def js_payload():
    """Objeto JSON com o styles.xml e os índices de estilo, injetado nos
    relatórios HTML que geram .xlsx em JavaScript (ver report.brand_report)."""
    import json
    return json.dumps({
        "styles": styles_xml(),
        "s": {"header": S_HEADER, "text": S_TEXT, "mono": S_MONO, "wrap": S_WRAP},
        "font": tokens.EXCEL["font"],
        "brand": BRAND,
        "print": u'<pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
                 u'<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>',
        "fitToPage": u'<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>',
    }, ensure_ascii=True)
