"""
Tablas dinamicas y segmentadores (slicers) nativos para archivos .xlsx.

openpyxl no sabe generar ninguna de las dos cosas, asi que las partes OOXML se
escriben a mano y el .xlsx se reempaqueta.

El flujo de uso es en dos tiempos, porque LibreOffice (que se usa para
recalcular y cachear los valores de las formulas) entiende y conserva las
tablas dinamicas pero descarta los segmentadores:

    1. openpyxl arma el libro
    2. inject_pivots(...)      -> tablas dinamicas
    3. recalc.py               -> LibreOffice calcula GETPIVOTDATA y cachea todo
    4. finalize(...)           -> segmentadores + retoques que LibreOffice pisa
"""
import re
import shutil
import zipfile
from xml.sax.saxutils import escape

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
NS_SLE = "http://schemas.microsoft.com/office/drawing/2010/slicer"
NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"

REL_PIVOT_CACHE = f"{NS_R}/pivotCacheDefinition"
REL_PIVOT_RECORDS = f"{NS_R}/pivotCacheRecords"
REL_PIVOT_TABLE = f"{NS_R}/pivotTable"
REL_DRAWING = f"{NS_R}/drawing"
REL_SLICER_CACHE = "http://schemas.microsoft.com/office/2007/relationships/slicerCache"
REL_SLICER = "http://schemas.microsoft.com/office/2007/relationships/slicer"

CT = {
    "cache": ("application/vnd.openxmlformats-officedocument"
              ".spreadsheetml.pivotCacheDefinition+xml"),
    "records": ("application/vnd.openxmlformats-officedocument"
                ".spreadsheetml.pivotCacheRecords+xml"),
    "pivot": ("application/vnd.openxmlformats-officedocument"
              ".spreadsheetml.pivotTable+xml"),
    "drawing": "application/vnd.openxmlformats-officedocument.drawing+xml",
    "slicer": "application/vnd.ms-excel.slicer+xml",
    "slicerCache": "application/vnd.ms-excel.slicerCache+xml",
}

DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'


def col_letter(idx):
    s = ""
    while idx:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s


class Pivot:
    """Tabla dinamica de una sola dimension en filas y una sola medida."""

    def __init__(self, name, sheet, anchor_row, anchor_col, row_field,
                 data_field, caption):
        self.name = name
        self.sheet = sheet
        self.anchor_row = anchor_row
        self.anchor_col = anchor_col
        self.row_field = row_field
        self.data_field = data_field
        self.caption = caption

    @property
    def top_left(self):
        return f"{col_letter(self.anchor_col)}{self.anchor_row}"


class Slicer:
    def __init__(self, field, caption, from_col, from_row, to_col, to_row,
                 column_count=1, style="SlicerStyleLight2"):
        self.field = field
        self.caption = caption
        self.from_col, self.from_row = from_col, from_row
        self.to_col, self.to_row = to_col, to_row
        self.column_count = column_count
        self.style = style
        self.cache_name = "Segm_" + re.sub(r"\W", "_", field)


class _Pkg:
    """Lectura/escritura del zip del .xlsx como un diccionario de partes."""

    def __init__(self, path):
        self.path = path
        with zipfile.ZipFile(path) as z:
            self.parts = {n: z.read(n) for n in z.namelist()}

    def txt(self, name):
        return self.parts[name].decode("utf-8")

    def put(self, name, text):
        self.parts[name] = text.encode("utf-8")

    def save(self):
        tmp = self.path + ".tmp"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            for name, data in self.parts.items():
                z.writestr(name, data)
        shutil.move(tmp, self.path)

    # -- helpers ------------------------------------------------------------
    def sheets(self):
        """nombre -> (indice 1-based en el orden del libro, sheetId)"""
        out = {}
        for i, m in enumerate(re.finditer(r"<sheet\b[^>]*/>",
                                          self.txt("xl/workbook.xml")), start=1):
            tag = m.group(0)
            name = re.search(r'name="([^"]*)"', tag).group(1)
            name = name.replace("&amp;", "&").replace("&quot;", '"')
            sid = int(re.search(r'sheetId="(\d+)"', tag).group(1))
            out[name] = (i, sid)
        return out

    def sheet_part(self, name):
        return f"xl/worksheets/sheet{self.sheets()[name][0]}.xml"

    def add_rels(self, part, new_rels):
        rp = re.sub(r"(.*/)([^/]+)$", r"\1_rels/\2.rels", part)
        if rp in self.parts:
            xml = self.txt(rp)
        else:
            xml = DECL + f'<Relationships xmlns="{NS_PKG}"></Relationships>'
        used = [int(x) for x in re.findall(r'Id="rId(\d+)"', xml)]
        nxt = max(used) + 1 if used else 1
        ids, add = [], ""
        for rtype, target in new_rels:
            rid = f"rId{nxt}"
            ids.append(rid)
            add += f'<Relationship Id="{rid}" Type="{rtype}" Target="{target}"/>'
            nxt += 1
        self.put(rp, xml.replace("</Relationships>", add + "</Relationships>"))
        return ids

    def add_content_types(self, overrides):
        ct = self.txt("[Content_Types].xml")
        add = "".join(f'<Override PartName="{p}" ContentType="{c}"/>'
                      for p, c in overrides if f'PartName="{p}"' not in ct)
        self.put("[Content_Types].xml", ct.replace("</Types>", add + "</Types>"))


# ---------------------------------------------------------------------------
#  PASO 1 - tablas dinamicas
# ---------------------------------------------------------------------------
class _CacheField:
    def __init__(self, name, values, is_measure):
        self.name = name
        self.raw = values
        self.is_measure = is_measure
        seen, items = set(), []
        for v in values:
            if v not in seen:
                seen.add(v)
                items.append(v)
        self.items = items
        self.index = {v: i for i, v in enumerate(items)}
        self.numeric = all(isinstance(v, (int, float)) for v in values)

    def xml(self):
        nm = escape(str(self.name))
        ia = ' containsInteger="1"'
        if self.is_measure:
            vals = [float(v) for v in self.raw]
            attr = ia if all(v.is_integer() for v in vals) else ""
            return (f'<cacheField name="{nm}" numFmtId="0"><sharedItems'
                    f' containsSemiMixedTypes="0" containsString="0"'
                    f' containsNumber="1"{attr} minValue="{min(vals):g}"'
                    f' maxValue="{max(vals):g}"/></cacheField>')
        if self.numeric:
            vals = [float(v) for v in self.items]
            attr = ia if all(v.is_integer() for v in vals) else ""
            body = "".join(f'<n v="{v:g}"/>' for v in vals)
            return (f'<cacheField name="{nm}" numFmtId="0"><sharedItems'
                    f' containsSemiMixedTypes="0" containsString="0"'
                    f' containsNumber="1"{attr} minValue="{min(vals):g}"'
                    f' maxValue="{max(vals):g}" count="{len(vals)}">{body}'
                    f"</sharedItems></cacheField>")
        body = "".join(f'<s v="{escape(str(v))}"/>' for v in self.items)
        return (f'<cacheField name="{nm}" numFmtId="0">'
                f'<sharedItems count="{len(self.items)}">{body}</sharedItems>'
                f"</cacheField>")

    def record(self, row):
        v = self.raw[row]
        return (f'<n v="{float(v):g}"/>' if self.is_measure
                else f'<x v="{self.index[v]}"/>')


def inject_pivots(path, source_sheet, columns, rows, pivots, slicer_fields=()):
    pkg = _Pkg(path)
    measures = {p.data_field for p in pivots}
    fields = [_CacheField(c, [r[i] for r in rows], c in measures)
              for i, c in enumerate(columns)]
    fidx = {f.name: i for i, f in enumerate(fields)}

    ref = f"A1:{col_letter(len(columns))}{len(rows) + 1}"
    pkg.put("xl/pivotCache/pivotCacheDefinition1.xml", DECL
            + f'<pivotCacheDefinition xmlns="{NS_MAIN}" xmlns:r="{NS_R}"'
            f' r:id="rId1" refreshOnLoad="1" refreshedVersion="6"'
            f' createdVersion="6" minRefreshableVersion="3"'
            f' recordCount="{len(rows)}"><cacheSource type="worksheet">'
            f'<worksheetSource ref="{ref}" sheet="{escape(source_sheet)}"/>'
            f'</cacheSource><cacheFields count="{len(fields)}">'
            + "".join(f.xml() for f in fields)
            + "</cacheFields></pivotCacheDefinition>")
    pkg.put("xl/pivotCache/pivotCacheRecords1.xml", DECL
            + f'<pivotCacheRecords xmlns="{NS_MAIN}" xmlns:r="{NS_R}"'
            f' count="{len(rows)}">'
            + "".join("<r>" + "".join(f.record(i) for f in fields) + "</r>"
                      for i in range(len(rows)))
            + "</pivotCacheRecords>")
    pkg.put("xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels", DECL
            + f'<Relationships xmlns="{NS_PKG}"><Relationship Id="rId1"'
            f' Type="{REL_PIVOT_RECORDS}" Target="pivotCacheRecords1.xml"/>'
            f"</Relationships>")

    need_items = set(slicer_fields) | {p.row_field for p in pivots}
    by_sheet = {}
    for pi, p in enumerate(pivots, start=1):
        rf, df = fidx[p.row_field], fidx[p.data_field]
        n = len(fields[rf].items)
        pf = []
        for i, f in enumerate(fields):
            if i == rf:
                it = "".join(f'<item x="{k}"/>' for k in range(n))
                pf.append(f'<pivotField axis="axisRow" showAll="0">'
                          f'<items count="{n + 1}">{it}<item t="default"/>'
                          f"</items></pivotField>")
            elif i == df:
                pf.append('<pivotField dataField="1" showAll="0"/>')
            elif f.name in need_items:
                c = len(f.items)
                it = "".join(f'<item x="{k}"/>' for k in range(c))
                pf.append(f'<pivotField showAll="0"'
                          f' multipleItemSelectionAllowed="1">'
                          f'<items count="{c + 1}">{it}<item t="default"/>'
                          f"</items></pivotField>")
            else:
                pf.append('<pivotField showAll="0"/>')
        ri = "".join("<i><x/></i>" if k == 0 else f'<i><x v="{k}"/></i>'
                     for k in range(n)) + '<i t="grand"><x/></i>'
        loc = (f"{col_letter(p.anchor_col)}{p.anchor_row}:"
               f"{col_letter(p.anchor_col + 1)}{p.anchor_row + n + 1}")
        pkg.put(f"xl/pivotTables/pivotTable{pi}.xml", DECL
                + f'<pivotTableDefinition xmlns="{NS_MAIN}"'
                f' name="{escape(p.name)}" cacheId="1" dataCaption="Valores"'
                f' updatedVersion="6" minRefreshableVersion="3"'
                f' createdVersion="6" applyNumberFormats="0"'
                f' applyBorderFormats="0" applyFontFormats="0"'
                f' applyPatternFormats="0" applyAlignmentFormats="0"'
                f' applyWidthHeightFormats="1" useAutoFormatting="1"'
                f' itemPrintTitles="1" indent="0" outline="1" outlineData="1"'
                f' multipleFieldFilters="0">'
                f'<location ref="{loc}" firstHeaderRow="1" firstDataRow="1"'
                f' firstDataCol="1"/>'
                f'<pivotFields count="{len(fields)}">' + "".join(pf)
                + f'</pivotFields><rowFields count="1"><field x="{rf}"/>'
                f'</rowFields><rowItems count="{n + 1}">{ri}</rowItems>'
                f'<colItems count="1"><i/></colItems>'
                f'<dataFields count="1"><dataField name="{escape(p.caption)}"'
                f' fld="{df}" baseField="0" baseItem="0"/></dataFields>'
                f'<pivotTableStyleInfo name="PivotStyleLight16"'
                f' showRowHeaders="1" showColHeaders="1" showRowStripes="0"'
                f' showColStripes="0" showLastColumn="1"/>'
                f"</pivotTableDefinition>")
        pkg.put(f"xl/pivotTables/_rels/pivotTable{pi}.xml.rels", DECL
                + f'<Relationships xmlns="{NS_PKG}"><Relationship Id="rId1"'
                f' Type="{REL_PIVOT_CACHE}"'
                f' Target="../pivotCache/pivotCacheDefinition1.xml"/>'
                f"</Relationships>")
        by_sheet.setdefault(p.sheet, []).append(pi)

    for sname, plist in by_sheet.items():
        pkg.add_rels(pkg.sheet_part(sname),
                     [(REL_PIVOT_TABLE, f"../pivotTables/pivotTable{pi}.xml")
                      for pi in plist])

    wb = pkg.txt("xl/workbook.xml")
    rid = pkg.add_rels("xl/workbook.xml",
                       [(REL_PIVOT_CACHE,
                         "pivotCache/pivotCacheDefinition1.xml")])[0]
    # openpyxl declara xmlns:r elemento por elemento, no en la raiz
    pc = (f'<pivotCaches><pivotCache xmlns:r="{NS_R}" cacheId="1"'
          f' r:id="{rid}"/></pivotCaches>')
    wb = wb.replace("</workbook>", pc + "</workbook>")
    pkg.put("xl/workbook.xml", wb)

    pkg.add_content_types(
        [("/xl/pivotCache/pivotCacheDefinition1.xml", CT["cache"]),
         ("/xl/pivotCache/pivotCacheRecords1.xml", CT["records"])]
        + [(f"/xl/pivotTables/pivotTable{i}.xml", CT["pivot"])
           for i in range(1, len(pivots) + 1)])
    pkg.save()


# ---------------------------------------------------------------------------
#  PASO 2 - segmentadores + retoques posteriores al round-trip de LibreOffice
# ---------------------------------------------------------------------------
def _shared_items(cache_xml, field):
    """Items compartidos de un campo, en el orden real del cache."""
    m = re.search(r'<cacheField name="%s"[^>]*>(.*?)</cacheField>'
                  % re.escape(field), cache_xml, re.S)
    if not m:
        raise KeyError(f"campo {field} ausente del pivotCache")
    return re.findall(r'<[sn] v="([^"]*)"/>', m.group(1))


def _valores_cacheados(pkg, hoja):
    """Valores ya calculados de una hoja: {'B75': 'texto'}."""
    xml = pkg.txt(pkg.sheet_part(hoja))
    out = {}
    for m in re.finditer(r'<c r="([A-Z]+\d+)"[^>]*>(.*?)</c>', xml, re.S):
        v = re.search(r"<v>(.*?)</v>", m.group(2), re.S)
        if v:
            out[m.group(1)] = v.group(1)
    return out


def titulos_dinamicos(path, titulos, color="1F3864", fuente="Arial"):
    """
    Deja cada grafico con un titulo que apunta a una celda (strRef) y ademas
    lleva el texto ya resuelto (strCache).

    Hay que hacerlo despues del recalculo por dos motivos: LibreOffice no sabe
    resolver un titulo por referencia y lo reemplaza por un literal
    "Chart Title", y el texto que hay que cachear recien existe una vez que las
    formulas se calcularon.

    titulos: [(marcador_presente_en_el_xml_del_grafico, "Hoja!$B$79"), ...]
    """
    pkg = _Pkg(path)
    valores, puestos = {}, 0
    for name in sorted(n for n in pkg.parts
                       if re.match(r"xl/charts/chart\d+\.xml$", n)):
        xml = pkg.txt(name)
        ref = next((celda for marca, celda in titulos if marca in xml), None)
        if ref is None:
            continue
        hoja, celda = ref.replace("$", "").split("!")
        if hoja not in valores:
            valores[hoja] = _valores_cacheados(pkg, hoja)
        texto = valores[hoja].get(celda, "")
        titulo = (
            f"<c:title><c:tx><c:strRef><c:f>{ref}</c:f><c:strCache>"
            f'<c:ptCount val="1"/><c:pt idx="0"><c:v>{texto}</c:v></c:pt>'
            f"</c:strCache></c:strRef></c:tx>"
            f'<c:overlay val="0"/><c:txPr><a:bodyPr rot="0"/><a:lstStyle/>'
            f'<a:p><a:pPr><a:defRPr sz="1000" b="1">'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:latin typeface="{fuente}"/></a:defRPr></a:pPr>'
            f'<a:endParaRPr lang="es-AR"/></a:p></c:txPr></c:title>')

        m = re.search(r"<c:chart>\s*<c:title>.*?</c:title>", xml, re.S)
        if m:
            xml = xml[:m.start()] + "<c:chart>" + titulo + xml[m.end():]
        else:
            xml = xml.replace("<c:chart>", "<c:chart>" + titulo, 1)
        xml = xml.replace('<c:autoTitleDeleted val="1"/>',
                          '<c:autoTitleDeleted val="0"/>')
        pkg.put(name, xml)
        puestos += 1
    pkg.save()
    return puestos


def finalize(path, slicers, slicer_sheet, gridless_sheets=()):
    pkg = _Pkg(path)
    sheets = pkg.sheets()

    # --- LibreOffice descarta refreshOnLoad; sin el, el cache queda congelado
    cname = "xl/pivotCache/pivotCacheDefinition1.xml"
    cache = pkg.txt(cname)
    if "refreshOnLoad" not in cache:
        cache = cache.replace("<pivotCacheDefinition ",
                              '<pivotCacheDefinition refreshOnLoad="1" ', 1)
        pkg.put(cname, cache)

    # --- las tablas dinamicas del archivo, con la hoja donde vive cada una
    pivot_parts = sorted(n for n in pkg.parts
                         if re.match(r"xl/pivotTables/pivotTable\d+\.xml$", n))
    pivot_sheet = {}
    for sname, (idx, _) in sheets.items():
        rp = f"xl/worksheets/_rels/sheet{idx}.xml.rels"
        if rp in pkg.parts:
            for t in re.findall(r'Target="\.\./pivotTables/(pivotTable\d+\.xml)"',
                                pkg.txt(rp)):
                pivot_sheet[f"xl/pivotTables/{t}"] = sname
    pivots = []
    for pn in pivot_parts:
        x = pkg.txt(pn)
        pivots.append((pn, re.search(r'name="([^"]*)"', x).group(1),
                       pivot_sheet.get(pn, slicer_sheet)))

    # --- cada campo con slicer necesita sus <items> en cada tabla dinamica,
    #     que es donde Excel guarda que items quedan seleccionados
    cache_fields = re.findall(r'<cacheField name="([^"]*)"', cache)
    for pn, _, _ in pivots:
        x = pkg.txt(pn)
        head, body, tail = re.match(
            r"(.*?<pivotFields[^>]*>)(.*?)(</pivotFields>.*)", x, re.S).groups()
        # ojo: <item .../> tambien termina en "/>", asi que la alternativa
        # de campo vacio tiene que excluir ">" dentro de los atributos
        fields_xml = re.findall(
            r"<pivotField\b[^>]*/>|<pivotField\b[^>]*>.*?</pivotField>",
            body, re.S)
        out = []
        for i, fx in enumerate(fields_xml):
            fname = cache_fields[i] if i < len(cache_fields) else None
            slicer_field = any(s.field == fname for s in slicers)
            if slicer_field and "<items" not in fx:
                n = len(_shared_items(cache, fname))
                it = "".join(f'<item x="{k}"/>' for k in range(n))
                fx = (f'<pivotField showAll="0" multipleItemSelectionAllowed="1">'
                      f'<items count="{n + 1}">{it}<item t="default"/></items>'
                      f"</pivotField>")
            out.append(fx)
        pkg.put(pn, head + "".join(out) + tail)

    # --- slicerCaches ------------------------------------------------------
    pt_refs = "".join(f'<pivotTable tabId="{sheets[sh][1]}" name="{escape(nm)}"/>'
                      for _, nm, sh in pivots)
    for si, s in enumerate(slicers, start=1):
        items = _shared_items(cache, s.field)
        pkg.put(f"xl/slicerCaches/slicerCache{si}.xml", DECL
                + f'<slicerCacheDefinition xmlns="{NS_X14}" xmlns:mc="{NS_MC}"'
                f' xmlns:x="{NS_MAIN}" mc:Ignorable="x" name="{s.cache_name}"'
                f' sourceName="{escape(s.field)}">'
                f"<pivotTables>{pt_refs}</pivotTables>"
                f'<data><tabular pivotCacheId="1"'
                f' crossFilter="showItemsWithDataAtTop">'
                f'<items count="{len(items)}">'
                + "".join(f'<i x="{k}"/>' for k in range(len(items)))
                + "</items></tabular></data></slicerCacheDefinition>")

    # --- slicers -----------------------------------------------------------
    pkg.put("xl/slicers/slicer1.xml", DECL
            + f'<slicers xmlns="{NS_X14}" xmlns:mc="{NS_MC}" xmlns:x="{NS_MAIN}"'
            f' mc:Ignorable="x">'
            + "".join(f'<slicer name="{escape(s.field)}" cache="{s.cache_name}"'
                      f' caption="{escape(s.caption)}"'
                      f' columnCount="{s.column_count}" style="{s.style}"'
                      f' rowHeight="234950"/>' for s in slicers)
            + "</slicers>")

    # --- formas de los slicers dentro del drawing de la hoja ---------------
    dash_part = pkg.sheet_part(slicer_sheet)
    dash_xml = pkg.txt(dash_part)
    m = re.search(r'<drawing[^>]*r:id="(rId\d+)"', dash_xml)
    if m:
        rp = re.sub(r"(.*/)([^/]+)$", r"\1_rels/\2.rels", dash_part)
        tgt = re.search(r'Id="%s"[^>]*Target="\.\./drawings/([^"]+)"' % m.group(1),
                        pkg.txt(rp)).group(1)
        dpart = f"xl/drawings/{tgt}"
        dxml = pkg.txt(dpart)
        next_id = max([int(v) for v in re.findall(r'<xdr:cNvPr id="(\d+)"', dxml)]
                      or [1]) + 1
        dr_id = None
    else:
        nums = [int(x.group(1)) for n in pkg.parts
                for x in [re.match(r"xl/drawings/drawing(\d+)\.xml$", n)] if x]
        dpart = f"xl/drawings/drawing{max(nums) + 1 if nums else 1}.xml"
        dxml = DECL + f'<xdr:wsDr xmlns:xdr="{NS_XDR}" xmlns:a="{NS_A}"/>'
        dxml = dxml.replace("/>", "></xdr:wsDr>")
        next_id = 2
        dr_id = "pending"

    anchors = ""
    for k, s in enumerate(slicers):
        anchors += (
            f"<xdr:twoCellAnchor>"
            f"<xdr:from><xdr:col>{s.from_col}</xdr:col><xdr:colOff>0</xdr:colOff>"
            f"<xdr:row>{s.from_row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
            f"<xdr:to><xdr:col>{s.to_col}</xdr:col><xdr:colOff>0</xdr:colOff>"
            f"<xdr:row>{s.to_row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>"
            f'<mc:AlternateContent xmlns:mc="{NS_MC}">'
            f'<mc:Choice xmlns:sle="{NS_SLE}" Requires="sle">'
            f'<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>'
            f'<xdr:cNvPr id="{next_id + k}" name="{escape(s.field)}"/>'
            f"<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>"
            f'<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
            f'<a:graphic><a:graphicData uri="{NS_SLE}">'
            f'<sle:slicer xmlns:sle="{NS_SLE}" name="{escape(s.field)}"/>'
            f"</a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice>"
            f'<mc:Fallback><xdr:sp macro="" textlink=""><xdr:nvSpPr>'
            f'<xdr:cNvPr id="{next_id + k}" name="{escape(s.field)}"/>'
            f'<xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr>'
            f'</xdr:nvSpPr><xdr:spPr>'
            f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
            f'<a:ln w="9525"><a:solidFill><a:srgbClr val="8C8C8C"/></a:solidFill>'
            f"</a:ln></xdr:spPr><xdr:txBody>"
            f'<a:bodyPr vertOverflow="clip" horzOverflow="clip"/><a:lstStyle/>'
            f'<a:p><a:r><a:rPr lang="es-AR" sz="1000"/><a:t>'
            f"Segmentador &quot;{escape(s.caption)}&quot;: requiere Excel 2010 "
            f"o posterior.</a:t></a:r></a:p></xdr:txBody></xdr:sp>"
            f"</mc:Fallback></mc:AlternateContent><xdr:clientData/>"
            f"</xdr:twoCellAnchor>")
    pkg.put(dpart, dxml.replace("</xdr:wsDr>", anchors + "</xdr:wsDr>"))

    new_rels = [(REL_SLICER, "../slicers/slicer1.xml")]
    if dr_id:
        new_rels.insert(0, (REL_DRAWING, f"../drawings/{dpart.split('/')[-1]}"))
    ids = pkg.add_rels(dash_part, new_rels)
    sl_id = ids[-1]

    ext = (f'<extLst><ext uri="{{A8765BA9-456A-4dab-B4F3-ACF1056F45CD}}"'
           f' xmlns:x14="{NS_X14}"><x14:slicerList>'
           f'<x14:slicer xmlns:r="{NS_R}" r:id="{sl_id}"/>'
           f"</x14:slicerList></ext></extLst>")
    tail = ext
    if dr_id:
        tail = f'<drawing xmlns:r="{NS_R}" r:id="{ids[0]}"/>' + ext
    dash_xml = pkg.txt(dash_part).replace("</worksheet>", tail + "</worksheet>")
    pkg.put(dash_part, dash_xml)

    # --- workbook: slicerCaches + recalculo forzado al abrir ---------------
    rids = pkg.add_rels("xl/workbook.xml",
                        [(REL_SLICER_CACHE, f"slicerCaches/slicerCache{i}.xml")
                         for i in range(1, len(slicers) + 1)])
    wb = pkg.txt("xl/workbook.xml")
    sc = (f'<extLst><ext uri="{{BBE1A952-AA13-448e-AADC-164F8A28A991}}"'
          f' xmlns:x14="{NS_X14}"><x14:slicerCaches>'
          + "".join(f'<x14:slicerCache xmlns:r="{NS_R}" r:id="{r}"/>'
                    for r in rids)
          + "</x14:slicerCaches></ext></extLst>")
    if re.search(r"<calcPr[^>]*/>", wb):
        wb = re.sub(r"<calcPr[^>]*/>",
                    '<calcPr calcId="171027" fullCalcOnLoad="1"/>', wb)
    else:
        wb = wb.replace("</workbook>",
                        '<calcPr calcId="171027" fullCalcOnLoad="1"/></workbook>')
    pkg.put("xl/workbook.xml", wb.replace("</workbook>", sc + "</workbook>"))

    # --- LibreOffice reactiva las lineas de cuadricula al guardar ---------
    for sname in gridless_sheets:
        pn = pkg.sheet_part(sname)
        x = pkg.txt(pn)
        if "showGridLines" in x:
            x = re.sub(r'showGridLines="[^"]*"', 'showGridLines="0"', x, count=1)
        else:
            x = x.replace("<sheetView ", '<sheetView showGridLines="0" ', 1)
        pkg.put(pn, x)

    pkg.add_content_types(
        [(f"/{dpart}", CT["drawing"]),
         ("/xl/slicers/slicer1.xml", CT["slicer"])]
        + [(f"/xl/slicerCaches/slicerCache{i}.xml", CT["slicerCache"])
           for i in range(1, len(slicers) + 1)])
    pkg.save()
