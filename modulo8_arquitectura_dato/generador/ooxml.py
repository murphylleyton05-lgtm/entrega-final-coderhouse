"""
Retoques OOXML que openpyxl no sabe hacer.

Son tres:

1. `cachear`  -> escribe el valor calculado dentro de las celdas con formula.
   openpyxl guarda la formula pero no su resultado, asi que un visor que no
   recalcula (la vista previa de GitHub, LibreOffice sin recalculo, Excel para
   web en modo lectura) mostraria las celdas vacias. Como todos los numeros ya
   estan calculados en `simulacion.py`, se inyectan como valor cacheado.

2. `inyectar_escenarios` -> el Administrador de escenarios de Excel. Los
   escenarios viven en el XML de la hoja, en un elemento `<scenarios>` que
   openpyxl no genera. Se arma a mano y se inserta respetando el orden que
   exige el esquema de SpreadsheetML (despues de `sheetData`, antes de
   `mergeCells`).

3. `calcular_al_abrir` -> marca el libro para recalculo completo al abrirlo,
   que es lo que hace que la celda de la ruta se resuelva sola en la maquina
   del evaluador en vez de mostrar la ruta de quien genero el archivo.
"""
import re
import shutil
import zipfile
from xml.sax.saxutils import escape, quoteattr

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

# lugares donde puede ir <scenarios>, en el orden en que hay que buscarlos:
# el esquema lo ubica despues de sheetData y antes de mergeCells
ANCLAS_SCENARIOS = ["<mergeCells", "<conditionalFormatting", "<dataValidations",
                    "<hyperlinks", "<printOptions", "<pageMargins", "<pageSetup",
                    "</worksheet>"]


def _mapa_de_hojas(z):
    """{nombre visible de la hoja: ruta de su XML dentro del zip}"""
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")

    destinos = {}
    for m in re.finditer(r'<Relationship\b[^>]*>', rels):
        tag = m.group(0)
        rid = re.search(r'Id="([^"]+)"', tag)
        target = re.search(r'Target="([^"]+)"', tag)
        if rid and target:
            ruta = target.group(1)
            if ruta.startswith("/"):
                ruta = ruta[1:]
            elif not ruta.startswith("xl/"):
                ruta = "xl/" + ruta.lstrip("./")
            destinos[rid.group(1)] = ruta

    hojas = {}
    for m in re.finditer(r'<sheet\b[^>]*/>', wb):
        tag = m.group(0)
        nombre = re.search(r'name="([^"]+)"', tag)
        rid = re.search(r'r:id="([^"]+)"', tag)
        if nombre and rid:
            hojas[nombre.group(1)] = destinos[rid.group(1)]
    return hojas


def _reescribir(ruta_xlsx, transformaciones, agregados=None):
    """Reescribe el zip aplicando `transformaciones[nombre_parte] -> texto`."""
    tmp = ruta_xlsx + ".tmp"
    with zipfile.ZipFile(ruta_xlsx) as origen:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as destino:
            for info in origen.infolist():
                datos = origen.read(info.filename)
                fn = transformaciones.get(info.filename)
                if fn:
                    datos = fn(datos.decode("utf-8")).encode("utf-8")
                destino.writestr(info, datos)
            for nombre, contenido in (agregados or {}).items():
                destino.writestr(nombre, contenido)
    shutil.move(tmp, ruta_xlsx)


def _formatear(valor):
    if isinstance(valor, bool):
        return "1" if valor else "0", "b"
    if isinstance(valor, (int,)):
        return str(valor), None
    if isinstance(valor, float):
        if valor == int(valor) and abs(valor) < 1e15:
            return str(int(valor)), None
        return repr(valor), None
    return escape(str(valor)), "str"


# Una celda con contenido. El `[^/]` antes del cierre descarta las celdas
# vacias autocerradas (`<c r="D52"/>`): sin eso el `.*?` se las tragaria hasta
# el siguiente `</c>` y el XML quedaria corrupto.
RE_CELDA = re.compile(r'<c\b[^>]*[^/]>.*?</c>', re.DOTALL)


def _cachear_hoja(xml, valores):
    """Agrega `<v>` a las celdas con formula de una hoja."""
    aplicados = 0

    def reemplazo(m):
        nonlocal aplicados
        celda = m.group(0)
        ref = re.search(r'\br="([A-Z]+\d+)"', celda)
        if not ref or ref.group(1) not in valores or "<f" not in celda:
            return celda
        # openpyxl deja un `<v />` vacio en toda celda con formula
        celda = re.sub(r'<v\s*/>', "", celda)
        if "<v>" in celda:
            return celda
        texto, tipo = _formatear(valores[ref.group(1)])
        celda = celda.replace("</c>", f"<v>{texto}</v></c>")
        if tipo:
            celda = re.sub(r'\s+t="[^"]*"', "", celda, count=1)
            celda = celda.replace("<c ", f'<c t="{tipo}" ', 1)
        aplicados += 1
        return celda

    return RE_CELDA.sub(reemplazo, xml), aplicados


def cachear(ruta_xlsx, mapa):
    """
    `mapa` es {nombre de hoja: {referencia de celda: valor}}.

    Devuelve cuantas celdas quedaron cacheadas. Si el numero no coincide con
    lo esperado, el pipeline corta: significa que una formula cambio de lugar.
    """
    with zipfile.ZipFile(ruta_xlsx) as z:
        hojas = _mapa_de_hojas(z)

    faltantes = set(mapa) - set(hojas)
    if faltantes:
        raise ValueError(f"hojas inexistentes: {sorted(faltantes)}")

    total = {"n": 0}
    transformaciones = {}
    for nombre, valores in mapa.items():
        def hacer(valores=valores):
            def fn(xml):
                nuevo, aplicados = _cachear_hoja(xml, valores)
                total["n"] += aplicados
                return nuevo
            return fn
        transformaciones[hojas[nombre]] = hacer()

    _reescribir(ruta_xlsx, transformaciones)
    return total["n"]


def _xml_escenarios(escenarios, sqref, actual=1, usuario="Lleyton Murphy"):
    partes = [f'<scenarios current="{actual}" show="{actual}" sqref="{sqref}">']
    for nombre, comentario, celdas in escenarios:
        partes.append(
            f'<scenario name={quoteattr(nombre)} locked="1" count="{len(celdas)}"'
            f' user={quoteattr(usuario)} comment={quoteattr(comentario)}>')
        for ref, valor in celdas:
            partes.append(f'<inputCells r="{ref}" val="{valor}"/>')
        partes.append('</scenario>')
    partes.append('</scenarios>')
    return "".join(partes)


def inyectar_escenarios(ruta_xlsx, hoja, escenarios, sqref, actual=1):
    """
    Agrega el bloque `<scenarios>` a una hoja.

    `escenarios` es una lista de (nombre, comentario, [(celda, valor), ...]).
    """
    bloque = _xml_escenarios(escenarios, sqref, actual)

    with zipfile.ZipFile(ruta_xlsx) as z:
        parte = _mapa_de_hojas(z)[hoja]

    def fn(xml):
        if "<scenarios" in xml:
            raise ValueError(f"la hoja {hoja} ya tiene escenarios")
        for ancla in ANCLAS_SCENARIOS:
            i = xml.find(ancla)
            if i != -1:
                return xml[:i] + bloque + xml[i:]
        raise ValueError("no se encontro donde insertar <scenarios>")

    _reescribir(ruta_xlsx, {parte: fn})
    return len(escenarios)


def calcular_al_abrir(ruta_xlsx):
    """
    Fuerza el recalculo completo al abrir.

    Importa por la celda de la ruta: se arma con CELL("filename"), que devuelve
    la carpeta de quien tiene el archivo abierto. Sin recalculo al abrir, el
    evaluador veria cacheada la ruta de la maquina que genero el libro.
    """
    def fn(xml):
        if "<calcPr" in xml:
            return re.sub(r'<calcPr\b[^>]*/>',
                          '<calcPr calcId="191029" fullCalcOnLoad="1"/>', xml, count=1)
        return xml.replace("</workbook>",
                           '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')

    _reescribir(ruta_xlsx, {"xl/workbook.xml": fn})


def resumen(ruta_xlsx):
    """Inventario de partes del archivo, para el chequeo final del pipeline."""
    with zipfile.ZipFile(ruta_xlsx) as z:
        nombres = z.namelist()
        hojas = _mapa_de_hojas(z)
        escenarios = {}
        formulas = 0
        cacheadas = 0
        for nombre, parte in hojas.items():
            xml = z.read(parte).decode("utf-8")
            n = xml.count("<scenario ")
            if n:
                escenarios[nombre] = n
            for celda in RE_CELDA.findall(xml):
                if "<f" in celda:
                    formulas += 1
                    if "<v>" in celda:
                        cacheadas += 1
    return {"partes": nombres, "hojas": list(hojas), "escenarios": escenarios,
            "formulas": formulas, "formulas_cacheadas": cacheadas}
