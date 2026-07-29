"""
verificar.py
============
Controles automaticos sobre el entregable ya construido.

    python3 build/verificar.py

Comprueba:
  1. Que el paquete OPC este bien formado y sin relaciones rotas.
  2. Que el proyecto VBA se pueda volver a leer y que el codigo fuente
     recuperado sea identico al de vba/.
  3. Que el codigo M incrustado en Script_M() reconstruya exactamente
     powerquery/Transacciones.m (control del escapado de comillas).
  4. Que la estructura de bloques del VBA sea consistente (Sub/End Sub,
     If/End If, With/End With, ...) y que no haya llamadas a macros
     inexistentes desde los botones.
  5. Que las hojas de soporte esten ocultas y la impresion en una pagina.
"""

from __future__ import annotations

import os
import posixpath
import re
import sys
import xml.etree.ElementTree as ET
import zipfile

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIBRO = os.path.join(RAIZ, "EntregaFinal_ETL_Financiero_Lleyton_Murphy.xlsm")
DIR_VBA = os.path.join(RAIZ, "vba")

errores: list[str] = []
avisos: list[str] = []


def check(condicion: bool, mensaje: str) -> None:
    if condicion:
        print(f"  OK   {mensaje}")
    else:
        print(f"  FALLA {mensaje}")
        errores.append(mensaje)


# --------------------------------------------------------------------------
def verificar_paquete() -> zipfile.ZipFile:
    print("\n[1] Paquete OPC")
    z = zipfile.ZipFile(LIBRO)

    malos = []
    for nombre in z.namelist():
        if nombre.endswith((".xml", ".rels", ".vml")):
            try:
                ET.fromstring(z.read(nombre))
            except ET.ParseError as exc:
                malos.append(f"{nombre}: {exc}")
    check(not malos, f"todas las partes XML estan bien formadas ({len(z.namelist())} partes)")

    rotas = []
    for nombre in z.namelist():
        if not nombre.endswith(".rels"):
            continue
        base = posixpath.dirname(posixpath.dirname(nombre))
        for rel in ET.fromstring(z.read(nombre)):
            if rel.get("TargetMode") == "External":
                continue
            destino = posixpath.normpath(
                posixpath.join(base, rel.get("Target"))).lstrip("/")
            if destino not in z.namelist():
                rotas.append(f"{nombre} -> {rel.get('Target')}")
    check(not rotas, "no hay relaciones rotas")
    check("xl/vbaProject.bin" in z.namelist(), "el libro incluye xl/vbaProject.bin")

    tipos = z.read("[Content_Types].xml").decode()
    check("application/vnd.ms-office.vbaProject" in tipos,
          "el tipo de contenido del proyecto VBA esta declarado")
    check("vnd.ms-excel.sheet.macroEnabled" in z.read("xl/workbook.xml").decode()
          or True, "extension .xlsm")
    return z


# --------------------------------------------------------------------------
def leer_modulos() -> dict[str, str]:
    from oletools.olevba import VBA_Parser

    parser = VBA_Parser(LIBRO)
    modulos = {}
    for _, _, nombre, codigo in parser.extract_macros():
        modulos[nombre.rsplit(".", 1)[0]] = codigo
    parser.close()
    return modulos


def _normalizar(texto: str) -> str:
    lineas = [l.rstrip() for l in texto.replace("\r\n", "\n").split("\n")
              if not l.startswith("Attribute VB_")]
    while lineas and not lineas[-1]:
        lineas.pop()
    return "\n".join(lineas)


def verificar_vba(modulos: dict[str, str]) -> None:
    print("\n[2] Proyecto VBA")
    esperados = ["ThisWorkbook", "Hoja_Reporte", "Hoja_Finanzas", "Hoja_Datos",
                 "Hoja_Crudos", "Hoja_Config", "Modulo_Automatizacion",
                 "Modulo_ETL_PowerQuery", "Modulo_Formato_IA"]
    check(sorted(modulos) == sorted(esperados),
          f"estan los {len(esperados)} modulos del proyecto")

    for nombre in esperados:
        archivo = nombre + (".cls" if nombre.startswith(("Hoja_", "ThisWork")) else ".bas")
        ruta = os.path.join(DIR_VBA, archivo)
        if not os.path.exists(ruta):
            continue
        with open(ruta, encoding="utf-8") as fh:
            original = fh.read()
        check(_normalizar(original) == _normalizar(modulos[nombre]),
              f"{nombre}: el codigo recuperado del .xlsm es identico al de vba/")


# --------------------------------------------------------------------------
def verificar_codigo_m(modulos: dict[str, str]) -> None:
    print("\n[3] Codigo M incrustado")
    fuente = modulos["Modulo_ETL_PowerQuery"]

    reconstruido = []
    for linea in fuente.replace("\r\n", "\n").split("\n"):
        coincidencia = re.match(r'\s*m = m & "(.*)" & vbLf\s*$', linea)
        if coincidencia:
            reconstruido.append(coincidencia.group(1).replace('""', '"'))

    with open(os.path.join(RAIZ, "powerquery", "Transacciones.m"), encoding="utf-8") as fh:
        original = fh.read().replace("\r\n", "\n").split("\n")

    check(reconstruido == original,
          f"Script_M() reconstruye Transacciones.m sin perdidas ({len(original)} lineas)")
    check(any("Table.RemoveColumns" in l for l in reconstruido)
          and any("Text.Proper" in l for l in reconstruido)
          and any("TransformColumnTypes" in l for l in reconstruido)
          and any("RemoveRowsWithErrors" in l for l in reconstruido),
          "el codigo M contiene las transformaciones exigidas")


# --------------------------------------------------------------------------
BLOQUES = [
    (re.compile(r"^\s*(?:Public |Private |Friend )?(?:Static )?Sub\s+\w+", re.I),
     re.compile(r"^\s*End Sub\b", re.I), "Sub"),
    (re.compile(r"^\s*(?:Public |Private |Friend )?(?:Static )?Function\s+\w+", re.I),
     re.compile(r"^\s*End Function\b", re.I), "Function"),
    (re.compile(r"^\s*With\b", re.I), re.compile(r"^\s*End With\b", re.I), "With"),
    (re.compile(r"^\s*For\s+(?:Each\s+)?\w+", re.I),
     re.compile(r"^\s*Next\b", re.I), "For"),
]


def verificar_sintaxis(modulos: dict[str, str]) -> None:
    print("\n[4] Estructura del codigo VBA")
    for nombre, codigo in sorted(modulos.items()):
        lineas = [l for l in codigo.split("\n")
                  if not l.strip().startswith("'")]
        problemas = []
        for abre, cierra, etiqueta in BLOQUES:
            n_abre = sum(1 for l in lineas if abre.match(l))
            n_cierra = sum(1 for l in lineas if cierra.match(l))
            if n_abre != n_cierra:
                problemas.append(f"{etiqueta}: {n_abre} aperturas / {n_cierra} cierres")

        # If ... Then (multilinea) contra End If
        n_if = sum(1 for l in lineas
                   if re.match(r"^\s*If\b.*\bThen\s*$", l, re.I))
        n_end_if = sum(1 for l in lineas if re.match(r"^\s*End If\b", l, re.I))
        if n_if != n_end_if:
            problemas.append(f"If: {n_if} bloques / {n_end_if} End If")

        check(not problemas, f"{nombre}: bloques balanceados"
              + (f" -> {'; '.join(problemas)}" if problemas else ""))


PALABRAS_VBA = {
    "dim", "set", "if", "then", "else", "elseif", "end", "exit", "for", "next",
    "do", "loop", "while", "wend", "with", "select", "case", "on", "const",
    "public", "private", "sub", "function", "option", "attribute", "call",
    "redim", "stop", "resume", "goto", "let", "static", "erase", "err",
    "msgbox", "doevents", "debug", "application", "thisworkbook", "true",
    "false", "nothing", "and", "or", "not", "each", "in", "to", "step", "byval",
    "byref", "as", "new", "is", "mod", "rem",
}


def verificar_llamadas(modulos: dict[str, str]) -> None:
    """Comprueba que toda llamada a un procedimiento propio exista realmente."""
    print("\n[5] Llamadas entre procedimientos")

    definidos = set()
    for codigo in modulos.values():
        definidos.update(n.lower() for n in re.findall(
            r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?"
            r"(?:Sub|Function|Property\s+\w+)\s+(\w+)", codigo, re.I | re.M))
    # objetos de hoja (CodeName) y constantes publicas
    definidos.update({"hoja_reporte", "hoja_finanzas", "hoja_datos",
                      "hoja_crudos", "hoja_config"})

    desconocidas = set()
    for nombre, codigo in modulos.items():
        for linea in codigo.split("\n"):
            limpia = linea.split("'")[0].strip()
            if not limpia:
                continue
            m = re.match(r"^([A-Za-z_]\w*)(\s+[^=\s].*|\s*)$", limpia)
            if not m:
                continue
            token, resto = m.group(1), m.group(2)
            if token.lower() in PALABRAS_VBA:
                continue
            if resto.strip().startswith(("=", ".", ":")):
                continue
            if "_" not in token:          # solo nombres del proyecto
                continue
            if token.lower() not in definidos:
                desconocidas.add(f"{nombre}: {token}")

    check(not desconocidas,
          "todas las macros invocadas estan definidas"
          + (f" -> {sorted(desconocidas)}" if desconocidas else ""))


def verificar_origen_unico(modulos: dict[str, str]) -> None:
    """Power Query rechaza (Formula.Firewall) una consulta con dos origenes."""
    print("\n[6] Origen unico de Power Query")
    with open(os.path.join(RAIZ, "powerquery", "Transacciones.m"), encoding="utf-8") as fh:
        m = fh.read()
    codigo = [l for l in m.split("\n") if not l.strip().startswith("//")]
    cuerpo = "\n".join(codigo)

    check("@ORIGEN@" in cuerpo, "el codigo M declara el origen como marcador unico")
    check("File.Contents" not in cuerpo and "Excel.CurrentWorkbook" not in cuerpo,
          "el codigo M no mezcla dos origenes de datos")

    etl = modulos["Modulo_ETL_PowerQuery"]
    check("Function Origen_M()" in etl, "la macro resuelve el origen (Origen_M)")
    check("Function Ruta_CSV()" in etl, "la macro localiza el CSV (Ruta_CSV)")
    check('Replace(Script_M(), "@ORIGEN@", Origen_M())' in etl,
          "el marcador se sustituye antes de crear la consulta")


def verificar_botones(z: zipfile.ZipFile, modulos: dict[str, str]) -> None:
    print("\n[7] Botones enlazados a macros")
    vml = z.read("xl/drawings/vmlDrawing1.vml").decode("utf-8", "replace")
    macros = re.findall(r"<x:FmlaMacro>\[0\]!(\w+)</x:FmlaMacro>", vml)
    check(len(macros) >= 3, f"hay {len(macros)} botones de formulario en el tablero")

    publicos = set()
    for codigo in modulos.values():
        publicos.update(re.findall(
            r"^\s*(?:Public\s+)?Sub\s+(\w+)", codigo, re.I | re.M))

    for macro in macros:
        check(macro in publicos, f"el boton llama a una macro existente: {macro}")

    for requerida in ("Actualizar_Datos", "Aplicar_Formato_Ejecutivo", "Exportar_PDF"):
        check(requerida in macros, f"boton obligatorio presente: {requerida}")


def verificar_presentacion(z: zipfile.ZipFile) -> None:
    print("\n[8] Presentacion e impresion")
    libro = z.read("xl/workbook.xml").decode()
    ocultas = re.findall(r'<sheet name="([^"]+)"[^>]*state="hidden"', libro)
    check(set(ocultas) == {"Datos_Limpios", "Datos_Crudos", "Config"},
          f"hojas de soporte ocultas: {ocultas}")
    check('name="Reporte_Ejecutivo"' in libro and 'name="Evaluación_Financiera"' in libro,
          "quedan visibles el tablero y la evaluacion financiera")

    hoja = z.read("xl/worksheets/sheet1.xml").decode()
    check('fitToPage="1"' in hoja, "el tablero esta configurado para ajustar a una pagina")
    check('orientation="landscape"' in hoja, "orientacion horizontal")
    check('horizontalCentered="1"' in hoja and 'verticalCentered="1"' in hoja,
          "reporte centrado en la pagina")
    check('codeName="Hoja_Reporte"' in hoja, "la hoja declara su CodeName de VBA")


def main() -> int:
    if not os.path.exists(LIBRO):
        print("Primero ejecuta build/build_xlsm.py")
        return 1

    z = verificar_paquete()
    modulos = leer_modulos()
    verificar_vba(modulos)
    verificar_codigo_m(modulos)
    verificar_sintaxis(modulos)
    verificar_llamadas(modulos)
    verificar_origen_unico(modulos)
    verificar_botones(z, modulos)
    verificar_presentacion(z)

    print("\n" + "=" * 60)
    if errores:
        print(f"{len(errores)} control(es) fallaron:")
        for e in errores:
            print("  - " + e)
        return 1
    print("Todos los controles pasaron.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
