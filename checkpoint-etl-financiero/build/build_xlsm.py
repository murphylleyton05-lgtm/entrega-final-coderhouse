"""
build_xlsm.py
=============
Construye el entregable `EntregaFinal_ETL_Financiero_Lleyton_Murphy.xlsm`.

    python3 build/build_xlsm.py

Que arma, en orden:

  1. El origen crudo `datos/transacciones_crudas.csv` (datos sucios).
  2. El proyecto VBA real (`vbaProject.bin`) a partir de los fuentes de
     `vba/`, con el codigo M de `powerquery/Transacciones.m` inyectado.
  3. El libro .xlsm: cinco hojas, la tabla de datos crudos, el modelo
     financiero con PAGO / VNA / TIR / VA, los botones de formulario
     enlazados a las macros y la configuracion de impresion en una pagina.
"""

from __future__ import annotations

import csv
import os
import sys
from datetime import date

import xlsxwriter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import datos_sinteticos as ds
import vba_writer as vw

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DIR_BUILD = os.path.join(RAIZ, "build")
DIR_VBA = os.path.join(RAIZ, "vba")
SALIDA = os.path.join(RAIZ, "EntregaFinal_ETL_Financiero_Lleyton_Murphy.xlsm")

AUTOR = "Lleyton Murphy"

# --------------------------------------------------------------------------
# Supuestos del modelo financiero
# --------------------------------------------------------------------------
INVERSION_INICIAL = 500_000.0
PCT_FINANCIADO = 0.70
TNA = 0.12
PLAZO_ANIOS = 5
CRECIMIENTO = 0.08
MARGEN_EBITDA = 0.28
TASA_DESCUENTO = 0.15
IMPUESTO = 0.30


# --------------------------------------------------------------------------
# Matematica financiera (para dejar los resultados en cache en el libro)
# --------------------------------------------------------------------------
def pago(tasa: float, nper: int, va: float) -> float:
    """Equivalente a PAGO()/PMT(): cuota constante de un prestamo."""
    return va * tasa / (1 - (1 + tasa) ** -nper)


def vna(tasa: float, flujos: list[float]) -> float:
    """Equivalente a VNA()/NPV(): descuenta desde el periodo 1."""
    return sum(f / (1 + tasa) ** (i + 1) for i, f in enumerate(flujos))


def tir(flujos: list[float]) -> float:
    """Equivalente a TIR()/IRR() por biseccion."""
    def valor(t: float) -> float:
        return sum(f / (1 + t) ** i for i, f in enumerate(flujos))

    bajo, alto = -0.9999, 10.0
    for _ in range(200):
        medio = (bajo + alto) / 2
        if valor(bajo) * valor(medio) <= 0:
            alto = medio
        else:
            bajo = medio
    return (bajo + alto) / 2


def va(tasa: float, nper: int, cuota: float) -> float:
    """Equivalente a VA()/PV()."""
    return -cuota * (1 - (1 + tasa) ** -nper) / tasa


# --------------------------------------------------------------------------
# 1. Origen crudo
# --------------------------------------------------------------------------
def escribir_csv(filas: list[dict]) -> str:
    ruta = os.path.join(RAIZ, "datos", "transacciones_crudas.csv")
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    with open(ruta, "w", newline="", encoding="utf-8") as fh:
        escritor = csv.DictWriter(fh, fieldnames=ds.COLUMNAS)
        escritor.writeheader()
        escritor.writerows(filas)
    return ruta


# --------------------------------------------------------------------------
# 2. Proyecto VBA
# --------------------------------------------------------------------------
def _leer_fuente(nombre: str) -> str:
    """Lee un fuente de vba/ y descarta las lineas Attribute (las agrega el escritor)."""
    with open(os.path.join(DIR_VBA, nombre), encoding="utf-8") as fh:
        lineas = fh.read().splitlines()
    return "\n".join(l for l in lineas if not l.startswith("Attribute VB_")) + "\n"


def _generar_modulo_etl() -> str:
    """Inyecta el codigo M dentro de Script_M() y deja el .bas listo en vba/."""
    with open(os.path.join(RAIZ, "powerquery", "Transacciones.m"), encoding="utf-8") as fh:
        m = fh.read()

    lineas_vba = []
    for linea in m.replace("\r\n", "\n").split("\n"):
        # En VBA las comillas dobles se escriben duplicadas.
        literal = linea.replace('"', '""')
        lineas_vba.append(f'    m = m & "{literal}" & vbLf')

    with open(os.path.join(DIR_BUILD, "plantilla_Modulo_ETL_PowerQuery.bas"),
              encoding="utf-8") as fh:
        plantilla = fh.read()

    completo = plantilla.replace("'@@SCRIPT_M@@", "\n".join(lineas_vba))
    destino = os.path.join(DIR_VBA, "Modulo_ETL_PowerQuery.bas")
    with open(destino, "w", encoding="utf-8") as fh:
        fh.write(completo)
    return completo


def construir_vba() -> str:
    _generar_modulo_etl()

    hojas = [
        ("Hoja_Reporte", "Hoja_Reporte.cls"),
        ("Hoja_Finanzas", "Hoja_Finanzas.cls"),
        ("Hoja_Datos", "Hoja_Datos.cls"),
        ("Hoja_Crudos", "Hoja_Crudos.cls"),
        ("Hoja_Config", "Hoja_Config.cls"),
    ]

    modulos = [vw.Modulo("ThisWorkbook", _leer_fuente("ThisWorkbook.cls"),
                         "document", vw.GUID_LIBRO)]
    for nombre, archivo in hojas:
        modulos.append(vw.Modulo(nombre, _leer_fuente(archivo), "document", vw.GUID_HOJA))
    for nombre in ("Modulo_Automatizacion", "Modulo_ETL_PowerQuery", "Modulo_Formato_IA"):
        modulos.append(vw.Modulo(nombre, _leer_fuente(nombre + ".bas")))

    binario = vw.construir_vba_project(modulos)
    ruta = os.path.join(DIR_BUILD, "vbaProject.bin")
    with open(ruta, "wb") as fh:
        fh.write(binario)
    return ruta


# --------------------------------------------------------------------------
# 3. Libro
# --------------------------------------------------------------------------
def _formatos(libro):
    azul = "#0B2A4A"
    celeste = "#008AB3"
    gris = "#F4F4F7"
    base = {"font_name": "Segoe UI"}

    def f(**extra):
        return libro.add_format({**base, **extra})

    return {
        "titulo": f(bold=True, font_size=20, font_color="white", bg_color=azul, valign="vcenter"),
        "subtitulo": f(font_size=11, font_color="#D8E4F0", bg_color=azul, valign="vcenter"),
        "seccion": f(bold=True, font_size=11, font_color="white", bg_color=celeste,
                     valign="vcenter", indent=1),
        "estado_lbl": f(bold=True, font_size=9, font_color="#5A5A5A", align="right"),
        "estado": f(italic=True, font_size=9, bg_color=gris, indent=1),
        "kpi_tit": f(bold=True, font_size=9, font_color=azul, bg_color=gris,
                     align="center", valign="vcenter", top=2, left=2, right=2,
                     border_color=azul),
        "kpi_val": f(bold=True, font_size=16, font_color=celeste, bg_color=gris,
                     align="center", valign="vcenter", left=2, right=2,
                     border_color=azul, num_format="$#,##0"),
        "kpi_val_pct": f(bold=True, font_size=16, font_color=celeste, bg_color=gris,
                         align="center", valign="vcenter", left=2, right=2,
                         border_color=azul, num_format="0.0%"),
        "kpi_val_num": f(bold=True, font_size=16, font_color=celeste, bg_color=gris,
                         align="center", valign="vcenter", left=2, right=2,
                         border_color=azul, num_format="#,##0"),
        "kpi_pie": f(font_size=8, font_color="#7A7A7A", bg_color=gris, align="center",
                     bottom=2, left=2, right=2, border_color=azul),
        "nota": f(font_size=9, font_color="#7A7A7A", italic=True),
        "etiqueta": f(font_size=10, indent=1),
        "etiqueta_b": f(font_size=10, bold=True, indent=1),
        "moneda": f(num_format="$#,##0", align="right"),
        "moneda2": f(num_format="$#,##0.00", align="right"),
        "porcentaje": f(num_format="0.0%", align="right"),
        "entero": f(num_format="#,##0", align="right"),
        "entrada": f(num_format="$#,##0", align="right", font_color="#0000C8"),
        "entrada_pct": f(num_format="0.0%", align="right", font_color="#0000C8"),
        "entrada_num": f(num_format="#,##0", align="right", font_color="#0000C8"),
        "destacado": f(bold=True, font_size=11, num_format="$#,##0", align="right",
                       bg_color=gris, border=1, border_color=azul),
        "destacado_pct": f(bold=True, font_size=11, num_format="0.0%", align="right",
                           bg_color=gris, border=1, border_color=azul),
        "destacado_lbl": f(bold=True, font_size=11, bg_color=gris, border=1,
                           border_color=azul, indent=1),
        "cab_tabla": f(bold=True, bg_color=gris, align="center", border=1,
                       border_color="#BEC8D2"),
        "celda_flujo": f(num_format="$#,##0", align="right", border=1,
                         border_color="#BEC8D2"),
        "celda_flujo_b": f(num_format="$#,##0", align="right", border=1, bold=True,
                           border_color="#BEC8D2", bg_color=gris),
        "concepto": f(indent=1, border=1, border_color="#BEC8D2"),
        "concepto_b": f(indent=1, bold=True, border=1, border_color="#BEC8D2",
                        bg_color=gris),
        "conclusion": f(bold=True, font_size=11, font_color="#008060", text_wrap=True,
                        valign="vcenter"),
        "cfg_tit": f(bold=True, font_size=12, font_color="white", bg_color=azul, indent=1),
        "cfg_lbl": f(bold=True, font_size=10, indent=1),
        "cfg_val": f(font_size=10, bg_color="#FFF8DC", border=1, border_color="#C8C8C8"),
        "mono": libro.add_format({"font_name": "Consolas", "font_size": 9}),
        "mono_com": libro.add_format({"font_name": "Consolas", "font_size": 9,
                                      "font_color": "#008000"}),
    }


def _suma_campo(campo: str) -> str:
    col = f'INDEX(Datos_Limpios!$A:$Z,0,MATCH("{campo}",Datos_Limpios!$1:$1,0))'
    return f"IFERROR(SUM({col}),0)"


def construir_libro(ruta_vba: str, crudo: list[dict], limpio: list[dict]) -> None:
    libro = xlsxwriter.Workbook(SALIDA)
    libro.set_vba_name("ThisWorkbook")
    libro.add_vba_project(ruta_vba)
    libro.set_properties({
        "title": "Reporte Ejecutivo - Analisis de Expansion Comercial",
        "subject": "Checkpoint: Flujo ETL, Analisis Financiero y Automatizacion",
        "author": AUTOR,
        "company": "Coderhouse",
        "comments": "Libro habilitado para macros. Power Query + VBA + Tabla dinamica.",
    })

    fmt = _formatos(libro)

    # ----- datos agregados que se guardan en cache ------------------------
    ventas = sum(f["Monto"] for f in limpio)
    margen = sum(f["Margen"] for f in limpio)
    ticket = ventas / len(limpio)

    hoja_rep = libro.add_worksheet("Reporte_Ejecutivo")
    hoja_fin = libro.add_worksheet("Evaluación_Financiera")
    hoja_lim = libro.add_worksheet("Datos_Limpios")
    hoja_cru = libro.add_worksheet("Datos_Crudos")
    hoja_cfg = libro.add_worksheet("Config")

    for hoja, nombre in ((hoja_rep, "Hoja_Reporte"), (hoja_fin, "Hoja_Finanzas"),
                         (hoja_lim, "Hoja_Datos"), (hoja_cru, "Hoja_Crudos"),
                         (hoja_cfg, "Hoja_Config")):
        hoja.set_vba_name(nombre)

    _hoja_reporte(hoja_rep, fmt, ventas, margen, ticket, len(limpio))
    valores = _hoja_financiera(hoja_fin, fmt, ventas)
    _hoja_datos_limpios(hoja_lim, fmt)
    _hoja_datos_crudos(hoja_cru, fmt, crudo)
    _hoja_config(hoja_cfg, fmt)

    libro.define_name("par_RutaCSV", "=Config!$C$6")
    libro.define_name("par_CarpetaLibro", "=Config!$C$7")

    hoja_lim.hide()
    hoja_cru.hide()
    hoja_cfg.hide()
    hoja_rep.activate()

    libro.close()
    return valores


# --------------------------------------------------------------------------
def _hoja_reporte(hoja, fmt, ventas, margen, ticket, transacciones):
    hoja.hide_gridlines(2)
    hoja.set_column("A:A", 2)
    hoja.set_column("B:K", 14)      # 10 columnas: 5 tarjetas de KPI de 2 columnas

    hoja.set_row(0, 30)
    hoja.set_row(1, 20)
    hoja.set_row(5, 32)
    hoja.set_row(7, 32)
    hoja.set_row(10, 16)
    hoja.set_row(11, 28)

    hoja.merge_range("A1:K1", "  REPORTE EJECUTIVO  |  ANÁLISIS DE EXPANSIÓN COMERCIAL",
                     fmt["titulo"])
    hoja.merge_range("A2:K2",
                     f"  Flujo ETL automatizado con Power Query  ·  Modelo financiero VAN / TIR / PAGO  ·  {AUTOR}",
                     fmt["subtitulo"])

    hoja.write("A4", "Estado:", fmt["estado_lbl"])
    hoja.merge_range("B4:K4",
                     "Presioná «1. ACTUALIZAR DATOS» para ejecutar el flujo ETL.",
                     fmt["estado"])

    # --- botones (Controles de formulario enlazados a macros) -------------
    botones = [
        ("B6", "1. ACTUALIZAR DATOS (ETL)", "Actualizar_Datos"),
        ("E6", "2. APLICAR FORMATO", "Aplicar_Formato_Ejecutivo"),
        ("H6", "3. EXPORTAR A PDF", "Exportar_PDF"),
        ("B8", "4. LIMPIAR FILTROS", "Limpiar_Filtros"),
        ("E8", "5. MOSTRAR/OCULTAR HOJAS", "Alternar_Hojas_Soporte"),
        ("H8", "6. INICIALIZAR MODELO", "Inicializar_Proyecto"),
    ]
    for celda, texto, macro in botones:
        hoja.insert_button(celda, {
            "macro": macro, "caption": texto,
            "width": 150, "height": 30,
            "x_offset": 2, "y_offset": 2,
        })

    # --- tarjetas de KPI --------------------------------------------------
    kpis = [
        ("B", "VENTAS TOTALES", f"={_suma_campo('Monto')}", ventas, "kpi_val",
         "Suma de Monto"),
        ("D", "MARGEN TOTAL", f"={_suma_campo('Margen')}", margen, "kpi_val",
         "Ventas - Costo"),
        ("F", "MARGEN %", f"=IFERROR({_suma_campo('Margen')}/{_suma_campo('Monto')},0)",
         margen / ventas, "kpi_val_pct", "Margen / Ventas"),
        ("H", "TICKET PROMEDIO", f"=IFERROR({_suma_campo('Monto')}/"
         f'MAX(1,COUNTA(INDEX(Datos_Limpios!$A:$Z,0,MATCH("ID_Transaccion",Datos_Limpios!$1:$1,0)))-1),0)',
         ticket, "kpi_val", "Venta media"),
        ("J", "TRANSACCIONES",
         '=IFERROR(MAX(0,COUNTA(INDEX(Datos_Limpios!$A:$Z,0,'
         'MATCH("ID_Transaccion",Datos_Limpios!$1:$1,0)))-1),0)',
         transacciones, "kpi_val_num", "Filas depuradas"),
    ]
    siguiente = {"B": "C", "D": "E", "F": "G", "H": "I", "J": "K"}
    for col, titulo, formula, valor, estilo, pie in kpis:
        par = siguiente[col]
        hoja.merge_range(f"{col}11:{par}11", titulo, fmt["kpi_tit"])
        hoja.merge_range(f"{col}12:{par}12", "", fmt[estilo])
        hoja.write_formula(f"{col}12", formula, fmt[estilo], valor)
        hoja.merge_range(f"{col}13:{par}13", pie, fmt["kpi_pie"])

    hoja.merge_range(
        "A16:K16",
        "  PREGUNTA DE NEGOCIO: ¿qué regiones y categorías concentran el margen? "
        "— usá los segmentadores para filtrar",
        fmt["seccion"])

    hoja.write("B44",
               "La tabla dinámica y los segmentadores se construyen automáticamente al abrir "
               "el libro con las macros habilitadas (botón 6).", fmt["nota"])
    hoja.write("B45",
               "Hojas de soporte (Datos_Crudos, Datos_Limpios, Config) ocultas: botón 5 para "
               "mostrarlas.", fmt["nota"])

    hoja.print_area("A1:K46")
    hoja.fit_to_pages(1, 1)
    hoja.set_landscape()
    hoja.set_paper(9)                     # A4
    hoja.center_horizontally()
    hoja.center_vertically()
    hoja.set_margins(0.3, 0.3, 0.4, 0.4)
    hoja.set_header('&C&"Segoe UI,Bold"&12Reporte Ejecutivo - Análisis de Expansión Comercial')
    hoja.set_footer(f'&L&"Segoe UI"&8{AUTOR}&C&"Segoe UI"&8Página &P de &N&R&"Segoe UI"&8&D')


# --------------------------------------------------------------------------
def _hoja_financiera(hoja, fmt, ventas_base):
    hoja.hide_gridlines(2)
    hoja.set_column("A:A", 2)
    hoja.set_column("B:B", 46)
    hoja.set_column("C:C", 15)
    hoja.set_column("D:H", 15)

    hoja.set_row(0, 28)
    hoja.merge_range("A1:H1", "  EVALUACIÓN FINANCIERA DEL PROYECTO DE EXPANSIÓN",
                     fmt["titulo"])
    hoja.write("B2", f"Modelo de viabilidad a 5 años  ·  {AUTOR}", fmt["nota"])

    # --- 1. Préstamo ------------------------------------------------------
    hoja.merge_range("A4:H4", "  1. ESTRUCTURA DE LA INVERSIÓN Y DEL PRÉSTAMO",
                     fmt["seccion"])

    prestamo = INVERSION_INICIAL * PCT_FINANCIADO
    aporte = INVERSION_INICIAL - prestamo
    tasa_mensual = TNA / 12
    cuotas = PLAZO_ANIOS * 12
    cuota = pago(tasa_mensual, cuotas, prestamo)
    total_pagado = cuota * cuotas
    intereses = total_pagado - prestamo
    servicio_anual = cuota * 12

    filas = [
        (5, "Inversión inicial del proyecto", INVERSION_INICIAL, "entrada"),
        (6, "% financiado con préstamo bancario", PCT_FINANCIADO, "entrada_pct"),
        (7, "Monto del préstamo", "=D5*D6", "moneda", prestamo),
        (8, "Aporte propio (capital del accionista)", "=D5-D7", "moneda", aporte),
        (9, "TNA — tasa nominal anual", TNA, "entrada_pct"),
        (10, "Plazo del préstamo (años)", PLAZO_ANIOS, "entrada_num"),
        (11, "Tasa mensual  (TNA / 12)", "=D9/12", "porcentaje", tasa_mensual),
        (12, "Cantidad de cuotas  (años × 12)", "=D10*12", "entero", cuotas),
    ]
    for fila in filas:
        hoja.write(f"B{fila[0]}", fila[1], fmt["etiqueta"])
        if isinstance(fila[2], str):
            hoja.write_formula(f"D{fila[0]}", fila[2], fmt[fila[3]], fila[4])
        else:
            hoja.write_number(f"D{fila[0]}", fila[2], fmt[fila[3]])

    hoja.write("B14", "CUOTA MENSUAL   =PAGO(tasa; nper; va)", fmt["destacado_lbl"])
    hoja.write_formula("D14", "=-PMT(D11,D12,D7)", fmt["destacado"], cuota)
    hoja.write("B15", "Total a pagar al banco", fmt["etiqueta"])
    hoja.write_formula("D15", "=D14*D12", fmt["moneda"], total_pagado)
    hoja.write("B16", "Intereses totales del financiamiento", fmt["etiqueta"])
    hoja.write_formula("D16", "=D15-D7", fmt["moneda"], intereses)
    hoja.write("B17", "Servicio anual de la deuda  (cuota × 12)", fmt["etiqueta"])
    hoja.write_formula("D17", "=D14*12", fmt["moneda"], servicio_anual)

    # --- 2. Supuestos -----------------------------------------------------
    hoja.merge_range("A19:H19", "  2. SUPUESTOS DE LA PROYECCIÓN", fmt["seccion"])
    hoja.write("B20", "Ventas anuales base (tomadas del flujo ETL)", fmt["etiqueta"])
    col = 'INDEX(Datos_Limpios!$A:$Z,0,MATCH("Monto",Datos_Limpios!$1:$1,0))'
    hoja.write_formula(
        "D20",
        f"=IFERROR(IF(SUM({col})>0,SUM({col}),{ventas_base:.2f}),{ventas_base:.2f})",
        fmt["moneda"], ventas_base)
    hoja.write("B21", "Crecimiento anual esperado de ventas", fmt["etiqueta"])
    hoja.write_number("D21", CRECIMIENTO, fmt["entrada_pct"])
    hoja.write("B22", "Margen operativo (EBITDA / ventas)", fmt["etiqueta"])
    hoja.write_number("D22", MARGEN_EBITDA, fmt["entrada_pct"])
    hoja.write("B23", "Tasa de descuento (costo de capital)", fmt["etiqueta"])
    hoja.write_number("D23", TASA_DESCUENTO, fmt["entrada_pct"])
    hoja.write("B24", "Impuesto a las ganancias", fmt["etiqueta"])
    hoja.write_number("D24", IMPUESTO, fmt["entrada_pct"])

    # --- 3. Flujo de caja -------------------------------------------------
    hoja.merge_range("A26:H26",
                     "  3. FLUJO DE CAJA LIBRE DEL ACCIONISTA — PROYECCIÓN A 5 AÑOS",
                     fmt["seccion"])

    hoja.write("B27", "Concepto", fmt["cab_tabla"])
    for i, titulo in enumerate(["Año 0", "Año 1", "Año 2", "Año 3", "Año 4", "Año 5"]):
        hoja.write(26, 2 + i, titulo, fmt["cab_tabla"])

    columnas = ["C", "D", "E", "F", "G", "H"]
    ingresos, ebitda, impuestos, servicio, flujo = [], [], [], [], []
    for n in range(6):
        ing = 0.0 if n == 0 else ventas_base * (1 + CRECIMIENTO) ** n
        eb = ing * MARGEN_EBITDA
        imp = -eb * IMPUESTO
        srv = 0.0 if n == 0 else -servicio_anual
        ingresos.append(ing)
        ebitda.append(eb)
        impuestos.append(imp)
        servicio.append(srv)
        flujo.append(-aporte if n == 0 else eb + imp + srv)

    hoja.write("B28", "Ingresos proyectados", fmt["concepto"])
    hoja.write("B29", "EBITDA  (ingresos × margen operativo)", fmt["concepto"])
    hoja.write("B30", "Impuesto a las ganancias", fmt["concepto"])
    hoja.write("B31", "Servicio de la deuda (capital + intereses)", fmt["concepto"])
    hoja.write("B32", "FLUJO DE CAJA NETO", fmt["concepto_b"])

    for n, letra in enumerate(columnas):
        if n == 0:
            hoja.write_number(f"{letra}28", 0, fmt["celda_flujo"])
            hoja.write_number(f"{letra}29", 0, fmt["celda_flujo"])
            hoja.write_number(f"{letra}30", 0, fmt["celda_flujo"])
            hoja.write_number(f"{letra}31", 0, fmt["celda_flujo"])
            hoja.write_formula(f"{letra}32", "=-D8", fmt["celda_flujo_b"], flujo[0])
        else:
            hoja.write_formula(f"{letra}28", f"=$D$20*(1+$D$21)^{n}",
                               fmt["celda_flujo"], ingresos[n])
            hoja.write_formula(f"{letra}29", f"={letra}28*$D$22",
                               fmt["celda_flujo"], ebitda[n])
            hoja.write_formula(f"{letra}30", f"=-{letra}29*$D$24",
                               fmt["celda_flujo"], impuestos[n])
            hoja.write_formula(f"{letra}31", "=-$D$17", fmt["celda_flujo"], servicio[n])
            hoja.write_formula(f"{letra}32", f"={letra}29+{letra}30+{letra}31",
                               fmt["celda_flujo_b"], flujo[n])

    # --- 4. Indicadores ---------------------------------------------------
    hoja.merge_range("A34:H34", "  4. INDICADORES DE VIABILIDAD", fmt["seccion"])

    van = vna(TASA_DESCUENTO, flujo[1:]) + flujo[0]
    tasa_interna = tir(flujo)
    va_deuda = va(TASA_DESCUENTO, PLAZO_ANIOS, -servicio_anual)

    hoja.write("B35", "VAN — Valor Actual Neto   =VNA(tasa; flujos) + inversión",
               fmt["destacado_lbl"])
    hoja.write_formula("D35", "=NPV(D23,D32:H32)+C32", fmt["destacado"], van)

    hoja.write("B36", "TIR — Tasa Interna de Retorno   =TIR(flujos)", fmt["destacado_lbl"])
    hoja.write_formula("D36", "=IRR(C32:H32)", fmt["destacado_pct"], tasa_interna)

    hoja.write("B37", "VA del servicio de la deuda   =VA(tasa; nper; pago)", fmt["etiqueta"])
    hoja.write_formula("D37", "=PV(D23,D10,-D17)", fmt["moneda"], va_deuda)

    hoja.write("B38", "Costo de capital de referencia", fmt["etiqueta"])
    hoja.write_formula("D38", "=D23", fmt["porcentaje"], TASA_DESCUENTO)

    hoja.write("B39", "Ratio VAN / aporte propio", fmt["etiqueta"])
    hoja.write_formula("D39", "=D35/D8", fmt["porcentaje"], van / aporte)

    conclusion = (
        '=IF(AND(D35>0,D36>D23),'
        '"CONCLUSIÓN: el proyecto se ACEPTA. El VAN es positivo ($"&TEXT(ROUND(D35,0),"0")&'
        '") y la TIR ("&TEXT(ROUND(D36*100,1),"0.0")&"%) supera el costo de capital ("&'
        'TEXT(ROUND(D23*100,1),"0.0")&"%), por lo que el proyecto genera valor económico.",'
        '"CONCLUSIÓN: el proyecto se RECHAZA. El VAN es $"&TEXT(ROUND(D35,0),"0")&'
        '" y la TIR ("&TEXT(ROUND(D36*100,1),"0.0")&"%) no supera el costo de capital ("&'
        'TEXT(ROUND(D23*100,1),"0.0")&"%).")')
    texto_conclusion = (
        f"CONCLUSIÓN: el proyecto se ACEPTA. El VAN es positivo (${van:,.0f}) y la TIR "
        f"({tasa_interna * 100:.1f}%) supera el costo de capital "
        f"({TASA_DESCUENTO * 100:.1f}%), por lo que el proyecto genera valor económico."
    ).replace(",", ".")

    hoja.set_row(39, 32)
    hoja.merge_range("B40:H40", "", fmt["conclusion"])
    hoja.write_formula("B40", conclusion, fmt["conclusion"], texto_conclusion)

    hoja.print_area("A1:H42")
    hoja.fit_to_pages(1, 1)
    hoja.set_paper(9)
    hoja.center_horizontally()
    hoja.set_margins(0.3, 0.3, 0.4, 0.4)

    return {"cuota": cuota, "van": van, "tir": tasa_interna, "flujo": flujo,
            "prestamo": prestamo, "aporte": aporte, "va_deuda": va_deuda}


# --------------------------------------------------------------------------
def _hoja_datos_limpios(hoja, fmt):
    """Hoja reservada: Power Query escribe aqui su tabla de salida (desde A1)."""
    hoja.set_column("A:P", 14)


# --------------------------------------------------------------------------
def _hoja_datos_crudos(hoja, fmt, crudo):
    hoja.set_column("A:L", 16)
    datos = [[fila[c] for c in ds.COLUMNAS] for fila in crudo]
    hoja.add_table(0, 0, len(datos), len(ds.COLUMNAS) - 1, {
        "name": "tbl_Datos_Crudos",
        "style": "Table Style Light 8",
        "columns": [{"header": c} for c in ds.COLUMNAS],
        "data": datos,
    })


# --------------------------------------------------------------------------
def _hoja_config(hoja, fmt):
    hoja.hide_gridlines(2)
    hoja.set_column("A:A", 2)
    hoja.set_column("B:B", 42)
    hoja.set_column("C:C", 70)

    hoja.set_row(0, 26)
    hoja.merge_range("A1:C1", "  CONFIGURACIÓN Y DOCUMENTACIÓN TÉCNICA", fmt["cfg_tit"])
    hoja.write("B2", "Hoja de soporte — se oculta automáticamente", fmt["nota"])

    hoja.merge_range("A4:C4", "  PARÁMETROS DEL FLUJO ETL", fmt["seccion"])
    hoja.write("B6", "Ruta completa del CSV de origen (opcional)", fmt["cfg_lbl"])
    hoja.write("C6", "", fmt["cfg_val"])
    hoja.write("B7", "Carpeta del libro (la completa la macro)", fmt["cfg_lbl"])
    hoja.write("C7", "", fmt["cfg_val"])
    hoja.write("B8", "Nombre de la consulta de Power Query", fmt["cfg_lbl"])
    hoja.write("C8", "Transacciones", fmt["cfg_val"])
    hoja.write("B9", "Tabla de respaldo dentro del libro", fmt["cfg_lbl"])
    hoja.write("C9", "tbl_Datos_Crudos", fmt["cfg_val"])
    hoja.write("B10", "Origen resuelto en la última ejecución", fmt["cfg_lbl"])
    hoja.write("C10", "(lo completa la macro al actualizar)", fmt["cfg_val"])
    hoja.write("B11",
               "Si C6 está vacío, la macro busca datos\\transacciones_crudas.csv junto al "
               "libro y, si tampoco existe, usa la tabla de respaldo incrustada.",
               fmt["nota"])

    hoja.merge_range("A13:C13", "  PROMPT UTILIZADO CON LA IA (PASO 6)", fmt["seccion"])
    prompt = [
        "Actuá como desarrollador VBA senior para Excel 365 en español.",
        "",
        "Escribí una macro llamada Aplicar_Formato_Ejecutivo que reciba un libro con:",
        "  - una hoja Reporte_Ejecutivo con KPIs en las filas 11 a 13 y una tabla",
        "    dinámica llamada TD_Margen,",
        "  - una hoja Evaluación_Financiera con un modelo de inversión,",
        "  - una hoja Datos_Limpios con la tabla de salida de Power Query.",
        "",
        "Requisitos:",
        "  1. Tipografía corporativa Segoe UI en todas las hojas.",
        "  2. Paleta: azul #0B2A4A para encabezados y celeste #008AB3 para acentos.",
        "  3. Bordes ejecutivos: contorno medio azul e interior fino gris.",
        "  4. Formato moneda $#,##0 en importes y 0.0% en porcentajes.",
        "  5. Todo referenciado por CodeName de hoja, nunca por ActiveSheet.",
        "  6. Manejo de errores con On Error para que no se corte si falta un objeto.",
        "",
        "Devolvé sólo el módulo, con Option Explicit y comentarios en español.",
        "",
        "DEPURACIÓN POSTERIOR (errores que devolvió la IA y cómo se corrigieron):",
        "  a) Usaba ActiveSheet y Select: se reemplazó por referencias por CodeName.",
        "  b) Declaraba Const AZUL = RGB(11,42,74): RGB no es constante en tiempo de",
        "     compilación, se reemplazó por el valor Long ya calculado (4860427).",
        "  c) Fallaba con error 1004 si la tabla dinámica todavía no existía: se agregó",
        "     la comprobación If td Is Nothing Then Exit Sub.",
        "  d) Aplicaba NumberFormat a columnas inexistentes: se encapsuló en el",
        "     procedimiento Formato_Columna con On Error Resume Next.",
    ]
    for i, linea in enumerate(prompt):
        estilo = fmt["mono_com"] if linea.strip().startswith(("DEPURACIÓN", "a)", "b)", "c)", "d)")) else fmt["mono"]
        hoja.write(14 + i, 1, linea, estilo)

    inicio_m = 14 + len(prompt) + 2
    hoja.merge_range(inicio_m - 1, 0, inicio_m - 1, 2,
                     "  CÓDIGO M DE LA CONSULTA (referencia — el original vive en la consulta)",
                     fmt["seccion"])
    with open(os.path.join(RAIZ, "powerquery", "Transacciones.m"), encoding="utf-8") as fh:
        for i, linea in enumerate(fh.read().splitlines()):
            estilo = fmt["mono_com"] if linea.strip().startswith("//") else fmt["mono"]
            hoja.write(inicio_m + i, 1, linea, estilo)


# --------------------------------------------------------------------------
def main() -> None:
    print("1/3  Generando el origen crudo...")
    crudo = ds.generar()
    limpio = ds.limpiar(crudo)
    ruta_csv = escribir_csv(crudo)
    print(f"     {len(crudo)} filas crudas -> {len(limpio)} filas limpias  ({ruta_csv})")

    print("2/3  Compilando el proyecto VBA...")
    ruta_vba = construir_vba()
    print(f"     {os.path.getsize(ruta_vba):,} bytes  ({ruta_vba})")

    print("3/3  Escribiendo el libro .xlsm...")
    valores = construir_libro(ruta_vba, crudo, limpio)
    print(f"     Cuota mensual: ${valores['cuota']:,.2f}")
    print(f"     VAN: ${valores['van']:,.2f}   TIR: {valores['tir']:.2%}")
    print(f"     {os.path.getsize(SALIDA):,} bytes  ({SALIDA})")


if __name__ == "__main__":
    main()
