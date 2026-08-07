"""
Capa de analisis: tablas dinamicas y la hoja `Calculos` que alimenta al tablero.

Como circula un filtro desde un segmentador hasta un grafico:

    segmentador -> tabla dinamica -> GETPIVOTDATA (Calculos) -> grafico

Los graficos no leen la tabla de hechos: leen una grilla de tamano fijo en
`Calculos` resuelta con GETPIVOTDATA contra las tablas dinamicas. Por eso el
layout del tablero no se rompe cuando un filtro deja menos filas.

Todas las tablas dinamicas -las visibles del entregable y las de soporte-
comparten un unico cache, que es lo que permite que un solo segmentador
recalcule todo el libro a la vez.
"""
import datos
from estilos import (GRAY, LIGHT, MONEY, NAVY, NUM, PCT1, PCT2, WHITE, banner,
                     f, imprimible)
from ooxml import Pivot, col_letter

CALC = "Calculos"
AUX = "Pivots_Aux"
TD = "Tablas_Dinamicas"
LIMPIOS = "Datos_Limpios"

# --- filas de la hoja Calculos (una sola fuente de verdad) -------------------
R_TOT = 4         # 4..13   totales por ejercicio
R_KPI = 17        # 17..20  KPIs contra meta ; 21..22 variaciones en p.p.
R_CASC = 28       # 28..33  puente de margen ; 34 control de cierre
R_MES = 38        # 38..49  evolucion mensual, acumulado YTD y variacion MoM
R_CAT = 54        # 54..57  categorias
R_PROD = 62       # 62..71  productos
R_REG = 76        # 76..80  regiones
R_TXT = 85        # 85..94  textos narrativos

CLAVES_TXT = ["titular", "bajada", "g1", "g2", "g3", "g4", "g5",
              "nota1", "nota2", "nota3", "pie5"]
T = {k: f"{CALC}!$B${R_TXT + i}" for i, k in enumerate(CLAVES_TXT)}

MEDIDAS = [("Ventas", "Ventas"), ("Margen_Bruto", "Margen"),
           ("Unidades", "Unidades")]


# =============================================================================
#  Definicion de las tablas dinamicas
# =============================================================================
def definir_pivots():
    """
    Devuelve (pivots, ref).

    `ref` mapea el nombre de cada tabla dinamica de soporte a su celda ancla,
    que es lo que GETPIVOTDATA necesita para apuntarle.
    """
    pivots, ref = [], {}

    # --- las que se entregan a la vista, en la hoja Tablas_Dinamicas -------
    visibles = [
        ("TD_Ventas_por_Region", "Region", "Region", MEDIDAS, 6, 2),
        ("TD_Ventas_por_Categoria", "Categoria", "Categoria", MEDIDAS, 6, 7),
        ("TD_Ventas_por_Producto", "Producto", "Producto", MEDIDAS, 16, 2),
        ("TD_Rentabilidad_por_Segmento", "Segmento_Rentabilidad",
         "Segmento de rentabilidad",
         [("Ventas", "Ventas"), ("Unidades", "Unidades")], 16, 7),
    ]
    for name, campo, caption, medidas, fila, col in visibles:
        pivots.append(Pivot(name, TD, fila, col, campo, data_fields=medidas,
                            row_caption=caption))

    # --- las de soporte: una dimension y una medida, en hoja oculta --------
    soporte = [
        ("PT_Anio_Ventas", "Anio", "Ventas", "Ventas"),
        ("PT_Anio_Margen", "Anio", "Margen_Bruto", "Margen"),
        ("PT_Anio_Unidades", "Anio", "Unidades", "Unidades"),
        ("PT_Anio_CostoProd", "Anio", "Costo_Producto", "CostoProd"),
        ("PT_Anio_CostoLog", "Anio", "Costo_Logistico", "CostoLog"),
        ("PT_Periodo_Ventas", "Periodo", "Ventas", "Ventas"),
        ("PT_Periodo_Margen", "Periodo", "Margen_Bruto", "Margen"),
        ("PT_CatAnio_Ventas", "Cat_Anio", "Ventas", "Ventas"),
        ("PT_CatAnio_Margen", "Cat_Anio", "Margen_Bruto", "Margen"),
        ("PT_ProdAnio_Ventas", "Prod_Anio", "Ventas", "Ventas"),
        ("PT_ProdAnio_Margen", "Prod_Anio", "Margen_Bruto", "Margen"),
        ("PT_Region_Ventas", "Region", "Ventas", "Ventas"),
        ("PT_Region_Margen", "Region", "Margen_Bruto", "Margen"),
    ]
    for i, (name, rowf, medida, cap) in enumerate(soporte):
        p = Pivot(name, AUX, 3, 1 + i * 3, rowf, data_field=medida, caption=cap)
        pivots.append(p)
        ref[name] = f"{AUX}!${col_letter(p.anchor_col)}${p.anchor_row}"
    return pivots, ref


def gpd(ref, campo, item, medida):
    """GETPIVOTDATA con red: si un segmentador deja el item fuera, devuelve 0."""
    it = item if isinstance(item, int) else f'"{item}"'
    return f'=IFERROR(GETPIVOTDATA("{medida}",{ref},"{campo}",{it}),0)'


# =============================================================================
#  Hoja de las tablas dinamicas visibles
# =============================================================================
def hoja_tablas_dinamicas(wb, P, n_limpias):
    from parametros import validacion

    ws = wb.create_sheet(TD)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    for c in "BCDEFGHIJ":
        ws.column_dimensions[c].width = 17
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["G"].width = 24
    ws.column_dimensions["K"].width = 3
    ws.column_dimensions["L"].width = 52

    f(ws, "B2", "Tablas dinamicas del reporte", bold=True, size=15, color=NAVY)
    f(ws, "B3", "Las cuatro resumen la tabla de hechos depurada y comparten el "
                "mismo cache: los segmentadores del tablero las filtran a "
                "todas a la vez. Los graficos del dashboard leen exactamente "
                "estos numeros.", size=9, color=GRAY, italic=True, wrap=True)
    ws.merge_cells("B3:J4")

    f(ws, "B5", "1 · Ventas, margen y unidades por REGION", bold=True, size=9.5,
      color=NAVY)
    f(ws, "G5", "2 · Ventas, margen y unidades por CATEGORIA", bold=True,
      size=9.5, color=NAVY)
    f(ws, "B15", "3 · Ventas, margen y unidades por PRODUCTO", bold=True,
      size=9.5, color=NAVY)
    f(ws, "G15", "4 · Ventas y unidades por SEGMENTO DE RENTABILIDAD "
                 "(columna calculada con SI/Y/O)", bold=True, size=9.5,
      color=NAVY)

    # --- bloque de consulta con listas desplegables -------------------------
    banner(ws, 31, 2, 6, "CONSULTA PUNTUAL  ·  elegi los valores en las listas "
                         "desplegables")
    f(ws, "B32", "Complementa a las tablas dinamicas: responde una pregunta "
                 "concreta sin tocar los filtros del tablero. Las tres celdas "
                 "azules son listas de validacion; el resto son formulas "
                 "SUMAR.SI.CONJUNTO sobre Datos_Limpios.",
      size=9, color=GRAY, italic=True, wrap=True, align="left")
    ws.merge_cells("B32:F33")

    entradas = [("B35", "Region", "=Lista_Region", "AMBA"),
                ("B36", "Categoria", "=Lista_Categoria", "Computacion"),
                ("B37", "Anio", "=Lista_Anio", "Todos")]
    for celda, etiqueta, lista, valor in entradas:
        fila = celda[1:]
        f(ws, f"B{fila}", etiqueta, bold=True, size=9.5, color=NAVY,
          align="left")
        f(ws, f"C{fila}", valor, size=9.5, color="0000FF", bold=True,
          align="center", fill="FFF7D6")
        dv = validacion(lista, f"Elegi {etiqueta}",
                        f"Usa la lista desplegable para elegir {etiqueta}.")
        ws.add_data_validation(dv)
        dv.add(ws[f"C{fila}"])

    # rangos de la tabla limpia que consulta el bloque
    import etl
    fin = n_limpias + 1
    rng = {c: f"{LIMPIOS}!${etl.CL[c]}$2:${etl.CL[c]}${fin}"
           for c in ("Region", "Categoria", "Anio", "Ventas", "Margen_Bruto",
                     "Unidades")}
    # el criterio de anio admite "Todos": en ese caso el filtro se vuelve ">0"
    crit_anio = '=IF($C$37="Todos",">0","="&$C$37)'
    f(ws, "F35", crit_anio, size=8, color=WHITE)   # celda tecnica, no se lee

    def sumif(medida):
        return (f"=SUMIFS({rng[medida]},{rng['Region']},$C$35,"
                f"{rng['Categoria']},$C$36,{rng['Anio']},$F$35)")

    salidas = [("Ventas", sumif("Ventas"), MONEY),
               ("Margen bruto", sumif("Margen_Bruto"), MONEY),
               ("Unidades", sumif("Unidades"), NUM),
               ("Margen bruto %", "=IFERROR(E36/E35,0)", PCT1),
               ("Operaciones",
                f"=COUNTIFS({rng['Region']},$C$35,{rng['Categoria']},$C$36,"
                f"{rng['Anio']},$F$35)", NUM)]
    for i, (etiqueta, formula, fmt) in enumerate(salidas):
        r = 35 + i
        f(ws, f"D{r}", etiqueta, size=9.5, color=GRAY, align="right")
        f(ws, f"E{r}", formula, size=11, bold=True, color=NAVY, fmt=fmt,
          align="center", fill=LIGHT)

    f(ws, "B41", "Nota: la opcion \"Todos\" del anio se resuelve con un "
                 "criterio dinamico (\">0\") en la celda tecnica F35, para que "
                 "una misma formula sirva para un anio puntual y para el total.",
      size=8.5, color=GRAY, italic=True, wrap=True, align="left")
    ws.merge_cells("B41:F42")

    # --- lectura de las tablas ---------------------------------------------
    f(ws, "L2", "Como leer esta hoja", bold=True, size=11, color=NAVY)
    lecturas = [
        "Las cuatro tablas salen del mismo cache de datos, asi que los "
        "cuatro segmentadores del tablero (Trimestre, Categoria, Region y "
        "Canal) las filtran simultaneamente.",
        "La tabla 4 resume una columna que no existe en el origen: "
        "Segmento_Rentabilidad se calcula con SI/Y/O en Datos_Limpios contra "
        "los umbrales de Parametros.",
        "Si los numeros no se mueven al abrir el archivo: Datos > Actualizar "
        "todo. El libro pide el refresco al abrir, pero algunos visores de "
        "terceros no ejecutan tablas dinamicas.",
    ]
    r = 4
    for texto in lecturas:
        f(ws, f"L{r}", texto, size=9, color=GRAY, wrap=True, align="left",
          vertical="top")
        ws.merge_cells(f"L{r}:L{r + 3}")
        r += 5
    imprimible(ws, "B2:J42")
    return ws


def hoja_aux(wb):
    ws = wb.create_sheet(AUX)
    ws.sheet_state = "hidden"
    f(ws, "A1", "Tablas dinamicas de soporte. Los graficos del tablero las "
                "leen con GETPIVOTDATA sobre una grilla de tamano fijo. Se "
                "ocultan porque son infraestructura, no lectura. No editar a "
                "mano.", size=9, color=GRAY)
    return ws


# =============================================================================
#  Hoja Calculos
# =============================================================================
def hoja_calculos(wb, ref, productos_riesgo, P):
    ws = wb.create_sheet(CALC)
    ws.sheet_state = "hidden"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 34
    for c in "BCDEFGHIJ":
        ws.column_dimensions[c].width = 15

    def par(rango):
        return f"Parametros!{rango}"

    f(ws, "A1", "Capa intermedia del tablero", bold=True, size=13, color=NAVY)
    f(ws, "A2", "Cadena del filtro:  segmentador -> tabla dinamica -> "
                "GETPIVOTDATA (esta hoja) -> grafico.",
      size=9, color=GRAY, italic=True)

    # --- totales por ejercicio ---------------------------------------------
    f(ws, "A3", "Totales por ejercicio", bold=True, color=WHITE, fill=NAVY)
    f(ws, "B3", 2024, bold=True, color=WHITE, fill=NAVY, align="center")
    f(ws, "C3", 2025, bold=True, color=WHITE, fill=NAVY, align="center")
    tot = [("Ventas", "PT_Anio_Ventas", "Ventas", MONEY),
           ("Costo de producto", "PT_Anio_CostoProd", "CostoProd", MONEY),
           ("Costo logistico", "PT_Anio_CostoLog", "CostoLog", MONEY),
           ("Margen bruto", "PT_Anio_Margen", "Margen", MONEY),
           ("Unidades", "PT_Anio_Unidades", "Unidades", NUM)]
    for i, (nombre, pt, med, fmt) in enumerate(tot):
        r = R_TOT + i
        f(ws, f"A{r}", nombre)
        for col, anio in (("B", 2024), ("C", 2025)):
            f(ws, f"{col}{r}", gpd(ref[pt], "Anio", anio, med), fmt=fmt)
    v, cp, cl, mg, un = (R_TOT, R_TOT + 1, R_TOT + 2, R_TOT + 3, R_TOT + 4)
    derivadas = [
        ("Margen bruto %", f"=IFERROR({{c}}{mg}/{{c}}{v},0)", PCT1),
        ("Precio promedio por unidad", f"=IFERROR({{c}}{v}/{{c}}{un},0)", MONEY),
        ("Costo de producto por unidad", f"=IFERROR({{c}}{cp}/{{c}}{un},0)",
         MONEY),
        ("Margen unitario", f"={{c}}{R_TOT + 6}-{{c}}{R_TOT + 7}", MONEY),
        ("Costo logistico / ventas", f"=IFERROR({{c}}{cl}/{{c}}{v},0)", PCT2),
    ]
    for i, (nombre, formula, fmt) in enumerate(derivadas):
        r = R_TOT + 5 + i
        f(ws, f"A{r}", nombre)
        for col in ("B", "C"):
            f(ws, f"{col}{r}", formula.format(c=col), fmt=fmt)
    F = {"ventas": v, "costoprod": cp, "costolog": cl, "margen": mg,
         "unid": un, "margen_pct": R_TOT + 5, "precio": R_TOT + 6,
         "costo_unit": R_TOT + 7, "margen_unit": R_TOT + 8,
         "log_pct": R_TOT + 9}

    # --- KPIs contra meta ----------------------------------------------------
    f(ws, f"A{R_KPI - 1}", "KPI", bold=True, color=WHITE, fill=NAVY)
    for col, tit in (("B", "Valor"), ("C", "Meta"), ("D", "Brecha"),
                     ("E", "Cumple")):
        f(ws, f"{col}{R_KPI - 1}", tit, bold=True, color=WHITE, fill=NAVY,
          align="center")
    kpis = [
        ("Ventas 2025", f"=C{F['ventas']}", f"={par(P['meta_ventas'])}",
         MONEY, "mayor"),
        ("Margen bruto %", f"=C{F['margen_pct']}", f"={par(P['meta_margen'])}",
         PCT1, "mayor"),
        ("Crecimiento interanual",
         f"=IFERROR(C{F['ventas']}/B{F['ventas']}-1,0)",
         f"={par(P['meta_yoy'])}", PCT1, "mayor"),
        ("Costo logistico / ventas", f"=C{F['log_pct']}",
         f"={par(P['techo_log'])}", PCT2, "menor"),
    ]
    for i, (nombre, val, meta, fmt, sentido) in enumerate(kpis):
        r = R_KPI + i
        f(ws, f"A{r}", nombre)
        f(ws, f"B{r}", val, fmt=fmt, bold=True, color=NAVY)
        f(ws, f"C{r}", meta, fmt=fmt)
        f(ws, f"D{r}", f"=B{r}-C{r}" if sentido == "mayor" else f"=C{r}-B{r}",
          fmt=fmt)
        f(ws, f"E{r}", f'=IF(D{r}>=0,"EN META","BAJO META")', align="center")
    K = {"ventas": R_KPI, "margen": R_KPI + 1, "yoy": R_KPI + 2,
         "log": R_KPI + 3}

    r = R_KPI + 4
    f(ws, f"A{r}", "Variacion del margen % (p.p.)")
    f(ws, f"B{r}", f"=(C{F['margen_pct']}-B{F['margen_pct']})*100", fmt='0.0')
    f(ws, f"A{r + 1}", "Variacion del costo logistico / ventas (p.p.)")
    f(ws, f"B{r + 1}", f"=(C{F['log_pct']}-B{F['log_pct']})*100", fmt='0.0')
    D = {"margen_pp": r, "log_pp": r + 1}

    # --- puente de margen (cascada) -----------------------------------------
    f(ws, f"A{R_CASC - 1}", "Puente de margen bruto 2024 -> 2025", bold=True,
      color=WHITE, fill=NAVY)
    for col, tit in (("B", "Efecto"), ("C", "Acumulado"), ("D", "Base"),
                     ("E", "Total"), ("F", "Suma"), ("G", "Resta")):
        f(ws, f"{col}{R_CASC - 1}", tit, bold=True, color=WHITE, fill=NAVY,
          align="center")
    pasos = [
        ("Margen 2024", f"=B{F['margen']}"),
        ("Mas unidades", f"=(C{F['unid']}-B{F['unid']})*B{F['margen_unit']}"),
        ("Precios mas altos",
         f"=(C{F['precio']}-B{F['precio']})*C{F['unid']}"),
        ("Costo de compra",
         f"=-(C{F['costo_unit']}-B{F['costo_unit']})*C{F['unid']}"),
        ("Costo logistico", f"=-(C{F['costolog']}-B{F['costolog']})"),
        ("Margen 2025", f"=C{F['margen']}"),
    ]
    for i, (etiqueta, formula) in enumerate(pasos):
        r = R_CASC + i
        f(ws, f"A{r}", etiqueta)
        f(ws, f"B{r}", formula, fmt=MONEY)
        es_total = i in (0, len(pasos) - 1)
        if i == 0 or es_total:
            f(ws, f"C{r}", f"=B{r}", fmt=MONEY)
        else:
            f(ws, f"C{r}", f"=C{r - 1}+B{r}", fmt=MONEY)
        # base invisible + los tres tramos visibles de la cascada
        f(ws, f"D{r}", 0 if es_total else f"=MIN(C{r - 1},C{r})", fmt=MONEY)
        f(ws, f"E{r}", f"=B{r}" if es_total else 0, fmt=MONEY)
        f(ws, f"F{r}", 0 if es_total else f"=MAX(B{r},0)", fmt=MONEY)
        f(ws, f"G{r}", 0 if es_total else f"=-MIN(B{r},0)", fmt=MONEY)
    r = R_CASC + len(pasos)
    f(ws, f"A{r}", "Control: los efectos deben cerrar en cero", size=9,
      color=GRAY, italic=True)
    f(ws, f"B{r}", f"=ROUND(SUM(B{R_CASC}:B{R_CASC + 4})-B{R_CASC + 5},2)",
      fmt=MONEY, size=9)

    # --- evolucion mensual, acumulado YTD y variacion MoM -------------------
    f(ws, f"A{R_MES - 1}", "Evolucion mensual", bold=True, color=WHITE,
      fill=NAVY)
    for col, tit in (("B", "Ventas 2024"), ("C", "Ventas 2025"),
                     ("D", "Margen % 2025"), ("E", "Margen 2025"),
                     ("F", "Margen 2024"), ("G", "YTD 2024"), ("H", "YTD 2025"),
                     ("I", "Var. mensual"), ("J", "Mes")):
        f(ws, f"{col}{R_MES - 1}", tit, bold=True, color=WHITE, fill=NAVY,
          align="center")
    for i, mes in enumerate(datos.MESES):
        r = R_MES + i
        f(ws, f"A{r}", mes)
        f(ws, f"B{r}", gpd(ref["PT_Periodo_Ventas"], "Periodo",
                           f"2024-{i + 1:02d}", "Ventas"), fmt=MONEY)
        f(ws, f"C{r}", gpd(ref["PT_Periodo_Ventas"], "Periodo",
                           f"2025-{i + 1:02d}", "Ventas"), fmt=MONEY)
        f(ws, f"E{r}", gpd(ref["PT_Periodo_Margen"], "Periodo",
                           f"2025-{i + 1:02d}", "Margen"), fmt=MONEY)
        f(ws, f"F{r}", gpd(ref["PT_Periodo_Margen"], "Periodo",
                           f"2024-{i + 1:02d}", "Margen"), fmt=MONEY)
        f(ws, f"D{r}", f"=IFERROR(E{r}/C{r},0)", fmt=PCT1)
        # acumulado del ano: la suma corre desde enero hasta el mes en curso
        f(ws, f"G{r}", f"=SUM($B${R_MES}:B{r})", fmt=MONEY)
        f(ws, f"H{r}", f"=SUM($C${R_MES}:C{r})", fmt=MONEY)
        # variacion contra el mes anterior (MoM); enero no tiene con que
        f(ws, f"I{r}", 0 if i == 0 else f"=IFERROR(C{r}/C{r - 1}-1,0)",
          fmt=PCT1)
        f(ws, f"J{r}", mes)

    # --- categorias ---------------------------------------------------------
    f(ws, f"A{R_CAT - 1}", "Margen % por categoria", bold=True, color=WHITE,
      fill=NAVY)
    for col, tit in (("B", "Margen % 2024"), ("C", "Margen % 2025"),
                     ("D", "Ventas 2025")):
        f(ws, f"{col}{R_CAT - 1}", tit, bold=True, color=WHITE, fill=NAVY,
          align="center")
    for i, cat in enumerate(datos.CATEGORIAS):
        r = R_CAT + i
        f(ws, f"A{r}", cat)
        for col, anio in (("B", 2024), ("C", 2025)):
            m = gpd(ref["PT_CatAnio_Margen"], "Cat_Anio", f"{cat} {anio}",
                    "Margen")[1:]
            vt = gpd(ref["PT_CatAnio_Ventas"], "Cat_Anio", f"{cat} {anio}",
                     "Ventas")[1:]
            f(ws, f"{col}{r}", f"=IFERROR(({m})/({vt}),0)", fmt=PCT1)
        f(ws, f"D{r}", gpd(ref["PT_CatAnio_Ventas"], "Cat_Anio",
                           f"{cat} 2025", "Ventas"), fmt=MONEY)

    # --- productos (dispersion) --------------------------------------------
    f(ws, f"A{R_PROD - 1}", "Productos 2025", bold=True, color=WHITE, fill=NAVY)
    for col, tit in (("B", "Ventas 2025"), ("C", "Crecimiento"),
                     ("D", "Margen %"), ("E", "X en meta"), ("F", "Y en meta"),
                     ("G", "X riesgo"), ("H", "Y riesgo")):
        f(ws, f"{col}{R_PROD - 1}", tit, bold=True, color=WHITE, fill=NAVY,
          align="center")
    for i, (prod, _, _, _) in enumerate(datos.PRODUCTOS):
        r = R_PROD + i
        v25 = gpd(ref["PT_ProdAnio_Ventas"], "Prod_Anio", f"{prod} 2025",
                  "Ventas")[1:]
        v24 = gpd(ref["PT_ProdAnio_Ventas"], "Prod_Anio", f"{prod} 2024",
                  "Ventas")[1:]
        m25 = gpd(ref["PT_ProdAnio_Margen"], "Prod_Anio", f"{prod} 2025",
                  "Margen")[1:]
        f(ws, f"A{r}", prod)
        f(ws, f"B{r}", f"={v25}", fmt=MONEY)
        f(ws, f"C{r}", f"=IFERROR(({v25})/({v24})-1,0)", fmt=PCT1)
        f(ws, f"D{r}", f"=IFERROR(({m25})/({v25}),0)", fmt=PCT1)
        # series separadas: las celdas vacias son las que el grafico saltea
        col_x, col_y = ("G", "H") if prod in productos_riesgo else ("E", "F")
        f(ws, f"{col_x}{r}", f"=C{r}", fmt=PCT1)
        f(ws, f"{col_y}{r}", f"=D{r}", fmt=PCT1)

    # --- regiones -----------------------------------------------------------
    f(ws, f"A{R_REG - 1}", "Regiones", bold=True, color=WHITE, fill=NAVY)
    for col, tit in (("B", "Ventas"), ("C", "Margen"), ("D", "Margen %")):
        f(ws, f"{col}{R_REG - 1}", tit, bold=True, color=WHITE, fill=NAVY,
          align="center")
    for i, region in enumerate(datos.REGIONES):
        r = R_REG + i
        f(ws, f"A{r}", region)
        f(ws, f"B{r}", gpd(ref["PT_Region_Ventas"], "Region", region, "Ventas"),
          fmt=MONEY)
        f(ws, f"C{r}", gpd(ref["PT_Region_Margen"], "Region", region, "Margen"),
          fmt=MONEY)
        f(ws, f"D{r}", f"=IFERROR(C{r}/B{r},0)", fmt=PCT1)
    n_reg = len(datos.REGIONES)
    r = R_REG + n_reg
    f(ws, f"A{r}", "Region lider por ventas", size=9, color=GRAY)
    f(ws, f"B{r}", f"=INDEX($A${R_REG}:$A${R_REG + n_reg - 1},"
                   f"MATCH(MAX($B${R_REG}:$B${R_REG + n_reg - 1}),"
                   f"$B${R_REG}:$B${R_REG + n_reg - 1},0))", size=9)
    f(ws, f"A{r + 1}", "Participacion de la region lider", size=9, color=GRAY)
    f(ws, f"B{r + 1}", f"=IFERROR(MAX($B${R_REG}:$B${R_REG + n_reg - 1})/"
                       f"SUM($B${R_REG}:$B${R_REG + n_reg - 1}),0)", fmt=PCT1,
      size=9)
    G = {"region_lider": r, "part_lider": r + 1}

    # --- narrativa dinamica -------------------------------------------------
    f(ws, f"A{R_TXT - 1}", "Textos que el tablero muestra", bold=True,
      color=WHITE, fill=NAVY)
    yoy, mgp = f"B{K['yoy']}", f"B{D['margen_pp']}"
    mg25, mg24 = f"C{F['margen_pct']}", f"B{F['margen_pct']}"
    # semestres: H1 = primeras 6 filas del bloque mensual, H2 = las otras 6
    h1_25 = (f'IFERROR(SUM(E{R_MES}:E{R_MES + 5})/'
             f'SUM(C{R_MES}:C{R_MES + 5}),0)')
    h2_25 = (f'IFERROR(SUM(E{R_MES + 6}:E{R_MES + 11})/'
             f'SUM(C{R_MES + 6}:C{R_MES + 11}),0)')
    h2_24 = (f'IFERROR(SUM(F{R_MES + 6}:F{R_MES + 11})/'
             f'SUM(B{R_MES + 6}:B{R_MES + 11}),0)')
    # tramos del puente
    e_vol, e_pre = f"B{R_CASC + 1}", f"B{R_CASC + 2}"
    e_cpr, e_log = f"B{R_CASC + 3}", f"B{R_CASC + 4}"
    peso_log = f"IFERROR(ABS({e_log})/(ABS({e_cpr})+ABS({e_log})),0)"
    ytd_gap = (f'IFERROR(H{R_MES + 11}/G{R_MES + 11}-1,0)')
    mom_max = f"MAX(I{R_MES + 1}:I{R_MES + 11})"

    textos = [
        ("titular", f'="Las ventas crecen "&TEXT({yoy},"0.0%")&'
                    f'" y el margen cae "&TEXT(ABS({mgp}),"0.0")&'
                    f'" p.p.: el costo de compra y la logistica del e-commerce'
                    f' se reparten el deterioro"'),
        ("bajada", f'="El margen bruto pasa de "&TEXT({mg24},"0.0%")&" a "&'
                   f'TEXT({mg25},"0.0%")&". La logistica explica "&'
                   f'TEXT({peso_log},"0%")&" del mayor costo y el resto es '
                   f'precio de compra: son dos causas, no una."'),
        ("g1", f'="El margen se recupera en el segundo semestre ("&'
               f'TEXT({h2_25},"0.0%")&") pero no vuelve al nivel de 2024 ("&'
               f'TEXT({h2_24},"0.0%")&")"'),
        ("g2", f'="El acumulado del ano cierra "&TEXT({ytd_gap},"0.0%")&'
               f'" arriba de 2024, y la brecha se abre desde el primer '
               f'trimestre"'),
        ("g3", f'="Volumen y precio sumaron "&'
               f'TEXT({e_vol}+{e_pre},"$#,##0")&" de margen; los costos se '
               f'llevaron "&TEXT(ABS({e_cpr})+ABS({e_log}),"$#,##0")'),
        ("g4", f'="Computacion concentra la caida: pierde "&'
               f'TEXT(ABS((C{R_CAT}-B{R_CAT})*100),"0.0")&" p.p. de margen"'),
        ("g5", '="Los productos que mas crecen son los que menos margen dejan"'),
        ("nota1", f'="EVIDENCIA: las ventas suben "&TEXT({yoy},"0.0%")&'
                  f'" y las unidades "&'
                  f'TEXT(IFERROR(C{F["unid"]}/B{F["unid"]}-1,0),"0.0%")&'
                  f'".  INSIGHT: el crecimiento es de volumen, no de precio, '
                  f'asi que la demanda es real y conviene sostener el plan '
                  f'comercial."'),
        ("nota2", f'="EVIDENCIA: el costo de compra resta "&'
                  f'TEXT(ABS({e_cpr}),"$#,##0")&" y la logistica "&'
                  f'TEXT(ABS({e_log}),"$#,##0")&" de margen.  INSIGHT: la '
                  f'caida tiene dos causas, no una; atribuirla solo al '
                  f'e-commerce sobreestima su peso y desviaria el presupuesto '
                  f'de correccion."'),
        ("nota3", f'="EVIDENCIA: el segundo semestre ("&TEXT({h2_25},"0.0%")&'
                  f'") supera al primero ("&TEXT({h1_25},"0.0%")&") pero '
                  f'sigue por debajo del mismo periodo de 2024 ("&'
                  f'TEXT({h2_24},"0.0%")&").  INSIGHT: la correccion del Q3 '
                  f'funciona pero es parcial; hay que sostenerla en el '
                  f'proximo semestre antes de fijar nuevas metas."'),
        # el pie del grafico de dispersion se arma en dos celdas: la lista de
        # productos en C y el texto que se muestra en B
        ("pie5", f'=IF(LEN(C{R_TXT + 10})=0,'
                 f'"Todos los productos quedan por encima del corte de '
                 f'margen.","Zona de riesgo (margen por debajo de "&'
                 f'TEXT({par(P["corte_prod"])},"0%")&"): "&'
                 f'LEFT(C{R_TXT + 10},LEN(C{R_TXT + 10})-2)&'
                 f'".  El resto de los productos esta en meta.")'),
    ]
    assert [c for c, _ in textos] == CLAVES_TXT, "orden de textos desalineado"
    for i, (clave, formula) in enumerate(textos):
        r = R_TXT + i
        f(ws, f"A{r}", clave, size=9, color=GRAY)
        f(ws, f"B{r}", formula, size=9)

    # lista de productos por debajo del corte: un termino por producto, para
    # que la lista se rehaga sola si cambia el corte o si un filtro mueve los
    # numeros. Sin TEXTJOIN, que no existe en todas las versiones.
    n_prod = len(datos.PRODUCTOS)
    concat = "&".join(
        f'IF(D{R_PROD + k}<{par(P["corte_prod"])},A{R_PROD + k}&", ","")'
        for k in range(n_prod))
    f(ws, f"C{R_TXT + 10}", f"={concat}", size=9)
    return F, K, D, G
