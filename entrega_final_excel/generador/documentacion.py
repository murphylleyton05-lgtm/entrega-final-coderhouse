"""
Hoja `Documentacion`: como esta hecho el libro y como se opera.

Un analisis reproducible es el que otra persona puede entender, revisar y
actualizar sin rehacerlo. Esta hoja es esa documentacion minima: el mapa de
hojas, el detalle del ETL, las decisiones de diseno y sus porques.
"""
from estilos import (GRAY, INK, LIGHT, NAVY, WHITE, banner,
                     encabezado_tabla, f, imprimible)

HOJA = "Documentacion"


def construir(wb):
    ws = wb.create_sheet(HOJA)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 30
    ws.column_dimensions["D"].width = 74
    ws.column_dimensions["E"].width = 2

    f(ws, "B2", "Documentacion del libro", bold=True, size=15, color=NAVY)
    f(ws, "B3", "Practica Integradora Final · Coderhouse, Data Analyst · "
                "Lleyton Murphy", size=9.5, color=GRAY, italic=True)

    r = 5

    def bloque(titulo, filas, ancho_b=None):
        nonlocal r
        banner(ws, r, 2, 4, titulo)
        r += 1
        for a, b, c in filas:
            f(ws, f"B{r}", a, bold=True, size=9.5, color=NAVY, align="left",
              vertical="top")
            f(ws, f"C{r}", b, size=9, color=INK, align="left", wrap=True,
              vertical="top")
            f(ws, f"D{r}", c, size=9, color=INK, align="left", wrap=True,
              vertical="top")
            ws.row_dimensions[r].height = max(
                26, 12.5 * (1 + max(len(b) // 34, len(c) // 88)))
            r += 1
        r += 1

    # --- mapa de hojas ------------------------------------------------------
    bloque("MAPA DEL LIBRO", [
        ("Hoja", "Rol", "Que contiene"),
        ("Dashboard", "El entregable",
         "Titular, 4 tarjetas de KPI, 4 segmentadores, 5 graficos y el "
         "reporte ejecutivo. Se abre en esta hoja."),
        ("Tablas_Dinamicas", "Analisis",
         "Las 4 tablas dinamicas del reporte mas un bloque de consulta con "
         "listas desplegables."),
        ("Datos_Limpios", "Tabla de hechos",
         "2.400 operaciones depuradas. Cada celda es una formula que "
         "reconstruye el valor desde el export crudo."),
        ("Datos_Origen", "Entrada",
         "El export del sistema transaccional, sin tocar. Los defectos estan "
         "resaltados con formato condicional."),
        ("Control_Calidad", "Auditoria",
         "Cuenta los defectos del origen y verifica con formulas que despues "
         "del ETL no quede ninguno."),
        ("Parametros", "Constantes",
         "Metas, umbrales, tasas, diccionarios de normalizacion y listas de "
         "validacion. Ningun numero fijo vive fuera de aca."),
        ("Documentacion", "Esta hoja", "Como esta hecho y como se opera."),
        ("Calculos (oculta)", "Soporte",
         "Grilla de tamano fijo resuelta con GETPIVOTDATA. Es lo que leen los "
         "graficos."),
        ("Pivots_Aux (oculta)", "Soporte",
         "Tablas dinamicas de una sola medida que alimentan a GETPIVOTDATA."),
    ])

    # --- el ETL paso a paso -------------------------------------------------
    bloque("EL ETL, PASO A PASO  ·  de Datos_Origen a Datos_Limpios", [
        ("Paso", "Funciones usadas", "Que hace y por que"),
        ("1. Quitar duplicados", "COINCIDIR / MATCH",
         "El sistema reenvia lotes y repite filas. Se conserva la primera "
         "aparicion de cada ID_Venta. La columna Pos_Origen guarda a que fila "
         "del export apunta cada operacion y la columna Control_ETL verifica, "
         "fila por fila, que esa sea de verdad la primera aparicion."),
        ("2. Normalizar fechas", "SI, ESNUMERO, EXTRAE, FECHA, VALOR",
         "El export mezcla tres formatos: fecha real, texto dd/mm/aaaa y "
         "texto aaaa-mm-dd. La formula detecta cual es cada uno y reconstruye "
         "una fecha real. Sin esto no hay agrupacion por mes ni por trimestre "
         "que funcione."),
        ("3. Estandarizar textos", "ESPACIOS, INDICE, COINCIDIR",
         "Region, Canal y Categoria llegan con mayusculas, minusculas, "
         "acentos y espacios mezclados. Se resuelven contra el diccionario de "
         "la hoja Parametros. Agregar una variante nueva es agregar una fila "
         "al diccionario: no hay que tocar ninguna formula."),
        ("4. Corregir tipos", "SI, ESTEXTO, VALOR",
         "Algunas unidades llegan como texto. Se convierten a numero solo "
         "cuando hace falta, para no romper las que ya venian bien."),
        ("5. Completar vacios", "SI, ESPACIOS, INDICE, COINCIDIR",
         "Las ciudades vacias se completan con la ciudad de referencia de la "
         "region. Se documenta como imputacion, no se presenta como dato "
         "original."),
        ("6. Derivar columnas", "ANIO, MES, REDONDEAR.MAS, TEXTO",
         "Anio, Trimestre, Mes, Periodo, Cat_Anio y Prod_Anio salen de la "
         "fecha y de los textos ya limpios. Son las que hacen posible el "
         "analisis temporal y la comparacion interanual."),
        ("7. Calcular medidas", "REDONDEAR, SUMAR.SI.CONJUNTO",
         "Ventas, Costo_Producto, Costo_Logistico y Margen_Bruto se calculan "
         "en la planilla, no se pegan como numeros. La tasa logistica no se "
         "repite fila a fila: se busca en la tabla de Parametros segun canal, "
         "anio y trimestre."),
        ("8. Clasificar", "SI, Y, O",
         "Tres columnas de logica condicional: Segmento_Rentabilidad, "
         "Alerta_Logistica y Ticket. Son las que convierten un numero en una "
         "decision, y son ademas dimensiones de analisis en las tablas "
         "dinamicas."),
    ])

    # --- columnas logicas ---------------------------------------------------
    bloque("LAS COLUMNAS CALCULADAS CON LOGICA CONDICIONAL", [
        ("Columna", "Regla", "Para que sirve"),
        ("Segmento_Rentabilidad",
         'SI(Y(margen% >= meta; ventas >= ticket alto); "Alta rentabilidad"; '
         'SI(O(margen% < piso; ventas < ticket bajo); "Revisar"; "Estandar"))',
         "Separa las operaciones que sostienen el negocio de las que hay que "
         "revisar. Usa Y para exigir las dos condiciones a la vez y O para "
         "que alcance con una sola para marcar el problema. Es la dimension "
         "de la cuarta tabla dinamica."),
        ("Alerta_Logistica",
         'SI(Y(canal = "E-commerce"; costo logistico / ventas > techo); '
         '"Alerta"; "OK")',
         "Marca las ventas online cuyo costo de envio se sale del umbral "
         "operativo. Es el detalle fila a fila de lo que el KPI de costo "
         "logistico muestra agregado."),
        ("Ticket",
         'SI(ventas >= ticket alto; "Alto"; SI(ventas < ticket bajo; "Bajo"; '
         '"Medio"))',
         "Clasifica el tamano de cada operacion. Permite mirar el mix de "
         "tickets sin recalcular nada."),
        ("Donde estan los umbrales", "Parametros!C11:C14",
         "Los cuatro cortes viven en la hoja Parametros. Cambiar un umbral "
         "reclasifica las 2.400 operaciones y actualiza las tablas dinamicas "
         "sin tocar una sola formula."),
    ])

    # --- como circula un filtro ---------------------------------------------
    bloque("COMO VIAJA UN FILTRO HASTA UN GRAFICO", [
        ("La cadena", "segmentador -> tabla dinamica -> GETPIVOTDATA "
                      "(Calculos) -> grafico",
         "Los graficos no leen la tabla de hechos: leen una grilla de tamano "
         "fijo en la hoja Calculos, resuelta con GETPIVOTDATA contra las "
         "tablas dinamicas. Por eso el layout no se rompe cuando un filtro "
         "deja menos filas, y por eso los numeros del tablero y los de la "
         "hoja Tablas_Dinamicas son siempre los mismos."),
        ("Un unico cache", "4 segmentadores, 17 tablas dinamicas",
         "Todas las tablas dinamicas del libro comparten un solo cache de "
         "datos. Eso es lo que hace que una sola seleccion recalcule a la vez "
         "las 4 tarjetas, los 5 graficos y las 4 tablas dinamicas visibles."),
        ("Por que Trimestre y no Anio",
         "El filtro de tiempo es el trimestre",
         "Un segmentador de anio dejaria sin sentido el KPI de crecimiento "
         "interanual: al elegir 2025 no habria 2024 contra que comparar. Con "
         "el trimestre, elegir Q1 compara Q1-2024 contra Q1-2025 y la "
         "comparacion se mantiene valida siempre."),
        ("Si los numeros no se mueven", "Datos > Actualizar todo",
         "El libro pide el refresco al abrir. Algunos visores de terceros "
         "(Google Sheets, LibreOffice, la vista previa de GitHub) no ejecutan "
         "tablas dinamicas ni dibujan segmentadores: el archivo esta pensado "
         "para Excel 2010 o posterior."),
    ])

    # --- criterios de diseno ------------------------------------------------
    bloque("CRITERIOS DE DISENO DEL TABLERO", [
        ("Decision", "Que se hizo", "Por que"),
        ("Paleta", "Tres colores y nada mas",
         "Azul institucional (ano corriente y valor principal), gris "
         "(contexto y ano anterior) y ambar (alerta y efectos que restan). El "
         "resto son neutros de fondo y de texto. Un color que no significa "
         "nada es ruido."),
        ("Interfaz", "Sin cuadricula ni encabezados de fila y columna",
         "El tablero tiene que leerse como un reporte, no como una planilla. "
         "Las hojas de soporte estan ocultas por la misma razon."),
        ("Tarjetas de KPI", "Cuatro, con meta y formato condicional",
         "Cada tarjeta muestra el valor, la meta y la brecha. La que no llega "
         "a la meta pasa entera a ambar: la alerta se ve antes de leer el "
         "numero."),
        ("Orden de lectura", "Que paso -> por que -> donde actuar",
         "Los graficos estan agrupados en tres secciones tituladas, en el "
         "orden en que un gerente hace las preguntas."),
        ("Titulos de grafico", "Son formulas, no texto fijo",
         "Cada titulo apunta a una celda que se reescribe con los datos, asi "
         "que dice el hallazgo y no solo el nombre del grafico."),
        ("Reporte ejecutivo", "Evidencia primero, insight despues",
         "Las tres notas del pie separan el dato cuantitativo de su "
         "significado estrategico, que es lo que permite decidir."),
        ("Vocabulario", "Metricas de negocio, no comandos",
         "En el tablero no aparece ningun tecnicismo de software: ni nombres "
         "de funciones, ni errores de actualizacion, ni instrucciones de "
         "herramienta."),
    ])

    # --- alcance y limites --------------------------------------------------
    bloque("ALCANCE, SUPUESTOS Y LIMITES", [
        ("Tema", "Detalle", "Aclaracion"),
        ("Origen de los datos", "Dataset simulado y determinista",
         "Extiende el modelo Ventas_Tech_DB del checkpoint de SQL -mismas "
         "categorias, productos y ciudades- a 24 meses, 5 regiones y 2 "
         "canales. Semilla fija: el mismo dataset se reproduce siempre igual. "
         "Todos los supuestos estan en la hoja Parametros."),
        ("Ciudades imputadas", "Se completan por region",
         "Las ciudades que el export deja vacias se rellenan con la ciudad de "
         "referencia de la region. Es una imputacion documentada, no un dato "
         "original: la hoja Control_Calidad informa cuantas fueron."),
        ("El ETL, en formulas", "Sin dependencias externas",
         "La limpieza esta escrita con formulas nativas de Excel a la vista, "
         "en lugar de quedar encerrada en un paso que no se puede auditar "
         "celda por celda. Se reejecuta sola al recalcular el libro y no "
         "depende de ningun conector."),
        ("Formato del archivo", ".xlsx, sin macros",
         "El entregable se envia en .xlsx, que es un formato sin macros. Toda "
         "la interactividad del tablero -segmentadores, formato condicional, "
         "refresco al abrir, listas de validacion- es nativa y no necesita "
         "que el usuario habilite nada."),
        ("Como reproducirlo", "generador/pipeline.py",
         "El libro se arma de cero con scripts versionados en el repositorio. "
         "El entregable es el .xlsx; los scripts existen para que el "
         "resultado sea auditable."),
    ])
    imprimible(ws, f"B2:D{r}")
    return ws
