"""
Las medidas DAX del modelo.

Igual que con el codigo M, viven en un solo lugar y desde aca se escriben al
archivo `dax/medidas.dax` y a la hoja `Medidas_DAX` del libro.

La idea del Paso 4 es que el modelo no se ponga pesado. Dos decisiones lo
sostienen:

  * las metricas de fila (venta bruta, venta neta, costo, margen) ya vienen
    calculadas desde Power Query, asi que las medidas base son SUM() sobre
    una columna comprimida y no SUMX() iterando la tabla de hechos;
  * cuando una medida necesita el mismo subtotal dos veces, se guarda en una
    variable con VAR en vez de repetir la expresion.

`COMPARACION` guarda el antes y el despues de la misma metrica: sirve para
mostrar por que VAR/RETURN no es solo cosmetica.
"""

MEDIDAS = [
    (
        "Total Ventas",
        "Total Ventas := SUM ( Hechos_Ventas[Venta_Neta] )",
        "Agregacion simple. La columna Venta_Neta ya viene resuelta desde "
        "Power Query, asi que el motor solo suma una columna comprimida: es la "
        "operacion mas barata que existe en VertiPaq.",
    ),
    (
        "Costo Total",
        "Costo Total := SUM ( Hechos_Ventas[Costo_Total] )",
        "Misma logica que Total Ventas.",
    ),
    (
        "Costo Logistico",
        "Costo Logistico := SUM ( Hechos_Ventas[Costo_Envio] )",
        "Se aisla el envio porque es la variable que mas mueve el margen entre "
        "canales, y es una de las celdas cambiantes de la simulacion.",
    ),
    (
        "Margen Bruto",
        "Margen Bruto := SUM ( Hechos_Ventas[Margen_Bruto] )",
        "Suma directa de la columna calculada en M. La alternativa "
        "SUMX(Hechos_Ventas, ...) daria el mismo numero iterando 5.200 filas "
        "en cada celda de cada tabla dinamica.",
    ),
    (
        "% Margen Bruto",
        "% Margen Bruto :=\n"
        "VAR VentasNetas = [Total Ventas]\n"
        "VAR MargenBruto = [Margen Bruto]\n"
        "RETURN\n"
        "    DIVIDE ( MargenBruto, VentasNetas )",
        "VAR/RETURN. Cada subtotal se resuelve una sola vez y se reutiliza. "
        "DIVIDE ademas evita el error de division por cero sin tener que "
        "envolver todo en un IF.",
    ),
    (
        "Ticket Promedio",
        "Ticket Promedio :=\n"
        "VAR VentasNetas = [Total Ventas]\n"
        "VAR CantidadOperaciones = DISTINCTCOUNT ( Hechos_Ventas[ID_Venta] )\n"
        "RETURN\n"
        "    DIVIDE ( VentasNetas, CantidadOperaciones )",
        "VAR/RETURN. Separar el numerador del denominador deja explicito que "
        "el ticket se mide por operacion y no por unidad.",
    ),
    (
        "Crecimiento Interanual %",
        "Crecimiento Interanual % :=\n"
        "VAR VentasPeriodoActual = [Total Ventas]\n"
        "VAR VentasPeriodoAnterior =\n"
        "    CALCULATE ( [Total Ventas], SAMEPERIODLASTYEAR ( Dim_Calendario[Fecha] ) )\n"
        "VAR Variacion = VentasPeriodoActual - VentasPeriodoAnterior\n"
        "RETURN\n"
        "    IF (\n"
        "        NOT ISBLANK ( VentasPeriodoAnterior ) && VentasPeriodoAnterior <> 0,\n"
        "        DIVIDE ( Variacion, VentasPeriodoAnterior )\n"
        "    )",
        "VAR/RETURN con inteligencia de tiempo. El subtotal del ano anterior "
        "es el calculo caro de la medida: guardarlo en una variable lo resuelve "
        "una vez y despues se usa tres veces (en la resta, en el ISBLANK y en "
        "el divisor). Necesita que Dim_Calendario este marcada como Tabla de "
        "fechas en Power Pivot.",
    ),
    (
        "% Costo Logistico s/Ventas",
        "% Costo Logistico s/Ventas :=\n"
        "VAR VentasNetas = [Total Ventas]\n"
        "VAR CostoEnvios = [Costo Logistico]\n"
        "RETURN\n"
        "    DIVIDE ( CostoEnvios, VentasNetas )",
        "Es el KPI que conecta el modelo con la hoja Simulacion: la tasa "
        "logistica que mueven los escenarios es exactamente esta medida.",
    ),
]

# El mismo indicador escrito de dos maneras. Es el ejemplo del Paso 4.
COMPARACION = {
    "lenta": (
        "% Margen Bruto (version lenta - NO USAR) :=\n"
        "DIVIDE (\n"
        "    SUMX (\n"
        "        FILTER ( Hechos_Ventas, Hechos_Ventas[Unidades] > 0 ),\n"
        "        Hechos_Ventas[Unidades] * Hechos_Ventas[Precio_Unitario]\n"
        "            * ( 1 - Hechos_Ventas[Descuento] )\n"
        "            - Hechos_Ventas[Unidades] * Hechos_Ventas[Costo_Unitario]\n"
        "            - Hechos_Ventas[Costo_Envio]\n"
        "    ),\n"
        "    SUMX (\n"
        "        FILTER ( Hechos_Ventas, Hechos_Ventas[Unidades] > 0 ),\n"
        "        Hechos_Ventas[Unidades] * Hechos_Ventas[Precio_Unitario]\n"
        "            * ( 1 - Hechos_Ventas[Descuento] )\n"
        "    )\n"
        ")"
    ),
    "rapida": (
        "% Margen Bruto :=\n"
        "VAR VentasNetas = [Total Ventas]\n"
        "VAR MargenBruto = [Margen Bruto]\n"
        "RETURN\n"
        "    DIVIDE ( MargenBruto, VentasNetas )"
    ),
    "problemas": [
        "FILTER ( Hechos_Ventas, ... ) materializa la tabla de hechos entera "
        "en memoria antes de filtrar; el filtro se puede resolver en la "
        "columna, sin construir la tabla.",
        "El mismo SUMX aparece dos veces: el motor lo evalua dos veces por "
        "cada celda del reporte.",
        "La aritmetica de fila se repite en cada consulta en vez de estar "
        "resuelta una sola vez en el ETL.",
        "Al no reutilizar medidas, cualquier cambio de definicion hay que "
        "replicarlo a mano en cada formula.",
    ],
    "resultado": (
        "Las dos versiones devuelven el mismo numero. La segunda lo hace con "
        "dos escaneos de columna en lugar de dos materializaciones completas "
        "de la tabla de hechos, y sin repetir una sola operacion aritmetica."
    ),
}

ENCABEZADO = """// ===========================================================================
//  TechStore - Medidas DAX del modelo en estrella
//  Modulo 8 - Arquitectura del dato, optimizacion y simulacion de escenarios
//
//  Como cargarlas: Power Pivot > Administrar > pestana Hechos_Ventas >
//  area de calculo > pegar una medida por celda.
//
//  Criterio de optimizacion:
//    1. la aritmetica de fila se resuelve en Power Query, no en DAX;
//    2. las medidas base son SUM() sobre columnas, no SUMX() sobre tablas;
//    3. todo subtotal que se usa mas de una vez va en un VAR.
// ===========================================================================

"""


def texto():
    partes = [ENCABEZADO]
    for nombre, codigo, explicacion in MEDIDAS:
        partes.append(f"// --- {nombre} ---\n")
        for linea in _envolver(explicacion, 74):
            partes.append(f"// {linea}\n")
        partes.append(codigo + "\n\n")

    partes.append("\n// ===========================================================================\n")
    partes.append("//  Antes y despues: la misma metrica sin optimizar y optimizada\n")
    partes.append("// ===========================================================================\n\n")
    partes.append(COMPARACION["lenta"] + "\n\n")
    partes.append("// Problemas de la version de arriba:\n")
    for p in COMPARACION["problemas"]:
        for i, linea in enumerate(_envolver(p, 72)):
            partes.append(f"//   {'- ' if i == 0 else '  '}{linea}\n")
    partes.append("\n")
    partes.append(COMPARACION["rapida"] + "\n")
    return "".join(partes)


def _envolver(texto_largo, ancho):
    palabras, linea, salida = texto_largo.split(), "", []
    for palabra in palabras:
        if len(linea) + len(palabra) + 1 > ancho:
            salida.append(linea)
            linea = palabra
        else:
            linea = f"{linea} {palabra}".strip()
    if linea:
        salida.append(linea)
    return salida


def escribir(ruta):
    with open(ruta, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(texto())
    return ruta


if __name__ == "__main__":
    print(texto())
