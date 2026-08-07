"""
Dataset de la Practica Integradora Final - cadena TechStore.

Este modulo produce DOS cosas:

  1. `generar_limpio()`  -> la tabla de hechos correcta (2.400 operaciones,
     2024-2025). Es la continuidad del modelo `Ventas_Tech_DB` del checkpoint
     de SQL y del tablero del Modulo 9.

  2. `ensuciar(filas)`   -> el export crudo tal como sale del sistema
     transaccional, con los problemas de calidad que pide la consigna:
     duplicados, fechas en formatos mixtos, categorias sin estandarizar,
     numeros cargados como texto y campos vacios.

El libro entregable arranca de (2) y reconstruye (1) con formulas visibles en
la hoja `Datos_Limpios`. Que las dos versiones coincidan se verifica en el
paso de auditoria del pipeline: si el ETL de la planilla no reproduce
exactamente `generar_limpio()`, el build falla.

La historia que codifica el dataset (y que el tablero debe hacer evidente):
  * las ventas crecen con fuerza en 2025,
  * pero el margen % cae, porque el mix se corre a e-commerce, que tiene un
    costo logistico por peso vendido mucho mas alto,
  * y a partir del Q3-2025 el margen se recupera cuando baja la tasa logistica
    de e-commerce (umbral de envio minimo).
"""
import random
from datetime import date
from decimal import ROUND_HALF_UP, Decimal


def r2(x):
    """
    Redondea a dos decimales como REDONDEAR/ROUND de Excel, no como Python.

    Dos diferencias, y las dos importan:

      1. Python redondea al par (round(0.125,2) -> 0.12); Excel se aleja
         siempre del cero (0.13).
      2. Excel primero recorta el numero a 15 digitos significativos y recien
         despues redondea. Sin ese paso, 4165*0.013 vale 54.144999999999996 en
         binario y daria 54.14, mientras que la planilla devuelve 54.15.

    La diferencia es de un centavo, pero aparece en decenas de filas y haria
    fallar la verificacion del ETL contra la planilla, que es justamente el
    control que garantiza el entregable.
    """
    return float(Decimal(f"{x:.15g}").quantize(Decimal("0.01"),
                                               rounding=ROUND_HALF_UP))


SEMILLA = 20250906
SEMILLA_SUCIEDAD = 4711

# --- dimensiones -------------------------------------------------------------
# (producto, categoria, precio_lista_2024, costo_sobre_precio)
# El costo/precio distinto por producto es lo que hace util el grafico de
# dispersion: sin dispersion de margen no hay anomalias que detectar.
PRODUCTOS = [
    ("Laptop Pro 15",      "Computacion",    1200.0, 0.620),
    ("Monitor 4K 27",      "Computacion",     450.0, 0.605),
    ("PC Gamer Ryzen",     "Computacion",    1650.0, 0.680),
    ("Mouse Inalambrico",  "Accesorios",       28.0, 0.480),
    ("Teclado Mecanico",   "Accesorios",       95.0, 0.505),
    ("Webcam Full HD",     "Accesorios",       65.0, 0.545),
    ("Auriculares BT Pro", "Audio",           120.0, 0.530),
    ("Parlante Bluetooth", "Audio",            85.0, 0.565),
    ("SSD Externo 1TB",    "Almacenamiento",  130.0, 0.585),
    ("Pendrive 256GB",     "Almacenamiento",   35.0, 0.510),
]

CATEGORIAS = ["Computacion", "Accesorios", "Audio", "Almacenamiento"]

# region -> (peso de demanda, ciudad de referencia del modelo SQL)
REGIONES = {
    "AMBA":    (0.38, "Buenos Aires"),
    "Centro":  (0.22, "Cordoba"),
    "Litoral": (0.16, "Rosario"),
    "Cuyo":    (0.12, "Mendoza"),
    "NOA":     (0.12, "Tucuman"),
}

CANALES = ["Tienda fisica", "E-commerce"]

MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun",
         "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]

# --- supuestos del modelo (se documentan en la hoja Parametros) --------------
ESTACIONALIDAD = {1: 0.85, 2: 0.80, 3: 0.95, 4: 0.92, 5: 1.00, 6: 1.05,
                  7: 1.10, 8: 0.98, 9: 1.00, 10: 1.05, 11: 1.25, 12: 1.45}

# participacion del canal e-commerce sobre las unidades
MIX_ECOM = {2024: 0.28, 2025: 0.41}

# costo logistico como % de las ventas, por canal / anio / trimestre
TASA_LOGISTICA = {
    ("Tienda fisica", 2024): {1: 0.012, 2: 0.012, 3: 0.013, 4: 0.013},
    ("Tienda fisica", 2025): {1: 0.013, 2: 0.014, 3: 0.013, 4: 0.013},
    ("E-commerce",    2024): {1: 0.042, 2: 0.044, 3: 0.046, 4: 0.048},
    # 2025: salta por la ultima milla y recien baja en Q3 con el envio minimo
    ("E-commerce",    2025): {1: 0.068, 2: 0.074, 3: 0.056, 4: 0.052},
}


def descuento(categoria, anio, trimestre):
    """Descuento sobre precio de lista: guerra de precios en Computacion H1-2025."""
    if anio == 2024:
        return 0.02
    if categoria == "Computacion" and trimestre in (1, 2):
        return 0.07
    if trimestre in (1, 2):
        return 0.04
    return 0.03


# aumento de precio de lista 2025 vs 2024
INFLACION_PRECIO = 0.035
# inflacion de costos de compra 2025 vs 2024
INFLACION_COSTO = 0.030
# crecimiento de unidades 2025 vs 2024
CRECIMIENTO_UNIDADES = 0.155
# escala del negocio (una cadena, no un local suelto)
ESCALA = 3.0

# --- el export crudo ---------------------------------------------------------
# Columnas tal como las emite el sistema transaccional. Ojo: no trae Anio,
# Trimestre, Periodo ni ninguna medida calculada; todo eso lo deriva el ETL.
COLUMNAS_CRUDAS = [
    "ID_Venta", "Fecha", "Region", "Ciudad", "Canal", "Categoria",
    "Producto", "Unidades", "Precio_Unitario", "Costo_Unitario",
]

# Variantes de texto que el sistema devuelve para un mismo valor real. La
# primera de cada lista es la forma canonica; el resto es lo que hay que
# estandarizar. La hoja Parametros publica esta misma tabla como diccionario
# de normalizacion y el ETL la consulta con INDICE/COINCIDIR.
VARIANTES_CATEGORIA = {
    "Computacion":    ["Computacion", "computacion", "COMPUTACION",
                       " Computacion ", "Computación", "computación"],
    "Accesorios":     ["Accesorios", "accesorios", "ACCESORIOS",
                       "Accesorios ", " accesorios"],
    "Audio":          ["Audio", "audio", "AUDIO", " Audio"],
    "Almacenamiento": ["Almacenamiento", "almacenamiento", "ALMACENAMIENTO",
                       " Almacenamiento ", "Almacenamiento "],
}
VARIANTES_REGION = {
    "AMBA":    ["AMBA", "amba", "Amba", " AMBA ", "A.M.B.A."],
    "Centro":  ["Centro", "centro", "CENTRO", " Centro"],
    "Litoral": ["Litoral", "litoral", "LITORAL", "Litoral "],
    "Cuyo":    ["Cuyo", "cuyo", "CUYO", " Cuyo "],
    "NOA":     ["NOA", "noa", "Noa", " NOA"],
}
VARIANTES_CANAL = {
    "Tienda fisica": ["Tienda fisica", "tienda fisica", "TIENDA FISICA",
                      "Tienda Fisica", " Tienda fisica ", "Tienda física"],
    "E-commerce":    ["E-commerce", "e-commerce", "E-COMMERCE", "Ecommerce",
                      "E commerce", " E-commerce "],
}

# cuantas filas del export vienen repetidas y cuantas con cada defecto
N_DUPLICADOS = 58
PROP_FECHA_TEXTO_DMA = 0.22    # "05/03/2024"
PROP_FECHA_TEXTO_ISO = 0.13    # "2024-03-05"
PROP_UNIDADES_TEXTO = 0.10     # "12" en vez de 12
PROP_CIUDAD_VACIA = 0.06       # celda en blanco

# columnas de la tabla limpia (las que el ETL reconstruye)
COLUMNAS = [
    "ID_Venta", "Fecha", "Anio", "Trimestre", "Mes_Num", "Mes", "Periodo",
    "Region", "Ciudad", "Canal", "Categoria", "Cat_Anio", "Producto",
    "Prod_Anio", "Unidades", "Precio_Unitario", "Costo_Unitario",
    "Ventas", "Costo_Producto", "Costo_Logistico", "Margen_Bruto",
    "Margen_Pct", "Segmento_Rentabilidad", "Alerta_Logistica", "Ticket",
]
IX = {c: i for i, c in enumerate(COLUMNAS)}
IX_CRUDO = {c: i for i, c in enumerate(COLUMNAS_CRUDAS)}

# umbrales de las columnas logicas (viven en la hoja Parametros)
META_MARGEN = 0.34          # margen bruto % objetivo
PISO_MARGEN = 0.28          # por debajo de esto la operacion se revisa
TICKET_ALTO = 3000.0        # venta grande, en USD
TICKET_BAJO = 300.0         # venta chica, en USD
TECHO_LOG_ECOM = 0.05       # costo logistico / ventas tolerado en e-commerce


def segmento_rentabilidad(margen_pct, ventas):
    """=SI(Y(...);...;SI(O(...);...;...)) - replica exacta de la formula."""
    if margen_pct >= META_MARGEN and ventas >= TICKET_ALTO:
        return "Alta rentabilidad"
    if margen_pct < PISO_MARGEN or ventas < TICKET_BAJO:
        return "Revisar"
    return "Estandar"


def alerta_logistica(canal, costo_log, ventas):
    """=SI(Y(canal="E-commerce"; costo_log/ventas > techo); ...)"""
    ratio = costo_log / ventas if ventas else 0.0
    return "Alerta" if (canal == "E-commerce" and ratio > TECHO_LOG_ECOM) else "OK"


def generar_limpio():
    """Devuelve (columnas, filas) de la tabla de hechos ya correcta."""
    rnd = random.Random(SEMILLA)
    filas = []
    idv = 0
    for anio in (2024, 2025):
        for mes in range(1, 13):
            trim = (mes - 1) // 3 + 1
            for region, (peso_reg, ciudad) in REGIONES.items():
                for prod, cat, precio24, ratio_costo in PRODUCTOS:
                    costo24 = precio24 * ratio_costo
                    # demanda base mensual del producto en la region
                    base = {"Computacion": 26, "Accesorios": 120,
                            "Audio": 55, "Almacenamiento": 70}[cat]
                    base *= peso_reg * ESTACIONALIDAD[mes] * ESCALA
                    if anio == 2025:
                        base *= (1 + CRECIMIENTO_UNIDADES)
                    for canal in CANALES:
                        mix = (MIX_ECOM[anio] if canal == "E-commerce"
                               else 1 - MIX_ECOM[anio])
                        u = base * mix * rnd.uniform(0.82, 1.18)
                        unidades = max(1, int(round(u)))

                        desc = descuento(cat, anio, trim)
                        precio = precio24 * (1 + INFLACION_PRECIO
                                             if anio == 2025 else 1.0)
                        precio = r2(precio * (1 - desc))
                        costo = r2(costo24 * (1 + INFLACION_COSTO)
                                   if anio == 2025 else costo24)

                        ventas = r2(unidades * precio)
                        costo_prod = r2(unidades * costo)
                        tasa = TASA_LOGISTICA[(canal, anio)][trim]
                        costo_log = r2(ventas * tasa)
                        margen = r2(ventas - costo_prod - costo_log)
                        margen_pct = margen / ventas if ventas else 0.0

                        idv += 1
                        # dia representativo: reparte las filas dentro del mes
                        dia = 1 + (idv * 7) % 27
                        filas.append([
                            idv,
                            date(anio, mes, dia),
                            anio,
                            f"Q{trim}",
                            mes,
                            MESES[mes - 1],
                            f"{anio}-{mes:02d}",
                            region,
                            ciudad,
                            canal,
                            cat,
                            f"{cat} {anio}",
                            prod,
                            f"{prod} {anio}",
                            unidades,
                            precio,
                            costo,
                            ventas,
                            costo_prod,
                            costo_log,
                            margen,
                            margen_pct,
                            segmento_rentabilidad(margen_pct, ventas),
                            alerta_logistica(canal, costo_log, ventas),
                            "Alto" if ventas >= TICKET_ALTO else
                            ("Bajo" if ventas < TICKET_BAJO else "Medio"),
                        ])
    return COLUMNAS, filas


def ensuciar(filas_limpias):
    """
    Devuelve (columnas_crudas, filas_crudas, informe).

    Toma la tabla correcta y la degrada a como sale del sistema: ordena por
    fecha de carga (no por ID), repite filas, escribe las fechas en tres
    formatos distintos, desestandariza los textos, carga algunas unidades como
    texto y deja ciudades vacias.

    `informe` resume cuantos defectos de cada tipo se inyectaron; la hoja
    Control_Calidad lo contrasta contra lo que las formulas detectan.
    """
    rnd = random.Random(SEMILLA_SUCIEDAD)

    def variante(mapa, valor):
        return rnd.choice(mapa[valor])

    crudas = []
    for f in filas_limpias:
        fecha = f[IX["Fecha"]]
        r = rnd.random()
        if r < PROP_FECHA_TEXTO_DMA:
            fecha_out = f"{fecha.day:02d}/{fecha.month:02d}/{fecha.year}"
        elif r < PROP_FECHA_TEXTO_DMA + PROP_FECHA_TEXTO_ISO:
            fecha_out = f"{fecha.year}-{fecha.month:02d}-{fecha.day:02d}"
        else:
            fecha_out = fecha

        unidades = f[IX["Unidades"]]
        if rnd.random() < PROP_UNIDADES_TEXTO:
            unidades = str(unidades)

        ciudad = "" if rnd.random() < PROP_CIUDAD_VACIA else f[IX["Ciudad"]]

        crudas.append([
            f[IX["ID_Venta"]],
            fecha_out,
            variante(VARIANTES_REGION, f[IX["Region"]]),
            ciudad,
            variante(VARIANTES_CANAL, f[IX["Canal"]]),
            variante(VARIANTES_CATEGORIA, f[IX["Categoria"]]),
            # el producto solo sufre espacios de mas
            rnd.choice([f[IX["Producto"]], f" {f[IX['Producto']]}",
                        f"{f[IX['Producto']]} "]),
            unidades,
            f[IX["Precio_Unitario"]],
            f[IX["Costo_Unitario"]],
        ])

    # --- duplicados: reenvios del sistema, fila identica repetida ------------
    indices_dup = rnd.sample(range(len(crudas)), N_DUPLICADOS)
    for i in sorted(indices_dup):
        crudas.append(list(crudas[i]))

    # el export no viene ordenado por ID: viene por lote de carga
    rnd.shuffle(crudas)

    informe = {
        "filas_crudas": len(crudas),
        "filas_unicas": len(filas_limpias),
        "duplicados": N_DUPLICADOS,
        "fechas_texto": sum(1 for r in crudas if isinstance(r[IX_CRUDO["Fecha"]], str)),
        "unidades_texto": sum(1 for r in crudas
                              if isinstance(r[IX_CRUDO["Unidades"]], str)),
        "ciudades_vacias": sum(1 for r in crudas if r[IX_CRUDO["Ciudad"]] == ""),
        "variantes_categoria": sum(len(v) for v in VARIANTES_CATEGORIA.values()),
        "variantes_region": sum(len(v) for v in VARIANTES_REGION.values()),
        "variantes_canal": sum(len(v) for v in VARIANTES_CANAL.values()),
    }
    return COLUMNAS_CRUDAS, crudas, informe


def resumen(filas):
    """Chequeo rapido de que la historia del dataset salio como se buscaba."""
    out = {}
    for anio in (2024, 2025):
        f = [r for r in filas if r[IX["Anio"]] == anio]
        v = sum(r[IX["Ventas"]] for r in f)
        m = sum(r[IX["Margen_Bruto"]] for r in f)
        cl = sum(r[IX["Costo_Logistico"]] for r in f)
        u = sum(r[IX["Unidades"]] for r in f)
        out[anio] = dict(ventas=v, margen=m, margen_pct=m / v,
                         costo_log=cl, costo_log_pct=cl / v, unidades=u)
    out["yoy_ventas"] = out[2025]["ventas"] / out[2024]["ventas"] - 1
    out["delta_margen_pp"] = (out[2025]["margen_pct"] - out[2024]["margen_pct"]) * 100
    for anio in (2024, 2025):
        for q in (1, 2, 3, 4):
            f = [r for r in filas
                 if r[IX["Anio"]] == anio and r[IX["Trimestre"]] == f"Q{q}"]
            v = sum(r[IX["Ventas"]] for r in f)
            m = sum(r[IX["Margen_Bruto"]] for r in f)
            out[f"{anio}-Q{q} margen%"] = m / v
    for seg in ("Alta rentabilidad", "Estandar", "Revisar"):
        out[f"seg {seg}"] = sum(1 for r in filas
                                if r[IX["Segmento_Rentabilidad"]] == seg)
    out["alertas logisticas"] = sum(1 for r in filas
                                    if r[IX["Alerta_Logistica"]] == "Alerta")
    return out


if __name__ == "__main__":
    cols, filas = generar_limpio()
    print(f"filas limpias: {len(filas)}")
    for k, v in resumen(filas).items():
        print(" ", k, v)
    _, crudas, inf = ensuciar(filas)
    print("export crudo:")
    for k, v in inf.items():
        print(" ", k, v)
