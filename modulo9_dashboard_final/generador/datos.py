"""
Generador del dataset simulado de TechStore (2024-2025).

El dataset es SIMULADO y determinista (semilla fija): reproduce el modelo de
Ventas_Tech_DB del checkpoint de SQL (categorias, productos y ciudades) y lo
extiende a 24 meses, 5 regiones y 2 canales para poder analizar margen,
estacionalidad y evolucion interanual.

La historia que codifica el dataset (y que el dashboard debe hacer evidente):
  * las ventas crecen con fuerza en 2025,
  * pero el margen % cae, porque el mix se corre a e-commerce, que tiene un
    costo logistico por peso vendido mucho mas alto,
  * y a partir del Q3-2025 el margen se recupera cuando baja la tasa logistica
    de e-commerce (umbral de envio minimo).
"""
import random
from datetime import date

SEMILLA = 20250906

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

# descuento sobre precio de lista: guerra de precios en Computacion en H1-2025
def descuento(categoria, anio, trimestre):
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

COLUMNAS = [
    "ID_Venta", "Fecha", "Anio", "Trimestre", "Mes_Num", "Mes", "Periodo",
    "Region", "Ciudad", "Canal", "Categoria", "Cat_Anio", "Producto",
    "Prod_Anio", "Unidades", "Precio_Unitario", "Costo_Unitario",
    "Ventas", "Costo_Producto", "Costo_Logistico", "Margen_Bruto",
]


def generar():
    """Devuelve (columnas, filas) con las filas ya calculadas."""
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
                        mix = MIX_ECOM[anio] if canal == "E-commerce" else 1 - MIX_ECOM[anio]
                        u = base * mix * rnd.uniform(0.82, 1.18)
                        unidades = max(1, int(round(u)))

                        desc = descuento(cat, anio, trim)
                        precio = precio24 * (1 + INFLACION_PRECIO if anio == 2025 else 1.0)
                        precio = round(precio * (1 - desc), 2)
                        costo = round(costo24 * (1 + INFLACION_COSTO)
                                      if anio == 2025 else costo24, 2)

                        ventas = round(unidades * precio, 2)
                        costo_prod = round(unidades * costo, 2)
                        tasa = TASA_LOGISTICA[(canal, anio)][trim]
                        costo_log = round(ventas * tasa, 2)
                        margen = round(ventas - costo_prod - costo_log, 2)

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
                        ])
    return COLUMNAS, filas


def resumen(filas):
    """Chequeo rapido de que la historia del dataset salio como se buscaba."""
    ix = {c: i for i, c in enumerate(COLUMNAS)}
    out = {}
    for anio in (2024, 2025):
        f = [r for r in filas if r[ix["Anio"]] == anio]
        v = sum(r[ix["Ventas"]] for r in f)
        m = sum(r[ix["Margen_Bruto"]] for r in f)
        cl = sum(r[ix["Costo_Logistico"]] for r in f)
        u = sum(r[ix["Unidades"]] for r in f)
        out[anio] = dict(ventas=v, margen=m, margen_pct=m / v,
                         costo_log=cl, costo_log_pct=cl / v, unidades=u)
    out["yoy_ventas"] = out[2025]["ventas"] / out[2024]["ventas"] - 1
    out["delta_margen_pp"] = (out[2025]["margen_pct"] - out[2024]["margen_pct"]) * 100
    for anio in (2024, 2025):
        for q in (1, 2, 3, 4):
            f = [r for r in filas
                 if r[ix["Anio"]] == anio and r[ix["Trimestre"]] == f"Q{q}"]
            v = sum(r[ix["Ventas"]] for r in f)
            m = sum(r[ix["Margen_Bruto"]] for r in f)
            out[f"{anio}-Q{q} margen%"] = m / v
    return out


if __name__ == "__main__":
    cols, filas = generar()
    print(f"filas: {len(filas)}")
    r = resumen(filas)
    for k, v in r.items():
        print(k, v)
