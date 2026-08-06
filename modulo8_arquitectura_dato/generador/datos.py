"""
Dataset simulado de TechStore para el Modulo 8 (arquitectura del dato).

Es el mismo universo de negocio que ya usan el checkpoint de SQL
(`Ventas_Tech_DB`) y el dashboard del Modulo 9: mismas categorias, productos y
ciudades. Aca se lo reorganiza con forma de **modelo dimensional**: una tabla de
hechos transaccional y tres dimensiones conformadas.

Dos decisiones importantes:

1. El dataset es determinista (semilla fija). Dos corridas del generador
   producen exactamente el mismo archivo, asi que los numeros del libro, del
   informe de estadistica descriptiva y del README siempre coinciden.

2. El CSV de origen sale **sucio a proposito**: filas anuladas, celdas vacias,
   numeros guardados como texto con espacios y filas duplicadas. Sin suciedad
   el paso de limpieza de Power Query no tendria nada que demostrar.
"""
import random
from datetime import date, timedelta

SEMILLA = 20260806

# --- dimension Productos -----------------------------------------------------
# (id, producto, categoria, marca, precio_lista, costo_sobre_precio)
PRODUCTOS = [
    ("P001", "Laptop Pro 15",      "Computacion",    "NovaTech", 1200.0, 0.620),
    ("P002", "Monitor 4K 27",      "Computacion",    "Visor",     450.0, 0.605),
    ("P003", "PC Gamer Ryzen",     "Computacion",    "NovaTech", 1650.0, 0.680),
    ("P004", "Mouse Inalambrico",  "Accesorios",     "Klik",       28.0, 0.480),
    ("P005", "Teclado Mecanico",   "Accesorios",     "Klik",       95.0, 0.505),
    ("P006", "Webcam Full HD",     "Accesorios",     "Visor",      65.0, 0.545),
    ("P007", "Auriculares BT Pro", "Audio",          "SonarLab",  120.0, 0.530),
    ("P008", "Parlante Bluetooth", "Audio",          "SonarLab",   85.0, 0.565),
    ("P009", "SSD Externo 1TB",    "Almacenamiento", "DataOne",   130.0, 0.585),
    ("P010", "Pendrive 256GB",     "Almacenamiento", "DataOne",    35.0, 0.510),
]

# region -> (peso de demanda, ciudad de referencia del modelo SQL)
REGIONES = {
    "AMBA":    (0.38, "Buenos Aires"),
    "Centro":  (0.22, "Cordoba"),
    "Litoral": (0.16, "Rosario"),
    "Cuyo":    (0.12, "Mendoza"),
    "NOA":     (0.12, "Tucuman"),
}

SEGMENTOS = ["Consumidor final", "PyME", "Corporativo", "Sector publico"]
CANALES = ["Tienda fisica", "E-commerce"]

# nombres de fantasia: el dataset es simulado, no hay clientes reales detras
NOMBRES = [
    "Alsina", "Barrancas", "Calchaqui", "Duarte", "Escalada", "Ferreyra",
    "Guemes", "Huergo", "Iriarte", "Juarez", "Krause", "Lavalle", "Mansilla",
    "Nogues", "Ortiz", "Pellegrini", "Quiroga", "Rivadavia", "Sarmiento",
    "Talcahuano", "Urquiza", "Videla", "Warnes", "Yrigoyen", "Zapiola",
]
SUFIJOS = ["SA", "SRL", "y Cia", "Servicios", "Group", "Distribuciones"]

MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
         "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
DIAS = ["Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado", "Domingo"]

DESDE = date(2024, 1, 1)
HASTA = date(2025, 12, 31)
FILAS_LIMPIAS = 5200

# El negocio crece en 2025: mas operaciones (45% / 55% del total) y tickets un
# poco mas grandes. Sin ese crecimiento la hoja de escenarios no tendria nada
# que proyectar para 2026.
PESO_ANIO = {2024: 0.45, 2025: 0.55}
CRECIMIENTO_TICKET = {2024: 1.00, 2025: 1.10}

# estacionalidad por mes (indice 1..12): pico en noviembre-diciembre
ESTACION = {1: 0.82, 2: 0.80, 3: 0.95, 4: 0.98, 5: 1.02, 6: 1.05,
            7: 1.08, 8: 1.00, 9: 0.97, 10: 1.06, 11: 1.28, 12: 1.42}


def _clientes(rnd):
    """60 clientes con region y segmento. ID_Cliente es la PK de la dimension."""
    filas = []
    regiones = list(REGIONES)
    for i in range(1, 61):
        region = regiones[i % len(regiones)]
        nombre = f"{NOMBRES[(i * 7) % len(NOMBRES)]} {SUFIJOS[i % len(SUFIJOS)]}"
        filas.append({
            "ID_Cliente": f"C{i:03d}",
            "Cliente": nombre,
            "Segmento": SEGMENTOS[i % len(SEGMENTOS)],
            "Region": region,
            "Ciudad": REGIONES[region][1],
            "Alta": (date(2022, 1, 1) + timedelta(days=rnd.randint(0, 900))).isoformat(),
        })
    return filas


def _productos():
    return [{
        "ID_Producto": pid,
        "Producto": nombre,
        "Categoria": cat,
        "Marca": marca,
        "Precio_Lista": round(precio, 2),
        "Costo_Estandar": round(precio * costo, 2),
    } for pid, nombre, cat, marca, precio, costo in PRODUCTOS]


def _calendario():
    """Dimension Calendario dia a dia. En el libro tambien se genera en M."""
    filas, d = [], DESDE
    while d <= HASTA:
        filas.append({
            "Fecha": d,
            "Anio": d.year,
            "Trimestre": f"Q{(d.month - 1) // 3 + 1}",
            "Mes": d.month,
            "Nombre_Mes": MESES[d.month - 1],
            "Anio_Mes": f"{d.year}-{d.month:02d}",
            "Dia_Semana": DIAS[d.weekday()],
            "Es_Fin_Semana": "Si" if d.weekday() >= 5 else "No",
        })
        d += timedelta(days=1)
    return filas


def _peso_regiones():
    ids, pesos = [], []
    for region, (peso, _) in REGIONES.items():
        ids.append(region)
        pesos.append(peso)
    return ids, pesos


def _ventas(rnd, clientes):
    """
    Tabla de hechos. La historia que codifica:

      * las ventas crecen en 2025 (mas unidades y precios actualizados),
      * el mix se corre hacia e-commerce, que tiene un costo de envio mucho
        mas alto por operacion,
      * y por eso el margen neto se erosiona aunque la venta crezca. Ese es
        justamente el problema que despues simula la hoja de escenarios.
    """
    por_region = {c["Region"]: [] for c in clientes}
    for c in clientes:
        por_region[c["Region"]].append(c["ID_Cliente"])
    regiones, pesos = _peso_regiones()

    anios, pesos_anio = list(PESO_ANIO), list(PESO_ANIO.values())
    filas = []
    for i in range(1, FILAS_LIMPIAS + 1):
        # primero el anio (el negocio crece), despues el dia dentro del anio
        # por muestreo de rechazo para respetar la estacionalidad
        anio = rnd.choices(anios, weights=pesos_anio)[0]
        inicio = date(anio, 1, 1)
        while True:
            f = inicio + timedelta(days=rnd.randint(0, 364))
            if rnd.random() < ESTACION[f.month] / 1.42:
                break

        region = rnd.choices(regiones, weights=pesos)[0]
        cliente = rnd.choice(por_region[region])

        pid, _, cat, _, precio_lista, costo_ratio = rnd.choices(
            PRODUCTOS, weights=[3, 4, 2, 9, 7, 5, 6, 6, 5, 8])[0]

        # el e-commerce gana peso en 2025: 38% -> 55% de las operaciones
        p_ecom = 0.38 if f.year == 2024 else 0.55
        canal = "E-commerce" if rnd.random() < p_ecom else "Tienda fisica"

        unidades = max(1, int(rnd.lognormvariate(1.05, 0.75) * CRECIMIENTO_TICKET[f.year]))
        if cat == "Accesorios":
            unidades += rnd.randint(0, 6)

        # 2025 actualiza precios ~6%; el costo de compra sube un poco mas
        ajuste_precio = 1.0 if f.year == 2024 else 1.06
        ajuste_costo = 1.0 if f.year == 2024 else 1.085
        precio = precio_lista * ajuste_precio * rnd.uniform(0.98, 1.03)
        costo = precio_lista * costo_ratio * ajuste_costo * rnd.uniform(0.98, 1.02)

        descuento = 0.0
        if canal == "E-commerce" and rnd.random() < 0.45:
            descuento = rnd.choice([0.05, 0.08, 0.10, 0.15])
        elif rnd.random() < 0.18:
            descuento = rnd.choice([0.05, 0.10])

        # el envio es el gran diferencial de costo entre canales
        if canal == "E-commerce":
            envio = round(rnd.uniform(9.0, 26.0) + unidades * rnd.uniform(1.2, 3.4), 2)
        else:
            envio = round(rnd.uniform(0.0, 4.0), 2)

        filas.append({
            "ID_Venta": f"V{i:05d}",
            "Fecha": f,
            "ID_Cliente": cliente,
            "ID_Producto": pid,
            "Canal": canal,
            "Estado": "Confirmada",
            "Unidades": unidades,
            "Precio_Unitario": round(precio, 2),
            "Descuento": round(descuento, 2),
            "Costo_Unitario": round(costo, 2),
            "Costo_Envio": envio,
        })
    return filas


def _ensuciar(rnd, ventas):
    """
    Devuelve las filas tal como salen del sistema transaccional: las 3000 filas
    limpias mas el ruido que Power Query tiene que sacar.

      * 45 operaciones anuladas   -> se filtran por Estado
      * 30 filas con Unidades o ID_Cliente vacios -> se eliminan como nulos
      * 15 filas duplicadas exactas -> se quitan con Table.Distinct

    Nada de esto se toca despues: es el estado "crudo" del archivo fuente.
    """
    sucias = []
    contador = FILAS_LIMPIAS

    for _ in range(45):
        contador += 1
        base = dict(rnd.choice(ventas))
        base["ID_Venta"] = f"V{contador:05d}"
        base["Estado"] = "Anulada"
        sucias.append(base)

    for k in range(30):
        contador += 1
        base = dict(rnd.choice(ventas))
        base["ID_Venta"] = f"V{contador:05d}"
        if k % 2 == 0:
            base["Unidades"] = ""
        else:
            base["ID_Cliente"] = ""
        sucias.append(base)

    duplicadas = [dict(f) for f in rnd.sample(ventas, 15)]

    todas = ventas + sucias + duplicadas
    rnd.shuffle(todas)
    return todas


def construir():
    rnd = random.Random(SEMILLA)
    clientes = _clientes(rnd)
    productos = _productos()
    calendario = _calendario()
    ventas = _ventas(rnd, clientes)
    crudas = _ensuciar(random.Random(SEMILLA + 1), ventas)
    return {
        "clientes": clientes,
        "productos": productos,
        "calendario": calendario,
        "ventas": ventas,      # las 3000 filas limpias (lo que deja Power Query)
        "crudas": crudas,      # lo que realmente contiene el CSV de origen
    }


# --- metricas derivadas ------------------------------------------------------
# Se calculan aca una sola vez para que el libro, el informe de estadistica y
# el README hablen siempre de los mismos numeros.

def enriquecer(ventas):
    for v in ventas:
        bruta = v["Unidades"] * v["Precio_Unitario"]
        neta = bruta * (1 - v["Descuento"])
        costo = v["Unidades"] * v["Costo_Unitario"]
        v["Venta_Bruta"] = round(bruta, 2)
        v["Venta_Neta"] = round(neta, 2)
        v["Costo_Total"] = round(costo, 2)
        v["Margen_Bruto"] = round(neta - costo - v["Costo_Envio"], 2)
    return ventas


if __name__ == "__main__":
    d = construir()
    enriquecer(d["ventas"])
    print("clientes  ", len(d["clientes"]))
    print("productos ", len(d["productos"]))
    print("calendario", len(d["calendario"]))
    print("ventas limpias", len(d["ventas"]))
    print("filas en el CSV crudo", len(d["crudas"]))
    print("venta neta total", round(sum(v["Venta_Neta"] for v in d["ventas"]), 2))
    print("margen bruto total", round(sum(v["Margen_Bruto"] for v in d["ventas"]), 2))
