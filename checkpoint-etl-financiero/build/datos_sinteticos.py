"""
datos_sinteticos.py
===================
Genera el archivo de transacciones "sucio" que alimenta el flujo ETL.

La suciedad es intencional: sin ella no habria nada que transformar en
Power Query. Se inyectan, de forma determinista (semilla fija):

  * mayusculas/minusculas y espacios inconsistentes en los textos
  * celdas vacias en columnas clave
  * importes no numericos ("N/D", "-")
  * registros anulados
  * identificadores duplicados
  * dos columnas irrelevantes para el analisis

Las filas validas suman aproximadamente un ano completo de operaciones.
"""

from __future__ import annotations

import random
from datetime import date, timedelta

COLUMNAS = [
    "ID_Transaccion", "Fecha", "Region", "Categoria", "Producto", "Vendedor",
    "Unidades", "Precio_Unitario", "Costo_Unitario", "Estado",
    "Notas_Internas", "Ref_Sistema",
]

REGIONES = ["Norte", "Sur", "Centro", "Litoral", "Cuyo"]

CATALOGO = {
    "Indumentaria": [("Campera Trekking", 1450, 0.42), ("Buzo Polar", 890, 0.40),
                     ("Remera Tecnica", 520, 0.45)],
    "Calzado":      [("Zapatilla Running", 1980, 0.35), ("Botin Sintetico", 1620, 0.33)],
    "Accesorios":   [("Mochila 30L", 1180, 0.44), ("Gorra Deportiva", 380, 0.50)],
    "Deportes":     [("Pelota Profesional", 760, 0.36), ("Guantes Arquero", 940, 0.38)],
}

VENDEDORES = ["Ana Gimenez", "Bruno Salas", "Carla Ortiz", "Diego Funes", "Elena Paz"]

# Variantes "sucias" con las que se escribe cada texto en el origen.
_ENSUCIAR = [
    lambda t: t,
    lambda t: t.upper(),
    lambda t: t.lower(),
    lambda t: "  " + t,
    lambda t: t + "   ",
    lambda t: "  " + t.upper() + " ",
]


def _ensuciar(texto: str, rnd: random.Random, probabilidad: float = 0.45) -> str:
    if rnd.random() < probabilidad:
        return rnd.choice(_ENSUCIAR)(texto)
    return texto


def generar(cantidad: int = 320, semilla: int = 2024) -> list[dict]:
    """Devuelve las filas del origen crudo, ya como texto."""
    rnd = random.Random(semilla)
    inicio = date(2024, 1, 1)
    filas: list[dict] = []

    for i in range(1, cantidad + 1):
        categoria = rnd.choice(list(CATALOGO))
        producto, precio_base, margen = rnd.choice(CATALOGO[categoria])

        precio = round(precio_base * rnd.uniform(0.92, 1.12), 2)
        costo = round(precio * (1 - margen * rnd.uniform(0.85, 1.15)), 2)
        unidades = rnd.choices([1, 2, 3, 4, 5, 6], weights=[30, 26, 18, 12, 9, 5])[0]
        fecha = inicio + timedelta(days=rnd.randint(0, 364))

        fila = {
            "ID_Transaccion": f"TRX-{i:05d}",
            "Fecha": fecha.strftime("%d/%m/%Y"),
            "Region": _ensuciar(rnd.choice(REGIONES), rnd),
            "Categoria": _ensuciar(categoria, rnd),
            "Producto": _ensuciar(producto, rnd, 0.25),
            "Vendedor": _ensuciar(rnd.choice(VENDEDORES), rnd, 0.3),
            "Unidades": str(unidades),
            "Precio_Unitario": f"{precio:.2f}",
            "Costo_Unitario": f"{costo:.2f}",
            "Estado": rnd.choice(["Confirmada", "CONFIRMADA", "confirmada"]),
            "Notas_Internas": rnd.choice(["", "revisar", "ok", "llamar al cliente"]),
            "Ref_Sistema": f"SYS{rnd.randint(100000, 999999)}",
        }
        filas.append(fila)

    # --- suciedad dirigida -------------------------------------------------
    posiciones = rnd.sample(range(len(filas)), 60)

    for p in posiciones[:12]:                      # region vacia
        filas[p]["Region"] = ""
    for p in posiciones[12:22]:                    # unidades no numericas
        filas[p]["Unidades"] = rnd.choice(["N/D", "-", ""])
    for p in posiciones[22:30]:                    # precio invalido
        filas[p]["Precio_Unitario"] = rnd.choice(["0", "", "s/d"])
    for p in posiciones[30:44]:                    # operaciones anuladas
        filas[p]["Estado"] = rnd.choice(["Anulada", "ANULADA", "anulada"])
    for p in posiciones[44:52]:                    # categoria vacia
        filas[p]["Categoria"] = ""
    for p in posiciones[52:]:                      # fecha ilegible
        filas[p]["Fecha"] = rnd.choice(["", "31/02/2024", "sin fecha"])

    # identificadores duplicados (mismo ID, fila repetida)
    for p in rnd.sample(range(len(filas)), 8):
        copia = dict(filas[p])
        copia["Notas_Internas"] = "duplicado de carga"
        filas.append(copia)

    rnd.shuffle(filas)
    return filas


def limpiar(filas: list[dict]) -> list[dict]:
    """Replica en Python las transformaciones del codigo M.

    Sirve para calcular los valores que se guardan en cache dentro del
    libro (KPI y modelo financiero) y para verificar que la consulta de
    Power Query devuelve exactamente el mismo resultado.
    """
    vistos: set[str] = set()
    limpias: list[dict] = []

    for fila in filas:
        # 2. estandarizacion de textos
        region = " ".join(fila["Region"].split()).title()
        categoria = " ".join(fila["Categoria"].split()).title()
        producto = " ".join(fila["Producto"].split()).title()
        vendedor = " ".join(fila["Vendedor"].split()).title()
        estado = fila["Estado"].strip().upper()

        # 3. tipos de datos
        try:
            dia, mes, anio = (int(x) for x in fila["Fecha"].split("/"))
            fecha = date(anio, mes, dia)
        except (ValueError, TypeError):
            continue
        try:
            unidades = int(fila["Unidades"])
            precio = float(fila["Precio_Unitario"])
            costo = float(fila["Costo_Unitario"])
        except ValueError:
            continue

        # 4. nulos, importes invalidos y anuladas
        if not region or not categoria or unidades <= 0 or precio <= 0 or costo < 0:
            continue
        if estado == "ANULADA":
            continue

        # 5. duplicados por clave de negocio
        if fila["ID_Transaccion"] in vistos:
            continue
        vistos.add(fila["ID_Transaccion"])

        # 6. columnas calculadas
        monto = unidades * precio
        costo_total = unidades * costo
        margen = monto - costo_total
        limpias.append({
            "ID_Transaccion": fila["ID_Transaccion"],
            "Fecha": fecha,
            "Anio": fecha.year,
            "Trimestre": f"T{(fecha.month - 1) // 3 + 1}",
            "Region": region,
            "Categoria": categoria,
            "Producto": producto,
            "Vendedor": vendedor,
            "Unidades": unidades,
            "Precio_Unitario": precio,
            "Costo_Unitario": costo,
            "Monto": monto,
            "Costo_Total": costo_total,
            "Margen": margen,
            "Margen_Pct": margen / monto if monto else 0,
        })

    limpias.sort(key=lambda f: f["Fecha"])
    return limpias
