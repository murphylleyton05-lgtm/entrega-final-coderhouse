"""
Auditoria del libro ya recalculado.

La hoja `Datos_Limpios` reconstruye la tabla de hechos con formulas. Que las
formulas *evaluen* sin error no prueba que devuelvan lo correcto: un rango
corrido por una fila da un archivo limpio y numeros equivocados.

Este modulo compara, celda por celda, lo que la planilla calculo contra lo que
el modelo dice que tendria que dar. Si algo no coincide, el build falla y el
entregable no se genera.
"""
import datetime as dt

import openpyxl

import analisis
import datos
import etl

TOL = 0.005          # tolerancia en USD para las medidas monetarias
MUESTRA_ERRORES = 8


class Fallo(Exception):
    pass


def _valor(ws, col, fila):
    return ws[f"{col}{fila}"].value


def _iguales(a, b):
    if isinstance(a, dt.datetime):
        a = a.date()
    if isinstance(b, dt.datetime):
        b = b.date()
    if isinstance(a, float) or isinstance(b, float):
        if a is None or b is None:
            return False
        return abs(float(a) - float(b)) <= TOL
    return a == b


def verificar(ruta, columnas, filas, crudas, informe):
    wb = openpyxl.load_workbook(ruta, data_only=True)
    problemas = []

    def check(cond, mensaje):
        if not cond:
            problemas.append(mensaje)

    # --- 1. el ETL reproduce exactamente la tabla de hechos ----------------
    ws = wb[etl.LIMPIOS]
    revisadas = 0
    for i, esperada in enumerate(filas):
        r = i + 2
        for campo in datos.COLUMNAS:
            obtenido = _valor(ws, etl.CL[campo], r)
            espera = esperada[datos.IX[campo]]
            revisadas += 1
            if not _iguales(obtenido, espera):
                problemas.append(
                    f"Datos_Limpios!{etl.CL[campo]}{r} ({campo}): "
                    f"la planilla dio {obtenido!r}, se esperaba {espera!r}")
                if len(problemas) > 40:
                    raise Fallo(_informe(problemas))
    print(f"     celdas del ETL comparadas: {revisadas:,}")

    # --- 2. la deduplicacion apunta a la primera aparicion -----------------
    malos = sum(1 for i in range(len(filas))
                if _valor(ws, etl.C_CTRL, i + 2) != "OK")
    check(malos == 0,
          f"{malos} filas con Control_ETL distinto de OK (deduplicacion)")

    # --- 3. la hoja de control de calidad dice lo que tiene que decir ------
    wc = wb[etl.CONTROL]
    esperado_control = {
        "duplicados_origen": informe["duplicados"],
        "fechas_texto_origen": informe["fechas_texto"],
        "unidades_texto_origen": informe["unidades_texto"],
        "ciudades_vacias_origen": informe["ciudades_vacias"],
    }
    # las cuatro primeras filas del diagnostico, columna D (en el origen)
    for j, (clave, esperado) in enumerate(esperado_control.items()):
        obtenido = _valor(wc, "D", 8 + j)
        check(_iguales(obtenido, esperado),
              f"Control_Calidad!D{8 + j} ({clave}): {obtenido!r}, "
              f"se esperaba {esperado}")
    # y las mismas cuatro despues del ETL: todas en cero
    for j in range(4):
        obtenido = _valor(wc, "E", 8 + j)
        check(_iguales(obtenido, 0),
              f"Control_Calidad!E{8 + j}: quedaron {obtenido!r} defectos "
              f"despues del ETL, se esperaba 0")
    # los valores estandarizados: 4 categorias, 5 regiones, 2 canales, 10 prod
    for j, esperado in enumerate((len(datos.CATEGORIAS), len(datos.REGIONES),
                                  len(datos.CANALES), len(datos.PRODUCTOS))):
        obtenido = _valor(wc, "E", 12 + j)
        check(_iguales(obtenido, esperado),
              f"Control_Calidad!E{12 + j}: {obtenido!r} valores distintos, "
              f"se esperaba {esperado}")

    # --- 4. los KPIs del tablero ------------------------------------------
    wk = wb[analisis.CALC]
    res = datos.resumen(filas)
    kpis = [
        (analisis.R_TOT, "B", res[2024]["ventas"], "Ventas 2024"),
        (analisis.R_TOT, "C", res[2025]["ventas"], "Ventas 2025"),
        (analisis.R_TOT + 3, "B", res[2024]["margen"], "Margen 2024"),
        (analisis.R_TOT + 3, "C", res[2025]["margen"], "Margen 2025"),
        (analisis.R_TOT + 4, "B", res[2024]["unidades"], "Unidades 2024"),
        (analisis.R_TOT + 4, "C", res[2025]["unidades"], "Unidades 2025"),
    ]
    for fila, col, esperado, nombre in kpis:
        obtenido = _valor(wk, col, fila)
        check(obtenido is not None and abs(float(obtenido) - esperado) <= 1.0,
              f"Calculos!{col}{fila} ({nombre}): {obtenido!r}, "
              f"se esperaba {esperado:,.2f}")

    # el puente de margen tiene que cerrar en cero
    cierre = _valor(wk, "B", analisis.R_CASC + len(
        ["2024", "vol", "precio", "compra", "log", "2025"]))
    check(cierre is not None and abs(float(cierre)) <= 1.0,
          f"el puente de margen no cierra: {cierre!r}")

    # --- 5. las tablas dinamicas se materializaron -------------------------
    # con varias medidas, la tabla dinamica gasta dos filas de encabezado
    # (la del campo virtual "Valores" y la de los nombres de medida), asi que
    # los datos arrancan dos filas debajo del ancla
    wt = wb[analisis.TD]
    total = sum(res[a]["ventas"] for a in (2024, 2025))
    for nombre, col, ancla, n in (("regiones", "C", 6, len(datos.REGIONES)),
                                  ("categorias", "H", 6, len(datos.CATEGORIAS)),
                                  ("productos", "C", 16, len(datos.PRODUCTOS))):
        leidos = [_valor(wt, col, ancla + 2 + k) for k in range(n)]
        if not all(isinstance(v, (int, float)) for v in leidos):
            problemas.append(f"la tabla dinamica de {nombre} no tiene valores "
                             f"cacheados: {leidos!r}")
            continue
        check(abs(sum(leidos) - total) <= 1.0,
              f"la tabla dinamica de {nombre} suma {sum(leidos):,.2f}, "
              f"se esperaba {total:,.2f}")

    # y los encabezados tienen que estar en castellano, no en el ingles que
    # deja LibreOffice al materializarlas
    for celda, esperado in (("C6", "Valores"), ("B13", "Total general")):
        obtenido = wt[celda].value
        check(obtenido == esperado,
              f"Tablas_Dinamicas!{celda}: {obtenido!r}, se esperaba "
              f"{esperado!r}")

    if problemas:
        raise Fallo(_informe(problemas))
    print("     verificacion OK: el ETL de la planilla reproduce el modelo")


def _informe(problemas):
    cabeza = "\n".join(f"  - {p}" for p in problemas[:MUESTRA_ERRORES])
    resto = (f"\n  ... y {len(problemas) - MUESTRA_ERRORES} mas"
             if len(problemas) > MUESTRA_ERRORES else "")
    return f"la verificacion encontro {len(problemas)} problemas:\n{cabeza}{resto}"


if __name__ == "__main__":
    import sys

    import construir
    cols, fs = datos.generar_limpio()
    _, cr, inf = datos.ensuciar(fs)
    verificar(sys.argv[1], cols, fs, cr, inf)
