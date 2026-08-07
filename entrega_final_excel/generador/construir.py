"""
Arma el libro base con openpyxl: hojas, formulas, formato y graficos.

Las tablas dinamicas y los segmentadores no los sabe escribir openpyxl; los
agrega `pipeline.py` sobre este archivo.
"""
import os

import openpyxl

import analisis
import dashboard
import datos
import documentacion
import etl
import parametros
from analisis import AUX, CALC, TD

AQUI = os.path.dirname(os.path.abspath(__file__))

# el tablero primero; despues el analisis, los datos y la documentacion;
# las hojas de soporte al final y ocultas
ORDEN = [dashboard.DASH, TD, etl.LIMPIOS, etl.ORIGEN, etl.CONTROL,
         parametros.HOJA, documentacion.HOJA, CALC, AUX]


def posiciones_unicas(filas_crudas):
    """
    Paso "quitar duplicados" del ETL.

    Devuelve, para cada ID_Venta ordenado, la posicion (1-based) de su PRIMERA
    aparicion dentro del export crudo. Es la unica decision del ETL que se
    resuelve fuera de la planilla; la columna Control_ETL la verifica despues
    con COINCIDIR, fila por fila.
    """
    primera = {}
    for pos, fila in enumerate(filas_crudas, start=1):
        idv = fila[datos.IX_CRUDO["ID_Venta"]]
        if idv not in primera:
            primera[idv] = pos
    return [primera[k] for k in sorted(primera)]


def main():
    columnas, filas = datos.generar_limpio()
    cols_crudas, crudas, informe = datos.ensuciar(filas)
    posiciones = posiciones_unicas(crudas)
    assert len(posiciones) == len(filas), "la deduplicacion no cierra"

    # productos por debajo del corte de margen -> serie ambar de la dispersion
    agg = {}
    for f in filas:
        if f[datos.IX["Anio"]] == 2025:
            k = f[datos.IX["Producto"]]
            v, m = agg.get(k, (0.0, 0.0))
            agg[k] = (v + f[datos.IX["Ventas"]], m + f[datos.IX["Margen_Bruto"]])
    corte = 0.33
    riesgo = {p for p, (v, m) in agg.items() if m / v < corte}

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # Parametros primero: todo lo demas referencia sus celdas
    P = parametros.construir(wb)

    pivots, ref = analisis.definir_pivots()
    analisis.hoja_aux(wb)
    F, K, D, G = analisis.hoja_calculos(wb, ref, riesgo, P)

    ws_dash = dashboard.hoja(wb, F, K, D)
    dashboard.graficos(wb, ws_dash, riesgo)
    dashboard.reporte_ejecutivo(ws_dash)

    analisis.hoja_tablas_dinamicas(wb, P, len(filas))

    ws_ori, fin_crudo = etl.hoja_origen(wb, cols_crudas, crudas)
    etl.leyenda_origen(ws_ori, fin_crudo)
    etl.hoja_limpios(wb, P, posiciones, len(crudas))
    etl.hoja_control(wb, len(crudas), len(filas), informe)

    documentacion.construir(wb)

    for pos, nombre in enumerate(ORDEN):
        wb.move_sheet(nombre, offset=pos - wb.sheetnames.index(nombre))
    wb.active = 0

    salida = os.path.join(AQUI, "libro.xlsx")
    wb.save(salida)
    print(f"libro base: {salida}")
    print(f"  operaciones limpias : {len(filas)}")
    print(f"  filas del export    : {len(crudas)} "
          f"({informe['duplicados']} duplicados)")
    print(f"  productos en riesgo : {sorted(riesgo)}")
    return salida, columnas, filas, pivots, crudas, informe


if __name__ == "__main__":
    main()
