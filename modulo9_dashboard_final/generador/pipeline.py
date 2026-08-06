"""
Arma el entregable de punta a punta.

    1. construir.py    -> libro con datos, formulas, formato y graficos
    2. inject_pivots   -> tablas dinamicas nativas
    3. recalc.py       -> LibreOffice calcula GETPIVOTDATA y cachea los valores
    4. finalize        -> segmentadores + lo que LibreOffice pisa al guardar

Los segmentadores van al final a proposito: LibreOffice entiende las tablas
dinamicas pero descarta los slicers al reescribir el archivo.
"""
import json
import os
import shutil
import subprocess
import sys

import construir
import ooxml

AQUI = os.path.dirname(os.path.abspath(__file__))
RECALC = "/root/.claude/skills/xlsx/scripts/recalc.py"

SEGMENTADORES = [
    # (campo, titulo, col_desde, col_hasta, columnas de la lista)
    ("Trimestre", "Trimestre", 1, 5, 2),
    ("Categoria", "Categoria", 5, 9, 2),
    ("Region", "Region", 9, 13, 2),
    ("Canal", "Canal", 13, 17, 1),
]
FILA_DESDE, FILA_HASTA = 14, 23   # base 0: filas 15 a 23 de la hoja


def main(destino):
    ruta, columnas, filas, pivots = construir.main()

    print("2/4  inyectando tablas dinamicas...")
    campos_slicer = [c for c, *_ in SEGMENTADORES]
    ooxml.inject_pivots(ruta, "Datos", columnas, filas, pivots,
                        slicer_fields=campos_slicer)

    print("3/4  recalculando con LibreOffice (tarda unos minutos)...")
    out = subprocess.run([sys.executable, RECALC, ruta, "900"],
                         capture_output=True, text=True, timeout=1200)
    print(out.stdout.strip() or out.stderr.strip())
    try:
        res = json.loads(out.stdout)
    except json.JSONDecodeError:
        raise SystemExit("recalc.py no devolvio JSON")
    if "error" in res:
        raise SystemExit("recalc fallo: " + res["error"])
    if res.get("total_errors"):
        raise SystemExit(f"hay {res['total_errors']} formulas con error")

    # copia previa a los segmentadores: permite reintentar el ultimo paso
    # sin volver a pagar los minutos de recalculo
    shutil.copy(ruta, os.path.join(AQUI, "post_recalc.xlsx"))

    # cada grafico se reconoce por un rango propio de la hoja Calculos
    n = ooxml.titulos_dinamicos(ruta, [
        (f"$A${construir.R_MES}:", construir.T_TITULOS["g1"]),
        (f"$A${construir.R_CASC}:", construir.T_TITULOS["g2"]),
        ("<c:scatterChart>", construir.T_TITULOS["g3"]),
        (f"$A${construir.R_CAT}:", construir.T_TITULOS["g4"]),
    ])
    print(f"     titulos dinamicos aplicados: {n}")
    if n != 4:
        raise SystemExit(f"se esperaban 4 titulos dinamicos, se aplicaron {n}")

    print("4/4  inyectando segmentadores...")
    slicers = [ooxml.Slicer(campo, titulo, c1, FILA_DESDE, c2, FILA_HASTA,
                            column_count=ncol)
               for campo, titulo, c1, c2, ncol in SEGMENTADORES]
    ooxml.finalize(ruta, slicers, construir.DASH,
                   gridless_sheets=[construir.DASH, "Analisis_IA",
                                    "Parametros", "Guia_de_uso"])

    os.makedirs(os.path.dirname(destino), exist_ok=True)
    shutil.copy(ruta, destino)
    print("listo ->", destino)
    return destino


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else os.path.join(AQUI, "final.xlsx"))
