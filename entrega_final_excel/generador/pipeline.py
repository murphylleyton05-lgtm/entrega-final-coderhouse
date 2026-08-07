"""
Arma el entregable de punta a punta.

    1. construir.py       -> libro con datos, formulas, formato y graficos
    2. inject_pivots      -> tablas dinamicas nativas (visibles y de soporte)
    3. recalc.py          -> LibreOffice calcula todo y cachea los valores
    4. verificar.py       -> el ETL de la planilla debe reproducir exactamente
                             los valores esperados, o el build falla
    5. finalize           -> segmentadores, titulos dinamicos, etiquetas de
                             datos y lo que LibreOffice pisa al guardar

Los segmentadores van al final a proposito: LibreOffice entiende las tablas
dinamicas pero descarta los slicers al reescribir el archivo.
"""
import json
import os
import shutil
import subprocess
import sys

import analisis
import construir
import dashboard
import documentacion
import etl
import ooxml
import parametros

AQUI = os.path.dirname(os.path.abspath(__file__))
RECALC = "/root/.claude/skills/xlsx/scripts/recalc.py"

SEGMENTADORES = [
    # (campo, titulo, col_desde, col_hasta, columnas de la lista)
    ("Trimestre", "Trimestre", 1, 5, 2),
    ("Categoria", "Categoria", 5, 9, 2),
    ("Region", "Region", 9, 13, 2),
    ("Canal", "Canal", 13, 17, 1),
]

SIN_CUADRICULA = [dashboard.DASH, analisis.TD, etl.CONTROL, parametros.HOJA,
                  documentacion.HOJA]
SIN_ENCABEZADOS = [dashboard.DASH]

# formatos de las etiquetas de datos: en miles, para que entren en el grafico
MONEY_ETIQ = '"$"#,##0," k"'
MONEY_ETIQ_CERO = '"$"#,##0," k";;'   # esconde ceros y negativos


def _recalcular(ruta, segundos=900):
    out = subprocess.run([sys.executable, RECALC, ruta, str(segundos)],
                         capture_output=True, text=True,
                         timeout=segundos + 300)
    texto = out.stdout.strip() or out.stderr.strip()
    print(texto)
    try:
        res = json.loads(out.stdout)
    except json.JSONDecodeError:
        raise SystemExit("recalc.py no devolvio JSON")
    if "error" in res:
        raise SystemExit("recalc fallo: " + res["error"])
    if res.get("total_errors"):
        raise SystemExit(f"hay {res['total_errors']} formulas con error")
    return res


def main(destino):
    print("1/5  construyendo el libro base...")
    ruta, columnas, filas, pivots, crudas, informe = construir.main()

    print("2/5  inyectando tablas dinamicas...")
    campos_slicer = [c for c, *_ in SEGMENTADORES]
    ooxml.inject_pivots(ruta, etl.LIMPIOS, columnas, filas, pivots,
                        slicer_fields=campos_slicer)

    print("3/5  recalculando con LibreOffice...")
    _recalcular(ruta)

    # copia previa al post-proceso: permite reintentar los ultimos pasos sin
    # volver a pagar el recalculo
    shutil.copy(ruta, os.path.join(AQUI, "post_recalc.xlsx"))

    # LibreOffice rotula las tablas dinamicas en ingles al materializarlas
    t = ooxml.traducir_cadenas(ruta, analisis.TD,
                               {"Data": "Valores",
                                "Total Result": "Total general"})
    print(f"     rotulos de tabla dinamica traducidos: {t}")

    print("4/5  verificando el ETL contra los valores esperados...")
    import verificar
    verificar.verificar(ruta, columnas, filas, crudas, informe)

    print("5/5  titulos, etiquetas de datos y segmentadores...")
    # cada grafico se reconoce por un rango propio de la hoja Calculos
    n = ooxml.titulos_dinamicos(ruta, [
        (f"$A${analisis.R_MES}:", analisis.T["g1"]),
        (f"$J${analisis.R_MES}:", analisis.T["g2"]),
        (f"$A${analisis.R_CASC}:", analisis.T["g3"]),
        (f"$A${analisis.R_CAT}:", analisis.T["g4"]),
        ("<c:scatterChart>", analisis.T["g5"]),
    ])
    print(f"     titulos dinamicos aplicados: {n}/5")
    if n != 5:
        raise SystemExit(f"se esperaban 5 titulos dinamicos, se aplicaron {n}")

    # las etiquetas de datos se reinyectan porque LibreOffice las reescribe
    # apagadas al guardar
    e = ooxml.etiquetas_de_datos(ruta, [
        # (marcador, series, formato, posicion, puntos sin etiqueta)
        # 1. mensual: solo la linea de margen %, que es la metrica que importa
        (f"$A${analisis.R_MES}:", [2], "0.0%", "t"),
        # 2. acumulado: solo el cierre del ano; 24 etiquetas serian ilegibles
        (f"$J${analisis.R_MES}:", None, MONEY_ETIQ, "t", range(11)),
        # 3. cascada: los tres tramos visibles, con los ceros ocultos
        (f"$A${analisis.R_CASC}:", [1, 2, 3], MONEY_ETIQ_CERO, "ctr"),
        # 4. categorias: las ocho barras entran comodas
        (f"$A${analisis.R_CAT}:", None, "0.0%", "outEnd"),
    ])
    print(f"     series con etiquetas de datos: {e}")
    if e != 8:
        raise SystemExit(f"se esperaban 8 series etiquetadas, hubo {e}")

    # la serie base de la cascada es invisible: fuera de la leyenda
    b = ooxml.borrar_leyenda(ruta, [(f"$A${analisis.R_CASC}:", [0])])
    print(f"     entradas de leyenda ocultadas: {b}")
    if b != 1:
        raise SystemExit(f"no se pudo ocultar la serie base ({b})")

    slicers = [ooxml.Slicer(campo, titulo, c1, 14, c2, 23, column_count=ncol)
               for campo, titulo, c1, c2, ncol in SEGMENTADORES]
    ooxml.finalize(ruta, slicers, dashboard.DASH,
                   gridless_sheets=SIN_CUADRICULA,
                   headerless_sheets=SIN_ENCABEZADOS)

    os.makedirs(os.path.dirname(os.path.abspath(destino)), exist_ok=True)
    shutil.copy(ruta, destino)
    tam = os.path.getsize(destino) / 1024 / 1024
    print(f"listo -> {destino}  ({tam:.1f} MB)")
    return destino


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1
         else os.path.join(AQUI, "..",
              "PracticaFinal_TechStore_Murphy_Lleyton.xlsx"))
