"""
Arma el entregable completo de punta a punta.

    1. fuentes.py    -> ventas_historico.csv y maestros.xlsx
    2. construir.py  -> el libro con todas las hojas
    3. ooxml.py      -> valores cacheados, escenarios y recalculo al abrir
    4. mashup.py     -> las consultas de Power Query embebidas
    5. verificacion  -> se relee el archivo terminado y se controla todo

El orden importa: el DataMashup va ultimo porque cualquier reescritura del zip
posterior tendria que preservarlo, y es la parte mas delicada del archivo.

    python generador/pipeline.py

Dependencias: openpyxl. Nada mas.
"""
import os
import sys

import construir
import consultas_m
import fuentes
import mashup
import medidas_dax
import ooxml
import simulacion

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(AQUI)

COMENTARIOS = {
    "Pesimista": "Se corta el ciclo de reposicion, sube el costo de envio y se "
                 "encarece el capital de trabajo",
    "Base": "Plan conservador: continua la tendencia de 2025 sin extrapolar el rebote",
    "Optimista": "Se sostiene la demanda y baja el costo logistico por mejor "
                 "umbral de envio",
}


def main():
    print("1/5  generando archivos fuente...")
    fuentes.main()
    consultas_m.escribir(os.path.join(RAIZ, "powerquery", "Section1.m"))
    medidas_dax.escribir(os.path.join(RAIZ, "dax", "medidas.dax"))

    print("2/5  construyendo el libro...")
    info = construir.main()
    ruta = info["ruta"]

    print("3/5  cacheando valores y escenarios...")
    esperadas = sum(len(v) for v in info["cache"].values())
    aplicadas = ooxml.cachear(ruta, info["cache"])
    if aplicadas != esperadas:
        raise SystemExit(f"se esperaban {esperadas} formulas cacheadas, "
                         f"se aplicaron {aplicadas}")
    print(f"     {aplicadas} formulas con valor cacheado")

    celdas = info["escenarios"]["celdas"]      # [(nombre, "C29", formato), ...]
    valores = info["escenarios"]["valores"]    # {escenario: (v1, v2, v3)}
    refs = [ref for _, ref, _ in celdas]
    escenarios = [
        (nombre, COMENTARIOS[nombre],
         [(ref, repr(valor)) for ref, valor in zip(refs, valores[nombre])])
        for nombre, _ in simulacion.ESCENARIOS
    ]
    sqref = f"{refs[0]}:{refs[-1]}"
    n = ooxml.inyectar_escenarios(ruta, "Simulacion", escenarios, sqref, actual=1)
    print(f"     {n} escenarios inyectados sobre {sqref}")

    ooxml.calcular_al_abrir(ruta)

    print("4/5  embebiendo las consultas de Power Query...")
    tamano = mashup.inyectar(ruta, consultas_m.SECCION, consultas_m.CONSULTAS)
    print(f"     DataMashup de {tamano:,} bytes")

    print("5/5  verificando el archivo terminado...")
    problemas = verificar(ruta, info)
    if problemas:
        for p in problemas:
            print("  FALLA:", p)
        raise SystemExit(f"{len(problemas)} verificaciones fallaron")

    print("listo ->", ruta, f"({os.path.getsize(ruta) / 1024:.0f} KB)")
    return ruta


HOJAS_ESPERADAS = [
    "Inicio", "Parametros", "Modelo_Estrella", "Medidas_DAX", "Codigo_M",
    "Estadistica_Descriptiva", "Simulacion", "Resumen de escenarios",
    "Hechos_Ventas", "Dim_Clientes", "Dim_Productos", "Dim_Calendario",
]


def verificar(ruta, info):
    """
    Controla el archivo terminado contra la lista de la consigna.

    Devuelve la lista de problemas; si vuelve vacia, el entregable cumple todo
    lo que se puede verificar sin abrir Excel.
    """
    problemas = []
    r = ooxml.resumen(ruta)

    faltan = [h for h in HOJAS_ESPERADAS if h not in r["hojas"]]
    if faltan:
        problemas.append(f"faltan hojas: {faltan}")

    if r["formulas"] != r["formulas_cacheadas"]:
        problemas.append(f"{r['formulas'] - r['formulas_cacheadas']} formulas "
                         f"quedaron sin valor cacheado")

    if r["escenarios"].get("Simulacion") != 3:
        problemas.append(f"escenarios en Simulacion: {r['escenarios']}")

    # 1. consultas de Power Query embebidas, con pasos renombrados y comentarios
    try:
        dm = mashup.verificar(ruta)
    except Exception as exc:                       # noqa: BLE001
        problemas.append(f"no se pudo releer el DataMashup: {exc}")
    else:
        seccion = dm["seccion"]
        if seccion != consultas_m.SECCION:
            problemas.append("el Section1.m embebido no coincide con el original")
        for consulta, _ in consultas_m.CONSULTAS:
            if f"shared {consulta} =" not in seccion:
                problemas.append(f"falta la consulta {consulta} en Section1.m")
        for paso in ("VentasFiltradas", "SinDuplicados", "TiposAjustados",
                     "ClavesUnicas", "EncabezadosPromovidos"):
            if paso not in seccion:
                problemas.append(f"falta el paso renombrado {paso}")
        if seccion.count("//") < 20:
            problemas.append("el codigo M tiene pocos comentarios")
        if "C:\\" in seccion or "/home/" in seccion:
            problemas.append("hay una ruta hardcodeada en el codigo M")

    # 2. medidas DAX con VAR / RETURN
    con_var = [n for n, codigo, _ in medidas_dax.MEDIDAS
               if "VAR " in codigo and "RETURN" in codigo]
    if len(con_var) < 1:
        problemas.append("ninguna medida usa VAR/RETURN")

    # 3. el bloque de Buscar Objetivo tiene que cerrar en la meta exacta
    objetivo = info["objetivo"]
    alcanzado = (objetivo["unidades"] * objetivo["contribucion_unitaria"]
                 - objetivo["gastos_fijos"])
    if abs(alcanzado - simulacion.META_BENEFICIO) > 0.01:
        problemas.append(f"Buscar Objetivo no cierra: {alcanzado:.2f} contra "
                         f"{simulacion.META_BENEFICIO:.2f}")

    # 4. los tres escenarios tienen que dar resultados distintos y ordenados
    beneficios = [info["proyecciones"][n]["beneficio"]
                  for n in ("Pesimista", "Base", "Optimista")]
    if not beneficios[0] < beneficios[1] < beneficios[2]:
        problemas.append(f"los escenarios no estan ordenados: {beneficios}")

    # 5. las fuentes tienen que estar donde las busca la consulta
    for archivo in ("ventas_historico.csv", "maestros.xlsx"):
        if not os.path.exists(os.path.join(RAIZ, "fuentes", archivo)):
            problemas.append(f"falta el archivo fuente {archivo}")

    return problemas


if __name__ == "__main__":
    sys.exit(0 if main() else 1)
