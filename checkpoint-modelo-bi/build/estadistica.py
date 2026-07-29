"""Replica exactamente el reporte que emite Analysis ToolPak > Estadistica
descriptiva, con las mismas etiquetas, el mismo orden y los mismos estimadores
que usa Excel (muestrales, no poblacionales)."""
import csv
import math
from datetime import datetime
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent


def cuantil_t(p, gl):
    """Cuantil de la t de Student por expansion de Cornish-Fisher.

    Con los grados de libertad de este dataset (>1000) el error frente al valor
    exacto esta en el sexto decimal, muy por debajo de lo que muestra Excel.
    """
    z = cuantil_normal(p)
    g1 = (z ** 3 + z) / 4
    g2 = (5 * z ** 5 + 16 * z ** 3 + 3 * z) / 96
    g3 = (3 * z ** 7 + 19 * z ** 5 + 17 * z ** 3 - 15 * z) / 384
    return z + g1 / gl + g2 / gl ** 2 + g3 / gl ** 3


def cuantil_normal(p):
    """Inversa de la normal estandar (algoritmo de Acklam)."""
    a = [-3.969683028665376e+01, 2.209460984245205e+02, -2.759285104469687e+02,
         1.383577518672690e+02, -3.066479806614716e+01, 2.506628277459239e+00]
    b = [-5.447609879822406e+01, 1.615858368580409e+02, -1.556989798598866e+02,
         6.680131188771972e+01, -1.328068155288572e+01]
    c = [-7.784894002430293e-03, -3.223964580411365e-01, -2.400758277161838e+00,
         -2.549732539343734e+00, 4.374664141464968e+00, 2.938163982698783e+00]
    d = [7.784695709041462e-03, 3.224671290700398e-01, 2.445134137142996e+00,
         3.754408661907416e+00]
    p_bajo, p_alto = 0.02425, 1 - 0.02425

    if p < p_bajo:
        q = math.sqrt(-2 * math.log(p))
        return (((((c[0]*q+c[1])*q+c[2])*q+c[3])*q+c[4])*q+c[5]) / \
               ((((d[0]*q+d[1])*q+d[2])*q+d[3])*q+1)
    if p > p_alto:
        q = math.sqrt(-2 * math.log(1 - p))
        return -(((((c[0]*q+c[1])*q+c[2])*q+c[3])*q+c[4])*q+c[5]) / \
                ((((d[0]*q+d[1])*q+d[2])*q+d[3])*q+1)
    q = p - 0.5
    r = q * q
    return (((((a[0]*r+a[1])*r+a[2])*r+a[3])*r+a[4])*r+a[5])*q / \
           (((((b[0]*r+b[1])*r+b[2])*r+b[3])*r+b[4])*r+1)


def descriptivos(valores, confianza=0.95):
    n = len(valores)
    ordenados = sorted(valores)
    media = sum(valores) / n

    desvios = [v - media for v in valores]
    varianza = sum(d ** 2 for d in desvios) / (n - 1)
    desvio = math.sqrt(varianza)
    error_tipico = desvio / math.sqrt(n)

    if n % 2:
        mediana = ordenados[n // 2]
    else:
        mediana = (ordenados[n // 2 - 1] + ordenados[n // 2]) / 2

    frecuencias = {}
    for v in valores:
        frecuencias[v] = frecuencias.get(v, 0) + 1
    tope = max(frecuencias.values())
    moda = min(v for v, f in frecuencias.items() if f == tope) if tope > 1 else "#N/A"

    # Curtosis y asimetria con las formulas exactas de Excel (KURT y COEFICIENTE.ASIMETRIA)
    suma_cubos = sum((d / desvio) ** 3 for d in desvios)
    asimetria = n / ((n - 1) * (n - 2)) * suma_cubos

    suma_cuartas = sum((d / desvio) ** 4 for d in desvios)
    curtosis = (n * (n + 1) / ((n - 1) * (n - 2) * (n - 3))) * suma_cuartas \
        - 3 * (n - 1) ** 2 / ((n - 2) * (n - 3))

    t_critico = cuantil_t(1 - (1 - confianza) / 2, n - 1)

    return [
        ("Media", media),
        ("Error tipico", error_tipico),
        ("Mediana", mediana),
        ("Moda", moda),
        ("Desviacion estandar", desvio),
        ("Varianza de la muestra", varianza),
        ("Curtosis", curtosis),
        ("Coeficiente de asimetria", asimetria),
        ("Rango", ordenados[-1] - ordenados[0]),
        ("Minimo", ordenados[0]),
        ("Maximo", ordenados[-1]),
        ("Suma", sum(valores)),
        ("Cuenta", n),
        (f"Nivel de confianza({confianza * 100:.1f}%)", t_critico * error_tipico),
    ]


def ventas_depuradas():
    """Aplica en Python la MISMA depuracion que hace la consulta Ventas en M,
    para que la estadistica describa exactamente las filas que entran al modelo."""
    ruta = RAIZ / "datos" / "ventas_historicas.csv"
    vistos = set()
    filas = []

    with ruta.open(encoding="utf-8") as fh:
        for fila in csv.DictReader(fh, delimiter=";"):
            if fila["Estado"] != "Confirmada":
                continue
            if not fila["Unidades"] or not fila["ID_Cliente"] or not fila["ID_Producto"]:
                continue
            try:
                unidades = int(fila["Unidades"])
                precio = float(fila["Precio_Unitario"])
                costo = float(fila["Costo_Unitario"])
                descuento = float(fila["Descuento"])
                fecha = datetime.strptime(fila["Fecha"], "%d/%m/%Y").date()
            except ValueError:
                continue
            if fila["ID_Venta"] in vistos:
                continue
            vistos.add(fila["ID_Venta"])

            bruto = unidades * precio
            filas.append({
                "ID_Venta": fila["ID_Venta"],
                "Fecha": fecha,
                "Canal": fila["Canal"].strip().capitalize(),
                "Unidades": unidades,
                "Importe_Bruto": bruto,
                "Importe_Neto": round(bruto * (1 - descuento), 2),
                "Costo_Total": round(unidades * costo, 2),
            })
    return filas


if __name__ == "__main__":
    filas = ventas_depuradas()
    print(f"Filas depuradas: {len(filas)}")
    for etiqueta, valor in descriptivos([f["Importe_Neto"] for f in filas]):
        print(f"  {etiqueta:<28} {valor}")
