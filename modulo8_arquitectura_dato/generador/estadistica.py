"""
Estadistica descriptiva con las mismas definiciones que usa Excel.

La hoja `Estadistica_Descriptiva` del libro reproduce exactamente la salida de
**Datos > Analisis de datos > Estadistica descriptiva** (Analysis ToolPak):
mismas 14 filas, mismo orden, mismos nombres. Para que los numeros coincidan
hasta el ultimo decimal hay que usar las definiciones de Excel, que no siempre
son las mas comunes:

  * Desviacion estandar y varianza son **muestrales** (divisor n-1).
  * Curtosis es la curtosis de exceso corregida por muestra, igual que KURT().
  * El coeficiente de asimetria es el G1 corregido, igual que COEFICIENTE.ASIMETRIA().
  * "Nivel de confianza (95,0%)" es CONFIANZA.T: t(0,025; n-1) * s / raiz(n),
    no el 1,96 de la normal.

El cuantil de la t de Student se calcula aca mismo (funcion beta incompleta
regularizada + biseccion) para no depender de scipy: el generador tiene que
poder correr con openpyxl y nada mas.
"""
import math
from collections import Counter

# nombres tal como los escribe el ToolPak en Excel en espanol
ETIQUETAS = [
    "Media",
    "Error tipico",
    "Mediana",
    "Moda",
    "Desviacion estandar",
    "Varianza de la muestra",
    "Curtosis",
    "Coeficiente de asimetria",
    "Rango",
    "Minimo",
    "Maximo",
    "Suma",
    "Cuenta",
    "Nivel de confianza (95,0%)",
]


# --- t de Student ------------------------------------------------------------

def _betacf(a, b, x, iteraciones=300, eps=1e-14):
    """Fraccion continua de Lentz para la beta incompleta."""
    qab, qap, qam = a + b, a + 1.0, a - 1.0
    c, d = 1.0, 1.0 - qab * x / qap
    if abs(d) < 1e-300:
        d = 1e-300
    d = 1.0 / d
    h = d
    for m in range(1, iteraciones + 1):
        m2 = 2 * m
        num = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1.0 + num * d
        if abs(d) < 1e-300:
            d = 1e-300
        c = 1.0 + num / c
        if abs(c) < 1e-300:
            c = 1e-300
        d = 1.0 / d
        h *= d * c
        num = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1.0 + num * d
        if abs(d) < 1e-300:
            d = 1e-300
        c = 1.0 + num / c
        if abs(c) < 1e-300:
            c = 1e-300
        d = 1.0 / d
        delta = d * c
        h *= delta
        if abs(delta - 1.0) < eps:
            break
    return h


def betainc(a, b, x):
    """Beta incompleta regularizada I_x(a, b)."""
    if x <= 0.0:
        return 0.0
    if x >= 1.0:
        return 1.0
    ln = (math.lgamma(a + b) - math.lgamma(a) - math.lgamma(b)
          + a * math.log(x) + b * math.log1p(-x))
    frente = math.exp(ln)
    if x < (a + 1.0) / (a + b + 2.0):
        return frente * _betacf(a, b, x) / a
    return 1.0 - frente * _betacf(b, a, 1.0 - x) / b


def t_cdf(t, gl):
    x = gl / (gl + t * t)
    cola = 0.5 * betainc(gl / 2.0, 0.5, x)
    return 1.0 - cola if t > 0 else cola


def t_ppf(p, gl):
    """Cuantil de la t de Student por biseccion sobre la CDF."""
    lo, hi = -1000.0, 1000.0
    for _ in range(200):
        medio = (lo + hi) / 2.0
        if t_cdf(medio, gl) < p:
            lo = medio
        else:
            hi = medio
    return (lo + hi) / 2.0


# --- estadisticos ------------------------------------------------------------

def _moda(valores):
    """MODA() de Excel: el valor mas frecuente; #N/D si ninguno se repite."""
    conteo = Counter(valores)
    valor, veces = conteo.most_common(1)[0]
    return valor if veces > 1 else "#N/D"


def describir(valores):
    """Devuelve las 14 medidas del ToolPak, en el orden de ETIQUETAS."""
    n = len(valores)
    if n < 4:
        raise ValueError("hacen falta al menos 4 observaciones")

    media = sum(valores) / n
    desvios = [v - media for v in valores]
    varianza = sum(d * d for d in desvios) / (n - 1)
    desvio = math.sqrt(varianza)
    error_tipico = desvio / math.sqrt(n)

    ordenados = sorted(valores)
    medio = n // 2
    mediana = (ordenados[medio] if n % 2
               else (ordenados[medio - 1] + ordenados[medio]) / 2)

    z3 = sum((d / desvio) ** 3 for d in desvios)
    asimetria = n / ((n - 1) * (n - 2)) * z3

    z4 = sum((d / desvio) ** 4 for d in desvios)
    curtosis = (n * (n + 1) / ((n - 1) * (n - 2) * (n - 3)) * z4
                - 3 * (n - 1) ** 2 / ((n - 2) * (n - 3)))

    confianza = t_ppf(0.975, n - 1) * error_tipico

    return dict(zip(ETIQUETAS, [
        media,
        error_tipico,
        mediana,
        _moda(valores),
        desvio,
        varianza,
        curtosis,
        asimetria,
        max(valores) - min(valores),
        min(valores),
        max(valores),
        sum(valores),
        n,
        confianza,
    ]))


if __name__ == "__main__":
    import datos
    d = datos.construir()
    datos.enriquecer(d["ventas"])
    for campo in ("Venta_Neta", "Margen_Bruto"):
        print("---", campo)
        for k, v in describir([v[campo] for v in d["ventas"]]).items():
            print(f"  {k:<28} {v}")
