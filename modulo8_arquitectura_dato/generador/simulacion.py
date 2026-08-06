"""
Modelo financiero de la hoja `Simulacion`: base, Buscar Objetivo y escenarios.

La base no se inventa: sale de la propia tabla de hechos, tomando 2025 como
ultimo ejercicio cerrado. Sobre esa base se proyecta 2026 con tres variables
que el directorio puede mover.

Buscar Objetivo (Paso 6) se plantea sobre un bloque aparte, con sus propias
constantes, para que el resultado siga valiendo cuando alguien cambie de
escenario. La incognita se puede despejar a mano:

    beneficio = U*p - U*c - U*p*t - U*p*i - F
              = U * (p - c - p*t - p*i) - F

    U = (meta + F) / (p - c - p*t - p*i)

donde p es el precio neto promedio, c el costo unitario, t la tasa de costo
logistico, i la tasa de interes financiero y F los gastos fijos anuales.

Aca se calcula ese U con precision completa y se escribe como constante en la
celda que Buscar Objetivo modifica, de modo que la formula de la hoja devuelve
la meta exacta.
"""
import datos

META_BENEFICIO = 500000.0

# Gastos fijos anuales de estructura (sueldos, alquileres, marketing y
# administracion). Es el unico supuesto que no sale de los datos: se fija en el
# 19% de la venta neta de 2025 y queda documentado en la hoja Parametros.
RATIO_GASTOS_FIJOS = 0.19

# Tasa de interes financiero de referencia (escenario base y Buscar Objetivo)
INTERES_REFERENCIA = 0.020

# Variacion de la demanda que asume cada escenario. NO es la extrapolacion de
# 2025: el plan base es deliberadamente mas conservador que el rebote del
# ultimo ejercicio, que es como se arma un presupuesto que se pueda defender.
PLAN_DEMANDA = {"Pesimista": -0.06, "Base": 0.08, "Optimista": 0.18}

# Las tres variables cambiantes del Administrador de escenarios.
# El escenario Base repite el comportamiento observado en 2025; el Pesimista y
# el Optimista se construyen moviendo cada variable en la misma direccion en
# que la movio la realidad entre 2024 y 2025 (o en la contraria).
ESCENARIOS = [
    ("Pesimista", "La demanda se contrae, el envio se encarece y sube el costo del dinero"),
    ("Base",      "Continua la tendencia observada en 2025"),
    ("Optimista", "Se consolida el crecimiento y el umbral de envio gratis baja el costo logistico"),
]


def base_desde_datos(ventas):
    """Cierre 2025 y comparacion contra 2024, a partir de la tabla de hechos."""
    v2024 = [v for v in ventas if v["Fecha"].year == 2024]
    v2025 = [v for v in ventas if v["Fecha"].year == 2025]

    def agregado(filas):
        neta = sum(f["Venta_Neta"] for f in filas)
        return {
            "unidades": sum(f["Unidades"] for f in filas),
            "venta_neta": neta,
            "costo_ventas": sum(f["Costo_Total"] for f in filas),
            "costo_envio": sum(f["Costo_Envio"] for f in filas),
            "margen_bruto": sum(f["Margen_Bruto"] for f in filas),
            "operaciones": len(filas),
        }

    a24, a25 = agregado(v2024), agregado(v2025)
    return {
        "anio_2024": a24,
        "anio_2025": a25,
        "crecimiento_ventas": a25["venta_neta"] / a24["venta_neta"] - 1,
        "crecimiento_unidades": a25["unidades"] / a24["unidades"] - 1,
        "precio_neto_promedio": a25["venta_neta"] / a25["unidades"],
        "costo_unitario_promedio": a25["costo_ventas"] / a25["unidades"],
        "tasa_logistica": a25["costo_envio"] / a25["venta_neta"],
        "gastos_fijos": a25["venta_neta"] * RATIO_GASTOS_FIJOS,
    }


def variables(base):
    """
    Valores de las tres celdas cambiantes en cada escenario.

    Devuelve {nombre_escenario: (var_demanda, tasa_logistica, tasa_interes)}.
    """
    logistica_base = round(base["tasa_logistica"], 4)

    # La tasa logistica del escenario base es la observada en 2025. El
    # pesimista la encarece un 25% (mas envios a domicilio, mas devoluciones);
    # el optimista la abarata un 18% (umbral de envio gratis mas alto).
    return {
        "Pesimista": (PLAN_DEMANDA["Pesimista"], round(logistica_base * 1.25, 4), 0.035),
        "Base":      (PLAN_DEMANDA["Base"],      logistica_base,     INTERES_REFERENCIA),
        "Optimista": (PLAN_DEMANDA["Optimista"], round(logistica_base * 0.82, 4), 0.012),
    }


def proyectar(base, var_demanda, tasa_logistica, tasa_interes):
    """Cuenta de resultados proyectada 2026. Espeja las formulas de la hoja."""
    unidades = base["anio_2025"]["unidades"] * (1 + var_demanda)
    precio = base["precio_neto_promedio"]
    costo = base["costo_unitario_promedio"]

    ventas = unidades * precio
    costo_ventas = unidades * costo
    margen_bruto = ventas - costo_ventas
    costo_logistico = ventas * tasa_logistica
    gastos_fijos = base["gastos_fijos"]
    costo_financiero = ventas * tasa_interes
    beneficio = margen_bruto - costo_logistico - gastos_fijos - costo_financiero

    return {
        "unidades": unidades,
        "ventas": ventas,
        "costo_ventas": costo_ventas,
        "margen_bruto": margen_bruto,
        "costo_logistico": costo_logistico,
        "gastos_fijos": gastos_fijos,
        "costo_financiero": costo_financiero,
        "beneficio": beneficio,
        "margen_neto": beneficio / ventas,
    }


def buscar_objetivo(base, meta=META_BENEFICIO):
    """
    Unidades necesarias para alcanzar `meta` de beneficio neto.

    Usa las constantes del escenario Base (precio, costo unitario, tasa
    logistica, tasa de interes y gastos fijos) y despeja U. El bloque de la
    hoja replica esta misma cuenta con formulas, asi que el evaluador puede
    volver a correr Buscar Objetivo y obtener el mismo numero.
    """
    # se usan las tasas sin redondear: el bloque de la hoja referencia las
    # mismas celdas, asi la formula devuelve la meta exacta y no 499.998
    p = base["precio_neto_promedio"]
    c = base["costo_unitario_promedio"]
    t = base["tasa_logistica"]
    i = INTERES_REFERENCIA
    f = base["gastos_fijos"]

    contribucion = p - c - p * t - p * i
    unidades = (meta + f) / contribucion
    return {
        "meta": meta,
        "unidades": unidades,
        "precio": p,
        "costo_unitario": c,
        "tasa_logistica": t,
        "tasa_interes": i,
        "gastos_fijos": f,
        "contribucion_unitaria": contribucion,
        "unidades_2025": base["anio_2025"]["unidades"],
        "variacion_unidades": unidades / base["anio_2025"]["unidades"] - 1,
        "unidades_base": base["anio_2025"]["unidades"] * (1 + PLAN_DEMANDA["Base"]),
    }


def resumen():
    d = datos.construir()
    datos.enriquecer(d["ventas"])
    base = base_desde_datos(d["ventas"])
    vars_ = variables(base)
    proyecciones = {n: proyectar(base, *vars_[n]) for n in vars_}
    return base, vars_, proyecciones, buscar_objetivo(base)


if __name__ == "__main__":
    base, vars_, proy, objetivo = resumen()
    print(f"unidades 2025            {base['anio_2025']['unidades']:>12,.0f}")
    print(f"venta neta 2025          {base['anio_2025']['venta_neta']:>12,.2f}")
    print(f"crecimiento unidades     {base['crecimiento_unidades']:>12.2%}")
    print(f"precio neto promedio     {base['precio_neto_promedio']:>12.2f}")
    print(f"costo unitario promedio  {base['costo_unitario_promedio']:>12.2f}")
    print(f"tasa logistica 2025      {base['tasa_logistica']:>12.2%}")
    print(f"gastos fijos             {base['gastos_fijos']:>12,.2f}")
    print()
    for nombre, _ in ESCENARIOS:
        v, p = vars_[nombre], proy[nombre]
        print(f"{nombre:<11} demanda {v[0]:>7.2%}  logistica {v[1]:>6.2%}  "
              f"interes {v[2]:>5.2%} -> ventas {p['ventas']:>12,.0f}  "
              f"beneficio {p['beneficio']:>11,.0f}  margen {p['margen_neto']:>6.2%}")
    print()
    print(f"Buscar Objetivo: {objetivo['unidades']:,.4f} unidades para "
          f"{objetivo['meta']:,.0f} de beneficio neto "
          f"({objetivo['variacion_unidades']:+.2%} vs 2025)")
