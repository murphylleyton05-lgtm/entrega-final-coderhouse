# Modulo 8 — Arquitectura del dato, optimizacion de modelos y simulacion de escenarios

Pre-entrega **Optimizacion y Modelado de Datos para Toma de Decisiones**, del
Sistema de Inteligencia de Negocios de la cadena de tecnologia **TechStore**.
Es el checkpoint donde el proyecto deja de ser una tabla plana y pasa a ser un
modelo: ETL con Power Query, esquema en estrella para Power Pivot, medidas DAX
optimizadas, diagnostico estadistico y capa de simulacion financiera.

## Que se entrega

| Archivo | Que es |
|---|---|
| [`PreEntrega_ModeloBI_MurphyLleyton.xlsx`](./PreEntrega_ModeloBI_MurphyLleyton.xlsx) | **El entregable.** Un unico libro con las consultas, el modelo, la estadistica y los escenarios |
| [`fuentes/ventas_historico.csv`](./fuentes) | Archivo transaccional historico (5.290 filas crudas) |
| [`fuentes/maestros.xlsx`](./fuentes) | Libro local con las tablas maestras `tbl_Clientes` y `tbl_Productos` |
| [`powerquery/Section1.m`](./powerquery/Section1.m) | El codigo M completo, en texto plano |
| [`dax/medidas.dax`](./dax/medidas.dax) | Las 8 medidas DAX, listas para pegar |
| [`generador/`](./generador) | Los scripts que arman todo de cero |

> **Descargar la carpeta entera, no solo el `.xlsx`.** Las consultas leen los
> archivos de `fuentes/`, que tienen que quedar al lado del libro.

## Las 12 hojas del libro

| Hoja | Rol |
|---|---|
| `Inicio` | Que es el archivo, indice y el paso final que se completa en Excel |
| `Parametros` | Ruta de las fuentes y todos los supuestos del modelo |
| `Modelo_Estrella` | Diagrama del esquema y la especificacion de las 3 relaciones 1:N |
| `Medidas_DAX` | Las 8 medidas, con el antes y despues de la optimizacion |
| `Codigo_M` | El `Section1.m` completo, legible sin abrir Power Query |
| `Estadistica_Descriptiva` | Salida del Analysis ToolPak + verificacion con formulas |
| `Simulacion` | Cuenta de resultados 2026, Buscar Objetivo y celdas cambiantes |
| `Resumen de escenarios` | Pesimista / Base / Optimista, uno al lado del otro |
| `Hechos_Ventas` | Tabla de hechos: 5.200 operaciones limpias, 2024-2025 |
| `Dim_Clientes` · `Dim_Productos` · `Dim_Calendario` | Las tres dimensiones |

## Como se cumple cada punto de la consigna

| Requisito | Donde se cumple |
|---|---|
| Conexion a fuentes heterogeneas (`.csv` + `.xlsx` local) | `Hechos_Ventas` lee el CSV; `Dim_Clientes` y `Dim_Productos` leen las tablas del libro de maestros |
| Limpieza de nulos, tipado y filtrado estrategico | El CSV trae 5.290 filas: se descartan 45 anuladas, 30 con nulos y 15 duplicadas. Quedan 5.200 |
| Orden eficiente de los pasos aplicados | Se limpia el texto, se filtra, se deduplica y **recien despues** se tipa: no se convierten filas que igual se van a descartar |
| Codigo M editado a mano en el Editor avanzado | Pasos con nombre de negocio (`VentasFiltradas`, `SinDuplicados`, `TiposAjustados`, `ClavesUnicas`) y comentarios `//` en cada bloque |
| Correccion de rutas hardcodeadas | Ninguna consulta tiene la ruta escrita adentro: la piden a `pRutaFuentes`, que la lee de la celda con nombre `RutaFuentes` |
| Carga al modelo de datos, no a hoja | Las consultas van como "Crear solo conexion" + "Agregar al Modelo de datos" (hoja `Inicio`, paso 1) |
| Esquema en estrella con relaciones 1:N | Hoja `Modelo_Estrella`: hechos al centro, 3 dimensiones alrededor, filtro fluyendo hacia los hechos |
| Medidas DAX con `VAR` / `RETURN` | 4 de las 8 medidas: `% Margen Bruto`, `Ticket Promedio`, `Crecimiento Interanual %`, `% Costo Logistico s/Ventas` |
| Reduccion del tamano del modelo | Aritmetica de fila resuelta en M, columna constante `Estado` descartada, cero `SUMX`/`FILTER` sobre la tabla completa |
| Estadistica descriptiva (Analysis ToolPak) | Hoja `Estadistica_Descriptiva`: media, mediana, desvio y rango sobre `Venta_Neta` y `Margen_Bruto` |
| Buscar Objetivo | Bloque 4 de la hoja `Simulacion`: cuantas unidades hacen falta para USD 500.000 de beneficio neto |
| Administrador de escenarios (3 escenarios) | Los tres escenarios estan guardados en el libro: Datos > Analisis de hipotesis > Administrador de escenarios |
| Resumen de escenarios en pestana propia | Hoja `Resumen de escenarios` |

## El paso que hay que dar dentro de Excel

El modelo tabular de Power Pivot es una base en memoria que **solo Excel puede
escribir**: ninguna herramienta externa puede generar ese componente. Todo lo
demas viaja dentro del archivo — las consultas M embebidas, los datos, las
relaciones especificadas, las medidas escritas, los escenarios guardados, el
resumen y la estadistica.

Faltan tres clics, documentados con detalle en la hoja `Inicio`:

1. **Cargar al modelo.** Datos > Consultas y conexiones. Sobre cada consulta:
   clic derecho > Cargar en... > *Crear solo conexion* + *Agregar estos datos
   al Modelo de datos*.
2. **Crear las relaciones.** Power Pivot > Administrar > Vista de diagrama.
   Arrastrar cada PK sobre su FK segun la tabla de `Modelo_Estrella`, y marcar
   `Dim_Calendario` como Tabla de fechas.
3. **Pegar las medidas.** Area de calculo de `Hechos_Ventas`, una medida por
   celda, desde `dax/medidas.dax`.

Si aparece el error `Formula.Firewall`, es porque `pRutaFuentes` lee una celda
del libro y las demas consultas leen archivos: Datos > Obtener datos > Opciones
de consulta > Privacidad > *Omitir siempre la configuracion de niveles de
privacidad*.

Las consultas viajan embebidas en el libro como parte `DataMashup`, que es el
formato que usa Excel para guardarlas ([MS-QDEFF]). El mismo script esta ademas
en [`powerquery/Section1.m`](./powerquery/Section1.m) y en la hoja `Codigo_M`:
si hiciera falta reconstruirlas, se pega el contenido en Power Query > Nueva
consulta > Consulta en blanco > Editor avanzado.

## El modelo

```
                      Dim_Calendario
                       Fecha (PK)
                            |
                          1 : N
                            v
   Dim_Clientes  --1:N-->  Hechos_Ventas  <--1:N--  Dim_Productos
   ID_Cliente (PK)          ID_Venta                 ID_Producto (PK)
                            Fecha        (FK)
                            ID_Cliente   (FK)
                            ID_Producto  (FK)
                            Unidades, Precio_Unitario, Descuento,
                            Costo_Unitario, Costo_Envio,
                            Venta_Bruta, Venta_Neta,
                            Costo_Total, Margen_Bruto
```

Las tres relaciones son de direccion unica y ninguna dimension se relaciona con
otra: eso es lo que distingue una estrella de un copo de nieve. Cada consulta de
dimension termina en `Table.Distinct` sobre su clave, que es lo que garantiza el
lado "1" de la relacion — sin ese paso Power Pivot rechaza el 1:N.

## Que dicen los numeros

**Diagnostico estadistico** (sobre `Venta_Neta`, 5.200 operaciones):

| Estadistico | Valor |
|---|---|
| Media | USD 915,10 |
| Mediana | USD 279,68 |
| Desviacion estandar | USD 2.375,18 |
| Rango | USD 79.734,05 (min 23,54 · max 79.757,59) |
| Coeficiente de asimetria | 10,95 |
| Curtosis | 258,21 |

La media es **3,3 veces la mediana**: la distribucion esta fuertemente sesgada a
la derecha y unas pocas operaciones corporativas muy grandes estiran la cola. El
desvio estandar supera a la media (coeficiente de variacion del 260%). La
conclusion operativa es directa: proyectar 2026 con un unico valor esperado
seria enganoso, y por eso el paso siguiente es una simulacion por escenarios.

**Simulacion 2026** (base: cierre 2025 con 13.310 unidades y USD 2.643.483 de
venta neta, +25,0% sobre 2024):

| | Pesimista | Base | Optimista |
|---|---|---|---|
| Variacion de la demanda | −6,0% | +8,0% | +18,0% |
| Tasa de costo logistico | 2,25% | 1,80% | 1,48% |
| Tasa de interes financiero | 3,5% | 2,0% | 1,2% |
| **Venta neta** | 2.484.874 | 2.854.961 | 3.119.310 |
| **Margen bruto** | 902.644 | 1.037.080 | 1.133.106 |
| **Beneficio neto** | **257.502** | **426.330** | **547.247** |
| Margen neto | 10,36% | 14,93% | 17,54% |

**Buscar Objetivo:** para cerrar 2026 con un beneficio neto de USD 500.000 hay
que vender **15.516 unidades**, un +16,6% sobre 2025 y un +7,9% por encima del
plan base. La meta cae entre el escenario Base y el Optimista: es alcanzable,
pero no con el plan conservador.

## Sobre los datos

El dataset es **simulado y determinista** (semilla fija). Extiende el modelo
`Ventas_Tech_DB` del [checkpoint de SQL](../checkpoint_sql_ventas_tech) —mismas
categorias, productos y ciudades— a 24 meses, 5 regiones y 2 canales.

El CSV de origen sale **sucio a proposito**: 45 operaciones anuladas, 30 filas
con `Unidades` o `ID_Cliente` vacios, 15 duplicadas exactas y numeros guardados
como texto con espacios sobrantes. Sin suciedad, el paso de limpieza de Power
Query no tendria nada que demostrar.

## Reproducir el entregable

La carpeta [`generador/`](./generador) arma todo de cero: los archivos fuente,
el libro, los valores cacheados, los escenarios y las consultas embebidas.

```bash
pip install openpyxl
python generador/pipeline.py
```

El pipeline termina con una bateria de verificaciones sobre el archivo ya
escrito: que esten las 12 hojas, que ninguna formula haya quedado sin valor
cacheado, que los 3 escenarios esten en el XML de la hoja, que el `Section1.m`
embebido sea identico al original, que existan los pasos renombrados, que no
haya quedado ninguna ruta hardcodeada, que Buscar Objetivo cierre exactamente en
la meta y que los tres escenarios den resultados ordenados.

| Script | Que hace |
|---|---|
| `datos.py` | Genera el dataset simulado y las dimensiones |
| `fuentes.py` | Escribe el CSV y el libro de maestros |
| `consultas_m.py` | El codigo M del flujo ETL |
| `medidas_dax.py` | Las medidas DAX |
| `estadistica.py` | Los 14 estadisticos del ToolPak, con las definiciones de Excel |
| `simulacion.py` | Modelo financiero, escenarios y Buscar Objetivo |
| `construir.py` | Arma el libro con openpyxl |
| `ooxml.py` | Valores cacheados, escenarios y recalculo al abrir |
| `mashup.py` | Empaqueta el codigo M como parte DataMashup ([MS-QDEFF]) |
| `pipeline.py` | Orquesta y verifica |

---

**Lleyton Murphy** · Coderhouse, Data Analyst · Modulo 8
