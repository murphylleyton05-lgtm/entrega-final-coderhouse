# Practica Integradora Final — Coderhouse, Data Analyst

**Entregable:** [`PracticaFinal_TechStore_Murphy_Lleyton.xlsx`](./PracticaFinal_TechStore_Murphy_Lleyton.xlsx)
· 1,2 MB · un unico archivo `.xlsx`, sin macros.

Sistema de reporte de ventas y margen de la cadena **TechStore**. Toma un
export crudo del sistema transaccional —con duplicados, fechas en tres
formatos y categorias sin estandarizar—, lo depura dentro del mismo libro con
formulas a la vista, lo resume en tablas dinamicas y lo comunica en un tablero
ejecutivo con segmentadores.

> El libro se abre en la pestana **Dashboard**. **Requiere Excel 2010 o
> posterior**: los segmentadores son objetos nativos de Excel y otros visores
> (Google Sheets, LibreOffice, la vista previa de GitHub) los descartan al
> abrir.

---

## Como se cumple cada punto de la consigna

| # | Requisito | Donde se cumple |
|---|---|---|
| 1 | Datos originales limpios, sin duplicados y con formatos de celda correctos | `Datos_Origen` conserva el export crudo con los defectos resaltados; `Datos_Limpios` lo reconstruye con formulas (2.400 operaciones unicas, fechas reales con formato `dd/mm/aaaa`, unidades numericas); `Control_Calidad` lo verifica con 8 controles, todos en OK |
| 2 | Al menos una columna calculada con funciones logicas (SI, Y u O) | Tres: `Segmento_Rentabilidad` (SI + Y + O anidados), `Alerta_Logistica` (SI + Y) y `Ticket` (SI anidado). Los umbrales viven en `Parametros` |
| 3 | Al menos dos tablas dinamicas con metricas clave | Cuatro a la vista en `Tablas_Dinamicas` (Region, Categoria, Producto y Segmento de rentabilidad), cada una con Ventas, Margen y Unidades. Mas 13 de soporte que alimentan los graficos |
| 4 | Dashboard con al menos dos graficos coherentes y un segmentador funcional | Cinco graficos y **cuatro** segmentadores nativos (Trimestre, Categoria, Region, Canal). Los graficos leen las mismas tablas dinamicas via GETPIVOTDATA, asi que muestran exactamente los numeros de la hoja `Tablas_Dinamicas` |
| 5 | Diseno profesional, sin cuadricula y con etiquetas de datos claras | Tablero sin lineas de cuadricula ni encabezados de fila/columna, hojas de soporte ocultas, paleta de tres colores, 8 series con etiquetas de datos y titulos de grafico que se reescriben con los datos |

### Del nivel recomendado

| Requisito | Donde se cumple |
|---|---|
| Inteligencia de tiempo | Acumulado del ano (YTD) 2024 vs 2025 como grafico propio, variacion mes contra mes y comparacion interanual en las tarjetas |
| KPIs con desvio y cumplimiento | Cuatro tarjetas con valor, meta y brecha; la que no llega a la meta pasa entera a ambar por formato condicional |
| UX de alta fidelidad | Interfaz limpia, paleta de tres colores, secciones tituladas en el orden que se lee un reporte, sin tecnicismos de software en el tablero |
| Validaciones y listas desplegables | Bloque de consulta puntual en `Tablas_Dinamicas`, con tres listas de validacion sobre nombres definidos y criterios dinamicos en SUMAR.SI.CONJUNTO |
| Narrativa de datos | Titular y bajada dinamicos, y un reporte ejecutivo de tres lineas con la estructura **Evidencia → Insight** |

### Que no incluye, y por que

- **Power Query / Power Pivot / DAX.** La limpieza esta implementada con
  formulas nativas a la vista en `Datos_Limpios`, no con un paso de Power
  Query. La decision fue deliberada: cada transformacion queda auditable
  celda por celda, el libro se recalcula solo y no depende de un motor que el
  corrector tenga que abrir y refrescar para ver el trabajo. El detalle de
  cada paso del ETL, con las funciones usadas, esta en la hoja
  `Documentacion`.
- **Macros VBA.** El formato aceptado para la entrega es `.xlsx`, que por
  definicion no admite macros (eso requiere `.xlsm`). Toda la interactividad
  del tablero —segmentadores, formato condicional, refresco al abrir, listas
  de validacion— es nativa y no necesita que el usuario habilite nada.

---

## Hojas del libro

| Hoja | Rol | Visible |
|---|---|---|
| `Dashboard` | El entregable: titular, 4 tarjetas de KPI, 4 segmentadores, 5 graficos y el reporte ejecutivo | si |
| `Tablas_Dinamicas` | Las 4 tablas dinamicas del reporte + bloque de consulta con listas desplegables | si |
| `Datos_Limpios` | Tabla de hechos depurada: 2.400 operaciones, cada celda una formula | si |
| `Datos_Origen` | El export crudo sin tocar, con los defectos resaltados | si |
| `Control_Calidad` | Diagnostico, remediacion y reconciliacion, todo con formulas | si |
| `Parametros` | Metas, umbrales, tasas, diccionarios de normalizacion y listas | si |
| `Documentacion` | Mapa del libro, ETL paso a paso y criterios de diseno | si |
| `Calculos` | Grilla que resuelve GETPIVOTDATA y alimenta los graficos | oculta |
| `Pivots_Aux` | Tablas dinamicas de soporte | oculta |

---

## El ETL, en una linea por paso

El libro arranca de un export de **2.458 filas** y llega a **2.400
operaciones unicas**:

| Paso | Funciones | Defectos que corrige |
|---|---|---|
| 1. Quitar duplicados | COINCIDIR | 58 reenvios del sistema, que inflaban la facturacion en USD 258.616 |
| 2. Normalizar fechas | SI, ESNUMERO, EXTRAE, FECHA, VALOR | 845 fechas cargadas como texto, en dos formatos distintos |
| 3. Estandarizar textos | ESPACIOS, INDICE, COINCIDIR | Region, Canal y Categoria con mayusculas, acentos y espacios mezclados |
| 4. Corregir tipos | SI, ESTEXTO, VALOR | 246 unidades cargadas como texto |
| 5. Completar vacios | SI, INDICE, COINCIDIR | 159 ciudades vacias, imputadas por region |
| 6. Derivar columnas | ANIO, MES, REDONDEAR.MAS, TEXTO | Anio, Trimestre, Mes, Periodo, Cat_Anio, Prod_Anio |
| 7. Calcular medidas | REDONDEAR, SUMAR.SI.CONJUNTO | Ventas, Costo_Producto, Costo_Logistico, Margen_Bruto |
| 8. Clasificar | **SI, Y, O** | Segmento_Rentabilidad, Alerta_Logistica, Ticket |

La deduplicacion es el unico paso que se resuelve fuera de la planilla (la
columna `Pos_Origen` guarda a que fila del export apunta cada operacion), y la
columna `Control_ETL` lo verifica con COINCIDIR fila por fila: las 2.400 dan
`OK`.

---

## Como viaja un filtro hasta un grafico

```
segmentador  ->  tabla dinamica  ->  GETPIVOTDATA (Calculos)  ->  grafico
```

Los graficos no leen la tabla de hechos: leen una grilla de tamano fijo en
`Calculos` resuelta con GETPIVOTDATA. Por eso el layout no se rompe cuando un
filtro deja menos filas, y por eso los numeros del tablero y los de
`Tablas_Dinamicas` son siempre los mismos. Las 17 tablas dinamicas comparten
un unico cache, asi que una sola seleccion recalcula a la vez las 4 tarjetas,
los 5 graficos y las 4 tablas visibles.

**Por que el filtro de tiempo es el Trimestre y no el Ano:** un segmentador de
ano dejaria sin sentido el KPI de crecimiento interanual (al elegir 2025 no
habria 2024 contra que comparar). Con el trimestre, elegir Q1 compara Q1-2024
contra Q1-2025 y la comparacion se mantiene valida siempre.

---

## Que dice el tablero

- Las ventas crecen **+17,8%** interanual (USD 4.726.315 → USD 5.566.046) y
  superan la meta de +15%. El crecimiento es de **volumen**: las unidades
  suben +14,4%.
- El margen bruto cae **2,5 p.p.** (35,8% → 33,3%) y queda por debajo de la
  meta de 34%.
- El puente de margen reparte esa caida: volumen +USD 258.146 y precio
  +USD 159.153 contra costo de compra −USD 172.525 y logistica −USD 80.921.
  **La logistica explica el 32% del mayor costo, no el total**: atribuir la
  caida solo al e-commerce sobreestima su peso.
- La correccion aplicada en el Q3 funciona pero es parcial: el segundo
  semestre (34,5%) mejora sobre el primero (31,8%) y sigue por debajo del
  mismo periodo de 2024 (35,7%).

---

## Sobre los datos

El dataset es **simulado y determinista** (semilla fija). Extiende el modelo
`Ventas_Tech_DB` del [checkpoint de SQL](../checkpoint_sql_ventas_tech)
—mismas categorias, productos y ciudades— a 24 meses, 5 regiones y 2 canales,
y es el mismo que se viene madurando desde el
[Modulo 9](../modulo9_dashboard_final). Todos los supuestos estan documentados
en la hoja `Parametros`.

---

## Reproducir el libro

La carpeta [`generador/`](./generador) arma el archivo de cero. El entregable
es el `.xlsx`; los scripts existen para que el resultado sea auditable.

```bash
pip install openpyxl
python generador/pipeline.py            # construye, recalcula, verifica y audita
python generador/auditar.py PracticaFinal_TechStore_Murphy_Lleyton.xlsx
```

| Modulo | Que hace |
|---|---|
| `datos.py` | Genera la tabla de hechos correcta y la degrada al export crudo |
| `parametros.py` | Hoja de constantes: metas, tasas, diccionarios y listas |
| `etl.py` | Hojas `Datos_Origen`, `Datos_Limpios` y `Control_Calidad` |
| `analisis.py` | Tablas dinamicas y la capa `Calculos` con GETPIVOTDATA |
| `dashboard.py` | El tablero: KPIs, graficos y reporte ejecutivo |
| `documentacion.py` | La hoja de documentacion del libro |
| `ooxml.py` | Tablas dinamicas, segmentadores y retoques de graficos a nivel OOXML |
| `verificar.py` | Compara las 60.000 celdas del ETL contra el modelo |
| `auditar.py` | 30 controles estructurales sobre el `.xlsx` terminado |

El build **falla a proposito** si el ETL de la planilla no reproduce
exactamente los valores esperados, si alguna formula da error, o si al archivo
le falta cualquiera de los elementos que pide la consigna.

---

**Lleyton Murphy** · Coderhouse, Data Analyst · Practica Integradora Final
