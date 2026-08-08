# Modulo 9 — Capa de visualizacion final del dashboard

Pre-entrega del **Sistema de Inteligencia de Negocios** de la cadena de
tecnologia **TechStore**. Es el checkpoint de diseno: se evalua la interfaz y
la capacidad narrativa del reporte, no el calculo.

## Que se entrega

| Archivo | Que es |
|---|---|
| [`TechStore_Dashboard_Final_Murphy_Lleyton.xlsx`](./TechStore_Dashboard_Final_Murphy_Lleyton.xlsx) | El libro con la pestana **Dashboard Final** terminada y funcional |
| [`Registro_Prompts_IA_Murphy_Lleyton.pdf`](./Registro_Prompts_IA_Murphy_Lleyton.pdf) | Documento anexo con los prompts usados en ChatGPT y las correcciones que produjeron |
| `Registro_Prompts_IA_Murphy_Lleyton.docx` | El mismo documento en Word, por si se pide editable |

> El libro se abre en la pestana `Dashboard Final`. **Requiere Excel 2010 o
> posterior**: los segmentadores son objetos nativos de Excel y otros visores
> (Google Sheets, LibreOffice, vista previa de GitHub) los descartan al abrir.

## Hojas del libro

| Hoja | Rol | Visible |
|---|---|---|
| `Dashboard Final` | El entregable: KPIs, 4 graficos, 4 segmentadores y narrativa | si |
| `Analisis_IA` | Registro del uso de IA como auditor cognitivo | si |
| `Datos` | Tabla de hechos: 2.400 operaciones, 2024-2025 | si |
| `Parametros` | Supuestos, metas y tasas: toda constante del modelo vive aca | si |
| `Guia_de_uso` | Como leer y operar el tablero | si |
| `Calculos` | Capa intermedia que alimenta los graficos | oculta |
| `Pivots` | Tablas dinamicas que los segmentadores filtran | oculta |

## Como se cumple cada punto de la consigna

| Requisito | Donde se cumple |
|---|---|
| Layout profesional, fondo sobrio, objetos alineados | Grilla de 16 columnas, lineas de cuadricula ocultas, hojas de soporte ocultas |
| Maximo 3 colores principales | Azul `#1F3864` (institucional / ano corriente), gris `#8C8C8C` (contexto / ano anterior), ambar `#E8A33D` (alerta y efectos que restan). El resto son neutros de fondo y texto |
| Tarjetas de KPI con formato condicional | 4 tarjetas: Ventas, Margen bruto %, Crecimiento interanual y Costo logistico / ventas. Las que no llegan a la meta de `Parametros` pasan a ambar |
| >= 3 segmentadores conectados a todos los graficos | 4 segmentadores nativos: **Trimestre, Categoria, Region y Canal**. Comparten un unico cache de tablas dinamicas, asi que cualquier seleccion recalcula a la vez las 4 tarjetas y los 4 graficos |
| Al menos 1 grafico avanzado | **Cascada** (puente de margen 2024 -> 2025) y **dispersion** (crecimiento contra margen por producto). Se entregan los dos |
| Capa narrativa | Titular y bajada dinamicos, los 4 titulos de grafico son formulas que apuntan a celdas, y 3 notas de analisis que se reescriben con los datos |
| Analisis asistido por IA | Hoja `Analisis_IA` + el documento anexo |

## Por que el filtro de tiempo es el trimestre y no el ano

Un segmentador de ano dejaria sin sentido el KPI de crecimiento interanual: al
elegir 2025 no habria 2024 contra que comparar. Con el trimestre, elegir Q1
compara Q1-2024 contra Q1-2025 y la comparacion se mantiene siempre valida.

## Como viaja un filtro hasta un grafico

```
Segmentador  ->  tabla dinamica (Pivots)  ->  GETPIVOTDATA (Calculos)  ->  grafico
```

Los graficos no leen la tabla de hechos: leen una grilla de tamano fijo en
`Calculos` que se resuelve con `GETPIVOTDATA` contra las tablas dinamicas. Por
eso el layout no se rompe cuando un filtro deja menos filas.

## Que dice el tablero

- Las ventas crecen **+17,8%** interanual (USD 4.726.315 -> USD 5.566.046) y
  superan la meta de +15%.
- El margen bruto cae **2,5 p.p.** (35,8% -> 33,3%) y queda por debajo de la
  meta de 34%.
- El puente de margen reparte esa caida: volumen +USD 258.146 y precio
  +USD 159.153 contra costo de compra -USD 172.525 y logistica -USD 80.921.
- La correccion aplicada en el Q3 funciona pero es parcial: el segundo semestre
  (34,5%) mejora sobre el primero (31,8%) y sigue por debajo del mismo periodo
  de 2024 (35,7%).

## Sobre los datos

El dataset es **simulado y determinista** (semilla fija). Extiende el modelo
`Ventas_Tech_DB` del [proyecto RetailPro](../RetailPro) —mismas
categorias, productos y ciudades— a 24 meses, 5 regiones y 2 canales, para
poder analizar margen, estacionalidad y evolucion interanual. Todos los
supuestos estan documentados en la hoja `Parametros`.

## Reproducir el libro

La carpeta [`generador/`](./generador) contiene los scripts que arman el
archivo de cero. El entregable es el `.xlsx`; los scripts estan para que el
resultado sea auditable y reproducible.

```bash
pip install openpyxl python-docx
python generador/pipeline.py salida.xlsx
```

---

**Lleyton Murphy** · Coderhouse, Data Analyst · Modulo 9
