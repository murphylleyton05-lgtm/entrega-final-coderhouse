# Checkpoint — Pipeline ETL en Power BI (TechStore)

Pipeline ETL de **TechStore**: se conecta a un dataset con problemas de calidad
reales, perfila, limpia (duplicados + nulos), tipa, renombra con estándar
`Dim_`/`Fact_`, hace el Merge y documenta la lógica en lenguaje M. La salida es
un `.pbix` limpio, base del modelo analítico del Módulo 8.

> **Importante — cómo leer esta carpeta**
> El entregable oficial es el archivo `Pipeline_ETL_Murphy_Lleyton.pbix`. Ese
> archivo es un binario de **Power BI Desktop** (app de Windows) y **se genera
> y guarda desde Power BI Desktop** — no se puede producir fuera de esa app. Esta
> carpeta contiene **todo lo necesario para construirlo en ~5 minutos**: el
> dataset, el código M de las 4 consultas ya escrito, y la justificación técnica.
> Seguí los pasos de abajo y guardás el `.pbix` en esta misma carpeta.

## Contenido de la carpeta

```
modulo_etl_powerbi/
├── README.md                     ← este archivo (pasos + checklist)
├── DECISIONES_LIMPIEZA.md        ← justificación de nulos, duplicados y tipos
├── Pipeline_ETL_Dataset.xlsx     ← dataset con los problemas intencionales (amarillo)
└── consultas_M/                  ← código Power Query listo para pegar
    ├── 00_Parametro_RutaDataset.pq
    ├── Dim_Categorias.pq
    ├── Dim_Clientes.pq
    ├── Dim_Productos.pq
    └── Fact_Ventas.pq
```

> **Sobre el dataset:** `Pipeline_ETL_Dataset.xlsx` de esta carpeta reproduce
> **exactamente** el spec del curso (clientes: 1 duplicado + 2 nulos; productos:
> 1 duplicado en `id_producto=103` + 2 nulos; ventas: 50 limpias; categorías: 4).
> Si tu comisión te da el archivo oficial, podés usar ese en su lugar: el código M
> es el mismo siempre que los nombres de columna coincidan (revisá la sección
> "Si los nombres de columna difieren").

## Paso a paso para generar el `.pbix`

1. **Abrí Power BI Desktop** → *Inicio* → *Obtener datos* → *Excel* → elegí
   `Pipeline_ETL_Dataset.xlsx`.
2. En el **Navegador**, tildá las 4 hojas (`clientes`, `productos`, `ventas`,
   `categorias`) → **Transformar datos** (NO "Cargar"). Entrás a Power Query.
3. **Perfilado** → pestaña *Vista* → activá *Calidad de columnas*, *Distribución
   de columnas* y *Perfil de columna* → cambiá "Primeras 1000 filas" a **Todo el
   conjunto de datos**. Vas a ver los % de vacíos/errores y los duplicados.
4. **Creá el parámetro de ruta** → *Inicio* → *Administrar parámetros* → *Nuevo*:
   nombre `RutaDataset`, tipo *Texto*, valor = la ruta completa a tu
   `Pipeline_ETL_Dataset.xlsx` (ver `consultas_M/00_Parametro_RutaDataset.pq`).
5. **Pegá el código M** de cada consulta: seleccioná la consulta en el panel
   izquierdo → *Inicio* → *Editor avanzado* → reemplazá todo por el contenido del
   `.pq` correspondiente → *Aceptar*. **Renombrá** cada consulta al estándar:
   `clientes`→`Dim_Clientes`, `productos`→`Dim_Productos`,
   `categorias`→`Dim_Categorias`, `ventas`→`Fact_Ventas`.
   - Orden recomendado: primero `Dim_Categorias`, `Dim_Clientes`, `Dim_Productos`,
     y **al final `Fact_Ventas`** (porque su Merge depende de `Dim_Productos`).
6. **Verificá los conteos** (barra inferior de cada consulta):
   `Dim_Clientes` = **11**, `Dim_Productos` = **12**, `Fact_Ventas` = **50**,
   `Dim_Categorias` = **4**. Ninguna consulta debe tener el triángulo de error.
7. **Cerrar y aplicar**. Confirmá que no aparecen errores en la vista de informe.
8. **Guardá como** `Pipeline_ETL_Murphy_Lleyton.pbix` dentro de esta carpeta
   (`modulo_etl_powerbi/`) y subilo al repo.

## Qué cumple cada criterio de aceptación

| Criterio | Dónde se cumple |
|----------|-----------------|
| 4 tablas cargadas sin errores con nomenclatura `Dim_`/`Fact_` | Paso 5 (renombrado) + los 4 `.pq` |
| Duplicados eliminados en `Dim_Clientes` y `Dim_Productos` | `Table.Distinct(..., {"id_cliente"})` / `{"id_producto"}` |
| Nulos resueltos con decisión justificada | `DECISIONES_LIMPIEZA.md` + comentarios en el M |
| Tipos de dato correctos en todas las columnas | `Table.TransformColumnTypes` en cada consulta |
| Merge que agrega `nombre_producto` y `categoria` a `Fact_Ventas` | `Fact_Ventas.pq` (NestedJoin + ExpandTableColumn) |
| ≥ 2 consultas con comentarios técnicos en el Editor Avanzado | Las 4 consultas están comentadas con `//` |

## Si los nombres de columna difieren (usando el .xlsx oficial del curso)

El código M referencia estos nombres de columna. Si el archivo oficial usa otros,
ajustalos en los pasos de `Table.TransformColumnTypes` y en el Merge:

- **clientes:** `id_cliente, nombre, email, ciudad, fecha_registro`
- **productos:** `id_producto, nombre_producto, categoria, precio, costo, stock, activo`
- **ventas:** `id_venta, id_cliente, id_producto, cantidad, precio_unitario, descuento, total_venta, fecha_venta, canal`
- **categorias:** `id_categoria, nombre_categoria`

## Verificación reproducible

La lógica de limpieza se validó de forma independiente contra el dataset: da
exactamente **11 / 12 / 50 / 4** filas, cero nulos remanentes, el merge matchea las
50 ventas y el precio crítico se imputa a **61,5** (promedio de su categoría).
