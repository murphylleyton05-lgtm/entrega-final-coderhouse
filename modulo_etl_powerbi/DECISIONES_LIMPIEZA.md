# Decisiones de limpieza — Pipeline ETL TechStore

Documento de justificación técnica del checkpoint **Pipeline ETL desde SQL con
Power Query y M**. Explica qué se hizo con cada problema de calidad del dataset
y por qué. Autor: Lleyton Murphy.

## Resumen de problemas y resolución

| Tabla | Problema | Decisión | Justificación |
|-------|----------|----------|---------------|
| `Dim_Clientes` | `id_cliente` duplicado (alta reintentada) | **Quitar duplicados por `id_cliente`** (conserva la 1ª aparición) | La PK de una dimensión debe ser única; si no, la relación 1:N con `Fact_Ventas` se vuelve muchos-a-muchos y las medidas DAX cuentan doble. |
| `Dim_Clientes` | `email` nulo (1 registro) | **Reemplazar por `"Sin dato"`** (no se borra la fila) | El email es un dato de contacto, no crítico. Borrar el cliente eliminaría su historial de ventas asociado por `id_cliente`. |
| `Dim_Clientes` | `ciudad` nula (1 registro) | **Reemplazar por `"Sin dato"`** | Igual criterio: la ciudad alimenta segmentadores; un nulo rompería filtros/visuales, pero no amerita perder al cliente. |
| `Dim_Productos` | `id_producto = 103` duplicado | **Quitar duplicados por `id_producto`** | Mismo motivo que en clientes: unicidad de PK para el modelo estrella. |
| `Dim_Productos` | `precio` nulo (1 registro — **crítico**) | **Imputar con el promedio de precio de su categoría** | Sin precio no hay ingreso, pero borrar la fila descuadra el catálogo y deja ventas huérfanas en el merge. El promedio por categoría es un valor plausible y trazable; mejor que `0` (sesga los KPI a la baja) o que un número inventado a mano. |
| `Dim_Productos` | `categoria` nula (1 registro) | **Reemplazar por `"Sin Categoría"`** | El producto existe en el catálogo aunque falte clasificarlo. Se resuelve **antes** que el precio para que la imputación por categoría no falle. |
| `Fact_Ventas` | — (tabla limpia) | Solo tipado + Merge | No tiene problemas de calidad; se enriquece con atributos de producto. |
| `Dim_Categorias` | — (tabla limpia) | Solo tipado | Tabla de referencia sin problemas. |

## Por qué el precio se imputa y el email/ciudad también se conservan (y no se borran)

La regla no es "todo nulo se borra" ni "todo nulo se reemplaza": **depende de la
criticidad del campo y del rol de la tabla**.

- En una **dimensión**, borrar una fila por un atributo secundario nulo (email,
  ciudad) destruye la clave y rompe relaciones. Se conserva la fila y se marca el
  faltante (`"Sin dato"` / `"Sin Categoría"`).
- Un **dato crítico** como el precio no puede quedar nulo (rompe cálculos de
  ingreso) ni borrarse (descuadra el conteo y el merge). La imputación por
  promedio de categoría lo resuelve sin distorsionar ni perder la fila.

## Conteo de filas esperado (verificación antes de *Cerrar y aplicar*)

| Consulta | Filas | Cómo se llega |
|----------|-------|---------------|
| `Dim_Clientes` | **11** | 12 crudas − 1 duplicado |
| `Dim_Productos` | **12** | 13 crudas − 1 duplicado (los 2 nulos se imputan, no se borran) |
| `Fact_Ventas` | **50** | 50 crudas (merge LEFT OUTER conserva todas) |
| `Dim_Categorias` | **4** | tabla de referencia limpia |

Estos conteos coinciden con los criterios de aceptación del checkpoint.

## Tipos de dato asignados

| Columna(s) | Tipo | Motivo |
|------------|------|--------|
| `fecha_venta`, `fecha_registro` | **Date** | Habilita la tabla calendario y la inteligencia de tiempo (M8). Si quedan como texto, no hay línea de tiempo. |
| `precio`, `costo`, `total_venta`, `descuento`, `precio_unitario` | **Número decimal** | Importes; permite `SUM`/`AVG` sin perder centavos. |
| IDs, `cantidad`, `stock`, `activo` | **Número entero** | Claves y conteos. IDs enteros dan relaciones limpias. |
| Nombres, `categoria`, `canal`, `email`, `ciudad` | **Texto** | Atributos descriptivos. |

## Merge aplicado

`Fact_Ventas` ← `Dim_Productos` por `id_producto` (LEFT OUTER), expandiendo solo
`nombre_producto` y `categoria`. Se usa LEFT OUTER para que ninguna venta se
pierda y se expanden solo esas dos columnas para no duplicar los demás atributos
del producto (que viven en `Dim_Productos`), respetando el esquema estrella.
