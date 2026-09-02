# Proyecto RetailPro — Base de datos `Ventas_Tech_DB` y consultas de negocio

Back-End del proyecto final de **Data Analyst (Coderhouse)**: la base de datos
relacional, limpia y normalizada (3NF) donde reside toda la informacion de
ventas de la cadena de tecnologia **TechStore**, y las consultas SQL que
extraen de ella las metricas del brief.

## Archivos

- [`ventas_tech_db.sql`](./ventas_tech_db.sql) — **M3**: script completo con las 3 secciones:
  1. **DDL** — `DROP TABLE` (orden inverso de dependencias) + `CREATE TABLE`
     (dimensiones primero, tabla de hechos al final).
  2. **Restricciones de integridad** — `PRIMARY KEY`, `FOREIGN KEY`,
     `NOT NULL`, `UNIQUE` y `DEFAULT`.
  3. **DML** — carga inicial de datos.
- [`m4_consultas_negocio.sql`](./m4_consultas_negocio.sql) — **M4**: consultas de
  agregacion que responden las preguntas de negocio del brief (ver mas abajo).
- [`m5_consultas_joins.sql`](./m5_consultas_joins.sql) — **M5**: consultas con
  `JOIN` (INNER / LEFT) y `UNION ALL` que cruzan las tablas para armar la vista
  enriquecida del dashboard (ver mas abajo).

## Modelo de datos

```
categorias (1) ── (N) productos (1) ── (N) ventas (N) ── (1) clientes
```

| Tabla        | Rol        | PK             | FK                                  |
|--------------|------------|----------------|-------------------------------------|
| `categorias` | Dimension  | `id_categoria` | —                                   |
| `clientes`   | Dimension  | `id_cliente`   | —                                   |
| `productos`  | Dimension  | `id_producto`  | `id_categoria` → `categorias`       |
| `ventas`     | **Hechos** | `id_venta`     | `id_cliente` → `clientes`, `id_producto` → `productos` |

Precios en `DECIMAL(10,2)` (nunca texto ni `FLOAT`) para poder usar `SUM`/`AVG`.

## Datos cargados

- 4 categorias · 7 clientes · 8 productos · **10 ventas**.
- A proposito, **2 clientes** (6 y 7) y **2 productos** (7 y 8) quedan **sin ventas**:
  son los casos que aislan las consultas `LEFT JOIN ... IS NULL` del M5.
- La tabla `clientes` incluye dos dimensiones para el dashboard: `ciudad`
  (geografica) y `segmento` (Consumo / PyME / Corporativo).

## Como ejecutarlo (PostgreSQL)

```bash
# 1) crear la base (una sola vez, conectado al servidor)
psql -U postgres -c "CREATE DATABASE ventas_tech_db;"

# 2) ejecutar el script (es repetible: se puede correr cuantas veces quieras)
psql -U postgres -d ventas_tech_db -f ventas_tech_db.sql

# 3) M4 — consultas de negocio sobre la base ya cargada
psql -U postgres -d ventas_tech_db -f m4_consultas_negocio.sql

# 4) M5 — consultas con JOINs / UNION ALL sobre la base ya cargada
psql -U postgres -d ventas_tech_db -f m5_consultas_joins.sql
```

## M4 — Consultas de negocio (`m4_consultas_negocio.sql`)

Consultas de agregacion sobre la tabla de hechos `ventas`
(`COUNT`, `SUM`, `AVG`, `MIN`, `MAX` + `GROUP BY`, `HAVING`, `ORDER BY`, `CASE WHEN`).
Los nombres de clientes y productos se resuelven con `JOIN` en M5: aca se trabaja con IDs.

| # | Consulta | Pregunta de negocio | Resultado sobre los datos cargados |
|---|----------|---------------------|------------------------------------|
| 1 | Resumen ejecutivo mensual | ¿Cuanto facturamos, cuantos pedidos y ticket promedio por mes? | Marzo 2024: 10 pedidos · USD 6.444,00 · ticket USD 644,40 |
| 2 | Ranking Top 5 de productos | ¿Que productos generan mas dinero? | Producto 1 (USD 3.600) · 3 (1.350) · 5 (390) · 6 (380) · 2 (364) |
| 3 | Clientes recurrentes (`HAVING COUNT(*) > 1`) | ¿Quienes compraron mas de una vez y cuanto gastaron? | Los 5 clientes, 2 pedidos c/u; del cliente 1 (USD 2.640) al 4 (USD 510) |
| 4 | Meses vs. promedio (`CASE WHEN`) | ¿Que meses rindieron mejor que un mes tipico? | Unico mes cargado (marzo) → `En el promedio` |
| + | Metricas generales | KPI cards para el dashboard de M6 | 29 unidades · ticket min USD 120 / max USD 2.400 |

Al final del archivo hay un bloque de comentarios con los **3 hallazgos** del analisis
(concentracion de la facturacion, volumen vs. ingreso, y recurrencia vs. valor del ticket).

El script fue ejecutado en **PostgreSQL 16** con `ON_ERROR_STOP=1`: las 5 consultas
devuelven resultados sin errores.

## M5 — Consultas con JOINs (`m5_consultas_joins.sql`)

Las consultas de M4 trabajaban sobre tablas individuales; en M5 las **cruzamos**
para armar la vista enriquecida que consume Power BI y responder tres preguntas
de negocio, mas una consolidacion por canal con `UNION ALL`.

| # | Consulta | Tecnica | Resultado sobre los datos cargados |
|---|----------|---------|------------------------------------|
| 1 | Vista base del proyecto | `INNER JOIN` de las 4 tablas | 10 ventas enriquecidas (fecha, cliente, region, segmento, producto, categoria, cantidad, precio, total) |
| 2 | Clientes sin ventas | `LEFT JOIN` + `WHERE ... IS NULL` | 2 clientes: Diego Fernandez y Sofia Ramirez |
| 3 | Productos sin ventas | `LEFT JOIN` + `WHERE ... IS NULL` | 2 productos: Webcam Full HD y Parlante BT Mini |
| 4 | Consolidado por canal | `UNION ALL` + `GROUP BY` (columna `canal` literal) | Online USD 3.620 (5 ventas) · Presencial USD 2.824 (5 ventas) |

En la Consulta 4 la columna `canal` **no se consulta, se crea**: es un texto fijo
(`'Online'` / `'Presencial'`) dentro de cada `SELECT`. Se usa `UNION ALL` y no
`UNION` para que **no** se eliminen filas repetidas (cada venta se cuenta una vez).
Al final del archivo hay un bloque con los **4 hallazgos** del analisis.

## Verificacion realizada

Los tres scripts fueron probados en **PostgreSQL 16** con `ON_ERROR_STOP=1`:

- Se ejecutan **sin errores** y M3 es **repetible** (corrido dos veces seguidas, OK).
- Conteos correctos: `categorias=4`, `clientes=7`, `productos=8`, `ventas=10`.
- Las **foreign keys** rechazan una venta con `id_producto` inexistente.
- El **`NOT NULL`** rechaza un producto con `precio` nulo.
- M5 corre sobre la base ya cargada: las 4 consultas devuelven los resultados de
  la tabla de arriba sin errores.
