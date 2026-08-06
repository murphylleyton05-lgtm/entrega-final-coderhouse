# Checkpoint — Script SQL de Ingenieria de Datos (`Ventas_Tech_DB`)

Back-End del proyecto final de **Data Analyst (Coderhouse)**: el script SQL que
crea la base de datos relacional, limpia y normalizada (3NF) donde reside toda
la informacion de ventas de la cadena de tecnologia **TechStore**.

## Archivo

- [`ventas_tech_db.sql`](./ventas_tech_db.sql) — script completo con las 3 secciones:
  1. **DDL** — `DROP TABLE` (orden inverso de dependencias) + `CREATE TABLE`
     (dimensiones primero, tabla de hechos al final).
  2. **Restricciones de integridad** — `PRIMARY KEY`, `FOREIGN KEY`,
     `NOT NULL`, `UNIQUE` y `DEFAULT`.
  3. **DML** — carga inicial de datos.

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

- 4 categorias · 5 clientes · 6 productos · **10 ventas**.

## Como ejecutarlo (PostgreSQL)

```bash
# 1) crear la base (una sola vez, conectado al servidor)
psql -U postgres -c "CREATE DATABASE ventas_tech_db;"

# 2) ejecutar el script (es repetible: se puede correr cuantas veces quieras)
psql -U postgres -d ventas_tech_db -f ventas_tech_db.sql
```

## Verificacion realizada

El script fue probado en **PostgreSQL 16** con `ON_ERROR_STOP=1`:

- Se ejecuta **sin errores** y es **repetible** (corrido dos veces seguidas, OK).
- Conteos correctos: `categorias=4`, `clientes=5`, `productos=6`, `ventas=10`.
- Las **foreign keys** rechazan una venta con `id_producto` inexistente.
- El **`NOT NULL`** rechaza un producto con `precio` nulo.
