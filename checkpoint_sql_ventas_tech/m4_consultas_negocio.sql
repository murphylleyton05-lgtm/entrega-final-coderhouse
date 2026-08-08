-- =============================================================================
--  m4_consultas_negocio.sql  —  Extrayendo metricas clave con SQL
-- -----------------------------------------------------------------------------
--  Autor    : Lleyton Murphy
--  Curso    : Data Analyst — Coderhouse  (Modulo 4 — Funciones de agregacion)
--  Base     : Ventas_Tech_DB  (creada en el Modulo 3 con ventas_tech_db.sql)
--  Tabla    : ventas (id_venta, id_cliente, id_producto, cantidad,
--                     precio_unitario, fecha_venta)
--
--  Objetivo : responder las preguntas de negocio del brief de M1 con consultas
--             de agregacion (COUNT, SUM, AVG, MIN, MAX) + GROUP BY, HAVING,
--             ORDER BY y CASE WHEN. Trabajamos SOLO sobre la tabla `ventas`:
--             los nombres de clientes y productos se resuelven con JOIN en M5.
--
--  Motor    : PostgreSQL 16 (probado con ON_ERROR_STOP=1, corre sin errores).
--  Ejecucion: psql -U postgres -d ventas_tech_db -f m4_consultas_negocio.sql
--
--  Regla de negocio: el importe de cada linea de venta es
--             total_linea = cantidad * precio_unitario
-- =============================================================================


-- =============================================================================
--  CONSULTA 1 — Resumen ejecutivo mensual
-- -----------------------------------------------------------------------------
--  Pregunta de negocio: "¿Cuanto facturamos por mes, cuantos pedidos hicimos
--  y cual fue el ticket promedio?"
--
--  Funciones usadas:
--    SUM   -> total facturado del mes (suma de cantidad * precio_unitario)
--    COUNT -> cantidad de pedidos (COUNT(*) cuenta filas, incluso con nulos)
--    AVG   -> ticket promedio = facturacion / cantidad de pedidos
-- =============================================================================
SELECT
    EXTRACT(MONTH FROM fecha_venta)             AS mes,
    COUNT(*)                                    AS cantidad_pedidos,
    SUM(cantidad * precio_unitario)             AS total_facturado,
    ROUND(AVG(cantidad * precio_unitario), 2)   AS ticket_promedio
FROM ventas
GROUP BY EXTRACT(MONTH FROM fecha_venta)
ORDER BY mes;


-- =============================================================================
--  CONSULTA 2 — Ranking de productos (Top 5 por facturacion)
-- -----------------------------------------------------------------------------
--  Pregunta de negocio: "¿Que productos generan mas dinero y cuantas unidades
--  se vendieron de cada uno?"
--
--  Ojo con la diferencia COUNT / SUM: `unidades_vendidas` NO es la cantidad de
--  pedidos, es la suma de las unidades de cada pedido (SUM(cantidad)).
--  Por eso agregamos tambien `cantidad_pedidos`: son metricas distintas.
-- =============================================================================
SELECT
    id_producto,
    SUM(cantidad)                     AS unidades_vendidas,
    COUNT(*)                          AS cantidad_pedidos,
    SUM(cantidad * precio_unitario)   AS total_facturado
FROM ventas
GROUP BY id_producto
ORDER BY total_facturado DESC
LIMIT 5;


-- =============================================================================
--  CONSULTA 3 — Clientes recurrentes
-- -----------------------------------------------------------------------------
--  Pregunta de negocio: "¿Que clientes compraron mas de una vez y cuanto
--  gastaron en total?"
--
--  HAVING filtra DESPUES de agrupar (sobre el resultado del COUNT), mientras
--  que WHERE filtraria fila por fila ANTES de agrupar: por eso la condicion
--  "mas de un pedido" solo puede escribirse con HAVING.
-- =============================================================================
SELECT
    id_cliente,
    COUNT(*)                                    AS cantidad_pedidos,
    SUM(cantidad * precio_unitario)             AS total_gastado,
    ROUND(AVG(cantidad * precio_unitario), 2)   AS ticket_promedio,
    MAX(fecha_venta)                            AS ultima_compra
FROM ventas
GROUP BY id_cliente
HAVING COUNT(*) > 1
ORDER BY total_gastado DESC;


-- =============================================================================
--  CONSULTA 4 — Meses por encima / por debajo del promedio
-- -----------------------------------------------------------------------------
--  Pregunta de negocio: "¿Que meses rindieron mejor que un mes tipico?"
--
--  Se resuelve en dos pasos:
--    1) el CTE `facturacion_mensual` calcula el total de cada mes;
--    2) la consulta final compara ese total contra el promedio mensual general
--       (subconsulta escalar) y lo etiqueta con CASE WHEN.
--
--  Se incluye la rama 'En el promedio' para el caso de empate exacto: sin ella,
--  un mes igual al promedio quedaria mal etiquetado como 'Por debajo'.
-- =============================================================================
WITH facturacion_mensual AS (
    SELECT
        EXTRACT(MONTH FROM fecha_venta)   AS mes,
        SUM(cantidad * precio_unitario)   AS total_facturado
    FROM ventas
    GROUP BY EXTRACT(MONTH FROM fecha_venta)
)
SELECT
    mes,
    total_facturado,
    ROUND((SELECT AVG(total_facturado) FROM facturacion_mensual), 2)
        AS promedio_mensual_general,
    CASE
        WHEN total_facturado > (SELECT AVG(total_facturado) FROM facturacion_mensual)
            THEN 'Por encima'
        WHEN total_facturado < (SELECT AVG(total_facturado) FROM facturacion_mensual)
            THEN 'Por debajo'
        ELSE 'En el promedio'
    END AS comparativa_vs_promedio
FROM facturacion_mensual
ORDER BY mes;


-- =============================================================================
--  CONSULTA EXTRA — Metricas generales del periodo (control / KPI cards)
-- -----------------------------------------------------------------------------
--  Las cinco funciones del modulo en una sola foto. Es la consulta que alimenta
--  las tarjetas de KPI del dashboard de Power BI en M6.
-- =============================================================================
SELECT
    COUNT(*)                                    AS pedidos_totales,
    COUNT(DISTINCT id_cliente)                  AS clientes_distintos,
    COUNT(DISTINCT id_producto)                 AS productos_distintos,
    SUM(cantidad)                               AS unidades_vendidas,
    SUM(cantidad * precio_unitario)             AS facturacion_total,
    ROUND(AVG(cantidad * precio_unitario), 2)   AS ticket_promedio,
    MIN(cantidad * precio_unitario)             AS ticket_minimo,
    MAX(cantidad * precio_unitario)             AS ticket_maximo,
    MIN(fecha_venta)                            AS primera_venta,
    MAX(fecha_venta)                            AS ultima_venta
FROM ventas;


-- =============================================================================
--  HALLAZGOS  (analisis de los resultados obtenidos al ejecutar el script)
-- =============================================================================
--
--  Periodo analizado: 05/03/2024 a 15/03/2024 — 10 pedidos, 29 unidades,
--  USD 6.444,00 facturados, ticket promedio USD 644,40.
--
--  1) La facturacion depende de un solo producto: el producto 1 (Laptop Pro 15)
--     genera USD 3.600 de los USD 6.444 totales, es decir el 55,9% de la
--     facturacion con apenas 3 unidades vendidas. Sumado al producto 3
--     (USD 1.350), entre dos productos concentran el 76,8% del ingreso: si cae
--     el stock de la laptop, se cae mas de la mitad del negocio del mes.
--
--  2) Vender muchas unidades no es vender mucho dinero. El producto 2 es el
--     mas vendido en volumen (13 unidades, casi la mitad de las 29 del periodo)
--     pero factura solo USD 364, el 5,6% del total; el producto 1 vende 3
--     unidades y factura 10 veces mas. El ranking por SUM(cantidad) y el
--     ranking por SUM(cantidad * precio_unitario) dan ordenes casi opuestos:
--     confundir COUNT/SUM de unidades con facturacion cambiaria la decision.
--
--  3) El 100% de los clientes es recurrente, pero el gasto esta muy disparejo.
--     Los 5 clientes tienen 2 pedidos cada uno (la Consulta 3 devuelve las 5
--     filas), sin embargo el cliente 1 gasto USD 2.640 y el cliente 4 solo
--     USD 510: una brecha de 5,2x entre el mejor y el peor cliente. La
--     frecuencia de compra ya esta ganada; la palanca de crecimiento es el
--     valor del ticket, no la recurrencia.
--
--  Nota metodologica: la carga inicial de M3 concentra las 10 ventas en marzo
--  de 2024, por lo que las Consultas 1 y 4 devuelven una unica fila y ese mes
--  queda etiquetado 'En el promedio' (es el unico mes, asi que es igual al
--  promedio). Las consultas ya estan escritas para varios meses: al ampliar la
--  carga de datos en M5/M6 el resumen mensual y la comparativa vs. promedio
--  devuelven una fila por mes sin tocar una sola linea de este archivo.
--
-- =============================================================================
--  FIN DEL SCRIPT
-- =============================================================================
