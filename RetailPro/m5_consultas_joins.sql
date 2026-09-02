-- =============================================================================
--  m5_consultas_joins.sql  —  Cruzando tablas para enriquecer el analisis
-- -----------------------------------------------------------------------------
--  Autor    : Lleyton Murphy
--  Curso    : Data Analyst — Coderhouse  (Modulo 5 — JOINs, UNION / UNION ALL)
--  Base     : Ventas_Tech_DB  (creada en el Modulo 3 con ventas_tech_db.sql)
--  Modelo   : categorias (1)-(N) productos (1)-(N) ventas (N)-(1) clientes
--
--  Objetivo : las consultas de M4 trabajaban sobre tablas individuales; aca las
--             CRUZAMOS para obtener la vista enriquecida que necesita Power BI
--             (venta con nombre de cliente, producto, categoria y region en una
--             sola fila) y respondemos tres preguntas de negocio con JOINs +
--             una consolidacion por canal con UNION ALL.
--
--  Motor    : PostgreSQL 16 (probado con ON_ERROR_STOP=1, corre sin errores).
--  Ejecucion: psql -U postgres -d ventas_tech_db -f m5_consultas_joins.sql
--
--  Regla de negocio: total_linea = cantidad * precio_unitario
-- =============================================================================


-- =============================================================================
--  CONSULTA 1 — Vista base del proyecto (INNER JOIN)
-- -----------------------------------------------------------------------------
--  Pregunta de negocio: "Necesito una sola tabla que combine toda la
--  informacion del negocio para armar el dashboard."
--
--  Se combinan las 4 tablas con INNER JOIN: partimos de `ventas` (hechos) y le
--  pegamos sus dimensiones. El INNER JOIN devuelve solo las ventas que tienen
--  cliente Y producto Y categoria validos — que, por las FOREIGN KEY del M3, son
--  TODAS las ventas (las 10 filas). Es la fuente de datos principal de Power BI.
--
--  Columnas para el dashboard:
--    - para AGRUPAR : ciudad, segmento, categoria  (dimensiones)
--    - para FILTRAR : fecha_venta, segmento, ciudad
--    - metrica      : total_venta = cantidad * precio_unitario
-- =============================================================================
SELECT
    v.fecha_venta,
    v.id_cliente,
    c.nombre                            AS cliente,
    c.ciudad                            AS region,          -- columna para agrupar/filtrar
    c.segmento                          AS segmento_cliente,-- columna para agrupar/filtrar
    p.nombre_producto                   AS producto,
    cat.nombre_categoria                AS categoria,
    v.cantidad,
    v.precio_unitario,
    (v.cantidad * v.precio_unitario)    AS total_venta
FROM ventas v
INNER JOIN clientes   c   ON c.id_cliente   = v.id_cliente
INNER JOIN productos  p   ON p.id_producto  = v.id_producto
INNER JOIN categorias cat ON cat.id_categoria = p.id_categoria
ORDER BY v.fecha_venta, v.id_venta;


-- =============================================================================
--  CONSULTA 2 — Clientes sin ventas (LEFT JOIN)
-- -----------------------------------------------------------------------------
--  Pregunta de negocio (area de CRM): "¿Que clientes se registraron pero todavia
--  NO compraron nada?" — para armar una campana de primera compra.
--
--  Tecnica: LEFT JOIN de `clientes` (todos) contra `ventas`. Los clientes sin
--  ninguna venta quedan con las columnas de `ventas` en NULL; con
--  WHERE v.id_venta IS NULL aislamos exactamente esos casos (clientes 6 y 7).
-- =============================================================================
SELECT
    c.id_cliente,
    c.nombre,
    c.email,
    c.fecha_registro
FROM clientes c
LEFT JOIN ventas v ON v.id_cliente = c.id_cliente
WHERE v.id_venta IS NULL          -- se queda solo con los clientes que nunca aparecen en ventas
ORDER BY c.id_cliente;


-- =============================================================================
--  CONSULTA 3 — Productos sin ventas (LEFT JOIN)
-- -----------------------------------------------------------------------------
--  Pregunta de negocio (area de producto): "¿Que articulos del catalogo no
--  tienen NINGUN movimiento?" — para decidir promociones o baja de catalogo.
--
--  Mismo patron que la Consulta 2, pero partiendo de `productos`. Sumamos un
--  JOIN a `categorias` para mostrar la categoria del producto. El LEFT JOIN se
--  hace contra `ventas`; WHERE v.id_venta IS NULL deja solo los productos que
--  nunca se vendieron (productos 7 y 8).
-- =============================================================================
SELECT
    p.id_producto,
    p.nombre_producto,
    cat.nombre_categoria   AS categoria,
    p.precio
FROM productos p
INNER JOIN categorias cat ON cat.id_categoria = p.id_categoria
LEFT  JOIN ventas     v   ON v.id_producto   = p.id_producto
WHERE v.id_venta IS NULL          -- productos que no aparecen en ninguna venta
ORDER BY p.id_producto;


-- =============================================================================
--  CONSULTA 4 — Consolidado por canal (UNION ALL)
-- -----------------------------------------------------------------------------
--  Pregunta de negocio: "¿Cuanto factura cada canal de venta?"
--
--  OJO: la columna `canal` NO existe en el esquema — se CREA aca como texto fijo
--  ('Online' / 'Presencial') dentro de cada SELECT. Partimos las ventas por un
--  criterio de negocio (periodo: hasta el 10/03 lo tratamos como Online, despues
--  como Presencial) y apilamos ambos SELECT con UNION ALL.
--
--  Usamos UNION ALL y NO UNION porque no queremos que se eliminen filas
--  repetidas: cada venta debe contarse una sola vez, aunque coincida con otra en
--  todos sus valores (UNION borraria esos duplicados legitimos y subcontaria).
--
--  Los dos SELECT devuelven la MISMA cantidad de columnas, en el mismo orden y
--  con tipos compatibles (date, decimal, text). Cerramos con GROUP BY canal para
--  obtener el total por cada origen.
-- =============================================================================
SELECT
    canal,
    COUNT(*)      AS cantidad_ventas,
    SUM(total)    AS total_por_canal
FROM (
    SELECT fecha_venta AS fecha,
           (cantidad * precio_unitario) AS total,
           'Online'     AS canal
    FROM ventas
    WHERE fecha_venta <= DATE '2024-03-10'
    UNION ALL
    SELECT fecha_venta AS fecha,
           (cantidad * precio_unitario) AS total,
           'Presencial' AS canal
    FROM ventas
    WHERE fecha_venta >  DATE '2024-03-10'
) AS ventas_por_canal
GROUP BY canal
ORDER BY total_por_canal DESC;


-- =============================================================================
--  HALLAZGOS  (analisis de los resultados obtenidos al ejecutar el script)
-- =============================================================================
--
--  1) Vista base (Consulta 1): el INNER JOIN devuelve las 10 ventas enriquecidas
--     con cliente, region, segmento, producto y categoria. Que sigan siendo 10
--     confirma la integridad referencial del M3: no hay ventas huerfanas. Esta
--     unica tabla es la que consume Power BI; ciudad/segmento/categoria sirven de
--     segmentadores y fecha_venta de filtro temporal.
--
--  2) Clientes sin ventas (Consulta 2): aparecen 2 de 7 clientes (Diego Fernandez
--     y Sofia Ramirez), un 29% de la base registrada que nunca compro. Son el
--     publico exacto de una campana de primera compra del area de CRM.
--
--  3) Productos sin ventas (Consulta 3): 2 de 8 productos del catalogo (Webcam
--     Full HD y Parlante BT Mini) no registran movimiento. El area de producto
--     puede decidir promocionarlos o darlos de baja. Nota: sin LEFT JOIN estos
--     casos serian invisibles — un INNER JOIN los ocultaria justamente por no
--     tener ventas, que es lo que queremos detectar.
--
--  4) Consolidado por canal (Consulta 4): al partir marzo en dos periodos, el
--     canal 'Online' (hasta el 10/03) concentra la mayor parte de la facturacion.
--     El valor de canal es inventado dentro del SELECT: muestra como etiquetar y
--     consolidar origenes sin que exista una columna fisica para ello.
--
-- =============================================================================
--  FIN DEL SCRIPT
-- =============================================================================
