-- =============================================================================
--  Ventas_Tech_DB  —  Script de Ingenieria de Datos (Back-End del Dashboard)
-- -----------------------------------------------------------------------------
--  Autor   : Lleyton Murphy
--  Curso   : Data Analyst — Coderhouse
--  Objetivo: Crear una base de datos relacional, limpia y normalizada (3NF)
--            para el modelo de Ventas de Tecnologia de la cadena "TechStore".
--
--  Contenido del script (3 secciones):
--    1. DDL          -> DROP + CREATE TABLE (dimensiones primero, hechos al final)
--    2. Restricciones-> PRIMARY KEY, FOREIGN KEY, NOT NULL, UNIQUE, DEFAULT
--    3. DML          -> Carga inicial de datos (categorias, clientes, productos, ventas)
--
--  Modelo:
--    categorias (1) ── (N) productos (1) ── (N) ventas (N) ── (1) clientes
--
--  Compatibilidad: escrito con SQL estandar para ejecutarse SIN ERRORES en
--  PostgreSQL (y compatible con SQL Server salvo ajustes menores de tipos).
--  El script es REPETIBLE: se puede ejecutar varias veces gracias a los
--  DROP TABLE IF EXISTS iniciales, que respetan el orden inverso de dependencias.
-- =============================================================================

-- -----------------------------------------------------------------------------
--  PASO 0 (opcional): Crear la base de datos.
--  Ejecutar esta linea por separado, conectado al servidor (no dentro de una
--  transaccion junto al resto). Luego conectarse a "Ventas_Tech_DB" y correr
--  el resto del script.
-- -----------------------------------------------------------------------------
-- CREATE DATABASE Ventas_Tech_DB;


-- =============================================================================
--  SECCION 1 — DDL
-- =============================================================================

-- ------------------------------------------------------------------
--  1.1  DROP TABLES (orden INVERSO a las dependencias)
--       Primero las tablas que TIENEN foreign keys (ventas, productos)
--       y al final las tablas referenciadas (clientes, categorias).
-- ------------------------------------------------------------------
DROP TABLE IF EXISTS ventas;
DROP TABLE IF EXISTS productos;
DROP TABLE IF EXISTS clientes;
DROP TABLE IF EXISTS categorias;


-- ------------------------------------------------------------------
--  1.2  CREATE TABLES (dimensiones primero, tabla de hechos al final)
-- ------------------------------------------------------------------

-- Dimension: CATEGORIAS -------------------------------------------------------
CREATE TABLE categorias (
    id_categoria     INTEGER      NOT NULL,
    nombre_categoria VARCHAR(50)  NOT NULL,
    descripcion      VARCHAR(200),
    CONSTRAINT pk_categorias PRIMARY KEY (id_categoria)
);

-- Dimension: CLIENTES ---------------------------------------------------------
CREATE TABLE clientes (
    id_cliente      INTEGER       NOT NULL,
    nombre          VARCHAR(100)  NOT NULL,
    email           VARCHAR(100),
    ciudad          VARCHAR(50),                 -- dimension GEOGRAFICA (para agrupar/filtrar)
    segmento        VARCHAR(20),                 -- dimension de SEGMENTACION (Consumo / PyME / Corporativo)
    fecha_registro  DATE          NOT NULL,
    CONSTRAINT pk_clientes    PRIMARY KEY (id_cliente),
    CONSTRAINT uq_clientes_email UNIQUE (email)
);

-- Dimension: PRODUCTOS --------------------------------------------------------
--  FK -> categorias. Precio monetario en DECIMAL(10,2) para permitir SUM/AVG.
CREATE TABLE productos (
    id_producto     INTEGER        NOT NULL,
    nombre_producto VARCHAR(100)   NOT NULL,
    id_categoria    INTEGER        NOT NULL,
    precio          DECIMAL(10,2)  NOT NULL,
    stock           INTEGER        DEFAULT 0,
    activo          SMALLINT       DEFAULT 1,   -- 1 = activo, 0 = inactivo
    CONSTRAINT pk_productos PRIMARY KEY (id_producto),
    CONSTRAINT fk_productos_categoria
        FOREIGN KEY (id_categoria) REFERENCES categorias (id_categoria)
);

-- Hechos: VENTAS --------------------------------------------------------------
--  Tabla central. FK -> clientes y FK -> productos garantizan integridad
--  referencial: no se puede registrar una venta de un cliente/producto inexistente.
CREATE TABLE ventas (
    id_venta        INTEGER        NOT NULL,
    id_cliente      INTEGER        NOT NULL,
    id_producto     INTEGER        NOT NULL,
    cantidad        INTEGER        NOT NULL,
    precio_unitario DECIMAL(10,2)  NOT NULL,
    fecha_venta     DATE           NOT NULL,
    CONSTRAINT pk_ventas PRIMARY KEY (id_venta),
    CONSTRAINT fk_ventas_cliente
        FOREIGN KEY (id_cliente)  REFERENCES clientes  (id_cliente),
    CONSTRAINT fk_ventas_producto
        FOREIGN KEY (id_producto) REFERENCES productos (id_producto)
);


-- =============================================================================
--  SECCION 3 — DML (Carga inicial de datos)
--  Orden: primero las tablas SIN dependencias (categorias, clientes),
--  luego productos (depende de categorias) y por ultimo ventas (depende de ambas).
-- =============================================================================

-- CATEGORIAS — 4 registros ----------------------------------------------------
INSERT INTO categorias (id_categoria, nombre_categoria, descripcion) VALUES (1, 'Computacion',    'Laptops, PCs y monitores');
INSERT INTO categorias (id_categoria, nombre_categoria, descripcion) VALUES (2, 'Accesorios',     'Perifericos y complementos');
INSERT INTO categorias (id_categoria, nombre_categoria, descripcion) VALUES (3, 'Audio',          'Auriculares y parlantes');
INSERT INTO categorias (id_categoria, nombre_categoria, descripcion) VALUES (4, 'Almacenamiento', 'Discos y memorias');

-- CLIENTES — 7 registros ------------------------------------------------------
--  Los clientes 6 y 7 se registraron pero AUN NO COMPRARON: son el caso que
--  aisla la Consulta 2 de M5 (LEFT JOIN + WHERE ... IS NULL — clientes sin ventas).
INSERT INTO clientes (id_cliente, nombre, email, ciudad, segmento, fecha_registro) VALUES (1, 'Maria Lopez',    'maria@mail.com',   'Buenos Aires', 'Consumo',     '2024-01-05');
INSERT INTO clientes (id_cliente, nombre, email, ciudad, segmento, fecha_registro) VALUES (2, 'Carlos Ruiz',    'carlos@mail.com',  'Cordoba',      'PyME',        '2024-01-10');
INSERT INTO clientes (id_cliente, nombre, email, ciudad, segmento, fecha_registro) VALUES (3, 'Ana Gomez',      'ana@mail.com',     'Rosario',      'Consumo',     '2024-02-01');
INSERT INTO clientes (id_cliente, nombre, email, ciudad, segmento, fecha_registro) VALUES (4, 'Pedro Sanz',     'pedro@mail.com',   'Mendoza',      'Corporativo', '2024-02-15');
INSERT INTO clientes (id_cliente, nombre, email, ciudad, segmento, fecha_registro) VALUES (5, 'Laura Torres',   'laura@mail.com',   'Tucuman',      'PyME',        '2024-03-01');
INSERT INTO clientes (id_cliente, nombre, email, ciudad, segmento, fecha_registro) VALUES (6, 'Diego Fernandez','diego@mail.com',   'La Plata',     'Consumo',     '2024-03-20');
INSERT INTO clientes (id_cliente, nombre, email, ciudad, segmento, fecha_registro) VALUES (7, 'Sofia Ramirez',  'sofia@mail.com',   'Salta',        'Corporativo', '2024-03-25');

-- PRODUCTOS — 8 registros -----------------------------------------------------
--  (id_producto, nombre_producto, id_categoria, precio, stock, activo)
--  Los productos 7 y 8 estan en el catalogo pero NO REGISTRAN VENTAS: son el
--  caso que aisla la Consulta 3 de M5 (LEFT JOIN + WHERE ... IS NULL).
INSERT INTO productos (id_producto, nombre_producto, id_categoria, precio, stock, activo) VALUES (1, 'Laptop Pro 15',      1, 1200.00, 15, 1);
INSERT INTO productos (id_producto, nombre_producto, id_categoria, precio, stock, activo) VALUES (2, 'Mouse Inalambrico',  2,   28.00, 80, 1);
INSERT INTO productos (id_producto, nombre_producto, id_categoria, precio, stock, activo) VALUES (3, 'Monitor 4K 27"',     1,  450.00, 12, 1);
INSERT INTO productos (id_producto, nombre_producto, id_categoria, precio, stock, activo) VALUES (4, 'Auriculares BT Pro', 3,  120.00, 35, 1);
INSERT INTO productos (id_producto, nombre_producto, id_categoria, precio, stock, activo) VALUES (5, 'SSD Externo 1TB',    4,  130.00, 18, 1);
INSERT INTO productos (id_producto, nombre_producto, id_categoria, precio, stock, activo) VALUES (6, 'Teclado Mecanico',   2,   95.00, 40, 1);
INSERT INTO productos (id_producto, nombre_producto, id_categoria, precio, stock, activo) VALUES (7, 'Webcam Full HD',     2,   65.00, 25, 1);
INSERT INTO productos (id_producto, nombre_producto, id_categoria, precio, stock, activo) VALUES (8, 'Parlante BT Mini',   3,   45.00, 30, 1);

-- VENTAS — 10 registros -------------------------------------------------------
--  (id_venta, id_cliente, id_producto, cantidad, precio_unitario, fecha_venta)
INSERT INTO ventas (id_venta, id_cliente, id_producto, cantidad, precio_unitario, fecha_venta) VALUES (1,  1, 1, 2, 1200.00, '2024-03-05');
INSERT INTO ventas (id_venta, id_cliente, id_producto, cantidad, precio_unitario, fecha_venta) VALUES (2,  2, 2, 5,   28.00, '2024-03-06');
INSERT INTO ventas (id_venta, id_cliente, id_producto, cantidad, precio_unitario, fecha_venta) VALUES (3,  3, 3, 1,  450.00, '2024-03-07');
INSERT INTO ventas (id_venta, id_cliente, id_producto, cantidad, precio_unitario, fecha_venta) VALUES (4,  1, 4, 2,  120.00, '2024-03-08');
INSERT INTO ventas (id_venta, id_cliente, id_producto, cantidad, precio_unitario, fecha_venta) VALUES (5,  4, 5, 3,  130.00, '2024-03-10');
INSERT INTO ventas (id_venta, id_cliente, id_producto, cantidad, precio_unitario, fecha_venta) VALUES (6,  2, 6, 4,   95.00, '2024-03-11');
INSERT INTO ventas (id_venta, id_cliente, id_producto, cantidad, precio_unitario, fecha_venta) VALUES (7,  5, 1, 1, 1200.00, '2024-03-12');
INSERT INTO ventas (id_venta, id_cliente, id_producto, cantidad, precio_unitario, fecha_venta) VALUES (8,  3, 2, 8,   28.00, '2024-03-13');
INSERT INTO ventas (id_venta, id_cliente, id_producto, cantidad, precio_unitario, fecha_venta) VALUES (9,  4, 4, 1,  120.00, '2024-03-14');
INSERT INTO ventas (id_venta, id_cliente, id_producto, cantidad, precio_unitario, fecha_venta) VALUES (10, 5, 3, 2,  450.00, '2024-03-15');


-- =============================================================================
--  VERIFICACION DE INTEGRIDAD
--  Ejecutar estas consultas tras el script para confirmar la carga de datos.
-- =============================================================================
-- SELECT * FROM categorias;   -- esperado: 4 filas
-- SELECT * FROM clientes;     -- esperado: 7 filas (2 sin ventas: clientes 6 y 7)
-- SELECT * FROM productos;    -- esperado: 8 filas (2 sin ventas: productos 7 y 8)
-- SELECT * FROM ventas;       -- esperado: 10 filas

-- Bonus (Modulo 5): cruce con JOIN para ver la venta con nombre de cliente/producto:
-- SELECT v.id_venta,
--        c.nombre            AS cliente,
--        p.nombre_producto   AS producto,
--        cat.nombre_categoria AS categoria,
--        v.cantidad,
--        v.precio_unitario,
--        (v.cantidad * v.precio_unitario) AS total,
--        v.fecha_venta
-- FROM ventas v
-- JOIN clientes   c   ON c.id_cliente   = v.id_cliente
-- JOIN productos  p   ON p.id_producto  = v.id_producto
-- JOIN categorias cat ON cat.id_categoria = p.id_categoria
-- ORDER BY v.id_venta;

-- =============================================================================
--  FIN DEL SCRIPT
-- =============================================================================
