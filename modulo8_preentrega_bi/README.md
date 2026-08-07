# Módulo 8 — Cierre de las correcciones de la pre-entrega BI

Devolución recibida: **59% (desaprobada)**. Los puntos fuertes ya están bien (código M, medidas escritas,
estadística descriptiva, escenarios, documentación del esquema estrella). Las 4 correcciones son todas
la misma cosa: **el modelo está documentado pero no existe dentro de Excel**.

`MaterializarModelo.bas` construye ese modelo. Se corre una vez, tarda ~1 minuto.

---

## Qué resuelve cada corrección

| Corrección de la profe | Qué hace la macro |
|---|---|
| "Materializá las 4 consultas M en Power Query y tildá *Agregar estos datos al Modelo de datos*" | Crea `Hechos_Ventas`, `Dim_Clientes`, `Dim_Productos`, `Dim_Calendario` con `Queries.Add` y las carga con `CreateModelConnection:=True` (equivale a *Solo crear conexión* + *Agregar al modelo*). Quedan visibles en **Datos > Consultas y conexiones**. |
| "Pegá las 8 medidas DAX en el área de cálculo de Power Pivot" | Crea las 8 medidas en la tabla `Hechos_Ventas` vía `Model.ModelMeasures.Add`. 4 usan `VAR`/`RETURN`. |
| "Creá las tres relaciones 1:N en la Vista de Diagrama" | Crea las 3 relaciones dimensión → hechos con `Model.ModelRelationships.Add`. |
| "Agregá una tabla dinámica que use [Total Ventas] filtrada por Dim_Calendario[Anio]" | Crea la hoja `TD_Ventas` con esa tabla dinámica sobre el modelo. |

Lo único que **no** se puede automatizar por VBA es *Marcar como tabla de fechas* — son dos clics y están abajo.

---

## Pasos (hacelo en este orden)

1. **Abrí tu libro** `PreEntrega_ModeloBI_MurphyLleyton.xlsx` y guardalo primero como **`.xlsm`**
   (Archivo > Guardar como > *Libro de Excel habilitado para macros*). Sin esto no podés correr la macro.
   Al final lo volvés a guardar como `.xlsx`.

2. Abrí el editor de VBA con **Alt + F11**.

3. Menú **Insertar > Módulo**. Se abre una hoja de código en blanco.

4. Pegá adentro **todo el contenido de `MaterializarModelo.bas`**.
   (Si copiás el archivo entero, borrá la primera línea `Attribute VB_Name = ...` — esa línea sólo
   sirve cuando el archivo se importa con *Archivo > Importar*, y molesta si pegás a mano.)

5. Volvé a Excel (**Alt + F11** otra vez), andá a **Vista > Macros > Ver macros**, elegí
   `MaterializarModelo` y **Ejecutar**. Confirmá el cartel.
   - Si Excel pregunta por niveles de privacidad de los orígenes, elegí **Público** para los dos y aceptá.
   - Si aparece "Se necesitan credenciales", elegí **Acceso anónimo / Windows** y *Conectar*.

6. Cuando termine, hacé el paso manual que falta:
   **Power Pivot > Administrar > pestaña Diseñar > Marcar como tabla de fechas > columna `Fecha`**,
   con la tabla `Dim_Calendario` seleccionada.

7. Verificá (esto es exactamente lo que va a mirar la evaluadora):
   - **Datos > Consultas y conexiones** → tienen que aparecer las 4 consultas.
   - **Power Pivot > Administrar > Vista de diagrama** → `Hechos_Ventas` en el centro, las 3 dimensiones
     alrededor, flechas apuntando hacia los hechos. Acomodá las cajas en forma de estrella y guardá.
   - **Área de cálculo de `Hechos_Ventas`** → las 8 medidas.
   - **Hoja `TD_Ventas`** → los importes cambian por año. Si cambian, el filtro fluye: las relaciones andan.
   - **Hoja `Verificacion_Modelo`** → log de todo lo que se creó (sirve de evidencia).

8. **Archivo > Guardar como > `.xlsx`**. Excel avisa que se pierden las macros: aceptá.
   El modelo de datos, las consultas y la tabla dinámica **se conservan**. El nombre final:
   `PreEntrega_ModeloBI_MurphyLleyton.xlsx`.

9. Al subir, subí también la carpeta **`Fuentes`** (queda al lado del libro, con
   `Ventas_Historico.csv` y `Maestras.xlsx`). Es la fuente local que pide la consigna.

---

## El modelo que queda

```
                    Dim_Calendario (Fecha PK)
                              |
                             1:N
                              v
  Dim_Clientes  --1:N-->  Hechos_Ventas  <--1:N--  Dim_Productos
  (ID_Cliente PK)         (FKs + métricas)         (ID_Producto PK)
```

**Hechos_Ventas**: `ID_Venta, Fecha, ID_Cliente, ID_Producto, Canal, Cantidad, Venta_Neta, Costo_Total, Margen_Bruto`
(~4.800 filas después de la limpieza; `Precio_Unitario`, `Descuento` y `Costo_Unitario` se descartan en el
ETL una vez absorbidas en las métricas — eso es la reducción de huella de memoria que pide el Paso 4).

**Medidas** (las 4 últimas con `VAR`/`RETURN`):
`Total Ventas` · `Total Costo` · `Margen Bruto` · `Unidades Vendidas` ·
`% Margen Bruto` · `Ticket Promedio` · `Ventas Anio Anterior` · `Var % vs Anio Anterior`

---

## Sobre las fuentes

La macro genera el `.csv` transaccional y el `.xlsx` de maestras en `\Fuentes`, con semilla fija
(dataset reproducible) e incluyendo filas sucias a propósito — nulos en `ID_Cliente`, cantidades vacías
y cantidades en cero — para que la limpieza del código M sea real y no decorativa.

Las rutas quedan escritas como absolutas dentro del código M, resueltas al momento de ejecutar la macro.
Es deliberado: combinar una ruta dinámica leída de una celda con `File.Contents` dispara el
*Formula.Firewall* de Power Query y rompe la actualización. Si movés el libro de carpeta, volvés a correr
la macro y las rutas se reescriben solas.

---

## Si algo falla

La macro no se corta en silencio: todo queda en la hoja `Verificacion_Modelo`, línea por línea
(fuentes generadas, consultas creadas, tablas cargadas, cada relación y cada medida con OK o FALLO).
Si algo dice FALLO, ese renglón dice exactamente qué pasó.
