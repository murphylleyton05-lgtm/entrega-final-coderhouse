# Módulo 8 — Pre-entrega BI (rehecha desde cero)

Devolución previa: **59% (desaprobada)**. Las 4 correcciones eran todas la misma:
el modelo estaba documentado en hojas pero no existía dentro de Excel.

Además, el `.xlsx` original tenía la parte interna de Power Query dañada — Excel devuelve
*"Las consultas de este libro están dañadas o se crearon en una versión más reciente de Excel"*
y se niega a crear consultas nuevas ahí. Por eso esta entrega arranca de un libro limpio.

## Los dos archivos

| Archivo | Qué es |
|---|---|
| `PreEntrega_ModeloBI_MurphyLleyton.xlsx` | Libro base limpio. Trae `LEEME`, `Codigo_M`, `Medidas_DAX` y `Modelo_Estrella`. |
| `MaterializarModelo.txt` | La macro. Construye todo lo que necesita el motor de Excel. |

`construir_libro.py` regenera el `.xlsx` si hace falta (requiere `openpyxl`).

## Pasos

1. Abrir `PreEntrega_ModeloBI_MurphyLleyton.xlsx`.
2. **Archivo > Guardar como > `.xlsm`** (Libro habilitado para macros).
3. **Alt+F11 > Insertar > Módulo**, pegar el contenido de `MaterializarModelo.txt`, cerrar con Alt+F11.
4. **Vista > Macros > `MaterializarModelo` > Ejecutar**. Confirmar.
   - Si pregunta por niveles de privacidad: **Público** en los dos orígenes.
5. **Power Pivot > Administrar > `Dim_Calendario` > Diseñar > Marcar como tabla de fechas > `Fecha`.**
   Único paso que VBA no puede hacer.
6. **Archivo > Guardar como > `.xlsx`**. Se pierden las macros, el modelo se conserva.
   Subir también la carpeta `Fuentes` que quedó al lado del libro.

## Qué crea la macro

**Fuentes locales** (`\Fuentes`, al lado del libro)
- `Ventas_Historico.csv` — 5.000 filas, semilla fija, con nulos y cantidades en cero
  a propósito para que la limpieza del ETL sea real.
- `Maestras.xlsx` — `Tbl_Clientes` (60) y `Tbl_Productos` (40).

**Power Query** — 4 consultas cargadas con *Solo crear conexión* + *Agregar al modelo de datos*:
`Hechos_Ventas`, `Dim_Clientes`, `Dim_Productos`, `Dim_Calendario`.
Paso renombrado a mano (`VentasFiltradas`) y comentarios `//` en el Editor avanzado.

**Power Pivot** — esquema estrella con 3 relaciones 1:N dimensión → hechos, y 8 medidas
en `Hechos_Ventas` (4 con `VAR`/`RETURN`).

**Verificación** — hoja `TD_Ventas` con `[Total Ventas]` por `Dim_Calendario[Anio]`.

**Estadística** — `Datos_Estadistica` (los datos crudos, por si se quiere correr el ToolPak
a mano) y `Estadistica_Descriptiva` con la salida completa: media, error típico, mediana, moda,
desvío, varianza, curtosis, asimetría, rango, mínimo, máximo, suma, cuenta y nivel de confianza,
sobre `Venta_Neta` y `Margen_Bruto`, más la lectura de los resultados.

**Simulación** — hoja `Simulacion` con tres celdas cambiantes (variación de demanda, tasa
logística, tasa financiera) y seis de resultado; Buscar Objetivo resuelto y documentado; los tres
escenarios en el Administrador de escenarios; y `Resumen de escenarios` en hoja aparte.

**Log** — `Verificacion_Modelo`, línea por línea, con OK o FALLO en cada relación y cada medida.

## Detalles de implementación

- Las rutas del código M se escriben absolutas al ejecutar la macro. Es deliberado: leer la ruta
  de una celda y pasarla a `File.Contents` dispara el *Formula.Firewall* de Power Query. Si el
  libro se mueve de carpeta, se vuelve a correr la macro y las rutas se reescriben.
- La macro es idempotente: borra consultas, conexiones, relaciones, medidas y hojas de salida
  homónimas antes de recrearlas. Se puede correr las veces que haga falta.
- Los bloques de código M se arman por concatenación incremental y no con `Array(...)` continuado,
  porque VBA acepta como máximo 25 líneas continuadas por instrucción.
- El `.bas` y el `.txt` son el mismo código; el `.txt` va sin la línea `Attribute VB_Name`
  para poder pegarlo directo en un módulo.
