# Checklist de entrega y respuestas a las preguntas de evaluación

## Respuestas a las 6 preguntas

### 1. El archivo entregado tiene la extensión .xlsm y contiene macros funcionales — **Sí**

`EntregaFinal_ETL_Financiero_Lleyton_Murphy.xlsm`, con `xl/vbaProject.bin` real
declarado en `[Content_Types].xml` como `application/vnd.ms-office.vbaProject`.

El proyecto VBA tiene **9 módulos**:

| Módulo | Tipo | Contenido |
|---|---|---|
| `ThisWorkbook` | documento | `Workbook_Open` → arranque automático |
| `Hoja_Reporte`, `Hoja_Finanzas`, `Hoja_Datos`, `Hoja_Crudos`, `Hoja_Config` | documento | módulos de hoja (CodeName) |
| `Modulo_Automatizacion` | estándar | macros de los botones |
| `Modulo_ETL_PowerQuery` | estándar | consulta M, carga, dinámica y segmentadores |
| `Modulo_Formato_IA` | estándar | formato corporativo (generado con IA, Paso 6) |

El código está **en módulos reales del editor de VBA**, no como texto en una celda.
Se comprueba con `Alt + F11`.

### 2. La consulta de Power Query se actualiza correctamente sin errores de origen — **Sí**

La consulta `Transacciones` declara **un único origen**, que la macro resuelve en
cada ejecución en este orden: ruta manual de `Config!C6` → `datos\transacciones_crudas.csv`
junto al libro → tabla `tbl_Datos_Crudos` incrustada. Siempre hay un origen válido,
así que la actualización nunca falla por archivo faltante.

Mantener un solo origen por consulta es lo que evita el error
`Formula.Firewall`, que aparecería si la consulta leyera la ruta desde una celda
y se la pasara a `File.Contents`.

### 3. Las funciones VAN, TIR y PAGO están correctamente formuladas y arrojan resultados coherentes — **Sí**

Hoja `Evaluación_Financiera`:

| Celda | Fórmula (como se ve en Excel en español) | Resultado |
|---|---|---|
| `D14` | `=-PAGO(D11; D12; D7)` | **$7.785,56** de cuota mensual |
| `D35` | `=VNA(D23; D32:H32) + C32` | **$108.455,39** |
| `D36` | `=TIR(C32:H32)` | **38,99 %** |
| `D37` | `=VA(D23; D10; -D17)` | $313.180,72 |

Modelo: inversión $500.000, 70 % financiado ($350.000) a TNA 12 % en 60 cuotas;
aporte propio $150.000 en el año 0 y flujo del accionista a 5 años.
Flujos: −150.000 / 55.059 / 66.938 / 79.767 / 93.622 / 108.586.

`VNA` descuenta desde el período 1 y se le suma aparte el flujo del año 0, que es
la forma correcta de usar la función. La conclusión de `B40` es una fórmula, no
texto fijo: se recalcula sola si se cambian los supuestos.

> **Conclusión:** el proyecto se **ACEPTA** — VAN positivo ($108.455) y TIR
> (39,0 %) muy por encima del costo de capital (15,0 %).

### 4. Existen al menos tres botones funcionales — **Sí, hay seis**

Controles de formulario sobre el tablero, cada uno con su macro asignada:
`Actualizar_Datos`, `Aplicar_Formato_Ejecutivo`, `Exportar_PDF`,
`Limpiar_Filtros`, `Alternar_Hojas_Soporte`, `Inicializar_Proyecto`.
Los tres obligatorios (Actualizar / Formato / PDF) son los botones 1, 2 y 3.

### 5. El reporte final está configurado para exportarse en una sola página de PDF — **Sí**

`Configurar_Impresion` fija, y el propio XML de la hoja ya trae:
`fitToPage="1"`, `FitToPagesWide = 1`, `FitToPagesTall = 1`, `Zoom = False`,
A4 horizontal, centrado horizontal y vertical, área de impresión `$A$1:$K$46`,
encabezado y pie de página con autor, fecha y numeración.

### 6. Las hojas de soporte están ocultas — **Sí**

`Datos_Crudos`, `Datos_Limpios` y `Config` quedan ocultas (`state="hidden"` ya en
el archivo, y `Ocultar_Hojas_Soporte` lo garantiza en cada apertura). Visibles
sólo `Reporte_Ejecutivo` y `Evaluación_Financiera`. El botón 5 permite mostrarlas
para revisión sin tener que tocar el menú de Excel.

---

## Checklist del entregable

| Requisito | Estado | Dónde |
|---|---|---|
| Archivo único `.xlsm` | ✅ | `EntregaFinal_ETL_Financiero_Lleyton_Murphy.xlsm` |
| Consulta Power Query con ≥ 3 transformaciones | ✅ (6) | `powerquery/Transacciones.m` |
| — eliminar columnas irrelevantes | ✅ | `Table.RemoveColumns` (`Notas_Internas`, `Ref_Sistema`) |
| — estandarizar textos | ✅ | `Text.Proper(Text.Trim(Text.Clean(...)))` |
| — corregir tipos de datos | ✅ | `Table.TransformColumnTypes` (fecha `es-AR`, números `en-US`) |
| — filtrar nulos | ✅ | `Table.RemoveRowsWithErrors` + `Table.SelectRows` |
| — quitar duplicados | ✅ | `Table.Distinct` por `ID_Transaccion` |
| — columnas calculadas | ✅ | Monto, Costo_Total, Margen, Margen_Pct, Anio, Trimestre, Mes |
| Carga al Modelo de datos | ✅ | `ETL_Cargar_Al_Modelo` (`CreateModelConnection:=True`) |
| Hoja `Evaluación_Financiera` con PAGO, VAN, TIR | ✅ | `D14`, `D35`, `D36` (+ `VA` en `D37`) |
| Conclusión de viabilidad | ✅ | `B40`, calculada con `SI(Y(...))` |
| Hoja `Reporte_Ejecutivo` con tabla dinámica | ✅ | `TD_Margen`: filas Región > Categoría, filtro Año |
| Agregaciones Suma / Promedio / % del total | ✅ | Ventas (suma), Margen (suma), Margen % (promedio), % del total |
| ≥ 2 segmentadores | ✅ (3) | Región, Categoría, Trimestre |
| 3 botones VBA funcionales | ✅ (6) | ver pregunta 4 |
| Exportación a PDF en una página | ✅ | `Exportar_PDF` |
| Hojas de soporte ocultas | ✅ | ver pregunta 6 |

## Sobre los datos

El origen `datos/transacciones_crudas.csv` es deliberadamente "sucio": **328 filas**
con mayúsculas y espacios inconsistentes, celdas vacías, importes no numéricos
(`N/D`, `s/d`), fechas ilegibles (`31/02/2024`), operaciones anuladas, IDs
duplicados y dos columnas irrelevantes.

Después del ETL quedan **260 filas limpias** (68 descartadas):
ventas $701.462,50 · margen $267.699,13 (38,2 %) · ticket promedio $2.697,93.

## Correcciones de la devolución anterior

| Corrección pedida | Cómo se resolvió |
|---|---|
| Integrar el VBA en un módulo real, no como texto en una celda | Proyecto VBA con 9 módulos dentro de `xl/vbaProject.bin` |
| Crear botones funcionales ligados a macros | 6 controles de formulario con macro asignada |
| Ocultar las hojas innecesarias | 3 hojas de soporte ocultas + botón para alternarlas |
| Guardar como `.xlsm` | El entregable es `.xlsm` habilitado para macros |
| Power Query (ETL) | Consulta `Transacciones` con 6 transformaciones documentadas |
| `=VNA()` y `=TIR()` en una proyección a 5 años | Sección 3 y 4 de `Evaluación_Financiera` |
| Reporte ejecutivo con dinámicas y segmentadores | `TD_Margen` + 3 segmentadores en el tablero |
