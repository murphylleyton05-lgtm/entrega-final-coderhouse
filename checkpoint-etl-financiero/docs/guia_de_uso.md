# Guía de uso

## 1. Abrir el libro

1. Descargá **`EntregaFinal_ETL_Financiero_Lleyton_Murphy.xlsm`** y, si querés
   probar el flujo desde el archivo externo, también la carpeta `datos/`
   (dejala al lado del `.xlsm`).
2. Abrilo en Excel. Como el archivo viene de internet, Windows lo marca como
   bloqueado: **clic derecho sobre el archivo → Propiedades → Desbloquear → Aceptar**.
   Sin este paso Excel abre en Vista protegida y no ejecuta nada.
3. Al abrirlo, Excel muestra la barra amarilla *"Advertencia de seguridad. Las
   macros se deshabilitaron"* → **Habilitar contenido**.

> Esa barra es exactamente la configuración que pide el Paso 1 de la consigna:
> Centro de confianza → *Deshabilitar todas las macros **con notificación***.
> Se comprueba en **Archivo → Opciones → Centro de confianza → Configuración del
> Centro de confianza → Configuración de macros**.

## 2. Qué pasa al habilitar las macros

`Workbook_Open` llama a `Inicializar_Proyecto`, que la primera vez:

1. crea la consulta de Power Query **Transacciones** con su código M,
2. la carga a la hoja `Datos_Limpios` **y** al Modelo de datos,
3. arma la tabla dinámica `TD_Margen` y los tres segmentadores,
4. aplica el formato corporativo,
5. oculta las hojas de soporte y deja el foco en el tablero.

Es **idempotente**: si el modelo ya está armado sólo lo refresca. La línea de
estado (celda `B4` del tablero) informa qué hizo y cuándo.

Si algo falla (por ejemplo, si las conexiones de datos siguen deshabilitadas),
el libro se abre igual y el detalle queda escrito en esa misma línea de estado:
alcanza con presionar el botón **1** o el **6** después de habilitar el contenido.

## 3. Los botones

| Botón | Macro | Qué hace |
|---|---|---|
| 1. ACTUALIZAR DATOS (ETL) | `Actualizar_Datos` | Resuelve el origen, `RefreshAll` en modo sincrónico, refresca la dinámica e informa cuántas filas cargó |
| 2. APLICAR FORMATO | `Aplicar_Formato_Ejecutivo` | Tipografía, paleta, bordes ejecutivos y formatos de moneda / porcentaje |
| 3. EXPORTAR A PDF | `Exportar_PDF` | Exporta el tablero a **una sola página** A4 horizontal, centrada, en la carpeta del libro |
| 4. LIMPIAR FILTROS | `Limpiar_Filtros` | Restablece segmentadores, filtros de dinámica y autofiltros |
| 5. MOSTRAR/OCULTAR HOJAS | `Alternar_Hojas_Soporte` | Alterna la visibilidad de `Datos_Crudos`, `Datos_Limpios` y `Config` |
| 6. INICIALIZAR MODELO | `Inicializar_Proyecto` | Reconstruye lo que falte (consulta, dinámica, segmentadores, formato) |

El PDF se guarda como `Reporte_Ejecutivo_aaaa-mm-dd_hhmm.pdf` junto al `.xlsm`.

## 4. El origen de datos

La consulta declara **un solo origen**, que la macro resuelve en cada ejecución:

1. la ruta escrita a mano en `Config!C6`, si el archivo existe;
2. si no, `datos\transacciones_crudas.csv` junto al libro;
3. si tampoco está, la tabla `tbl_Datos_Crudos` incrustada en la hoja `Datos_Crudos`.

Con cualquiera de las tres opciones la actualización termina sin errores de
origen, y la hoja `Config` (celda `C10`) deja registrado cuál se usó.

> **Por qué un solo origen y no un `try ... otherwise` con los dos:** si la
> consulta leyera la ruta de una celda del libro y se la pasara a `File.Contents`,
> Power Query estaría mezclando dos orígenes y aborta con
> `Formula.Firewall: ... may not directly access a data source`. Por eso la ruta
> se inyecta como literal desde VBA (`Origen_M`) y la consulta queda con un
> origen único.

## 5. Ver el código

- **Macros:** `Alt + F11` → proyecto `VBAProject` → `Modulo_Automatizacion`,
  `Modulo_ETL_PowerQuery`, `Modulo_Formato_IA`, más los módulos de cada hoja.
- **Power Query:** pestaña `Datos → Consultas y conexiones → Transacciones`
  (doble clic para abrir el editor y ver los pasos aplicados).
- Los mismos fuentes están, sin comprimir, en `vba/` y `powerquery/` del repo.

## 6. Reconstruir el archivo desde cero

```bash
pip install XlsxWriter oletools
python3 build/build_xlsm.py     # genera CSV + vbaProject.bin + .xlsm
python3 build/verificar.py      # 40 controles automáticos sobre el resultado
```
