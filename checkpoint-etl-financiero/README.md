# Checkpoint — Flujo ETL, Análisis Financiero y Automatización del Reporte Ejecutivo

**Lleyton Murphy** · Entregable: un único libro de Excel habilitado para macros.

> **[`EntregaFinal_ETL_Financiero_Lleyton_Murphy.xlsm`](./EntregaFinal_ETL_Financiero_Lleyton_Murphy.xlsm)**

Herramienta que toma un archivo de transacciones crudo, lo depura con Power Query,
evalúa la viabilidad de un proyecto de expansión con PAGO / VNA / TIR y publica un
reporte ejecutivo interactivo que se actualiza, se formatea y se exporta a PDF con
un clic.

---

## Qué hay adentro

| Hoja | Visible | Contenido |
|---|:--:|---|
| `Reporte_Ejecutivo` | ✅ | Tablero: 5 KPI, tabla dinámica `TD_Margen`, 3 segmentadores y 6 botones |
| `Evaluación_Financiera` | ✅ | Préstamo (PAGO), flujo a 5 años, VAN / TIR / VA y conclusión calculada |
| `Datos_Limpios` | ⬚ | Salida de Power Query |
| `Datos_Crudos` | ⬚ | Tabla `tbl_Datos_Crudos`: respaldo del origen sucio |
| `Config` | ⬚ | Parámetros del ETL, prompt de IA y copia del código M |

## Cómo se usa

Abrir el `.xlsm` → **Habilitar contenido**. El libro se arma solo: crea la consulta,
la carga a la hoja y al Modelo de datos, construye la dinámica y los segmentadores,
aplica el formato y oculta las hojas de soporte.

| Botón | Macro |
|---|---|
| 1. ACTUALIZAR DATOS (ETL) | `Actualizar_Datos` |
| 2. APLICAR FORMATO | `Aplicar_Formato_Ejecutivo` |
| 3. EXPORTAR A PDF | `Exportar_PDF` |
| 4. LIMPIAR FILTROS | `Limpiar_Filtros` |
| 5. MOSTRAR/OCULTAR HOJAS | `Alternar_Hojas_Soporte` |
| 6. INICIALIZAR MODELO | `Inicializar_Proyecto` |

Detalle completo en **[docs/guia_de_uso.md](./docs/guia_de_uso.md)**.

## El flujo ETL

`datos/transacciones_crudas.csv` — 328 filas con mayúsculas inconsistentes, nulos,
importes no numéricos, fechas ilegibles, anuladas y duplicados — pasa por seis
transformaciones en M y quedan **260 filas limpias**:

1. eliminar columnas irrelevantes → `Table.RemoveColumns`
2. estandarizar textos → `Text.Proper(Text.Trim(Text.Clean(...)))`
3. corregir tipos → `Table.TransformColumnTypes` (fecha `es-AR`, números `en-US`)
4. filtrar nulos y errores → `Table.RemoveRowsWithErrors` + `Table.SelectRows`
5. quitar duplicados → `Table.Distinct` por `ID_Transaccion`
6. enriquecer → Monto, Costo_Total, Margen, Margen_Pct, Anio, Trimestre, Mes

Código completo: **[powerquery/Transacciones.m](./powerquery/Transacciones.m)**.

La consulta declara **un solo origen**, resuelto por la macro (CSV externo si está,
tabla incrustada si no). Así la actualización nunca falla por un archivo faltante
y se evita el error `Formula.Firewall` de Power Query.

## El modelo financiero

| Concepto | Valor |
|---|---|
| Inversión inicial / financiado | $500.000 · 70 % ($350.000 a TNA 12 %, 60 cuotas) |
| **Cuota mensual** `=PAGO()` | **$7.785,56** |
| **VAN** `=VNA()` al 15 % | **$108.455,39** |
| **TIR** `=TIR()` | **38,99 %** |
| Conclusión | Proyecto **viable**: se acepta |

Las ventas base salen del propio flujo ETL, así que el modelo se recalcula solo
cuando cambian los datos.

## Estructura del repositorio

```
checkpoint-etl-financiero/
├── EntregaFinal_ETL_Financiero_Lleyton_Murphy.xlsm   <- el entregable
├── datos/transacciones_crudas.csv                    <- origen sucio
├── powerquery/Transacciones.m                        <- consulta ETL
├── vba/                                              <- fuentes de los 9 módulos
│   ├── ThisWorkbook.cls          Hoja_*.cls
│   ├── Modulo_Automatizacion.bas
│   ├── Modulo_ETL_PowerQuery.bas
│   └── Modulo_Formato_IA.bas
├── docs/
│   ├── guia_de_uso.md
│   ├── prompt_chatgpt.md                             <- Paso 6: prompt y depuración
│   └── checklist_consigna.md                         <- respuestas a la rúbrica
└── build/                                            <- generación reproducible
    ├── build_xlsm.py         arma el libro completo
    ├── vba_writer.py         escribe el vbaProject.bin (MS-OVBA + OLE/CFB)
    ├── datos_sinteticos.py   genera el CSV sucio
    └── verificar.py          controles automáticos sobre el resultado
```

## Reconstruir el entregable

El `.xlsm` no se editó a mano: se genera por script, así que es reproducible y
todo el código fuente es revisable en texto plano.

```bash
pip install XlsxWriter oletools
python3 build/build_xlsm.py     # CSV + vbaProject.bin + .xlsm
python3 build/verificar.py      # controles sobre el archivo generado
```

`verificar.py` comprueba que el paquete OPC esté íntegro, que el código VBA
recuperado del `.xlsm` sea idéntico al de `vba/`, que el M incrustado reconstruya
`Transacciones.m` sin pérdidas, que los bloques del VBA estén balanceados, que
cada botón apunte a una macro existente y que la configuración de impresión y las
hojas ocultas sean las pedidas.
