# Checkpoint — Arquitectura del Dato, Optimización de Modelos y Simulación

**Lleyton Murphy** · Entregable: `PreEntrega_ModeloBI_MurphyLleyton.xlsx`

Flujo completo de BI en Excel: ETL con Power Query sobre dos fuentes
heterogéneas, modelo dimensional en estrella en Power Pivot, medidas DAX
optimizadas con `VAR`/`RETURN`, diagnóstico estadístico con Analysis ToolPak y
simulación de escenarios para la dirección.

---

## Importante antes de entregar

**El libro requiere ~15 minutos de trabajo dentro de Excel** para quedar
completo. La hoja `PASOS_EN_EXCEL` los detalla uno por uno.

El motivo es técnico, no una omisión: el modelo de Power Pivot vive en
`xl/model/item.data`, una base de datos tabular en memoria que escribe el motor
xVelocity de Excel. No es XML editable y ninguna librería externa puede
generarla. Lo mismo vale para las medidas DAX, que se guardan dentro de ese
modelo. Todo lo demás está resuelto y verificado.

| Requisito de la rúbrica | Peso | Estado |
|---|:--:|---|
| Power Query + lenguaje M | 25% | Código escrito y documentado — **pegar en 4 consultas + 1 parámetro** |
| Modelado estrella en Power Pivot | 25% | Diagrama y relaciones especificadas — **crear en Vista de diagrama** |
| Medidas DAX con VAR/RETURN | 20% | 7 medidas escritas — **pegar en Power Pivot** |
| Simulación de escenarios + resumen | 15% | ✅ Completo en el archivo |
| Estadística con ToolPak | 10% | ✅ Completo en el archivo |
| Revisabilidad / entrega | 5% | ✅ Completo en el archivo |

---

## Contenido del libro

| Hoja | Contenido |
|---|---|
| `Inicio` | Portada e índice |
| `Config` | Valor a cargar en el parámetro `RutaDatos` de Power Query |
| `Modelo_Estrella` | Diagrama del esquema y tabla de relaciones 1:N a crear |
| `Simulacion` | Modelo financiero con las celdas cambiantes y Buscar Objetivo |
| `Resumen_Escenarios` | Comparativa Pesimista / Base / Optimista |
| `Estadistica_Descriptiva` | Reporte del ToolPak sobre `Importe_Neto` y `Margen_Bruto` |
| `Anexo_Codigo_M` | El parámetro y las 4 consultas en M, listos para copiar |
| `Anexo_Medidas_DAX` | Las 7 medidas, listas para copiar |
| `PASOS_EN_EXCEL` | Los pasos manuales, en orden |

## El flujo ETL

Dos fuentes heterogéneas, como pide la consigna:

- `datos/ventas_historicas.csv` — 1.218 filas transaccionales con suciedad
  deliberada: nulos en métricas, claves foráneas vacías, importes no numéricos,
  ventas anuladas, canal con mayúsculas inconsistentes y duplicados exactos.
  Quedan **1.149 filas limpias**.
- `datos/maestras.xlsx` — tablas maestras `tbl_Clientes` y `tbl_Productos`.

La dimensión `Calendario` no tiene archivo detrás: se construye entera en M con
`List.Dates`, tomando los límites reales de la tabla de hechos. Si mañana entran
ventas de 2027, el calendario se extiende solo al actualizar.

**Ruta no hardcodeada:** la carpeta de origen vive en un parámetro `RutaDatos` y
las cuatro consultas la concatenan con el nombre de archivo. Mover el proyecto
de carpeta es editar el parámetro, no cuatro consultas.

Se usa un parámetro y no una celda leída con `Excel.CurrentWorkbook()` a
propósito: con la celda, `Ventas` quedaría referenciando otra consulta **y**
accediendo a un origen externo en el mismo paso, que es la condición exacta que
dispara `Formula.Firewall` al actualizar. El parámetro se resuelve como valor
literal antes de evaluar la consulta, así que no cuenta como referencia.

**Pasos renombrados:** el paso que filtra las ventas anuladas se llama
`VentasFiltradas` en lugar de `#"Filas filtradas2"`, y el código tiene 69
comentarios `//` explicando cada decisión.

## El modelo dimensional

```
                    ┌──────────────┐
                    │  Calendario  │
                    │  PK Fecha    │
                    └──────┬───────┘
                           │ 1:N
   ┌──────────────┐  ┌─────┴────────┐  ┌───────────────┐
   │  Clientes    │──│    Ventas    │──│  Productos    │
   │ PK ID_Cliente│1:N│   (HECHOS)  │N:1│ PK ID_Producto│
   └──────────────┘  └──────────────┘  └───────────────┘
```

El filtro fluye siempre desde las dimensiones hacia los hechos.

## Optimización aplicada

La decisión de calcular `Importe_Neto` y `Costo_Total` en Power Query, del lado
del ETL, permite que las medidas usen `SUM` en vez de `SUMX`:

```dax
// Sin optimizar: dos iteradores recorriendo los hechos fila por fila
Margen Neto = SUMX ( Ventas, Ventas[Unidades] * Ventas[Precio_Unitario] * (1 - Ventas[Descuento]) )
            - SUMX ( Ventas, Ventas[Unidades] * Ventas[Costo_Unitario] )

// Optimizado: dos agregaciones sobre columnas ya materializadas
Margen Neto =
VAR VentasNetas   = [Total Ventas]
VAR CostoDeVentas = [Costo Total]
RETURN VentasNetas - CostoDeVentas
```

Combinado con el `Table.SelectColumns` final, la tabla de hechos entra al modelo
con 9 columnas en lugar de 13.

## La simulación

El modelo financiero se ancla en los datos reales: la base son las 7.900
unidades y los precios promedio del último año del dataset.

| Escenario | Tasa | Envío | Demanda | Beneficio neto |
|---|--:|--:|--:|--:|
| Pesimista | 18,0% | 9,0% | −15% | $38.229.103 |
| Base | 12,0% | 6,0% | 0% | $93.569.304 |
| Optimista | 9,0% | 4,5% | +20% | $149.520.048 |

Los tres escenarios están guardados de verdad en el libro (elemento
`<scenarios>` del XML de la hoja), así que aparecen al abrir
**Datos > Análisis de hipótesis > Administrador de escenarios**.

**Buscar Objetivo:** para alcanzar la meta de $126.000.000 de beneficio neto
hacen falta **9.181 unidades**, un 16,2% sobre la base.

## Reconstruir y verificar

```bash
pip install openpyxl
python3 build/datos_fuente.py      # CSV sucio + maestras.xlsx
python3 build/construir_libro.py   # el .xlsx
python3 build/verificar.py         # 65 controles automáticos
```

`verificar.py` comprueba la integridad del paquete OPC, que los 17 XML sean
válidos, que el bloque `<scenarios>` tenga los 3 escenarios en la posición que
exige el esquema, que ningún File.Contents tenga la ruta hardcodeada, que el
resumen coincida con el modelo financiero, que la estadística coincida con el
cálculo directo sobre el CSV, y que la cadena de fórmulas de la simulación
devuelva lo esperado (con un evaluador propio, incluido el resultado de Buscar
Objetivo).

Las fórmulas de estadística descriptiva están validadas contra valores conocidos
de Excel: para la serie 1..5, `SKEW` da 0 y `KURT` da −1,2.

## Estructura

```
checkpoint-modelo-bi/
├── PreEntrega_ModeloBI_MurphyLleyton.xlsx   <- el entregable
├── datos/
│   ├── ventas_historicas.csv                <- hechos (fuente sucia)
│   └── maestras.xlsx                        <- dimensiones Clientes y Productos
├── powerquery/
│   ├── 01_RutaDatos.m                       <- parámetro de ruta (no consulta)
│   ├── 02_Ventas.m                          <- tabla de hechos
│   └── 03_Dimensiones.m                     <- Clientes, Productos, Calendario
├── dax/medidas.dax                          <- las 7 medidas
└── build/
    ├── datos_fuente.py       genera el CSV y las maestras
    ├── estadistica.py        estimadores de Analysis ToolPak
    ├── construir_libro.py    arma el .xlsx e inyecta los escenarios
    └── verificar.py          65 controles sobre el resultado
```
