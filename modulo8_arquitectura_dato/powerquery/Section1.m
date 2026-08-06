section Section1;

// ===========================================================================
//  TechStore - Flujo ETL del modelo de Inteligencia de Negocios
//  Modulo 8 - Arquitectura del dato, optimizacion y simulacion de escenarios
//
//  Origen 1 (hechos)      : fuentes\ventas_historico.csv
//  Origen 2 (dimensiones) : fuentes\maestros.xlsx
//  Destino                : Modelo de datos (Power Pivot), esquema en estrella
//
//  Este script fue editado a mano en el Editor avanzado: los pasos tienen
//  nombres de negocio (VentasFiltradas, TiposAjustados, ClavesUnicas...) en
//  lugar de los #"Filas filtradas1" que genera la interfaz.
// ===========================================================================


// ---------------------------------------------------------------------------
//  pRutaFuentes
//  Ninguna consulta lleva la ruta escrita adentro. Todas la piden aca, y aca
//  se lee de la celda con nombre RutaFuentes de la hoja Parametros, que a su
//  vez se arma con CELL("filename"): la ruta sigue al libro a donde lo muevan.
//  Cambiar de carpeta es cambiar un dato, no cinco consultas.
// ---------------------------------------------------------------------------
shared pRutaFuentes = let
    CeldaDelLibro = Excel.CurrentWorkbook(){[Name = "RutaFuentes"]}[Content],
    RutaLeida     = Text.Trim(Text.From(CeldaDelLibro{0}[Column1])),
    // si alguien la escribe sin la barra final, la consulta no se rompe
    RutaNormalizada = if Text.EndsWith(RutaLeida, "\") then RutaLeida else RutaLeida & "\"
in
    RutaNormalizada;


// ---------------------------------------------------------------------------
//  Hechos_Ventas
//  Tabla de hechos. Del CSV crudo (3.090 filas) quedan 3.000 utiles.
//  Orden deliberado de los pasos: limpiar texto -> filtrar -> deduplicar ->
//  tipar -> calcular. Tipar al final evita convertir filas que igual se van a
//  descartar, y evita que una celda vacia haga fallar la conversion.
// ---------------------------------------------------------------------------
shared Hechos_Ventas = let
    OrigenCSV = Csv.Document(
        File.Contents(pRutaFuentes & "ventas_historico.csv"),
        [Delimiter = ",", Columns = 11, Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),

    EncabezadosPromovidos = Table.PromoteHeaders(OrigenCSV, [PromoteAllScalars = true]),

    // el export trae numeros con espacios sobrantes ("  1 "): se limpian antes
    // de tipar, si no Number.From falla
    TextoLimpio = Table.TransformColumns(EncabezadosPromovidos, {}, Text.Trim),

    // un solo paso resuelve las tres reglas de negocio: fuera las anuladas y
    // fuera las filas sin unidades o sin cliente (nulos que romperian el 1:N)
    VentasFiltradas = Table.SelectRows(TextoLimpio, each
        [Estado] = "Confirmada"
        and [Unidades] <> null and [Unidades] <> ""
        and [ID_Cliente] <> null and [ID_Cliente] <> ""),

    // el sistema transaccional reenvia operaciones: la PK manda
    SinDuplicados = Table.Distinct(VentasFiltradas, {"ID_Venta"}),

    // Culture en-US fija el punto como separador decimal: el libro abre igual
    // en un Excel en espanol que en uno en ingles
    TiposAjustados = Table.TransformColumnTypes(SinDuplicados, {
        {"ID_Venta", type text},
        {"Fecha", type date},
        {"ID_Cliente", type text},
        {"ID_Producto", type text},
        {"Canal", type text},
        {"Estado", type text},
        {"Unidades", Int64.Type},
        {"Precio_Unitario", type number},
        {"Descuento", type number},
        {"Costo_Unitario", type number},
        {"Costo_Envio", type number}}, "en-US"),

    // Las cuatro metricas de fila se resuelven ACA y no con DAX. Es la decision
    // de optimizacion del Paso 4: si la multiplicacion ya viene hecha, las
    // medidas son SUM() sobre una columna comprimida y no SUMX() iterando
    // 3.000 filas en cada celda de cada tabla dinamica.
    VentaBruta   = Table.AddColumn(TiposAjustados, "Venta_Bruta",
                                   each [Unidades] * [Precio_Unitario], type number),
    VentaNeta    = Table.AddColumn(VentaBruta, "Venta_Neta",
                                   each [Venta_Bruta] * (1 - [Descuento]), type number),
    CostoTotal   = Table.AddColumn(VentaNeta, "Costo_Total",
                                   each [Unidades] * [Costo_Unitario], type number),
    MargenBruto  = Table.AddColumn(CostoTotal, "Margen_Bruto",
                                   each [Venta_Neta] - [Costo_Total] - [Costo_Envio], type number),

    // Estado ya vale "Confirmada" en todas las filas: como columna no aporta
    // nada y si ocupa memoria en el modelo
    ColumnasFinales = Table.RemoveColumns(MargenBruto, {"Estado"})
in
    ColumnasFinales;


// ---------------------------------------------------------------------------
//  Dim_Clientes
//  Dimension de clientes. ClavesUnicas es el paso que sostiene el modelo: sin
//  el, Power Pivot rechaza la relacion 1:N por clave duplicada.
// ---------------------------------------------------------------------------
shared Dim_Clientes = let
    OrigenLibro = Excel.Workbook(File.Contents(pRutaFuentes & "maestros.xlsx"), null, true),

    // se busca la tabla por nombre, no por posicion de hoja: si manana alguien
    // agrega una hoja adelante, la consulta sigue funcionando
    TablaClientes = OrigenLibro{[Item = "tbl_Clientes", Kind = "Table"]}[Data],

    TiposAjustados = Table.TransformColumnTypes(TablaClientes, {
        {"ID_Cliente", type text},
        {"Cliente", type text},
        {"Segmento", type text},
        {"Region", type text},
        {"Ciudad", type text},
        {"Alta", type date}}, "en-US"),

    SinNulos     = Table.SelectRows(TiposAjustados, each [ID_Cliente] <> null and [ID_Cliente] <> ""),
    ClavesUnicas = Table.Distinct(SinNulos, {"ID_Cliente"})
in
    ClavesUnicas;


// ---------------------------------------------------------------------------
//  Dim_Productos
//  Dimension de productos, misma logica que Clientes.
// ---------------------------------------------------------------------------
shared Dim_Productos = let
    OrigenLibro     = Excel.Workbook(File.Contents(pRutaFuentes & "maestros.xlsx"), null, true),
    TablaProductos  = OrigenLibro{[Item = "tbl_Productos", Kind = "Table"]}[Data],

    TiposAjustados = Table.TransformColumnTypes(TablaProductos, {
        {"ID_Producto", type text},
        {"Producto", type text},
        {"Categoria", type text},
        {"Marca", type text},
        {"Precio_Lista", type number},
        {"Costo_Estandar", type number}}, "en-US"),

    SinNulos     = Table.SelectRows(TiposAjustados, each [ID_Producto] <> null and [ID_Producto] <> ""),
    ClavesUnicas = Table.Distinct(SinNulos, {"ID_Producto"})
in
    ClavesUnicas;


// ---------------------------------------------------------------------------
//  Dim_Calendario
//  No se importa de ningun archivo: se genera con M a partir del rango real de
//  fechas de la tabla de hechos, redondeado a anios completos. Un calendario
//  con anios incompletos rompe cualquier calculo de acumulado o interanual.
// ---------------------------------------------------------------------------
shared Dim_Calendario = let
    FechaMinima = List.Min(Hechos_Ventas[Fecha]),
    FechaMaxima = List.Max(Hechos_Ventas[Fecha]),

    InicioAnio = #date(Date.Year(FechaMinima), 1, 1),
    FinAnio    = #date(Date.Year(FechaMaxima), 12, 31),

    ListaFechas = List.Dates(InicioAnio,
                             Duration.Days(FinAnio - InicioAnio) + 1,
                             #duration(1, 0, 0, 0)),

    TablaFechas  = Table.FromList(ListaFechas, Splitter.SplitByNothing(), {"Fecha"}),
    FechaTipada  = Table.TransformColumnTypes(TablaFechas, {{"Fecha", type date}}),

    ConAnio      = Table.AddColumn(FechaTipada, "Anio", each Date.Year([Fecha]), Int64.Type),
    ConTrimestre = Table.AddColumn(ConAnio, "Trimestre", each "Q" & Text.From(Date.QuarterOfYear([Fecha])), type text),
    ConMes       = Table.AddColumn(ConTrimestre, "Mes", each Date.Month([Fecha]), Int64.Type),
    ConNombreMes = Table.AddColumn(ConMes, "Nombre_Mes", each Date.MonthName([Fecha], "es-AR"), type text),
    ConAnioMes   = Table.AddColumn(ConNombreMes, "Anio_Mes",
                                   each Text.From(Date.Year([Fecha])) & "-" & Text.PadStart(Text.From(Date.Month([Fecha])), 2, "0"),
                                   type text),
    ConDiaSemana = Table.AddColumn(ConAnioMes, "Dia_Semana", each Date.DayOfWeekName([Fecha], "es-AR"), type text),
    ConFinDeSemana = Table.AddColumn(ConDiaSemana, "Es_Fin_Semana",
                                     each if Date.DayOfWeek([Fecha], Day.Monday) >= 5 then "Si" else "No",
                                     type text),

    // la PK va primera por prolijidad al mirar la tabla en Power Pivot
    OrdenDeColumnas = Table.SelectColumns(ConFinDeSemana,
        {"Fecha", "Anio", "Trimestre", "Mes", "Nombre_Mes", "Anio_Mes", "Dia_Semana", "Es_Fin_Semana"})
in
    OrdenDeColumnas;
