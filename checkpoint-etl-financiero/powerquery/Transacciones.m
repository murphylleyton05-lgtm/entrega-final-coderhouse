// =========================================================================
//  Consulta de Power Query: "Transacciones"
//  Checkpoint: Flujo ETL, Analisis Financiero y Reporte Ejecutivo
//  Autor: Lleyton Murphy
//
//  Origen  : datos/transacciones_crudas.csv  (Datos > Obtener datos > Desde
//            texto/CSV). Si el CSV no esta junto al libro, la macro sustituye
//            el origen por la tabla tbl_Datos_Crudos incrustada, de modo que la
//            actualizacion NUNCA falla por un error de origen.
//  Destino : hoja Datos_Limpios + Modelo de datos.
//
//  IMPORTANTE: la consulta declara UN SOLO origen de datos. Leer la ruta desde
//  una celda del libro y pasarla a File.Contents mezclaria dos origenes en la
//  misma consulta y Power Query lo rechazaria con el error Formula.Firewall
//  ("references other queries or steps, so it may not directly access a data
//  source"). Por eso la ruta se inyecta como literal: ver Origen_M() en el
//  modulo Modulo_ETL_PowerQuery.
//
//  Transformaciones aplicadas (las 3 pedidas y algunas mas):
//    1. Eliminar columnas irrelevantes
//    2. Estandarizar textos (Clean + Trim + mayusculas/minusculas)
//    3. Corregir tipos de datos (fecha con cultura es-AR, numeros en-US)
//    4. Filtrar nulos, errores, importes invalidos y registros anulados
//    5. Quitar duplicados por clave de negocio
//    6. Enriquecer con columnas calculadas (Monto, Margen, Anio, Trimestre)
// =========================================================================
let
    // ---------------------------------------------------------------------
    // E X T R A C C I O N
    // ---------------------------------------------------------------------
    // La macro sustituye el marcador de la linea Origen por UNA de estas dos
    // expresiones (el marcador aparece una sola vez en toda la consulta):
    //
    //   CSV    Table.PromoteHeaders(
    //              Csv.Document(
    //                  File.Contents("C:\ruta\datos\transacciones_crudas.csv"),
    //                  [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    //              [PromoteAllScalars = true])
    //
    //   Libro  Excel.CurrentWorkbook(){[Name = "tbl_Datos_Crudos"]}[Content]
    //
    // Nunca las dos a la vez: un origen unico por consulta evita Formula.Firewall.
    Origen = @ORIGEN@,

    // ---------------------------------------------------------------------
    // T R A N S F O R M A C I O N
    // ---------------------------------------------------------------------

    // 1. Columnas irrelevantes para el analisis
    SinColumnas = Table.RemoveColumns(
        Origen, {"Notas_Internas", "Ref_Sistema"}, MissingField.Ignore),

    // 2. Estandarizacion de textos: quita caracteres no imprimibles, espacios
    //    sobrantes y unifica el uso de mayusculas ("  norte" y "NORTE" -> "Norte")
    Estandarizado = Table.TransformColumns(
        SinColumnas,
        {
            {"Region",    each if _ = null then null else Text.Proper(Text.Trim(Text.Clean(Text.From(_)))), type text},
            {"Categoria", each if _ = null then null else Text.Proper(Text.Trim(Text.Clean(Text.From(_)))), type text},
            {"Producto",  each if _ = null then null else Text.Proper(Text.Trim(Text.Clean(Text.From(_)))), type text},
            {"Vendedor",  each if _ = null then null else Text.Proper(Text.Trim(Text.Clean(Text.From(_)))), type text},
            {"Estado",    each if _ = null then null else Text.Upper(Text.Trim(Text.From(_))),              type text}
        },
        MissingField.Ignore),

    // 3a. Tipos de datos: la fecha viene como texto dd/mm/aaaa
    TipoFecha = Table.TransformColumnTypes(
        Estandarizado, {{"Fecha", type date}}, "es-AR"),

    // 3b. Tipos de datos: los importes usan el punto como separador decimal
    TipoNumeros = Table.TransformColumnTypes(
        TipoFecha,
        {
            {"ID_Transaccion", type text},
            {"Region", type text},
            {"Categoria", type text},
            {"Producto", type text},
            {"Vendedor", type text},
            {"Estado", type text},
            {"Unidades", Int64.Type},
            {"Precio_Unitario", type number},
            {"Costo_Unitario", type number}
        },
        "en-US"),

    // 4. Limpieza de nulos, errores de conversion y registros no vendibles
    SinErrores = Table.RemoveRowsWithErrors(TipoNumeros),

    SinNulos = Table.SelectRows(
        SinErrores,
        each [Fecha] <> null
            and [Region] <> null and [Region] <> ""
            and [Categoria] <> null and [Categoria] <> ""
            and [Unidades] <> null and [Unidades] > 0
            and [Precio_Unitario] <> null and [Precio_Unitario] > 0
            and [Costo_Unitario] <> null and [Costo_Unitario] >= 0),

    SinAnuladas = Table.SelectRows(SinNulos, each [Estado] <> "ANULADA"),

    // 5. Duplicados por clave de negocio
    SinDuplicados = Table.Distinct(SinAnuladas, {"ID_Transaccion"}),

    // 6. Columnas calculadas para el analisis
    ConMonto = Table.AddColumn(SinDuplicados, "Monto",
        each [Unidades] * [Precio_Unitario], type number),
    ConCosto = Table.AddColumn(ConMonto, "Costo_Total",
        each [Unidades] * [Costo_Unitario], type number),
    ConMargen = Table.AddColumn(ConCosto, "Margen",
        each [Monto] - [Costo_Total], type number),
    ConMargenPct = Table.AddColumn(ConMargen, "Margen_Pct",
        each if [Monto] = 0 then 0 else [Margen] / [Monto], type number),
    ConAnio = Table.AddColumn(ConMargenPct, "Anio",
        each Date.Year([Fecha]), Int64.Type),
    ConTrimestre = Table.AddColumn(ConAnio, "Trimestre",
        each "T" & Text.From(Date.QuarterOfYear([Fecha])), type text),
    ConMes = Table.AddColumn(ConTrimestre, "Mes",
        each Text.Proper(Date.MonthName([Fecha], "es-AR")), type text),

    Ordenado = Table.Sort(ConMes, {{"Fecha", Order.Ascending}}),

    // ---------------------------------------------------------------------
    // C A R G A
    // ---------------------------------------------------------------------
    Final = Table.SelectColumns(Ordenado, {
        "ID_Transaccion", "Fecha", "Anio", "Trimestre", "Mes",
        "Region", "Categoria", "Producto", "Vendedor",
        "Unidades", "Precio_Unitario", "Costo_Unitario",
        "Monto", "Costo_Total", "Margen", "Margen_Pct"})
in
    Final
