// ============================================================================
// LAS TRES DIMENSIONES DEL ESQUEMA EN ESTRELLA
// Cada bloque es una consulta independiente. Copiar cada let...in por separado
// en su propio Editor Avanzado.
// Destino de las tres: Solo crear conexion + Agregar al Modelo de datos
// ============================================================================


// ############################################################################
// Consulta: Clientes        Origen: maestras.xlsx, tabla tbl_Clientes
// ############################################################################
let
    LibroMaestras = Excel.Workbook(
        File.Contents(RutaDatos & "maestras.xlsx"), null, true
    ),

    // Se toma la TABLA con nombre, no la hoja: si alguien inserta filas arriba
    // la consulta no se rompe, cosa que si pasa al apuntar a la hoja entera.
    TablaClientes = LibroMaestras{[Item = "tbl_Clientes", Kind = "Table"]}[Data],

    TiposClientes = Table.TransformColumnTypes(
        TablaClientes,
        {
            {"ID_Cliente", type text},
            {"Cliente", type text},
            {"Region", type text},
            {"Segmento", type text},
            {"Antiguedad_Anios", Int64.Type}
        }
    ),

    // La clave primaria no admite nulos ni repetidos: si los tuviera, la
    // relacion 1:N contra la tabla de hechos fallaria al crearse.
    ClavesValidas = Table.SelectRows(
        TiposClientes, each [ID_Cliente] <> null and [ID_Cliente] <> ""
    ),
    DimClientes = Table.Distinct(ClavesValidas, {"ID_Cliente"})
in
    DimClientes


// ############################################################################
// Consulta: Productos       Origen: maestras.xlsx, tabla tbl_Productos
// ############################################################################
let
    LibroMaestras = Excel.Workbook(
        File.Contents(RutaDatos & "maestras.xlsx"), null, true
    ),

    TablaProductos = LibroMaestras{[Item = "tbl_Productos", Kind = "Table"]}[Data],

    TiposProductos = Table.TransformColumnTypes(
        TablaProductos,
        {
            {"ID_Producto", type text},
            {"Producto", type text},
            {"Categoria", type text},
            {"Subcategoria", type text},
            {"Costo_Unitario", Currency.Type},
            {"Precio_Lista", Currency.Type}
        }
    ),

    ClavesValidas = Table.SelectRows(
        TiposProductos, each [ID_Producto] <> null and [ID_Producto] <> ""
    ),

    // Columna de apoyo para ordenar y agrupar el margen teorico por producto
    ConMargenLista = Table.AddColumn(
        ClavesValidas, "Margen_Lista_Pct",
        each if [Precio_Lista] = 0 then null
             else ([Precio_Lista] - [Costo_Unitario]) / [Precio_Lista],
        Percentage.Type
    ),

    DimProductos = Table.Distinct(ConMargenLista, {"ID_Producto"})
in
    DimProductos


// ############################################################################
// Consulta: Calendario      Origen: NINGUNO - generada 100% en lenguaje M
// ----------------------------------------------------------------------------
// No hay archivo detras de esta tabla: se construye con List.Dates a partir del
// rango real de fechas de la tabla de hechos. Si manana llegan ventas de 2027,
// el calendario se extiende solo al actualizar. Es el ejemplo mas claro de por
// que conviene escribir M a mano en vez de depender de los botones.
// ############################################################################
let
    // Se leen los limites reales desde Ventas para no fijar fechas a mano
    FechaMinima = Date.StartOfYear(List.Min(Ventas[Fecha])),
    FechaMaxima = Date.EndOfYear(List.Max(Ventas[Fecha])),
    CantidadDias = Duration.Days(FechaMaxima - FechaMinima) + 1,

    ListaFechas = List.Dates(FechaMinima, CantidadDias, #duration(1, 0, 0, 0)),
    TablaBase = Table.FromList(
        ListaFechas, Splitter.SplitByNothing(), {"Fecha"}, null, ExtraValues.Error
    ),
    FechaTipada = Table.TransformColumnTypes(TablaBase, {{"Fecha", type date}}),

    ConAnio = Table.AddColumn(FechaTipada, "Anio", each Date.Year([Fecha]), Int64.Type),
    ConTrimestre = Table.AddColumn(
        ConAnio, "Trimestre", each "T" & Text.From(Date.QuarterOfYear([Fecha])), type text
    ),
    ConMes = Table.AddColumn(ConTrimestre, "Mes", each Date.Month([Fecha]), Int64.Type),

    // Nombre de mes forzado a es-AR: sin la cultura, el idioma depende de la
    // maquina que actualice y el reporte alterna entre "enero" y "January".
    ConNombreMes = Table.AddColumn(
        ConMes, "Nombre_Mes",
        each Text.Proper(Date.ToText([Fecha], "MMMM", "es-AR")), type text
    ),
    ConAnioMes = Table.AddColumn(
        ConNombreMes, "Anio_Mes",
        each Text.From([Anio]) & "-" & Text.PadStart(Text.From([Mes]), 2, "0"), type text
    ),
    ConDiaSemana = Table.AddColumn(
        ConAnioMes, "Dia_Semana",
        each Text.Proper(Date.ToText([Fecha], "dddd", "es-AR")), type text
    ),
    DimCalendario = Table.AddColumn(
        ConDiaSemana, "Es_Fin_De_Semana",
        each Date.DayOfWeek([Fecha], Day.Monday) >= 5, type logical
    )
in
    DimCalendario
