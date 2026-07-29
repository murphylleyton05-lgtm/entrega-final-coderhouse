// ============================================================================
// Consulta: Ventas   (TABLA DE HECHOS)
// ----------------------------------------------------------------------------
// Origen: ventas_historicas.csv (transaccional historico, delimitado por ";")
// Destino: Solo crear conexion + Agregar al Modelo de datos
//
// Los pasos estan renombrados a mano en el Editor Avanzado: en lugar de los
// #"Tipo cambiado", #"Filas filtradas1", #"Personalizada agregada2" que genera
// la interfaz, cada paso dice que hace. El orden tambien esta pensado: se filtra
// ANTES de calcular columnas, para no gastar CPU en filas que se van a descartar.
// ============================================================================
let
    OrigenCSV = Csv.Document(
        File.Contents(RutaDatos & "ventas_historicas.csv"),
        [Delimiter = ";", Encoding = 65001, QuoteStyle = QuoteStyle.None]
    ),

    EncabezadosPromovidos = Table.PromoteHeaders(OrigenCSV, [PromoteAllScalars = true]),

    // --- 1. Tipado explicito -------------------------------------------------
    // Cultura "es-AR" para la fecha (dd/mm/aaaa) y "en-US" para los importes,
    // que en el CSV vienen con punto decimal. Sin esto Excel interpreta mal ambos.
    TiposCorregidos = Table.TransformColumnTypes(
        EncabezadosPromovidos,
        {
            {"ID_Venta", type text},
            {"ID_Cliente", type text},
            {"ID_Producto", type text},
            {"Unidades", Int64.Type},
            {"Descuento", type number},
            {"Canal", type text},
            {"Estado", type text}
        }
    ),
    FechasTipadas = Table.TransformColumnTypes(
        TiposCorregidos, {{"Fecha", type date}}, "es-AR"
    ),
    ImportesTipados = Table.TransformColumnTypes(
        FechasTipadas,
        {{"Precio_Unitario", type number}, {"Costo_Unitario", type number}},
        "en-US"
    ),

    // --- 2. Limpieza de nulos y errores --------------------------------------
    // "n/d" en Precio_Unitario quedo como error tras el tipado: se elimina aca.
    SinErroresDeTipado = Table.RemoveRowsWithErrors(
        ImportesTipados, {"Precio_Unitario", "Costo_Unitario", "Fecha"}
    ),

    SinNulosCriticos = Table.SelectRows(
        SinErroresDeTipado,
        each [Unidades] <> null
            and [ID_Cliente] <> null and [ID_Cliente] <> ""
            and [ID_Producto] <> null and [ID_Producto] <> ""
    ),

    // --- 3. Filtrado estrategico ---------------------------------------------
    // ESTE es el paso renombrado que pide la consigna (la interfaz lo llamaba
    // #"Filas filtradas2"). Regla de negocio: las ventas anuladas no son ventas.
    VentasFiltradas = Table.SelectRows(SinNulosCriticos, each [Estado] = "Confirmada"),

    // Duplicados exactos por ID_Venta: se conserva la primera ocurrencia
    SinDuplicados = Table.Distinct(VentasFiltradas, {"ID_Venta"}),

    // --- 4. Estandarizacion de texto -----------------------------------------
    // "online" / "ONLINE" / "Online" son el mismo canal: si no se unifica, la
    // dimension se rompe y el grafico muestra tres barras donde deberia haber una.
    CanalEstandarizado = Table.TransformColumns(
        SinDuplicados, {{"Canal", each Text.Proper(Text.Trim(_)), type text}}
    ),

    // --- 5. Columnas calculadas (solo sobre filas ya depuradas) --------------
    ConImporteBruto = Table.AddColumn(
        CanalEstandarizado, "Importe_Bruto",
        each [Unidades] * [Precio_Unitario], type number
    ),
    ConImporteNeto = Table.AddColumn(
        ConImporteBruto, "Importe_Neto",
        each [Importe_Bruto] * (1 - [Descuento]), Currency.Type
    ),
    ConCostoTotal = Table.AddColumn(
        ConImporteNeto, "Costo_Total",
        each [Unidades] * [Costo_Unitario], Currency.Type
    ),

    // --- 6. Reduccion del modelo ---------------------------------------------
    // Menos columnas = menos memoria en el motor xVelocity. Precio_Unitario y
    // Costo_Unitario ya estan absorbidos en las columnas de importe, y las
    // descriptivas (Cliente, Producto) viven en las dimensiones, no en los hechos.
    HechosFinales = Table.SelectColumns(
        ConCostoTotal,
        {
            "ID_Venta", "Fecha", "ID_Cliente", "ID_Producto",
            "Canal", "Unidades", "Importe_Bruto", "Importe_Neto", "Costo_Total"
        }
    )
in
    HechosFinales
