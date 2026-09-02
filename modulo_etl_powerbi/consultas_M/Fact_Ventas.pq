// =============================================================================
//  CONSULTA: Fact_Ventas   (tabla de HECHOS — 50 filas, limpia + Merge)
// -----------------------------------------------------------------------------
//  Pegar en: Inicio > Editor avanzado (sobre la consulta "ventas").
//  Renombrar la consulta a Fact_Ventas.
//  IMPORTANTE: esta consulta depende de Dim_Productos (para el Merge), asi que
//  pega y renombra Dim_Productos ANTES que esta.
// =============================================================================
let
    // ---------------------------------------------------------------------
    //  ORIGEN
    // ---------------------------------------------------------------------
    Origen = Excel.Workbook(File.Contents(RutaDataset), null, true),
    HojaVentas = Origen{[Item = "ventas", Kind = "Sheet"]}[Data],
    EncabezadosPromovidos = Table.PromoteHeaders(HojaVentas, [PromoteAllScalars = true]),

    // ---------------------------------------------------------------------
    //  TIPADO DE LA TABLA DE HECHOS
    //  fecha_venta como Date es lo mas critico: es el eje temporal que conecta
    //  con la tabla calendario del Modulo 8; si queda como texto, ninguna medida
    //  de inteligencia de tiempo (YTD, MoM) va a funcionar. Importes en Decimal
    //  para agregaciones DAX exactas; cantidad e ids como enteros.
    // ---------------------------------------------------------------------
    TiposAsignados = Table.TransformColumnTypes(EncabezadosPromovidos, {
        {"id_venta", Int64.Type},
        {"id_cliente", Int64.Type},
        {"id_producto", Int64.Type},
        {"cantidad", Int64.Type},
        {"precio_unitario", type number},
        {"descuento", type number},
        {"total_venta", type number},
        {"fecha_venta", type date},
        {"canal", type text}
    }),

    // ---------------------------------------------------------------------
    //  MERGE: enriquecer los hechos con atributos del producto
    //  Fact_Ventas trae id_producto pero no el nombre ni la categoria. Unimos con
    //  Dim_Productos por id_producto usando LEFT OUTER para CONSERVAR las 50 ventas
    //  aunque algun producto no matchee (una venta nunca debe desaparecer en la
    //  limpieza). Expandimos SOLO nombre_producto y categoria: el resto de los
    //  atributos del producto vive en Dim_Productos, evitando duplicar datos y
    //  respetando el esquema estrella.
    // ---------------------------------------------------------------------
    VentasCombinadas = Table.NestedJoin(TiposAsignados, {"id_producto"}, Dim_Productos, {"id_producto"}, "Dim_Productos", JoinKind.LeftOuter),
    VentasEnriquecidas = Table.ExpandTableColumn(VentasCombinadas, "Dim_Productos", {"nombre_producto", "categoria"}, {"nombre_producto", "categoria"})
in
    VentasEnriquecidas
