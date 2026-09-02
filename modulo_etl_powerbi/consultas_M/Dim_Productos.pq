// =============================================================================
//  CONSULTA: Dim_Productos   (dimension — 13 filas crudas -> 12 limpias)
// -----------------------------------------------------------------------------
//  Pegar en: Inicio > Editor avanzado (sobre la consulta "productos").
//  Renombrar la consulta a Dim_Productos.
//  Problemas que resuelve: 1 duplicado (id_producto = 103) + 2 nulos
//  (precio = CRITICO, categoria = no critico).
// =============================================================================
let
    // ---------------------------------------------------------------------
    //  ORIGEN
    // ---------------------------------------------------------------------
    Origen = Excel.Workbook(File.Contents(RutaDataset), null, true),
    HojaProductos = Origen{[Item = "productos", Kind = "Sheet"]}[Data],
    EncabezadosPromovidos = Table.PromoteHeaders(HojaProductos, [PromoteAllScalars = true]),

    // ---------------------------------------------------------------------
    //  TIPADO
    //  precio y costo como Numero decimal para poder sumar/promediar (nunca texto
    //  ni entero, se perderian los centavos). ids/stock/activo como enteros.
    // ---------------------------------------------------------------------
    TiposAsignados = Table.TransformColumnTypes(EncabezadosPromovidos, {
        {"id_producto", Int64.Type},
        {"nombre_producto", type text},
        {"categoria", type text},
        {"precio", type number},
        {"costo", type number},
        {"stock", Int64.Type},
        {"activo", Int64.Type}
    }),

    // ---------------------------------------------------------------------
    //  DEDUPLICACION
    //  id_producto = 103 llega repetido (mismo articulo cargado dos veces).
    //  Quitamos duplicados por id_producto: la PK de la dimension debe ser unica.
    // ---------------------------------------------------------------------
    SinDuplicados = Table.Distinct(TiposAsignados, {"id_producto"}),

    // ---------------------------------------------------------------------
    //  NULO NO CRITICO: categoria
    //  Se resuelve ANTES que el precio para que TODA fila tenga categoria y la
    //  imputacion del precio (que promedia por categoria) no falle. No se borra:
    //  el producto existe en el catalogo aunque falte clasificarlo -> "Sin Categoria".
    // ---------------------------------------------------------------------
    CategoriaResuelta = Table.ReplaceValue(SinDuplicados, null, "Sin Categoria", Replacer.ReplaceValue, {"categoria"}),

    // ---------------------------------------------------------------------
    //  NULO CRITICO: precio
    //  Sin precio no hay ingreso, pero borrar la fila descuadraria el conteo del
    //  catalogo y dejaria ventas huerfanas en el merge. Decision: IMPUTAR el
    //  precio faltante con el PROMEDIO de su misma categoria (calculado solo sobre
    //  productos con precio valido). Es un valor plausible y trazable, mucho mejor
    //  que 0 (sesga los KPIs a la baja) o que un numero inventado a mano.
    // ---------------------------------------------------------------------
    PromedioPorCategoria = Table.Group(
        Table.SelectRows(CategoriaResuelta, each [precio] <> null),
        {"categoria"},
        {{"precio_prom", each List.Average([precio]), type number}}
    ),
    ConPromedio = Table.NestedJoin(CategoriaResuelta, {"categoria"}, PromedioPorCategoria, {"categoria"}, "aux_prom", JoinKind.LeftOuter),
    ConPromedioExpandido = Table.ExpandTableColumn(ConPromedio, "aux_prom", {"precio_prom"}, {"precio_prom"}),
    PrecioImputado = Table.AddColumn(
        ConPromedioExpandido, "precio_final",
        each if [precio] = null then Number.Round([precio_prom], 2) else [precio],
        type number
    ),

    // Reemplazamos la columna original por la imputada y restauramos el orden.
    ColumnasAuxiliaresQuitadas = Table.RemoveColumns(PrecioImputado, {"precio", "precio_prom"}),
    PrecioRenombrado = Table.RenameColumns(ColumnasAuxiliaresQuitadas, {{"precio_final", "precio"}}),
    ColumnasReordenadas = Table.ReorderColumns(PrecioRenombrado, {
        "id_producto", "nombre_producto", "categoria", "precio", "costo", "stock", "activo"
    })
in
    ColumnasReordenadas
