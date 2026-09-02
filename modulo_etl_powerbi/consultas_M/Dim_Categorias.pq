// =============================================================================
//  CONSULTA: Dim_Categorias   (dimension de referencia — tabla limpia)
// -----------------------------------------------------------------------------
//  Pegar en: Inicio > Editor avanzado (sobre la consulta "categorias").
//  Renombrar la consulta a Dim_Categorias en el panel de la izquierda.
// =============================================================================
let
    // ORIGEN: leemos la hoja "categorias" del Excel via el parametro RutaDataset.
    Origen = Excel.Workbook(File.Contents(RutaDataset), null, true),
    HojaCategorias = Origen{[Item = "categorias", Kind = "Sheet"]}[Data],
    EncabezadosPromovidos = Table.PromoteHeaders(HojaCategorias, [PromoteAllScalars = true]),

    // TIPADO: id como numero entero (clave), nombre como texto.
    TiposAsignados = Table.TransformColumnTypes(EncabezadosPromovidos, {
        {"id_categoria", Int64.Type},
        {"nombre_categoria", type text}
    })
in
    TiposAsignados
