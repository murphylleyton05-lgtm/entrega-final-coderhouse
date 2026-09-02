// =============================================================================
//  CONSULTA: Dim_Clientes   (dimension — 12 filas crudas -> 11 limpias)
// -----------------------------------------------------------------------------
//  Pegar en: Inicio > Editor avanzado (sobre la consulta "clientes").
//  Renombrar la consulta a Dim_Clientes.
//  Problemas que resuelve: 1 duplicado (id_cliente) + 2 nulos (email, ciudad).
// =============================================================================
let
    // ---------------------------------------------------------------------
    //  ORIGEN
    //  Conexion parametrizada al Excel (RutaDataset) para no hardcodear la ruta.
    // ---------------------------------------------------------------------
    Origen = Excel.Workbook(File.Contents(RutaDataset), null, true),
    HojaClientes = Origen{[Item = "clientes", Kind = "Sheet"]}[Data],
    EncabezadosPromovidos = Table.PromoteHeaders(HojaClientes, [PromoteAllScalars = true]),

    // ---------------------------------------------------------------------
    //  TIPADO
    //  fecha_registro DEBE ser Date (no texto): sin esto no se puede armar la
    //  tabla calendario ni la linea de tiempo del Modulo 7. id_cliente Int64
    //  para que la relacion 1:N con Fact_Ventas sea limpia.
    // ---------------------------------------------------------------------
    TiposAsignados = Table.TransformColumnTypes(EncabezadosPromovidos, {
        {"id_cliente", Int64.Type},
        {"nombre", type text},
        {"email", type text},
        {"ciudad", type text},
        {"fecha_registro", type date}
    }),

    // ---------------------------------------------------------------------
    //  DEDUPLICACION
    //  Quitamos duplicados POR id_cliente (no por todas las columnas a la vez):
    //  Dim_Clientes es una dimension y su clave primaria debe ser unica, o la
    //  relacion con la tabla de hechos degeneraria en muchos-a-muchos y las
    //  medidas DAX del Modulo 8 contarian doble. El origen duplica el alta al
    //  reintentar el registro; Table.Distinct conserva la primera aparicion.
    // ---------------------------------------------------------------------
    SinDuplicados = Table.Distinct(TiposAsignados, {"id_cliente"}),

    // ---------------------------------------------------------------------
    //  NULOS (no criticos -> se conservan las filas)
    //  email/ciudad nulos NO justifican borrar un cliente real: se perderia su
    //  historial de ventas asociado por id_cliente. Los reemplazamos por
    //  "Sin dato" para no romper filtros ni segmentadores y mantener el registro
    //  trazable. (Distinto criterio que un dato critico como el precio.)
    // ---------------------------------------------------------------------
    NulosResueltos = Table.ReplaceValue(SinDuplicados, null, "Sin dato", Replacer.ReplaceValue, {"email", "ciudad"})
in
    NulosResueltos
