// ============================================================================
// Consulta: RutaDatos   (parametro)
// ----------------------------------------------------------------------------
// Evita la ruta de origen "hardcodeada". En lugar de escribir C:\Users\...\datos
// dentro de cada consulta, el path se lee desde la celda con nombre RutaDatos
// de la hoja Config del propio libro.
//
// Resultado: el evaluador puede mover la carpeta del proyecto a cualquier lado,
// cambiar una sola celda y las cuatro consultas siguen funcionando.
// Destino: Solo crear conexion (no se carga al modelo).
// ============================================================================
let
    // Excel.CurrentWorkbook lee rangos con nombre del libro que contiene la consulta
    TablaParametro = Excel.CurrentWorkbook(){[Name = "RutaDatos"]}[Content],

    ValorCrudo = Table.FirstValue(TablaParametro, ""),

    // Normaliza la barra final para poder concatenar el nombre de archivo sin riesgo
    RutaNormalizada =
        if Text.EndsWith(ValorCrudo, "\") then ValorCrudo
        else ValorCrudo & "\"
in
    RutaNormalizada
