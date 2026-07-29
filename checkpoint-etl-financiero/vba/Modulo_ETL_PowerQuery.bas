Attribute VB_Name = "Modulo_ETL_PowerQuery"
Option Explicit

'==========================================================================
'  Modulo_ETL_PowerQuery
'
'  Automatiza el flujo ETL completo:
'     E - conexion al origen (CSV externo, con respaldo en el libro)
'     T - transformaciones escritas en lenguaje M
'     L - carga a la hoja Datos_Limpios y al Modelo de datos
'
'  Ademas construye la tabla dinamica y los segmentadores del tablero.
'  Todo el modulo es idempotente: si el objeto ya existe, lo reutiliza.
'
'  ATENCION: este archivo se genera desde powerquery/Transacciones.m
'  (ver build/build_xlsm.py). No editar Script_M a mano.
'==========================================================================

Private Const CADENA_MASHUP As String = _
    "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location="


'--------------------------------------------------------------------------
'  Crea la consulta de Power Query o actualiza su codigo M si ya existe.
'--------------------------------------------------------------------------
Public Sub ETL_Asegurar_Consulta()
    Dim m As String
    Dim existe As Boolean

    ' El origen se resuelve en cada ejecucion: si el CSV volvio a aparecer (o
    ' desaparecio), la consulta se reescribe sola y nunca queda apuntando a un
    ' archivo inexistente.
    m = Replace(Script_M(), "@ORIGEN@", Origen_M())

    On Error Resume Next
    existe = (Len(ThisWorkbook.Queries(NOMBRE_CONSULTA).Name) > 0)
    On Error GoTo 0

    On Error GoTo SinSoporte
    If existe Then
        ThisWorkbook.Queries(NOMBRE_CONSULTA).Formula = m
    Else
        ThisWorkbook.Queries.Add Name:=NOMBRE_CONSULTA, Formula:=m, _
            Description:="Flujo ETL de transacciones - Checkpoint Coderhouse"
    End If
    Exit Sub

SinSoporte:
    ' Excel 2013 o anterior no expone el objeto Queries.
    Err.Clear
End Sub


'--------------------------------------------------------------------------
'  Carga el resultado de la consulta a la hoja Datos_Limpios.
'--------------------------------------------------------------------------
Public Sub ETL_Cargar_A_Hoja()
    Dim tabla As ListObject

    If Hoja_Datos.ListObjects.Count > 0 Then Exit Sub

    On Error GoTo Fallo
    Hoja_Datos.Visible = xlSheetVisible
    Set tabla = Hoja_Datos.ListObjects.Add( _
        SourceType:=0, _
        Source:=CADENA_MASHUP & NOMBRE_CONSULTA & ";Extended Properties=""""", _
        Destination:=Hoja_Datos.Range("$A$1"))

    With tabla.QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & NOMBRE_CONSULTA & "]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "tbl_Transacciones"
        .Refresh BackgroundQuery:=False
    End With
    Exit Sub

Fallo:
    Err.Clear
End Sub


'--------------------------------------------------------------------------
'  Carga la misma consulta al Modelo de datos (Power Pivot).
'--------------------------------------------------------------------------
Public Sub ETL_Cargar_Al_Modelo()
    Dim nombre As String
    Dim conexion As WorkbookConnection

    nombre = "Consulta - " & NOMBRE_CONSULTA & " (Modelo)"

    On Error Resume Next
    Set conexion = ThisWorkbook.Connections(nombre)
    On Error GoTo 0
    If Not conexion Is Nothing Then Exit Sub

    On Error GoTo Fallo
    ThisWorkbook.Connections.Add2 _
        Name:=nombre, _
        Description:="Carga de la consulta " & NOMBRE_CONSULTA & " al Modelo de datos", _
        ConnectionString:=CADENA_MASHUP & NOMBRE_CONSULTA & ";Extended Properties=""""", _
        CommandText:="SELECT * FROM [" & NOMBRE_CONSULTA & "]", _
        lCmdtype:=xlCmdSql, _
        CreateModelConnection:=True, _
        ImportRelationships:=False
    Exit Sub

Fallo:
    Err.Clear
End Sub


'--------------------------------------------------------------------------
'  Tabla dinamica del tablero. Responde a la pregunta de negocio:
'  que regiones y categorias concentran el margen.
'
'  Filas   : Region > Categoria
'  Filtro  : Anio
'  Valores : Suma de ventas, suma de margen, PROMEDIO de margen % y
'            PORCENTAJE del total  (las tres agregaciones pedidas)
'--------------------------------------------------------------------------
Public Sub ETL_Asegurar_TablaDinamica()
    Dim tabla As ListObject
    Dim cache As PivotCache
    Dim td As PivotTable
    Dim campo As PivotField

    If Hoja_Datos.ListObjects.Count = 0 Then ETL_Cargar_A_Hoja
    If Hoja_Datos.ListObjects.Count = 0 Then Exit Sub

    Set tabla = Hoja_Datos.ListObjects(1)
    If tabla.ListRows.Count = 0 Then Exit Sub

    On Error Resume Next
    Set td = Hoja_Reporte.PivotTables(NOMBRE_TD)
    On Error GoTo 0

    If td Is Nothing Then
        On Error GoTo Fallo
        Set cache = ThisWorkbook.PivotCaches.Create( _
            SourceType:=xlDatabase, SourceData:=tabla.Name)
        Set td = cache.CreatePivotTable( _
            TableDestination:=Hoja_Reporte.Range("B18"), TableName:=NOMBRE_TD)

        With td
            .PivotFields("Region").Orientation = xlRowField
            .PivotFields("Region").Position = 1
            .PivotFields("Categoria").Orientation = xlRowField
            .PivotFields("Categoria").Position = 2
            .PivotFields("Anio").Orientation = xlPageField

            Set campo = .AddDataField(.PivotFields("Monto"), "Ventas", xlSum)
            campo.NumberFormat = "$#,##0"

            Set campo = .AddDataField(.PivotFields("Margen"), "Margen", xlSum)
            campo.NumberFormat = "$#,##0"

            Set campo = .AddDataField(.PivotFields("Margen_Pct"), "Margen % prom.", xlAverage)
            campo.NumberFormat = "0.0%"

            Set campo = .AddDataField(.PivotFields("Monto"), "% del total", xlSum)
            campo.Calculation = xlPercentOfTotal
            campo.NumberFormat = "0.0%"

            .RowAxisLayout xlTabularRow
            .TableStyle2 = "PivotStyleMedium9"
            ' sin autoajuste: conserva el ancho de columna del tablero
            .HasAutoFormat = False
            .ShowTableStyleRowStripes = True
            .RepeatAllLabels xlRepeatLabels
        End With
    Else
        td.RefreshTable
    End If

    ETL_Asegurar_Segmentadores td
    Exit Sub

Fallo:
    Err.Clear
End Sub


'--------------------------------------------------------------------------
'  Segmentadores de datos (slicers) vinculados a la tabla dinamica.
'--------------------------------------------------------------------------
'  Las coordenadas estan en puntos y ubican los segmentadores a la derecha
'  de la tabla dinamica, dentro del area de impresion A1:K46.
Public Sub ETL_Asegurar_Segmentadores(ByVal td As PivotTable)
    Crear_Segmentador td, "Region", 325, 484, 145, 175
    Crear_Segmentador td, "Categoria", 325, 635, 145, 175
    Crear_Segmentador td, "Trimestre", 506, 484, 296, 105
End Sub


Private Sub Crear_Segmentador(ByVal td As PivotTable, ByVal campo As String, _
                              ByVal arriba As Single, ByVal izquierda As Single, _
                              ByVal ancho As Single, ByVal alto As Single)
    Dim cache As SlicerCache
    Dim nombre As String

    nombre = "Seg_" & campo

    On Error Resume Next
    Set cache = ThisWorkbook.SlicerCaches(nombre)
    On Error GoTo 0
    If Not cache Is Nothing Then Exit Sub

    On Error GoTo Fallo
    Set cache = ThisWorkbook.SlicerCaches.Add2(td, campo, nombre)
    cache.Slicers.Add _
        SlicerDestination:=Hoja_Reporte, _
        Name:=nombre & "_Ctrl", _
        Caption:=campo, _
        Top:=arriba, Left:=izquierda, Width:=ancho, Height:=alto
    Exit Sub

Fallo:
    Err.Clear
End Sub


'--------------------------------------------------------------------------
'  Expresion M del origen de datos. Devuelve el CSV si esta disponible y, si
'  no, la tabla de respaldo incrustada en el libro.
'--------------------------------------------------------------------------
Public Function Origen_M() As String
    Dim ruta As String

    ruta = Ruta_CSV()
    If Len(ruta) > 0 Then
        Origen_M = "Table.PromoteHeaders(Csv.Document(File.Contents(""" & _
                   Replace(ruta, """", """""") & """), " & _
                   "[Delimiter = "","", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]), " & _
                   "[PromoteAllScalars = true])"
    Else
        Origen_M = "Excel.CurrentWorkbook(){[Name = ""tbl_Datos_Crudos""]}[Content]"
    End If
End Function


'--------------------------------------------------------------------------
'  Busca el CSV de origen: primero la ruta manual de la hoja Config, despues
'  la subcarpeta datos\ del libro y por ultimo la carpeta del libro.
'--------------------------------------------------------------------------
Public Function Ruta_CSV() As String
    Dim manual As String
    Dim carpeta As String
    Dim candidato As String

    On Error Resume Next
    manual = Trim$(CStr(Hoja_Config.Range("par_RutaCSV").Value))
    On Error GoTo 0

    If Len(manual) > 0 Then
        If Len(Dir$(manual)) > 0 Then
            Ruta_CSV = manual
            Exit Function
        End If
    End If

    carpeta = Carpeta_Del_Libro()
    If Len(carpeta) = 0 Then Exit Function

    candidato = carpeta & "datos" & Application.PathSeparator & "transacciones_crudas.csv"
    If Len(Dir$(candidato)) > 0 Then
        Ruta_CSV = candidato
        Exit Function
    End If

    candidato = carpeta & "transacciones_crudas.csv"
    If Len(Dir$(candidato)) > 0 Then Ruta_CSV = candidato
End Function


'--------------------------------------------------------------------------
'  Deja en la hoja Config los parametros de referencia del flujo.
'--------------------------------------------------------------------------
Public Sub ETL_Escribir_Parametros()
    On Error Resume Next
    Hoja_Config.Range("par_CarpetaLibro").Value = Carpeta_Del_Libro()
    Hoja_Config.Range("C10").Value = IIf(Len(Ruta_CSV()) > 0, _
        "CSV externo: " & Ruta_CSV(), "Tabla incrustada: tbl_Datos_Crudos")
    On Error GoTo 0
End Sub


'--------------------------------------------------------------------------
'  Codigo M de la consulta. GENERADO - ver encabezado del modulo.
'--------------------------------------------------------------------------
Public Function Script_M() As String
    Dim m As String

    m = m & "// =========================================================================" & vbLf
    m = m & "//  Consulta de Power Query: ""Transacciones""" & vbLf
    m = m & "//  Checkpoint: Flujo ETL, Analisis Financiero y Reporte Ejecutivo" & vbLf
    m = m & "//  Autor: Lleyton Murphy" & vbLf
    m = m & "//" & vbLf
    m = m & "//  Origen  : datos/transacciones_crudas.csv  (Datos > Obtener datos > Desde" & vbLf
    m = m & "//            texto/CSV). Si el CSV no esta junto al libro, la macro sustituye" & vbLf
    m = m & "//            el origen por la tabla tbl_Datos_Crudos incrustada, de modo que la" & vbLf
    m = m & "//            actualizacion NUNCA falla por un error de origen." & vbLf
    m = m & "//  Destino : hoja Datos_Limpios + Modelo de datos." & vbLf
    m = m & "//" & vbLf
    m = m & "//  IMPORTANTE: la consulta declara UN SOLO origen de datos. Leer la ruta desde" & vbLf
    m = m & "//  una celda del libro y pasarla a File.Contents mezclaria dos origenes en la" & vbLf
    m = m & "//  misma consulta y Power Query lo rechazaria con el error Formula.Firewall" & vbLf
    m = m & "//  (""references other queries or steps, so it may not directly access a data" & vbLf
    m = m & "//  source""). Por eso la ruta se inyecta como literal: ver Origen_M() en el" & vbLf
    m = m & "//  modulo Modulo_ETL_PowerQuery." & vbLf
    m = m & "//" & vbLf
    m = m & "//  Transformaciones aplicadas (las 3 pedidas y algunas mas):" & vbLf
    m = m & "//    1. Eliminar columnas irrelevantes" & vbLf
    m = m & "//    2. Estandarizar textos (Clean + Trim + mayusculas/minusculas)" & vbLf
    m = m & "//    3. Corregir tipos de datos (fecha con cultura es-AR, numeros en-US)" & vbLf
    m = m & "//    4. Filtrar nulos, errores, importes invalidos y registros anulados" & vbLf
    m = m & "//    5. Quitar duplicados por clave de negocio" & vbLf
    m = m & "//    6. Enriquecer con columnas calculadas (Monto, Margen, Anio, Trimestre)" & vbLf
    m = m & "// =========================================================================" & vbLf
    m = m & "let" & vbLf
    m = m & "    // ---------------------------------------------------------------------" & vbLf
    m = m & "    // E X T R A C C I O N" & vbLf
    m = m & "    // ---------------------------------------------------------------------" & vbLf
    m = m & "    // La macro sustituye el marcador de la linea Origen por UNA de estas dos" & vbLf
    m = m & "    // expresiones (el marcador aparece una sola vez en toda la consulta):" & vbLf
    m = m & "    //" & vbLf
    m = m & "    //   CSV    Table.PromoteHeaders(" & vbLf
    m = m & "    //              Csv.Document(" & vbLf
    m = m & "    //                  File.Contents(""C:\ruta\datos\transacciones_crudas.csv"")," & vbLf
    m = m & "    //                  [Delimiter = "","", Encoding = 65001, QuoteStyle = QuoteStyle.Csv])," & vbLf
    m = m & "    //              [PromoteAllScalars = true])" & vbLf
    m = m & "    //" & vbLf
    m = m & "    //   Libro  Excel.CurrentWorkbook(){[Name = ""tbl_Datos_Crudos""]}[Content]" & vbLf
    m = m & "    //" & vbLf
    m = m & "    // Nunca las dos a la vez: un origen unico por consulta evita Formula.Firewall." & vbLf
    m = m & "    Origen = @ORIGEN@," & vbLf
    m = m & "" & vbLf
    m = m & "    // ---------------------------------------------------------------------" & vbLf
    m = m & "    // T R A N S F O R M A C I O N" & vbLf
    m = m & "    // ---------------------------------------------------------------------" & vbLf
    m = m & "" & vbLf
    m = m & "    // 1. Columnas irrelevantes para el analisis" & vbLf
    m = m & "    SinColumnas = Table.RemoveColumns(" & vbLf
    m = m & "        Origen, {""Notas_Internas"", ""Ref_Sistema""}, MissingField.Ignore)," & vbLf
    m = m & "" & vbLf
    m = m & "    // 2. Estandarizacion de textos: quita caracteres no imprimibles, espacios" & vbLf
    m = m & "    //    sobrantes y unifica el uso de mayusculas (""  norte"" y ""NORTE"" -> ""Norte"")" & vbLf
    m = m & "    Estandarizado = Table.TransformColumns(" & vbLf
    m = m & "        SinColumnas," & vbLf
    m = m & "        {" & vbLf
    m = m & "            {""Region"",    each if _ = null then null else Text.Proper(Text.Trim(Text.Clean(Text.From(_)))), type text}," & vbLf
    m = m & "            {""Categoria"", each if _ = null then null else Text.Proper(Text.Trim(Text.Clean(Text.From(_)))), type text}," & vbLf
    m = m & "            {""Producto"",  each if _ = null then null else Text.Proper(Text.Trim(Text.Clean(Text.From(_)))), type text}," & vbLf
    m = m & "            {""Vendedor"",  each if _ = null then null else Text.Proper(Text.Trim(Text.Clean(Text.From(_)))), type text}," & vbLf
    m = m & "            {""Estado"",    each if _ = null then null else Text.Upper(Text.Trim(Text.From(_))),              type text}" & vbLf
    m = m & "        }," & vbLf
    m = m & "        MissingField.Ignore)," & vbLf
    m = m & "" & vbLf
    m = m & "    // 3a. Tipos de datos: la fecha viene como texto dd/mm/aaaa" & vbLf
    m = m & "    TipoFecha = Table.TransformColumnTypes(" & vbLf
    m = m & "        Estandarizado, {{""Fecha"", type date}}, ""es-AR"")," & vbLf
    m = m & "" & vbLf
    m = m & "    // 3b. Tipos de datos: los importes usan el punto como separador decimal" & vbLf
    m = m & "    TipoNumeros = Table.TransformColumnTypes(" & vbLf
    m = m & "        TipoFecha," & vbLf
    m = m & "        {" & vbLf
    m = m & "            {""ID_Transaccion"", type text}," & vbLf
    m = m & "            {""Region"", type text}," & vbLf
    m = m & "            {""Categoria"", type text}," & vbLf
    m = m & "            {""Producto"", type text}," & vbLf
    m = m & "            {""Vendedor"", type text}," & vbLf
    m = m & "            {""Estado"", type text}," & vbLf
    m = m & "            {""Unidades"", Int64.Type}," & vbLf
    m = m & "            {""Precio_Unitario"", type number}," & vbLf
    m = m & "            {""Costo_Unitario"", type number}" & vbLf
    m = m & "        }," & vbLf
    m = m & "        ""en-US"")," & vbLf
    m = m & "" & vbLf
    m = m & "    // 4. Limpieza de nulos, errores de conversion y registros no vendibles" & vbLf
    m = m & "    SinErrores = Table.RemoveRowsWithErrors(TipoNumeros)," & vbLf
    m = m & "" & vbLf
    m = m & "    SinNulos = Table.SelectRows(" & vbLf
    m = m & "        SinErrores," & vbLf
    m = m & "        each [Fecha] <> null" & vbLf
    m = m & "            and [Region] <> null and [Region] <> """"" & vbLf
    m = m & "            and [Categoria] <> null and [Categoria] <> """"" & vbLf
    m = m & "            and [Unidades] <> null and [Unidades] > 0" & vbLf
    m = m & "            and [Precio_Unitario] <> null and [Precio_Unitario] > 0" & vbLf
    m = m & "            and [Costo_Unitario] <> null and [Costo_Unitario] >= 0)," & vbLf
    m = m & "" & vbLf
    m = m & "    SinAnuladas = Table.SelectRows(SinNulos, each [Estado] <> ""ANULADA"")," & vbLf
    m = m & "" & vbLf
    m = m & "    // 5. Duplicados por clave de negocio" & vbLf
    m = m & "    SinDuplicados = Table.Distinct(SinAnuladas, {""ID_Transaccion""})," & vbLf
    m = m & "" & vbLf
    m = m & "    // 6. Columnas calculadas para el analisis" & vbLf
    m = m & "    ConMonto = Table.AddColumn(SinDuplicados, ""Monto""," & vbLf
    m = m & "        each [Unidades] * [Precio_Unitario], type number)," & vbLf
    m = m & "    ConCosto = Table.AddColumn(ConMonto, ""Costo_Total""," & vbLf
    m = m & "        each [Unidades] * [Costo_Unitario], type number)," & vbLf
    m = m & "    ConMargen = Table.AddColumn(ConCosto, ""Margen""," & vbLf
    m = m & "        each [Monto] - [Costo_Total], type number)," & vbLf
    m = m & "    ConMargenPct = Table.AddColumn(ConMargen, ""Margen_Pct""," & vbLf
    m = m & "        each if [Monto] = 0 then 0 else [Margen] / [Monto], type number)," & vbLf
    m = m & "    ConAnio = Table.AddColumn(ConMargenPct, ""Anio""," & vbLf
    m = m & "        each Date.Year([Fecha]), Int64.Type)," & vbLf
    m = m & "    ConTrimestre = Table.AddColumn(ConAnio, ""Trimestre""," & vbLf
    m = m & "        each ""T"" & Text.From(Date.QuarterOfYear([Fecha])), type text)," & vbLf
    m = m & "    ConMes = Table.AddColumn(ConTrimestre, ""Mes""," & vbLf
    m = m & "        each Text.Proper(Date.MonthName([Fecha], ""es-AR"")), type text)," & vbLf
    m = m & "" & vbLf
    m = m & "    Ordenado = Table.Sort(ConMes, {{""Fecha"", Order.Ascending}})," & vbLf
    m = m & "" & vbLf
    m = m & "    // ---------------------------------------------------------------------" & vbLf
    m = m & "    // C A R G A" & vbLf
    m = m & "    // ---------------------------------------------------------------------" & vbLf
    m = m & "    Final = Table.SelectColumns(Ordenado, {" & vbLf
    m = m & "        ""ID_Transaccion"", ""Fecha"", ""Anio"", ""Trimestre"", ""Mes""," & vbLf
    m = m & "        ""Region"", ""Categoria"", ""Producto"", ""Vendedor""," & vbLf
    m = m & "        ""Unidades"", ""Precio_Unitario"", ""Costo_Unitario""," & vbLf
    m = m & "        ""Monto"", ""Costo_Total"", ""Margen"", ""Margen_Pct""})" & vbLf
    m = m & "in" & vbLf
    m = m & "    Final" & vbLf
    m = m & "" & vbLf

    Script_M = m
End Function
