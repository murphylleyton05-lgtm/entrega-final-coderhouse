Attribute VB_Name = "modModeloBI"
' =====================================================================================
'  MaterializarModelo.bas  -  Pre-entrega Modulo 8 (Arquitectura del Dato / Power Pivot)
'  Lleyton Murphy - CoderHouse
'
'  QUE HACE ESTA MACRO (una sola corrida, sobre el libro activo):
'    1. Genera las fuentes locales en la subcarpeta \Fuentes del libro:
'         - Ventas_Historico.csv   (transaccional, ~5.000 filas, con nulos a limpiar)
'         - Maestras.xlsx          (Tbl_Clientes y Tbl_Productos)
'    2. Crea las 4 consultas de Power Query (codigo M) y las carga AL MODELO DE DATOS
'       con "Crear solo conexion" + "Agregar estos datos al Modelo de datos".
'    3. Crea las 3 relaciones 1:N (dimension -> hechos).
'    4. Crea las 8 medidas DAX en la tabla Hechos_Ventas (4 con VAR/RETURN).
'    5. Crea la tabla dinamica de verificacion [Total Ventas] x Dim_Calendario[Anio].
'    6. Deja el codigo M generado y un log de verificacion en hojas del libro.
'
'  NOTA DE LECTURA: dentro de este modulo el caracter "^" se reemplaza por comillas
'  dobles al construir el codigo M (funcion Q). Se hace asi para que el script M
'  sea legible en el .bas; el codigo que realmente se carga en Power Query queda
'  volcado tal cual en la hoja "Codigo_M_Modelo".
' =====================================================================================

Option Explicit

Private Const NOM_HECHOS As String = "Hechos_Ventas"
Private Const NOM_CLIENTES As String = "Dim_Clientes"
Private Const NOM_PRODUCTOS As String = "Dim_Productos"
Private Const NOM_CALENDARIO As String = "Dim_Calendario"
Private Const HOJA_LOG As String = "Verificacion_Modelo"
Private Const HOJA_M As String = "Codigo_M_Modelo"
Private Const HOJA_TD As String = "TD_Ventas"

Private Const ANIO_INI As Long = 2023
Private Const ANIO_FIN As Long = 2025
Private Const FILAS_VENTAS As Long = 5000

Private gLog As Collection
Private gRutaFuentes As String

' -------------------------------------------------------------------------------------
' PUNTO DE ENTRADA
' -------------------------------------------------------------------------------------
Public Sub MaterializarModelo()
    Dim wb As Workbook
    Dim t0 As Single

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "No hay ningun libro activo.", vbCritical
        Exit Sub
    End If

    If Len(wb.Path) = 0 Then
        MsgBox "Primero guarda el libro en una carpeta (las fuentes se crean al lado del archivo)." & vbLf & _
               "Guardalo y volve a ejecutar la macro.", vbCritical
        Exit Sub
    End If

    If MsgBox("Se va a construir el modelo de Power Pivot sobre el libro:" & vbLf & vbLf & _
              wb.Name & vbLf & wb.Path & vbLf & vbLf & _
              "Se crearan las consultas, las relaciones, las medidas DAX y una tabla dinamica." & vbLf & _
              "Las consultas/medidas/relaciones existentes con estos mismos nombres se reemplazan." & vbLf & vbLf & _
              "Continuar?", vbQuestion + vbYesNo, "Materializar modelo BI") <> vbYes Then Exit Sub

    t0 = Timer
    Set gLog = New Collection

    On Error GoTo Fallo
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    gRutaFuentes = wb.Path & Application.PathSeparator & "Fuentes"

    Registrar "Libro destino: " & wb.FullName
    GenerarFuentes wb
    LimpiarModeloPrevio wb
    CrearConsultas wb
    CargarAlModelo wb
    CrearRelaciones wb
    CrearMedidas wb
    CrearTablaDinamica wb
    VolcarCodigoM wb
    EscribirLog wb

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Modelo materializado en " & Format$(Timer - t0, "0.0") & " s." & vbLf & vbLf & _
           "Revisa la hoja '" & HOJA_LOG & "'." & vbLf & vbLf & _
           "PASO MANUAL QUE FALTA (no se puede automatizar por VBA):" & vbLf & _
           "Power Pivot > Administrar > tabla Dim_Calendario > pestania Diseniar >" & vbLf & _
           "Marcar como tabla de fechas > columna Fecha." & vbLf & vbLf & _
           "Despues: Archivo > Guardar como > .xlsx (el modelo se conserva).", _
           vbInformation, "Listo"
    Exit Sub

Fallo:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Registrar "ERROR " & Err.Number & ": " & Err.Description
    On Error Resume Next
    EscribirLog wb
    MsgBox "Se corto en: " & Err.Description & " (error " & Err.Number & ")" & vbLf & vbLf & _
           "El detalle de lo que si se completo esta en la hoja '" & HOJA_LOG & "'.", vbCritical
End Sub

' -------------------------------------------------------------------------------------
' 1. FUENTES LOCALES
' -------------------------------------------------------------------------------------
Private Sub GenerarFuentes(wb As Workbook)
    If Dir(gRutaFuentes, vbDirectory) = "" Then MkDir gRutaFuentes
    GenerarCsvVentas
    GenerarMaestras
    Registrar "Fuentes generadas en: " & gRutaFuentes
End Sub

Private Sub GenerarCsvVentas()
    Dim ff As Integer, r As Long, ruta As String
    Dim fec As Date, idc As Long, idp As Long, cant As Long
    Dim precio As Double, costo As Double, desc As Double, base As Double
    Dim sCant As String, sCli As String
    Dim descartadas As Long

    ruta = gRutaFuentes & Application.PathSeparator & "Ventas_Historico.csv"
    ff = FreeFile

    Rnd -1
    Randomize 42                      ' semilla fija -> dataset reproducible

    Open ruta For Output As #ff
    Print #ff, "ID_Venta,Fecha,ID_Cliente,ID_Producto,Canal,Cantidad,Precio_Unitario,Descuento,Costo_Unitario"

    For r = 1 To FILAS_VENTAS
        fec = DateSerial(ANIO_INI, 1, 1) + Int(Rnd * (DateSerial(ANIO_FIN, 12, 31) - DateSerial(ANIO_INI, 1, 1) + 1))
        idc = 1 + Int(Rnd * 60)
        idp = 1 + Int(Rnd * 40)

        ' distribucion con cola a la derecha: la mayoria son ordenes chicas y
        ' un 4% son ordenes mayoristas grandes (genera el sesgo/curtosis del analisis)
        If Rnd < 0.04 Then
            cant = 20 + Int(Rnd * 45)
        Else
            cant = 1 + Int(Rnd * 8)
        End If

        base = 14500 + idp * 3700
        precio = base * (0.9 + 0.2 * Rnd)
        costo = precio * (0.55 + 0.15 * Rnd)
        desc = Choose(1 + Int(Rnd * 4), 0#, 0.05, 0.1, 0.15)

        sCli = "C" & Format$(idc, "000")
        sCant = CStr(cant)

        ' suciedad controlada: el ETL de Power Query tiene que poder limpiarla
        If r Mod 37 = 0 Then sCli = "": descartadas = descartadas + 1
        If r Mod 53 = 0 Then sCant = "": descartadas = descartadas + 1
        If r Mod 71 = 0 Then sCant = "0": descartadas = descartadas + 1

        Print #ff, "V" & Format$(r, "00000") & "," & _
                   Format$(fec, "yyyy-mm-dd") & "," & _
                   sCli & "," & _
                   "P" & Format$(idp, "000") & "," & _
                   Choose(1 + Int(Rnd * 3), "Online", "Sucursal", "Mayorista") & "," & _
                   sCant & "," & _
                   Num(precio, 2) & "," & _
                   Num(desc, 2) & "," & _
                   Num(costo, 2)
    Next r

    Close #ff
    Registrar "Ventas_Historico.csv: " & FILAS_VENTAS & " filas (aprox. " & descartadas & " con nulos/cantidad invalida para limpiar en M)"
End Sub

Private Sub GenerarMaestras()
    Dim wbM As Workbook, ws As Worksheet, i As Long, ruta As String
    Dim cats As Variant, marcas As Variant, segs As Variant, provs As Variant

    cats = Array("Notebooks", "Perifericos", "Monitores", "Almacenamiento", "Redes")
    marcas = Array("TechStore", "Lenova", "Delta", "AserX", "Kingsom")
    segs = Array("Retail", "Mayorista", "Corporativo")
    provs = Array("Buenos Aires", "CABA", "Cordoba", "Santa Fe", "Mendoza", "Tucuman")

    ruta = gRutaFuentes & Application.PathSeparator & "Maestras.xlsx"

    Set wbM = Workbooks.Add(xlWBATWorksheet)

    Set ws = wbM.Worksheets(1)
    ws.Name = "Clientes"
    ws.Range("A1:E1").Value = Array("ID_Cliente", "Cliente", "Segmento", "Provincia", "Pais")
    For i = 1 To 60
        ws.Cells(i + 1, 1).Value = "C" & Format$(i, "000")
        ws.Cells(i + 1, 2).Value = "Cliente " & Format$(i, "000")
        ws.Cells(i + 1, 3).Value = segs((i - 1) Mod 3)
        ws.Cells(i + 1, 4).Value = provs((i - 1) Mod 6)
        ws.Cells(i + 1, 5).Value = "Argentina"
    Next i
    ws.ListObjects.Add(xlSrcRange, ws.Range("A1:E61"), , xlYes).Name = "Tbl_Clientes"

    Set ws = wbM.Worksheets.Add(After:=wbM.Worksheets(wbM.Worksheets.Count))
    ws.Name = "Productos"
    ws.Range("A1:E1").Value = Array("ID_Producto", "Producto", "Categoria", "Marca", "Precio_Lista")
    For i = 1 To 40
        ws.Cells(i + 1, 1).Value = "P" & Format$(i, "000")
        ws.Cells(i + 1, 2).Value = cats((i - 1) Mod 5) & " Serie " & Format$(i, "000")
        ws.Cells(i + 1, 3).Value = cats((i - 1) Mod 5)
        ws.Cells(i + 1, 4).Value = marcas((i - 1) Mod 5)
        ws.Cells(i + 1, 5).Value = 14500 + i * 3700
    Next i
    ws.ListObjects.Add(xlSrcRange, ws.Range("A1:E41"), , xlYes).Name = "Tbl_Productos"

    wbM.SaveAs Filename:=ruta, FileFormat:=51
    wbM.Close SaveChanges:=False
    Registrar "Maestras.xlsx: Tbl_Clientes (60) y Tbl_Productos (40)"
End Sub

' -------------------------------------------------------------------------------------
' 2. LIMPIEZA DE UNA CORRIDA ANTERIOR (idempotencia)
' -------------------------------------------------------------------------------------
Private Sub LimpiarModeloPrevio(wb As Workbook)
    Dim i As Long, nombres As Variant, n As Variant

    nombres = Array(NOM_HECHOS, NOM_CLIENTES, NOM_PRODUCTOS, NOM_CALENDARIO)

    On Error Resume Next
    ' medidas
    For i = wb.Model.ModelMeasures.Count To 1 Step -1
        wb.Model.ModelMeasures(i).Delete
    Next i
    ' relaciones
    For i = wb.Model.ModelRelationships.Count To 1 Step -1
        wb.Model.ModelRelationships(i).Delete
    Next i
    ' tabla dinamica y hojas de salida
    wb.Worksheets(HOJA_TD).Delete
    ' conexiones de las consultas
    For i = wb.Connections.Count To 1 Step -1
        For Each n In nombres
            If wb.Connections(i).Name = "Consulta - " & n Then
                wb.Connections(i).Delete
                Exit For
            End If
        Next n
    Next i
    ' consultas
    For Each n In nombres
        wb.Queries(CStr(n)).Delete
    Next n
    On Error GoTo 0

    Registrar "Limpieza previa completada (consultas, conexiones, relaciones y medidas homonimas)"
End Sub

' -------------------------------------------------------------------------------------
' 3. CONSULTAS POWER QUERY (LENGUAJE M)
' -------------------------------------------------------------------------------------
Private Sub CrearConsultas(wb As Workbook)
    wb.Queries.Add Name:=NOM_HECHOS, Formula:=M_Hechos(), _
        Description:="Tabla de hechos. Origen .csv transaccional. Limpieza + metricas de fila calculadas en el ETL."
    wb.Queries.Add Name:=NOM_CLIENTES, Formula:=M_Clientes(), _
        Description:="Dimension Clientes. Origen Maestras.xlsx. Table.Distinct para garantizar el lado 1 de la relacion."
    wb.Queries.Add Name:=NOM_PRODUCTOS, Formula:=M_Productos(), _
        Description:="Dimension Productos. Origen Maestras.xlsx. Table.Distinct para garantizar el lado 1 de la relacion."
    wb.Queries.Add Name:=NOM_CALENDARIO, Formula:=M_Calendario(), _
        Description:="Dimension Calendario generada en M. Una fila por dia, sin huecos."

    Registrar "Consultas M creadas: " & NOM_HECHOS & ", " & NOM_CLIENTES & ", " & NOM_PRODUCTOS & ", " & NOM_CALENDARIO
End Sub

Private Function M_Hechos() As String
    Dim s As String
    Dim ruta As String
    ruta = gRutaFuentes & "\Ventas_Historico.csv"

    s = s & "let" & vbCrLf
    s = s & "    // ORIGEN: archivo transaccional .csv ubicado en la subcarpeta \Fuentes del libro." & vbCrLf
    s = s & "    // La ruta la reescribe la macro MaterializarModelo al ejecutarse, de modo que el" & vbCrLf
    s = s & "    // libro se puede mover de carpeta y solo hay que volver a correrla." & vbCrLf
    s = s & "    Origen = Csv.Document(File.Contents(^" & ruta & "^), [Delimiter=^,^, Columns=9, Encoding=65001, QuoteStyle=QuoteStyle.Csv])," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    EncabezadosPromovidos = Table.PromoteHeaders(Origen, [PromoteAllScalars=true])," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    // Tipado explicito con Culture=^en-US^: el .csv usa punto decimal y fecha ISO," & vbCrLf
    s = s & "    // asi el tipado no depende de la configuracion regional del equipo que abra el libro." & vbCrLf
    s = s & "    TiposDefinidos = Table.TransformColumnTypes(EncabezadosPromovidos,{" & vbCrLf
    s = s & "        {^ID_Venta^, type text}, {^Fecha^, type date}, {^ID_Cliente^, type text}," & vbCrLf
    s = s & "        {^ID_Producto^, type text}, {^Canal^, type text}, {^Cantidad^, Int64.Type}," & vbCrLf
    s = s & "        {^Precio_Unitario^, type number}, {^Descuento^, type number}, {^Costo_Unitario^, type number}" & vbCrLf
    s = s & "    }, ^en-US^)," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    // PASO RENOMBRADO MANUALMENTE (era ^Filas filtradas^): consolida en un unico" & vbCrLf
    s = s & "    // Table.SelectRows lo que la interfaz genera como tres pasos encadenados" & vbCrLf
    s = s & "    // (quitar nulos, quitar vacios, filtrar cantidad). Un solo recorrido de la tabla." & vbCrLf
    s = s & "    VentasFiltradas = Table.SelectRows(TiposDefinidos, each" & vbCrLf
    s = s & "        [Fecha] <> null and" & vbCrLf
    s = s & "        [ID_Cliente] <> null and [ID_Cliente] <> ^^ and" & vbCrLf
    s = s & "        [ID_Producto] <> null and [ID_Producto] <> ^^ and" & vbCrLf
    s = s & "        [Cantidad] <> null and [Cantidad] > 0)," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    // METRICAS DE FILA RESUELTAS EN EL ETL, NO EN DAX." & vbCrLf
    s = s & "    // Materializarlas aca evita iteradores (SUMX/FILTER) sobre la tabla de hechos:" & vbCrLf
    s = s & "    // en el modelo alcanza con SUM() sobre columnas ya calculadas." & vbCrLf
    s = s & "    VentaBruta = Table.AddColumn(VentasFiltradas, ^Venta_Bruta^, each [Cantidad] * [Precio_Unitario], type number)," & vbCrLf
    s = s & "    VentaNeta = Table.AddColumn(VentaBruta, ^Venta_Neta^, each [Venta_Bruta] * (1 - [Descuento]), type number)," & vbCrLf
    s = s & "    CostoTotal = Table.AddColumn(VentaNeta, ^Costo_Total^, each [Cantidad] * [Costo_Unitario], type number)," & vbCrLf
    s = s & "    MargenBruto = Table.AddColumn(CostoTotal, ^Margen_Bruto^, each [Venta_Neta] - [Costo_Total], type number)," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    // REDUCCION DE HUELLA: se descartan las columnas que ya quedaron absorbidas en las" & vbCrLf
    s = s & "    // metricas (Precio_Unitario, Descuento, Costo_Unitario, Venta_Bruta). Menos columnas" & vbCrLf
    s = s & "    // de alta cardinalidad = menor tamanio del modelo en memoria." & vbCrLf
    s = s & "    ColumnasFinales = Table.SelectColumns(MargenBruto,{" & vbCrLf
    s = s & "        ^ID_Venta^, ^Fecha^, ^ID_Cliente^, ^ID_Producto^, ^Canal^," & vbCrLf
    s = s & "        ^Cantidad^, ^Venta_Neta^, ^Costo_Total^, ^Margen_Bruto^})" & vbCrLf
    s = s & "in" & vbCrLf
    s = s & "    ColumnasFinales" & vbCrLf

    M_Hechos = Q(s)
End Function

Private Function M_Clientes() As String
    Dim s As String
    Dim ruta As String
    ruta = gRutaFuentes & "\Maestras.xlsx"

    s = s & "let" & vbCrLf
    s = s & "    // ORIGEN: libro local de tablas maestras (.xlsx), segunda fuente heterogenea del flujo." & vbCrLf
    s = s & "    Origen = Excel.Workbook(File.Contents(^" & ruta & "^), null, true)," & vbCrLf
    s = s & "    TablaClientes = Origen{[Item=^Tbl_Clientes^, Kind=^Table^]}[Data]," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    TiposDefinidos = Table.TransformColumnTypes(TablaClientes,{" & vbCrLf
    s = s & "        {^ID_Cliente^, type text}, {^Cliente^, type text}, {^Segmento^, type text}," & vbCrLf
    s = s & "        {^Provincia^, type text}, {^Pais^, type text}}, ^en-US^)," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ClavesValidas = Table.SelectRows(TiposDefinidos, each [ID_Cliente] <> null and [ID_Cliente] <> ^^)," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    // Table.Distinct sobre la PK: sin unicidad garantizada Power Pivot rechaza la" & vbCrLf
    s = s & "    // relacion 1:N y la degrada a N:N. Este paso es el que sostiene el lado ^1^." & vbCrLf
    s = s & "    ClientesUnicos = Table.Distinct(ClavesValidas, {^ID_Cliente^})" & vbCrLf
    s = s & "in" & vbCrLf
    s = s & "    ClientesUnicos" & vbCrLf

    M_Clientes = Q(s)
End Function

Private Function M_Productos() As String
    Dim s As String
    Dim ruta As String
    ruta = gRutaFuentes & "\Maestras.xlsx"

    s = s & "let" & vbCrLf
    s = s & "    Origen = Excel.Workbook(File.Contents(^" & ruta & "^), null, true)," & vbCrLf
    s = s & "    TablaProductos = Origen{[Item=^Tbl_Productos^, Kind=^Table^]}[Data]," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    TiposDefinidos = Table.TransformColumnTypes(TablaProductos,{" & vbCrLf
    s = s & "        {^ID_Producto^, type text}, {^Producto^, type text}, {^Categoria^, type text}," & vbCrLf
    s = s & "        {^Marca^, type text}, {^Precio_Lista^, type number}}, ^en-US^)," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ClavesValidas = Table.SelectRows(TiposDefinidos, each [ID_Producto] <> null and [ID_Producto] <> ^^)," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    // Misma logica que en Dim_Clientes: unicidad de PK garantizada en el ETL." & vbCrLf
    s = s & "    ProductosUnicos = Table.Distinct(ClavesValidas, {^ID_Producto^})" & vbCrLf
    s = s & "in" & vbCrLf
    s = s & "    ProductosUnicos" & vbCrLf

    M_Productos = Q(s)
End Function

Private Function M_Calendario() As String
    Dim s As String

    s = s & "let" & vbCrLf
    s = s & "    // DIMENSION CALENDARIO generada integramente en M (no hay archivo de origen)." & vbCrLf
    s = s & "    // El rango cubre el historico completo del .csv transaccional." & vbCrLf
    s = s & "    FechaInicio = #date(" & ANIO_INI & ", 1, 1)," & vbCrLf
    s = s & "    FechaFin = #date(" & ANIO_FIN & ", 12, 31)," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    // Una fila por dia y sin huecos: requisito para marcarla como Tabla de fechas" & vbCrLf
    s = s & "    // y para que funcione la inteligencia de tiempo (DATEADD / SAMEPERIODLASTYEAR)." & vbCrLf
    s = s & "    ListaDias = List.Dates(FechaInicio, Duration.Days(FechaFin - FechaInicio) + 1, #duration(1,0,0,0))," & vbCrLf
    s = s & "    TablaFechas = Table.FromList(ListaDias, Splitter.SplitByNothing(), {^Fecha^})," & vbCrLf
    s = s & "    FechaTipada = Table.TransformColumnTypes(TablaFechas,{{^Fecha^, type date}}, ^en-US^)," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ConAnio = Table.AddColumn(FechaTipada, ^Anio^, each Date.Year([Fecha]), Int64.Type)," & vbCrLf
    s = s & "    ConTrimestre = Table.AddColumn(ConAnio, ^Trimestre^, each ^T^ & Text.From(Date.QuarterOfYear([Fecha])), type text)," & vbCrLf
    s = s & "    ConMes = Table.AddColumn(ConTrimestre, ^Mes^, each Date.Month([Fecha]), Int64.Type)," & vbCrLf
    s = s & "    ConNombreMes = Table.AddColumn(ConMes, ^Nombre_Mes^, each Date.MonthName([Fecha], ^es-AR^), type text)," & vbCrLf
    s = s & "    ConAnioMes = Table.AddColumn(ConNombreMes, ^Anio_Mes^, each Text.From(Date.Year([Fecha])) & ^-^ & Text.PadStart(Text.From(Date.Month([Fecha])), 2, ^0^), type text)," & vbCrLf
    s = s & "    ConDiaSemana = Table.AddColumn(ConAnioMes, ^Dia_Semana^, each Date.DayOfWeekName([Fecha], ^es-AR^), type text)," & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    // Consistencia con el resto de las dimensiones: PK unica explicita." & vbCrLf
    s = s & "    CalendarioUnico = Table.Distinct(ConDiaSemana, {^Fecha^})" & vbCrLf
    s = s & "in" & vbCrLf
    s = s & "    CalendarioUnico" & vbCrLf

    M_Calendario = Q(s)
End Function

' -------------------------------------------------------------------------------------
' 4. CARGA AL MODELO DE DATOS ("solo conexion" + agregar al modelo)
' -------------------------------------------------------------------------------------
Private Sub CargarAlModelo(wb As Workbook)
    CargarConsulta wb, NOM_CALENDARIO
    CargarConsulta wb, NOM_CLIENTES
    CargarConsulta wb, NOM_PRODUCTOS
    CargarConsulta wb, NOM_HECHOS

    Application.CalculateUntilAsyncQueriesDone
    DoEvents

    Registrar "Tablas en el modelo de datos: " & wb.Model.ModelTables.Count
End Sub

Private Sub CargarConsulta(wb As Workbook, ByVal nombre As String)
    wb.Connections.Add2 _
        Name:="Consulta - " & nombre, _
        Description:="Conexion con la consulta '" & nombre & "' del libro. Solo conexion + modelo de datos.", _
        ConnectionString:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & nombre & ";Extended Properties=" & Chr(34) & Chr(34), _
        CommandText:="SELECT * FROM [" & nombre & "]", _
        lCmdtype:=xlCmdSql, _
        CreateModelConnection:=True, _
        ImportRelationships:=False
    Application.CalculateUntilAsyncQueriesDone
End Sub

' -------------------------------------------------------------------------------------
' 5. RELACIONES 1:N  (PK dimension  ->  FK hechos)
' -------------------------------------------------------------------------------------
Private Sub CrearRelaciones(wb As Workbook)
    Relacionar wb, NOM_CALENDARIO, "Fecha", "Fecha"
    Relacionar wb, NOM_CLIENTES, "ID_Cliente", "ID_Cliente"
    Relacionar wb, NOM_PRODUCTOS, "ID_Producto", "ID_Producto"
    Registrar "Relaciones activas en el modelo: " & wb.Model.ModelRelationships.Count & " (esperadas: 3)"
End Sub

Private Sub Relacionar(wb As Workbook, ByVal tablaDim As String, ByVal colPK As String, ByVal colFK As String)
    On Error GoTo Fallo
    wb.Model.ModelRelationships.Add _
        ForeignKeyColumn:=wb.Model.ModelTables(NOM_HECHOS).ModelTableColumns(colFK), _
        PrimaryKeyColumn:=wb.Model.ModelTables(tablaDim).ModelTableColumns(colPK)
    Registrar "  OK  " & tablaDim & "[" & colPK & "]  1 --> N  " & NOM_HECHOS & "[" & colFK & "]"
    Exit Sub
Fallo:
    Registrar "  FALLO relacion " & tablaDim & "[" & colPK & "] -> " & NOM_HECHOS & "[" & colFK & "]: " & Err.Description
End Sub

' -------------------------------------------------------------------------------------
' 6. MEDIDAS DAX (8 medidas, 4 con VAR/RETURN)
' -------------------------------------------------------------------------------------
Private Sub CrearMedidas(wb As Workbook)
    Dim fmtMoneda As Object, fmtEntero As Object, fmtPct As Object

    Set fmtMoneda = Nothing: Set fmtEntero = Nothing: Set fmtPct = Nothing
    On Error Resume Next
    Set fmtMoneda = wb.Model.ModelFormatCurrency("$", 0)
    Set fmtEntero = wb.Model.ModelFormatWholeNumber(True)
    Set fmtPct = wb.Model.ModelFormatPercentageNumber(False, 1)
    On Error GoTo 0

    ' --- agregaciones simples sobre columnas pre-calculadas en el ETL ---
    AgregarMedida wb, "Total Ventas", _
        "SUM(Hechos_Ventas[Venta_Neta])", fmtMoneda, _
        "Suma directa de la columna materializada en Power Query. Sin iteradores."

    AgregarMedida wb, "Total Costo", _
        "SUM(Hechos_Ventas[Costo_Total])", fmtMoneda, _
        "Agregacion simple sobre columna pre-calculada."

    AgregarMedida wb, "Margen Bruto", _
        "SUM(Hechos_Ventas[Margen_Bruto])", fmtMoneda, _
        "Agregacion simple. Reemplaza a SUMX(Hechos_Ventas, [Venta_Neta]-[Costo_Total])."

    AgregarMedida wb, "Unidades Vendidas", _
        "SUM(Hechos_Ventas[Cantidad])", fmtEntero, _
        "Volumen fisico vendido."

    ' --- medidas con VAR / RETURN ---
    AgregarMedida wb, "% Margen Bruto", _
        Join(Array( _
            "VAR VentasTotales = [Total Ventas]", _
            "VAR MargenTotal = [Margen Bruto]", _
            "RETURN", _
            "    DIVIDE(MargenTotal, VentasTotales, 0)"), vbLf), _
        fmtPct, _
        "VAR/RETURN: cada agregado se evalua UNA sola vez y se reutiliza. DIVIDE evita el error por division por cero."

    AgregarMedida wb, "Ticket Promedio", _
        Join(Array( _
            "VAR VentasTotales = [Total Ventas]", _
            "VAR CantidadTickets = DISTINCTCOUNT(Hechos_Ventas[ID_Venta])", _
            "RETURN", _
            "    DIVIDE(VentasTotales, CantidadTickets, 0)"), vbLf), _
        fmtMoneda, _
        "VAR/RETURN: el DISTINCTCOUNT se materializa una vez en lugar de recalcularse dentro del cociente."

    AgregarMedida wb, "Ventas Anio Anterior", _
        Join(Array( _
            "VAR VentasPrevias =", _
            "    CALCULATE([Total Ventas], DATEADD(Dim_Calendario[Fecha], -1, YEAR))", _
            "RETURN", _
            "    IF(ISBLANK(VentasPrevias), BLANK(), VentasPrevias)"), vbLf), _
        fmtMoneda, _
        "VAR/RETURN + inteligencia de tiempo apoyada en la relacion con Dim_Calendario."

    AgregarMedida wb, "Var % vs Anio Anterior", _
        Join(Array( _
            "VAR VentasActuales = [Total Ventas]", _
            "VAR VentasPrevias = [Ventas Anio Anterior]", _
            "RETURN", _
            "    IF(", _
            "        ISBLANK(VentasPrevias) || VentasPrevias = 0,", _
            "        BLANK(),", _
            "        DIVIDE(VentasActuales - VentasPrevias, VentasPrevias)", _
            "    )"), vbLf), _
        fmtPct, _
        "VAR/RETURN: guarda ambos periodos en variables y evita recalcular la medida base dentro del IF y del DIVIDE."
End Sub

Private Sub AgregarMedida(wb As Workbook, ByVal nombre As String, ByVal formula As String, _
                          ByVal fmt As Object, ByVal descripcion As String)
    On Error Resume Next
    Err.Clear

    If fmt Is Nothing Then Set fmt = wb.Model.ModelFormatGeneral()

    wb.Model.ModelMeasures.Add MeasureName:=nombre, _
                               AssociatedTable:=wb.Model.ModelTables(NOM_HECHOS), _
                               Formula:=formula, _
                               FormatInformation:=fmt, _
                               Description:=descripcion

    If Err.Number <> 0 Then
        Err.Clear
        wb.Model.ModelMeasures.Add MeasureName:=nombre, _
                                   AssociatedTable:=wb.Model.ModelTables(NOM_HECHOS), _
                                   Formula:=formula, _
                                   FormatInformation:=wb.Model.ModelFormatGeneral(), _
                                   Description:=descripcion
    End If

    If Err.Number <> 0 Then
        Registrar "  FALLO medida [" & nombre & "]: " & Err.Description
        Err.Clear
    Else
        Registrar "  OK  medida [" & nombre & "]"
    End If
    On Error GoTo 0
End Sub

' -------------------------------------------------------------------------------------
' 7. TABLA DINAMICA DE VERIFICACION
' -------------------------------------------------------------------------------------
Private Sub CrearTablaDinamica(wb As Workbook)
    Dim ws As Worksheet, pc As PivotCache, pt As PivotTable, cn As WorkbookConnection, i As Long

    On Error GoTo Fallo

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = HOJA_TD

    ws.Range("B2").Value = "Verificacion del modelo: [Total Ventas] por Dim_Calendario[Anio]"
    ws.Range("B2").Font.Bold = True
    ws.Range("B3").Value = "Si los importes cambian por anio, el filtro fluye desde la dimension hacia la tabla de hechos: las relaciones 1:N funcionan."

    For i = 1 To wb.Connections.Count
        If InStr(1, wb.Connections(i).Name, "DataModel", vbTextCompare) > 0 Then
            Set cn = wb.Connections(i)
            Exit For
        End If
    Next i
    If cn Is Nothing Then Err.Raise 5, , "No se encontro la conexion del modelo de datos."

    Set pc = wb.PivotCaches.Create(SourceType:=xlExternal, SourceData:=cn)
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("B5"), TableName:="TD_Verificacion")

    pt.CubeFields("[" & NOM_CALENDARIO & "].[Anio]").Orientation = xlRowField
    pt.CubeFields("[Measures].[Total Ventas]").Orientation = xlDataField
    pt.CubeFields("[Measures].[Margen Bruto]").Orientation = xlDataField
    pt.CubeFields("[Measures].[% Margen Bruto]").Orientation = xlDataField

    ws.Columns("B:E").AutoFit
    Registrar "Tabla dinamica de verificacion creada en la hoja '" & HOJA_TD & "'"
    Exit Sub

Fallo:
    Registrar "  FALLO tabla dinamica: " & Err.Description & " (se puede armar a mano: Insertar > Tabla dinamica > Usar el modelo de datos de este libro)"
End Sub

' -------------------------------------------------------------------------------------
' 8. SALIDAS DE DOCUMENTACION
' -------------------------------------------------------------------------------------
Private Sub VolcarCodigoM(wb As Workbook)
    Dim ws As Worksheet, f As Long
    Set ws = HojaLimpia(wb, HOJA_M)
    ws.Columns(1).NumberFormat = "@"   ' texto: evita que Excel interprete el codigo M como formula

    f = 1
    ws.Cells(f, 1).Value = "CODIGO M CARGADO EN POWER QUERY (generado por MaterializarModelo)"
    ws.Cells(f, 1).Font.Bold = True
    f = f + 2
    f = VolcarTexto(ws, f, "shared " & NOM_HECHOS & " =", M_Hechos())
    f = VolcarTexto(ws, f, "shared " & NOM_CLIENTES & " =", M_Clientes())
    f = VolcarTexto(ws, f, "shared " & NOM_PRODUCTOS & " =", M_Productos())
    f = VolcarTexto(ws, f, "shared " & NOM_CALENDARIO & " =", M_Calendario())

    ws.Columns(1).ColumnWidth = 120
End Sub

Private Function VolcarTexto(ws As Worksheet, ByVal fila As Long, ByVal titulo As String, ByVal texto As String) As Long
    Dim lineas As Variant, i As Long

    ws.Cells(fila, 1).Value = titulo
    ws.Cells(fila, 1).Font.Bold = True
    fila = fila + 1

    lineas = Split(Replace(texto, vbCrLf, vbLf), vbLf)
    For i = LBound(lineas) To UBound(lineas)
        ws.Cells(fila, 1).Value = lineas(i)
        fila = fila + 1
    Next i

    VolcarTexto = fila + 2
End Function

Private Sub EscribirLog(wb As Workbook)
    Dim ws As Worksheet, i As Long
    Set ws = HojaLimpia(wb, HOJA_LOG)

    ws.Range("A1").Value = "VERIFICACION DEL MODELO DE DATOS - " & Format$(Now, "yyyy-mm-dd hh:nn")
    ws.Range("A1").Font.Bold = True

    For i = 1 To gLog.Count
        ws.Cells(i + 2, 1).Value = gLog(i)
    Next i

    i = i + 3
    ws.Cells(i, 1).Value = "PENDIENTE MANUAL: Power Pivot > Administrar > Dim_Calendario > Diseniar > Marcar como tabla de fechas (columna Fecha)."
    ws.Cells(i, 1).Font.Bold = True
    ws.Cells(i + 1, 1).Value = "Donde se comprueba cada punto: Datos > Consultas y conexiones (las 4 consultas) | Power Pivot > Administrar > Vista de diagrama (estrella + 3 relaciones) | area de calculo de Hechos_Ventas (8 medidas) | hoja " & HOJA_TD & " (tabla dinamica)."

    ws.Columns(1).ColumnWidth = 130
    ws.Activate
End Sub

' -------------------------------------------------------------------------------------
' UTILIDADES
' -------------------------------------------------------------------------------------
Private Function HojaLimpia(wb As Workbook, ByVal nombre As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nombre)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = nombre
    Else
        ws.Cells.Clear
    End If

    Set HojaLimpia = ws
End Function

Private Sub Registrar(ByVal texto As String)
    If gLog Is Nothing Then Set gLog = New Collection
    gLog.Add texto
End Sub

' Reemplaza el marcador ^ por comillas dobles al construir el codigo M.
Private Function Q(ByVal s As String) As String
    Q = Replace(s, "^", Chr(34))
End Function

' Numero con punto decimal, independiente de la configuracion regional del equipo.
Private Function Num(ByVal v As Double, ByVal dec As Integer) As String
    Dim mascara As String
    If dec = 0 Then mascara = "0" Else mascara = "0." & String$(dec, "0")
    Num = Replace(Format$(v, mascara), ",", ".")
End Function
