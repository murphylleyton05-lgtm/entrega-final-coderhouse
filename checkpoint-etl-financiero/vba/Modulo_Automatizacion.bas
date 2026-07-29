Attribute VB_Name = "Modulo_Automatizacion"
Option Explicit

'==========================================================================
'  Modulo_Automatizacion
'  Checkpoint: Flujo ETL, Analisis Financiero y Reporte Ejecutivo
'  Autor: Lleyton Murphy
'
'  Contiene las macros enlazadas a los botones (Controles de formulario)
'  de la hoja Reporte_Ejecutivo. Ningun procedimiento depende de la hoja
'  activa: todas las hojas se referencian por su CodeName, de modo que el
'  codigo sigue funcionando aunque se renombren las pestanas.
'==========================================================================

Public Const NOMBRE_CONSULTA As String = "Transacciones"
Public Const NOMBRE_TD As String = "TD_Margen"
Public Const CELDA_ESTADO As String = "B4"   ' celda superior izquierda del rango combinado B4:K4


'--------------------------------------------------------------------------
'  BOTON 1 - ACTUALIZAR DATOS (ETL)
'  Equivale a Datos > Actualizar todo, pero ademas garantiza que la
'  consulta, la tabla dinamica y los segmentadores existan.
'--------------------------------------------------------------------------
Public Sub Actualizar_Datos()
    Dim reloj As Double

    On Error GoTo Fallo
    reloj = Timer
    Preparar_Entorno False
    Estado "Actualizando el modelo de datos..."

    ETL_Asegurar_Consulta            ' crea o actualiza la consulta de Power Query
    ETL_Cargar_Al_Modelo             ' garantiza la conexion al Modelo de datos
    Actualizar_Todo_Sincronico       ' Datos > Actualizar todo (esperando el resultado)
    ETL_Asegurar_TablaDinamica       ' tabla dinamica + segmentadores
    Refrescar_Dinamicas

    Preparar_Entorno True
    Estado "Datos actualizados en " & Format$(Timer - reloj, "0.0") & " s"
    MsgBox "Modelo de datos actualizado correctamente." & vbCrLf & vbCrLf & _
           "Consulta Power Query: " & NOMBRE_CONSULTA & vbCrLf & _
           "Filas cargadas: " & Format$(Filas_Cargadas(), "#,##0"), _
           vbInformation, "Actualizar datos"
    Exit Sub

Fallo:
    Preparar_Entorno True
    Estado "Error al actualizar: " & Err.Description
    MsgBox "No se pudo completar la actualizacion." & vbCrLf & vbCrLf & _
           "Detalle: " & Err.Description, vbExclamation, "Actualizar datos"
End Sub


'--------------------------------------------------------------------------
'  BOTON 4 - LIMPIAR FILTROS
'  Devuelve el tablero a su estado base: segmentadores, filtros de la
'  tabla dinamica y autofiltros de todas las hojas.
'--------------------------------------------------------------------------
Public Sub Limpiar_Filtros()
    Dim cache As SlicerCache
    Dim hoja As Worksheet
    Dim td As PivotTable
    Dim tabla As ListObject
    Dim limpiados As Long

    Preparar_Entorno False

    ' 1) Segmentadores (slicers)
    For Each cache In ThisWorkbook.SlicerCaches
        On Error Resume Next
        cache.ClearManualFilter          ' segmentadores OLAP / modelo de datos
        cache.ClearAllFilters            ' segmentadores clasicos
        On Error GoTo 0
        limpiados = limpiados + 1
    Next cache

    ' 2) Filtros de tablas dinamicas y autofiltros
    For Each hoja In ThisWorkbook.Worksheets
        For Each td In hoja.PivotTables
            On Error Resume Next
            td.ClearAllFilters
            On Error GoTo 0
        Next td

        On Error Resume Next
        If hoja.FilterMode Then hoja.ShowAllData
        For Each tabla In hoja.ListObjects
            If Not tabla.AutoFilter Is Nothing Then
                If tabla.AutoFilter.FilterMode Then tabla.AutoFilter.ShowAllData
            End If
        Next tabla
        On Error GoTo 0
    Next hoja

    Preparar_Entorno True
    Estado "Filtros y segmentadores restablecidos"
    MsgBox "Se restablecieron " & limpiados & " segmentador(es) y todos los " & _
           "filtros del tablero.", vbInformation, "Limpiar filtros"
End Sub


'--------------------------------------------------------------------------
'  BOTON 5 - EXPORTAR A PDF
'  Exporta el Reporte Ejecutivo ajustado a UNA sola pagina, centrado, y lo
'  guarda en la carpeta local del proyecto.
'--------------------------------------------------------------------------
Public Sub Exportar_PDF()
    Dim carpeta As String
    Dim archivo As String

    On Error GoTo Fallo
    carpeta = Carpeta_Del_Libro()
    If Len(carpeta) = 0 Then
        MsgBox "Guarda el libro en una carpeta antes de exportar el PDF.", _
               vbExclamation, "Exportar a PDF"
        Exit Sub
    End If

    Preparar_Entorno False
    Configurar_Impresion Hoja_Reporte
    archivo = carpeta & "Reporte_Ejecutivo_" & Format$(Now, "yyyy-mm-dd_hhnn") & ".pdf"

    Hoja_Reporte.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=archivo, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    Preparar_Entorno True
    Estado "PDF exportado: " & Mid$(archivo, InStrRev(archivo, Application.PathSeparator) + 1)
    MsgBox "Reporte exportado en una sola pagina:" & vbCrLf & vbCrLf & archivo, _
           vbInformation, "Exportar a PDF"
    Exit Sub

Fallo:
    Preparar_Entorno True
    MsgBox "No se pudo exportar el PDF." & vbCrLf & vbCrLf & _
           "Detalle: " & Err.Description, vbExclamation, "Exportar a PDF"
End Sub


'--------------------------------------------------------------------------
'  BOTON 6 - MOSTRAR / OCULTAR HOJAS DE SOPORTE
'--------------------------------------------------------------------------
Public Sub Alternar_Hojas_Soporte()
    If Hoja_Crudos.Visible = xlSheetVisible Then
        Ocultar_Hojas_Soporte
        Estado "Hojas de soporte ocultas"
    Else
        Mostrar_Hojas_Soporte
        Estado "Hojas de soporte visibles (modo mantenimiento)"
    End If
End Sub


'  Deja a la vista unicamente el Reporte Ejecutivo y la Evaluacion Financiera.
Public Sub Ocultar_Hojas_Soporte()
    On Error Resume Next
    Hoja_Reporte.Visible = xlSheetVisible
    Hoja_Finanzas.Visible = xlSheetVisible
    Hoja_Datos.Visible = xlSheetHidden
    Hoja_Crudos.Visible = xlSheetHidden
    Hoja_Config.Visible = xlSheetHidden
    Hoja_Reporte.Activate
    On Error GoTo 0
End Sub


Public Sub Mostrar_Hojas_Soporte()
    On Error Resume Next
    Hoja_Datos.Visible = xlSheetVisible
    Hoja_Crudos.Visible = xlSheetVisible
    Hoja_Config.Visible = xlSheetVisible
    On Error GoTo 0
End Sub


'--------------------------------------------------------------------------
'  Puesta en marcha completa. Se ejecuta desde Workbook_Open y desde el
'  boton "Inicializar". Es idempotente: si el modelo ya esta armado, solo
'  refresca lo necesario.
'--------------------------------------------------------------------------
Public Sub Inicializar_Proyecto()
    Dim habia_modelo As Boolean

    On Error GoTo Fallo
    habia_modelo = Modelo_Construido()
    Preparar_Entorno False

    ETL_Escribir_Parametros          ' deja la ruta del libro a mano para Power Query

    If Not habia_modelo Then
        Estado "Construyendo el modelo por primera vez..."
        ETL_Asegurar_Consulta
        ETL_Cargar_A_Hoja
        ETL_Cargar_Al_Modelo
        ETL_Asegurar_TablaDinamica
        Aplicar_Formato_Silencioso
    Else
        Refrescar_Dinamicas
    End If

    Ocultar_Hojas_Soporte
    Configurar_Impresion Hoja_Reporte
    Preparar_Entorno True
    Estado IIf(habia_modelo, "Listo", "Modelo construido correctamente")
    Exit Sub

Fallo:
    Preparar_Entorno True
    Estado "Inicializacion incompleta: " & Err.Description & _
           " - usa el boton ACTUALIZAR DATOS"
End Sub


'==========================================================================
'  Utilidades internas
'==========================================================================

'  Datos > Actualizar todo, pero en modo sincronico: sin esto la macro
'  seguiria adelante mientras la consulta aun se esta refrescando.
Private Sub Actualizar_Todo_Sincronico()
    Dim conexion As WorkbookConnection

    For Each conexion In ThisWorkbook.Connections
        On Error Resume Next
        If conexion.Type = xlConnectionTypeOLEDB Then
            conexion.OLEDBConnection.BackgroundQuery = False
        ElseIf conexion.Type = xlConnectionTypeODBC Then
            conexion.ODBCConnection.BackgroundQuery = False
        End If
        On Error GoTo 0
    Next conexion

    ThisWorkbook.RefreshAll
    On Error Resume Next
    Application.CalculateUntilAsyncQueriesDone
    On Error GoTo 0
End Sub


Private Sub Refrescar_Dinamicas()
    Dim hoja As Worksheet
    Dim td As PivotTable

    On Error Resume Next
    For Each hoja In ThisWorkbook.Worksheets
        For Each td In hoja.PivotTables
            td.RefreshTable
        Next td
    Next hoja
    On Error GoTo 0
End Sub


'  Configuracion de impresion profesional: una sola pagina y centrada.
Public Sub Configurar_Impresion(ByVal hoja As Worksheet)
    On Error Resume Next
    Application.PrintCommunication = False
    With hoja.PageSetup
        .PrintArea = "$A$1:$K$46"
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .Zoom = False                    ' obligatorio antes de usar FitToPages
        .FitToPagesWide = 1
        .FitToPagesTall = 1              ' <- todo el reporte en UNA pagina
        .CenterHorizontally = True
        .CenterVertically = True
        .LeftMargin = Application.InchesToPoints(0.3)
        .RightMargin = Application.InchesToPoints(0.3)
        .TopMargin = Application.InchesToPoints(0.4)
        .BottomMargin = Application.InchesToPoints(0.4)
        .HeaderMargin = Application.InchesToPoints(0.2)
        .FooterMargin = Application.InchesToPoints(0.2)
        .CenterHeader = "&""Segoe UI,Negrita""&12Reporte Ejecutivo - Analisis de Expansion Comercial"
        .LeftFooter = "&""Segoe UI""&8Lleyton Murphy"
        .RightFooter = "&""Segoe UI""&8&D &T"
        .CenterFooter = "&""Segoe UI""&8Pagina &P de &N"
        .PrintGridlines = False
        .BlackAndWhite = False
    End With
    Application.PrintCommunication = True
    On Error GoTo 0
End Sub


Public Sub Estado(ByVal texto As String)
    On Error Resume Next
    Hoja_Reporte.Range(CELDA_ESTADO).Value = texto & _
        "   |   " & Format$(Now, "dd/mm/yyyy hh:nn")
    Application.StatusBar = texto
    DoEvents
    On Error GoTo 0
End Sub


Private Sub Preparar_Entorno(ByVal restaurar As Boolean)
    On Error Resume Next
    Application.ScreenUpdating = restaurar
    Application.EnableEvents = restaurar
    Application.DisplayAlerts = restaurar
    If restaurar Then
        Application.Cursor = xlDefault
        Application.StatusBar = False
    Else
        Application.Cursor = xlWait
    End If
    On Error GoTo 0
End Sub


Public Function Carpeta_Del_Libro() As String
    Dim ruta As String
    ruta = ThisWorkbook.Path
    If Len(ruta) = 0 Then
        Carpeta_Del_Libro = ""
    ElseIf Right$(ruta, 1) = Application.PathSeparator Then
        Carpeta_Del_Libro = ruta
    Else
        Carpeta_Del_Libro = ruta & Application.PathSeparator
    End If
End Function


Public Function Filas_Cargadas() As Long
    On Error Resume Next
    If Hoja_Datos.ListObjects.Count > 0 Then
        Filas_Cargadas = Hoja_Datos.ListObjects(1).ListRows.Count
    End If
    On Error GoTo 0
End Function


Public Function Modelo_Construido() As Boolean
    Dim td As PivotTable
    On Error Resume Next
    Set td = Hoja_Reporte.PivotTables(NOMBRE_TD)
    Modelo_Construido = (Not td Is Nothing) And (Filas_Cargadas() > 0)
    On Error GoTo 0
End Function
