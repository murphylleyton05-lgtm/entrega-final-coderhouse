Attribute VB_Name = "Modulo_Formato_IA"
Option Explicit

'==========================================================================
'  Modulo_Formato_IA
'
'  Modulo generado con asistencia de IA (ChatGPT) y depurado a mano.
'  El prompt utilizado y las correcciones aplicadas estan documentados en
'  docs/prompt_chatgpt.md y en la hoja Config del libro.
'
'  Aplica la identidad visual corporativa a todo el entregable: tipografia,
'  paleta, bordes ejecutivos y formatos de moneda / porcentaje.
'==========================================================================

' --- Paleta corporativa ---------------------------------------------------
Private Const AZUL_CORP As Long = 4860427      ' RGB(11, 42, 74)   - encabezados
Private Const CELESTE   As Long = 11766272     ' RGB(0, 138, 179)  - acentos
Private Const GRIS_SUAVE As Long = 16250100    ' RGB(244, 244, 247) - fondos
Private Const VERDE_OK  As Long = 6324224      ' RGB(0, 128, 96)   - resultado positivo
Private Const ROJO_ALERTA As Long = 2237106    ' RGB(178, 34, 34)  - resultado negativo
Private Const TIPOGRAFIA As String = "Segoe UI"


'--------------------------------------------------------------------------
'  BOTON 2 - APLICAR FORMATO EJECUTIVO
'--------------------------------------------------------------------------
Public Sub Aplicar_Formato_Ejecutivo()
    On Error GoTo Fallo
    Aplicar_Formato_Silencioso
    MsgBox "Formato ejecutivo aplicado al reporte, a la tabla dinamica y a " & _
           "la evaluacion financiera.", vbInformation, "Aplicar formato"
    Exit Sub

Fallo:
    Application.ScreenUpdating = True
    MsgBox "No se pudo aplicar el formato." & vbCrLf & vbCrLf & _
           "Detalle: " & Err.Description, vbExclamation, "Aplicar formato"
End Sub


'--------------------------------------------------------------------------
'  Misma rutina, sin cuadros de dialogo: la usa Workbook_Open.
'--------------------------------------------------------------------------
Public Sub Aplicar_Formato_Silencioso()
    On Error Resume Next
    Application.ScreenUpdating = False

    Formato_Tablero
    Formato_TablaDinamica
    Formato_Evaluacion_Financiera
    Formato_Datos_Limpios

    Application.ScreenUpdating = True
    Estado "Formato corporativo aplicado"
    On Error GoTo 0
End Sub


'--------------------------------------------------------------------------
Private Sub Formato_Tablero()
    Dim hoja As Worksheet
    Set hoja = Hoja_Reporte

    With hoja.Cells
        .Font.Name = TIPOGRAFIA
        .Font.Size = 10
    End With

    ' Banda de titulo
    With hoja.Range("A1:K2")
        .Interior.Color = AZUL_CORP
        .Font.Color = vbWhite
    End With
    With hoja.Range("A1")
        .Font.Size = 20
        .Font.Bold = True
    End With
    hoja.Range("A2").Font.Size = 11
    hoja.Range("A3:K3").Font.Color = RGB(90, 90, 90)

    ' Linea de estado
    With hoja.Range("B4:K4")
        .Interior.Color = GRIS_SUAVE
        .Font.Italic = True
        .Font.Size = 9
    End With

    ' Tarjetas de KPI
    Dim i As Long, col As Long
    Dim formatos As Variant
    formatos = Array("$#,##0", "$#,##0", "0.0%", "$#,##0", "#,##0")

    For i = 0 To 4
        col = 2 + i * 2                          ' B, D, F, H, J (tarjetas de 2 columnas)
        With hoja.Range(hoja.Cells(11, col), hoja.Cells(13, col + 1))
            .Interior.Color = GRIS_SUAVE
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        With hoja.Cells(11, col)
            .Font.Size = 9
            .Font.Bold = True
            .Font.Color = AZUL_CORP
        End With
        With hoja.Cells(12, col)
            .Font.Size = 16
            .Font.Bold = True
            .Font.Color = CELESTE
            .NumberFormat = formatos(i)
        End With
        Bordes_Ejecutivos hoja.Range(hoja.Cells(11, col), hoja.Cells(13, col + 1))
    Next i

    ' Titulo de la seccion analitica
    With hoja.Range("A16:K16")
        .Interior.Color = CELESTE
        .Font.Color = vbWhite
        .Font.Bold = True
    End With

    hoja.Rows("1:2").RowHeight = 22
    hoja.Rows("11").RowHeight = 16
    hoja.Rows("12").RowHeight = 26

    On Error Resume Next
    hoja.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = True
    On Error GoTo 0
End Sub


'--------------------------------------------------------------------------
Private Sub Formato_TablaDinamica()
    Dim td As PivotTable

    On Error Resume Next
    Set td = Hoja_Reporte.PivotTables(NOMBRE_TD)
    On Error GoTo 0
    If td Is Nothing Then Exit Sub

    On Error Resume Next
    td.TableStyle2 = "PivotStyleMedium9"
    td.ShowTableStyleRowStripes = True

    With td.TableRange1
        .Font.Name = TIPOGRAFIA
        .Font.Size = 10
    End With

    ' Encabezados de la dinamica con la paleta corporativa
    With td.TableRange1.Rows(1)
        .Interior.Color = AZUL_CORP
        .Font.Color = vbWhite
        .Font.Bold = True
        .WrapText = True
    End With

    td.DataBodyRange.HorizontalAlignment = xlRight
    Bordes_Ejecutivos td.TableRange1
    On Error GoTo 0
End Sub


'--------------------------------------------------------------------------
Private Sub Formato_Evaluacion_Financiera()
    Dim hoja As Worksheet
    Dim fila As Variant
    Set hoja = Hoja_Finanzas

    With hoja.Cells
        .Font.Name = TIPOGRAFIA
        .Font.Size = 10
    End With

    With hoja.Range("A1:H1")
        .Interior.Color = AZUL_CORP
        .Font.Color = vbWhite
        .Font.Bold = True
        .Font.Size = 16
    End With

    ' Encabezados de seccion
    For Each fila In Array(4, 19, 26, 34)
        With hoja.Range("A" & fila & ":H" & fila)
            .Interior.Color = CELESTE
            .Font.Color = vbWhite
            .Font.Bold = True
        End With
    Next fila

    ' Encabezado de la proyeccion
    With hoja.Range("B27:H27")
        .Interior.Color = GRIS_SUAVE
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    ' Celdas de entrada en azul (convencion de modelado financiero)
    hoja.Range("D5,D6,D9,D10,D20,D21,D22,D23,D24").Font.Color = RGB(0, 0, 200)

    ' Resultados destacados: PAGO, VAN y TIR
    For Each fila In Array(14, 35, 36)
        With hoja.Range("B" & fila & ":D" & fila)
            .Interior.Color = GRIS_SUAVE
            .Font.Bold = True
            .Font.Size = 11
        End With
        Bordes_Ejecutivos hoja.Range("B" & fila & ":D" & fila)
    Next fila

    ' Conclusion
    With hoja.Range("B40")
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = VERDE_OK
    End With

    Bordes_Ejecutivos hoja.Range("B27:H32")
End Sub


'--------------------------------------------------------------------------
Private Sub Formato_Datos_Limpios()
    Dim tabla As ListObject

    If Hoja_Datos.ListObjects.Count = 0 Then Exit Sub
    Set tabla = Hoja_Datos.ListObjects(1)

    On Error Resume Next
    tabla.TableStyle = "TableStyleMedium2"
    tabla.Range.Font.Name = TIPOGRAFIA
    tabla.Range.Font.Size = 9

    Formato_Columna tabla, "Precio_Unitario", "$#,##0.00"
    Formato_Columna tabla, "Costo_Unitario", "$#,##0.00"
    Formato_Columna tabla, "Monto", "$#,##0.00"
    Formato_Columna tabla, "Costo_Total", "$#,##0.00"
    Formato_Columna tabla, "Margen", "$#,##0.00"
    Formato_Columna tabla, "Margen_Pct", "0.0%"
    Formato_Columna tabla, "Fecha", "dd/mm/yyyy"
    Formato_Columna tabla, "Unidades", "#,##0"

    tabla.Range.EntireColumn.AutoFit
    On Error GoTo 0
End Sub


Private Sub Formato_Columna(ByVal tabla As ListObject, ByVal columna As String, _
                            ByVal formato As String)
    On Error Resume Next
    tabla.ListColumns(columna).DataBodyRange.NumberFormat = formato
    On Error GoTo 0
End Sub


'--------------------------------------------------------------------------
'  Bordes "ejecutivos": contorno medio y division interior fina.
'--------------------------------------------------------------------------
Private Sub Bordes_Ejecutivos(ByVal rango As Range)
    Dim borde As Variant

    On Error Resume Next
    For Each borde In Array(xlEdgeLeft, xlEdgeTop, xlEdgeRight, xlEdgeBottom)
        With rango.Borders(borde)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = AZUL_CORP
        End With
    Next borde

    For Each borde In Array(xlInsideVertical, xlInsideHorizontal)
        With rango.Borders(borde)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(190, 200, 210)
        End With
    Next borde
    On Error GoTo 0
End Sub
