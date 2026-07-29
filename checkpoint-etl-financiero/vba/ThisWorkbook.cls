Attribute VB_Name = "ThisWorkbook"
Option Explicit

'==========================================================================
'  ThisWorkbook
'  Al abrir el libro se pone en marcha el proyecto: se verifica que la
'  consulta de Power Query, la tabla dinamica y los segmentadores existan,
'  se ocultan las hojas de soporte y se deja el foco en el tablero.
'
'  La rutina es tolerante a fallos: si algo no puede construirse (por
'  ejemplo, si las conexiones de datos siguen deshabilitadas), el libro
'  se abre igual y el detalle queda escrito en la linea de estado.
'==========================================================================

Private Sub Workbook_Open()
    On Error Resume Next
    Inicializar_Proyecto
    Hoja_Reporte.Activate
    On Error GoTo 0
End Sub
