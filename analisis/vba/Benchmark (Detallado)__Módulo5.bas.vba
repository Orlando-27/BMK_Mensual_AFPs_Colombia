Attribute VB_Name = "Módulo5"
Sub calcular()
Selection.Calculate
End Sub




Public Sub SaveWorksheetsAsCsv()

Application.ScreenUpdating = False
Application.Calculation = xlManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayAlerts = False
Dim xWs As Worksheet

'Dim Fecha_Valoración As Date
'Fecha_Valoración = Range("Fecha_Valoración")


Copiar_paste

    Dim A_1 As String
    Dim R_1 As String
    Dim hoja_1 As String
    Dim Fecha_Valoración, fv As Date
   
    Sheets("Procedimiento").Select
    R_1 = Range("R_1")
    Fecha_Valoración = Range("Fecha_Valoración")
    'fv = Fecha_Valoración.NumberFormat = "dd-mm-yy;@"
    '.NumberFormat = "dd-mm-yy;@"
    
    Sheets(Array("Procedimiento", "Controles", "Voluntarias", "Formato de Apuestas PO-CS", "Formato de Apuestas", "Formato de Apuestas CP", "Opt Colfondos", "Opt Ind", "SWAP", "Fwd Colfondos", "Fwd Industria", "Curvas", _
        "Parametros", "VECTOR DE PRECIOS", "TASAS", "CUF")).Select
    Application.CutCopyMode = False
    ActiveWindow.SelectedSheets.Delete
      
    xDir = R_1
    
    For Each xWs In Application.ActiveWorkbook.Worksheets
    xWs.SaveAs xDir & "\" & xWs.Name, xlCSV
    'xWs.SaveAs xDir & "\" & xWs.Name & fv, xlCSV
    Next
    
    
   
    ActiveWorkbook.Close
    
    Application.DisplayAlerts = True

End Sub






'Public Sub SaveWorksheetsAsCsv()
'Dim xWs As Worksheet
'Dim xDir As String
'Dim folder As FileDialog
'Set folder = Application.FileDialog(msoFileDialogFolderPicker)
'
'If folder.Show <> -1 Then Exit Sub
'
'xDir = folder.SelectedItems(1)
'
'For Each xWs In Application.ActiveWorkbook.Worksheets
'
'xWs.SaveAs xDir & "\" & xWs.Name, xlCSV
'
'Next
'
'End Sub
'
'
'
