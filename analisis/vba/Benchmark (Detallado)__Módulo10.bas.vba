Attribute VB_Name = "Módulo10"
Sub Vector_duraciones_SWAPs()
'
' Macro1 Macro
 
 
    Dim A_1 As String
    Dim R_1 As String

    Sheets("Procedimiento").Select
    A_dur_c = Range("A_dur_c")
    A_dur_b = Range("A_dur_b")
    R_1 = Range("R_1")
 
'''''' Copia las hojas de Colfondos y Bmk en un nuevo archivo
    Sheets(Array("Colfondos", "Benchmark")).Select
    Sheets("Benchmark").Activate
    Sheets(Array("Colfondos", "Benchmark")).Copy
'''''' Se acomoda la hoja de colfondos para dejar la informacion que necesitamos
 
    Sheets("Colfondos").Activate
    Range("Q7").Select
    ActiveSheet.Range("A:BW").AutoFilter Field:=17, Criteria1:=Array("CAJA", "FORWARD", "Notas estructuradas", "OPCIONES", "R FIJA", "R VARIABLE", "="), Operator:=xlFilterValues
    Rows("8:8").Select
    'Range("C8").Activate
    Range(Selection, Selection.End(xlDown)).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.ShowAllData
    Columns("A:C").Select
    Selection.Delete Shift:=xlToLeft
    Columns("B:Y").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:C").Select
    Selection.Delete Shift:=xlToLeft
    Columns("D:AU").Select
    Selection.Delete Shift:=xlToLeft
    Rows("1:6").Select
    'Range(Selection, Selection.End(xlUp)).Select
    Selection.Delete Shift:=xlUp
 
'''''' Se acomoda la hoja de BMK para dejar la informacion que necesitamos
 
    Sheets("Benchmark").Select
    ActiveSheet.Range("A:BO").AutoFilter Field:=18, Criteria1:=Array("CAJA", "FORWARD", "Notas estructuradas", "OPCIONES", "R FIJA", "R VARIABLE"), Operator:=xlFilterValues
    Rows("15:15").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.ShowAllData
    Columns("A:B").Select
    Selection.Delete Shift:=xlToLeft
    Columns("B:X").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:C").Select
    Selection.Delete Shift:=xlToLeft
    Columns("D:AO").Select
    Selection.Delete Shift:=xlToLeft
    Rows("1:13").Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.Delete Shift:=xlUp
    Range("A12").Select
    Sheets("Colfondos").Select
    ActiveSheet.Copy
 
    ActiveWorkbook.SaveAs Filename:=R_1 & "\" & A_dur_c & ".CSV", FileFormat:=xlCSV, CreateBackup:=False
    ActiveWorkbook.Close
    'ActiveWorkbook.SaveAs GuardarComo, xlCSV
    Sheets("Benchmark").Select
    ActiveSheet.Copy
 
    ActiveWorkbook.SaveAs Filename:=R_1 & "\" & A_dur_b & ".CSV", FileFormat:=xlCSV, CreateBackup:=False
    ActiveWorkbook.Close
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
 
End Sub
