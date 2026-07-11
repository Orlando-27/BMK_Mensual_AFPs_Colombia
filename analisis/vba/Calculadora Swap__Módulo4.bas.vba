Attribute VB_Name = "Módulo4"
'JOSE MANUEL IBAÑEZ DELGADO
'INGENIERIA FINANCIERA
'CALCULADORA SWAP - 2016 - 1

Sub PEGADO()
Attribute PEGADO.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro6 Macro
'
Calculate
'
Sheets("SWAPIND").Select
    Range("B4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("INICIO").Select
    Range("B26").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("SWAPIND").Select
    Range("N4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("INICIO").Select
    Range("C26").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("SWAPIND").Select
    Range("V4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("INICIO").Select
    Range("D26").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    'pega duracion modif der en hoja inicio
    Sheets("SWAPIND").Select
    Range("AT4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("INICIO").Select
    Range("W26").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    'pega duracion modif obl en hoja inicio
    Sheets("SWAPIND").Select
    Range("AU4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("INICIO").Select
    Range("Y26").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        
         Sheets("SWAPIND").Select
    Range("W4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("INICIO").Select
    Range("E26").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub


Sub GUARDAR()

Dim RUTA_GUARDAR As String
Dim N_GUARDAR As String


RUTA_GUARDAR = Range("RUTA_GUARDAR")
N_GUARDAR = Range("N_GUARDAR")
N_GUARDAR_DM = Range("N_GUARDAR_DM")

'precios
Range("g26:j26").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Dim Libro_inicial, Libro_copia As Workbook
    
    Set Libro_inicial = ThisWorkbook
    
    Workbooks.Add
        Set Libro_copia = ActiveWorkbook

    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

ActiveWorkbook.SaveAs Filename:=RUTA_GUARDAR & "\" & N_GUARDAR & ".xlsb", FileFormat:=xlExcel12, CreateBackup:=False

ActiveWorkbook.Close

'duracion modificada
Range("v26:y26").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
    
    Set Libro_inicial = ThisWorkbook
    
    Workbooks.Add
        Set Libro_copia = ActiveWorkbook

    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

ActiveWorkbook.SaveAs Filename:=RUTA_GUARDAR & "\" & N_GUARDAR_DM & ".xlsb", FileFormat:=xlExcel12, CreateBackup:=False

ActiveWorkbook.Close

End Sub


