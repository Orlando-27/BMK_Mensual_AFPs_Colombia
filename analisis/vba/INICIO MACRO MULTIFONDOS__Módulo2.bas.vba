Attribute VB_Name = "Módulo2"
Sub Importar_CSA()
'Jhonathan Higuera

Application.DisplayAlerts = False




Dim N_CSA As String
Dim R_CSA As String
Dim Macro As String
Dim N_MACRO As String
Dim R_MACRO As String

Sheets("PROCEDIMIENTO").Select

N_CSA = Range("N_CSA")
R_CSA = Range("R_CSA")
N_MACRO = Range("N_MACRO")
R_MACRO = Range("R_MACRO")

Dim workbookdestino As Workbook


CSA = R_CSA & "\" & N_CSA '& ".xlsx"
Set workbookdestino = Workbooks.Open(CSA)

Sheets("Balance Esperado").Select

Range("B3").Select

If Selection = "Portafolio" Then

   
Range("B3").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Workbooks(N_MACRO).Activate
Worksheets("Resumen CV").Select
Range("e47").Select
Selection.PasteSpecial xlValues


Workbooks(N_CSA).Activate
Workbooks(N_CSA).Close

Else

MsgBox ("Revisar CSA")


End If
Workbooks(N_MACRO).Activate
Worksheets("PROCEDIMIENTO").Select

End Sub

Sub CSA_balance()
'Katherin Taborda 09/2025

Application.DisplayAlerts = False

    Dim N_CSA As String, R_CSA As String
    Dim N_MACRO As String, R_MACRO As String
    Dim wbCSA As Workbook, wsBalance As Worksheet
    Dim wsDestino As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim datosExtraidos() As Variant

    ' Obtener rutas desde hoja PROCEDIMIENTO
    Sheets("PROCEDIMIENTO").Select
    N_CSA = Range("N_CSA").Value
    R_CSA = Range("R_CSA").Value
    N_MACRO = Range("N_MACRO").Value
    R_MACRO = Range("R_MACRO").Value

    ' Abrir archivo CSA
    Set wbCSA = Workbooks.Open(R_CSA & "\" & N_CSA)
    Set wsBalance = wbCSA.Sheets("Balances Esperados")
    
    ' Determinar última fila con datos en columna A
    lastRow = wsBalance.Cells(wsBalance.Rows.Count, "A").End(xlUp).Row

    ' Redimensionar matriz para almacenar columnas A,B,D,E
    ReDim datosExtraidos(1 To lastRow - 2, 1 To 4)
    
    For i = 3 To lastRow
        datosExtraidos(i - 2, 1) = wsBalance.Cells(i, 1).Value ' Columna A
        datosExtraidos(i - 2, 2) = wsBalance.Cells(i, 2).Value ' Columna B
        datosExtraidos(i - 2, 3) = wsBalance.Cells(i, 4).Value ' Columna D
        datosExtraidos(i - 2, 4) = wsBalance.Cells(i, 5).Value ' Columna E
    Next i

    ' Pegar en hoja destino
    Set wsDestino = Workbooks(N_MACRO).Sheets("Resumen CV")
    wsDestino.Range("AQ3").Resize(UBound(datosExtraidos, 1), 4).Value = datosExtraidos

    ' Cerrar CSA sin guardar
    wbCSA.Close SaveChanges:=False

    ' Volver a hoja PROCEDIMIENTO
    Workbooks(N_MACRO).Activate
    Sheets("PROCEDIMIENTO").Select

End Sub
