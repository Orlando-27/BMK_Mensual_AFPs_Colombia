Attribute VB_Name = "Módulo9"
Sub ReemplazarPalabraNotas()
    Dim ws As Worksheet
    Dim buscarTexto As String
    Dim reemplazarTexto As String
    Dim hojasObjetivo As Variant
    hojasObjetivo = Array("NE")
    
    For Each hojaNombre In hojasObjetivo

        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(hojaNombre)
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            ws.Cells.Replace What:="PROTECCIËN", Replacement:="PROTECCIÓN", _
                LookAt:=xlPart, MatchCase:=False

 Else
            MsgBox "La hoja '" & hojaNombre & "' no se encontró en el libro.", vbExclamation
        End If
    Next hojaNombre
End Sub

Sub Consolidar_Cupones_Swaps()
Calculate
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Cupónes Swaps")
ws.Range("A3:P3").Copy
ultimaFilaOrigen = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
ws.Cells(ultimaFilaOrigen, "A").PasteSpecial Paste:=xlPasteValues

End Sub
Sub Eliminar_Cupones_Swaps()

Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Cupónes Swaps")
ws.Range("A4:P10000").ClearContents

End Sub

