Attribute VB_Name = "Module2"
Sub Ventas()
Sheets("T597").Select
Ultimaf = Application.WorksheetFunction.CountA(Range("ak:ak"))
k = 2
For i = 4 To Ultimaf
If Trim(Cells(i, 36)) = "X" Then
Worksheets("VENTAS").Cells(k, 2).Value = "Venta" 'OPERACION
Worksheets("VENTAS").Cells(k, 6).Value = Cells(i, 6).Value 'TITULO
Worksheets("VENTAS").Cells(k, 7).Value = Cells(i, 14).Value 'Nominal
Worksheets("VENTAS").Cells(k, 8).Value = Cells(i, 12).Value 'Emision
Worksheets("VENTAS").Cells(k, 9).Value = Cells(i, 13).Value 'Vencimiento
Worksheets("VENTAS").Cells(k, 10).Value = Cells(i, 15).Value 'Facial
Worksheets("VENTAS").Cells(k, 13).Value = Cells(i, 17).Value 'Compromiso
Worksheets("VENTAS").Cells(k, 17).Value = Cells(i, 25).Value 'Valor de Mercado
Worksheets("VENTAS").Cells(k, 22).Value = Cells(i, 37).Value 'Portafolio
Worksheets("VENTAS").Cells(k, 4).Value = Cells(i, 3).Value 'Serie
k = k + 1
End If
Next
ActiveCell(4, 36).Select
Counter = Ultimaf
For i = 1 To Counter
    ' Checks to see if the active cell is blank.
    If ActiveCell = "X" Then
        Selection.EntireRow.Delete
        ' Decrements count each time a row is deleted. This ensures
        ' that the macro will not run past the last row.
        Counter = Counter - 1
    Else
        ' Selects the next cell.
        ActiveCell.Offset(1, 0).Select
    End If
Next i
End Sub
Sub Compras()
Sheets("T620").Select
Ultimafta = Application.WorksheetFunction.CountA(Range("a:a"))
Sheets("T597").Select
Ultimaftaa = Application.WorksheetFunction.CountA(Range("a:a")) + 1
For i = 4 To Ultimafta + 1
Set B = Worksheets("T620").Cells(i, 2)
If Trim(B) = "Compra" Then
k = 4
Set a = Worksheets("T620").Cells(i, 4)
Do Until Trim(a) = Trim(Cells(k, 3)) Or k = Ultimaftaa
k = k + 1
Loop
Dim TN As Range
Set TN = Sheets("TITULOS NUEVOS").Columns("A:D")
If k = Ultimaftaa Then
Ultimaftaa = Ultimaftaa + 1
Cells(Ultimaftaa, 3) = Application.VLookup(Trim(a), TN, 2, 0)
Cells(Ultimaftaa, 8) = Worksheets("T620").Cells(i, 6).Value
Cells(Ultimaftaa, 9) = Worksheets("T620").Cells(i, 5).Value
Cells(Ultimaftaa, 6) = Worksheets("T620").Cells(i, 8).Value
Cells(Ultimaftaa, 14) = Worksheets("T620").Cells(i, 9).Value
Cells(Ultimaftaa, 18) = Worksheets("T620").Cells(i, 9).Value
Cells(Ultimaftaa, 12) = Worksheets("T620").Cells(i, 10).Value
Cells(Ultimaftaa, 13) = Worksheets("T620").Cells(i, 11).Value
Cells(Ultimaftaa, 15) = Worksheets("T620").Cells(i, 12).Value
Cells(Ultimaftaa, 17) = Worksheets("T620").Cells(i, 15).Value
Cells(Ultimaftaa, 37) = Worksheets("T620").Cells(i, 24).Value
Cells(Ultimaftaa, 25) = Worksheets("T620").Cells(i, 19).Value
Cells(Ultimaftaa, 33) = Worksheets("T620").Cells(i, 13).Value
Cells(Ultimaftaa, 20) = Application.VLookup(Trim(a), TN, 4, 0)
Cells(Ultimaftaa, 4) = "No aplica"
Cells(Ultimaftaa, 5) = Trim(a)
Cells(Ultimaftaa, 16) = Application.VLookup(Trim(a), TN, 3, 0)
Else
If Trim(Worksheets("T620").Cells(i, 4)) = Trim(Cells(k, 3)) Then
Ultimaftaa = Ultimaftaa + 1
Cells(Ultimaftaa, 3) = Cells(k, 3)
Cells(Ultimaftaa, 8) = Cells(k, 8)
Cells(Ultimaftaa, 9) = Cells(k, 9)
Cells(Ultimaftaa, 6) = Worksheets("T620").Cells(i, 8).Value
Cells(Ultimaftaa, 14) = Worksheets("T620").Cells(i, 9).Value
Cells(Ultimaftaa, 12) = Worksheets("T620").Cells(i, 10).Value
Cells(Ultimaftaa, 13) = Worksheets("T620").Cells(i, 11).Value
Cells(Ultimaftaa, 15) = Worksheets("T620").Cells(i, 12).Value
Cells(Ultimaftaa, 17) = Worksheets("T620").Cells(i, 15).Value
Cells(Ultimaftaa, 37) = Worksheets("T620").Cells(i, 24).Value
Cells(Ultimaftaa, 25) = Worksheets("T620").Cells(i, 19).Value
Cells(Ultimaftaa, 33) = Worksheets("T620").Cells(i, 13).Value
Cells(Ultimaftaa, 20) = Cells(k, 20)
Cells(Ultimaftaa, 4) = Cells(k, 4)
Cells(Ultimaftaa, 5) = Cells(k, 5)
Cells(Ultimaftaa, 16) = Cells(k, 16)
End If
End If
End If
Next
End Sub
