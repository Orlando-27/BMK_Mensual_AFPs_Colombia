Attribute VB_Name = "Module3"
Sub PORTAFOLIOS()
Sheets("T597").Select
k = 2
Ultimaf = Application.WorksheetFunction.CountA(Range("ak:ak"))

For i = 4 To Ultimaf + 2

Worksheets("PORTAFOLIOS").Cells(k, 1).Value = Trim(Cells(i, 37))
Worksheets("PORTAFOLIOS").Cells(k, 3).Value = Trim(Cells(i, 9))
Worksheets("PORTAFOLIOS").Cells(k, 4).Value = Cells(i, 6)
Worksheets("PORTAFOLIOS").Cells(k, 5).Value = Trim(Cells(i, 5))
Worksheets("PORTAFOLIOS").Cells(k, 6).Value = Cells(i, 3)
Worksheets("PORTAFOLIOS").Cells(k, 7).Value = Cells(i, 17)
Worksheets("PORTAFOLIOS").Cells(k, 8).Value = Cells(i, 12)
Worksheets("PORTAFOLIOS").Cells(k, 9).Value = Cells(i, 13)
Worksheets("PORTAFOLIOS").Cells(k, 10).Value = Cells(i, 20)
Worksheets("PORTAFOLIOS").Cells(k, 11).Value = Cells(i, 14)
Worksheets("PORTAFOLIOS").Cells(k, 12).Value = Cells(i, 25)
Worksheets("PORTAFOLIOS").Cells(k, 21).Value = Cells(i, 33)
Worksheets("PORTAFOLIOS").Cells(k, 22).Value = Cells(i, 16)
Worksheets("PORTAFOLIOS").Cells(k, 19).Value = Trim(Cells(i, 15))

If Trim(Cells(i, 8)) = "" Or Trim(Cells(i, 8)) = "NA" Then
If Trim(Cells(i, 9)) = "NA" Or Trim(Cells(i, 9)) = "" Then
Worksheets("PORTAFOLIOS").Cells(k, 2).Value = Trim(Cells(i, 3))
Else
Worksheets("PORTAFOLIOS").Cells(k, 2).Value = Trim(Cells(i, 9))
End If
Else
Worksheets("PORTAFOLIOS").Cells(k, 2).Value = Trim(Cells(i, 8))
End If

k = k + 1

Next
Module6.clasificacionPO
End Sub
