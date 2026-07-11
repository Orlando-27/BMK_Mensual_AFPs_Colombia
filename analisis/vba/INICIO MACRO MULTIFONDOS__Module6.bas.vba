Attribute VB_Name = "Module6"
Sub clasificacionPO()
Dim number
Dim RNG As Range
Sheets("PORTAFOLIOS").Select
Ultimaf = Application.WorksheetFunction.CountA(Range("C:C"))
Set RNG = Sheets("CLASIFICACION").Columns("A:M")
For i = 2 To Ultimaf
If IsError(Application.VLookup(Cells(i, 2), RNG, 2, 0)) Then
Cells(i, 2) = Trim(Cells(i, 6))
End If
Next

For i = 2 To Ultimaf

Cells(i, 13) = Application.VLookup(Cells(i, 2), RNG, 2, 0)
Cells(i, 14) = Application.VLookup(Cells(i, 2), RNG, 3, 0)
Cells(i, 15) = Application.VLookup(Cells(i, 2), RNG, 4, 0)
Cells(i, 16) = Application.VLookup(Cells(i, 2), RNG, 5, 0)
Cells(i, 17) = Application.VLookup(Cells(i, 2), RNG, 6, 0)
Cells(i, 18) = Application.VLookup(Cells(i, 2), RNG, 7, 0)
Cells(i, 23) = Application.VLookup(Cells(i, 2), RNG, 8, 0)
Cells(i, 24) = Application.VLookup(Cells(i, 2), RNG, 9, 0)
Cells(i, 25) = Application.VLookup(Cells(i, 2), RNG, 10, 0)
Cells(i, 26) = Application.VLookup(Cells(i, 2), RNG, 11, 0)
Cells(i, 27) = Application.VLookup(Cells(i, 2), RNG, 12, 0)
Cells(i, 29) = "=IF(RC[-6]=""r fija"",IF(OR(RC[-26]=""NA"",RC[-26]=""COB51CB00452""),1,COUNTIFS(C[-26],RC[-26],C[-20],RC[-20])/COUNTIFS(C[-26],RC[-26])),""NA"")"

If Cells(i, 19) = "No aplica" Or Cells(i, 19) = "Varias" Then
Cells(i, 20) = Cells(i, 19)
Else
number = Len(Cells(i, 19))
Select Case number
Case 6
Cells(i, 20) = Mid(Trim(Cells(i, 19)), 5, 1)
Case 7
Cells(i, 20) = Mid(Trim(Cells(i, 19)), 6, 1)
Case 8
Cells(i, 20) = Mid(Trim(Cells(i, 19)), 5, 3)
Case 9
If Left(Trim(Cells(i, 19)), 4) = "IBR0" Then
Cells(i, 20) = Mid(Trim(Cells(i, 19)), 6, 3)
Else
Cells(i, 20) = Mid(Trim(Cells(i, 19)), 5, 4)
End If
Case 10
Cells(i, 20) = Mid(Trim(Cells(i, 19)), 5, 5)
Case 0
Cells(i, 19) = "0.00000001 - Ef"
Cells(i, 20) = 0.00000001
Case Else
Cells(i, 20) = Left(Trim(Cells(i, 19)), 9)
End Select
End If
Next

'Sheets("PORTAFOLIOS").Select
'Range("A:AA").Select
'    Selection.AutoFilter
'    ActiveSheet.Range("A1:AA1").AutoFilter Field:=1, Criteria1:="=AD", Operator:=xlFilterValues
'    Range("A2").Select
'
'    Dim COUNTULT As String
'    COUNTULT = Worksheets("PORTAFOLIOS").Cells(Rows.Count, 1).End(xlUp).Row
'
'    Range(("A2:AA" + Format(COUNTULT, "0"))).Select
'    Selection.SpecialCells(xlCellTypeVisible).Select
'    Selection.EntireRow.Delete
'    Selection.AutoFilter
    
End Sub
