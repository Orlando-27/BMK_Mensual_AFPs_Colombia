Attribute VB_Name = "Module4"
Sub copia()
Attribute copia.VB_Description = "Macro recorded 3/3/2011 by JR47378"
Attribute copia.VB_ProcData.VB_Invoke_Func = " \n14"
        Sheets("597").Select
    Cells.Select
    Selection.Copy
    Sheets("T597").Select
    Cells.Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Sheets("620").Select
    Cells.Select
    Selection.Copy
    Sheets("T620").Select
    Cells.Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
End Sub
Sub portafolio()
'Rutina para generar y estandarizar los portafolios de los archivos T620 y T597
Sheets("T597").Select
Ultimaf = Application.WorksheetFunction.CountA(Range("ak:ak"))
For i = 4 To Ultimaf + 1

If Mid(Cells(i, 37), 4, 1) = "-" Then
Cells(i, 37) = Left(Trim(Cells(i, 37)), 2)
Else
Cells(i, 37) = Left(Trim(Cells(i, 37)), 3)
End If
'Cells(i, 37) = Left(Trim(Cells(i, 37)), 2)
Cells(i, 12) = DateSerial(Left(Cells(i, 12), 4), Mid(Cells(i, 12), 5, 2), Right(Cells(i, 12), 2))
Cells(i, 17) = DateSerial(Left(Cells(i, 17), 4), Mid(Cells(i, 17), 5, 2), Right(Cells(i, 17), 2))
If Trim(Cells(i, 13)) = "" Then
Cells(i, 13) = DateSerial(1900, 1, 1)
Else
Cells(i, 13) = DateSerial(Left(Cells(i, 13), 4), Mid(Cells(i, 13), 5, 2), Right(Cells(i, 13), 2))
End If
Cells(i, 36) = Trim(Cells(i, 36))
Next
Sheets("T620").Select
Ultimafa = Application.WorksheetFunction.CountA(Range("a:a"))
Counter = Ultimafa + 1
For i = 1 To Counter
    If Trim(Cells(i, 1)) = "NV" Then
        For j = 1 To 25
        Cells(i, j) = ""
        Next
    End If
Next i
Ultimafa = Application.WorksheetFunction.CountA(Range("a:a"))
For i = 4 To Ultimafa + 1
    If Mid(Cells(i, 24), 4, 1) = "-" Then
    Cells(i, 24) = Left(Trim(Cells(i, 24)), 2)
    Else
    Cells(i, 24) = Left(Trim(Cells(i, 24)), 3)
    End If
'Cells(i, 22) = Left(Trim(Cells(i, 22)), 2)
Next

End Sub

