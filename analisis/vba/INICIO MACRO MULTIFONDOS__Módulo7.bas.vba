Attribute VB_Name = "Módulo7"
Sub Copy()
Attribute Copy.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Sheets("T597").Select
    B = Application.WorksheetFunction.CountA(Range("A:A"))
    
    Sheets("VECTOR DE PRECIOS").Select
    
''''''''BORRA LA INFORMACION DE LA HOJA VECTOR DE PRECIOS'''''''''''
    
    Range("A3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = 3 To B
    Cells(i, 1) = Sheets("T597").Cells(i + 1, 3)
    Cells(i, 2) = Sheets("T597").Cells(i + 1, 8)
    Cells(i, 3) = Sheets("T597").Cells(i + 1, 9)
    Cells(i, 4) = "=IF(OR(TRIM(RC[-2])="""",TRIM(RC[-2])=""NA"",TRIM(RC[-1])=""NA""),IF(OR(TRIM(RC[-1])=""NA"",TRIM(RC[-2])=""NA""),TRIM(RC[-3]),TRIM(RC[-1])),TRIM(RC[-2]))"
    Cells(i, 5) = Sheets("T597").Cells(i + 1, 14)
    Cells(i, 6) = Trim(Sheets("T597").Cells(i + 1, 20))
    Cells(i, 7) = Sheets("T597").Cells(i + 1, 25)
    Cells(i, 8) = "=IF(RC[-2]=""$"",1,VLOOKUP(RC[-2],'IND2'!C1:C9,5,FALSE))"
    Cells(i, 9) = "=+RC[-2]/RC[-1]/RC[-4]"
    Cells(i, 10) = Sheets("T597").Cells(i + 1, 6)
    
    Next
    Calculate
    ActiveWorkbook.RefreshAll
    
End Sub
