Attribute VB_Name = "Módulo7"
Sub futuros_trm()
'
Sheets("Fwd Industria").Select

 Sheets("Fwd Industria").Select
    Range("A22").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.ClearContents
    
    Range("AG22").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.ClearContents

Dim Destino8 As Range
Set Destino8 = Worksheets("Fwd Industria").Range("A" & Rows.Count).End(xlUp).Offset(2)
Dim Destino9 As Range
Set Destino9 = Worksheets("Fwd Industria").Range("B" & Rows.Count).End(xlUp).Offset(2)
Dim Destino10 As Range
Set Destino10 = Worksheets("Fwd Industria").Range("C" & Rows.Count).End(xlUp).Offset(2)
Dim Destino11 As Range
Set Destino11 = Worksheets("Fwd Industria").Range("D" & Rows.Count).End(xlUp).Offset(2)
Dim Destino12 As Range
Set Destino12 = Worksheets("Fwd Industria").Range("E" & Rows.Count).End(xlUp).Offset(2)
Dim Destino13 As Range
Set Destino13 = Worksheets("Fwd Industria").Range("F" & Rows.Count).End(xlUp).Offset(2)
Dim Destino14 As Range
Set Destino14 = Worksheets("Fwd Industria").Range("H" & Rows.Count).End(xlUp).Offset(2)
Dim Destino15 As Range
Set Destino15 = Worksheets("Fwd Industria").Range("AG" & Rows.Count).End(xlUp).Offset(2)

Sheets("Fut Colf").Select
ActiveSheet.Range("$A$1:$BD$5000").AutoFilter Field:=21, Criteria1:=Array( _
        "TRM"), Operator:=xlFilterValues
Range("C2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Destino8.PasteSpecial xlPasteValues

Sheets("Fut Colf").Select
Range("BE3").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Destino9.PasteSpecial xlPasteValues

Sheets("Fut Colf").Select
Range("AH3").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Destino10.PasteSpecial xlPasteValues


Sheets("Fut Colf").Select
Range("K2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Destino11.PasteSpecial xlPasteValues


Sheets("Fut Colf").Select
Range("G2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Destino12.PasteSpecial xlPasteValues


Sheets("Fut Colf").Select
Range("R3").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Destino13.PasteSpecial xlPasteValues

Sheets("Fut Colf").Select
Range("D2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Destino14.PasteSpecial xlPasteValues

Sheets("Fut Colf").Select
Range("B2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Destino15.PasteSpecial xlPasteValues

Sheets("Fwd Industria").Select
LastRow22 = Range("A" & Rows.Count).End(xlUp).Row
Range("G23:G23").AutoFill Destination:=Range("G23:G" & LastRow22), Type:=xlFillDefault

LastRow23 = Range("A" & Rows.Count).End(xlUp).Row
Range("I23:AF23").AutoFill Destination:=Range("I23:AF" & LastRow23), Type:=xlFillDefault

LastRow24 = Range("A" & Rows.Count).End(xlUp).Row
Range("AH23:AH23").AutoFill Destination:=Range("AH23:AH" & LastRow24), Type:=xlFillDefault

Sheets("Fut Colf").Select

ActiveSheet.ShowAllData

Sheets("Controles").Select

End Sub


