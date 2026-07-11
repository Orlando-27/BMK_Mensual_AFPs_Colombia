Attribute VB_Name = "Módulo4"


Sub Macro1()
'
' Macro1 Macro
'

'
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
End Sub

Sub traer_sensibilidad()
Dim CP As Double
Dim LP As Double

Ruta_BMK = Range("Ruta_Bmk").Value
CP = Range("CP").Value
LP = Range("LP").Value
Sensibilidad = Range("Sensibilidad").Value

FileSensibilidad = Ruta_BMK & Sensibilidad
BMK = ActiveWorkbook.Name

Sheets("Fwd Colfondos").Select
Range("U21:Y21").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

Range("U21").Value = "Plazo"
Range("V21").Value = "KRD Local"
Range("W21").Value = "KRD Foraneo"
Range("X21").Value = "Sens Local"
Range("Y21").Value = "Sens Foraneo"

Sheets("Fwd Industria").Select
Range("S22:W22").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

Range("S22").Value = "Plazo"
Range("T22").Value = "KRD Local"
Range("U22").Value = "KRD Foraneo"
Range("V22").Value = "Sens Local"
Range("W22").Value = "Sens Foraneo"

Workbooks.Open (FileSensibilidad)
Sheets("Ind").Select

rangocopia = Worksheets("Ind").Cells(Rows.Count, 1).End(xlUp).Row
Range("Z3:AA" + Format(rangocopia, "00")).Copy
Workbooks(BMK).Activate
Range("T23").PasteSpecial xlPasteValues
Range("T23").PasteSpecial xlPasteFormats
Range("S23").FormulaR1C1 = "=IF((RC[-10]-R18C2)/365=0,""NA"",IF((RC[-10]-R18C2)/365<CP,""CORTO PLAZO"",IF((RC[-10]-R18C2)/365>=LP,""LARGO PLAZO"",""MEDIANO PLAZO"")))"

Range("V23").FormulaR1C1 = "=IF(RC[-2]="""","""",RC[-2]*(IF(RC[-19]=""Compra"",ABS(RC[-8]),ABS(RC[-9]))/INDEX(Benchmark!R2C9:R6C11,MATCH('Fwd Industria'!RC[-6],Benchmark!R2C8:R6C8,0),MATCH('Fwd Industria'!RC[11],Benchmark!R1C9:R1C11,0))))"
Range("W23").FormulaR1C1 = "=IF(RC[-2]="""","""",RC[-2]*(IF(RC[-20]=""Compra"",ABS(RC[-10]),ABS(RC[-9]))/INDEX(Benchmark!R2C9:R6C11,MATCH('Fwd Industria'!RC[-7],Benchmark!R2C8:R6C8,0),MATCH('Fwd Industria'!RC[10],Benchmark!R1C9:R1C11,0))))"


Refer = Worksheets("Fwd Industria").Cells(Rows.Count, 1).End(xlUp).Row
Range("S23").Copy
Range("S23:S" + Format(Refer, "00")).PasteSpecial xlPasteFormulas
Range("S23:S" + Format(Refer, "00")).Select
Selection.Calculate
Range("S23:S" + Format(Refer, "00")).Copy
Range("S23").PasteSpecial xlPasteValues

Range("V23:W23").Copy
Range("V23:W" + Format(Refer, "00")).PasteSpecial xlPasteFormulas


Sheets("Fwd Colfondos").Select

Range("U22").FormulaR1C1 = "=IF(RC[-20]/365=0,""NA"",IF(RC[-20]/365<CP,""CORTO PLAZO"",IF(RC[-20]/365>=LP,""LARGO PLAZO"",""MEDIANO PLAZO"")))"
Range("V22").FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-18],'[Sensibilidad_Fwds.xlsb]572'!C5:C43,38,0),"""")"
Range("W22").FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-19],'[Sensibilidad_Fwds.xlsb]572'!C5:C43,39,0),"""")"
Range("X22").FormulaR1C1 = "=IF(RC[-2]="""","""",RC[-2]*(IF(RC[-22]=""Compra"",ABS(RC[-5]),ABS(RC[-6]))/HLOOKUP(IF(ISNUMBER(RC[-7]),TEXT(RC[-7],""##""),RC[-7]),Colfondos!R1C7:R2C16384,2,0)))"
Range("Y22").FormulaR1C1 = "=IF(RC[-2]="""","""",RC[-2]*(IF(RC[-23]=""Compra"",ABS(RC[-7]),ABS(RC[-6]))/HLOOKUP(IF(ISNUMBER(RC[-8]),TEXT(RC[-8],""##""),RC[-8]),Colfondos!R1C7:R2C16384,2,0)))"
    
Refer1 = Worksheets("Fwd Colfondos").Cells(Rows.Count, 1).End(xlUp).Row
Range("U22:Y22").Copy
Range("U23:Y" + Format(Refer1, "00")).PasteSpecial xlPasteFormulas
Range("U23:Y" + Format(Refer1, "00")).Select
Selection.Calculate
Range("U22:W" + Format(Refer1, "00")).Copy
Range("U22").PasteSpecial xlPasteValues

Workbooks(Sensibilidad).Close SaveChanges:=False

Application.CutCopyMode = False


End Sub






