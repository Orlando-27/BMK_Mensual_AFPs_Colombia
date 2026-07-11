Attribute VB_Name = "Módulo1"
Sub ImportarCurvas()
Calculate
'Asigna nombres
Ruta = Range("Ruta_Inf").Value
Ruta_Indica = Range("Ruta_Indica").Value
Pt_Fwd = Range("Pt_Fwd").Value
Pt_FwdEUR = Range("Pt_FwdEUR").Value
SWP_IBR = Range("SWP_IBR").Value
SWP_Impli = Range("SWP_Impli").Value
SWP_ImpliEUR = Range("SWP_ImpliEUR").Value
Indicadores = Range("Indicadores").Value


PuntosFwd = Ruta & Pt_Fwd
PuntosFwdEUR = Ruta & Pt_FwdEUR
SwapIBR = Ruta & SWP_IBR
SwapImplicita = Ruta & SWP_Impli
SwapImplicitaEUR = Ruta & SWP_ImpliEUR
File_Indicadores = Ruta_Indica & Indicadores

'Borra la información de la hoja Curvas entre A:AA
Sheets("Curvas").Select
Range("A:AZ").ClearContents

'Trae los puntos FWD en la columna A:B
Range("A1").Value = "Puntos Fwd USDCOP"
Range("A2").Select
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" + PuntosFwd, Destination:=Range("$A$2"))
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 850
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = True
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = True
        .TextFileColumnDataTypes = Array(1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
'Trae la curva SWAP IBR en las columnas D:E
Range("D1").Value = "Curva SWAP IBR"
Range("D2").Select

   With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + SwapIBR, Destination:=Range("$d$2"))
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 850
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = True
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = True
        .TextFileColumnDataTypes = Array(1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
End With

'Trae la curva implicita USDCOP en las columnas G:H
Range("G1").Value = "Curva SWAP Implicita USDCOP"
Range("G2").Select
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" + SwapImplicita, Destination:=Range("$G$2"))
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 850
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = True
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = True
        .TextFileColumnDataTypes = Array(1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With



'Trae los indicadores del día en la columna R:Z
With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + File_Indicadores, Destination:=Range("$R$1"))
       
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 1252
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierNone
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = True
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With

'Trae los puntos FWD EURCOP en la columna AC:AD
Range("AC1").Value = "Puntos Fwd EURCOP"
Range("AC2").Select
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" + PuntosFwdEUR, Destination:=Range("$AC$2"))
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 850
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = True
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = True
        .TextFileColumnDataTypes = Array(1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With

Range("AE:AF").ClearContents


'Trae la curva implicita EURCOP en las columnas AF:AG
Range("AF1").Value = "Curva SWAP Implicita EURCOP"
Range("AF2").Select

LIBRO = ActiveWorkbook.Name

'importa Tasas_EURCOP_Diaria_
'
'    With ActiveSheet.QueryTables.Add(Connection:= _
'        "TEXT;" + SwapImplicitaEUR, Destination:=Range("$AF$2"))
'        .FieldNames = True
'        .RowNumbers = False
'        .FillAdjacentFormulas = False
'        .PreserveFormatting = True
'        .RefreshOnFileOpen = False
'        .RefreshStyle = xlInsertDeleteCells
'        .SavePassword = False
'        .SaveData = True
'        .AdjustColumnWidth = True
'        .RefreshPeriod = 0
'        .TextFilePromptOnRefresh = False
'        .TextFilePlatform = 850
'        .TextFileStartRow = 1
'        .TextFileParseType = xlDelimited
'        .TextFileTextQualifier = xlTextQualifierDoubleQuote
'        .TextFileConsecutiveDelimiter = True
'        .TextFileTabDelimiter = True
'        .TextFileSemicolonDelimiter = False
'        .TextFileCommaDelimiter = False
'        .TextFileSpaceDelimiter = True
'        .TextFileColumnDataTypes = Array(1, 1)
'        .TextFileTrailingMinusNumbers = True
'        .Refresh BackgroundQuery:=False
'    End With

'Importa SwapCC2_EURCOP_Diaria_
Workbooks.Open SwapImplicitaEUR
    EURCOP = ActiveWorkbook.Name
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
        Array(7, 1), Array(8, 1), Array(9, 1)), TrailingMinusNumbers:=True
    Range("G2:H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

Workbooks(LIBRO).Activate
    Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False

Workbooks(EURCOP).Close SaveChanges:=False

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Range("AH:AI").ClearContents

Calculate

'Establece el largo de las columnas para hacer calculos con base en la cantidad mínima para USDCOP
rango1 = Worksheets("Curvas").Cells(Rows.Count, 1).End(xlUp).Row
rango2 = Worksheets("Curvas").Cells(Rows.Count, 4).End(xlUp).Row
rango3 = Worksheets("Curvas").Cells(Rows.Count, 7).End(xlUp).Row

rango = WorksheetFunction.Min(rango1, rango2, rango3)

'Rótulos en las celdas J1:P1
Range("J1").Value = "Tasa FWD"
Range("K1").Value = "Puntos Fwd"
Range("L1").Value = "Loc +100pb"
Range("M1").Value = "Loc -100pb"
Range("N1").Value = "For +100pb"
Range("O1").Value = "For -100pb"
Range("P1").Value = "Validador"

'Fórmulas en J2:P2 para estimar la curva FWD con movimientos de 100 pbs tanto local como foranea
Range("J2").FormulaR1C1 = "=TRM*((1+RC[-5]*RC[-6]/360)/(1+RC[-2]*RC[-3]/360))"
Range("K2").FormulaR1C1 = "=+RC[-1]-TRM"
Range("L2").FormulaR1C1 = "=TRM*((1+(1%+RC[-7])*RC[-8]/360)/(1+RC[-4]*RC[-5]/360))-TRM"
Range("M2").FormulaR1C1 = "=TRM*((1+(-1%+RC[-8])*RC[-9]/360)/(1+RC[-5]*RC[-6]/360))-TRM"
Range("N2").FormulaR1C1 = "=TRM*((1+RC[-9]*RC[-10]/360)/(1+(1%+RC[-6])*RC[-7]/360))-TRM"
Range("O2").FormulaR1C1 = "=TRM*((1+RC[-10]*RC[-11]/360)/(1+(-1%+RC[-7])*RC[-8]/360))-TRM"
Range("P2").FormulaR1C1 = "=ROUND(RC[-5],5)=ROUND(RC[-14],5)"

'Copia las fórmulas y las pega como valores
Range("J2:P2").Copy
Range("J2:P" + Format(rango, "00")).PasteSpecial xlPasteFormulas
Range("J2:P" + Format(rango, "00")).PasteSpecial xlPasteFormats
Calculate
Range("J3:P" + Format(rango, "00")).Copy
Range("J3:P" + Format(rango, "00")).PasteSpecial xlPasteValues
Range("A1").Select


'Establece el largo de las columnas para hacer calculos con base en la cantidad mínima para EURCOP
rango4 = Worksheets("Curvas").Cells(Rows.Count, 29).End(xlUp).Row
rango5 = Worksheets("Curvas").Cells(Rows.Count, 4).End(xlUp).Row
rango6 = Worksheets("Curvas").Cells(Rows.Count, 32).End(xlUp).Row

rangoEUR = WorksheetFunction.Min(rango4, rango5, rango6)

'Rótulos en las celdas J1:P1
Range("AI1").Value = "Tasa FWD"
Range("AJ1").Value = "Puntos Fwd"
Range("AK1").Value = "Loc +100pb"
Range("AL1").Value = "Loc -100pb"
Range("AM1").Value = "For +100pb"
Range("AN1").Value = "For -100pb"
Range("AO1").Value = "Validador"

'Fórmulas en J2:P2 para estimar la curva FWD con movimientos de 100 pbs tanto local como foranea
Range("AI2").FormulaR1C1 = "=EUR*((1+RC[-30]*RC[-31]/360)/(1+RC[-2]*RC[-3]/360))"
Range("AJ2").FormulaR1C1 = "=RC[-1]-EUR"
Range("AK2").FormulaR1C1 = "=EUR*((1+(1%+RC[-32])*RC[-33]/360)/(1+RC[-4]*RC[-5]/360))-EUR"
Range("AL2").FormulaR1C1 = "=EUR*((1+(-1%+RC[-33])*RC[-34]/360)/(1+RC[-5]*RC[-6]/360))-EUR"
Range("AM2").FormulaR1C1 = "=EUR*((1+RC[-34]*RC[-35]/360)/(1+(1%+RC[-6])*RC[-7]/360))-EUR"
Range("AN2").FormulaR1C1 = "=EUR*((1+RC[-35]*RC[-36]/360)/(1+(-1%+RC[-7])*RC[-8]/360))-EUR"
Range("AO2").FormulaR1C1 = "=ROUND(RC[-11],3)=ROUND(RC[-5],3)"


'Copia las fórmulas y las pega como valores
Range("AI2:AO2").Copy
Range("AI2:AO" + Format(rangoEUR, "00")).PasteSpecial xlPasteFormulas
Range("AI2:AO" + Format(rangoEUR, "00")).PasteSpecial xlPasteFormats
Calculate
Range("AI3:AO" + Format(rangoEUR, "00")).Copy
Range("AI3:AO" + Format(rangoEUR, "00")).PasteSpecial xlPasteValues
Range("A1").Select


Sheets("Control").Select
Limpiar_Datos_externos
End Sub
Sub Limpiar_Datos_externos()
    Dim nm As Name
    Dim Hoja

    For Each nm In ThisWorkbook.Names
        posicion_busqueda = InStr(1, nm.Name, "DatosExternos")
        If posicion_busqueda <> 0 Then
        nm.Delete
        End If
    Next nm

End Sub
Sub Importar572()

Dim ultimaColumna As Long
Dim ultimaFila As Long

'Asigna nombres
Ruta_572 = Range("Ruta_572").Value
Plano572 = Range("Plano572").Value

File_572 = Ruta_572 & Plano572

'Borra la información de la hoja 572
Sheets("572").Select
Cells.ClearContents

Sensibilidad = ActiveWorkbook.Name

'Crea un archivo nuevo para abrir ahí el 572
Workbooks.Add


With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + File_572, Destination:=Range("$A$1"))
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 850
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = True
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1) 'se cambia el numero de columnas
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With

Range("AA:AD").Columns.Delete
'Elimina los fwd diferentes a USDCOP
ActiveSheet.Range("A$3:z$20000").AutoFilter field:=12, Criteria1:="US$   "
Rows("4:2000").EntireRow.Select
Selection.SpecialCells(xlCellTypeVisible).Select
Selection.Delete


'Copia la información de FWD USDCOP
'ActiveSheet.ShowAllData
'Range("A2").Select
'Range(Selection, Selection.End(xlToRight)).Select
'Range(Selection, Selection.End(xlDown)).Select
'Selection.Copy


'Se cambia por este - Copia solo las primeras 29 columnas (A a AC)
ActiveSheet.ShowAllData
Range("A2").Select
Range(Selection, Selection.End(xlToRight).Offset(0, -1)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

borrar = ActiveWorkbook.Name

'Pega la base de FWD en la hoja 572 del archivo Sensibilidad_Fwds
Workbooks(Sensibilidad).Activate
Sheets("572").Select
Range("a1").PasteSpecial xlPasteAll
Application.CutCopyMode = False


'Cierra el archivo creado sin cambios
Workbooks(borrar).Close SaveChanges:=False

Workbooks(Sensibilidad).Activate


ActiveSheet.Range("a1").End(xlToRight).Offset(0, 2).Select

'Rótulos en las celdas AB1:AQ2
ActiveCell.Value = "Fondo"
ActiveCell.Offset(0, 1).Value = "Tipo Operación"
ActiveCell.Offset(0, 2).Value = "Plazo"
ActiveCell.Offset(0, 3).Value = "Strike"
ActiveCell.Offset(0, 4).Value = "Valoración"
ActiveCell.Offset(0, 5).Value = "Valoración"
ActiveCell.Offset(0, 6).Value = "Sube IBR"
ActiveCell.Offset(0, 7).Value = "Sube IBR"
ActiveCell.Offset(0, 8).Value = "Baja IBR"
ActiveCell.Offset(0, 9).Value = "Baja IBR"
ActiveCell.Offset(0, 10).Value = "Sube Impli"
ActiveCell.Offset(0, 11).Value = "Sube Impli"
ActiveCell.Offset(0, 12).Value = "Baja Impli"
ActiveCell.Offset(0, 13).Value = "Baja Impli"
ActiveCell.Offset(0, 14).Value = "KRD Local"
ActiveCell.Offset(0, 15).Value = "KRD Foran"
ActiveCell.Offset(1, 4).Value = "Derecho"
ActiveCell.Offset(1, 5).Value = "Obligación"
ActiveCell.Offset(1, 6).Value = "Derecho"
ActiveCell.Offset(1, 7).Value = "Obligación"
ActiveCell.Offset(1, 8).Value = "Derecho"
ActiveCell.Offset(1, 9).Value = "Obligación"
ActiveCell.Offset(1, 10).Value = "Derecho"
ActiveCell.Offset(1, 11).Value = "Obligación"
ActiveCell.Offset(1, 12).Value = "Derecho"
ActiveCell.Offset(1, 13).Value = "Obligación"

'Fórmulas en las celdas AB3:AQ3 para calcular la tasas fwd de cada operación y el KRD
Range("AB3").FormulaR1C1 = "=MID(RC[-4],2,2)"
Range("AC3").FormulaR1C1 = "=TRIM(RC[-27])"
Range("AD3").FormulaR1C1 = "=RC[-23]-Fecha"
Range("AE3").FormulaR1C1 = "=RC[-18]"
Range("AF3").FormulaR1C1 = "=IF(RC29=""Compra"",(TRM+VLOOKUP(RC30,Curvas!C7:C11,5,0))/(1+VLOOKUP(RC30,Curvas!C4:C5,2,0)*RC30/360),IF(RC29=""Venta"",RC31/(1+VLOOKUP(RC30,Curvas!C4:C5,2,0)*RC30/360),""Error""))"
Range("AG3").FormulaR1C1 = "=IF(RC29=""Venta"",(TRM+VLOOKUP(RC30,Curvas!C7:C11,5,0))/(1+VLOOKUP(RC30,Curvas!C4:C5,2,0)*RC30/360),IF(RC29=""Compra"",RC31/(1+VLOOKUP(RC30,Curvas!C4:C5,2,0)*RC30/360),""Error""))"
Range("AH3").FormulaR1C1 = "=IF(RC29=""Compra"",(TRM+VLOOKUP(RC30,Curvas!C7:C15,6,0))/(1+(1%+VLOOKUP(RC30,Curvas!C4:C5,2,0))*RC30/360),IF(RC29=""Venta"",RC31/(1+(1%+VLOOKUP(RC30,Curvas!C4:C5,2,0))*RC30/360),""Error""))"
Range("AI3").FormulaR1C1 = "=IF(RC29=""Venta"",(TRM+VLOOKUP(RC30,Curvas!C7:C15,6,0))/(1+(1%+VLOOKUP(RC30,Curvas!C4:C5,2,0))*RC30/360),IF(RC29=""Compra"",RC31/(1+(1%+VLOOKUP(RC30,Curvas!C4:C5,2,0))*RC30/360),""Error""))"
Range("AJ3").FormulaR1C1 = "=IF(RC29=""Compra"",(TRM+VLOOKUP(RC30,Curvas!C7:C15,7,0))/(1+(-1%+VLOOKUP(RC30,Curvas!C4:C5,2,0))*RC30/360),IF(RC29=""Venta"",RC31/(1+(-1%+VLOOKUP(RC30,Curvas!C4:C5,2,0))*RC30/360),""Error""))"
Range("AK3").FormulaR1C1 = "=IF(RC29=""Venta"",(TRM+VLOOKUP(RC30,Curvas!C7:C15,7,0))/(1+(-1%+VLOOKUP(RC30,Curvas!C4:C5,2,0))*RC30/360),IF(RC29=""Compra"",RC31/(1+(-1%+VLOOKUP(RC30,Curvas!C4:C5,2,0))*RC30/360),""Error""))"
Range("AL3").FormulaR1C1 = "=IF(RC29=""Compra"",(TRM+VLOOKUP(RC30,Curvas!C7:C15,8,0))/(1+VLOOKUP(RC30,Curvas!C4:C5,2,0)*RC30/360),IF(RC29=""Venta"",RC31/(1+VLOOKUP(RC30,Curvas!C4:C5,2,0)*RC30/360),""Error""))"
Range("AM3").FormulaR1C1 = "=IF(RC29=""Venta"",(TRM+VLOOKUP(RC30,Curvas!C7:C15,8,0))/(1+VLOOKUP(RC30,Curvas!C4:C5,2,0)*RC30/360),IF(RC29=""Compra"",RC31/(1+VLOOKUP(RC30,Curvas!C4:C5,2,0)*RC30/360),""Error""))"
Range("AN3").FormulaR1C1 = "=IF(RC29=""Compra"",(TRM+VLOOKUP(RC30,Curvas!C7:C15,9,0))/(1+VLOOKUP(RC30,Curvas!C4:C5,2,0)*RC30/360),IF(RC29=""Venta"",RC31/(1+VLOOKUP(RC30,Curvas!C4:C5,2,0)*RC30/360),""Error""))"
Range("AO3").FormulaR1C1 = "=IF(RC29=""Venta"",(TRM+VLOOKUP(RC30,Curvas!C7:C15,9,0))/(1+VLOOKUP(RC30,Curvas!C4:C5,2,0)*RC30/360),IF(RC29=""Compra"",RC31/(1+VLOOKUP(RC30,Curvas!C4:C5,2,0)*RC30/360),""Error""))"
Range("AP3").FormulaR1C1 = "=IF(RC29=""Compra"",-1*((RC[-5]-RC[-7])/(2*1%*RC33)),IF(RC29=""Venta"",(RC[-6]-RC[-8])/(2*1%*RC32),""Error""))"
Range("AQ3").FormulaR1C1 = "=IF(RC29=""Venta"",-1*((RC[-2]-RC[-4])/(2*1%*RC33)),IF(RC29=""Compra"",(RC[-3]-RC[-5])/(2*1%*RC32),""Error""))"

rango572 = Worksheets("572").Cells(Rows.Count, 1).End(xlUp).Row

'Copia las formulas y las pega como valores
Range("AB3:AQ3").Copy
Range("AB4:AQ" + Format(rango572, "00")).PasteSpecial xlPasteFormulas
Range("AB4:AQ" + Format(rango572, "00")).PasteSpecial xlPasteFormats
Calculate
Range("AB4:AQ" + Format(rango572, "00")).Copy
Range("AB4:AQ" + Format(rango572, "00")).PasteSpecial xlPasteValues
Application.CutCopyMode = False

'Corre la macro que calcula la valoración de la industria
Range("AC3").Select
Call CalcularInd
Sheets("Control").Select
'
    
End Sub

Sub Actualizar_Industria()

'Asigna nombres
Ruta_Ind = Range("Ruta_Ind").Value
BMK = Range("BMK").Value
FileBMK = Ruta_Ind & BMK
Sensibilidad = ActiveWorkbook.Name

'Borra la información de la hoja Ind
Sheets("Ind").Select
Range("A3:AA3").Select
Range(Selection, Selection.End(xlDown)).ClearContents

'Abre el archivo Bencharmk detallado y copia los fwds de la industria
Workbooks.Open (FileBMK)
Sheets("Fwd Industria").Select
Range("A23:H23").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

'pega la base en la hoja Ind
Workbooks(Sensibilidad).Activate

Sheets("Ind").Select
Range("A3").PasteSpecial xlPasteValues
Range("A3").PasteSpecial xlPasteFormats

Application.CutCopyMode = False

'Cierra el benchmark sin cambios
Workbooks(BMK).Close SaveChanges:=False
Workbooks(Sensibilidad).Activate

'Solicita el calculo de KRD para la industria
Call CalcularInd

End Sub
Sub CalcularInd()

rangoIND = Worksheets("Ind").Cells(Rows.Count, 1).End(xlUp).Row

Sheets("Ind").Select
Range("L:AA").ClearContents
ActiveSheet.Range("A2").End(xlToRight).Offset(-1, 4).Select

'Rótulos en L1:AA2
ActiveCell.Value = "Fondo"
ActiveCell.Offset(0, 1).Value = "Tipo Operación"
ActiveCell.Offset(0, 2).Value = "Plazo"
ActiveCell.Offset(0, 3).Value = "Strike"
ActiveCell.Offset(0, 4).Value = "Valoración"
ActiveCell.Offset(0, 5).Value = "Valoración"
ActiveCell.Offset(0, 6).Value = "Sube IBR"
ActiveCell.Offset(0, 7).Value = "Sube IBR"
ActiveCell.Offset(0, 8).Value = "Baja IBR"
ActiveCell.Offset(0, 9).Value = "Baja IBR"
ActiveCell.Offset(0, 10).Value = "Sube Impli"
ActiveCell.Offset(0, 11).Value = "Sube Impli"
ActiveCell.Offset(0, 12).Value = "Baja Impli"
ActiveCell.Offset(0, 13).Value = "Baja Impli"
ActiveCell.Offset(0, 14).Value = "KRD Local"
ActiveCell.Offset(0, 15).Value = "KRD Foran"
ActiveCell.Offset(1, 4).Value = "Derecho"
ActiveCell.Offset(1, 5).Value = "Obligación"
ActiveCell.Offset(1, 6).Value = "Derecho"
ActiveCell.Offset(1, 7).Value = "Obligación"
ActiveCell.Offset(1, 8).Value = "Derecho"
ActiveCell.Offset(1, 9).Value = "Obligación"
ActiveCell.Offset(1, 10).Value = "Derecho"
ActiveCell.Offset(1, 11).Value = "Obligación"
ActiveCell.Offset(1, 12).Value = "Derecho"
ActiveCell.Offset(1, 13).Value = "Obligación"

'Fórmulas en J3:AA3
Range("J3").FormulaR1C1 = "=IF(RC[-2]-Fecha<=0,Fecha+(31),RC[-2])"
Range("L3").FormulaR1C1 = "=RC[-6]"
Range("M3").FormulaR1C1 = "=IF(RC[-6]=""SELL"",""Venta"",IF(RC[-6]=""Buy"",""Compra"",""Error""))"
Range("N3").FormulaR1C1 = "=+RC[-4]-Fecha"
Range("O3").FormulaR1C1 = "=RC[-10]"
Range("P3").FormulaR1C1 = "=IF(RC13=""Compra"",(IF(RC[-14]=""EURCOP"",EUR,TRM)+VLOOKUP(RC14,IF(RC[-14]=""EURCOP"",Curvas!C32:C36,Curvas!C7:C11),5,0))/(1+VLOOKUP(RC14,Curvas!C4:C5,2,0)*RC14/360),IF(RC13=""Venta"",RC15/(1+VLOOKUP(RC14,Curvas!C4:C5,2,0)*RC14/360),""Error""))"
Range("Q3").FormulaR1C1 = "=IF(RC13=""Venta"",(IF(RC[-15]=""EURCOP"",EUR,TRM)+VLOOKUP(RC14,IF(RC[-15]=""EURCOP"",Curvas!C32:C36,Curvas!C7:C11),5,0))/(1+VLOOKUP(RC14,Curvas!C4:C5,2,0)*RC14/360),IF(RC13=""Compra"",RC15/(1+VLOOKUP(RC14,Curvas!C4:C5,2,0)*RC14/360),""Error""))"
Range("R3").FormulaR1C1 = "=IF(RC13=""Compra"",(IF(RC[-16]=""EURCOP"",EUR,TRM)+VLOOKUP(RC14,IF(RC[-16]=""EURCOP"",Curvas!C32:C40,Curvas!C7:C15),6,0))/(1+(1%+VLOOKUP(RC14,Curvas!C4:C5,2,0))*RC14/360),IF(RC13=""Venta"",RC15/(1+(1%+VLOOKUP(RC14,Curvas!C4:C5,2,0))*RC14/360),""Error""))"
Range("S3").FormulaR1C1 = "=IF(RC13=""Venta"",(IF(RC[-17]=""EURCOP"",EUR,TRM)+VLOOKUP(RC14,IF(RC[-17]=""EURCOP"",Curvas!C32:C40,Curvas!C7:C15),6,0))/(1+(1%+VLOOKUP(RC14,Curvas!C4:C5,2,0))*RC14/360),IF(RC13=""Compra"",RC15/(1+(1%+VLOOKUP(RC14,Curvas!C4:C5,2,0))*RC14/360),""Error""))"
Range("T3").FormulaR1C1 = "=IF(RC13=""Compra"",(IF(RC[-18]=""EURCOP"",EUR,TRM)+VLOOKUP(RC14,IF(RC[-18]=""EURCOP"",Curvas!C32:C40,Curvas!C7:C15),7,0))/(1+(-1%+VLOOKUP(RC14,Curvas!C4:C5,2,0))*RC14/360),IF(RC13=""Venta"",RC15/(1+(-1%+VLOOKUP(RC14,Curvas!C4:C5,2,0))*RC14/360),""Error""))"
Range("U3").FormulaR1C1 = "=IF(RC13=""Venta"",(IF(RC[-19]=""EURCOP"",EUR,TRM)+VLOOKUP(RC14,IF(RC[-19]=""EURCOP"",Curvas!C32:C40,Curvas!C7:C15),7,0))/(1+(-1%+VLOOKUP(RC14,Curvas!C4:C5,2,0))*RC14/360),IF(RC13=""Compra"",RC15/(1+(-1%+VLOOKUP(RC14,Curvas!C4:C5,2,0))*RC14/360),""Error""))"
Range("V3").FormulaR1C1 = "=IF(RC13=""Compra"",(IF(RC[-20]=""EURCOP"",EUR,TRM)+VLOOKUP(RC14,IF(RC[-20]=""EURCOP"",Curvas!C32:C40,Curvas!C7:C15),8,0))/(1+VLOOKUP(RC14,Curvas!C4:C5,2,0)*RC14/360),IF(RC13=""Venta"",RC15/(1+VLOOKUP(RC14,Curvas!C4:C5,2,0)*RC14/360),""Error""))"
Range("W3").FormulaR1C1 = "=IF(RC13=""Venta"",(IF(RC[-21]=""EURCOP"",EUR,TRM)+VLOOKUP(RC14,IF(RC[-21]=""EURCOP"",Curvas!C32:C40,Curvas!C7:C15),8,0))/(1+VLOOKUP(RC14,Curvas!C4:C5,2,0)*RC14/360),IF(RC13=""Compra"",RC15/(1+VLOOKUP(RC14,Curvas!C4:C5,2,0)*RC14/360),""Error""))"
Range("X3").FormulaR1C1 = "=IF(RC13=""Compra"",(IF(RC[-22]=""EURCOP"",EUR,TRM)+VLOOKUP(RC14,IF(RC[-22]=""EURCOP"",Curvas!C32:C40,Curvas!C7:C15),9,0))/(1+VLOOKUP(RC14,Curvas!C4:C5,2,0)*RC14/360),IF(RC13=""Venta"",RC15/(1+VLOOKUP(RC14,Curvas!C4:C5,2,0)*RC14/360),""Error""))"
Range("Y3").FormulaR1C1 = "=IF(RC13=""Venta"",(IF(RC[-23]=""EURCOP"",EUR,TRM)+VLOOKUP(RC14,IF(RC[-23]=""EURCOP"",Curvas!C32:C40,Curvas!C7:C15),9,0))/(1+VLOOKUP(RC14,Curvas!C4:C5,2,0)*RC14/360),IF(RC13=""Compra"",RC15/(1+VLOOKUP(RC14,Curvas!C4:C5,2,0)*RC14/360),""Error""))"
Range("Z3").FormulaR1C1 = "=IF(OR(RC[-24]=""USDCOP"",RC[-24]=""EURCOP""),IF(RC13=""Compra"",-1*((RC[-5]-RC[-7])/(2*1%*RC17)),IF(RC13=""Venta"",(RC[-6]-RC[-8])/(2*1%*RC16),""Error"")),"""")"
Range("AA3").FormulaR1C1 = "=IF(OR(RC[-25]=""USDCOP"",RC[-25]=""EURCOP""),IF(RC13=""Venta"",-1*((RC[-2]-RC[-4])/(2*1%*RC17)),IF(RC13=""Compra"",(RC[-3]-RC[-5])/(2*1%*RC16),""Error"")),"""")"



'Copia las fórmulas
Range("J3:AA3").Copy
Range("J4:AA" + Format(rangoIND, "00")).PasteSpecial xlPasteFormulas
Range("J4:AA" + Format(rangoIND, "00")).PasteSpecial xlPasteFormats
Calculate
Range("J4:AA" + Format(rangoIND, "00")).Copy
Range("J4:AA" + Format(rangoIND, "00")).PasteSpecial xlPasteValues
Application.CutCopyMode = False

Range("L3").Select

Sheets("Control").Select

End Sub



