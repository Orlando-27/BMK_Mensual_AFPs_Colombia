Attribute VB_Name = "Módulo3"
'Consolida el vector que necesitamos para el precio y tasa facial en activos,
'trae la hoja sx y r int del vector, además de los isines consolidados y sus precios

Sub V_PRECIOS_IMPORTAR_COMPLETO()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    
    StartTime = Timer
    
    Application.Calculation = xlManual
        Application.ScreenUpdating = False
        
    Calculate

    Dim hojas_v(1 To 2) As String
    hojas_v(1) = "SX"
    hojas_v(2) = "R.Fija Int"

    Dim rutaArchivo As String
    rutaArchivo = Sheets("Parámetros").Range("C14").Value

    Dim wbOrigen As Workbook
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim i As Integer

    Set wbOrigen = Workbooks.Open(Filename:=rutaArchivo, ReadOnly:=True)

    'Hojas SX y R.Fija Int
For i = 1 To 2
    Set wsDestino = ThisWorkbook.Sheets(hojas_v(i))
    Set wsOrigen = wbOrigen.Sheets(hojas_v(i))

    wsDestino.Cells.ClearContents
    Dim lastRow As Long, lastCol As Long
    With wsOrigen.UsedRange
        lastRow = .Rows(.Rows.Count).Row
        lastCol = .Columns(.Columns.Count).Column
    End With
    Dim inicioFila As Long
    If hojas_v(i) = "SX" Then
        inicioFila = 2
    Else
        inicioFila = 1
    End If
    Dim numFilas As Long
    numFilas = lastRow - inicioFila + 1

    wsDestino.Cells(1, 1).Resize(numFilas, lastCol).Value = _
        wsOrigen.Cells(inicioFila, 1).Resize(numFilas, lastCol).Value
    Dim arr As Variant
    arr = wsDestino.Range("A2", wsDestino.Cells(numFilas, lastCol)).Value

    Dim r As Long, c As Long
    For r = 1 To UBound(arr, 1)
        For c = 1 To UBound(arr, 2)
            If arr(r, c) = 0 Then arr(r, c) = ""
        Next c
    Next r
    wsDestino.Range("A2").Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr
Next i

    'A y B Vector (ISIN, PRECIO)
    Set wsDestino = ThisWorkbook.Sheets("VECTOR")
    Set wsOrigen = wbOrigen.Sheets("Vector")
    wsDestino.Range("A4:B800000").ClearContents
    
    lastRow = wsOrigen.Cells(wsOrigen.Rows.Count, 1).End(xlUp).Row
    wsOrigen.Range("A4:B" & lastRow).Copy
    wsDestino.Range("A4:B" & lastRow).PasteSpecial Paste:=xlPasteValues

    wbOrigen.Close SaveChanges:=False

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Call FormulaVECTOR
    
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
     
    SecondsElapsed = Round(Timer - StartTime, 2)
    Sheets("Parámetros").Activate
    Sheets("Parámetros").Range("L6") = SecondsElapsed
    
    Sheets("Parámetros").Select
    
End Sub
'Arrastra la formula para el nuevo vector importado para buscar la tasa facial de RF
Sub FormulaVECTOR()
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = Sheets("VECTOR")
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    ws.Range("C5:C" & lastRow).ClearContents

    ws.Range("C4").AutoFill Destination:=ws.Range("C4:C" & lastRow), Type:=xlFillDefault
    Calculate
    Call Importar_indicadores
End Sub
'En esta parte nos desahacemos del portafolio fondo alternativo y los titulos con valor de mercado menor a 1000
Sub F_Alternativo_Mercado_1000()
    Dim wsClasif As Worksheet, wsEliminados As Worksheet, wsFwd As Worksheet
    Dim lastRowClasif As Long, lastRowFwd As Long

    Dim StartTime As Double
    Dim SecondsElapsed As Double
    
    StartTime = Timer
    
    Application.Calculation = xlManual
        Application.ScreenUpdating = False
        
    Calculate

    ThisWorkbook.Activate
    Set wsClasif = Sheets("Clasificación")
    Set wsEliminados = Sheets("Eliminados")
    Set wsFwd = Sheets("Fwd Industria")
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    wsEliminados.Range("A2:O10000").ClearContents
    '2. FILTRAR POR VALORES < 1000 EN CLASIFICACIÓN!M1 Y MOVER A ELIMINADOS
    lastRowClasif = wsClasif.Cells(wsClasif.Rows.Count, 13).End(xlUp).Row
    wsClasif.Range("A1:O" & lastRowClasif).AutoFilter Field:=13, Criteria1:="<1000"

    On Error Resume Next
    wsClasif.Range("A2:O" & lastRowClasif).SpecialCells(xlCellTypeVisible).Copy
    wsEliminados.Range("A2").PasteSpecial Paste:=xlPasteValues
    wsClasif.Range("A2:O" & lastRowClasif).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    On Error GoTo 0

    wsClasif.AutoFilterMode = False

    '3. FILTRAR POR TEXTO "alternativo" EN FWD INDUSTRIA!AS22 Y ELIMINAR
    lastRowFwd = wsFwd.Cells(wsFwd.Rows.Count, 45).End(xlUp).Row
    wsFwd.Range("AK22:AU" & lastRowFwd).AutoFilter Field:=9, Criteria1:="alternativo"

    On Error Resume Next
    wsFwd.Range("A23:AU" & lastRowFwd).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    On Error GoTo 0

    wsFwd.AutoFilterMode = False

    Application.CutCopyMode = True
    
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
     
    SecondsElapsed = Round(Timer - StartTime, 2)
    Sheets("Parámetros").Activate
    Sheets("Parámetros").Range("L7") = SecondsElapsed
    
    Sheets("Parámetros").Select
  
End Sub
'Esta macro simplemente nos deja los fwds en el formato del detallado despues de haber arreglado paridades
Sub ReorganizarColumnasFwdIndustria()
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = Sheets("Fwd Industria")
    ws.Range("A23:AH100000").ClearContents
    lastRow = ws.Cells(ws.Rows.Count, "AK").End(xlUp).Row
    If lastRow < 23 Then Exit Sub

    ws.Range("A23:A" & lastRow).Value = ws.Range("AK23:AK" & lastRow).Value
    ws.Range("B23:B" & lastRow).Value = ws.Range("AY23:AY" & lastRow).Value
    ws.Range("C23:C" & lastRow).Value = ws.Range("AN23:AN" & lastRow).Value
    ws.Range("D23:D" & lastRow).Value = ws.Range("AQ23:AQ" & lastRow).Value
    ws.Range("E23:E" & lastRow).Value = ws.Range("AX23:AX" & lastRow).Value
    ws.Range("F23:F" & lastRow).Value = ws.Range("AS23:AS" & lastRow).Value
    ws.Range("H23:H" & lastRow).Value = ws.Range("AT23:AT" & lastRow).Value
    ws.Range("AG23:AG" & lastRow).Value = ws.Range("AU23:AU" & lastRow).Value
    ws.Range("AH23:AH" & lastRow).Value = ws.Range("AZ23:AZ" & lastRow).Value
End Sub

Sub Importar_indicadores()

Dim R_IND As String
Dim FechaActual As String

Application.ScreenUpdating = False

R_IND_1 = Range("R_IND")

Worksheets("VECTOR").Select
Range("U:AI").ClearContents

Archivo = R_IND_1 & ".CSV"

    With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + Archivo, Destination:=Range("$U$1"))
       
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

End Sub
'VALIRDAR COLUMNAS
Sub Columnas()

    Dim wbOrigen As Workbook
    Dim wbDestino As Workbook
    Dim rutaArchivo As String
    Dim hojaParametros As Worksheet
    Dim hojaDestino As Worksheet
    Dim hoja351 As Worksheet
    Dim hoja415 As Worksheet
    Dim ultimaColumna As Long
    Dim encabezados As Variant
    Dim i As Long

    Set wbDestino = ThisWorkbook
    Set hojaParametros = wbDestino.Sheets("Parámetros")
    Set hojaDestino = wbDestino.Sheets("columnas")

    rutaArchivo = hojaParametros.Range("C17").Value

    Set wbOrigen = Workbooks.Open(rutaArchivo)

    Set hoja351 = wbOrigen.Sheets("Fmto-351")
    ultimaColumna = hoja351.Cells(1, hoja351.Columns.Count).End(xlToLeft).Column
    encabezados = hoja351.Range(hoja351.Cells(1, 1), hoja351.Cells(1, ultimaColumna)).Value

    For i = 1 To ultimaColumna
        hojaDestino.Cells(2, i).Value = encabezados(1, i)
    Next i
    Set hoja415 = wbOrigen.Sheets("Fmto-415")
    ultimaColumna = hoja415.Cells(1, hoja415.Columns.Count).End(xlToLeft).Column
    encabezados = hoja415.Range(hoja415.Cells(1, 1), hoja415.Cells(1, ultimaColumna)).Value

    For i = 1 To ultimaColumna
        hojaDestino.Cells(5, i).Value = encabezados(1, i)
    Next i
    wbOrigen.Close SaveChanges:=False
End Sub

