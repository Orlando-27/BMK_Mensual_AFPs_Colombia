Attribute VB_Name = "Módulo1"

'JOSE MANUEL IBAÑEZ DELGADO
'INGENIERIA FINANCIERA
'CALCULADORA SWAP - 2016 - 1

Sub EXPORTARDATOS()

Application.ScreenUpdating = False

EXPORTAR_IBR
EXPORTAR_IBRUVR
EXPORTAR_DTF
'EXPORTAR_IPC
EXPORTAR_USDOIS
EXPORTAR_LIBORUVR
EXPORTAR_LIBORCOP
EXPORTAR_USDCO
'EXPORTAR_USDIBR
EXPORTAR_MONEDA
EXPORTAR_EURCOP
'EXPORTAR_LIBORUSD3M
EXPORTAR_CCS_COLATERAL
Sheets("INICIO").Select

Calculate
End Sub
Sub EXPORTAR_IBR()

RUTA_IBR = Range("ruta_ibr")
año = Format(Year(Range("Fecha_valoracion")), "0000")
mes = Format(Month(Range("Fecha_valoracion")), "00")
dia = Format(Day(Range("Fecha_valoracion")), "00")


Sheets("DATOS").Select
    Range("d2:e2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

  Range("d2").Select
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & RUTA_IBR & año & mes & dia & ".txt", Destination _
        :=Range("$d$2"))
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

End Sub


Sub EXPORTAR_DTF()

RUTA_DTF = Range("ruta_dtf")
año = Format(Year(Range("Fecha_valoracion")), "0000")
mes = Format(Month(Range("Fecha_valoracion")), "00")
dia = Format(Day(Range("Fecha_valoracion")), "00")


Sheets("DATOS").Select
    Range("j2:k2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

  Range("j2").Select
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & RUTA_DTF & año & mes & dia & ".txt", Destination _
        :=Range("$j$2"))
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

End Sub
Sub EXPORTAR_CCS_COLATERAL()

CCS_COLATERAL = Range("CCS_COLATERAL")
año = Format(Year(Range("Fecha_valoracion")), "0000")
mes = Format(Month(Range("Fecha_valoracion")), "00")
dia = Format(Day(Range("Fecha_valoracion")), "00")


Sheets("DATOS").Select
    Range("AZ2:BA2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

  Range("AZ2").Select
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & CCS_COLATERAL & año & mes & dia & ".txt", Destination _
        :=Range("$AZ$2"))
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

End Sub

Sub EXPORTAR_IBRUVR()

RUTA_IBRUVR = Range("ruta_ibruvr")
año = Format(Year(Range("Fecha_valoracion")), "0000")
mes = Format(Month(Range("Fecha_valoracion")), "00")
dia = Format(Day(Range("Fecha_valoracion")), "00")


Sheets("DATOS").Select
    Range("G2:H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

  Range("G2").Select
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & RUTA_IBRUVR & año & mes & dia & ".txt", Destination _
        :=Range("$G$2"))
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

End Sub

Sub EXPORTAR_USDOIS()

RUTA_USDOIS = Range("ruta_usdois")
año = Format(Year(Range("Fecha_valoracion")), "0000")
mes = Format(Month(Range("Fecha_valoracion")), "00")
dia = Format(Day(Range("Fecha_valoracion")), "00")


Sheets("DATOS").Select
    Range("M2:N2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

  Range("M2").Select
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & RUTA_USDOIS & año & mes & dia & ".txt", Destination _
        :=Range("$M$2"))
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
    
End Sub

Sub EXPORTAR_IPC()

RUTA_IPC = Range("ruta_ipc")
año = Format(Year(Range("Fecha_valoracion")), "0000")
mes = Format(Month(Range("Fecha_valoracion")), "00")
dia = Format(Day(Range("Fecha_valoracion")), "00")


Sheets("DATOS").Select
    Range("P2:Q2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

  Range("P2").Select
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & RUTA_IPC & año & mes & dia & ".txt", Destination _
        :=Range("$P$2"))
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

End Sub
Sub EXPORTAR_LIBORUVR()

RUTA_LIBORUVR = Range("ruta_liboruvr")
año = Format(Year(Range("Fecha_valoracion")), "0000")
mes = Format(Month(Range("Fecha_valoracion")), "00")
dia = Format(Day(Range("Fecha_valoracion")), "00")


Sheets("DATOS").Select
    Range("S2:T2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

  Range("S2").Select
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & RUTA_LIBORUVR & año & mes & dia & ".txt", Destination _
        :=Range("$S$2"))
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

End Sub

Sub EXPORTAR_LIBORCOP()

RUTA_LIBORCOP = Range("ruta_liborcop")
año = Format(Year(Range("Fecha_valoracion")), "0000")
mes = Format(Month(Range("Fecha_valoracion")), "00")
dia = Format(Day(Range("Fecha_valoracion")), "00")


Sheets("DATOS").Select
    Range("V2:W2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

  Range("V2").Select
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & RUTA_LIBORCOP & año & mes & dia & ".txt", Destination _
        :=Range("$V$2"))
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

End Sub

Sub EXPORTAR_USDCO()

RUTA_USDCO = Range("ruta_usdco")
año = Format(Year(Range("Fecha_valoracion")), "0000")
mes = Format(Month(Range("Fecha_valoracion")), "00")
dia = Format(Day(Range("Fecha_valoracion")), "00")


Sheets("DATOS").Select
    Range("y2:z2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

  Range("y2").Select
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & RUTA_USDCO & año & mes & dia & ".txt", Destination _
        :=Range("$y$2"))
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

End Sub

Sub EXPORTAR_USDIBR()

RUTA_USDIBR = Range("ruta_usdibr")
año = Format(Year(Range("Fecha_valoracion")), "0000")
mes = Format(Month(Range("Fecha_valoracion")), "00")
dia = Format(Day(Range("Fecha_valoracion")), "00")


Sheets("DATOS").Select
    Range("ab2:ac2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

  Range("ab2").Select
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & RUTA_USDIBR & año & mes & dia & ".csv", Destination _
        :=Range("$ab$2"))
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

End Sub

Sub EXPORTAR_MONEDA()

RUTA_CARPETAMONEDA = Range("ruta_moneda")
año = Format(Year(Range("Fecha_valoracion")), "0000")
mes = Format(Month(Range("Fecha_valoracion")), "00")
dia = Format(Day(Range("Fecha_valoracion")), "00")

Sheets("MONEDA").Select
Range("A:U").Select
Selection.ClearContents
Ruta_MONEDA = RUTA_CARPETAMONEDA + "\" + año + "-" + mes + "-" + dia + " IND.CSV"
Range("A1").Select


With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + Ruta_MONEDA, Destination:=Range("$A$1"))
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
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With


End Sub


Sub EXPORTAR_LIBORUSD3M()

RUTA_LIBORUSD3M = Range("ruta_liborusd3M")
año = Format(Year(Range("Fecha_valoracion")), "0000")
mes = Format(Month(Range("Fecha_valoracion")), "00")
dia = Format(Day(Range("Fecha_valoracion")), "00")


Sheets("DATOS").Select
    Range("AR2:AS2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

  Range("AR2").Select
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & RUTA_LIBORUSD3M & año & mes & dia & ".txt", Destination _
        :=Range("$AR$2"))
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
        .TextFileColumnDataTypes = Array(1, 1, 9, 9)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With

End Sub

Sub EXPORTAR_EURCOP()

RUTA_EURCOP = Range("ruta_eurcop")
año = Format(Year(Range("Fecha_valoracion")), "0000")
mes = Format(Month(Range("Fecha_valoracion")), "00")
dia = Format(Day(Range("Fecha_valoracion")), "00")

LIBRO = ActiveWorkbook.Name
Sheets("DATOS").Select
    Range("ae2:af2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

  Range("ae2").Select

'IMPORTA EL ARCHIVO DE SwapCC_EURCOP_Diaria
'
'    With ActiveSheet.QueryTables.Add(Connection:= _
'        "TEXT;" & RUTA_EURCOP & año & mes & dia & ".txt", Destination _
'        :=Range("$ae$2"))
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

'Abre el CSV SwapCC2_EURCOP_Diaria
Workbooks.Open RUTA_EURCOP & año & mes & dia & ".csv"
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

End Sub

