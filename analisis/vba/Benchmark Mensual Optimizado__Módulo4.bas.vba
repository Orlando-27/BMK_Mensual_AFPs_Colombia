Attribute VB_Name = "Módulo4"
Sub Importar()
    Dim progreso As frmProgreso
    Set progreso = New frmProgreso
    progreso.Show vbModeless
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    
    StartTime = Timer
    
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
        
    Calculate
    progreso.ActualizarProgreso 1, 7, "Paso 1 de 7: Controles de Importación + Derivados..."
    DoEvents
    DiccionarioSFC

    progreso.ActualizarProgreso 2, 7, "Paso 2 de 7: Importando Activos + Formulas..."
    DoEvents
    Importar_activos_industria

    progreso.ActualizarProgreso 3, 7, "Paso 3 de 7: Nombres AFPs y Portafolios..."
    DoEvents
    ReemplazarDiccionarioSFC
    
    progreso.ActualizarProgreso 4, 7, "Paso 4 de 7: Vector Completo..."
    DoEvents
    V_PRECIOS_IMPORTAR_COMPLETO
    
    progreso.ActualizarProgreso 5, 7, "Paso 5 de 7: Quitando F Alternativo y VM <1000..."
    DoEvents
    F_Alternativo_Mercado_1000
    
    progreso.ActualizarProgreso 6, 7, "Paso 6 de 7: Aplicando Formulas de Clasificación..."
    DoEvents
    Formulas_Clasificación

    progreso.ActualizarProgreso 7, 7, "Paso 7 de 7: Importando Swaps y CSA..."
    DoEvents
    Importar_Swaps
    
    Unload progreso
    
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
     
    SecondsElapsed = Round(Timer - StartTime, 2)
    ThisWorkbook.Sheets("Parámetros").Activate
    Sheets("Parámetros").Range("L19") = SecondsElapsed
End Sub

Sub Actualizar_detallado()
    Dim wbOrigen As Workbook, wbDestino As Workbook
    Dim rutaArchivo As String
    Dim hojaFwd As Worksheet, hojaFutInt As Worksheet, hojaFutLoc As Worksheet, hojaBenchmark As Worksheet
    Dim hojaOrigenFwd As Worksheet, hojaOrigenFutInt As Worksheet, hojaOrigenFutLoc As Worksheet, hojaClasif As Worksheet
    Dim ultimaFilaOrigen As Long, ultimaFilaDestino As Long, ultimaFilaNueva As Long
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    
    StartTime = Timer
    
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    
    Set wbOrigen = ThisWorkbook
    rutaArchivo = wbOrigen.Sheets("Parámetros").Range("C16").Value
    
    Set celdaControlexportación = ThisWorkbook.Sheets("Parámetros").Range("I19")

    If celdaControlexportación.Value <> True Then
        MsgBox "Revisar controles antes de exportar.", vbExclamation
        Exit Sub
    End If
    
    If rutaArchivo = "" Then
        MsgBox "La celda C16 está vacía. Ingresa la ruta del archivo.", vbExclamation
        Exit Sub
    End If
    If Dir(rutaArchivo) = "" Then
        MsgBox "El archivo no existe: " & rutaArchivo, vbCritical
        Exit Sub
    End If
    
    Set wbDestino = Workbooks.Open(rutaArchivo)
    
    Set hojaFwd = wbDestino.Sheets("Fwd Industria")
    Set hojaFutInt = wbDestino.Sheets("FutInt Industria")
    Set hojaFutLoc = wbDestino.Sheets("FutLoc Industria")
    Set hojaBenchmark = wbDestino.Sheets("Benchmark")
    
    Set hojaOrigenFwd = wbOrigen.Sheets("Fwd Industria")
    Set hojaOrigenFutInt = wbOrigen.Sheets("FutInt Industria")
    Set hojaOrigenFutLoc = wbOrigen.Sheets("FutLoc Industria")
    Set hojaClasif = wbOrigen.Sheets("Clasificación")
    Set hojaCSA = wbOrigen.Sheets("CSA BMK")
    
    ' FWDS
    ultimaFilaOrigen = hojaOrigenFwd.Cells(Rows.Count, "A").End(xlUp).Row
    ultimaFilaDestino = hojaFwd.Cells(Rows.Count, "A").End(xlUp).Row
    
    hojaFwd.Range("A23:F" & ultimaFilaDestino).ClearContents
    hojaFwd.Range("H23:H" & ultimaFilaDestino).ClearContents
    hojaFwd.Range("AG23:AG" & ultimaFilaDestino).ClearContents
    hojaFwd.Range("AH23:AH" & ultimaFilaDestino).ClearContents
    hojaFwd.Range("G24:G" & ultimaFilaDestino).ClearContents
    hojaFwd.Range("I24:AF" & ultimaFilaDestino).ClearContents
    
    'FILTRO
    hojaOrigenFwd.Rows(22).AutoFilter
    hojaOrigenFwd.Range("A22:AH" & ultimaFilaOrigen).AutoFilter Field:=33, _
        Criteria1:=Array("PROTECCIÓN", "PORVENIR", "OLD MUTUAL"), Operator:=xlFilterValues
    
    ' A:F
    hojaOrigenFwd.Range("A23:F" & ultimaFilaOrigen).SpecialCells(xlCellTypeVisible).Copy
    hojaFwd.Range("A23").PasteSpecial xlPasteValues
    
    ' H
    hojaOrigenFwd.Range("H23:H" & ultimaFilaOrigen).SpecialCells(xlCellTypeVisible).Copy
    hojaFwd.Range("H23").PasteSpecial xlPasteValues
    
    ' AG
    hojaOrigenFwd.Range("AG23:AG" & ultimaFilaOrigen).SpecialCells(xlCellTypeVisible).Copy
    hojaFwd.Range("AG23").PasteSpecial xlPasteValues
    
    ' AH
    hojaOrigenFwd.Range("AH23:AH" & ultimaFilaOrigen).SpecialCells(xlCellTypeVisible).Copy
    hojaFwd.Range("AH23").PasteSpecial xlPasteValues
    
    'FORMULAS FWDS EN DETALLADO PILAS
    ultimaFilaNueva = hojaFwd.Cells(Rows.Count, "A").End(xlUp).Row
    
    hojaFwd.Range("G23").AutoFill Destination:=hojaFwd.Range("G23:G" & ultimaFilaNueva)
    hojaFwd.Range("I23:AF23").AutoFill Destination:=hojaFwd.Range("I23:AF" & ultimaFilaNueva)
    
    'FORMATO FWDS
    hojaFwd.Range("A23:AH23").Copy
    hojaFwd.Range("A24:AH" & ultimaFilaNueva).PasteSpecial Paste:=xlPasteFormats
    
    ' Quitar filtro
    hojaOrigenFwd.AutoFilterMode = False
    
    'FUTUROS INT
    ultimaFilaOrigen = hojaOrigenFutInt.Cells(Rows.Count, "E").End(xlUp).Row
    ultimaFilaDestino = hojaFutInt.Cells(Rows.Count, "E").End(xlUp).Row
    
    hojaFutInt.Range("A2:R" & ultimaFilaDestino).ClearContents
    hojaFutInt.Range("s3:Ak1000").ClearContents
    hojaOrigenFutInt.Range("E2:R" & ultimaFilaOrigen).Copy
    hojaFutInt.Range("E2").PasteSpecial xlPasteValues
    
    hojaFutInt.Range("S2:AK2").AutoFill Destination:=hojaFutInt.Range("S2:AK" & ultimaFilaOrigen)
    

    'FUTUROS LOC
    ultimaFilaOrigen = hojaOrigenFutLoc.Cells(Rows.Count, "A").End(xlUp).Row
    ultimaFilaDestino = hojaFutLoc.Cells(Rows.Count, "A").End(xlUp).Row
    
    hojaFutLoc.Range("A2:N" & ultimaFilaDestino).ClearContents
    hojaFutLoc.Range("O3:AF1000").ClearContents
    
    hojaOrigenFutLoc.Rows(1).AutoFilter
    hojaOrigenFutLoc.Range("A1:AF" & ultimaFilaOrigen).AutoFilter Field:=14, _
        Criteria1:=Array("PROTECCIÓN", "PORVENIR", "OLD MUTUAL"), Operator:=xlFilterValues
    
    hojaOrigenFutLoc.Range("A2:N" & ultimaFilaOrigen).SpecialCells(xlCellTypeVisible).Copy
    hojaFutLoc.Range("A2").PasteSpecial xlPasteValues
    
    Dim ultimaFilaNuevaLoc As Long
    ultimaFilaNuevaLoc = hojaFutLoc.Cells(Rows.Count, "A").End(xlUp).Row
    
    hojaFutLoc.Range("O2:AF2").AutoFill Destination:=hojaFutLoc.Range("O2:AF" & ultimaFilaNuevaLoc)
    
    hojaOrigenFutLoc.AutoFilterMode = False

    'HOJA BMK
    hojaBenchmark.Range("A255:V300000").ClearContents
    hojaBenchmark.Range("W256:BS300000").ClearContents
    
    hojaClasif.Rows(1).AutoFilter
    hojaClasif.Range("A1").AutoFilter Field:=17, Criteria1:=Array("PROTECCIÓN", "OLD MUTUAL", "PORVENIR"), Operator:=xlFilterValues
    
    Dim ultimaFilaClasif As Long
    ultimaFilaClasif = hojaClasif.Cells(Rows.Count, "P").End(xlUp).Row
    
    hojaClasif.Range("P2:Q" & ultimaFilaClasif).Copy
    hojaBenchmark.Range("A255").PasteSpecial xlPasteValues
    
    hojaClasif.Range("R2:AH" & ultimaFilaClasif).Copy
    hojaBenchmark.Range("D255").PasteSpecial xlPasteValues
    
    hojaClasif.Range("AJ2:AJ" & ultimaFilaClasif).Copy
    hojaBenchmark.Range("U255").PasteSpecial xlPasteValues
    
    hojaClasif.Range("AO2:AO" & ultimaFilaClasif).Copy
    hojaBenchmark.Range("V255").PasteSpecial xlPasteValues
    
    hojaClasif.AutoFilterMode = False
    
    'CSA
    Dim ultimaFilaBenchmark As Long
    ultimaFilaBenchmark = hojaBenchmark.Cells(Rows.Count, "A").End(xlUp).Row + 1
    hojaCSA.Rows(2).AutoFilter
    hojaCSA.Range("A2").AutoFilter Field:=16, Criteria1:=Array("PROTECCIÓN", "OLD MUTUAL", "PORVENIR"), Operator:=xlFilterValues
    hojaCSA.Range("A2").AutoFilter Field:=25, Criteria1:=Array("<>0")
    Dim ultimaFilaCSA As Long
    ultimaFilaCSA = hojaCSA.Cells(Rows.Count, "P").End(xlUp).Row
    hojaCSA.Range("O3:AJ" & ultimaFilaCSA).Copy
    hojaBenchmark.Range("A" & ultimaFilaBenchmark).PasteSpecial xlPasteValues
    
    hojaCSA.AutoFilterMode = False
    
    ' Arrastrar fórmulas en Benchmark
    Dim ultimaFilaBenchmark1 As Long
    ultimaFilaBenchmark1 = hojaBenchmark.Cells(Rows.Count, "A").End(xlUp).Row
    hojaBenchmark.Range("W255:BS255").AutoFill Destination:=hojaBenchmark.Range("W255:BS" & ultimaFilaBenchmark1)
    
    Application.CutCopyMode = True
    
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
     
    SecondsElapsed = Round(Timer - StartTime, 2)
    ThisWorkbook.Sheets("Parámetros").Activate
    Sheets("Parámetros").Range("L14") = SecondsElapsed
    
End Sub

Sub Importar_CSA_BMK()

    Dim wbOrigen As Workbook
    Dim wbDestino As Workbook
    Dim rutaArchivo As String
    Dim hojaParametros As Worksheet
    Dim hojaDestino As Worksheet
    Dim hojaCSA As Worksheet
    Dim i As Long

    Set wbDestino = ThisWorkbook
    Set hojaParametros = wbDestino.Sheets("Parámetros")
    Set hojaDestino = wbDestino.Sheets("CSA BMK")

    rutaArchivo = hojaParametros.Range("C17").Value

    Set wbOrigen = Workbooks.Open(rutaArchivo)

    Set hojaCSA = wbOrigen.Sheets("Cuentas CSA")
    hojaCSA.Range("A1:G200").Copy
    hojaDestino.Range("A1").PasteSpecial xlPasteValues
    wbDestino.Sheets("Parámetros").Activate
    
End Sub

Sub exportar_swaps_para_julián()
    Dim wbOrigen As Workbook
    Dim wsOrigen As Worksheet
    Dim rngOrigen As Range
    Dim wbNuevo As Workbook
    Dim wsNuevo As Worksheet
    Dim ruta As String
    Dim nombreArchivo As String
    Dim formato As XlFileFormat
    
    Set wbOrigen = ThisWorkbook
    Set wsOrigen = wbOrigen.Worksheets("Swaps")
    Set rngOrigen = wsOrigen.Range("BX1:CU100000")
    
    ruta = "M:\COB\Gerencia Estrategia\Analisis Cuantitativo\GERENCIA DE ANALISIS\ANALISIS FINANCIERO\ANALISIS CUANTITATIVO\Benchmark\Actualización Mensual\"
    nombreArchivo = "Swaps Formato Calculadora" & ".xlsx"
    formato = xlOpenXMLWorkbook
  
    Set wbNuevo = Workbooks.Add(xlWBATWorksheet)
    Set wsNuevo = wbNuevo.Worksheets(1)
    wsNuevo.Name = "Datos"

    rngOrigen.Copy
    wsNuevo.Range("A1").PasteSpecial Paste:=xlPasteValues

    wsNuevo.Columns.AutoFit

    Application.DisplayAlerts = False
    wbNuevo.SaveAs Filename:=ruta & nombreArchivo, FileFormat:=formato
    Application.DisplayAlerts = True
    
    End Sub
