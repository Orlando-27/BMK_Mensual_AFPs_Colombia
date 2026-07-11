Attribute VB_Name = "Módulo1"
'By Jhon Guevara_

'Esta macro inicia actualizando los controles sobre los tipos de derivados, clases de inversión,
'portafolios y AFPs que existen en el formato de la SFC, de tal manera que nos aseguremos de que siga estando en los mismos terminos
'en caso de que no sea así podamos darnos cuenta y ajustemos las cosas nuevas

Sub DiccionarioSFC()
    Dim wbOrigen As Workbook, wbDestino As Workbook
    Dim rutaArchivo As String
    Dim ultimaFila As Long
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    
    StartTime = Timer
    
    Application.Calculation = xlManual
        Application.ScreenUpdating = False
        
        Calculate

    Set ws = Sheets("Clasificación")
    ws.Range("P3:AS100000").ClearContents
    
    Set ws = Sheets("VECTOR")
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    ws.Range("C5:C" & lastRow).ClearContents
    
    Set wbDestino = ThisWorkbook
    rutaArchivo = wbDestino.Sheets("Parámetros").Range("C17").Value

    If Dir(rutaArchivo) = "" Then
        MsgBox "Archivo no encontrado en la ruta especificada.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Set wbOrigen = Workbooks.Open(rutaArchivo)

   
    With wbOrigen.Sheets("Fmto-351")
        ultimaFila = .Cells(.Rows.Count, "C").End(xlUp).Row
        .Range("C2:C" & ultimaFila).Copy
        With wbDestino.Sheets("Diccionario SFC")
            .Range("B9").PasteSpecial xlPasteValues
            .Range("B9:B" & .Cells(.Rows.Count, "B").End(xlUp).Row).RemoveDuplicates Columns:=1, Header:=xlNo
        End With
    End With

    With wbOrigen.Sheets("Fmto-415")
        ultimaFila = .Cells(.Rows.Count, "A").End(xlUp).Row
        .Range("A2:A" & ultimaFila).Copy
        With wbDestino.Sheets("Diccionario SFC")
            .Range("A9").PasteSpecial xlPasteValues
            .Range("A9:A" & .Cells(.Rows.Count, "A").End(xlUp).Row).RemoveDuplicates Columns:=1, Header:=xlNo
        End With
    End With

    With wbOrigen.Sheets("Fmto-351")
        ultimaFila = .Cells(.Rows.Count, "G").End(xlUp).Row
        .Range("G2:G" & ultimaFila).Copy
        With wbDestino.Sheets("Diccionario SFC")
            .Range("G26").PasteSpecial xlPasteValues
            .Range("G26:G" & .Cells(.Rows.Count, "G").End(xlUp).Row).RemoveDuplicates Columns:=1, Header:=xlNo
        End With
    End With

    With wbOrigen.Sheets("Fmto-415")
        ultimaFila = .Cells(.Rows.Count, "F").End(xlUp).Row
        .Range("F2:F" & ultimaFila).Copy
        With wbDestino.Sheets("Diccionario SFC")
            .Range("F26").PasteSpecial xlPasteValues
            .Range("F26:F" & .Cells(.Rows.Count, "F").End(xlUp).Row).RemoveDuplicates Columns:=1, Header:=xlNo
        End With
    End With
    
    With wbOrigen.Sheets("Fmto-351")
        ultimaFila = .Cells(.Rows.Count, "A").End(xlUp).Row
        .Range("Z2:Z" & ultimaFila).Copy
        With wbDestino.Sheets("Diccionario SFC")
            .Range("O2").PasteSpecial xlPasteValues
            .Range("O2:O" & .Cells(.Rows.Count, "O").End(xlUp).Row).RemoveDuplicates Columns:=1, Header:=xlNo
            .Range("P2").AutoFill Destination:=.Range("P2:P" & .Cells(.Rows.Count, "O").End(xlUp).Row)
        End With
    End With

    With wbOrigen.Sheets("Fmto-415")
        ultimaFila = .Cells(.Rows.Count, "A").End(xlUp).Row
        .Range("J2:J" & ultimaFila).Copy
        With wbDestino.Sheets("Diccionario SFC")
            .Range("S3").PasteSpecial xlPasteValues
            .Range("S3:S" & ultimaFila).RemoveDuplicates Columns:=1, Header:=xlNo
   
            .Range("T3").AutoFill Destination:=.Range("T3:T" & .Cells(.Rows.Count, "S").End(xlUp).Row)
        End With
    End With

    wbOrigen.Close SaveChanges:=False
    Application.ScreenUpdating = False
    Calculate
    
    Call Columnas
    Call Importar_derivados_industria
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
     
    SecondsElapsed = Round(Timer - StartTime, 2)
    Sheets("Parámetros").Activate
    Sheets("Parámetros").Range("L3") = SecondsElapsed
    
    Sheets("Parámetros").Select
End Sub

'Esta macro valida si los controles anteriores son correctos, si es así empíeza importandos todos los
'derivados nuevos tanto de la industria como de colfondos como industria, dado el caso de que aparezca un nuevo derivado, este se puede agregar
'la parte de abajo copiando el mismo bloque de cada hoja escrita en verde, solo hay que cambiar el nombre de la hoja destino y el filtro en el origen (J)

Sub Importar_derivados_industria()
    Dim wbOrigen As Workbook, wbDestino As Workbook
    Dim rutaArchivo As String
    Dim hojaOrigen As Worksheet
    Dim celdaControl As Range
    Dim filaFinalFwd As Long

    Set wbDestino = ThisWorkbook
    rutaArchivo = wbDestino.Sheets("Parámetros").Range("C17").Value
    Set celdaControl = wbDestino.Sheets("Parámetros").Range("F19")

    If celdaControl.Value <> True Then
        MsgBox "Revisar derivados antes de importar.", vbExclamation
        Exit Sub
    End If

    If Dir(rutaArchivo) = "" Then
        MsgBox "Archivo no encontrado en la ruta especificada.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Set wbOrigen = Workbooks.Open(rutaArchivo)
    Set hojaOrigen = wbOrigen.Sheets("Fmto-415")
    
    With wbDestino.Sheets("FutLoc Industria")
        .Range("A2:B10000,F2:N10000").ClearContents
    End With

    With wbDestino.Sheets("FutInt Industria")
        .Range("E2:F10000,G2:O10000,R2:R10000").ClearContents
    End With

    With wbDestino.Sheets("Fwd Industria")
        .Range("AK23:AK10000,AL23:AL10000,AM23:AM10000,AO23:AO10000,AP23:AP10000,AR23:AR10000,AS23:AS10000,AT23:AT10000,AU23:AU10000,AN24:AN10000,AQ24:AQ10000,AV24:AZ10000").ClearContents
       
    End With

    'Empezamos con FILTRO POR 5 en J = FutLoc Industria
    With hojaOrigen
        .AutoFilterMode = False
        .Range("J1").AutoFilter Field:=10, Criteria1:=5

        With wbDestino.Sheets("FutLoc Industria")
            hojaOrigen.Range("B2:B" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("A2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("Z2:Z" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("B2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("F2:F" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("F2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("AG2:AG" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("I2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("AE2:AE" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("J2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("AB2:AB" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("K2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("AC2:AC" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("L2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("X2:X" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("M2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("A2:A" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("N2").PasteSpecial Paste:=xlPasteValues
        End With
    End With

    'Sigue con FILTRO POR 6 en J = FutInt Industria
    
    With hojaOrigen
        .AutoFilterMode = False
        .Range("J1").AutoFilter Field:=10, Criteria1:=6

        With wbDestino.Sheets("FutInt Industria")
            hojaOrigen.Range("F2:F" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("E2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("AF2:AF" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("F2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("AG2:AG" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("G2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("Z2:Z" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("H2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("P2:P" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("J2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("W2:W" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("K2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("X2:X" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("L2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("Z2:Z" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("M2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("AE2:AE" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("N2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("AD2:AD" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("O2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("A2:A" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("R2").PasteSpecial Paste:=xlPasteValues
        End With
    End With

    'Sigue con FILTRO POR 1 en J Fwd Industria
    
    With hojaOrigen
        .AutoFilterMode = False
        .Range("J1").AutoFilter Field:=10, Criteria1:=1

        With wbDestino.Sheets("Fwd Industria")
            hojaOrigen.Range("W2:W" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("AK23").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("AA2:AA" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("AL23").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("AC2:AC" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("AM23").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("AB2:AB" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("AO23").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("AD2:AD" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("AP23").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("AE2:AE" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("AR23").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("F2:F" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("AS23").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("X2:X" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("AT23").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("A2:A" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("AU23").PasteSpecial Paste:=xlPasteValues
        End With
    End With

    'Sigue con FILTRO POR 4 en J y llevamos a hoja destino "Fwd Industria" al final (Futuros TRM)
    
    With hojaOrigen
        .AutoFilterMode = False
        .Range("J1").AutoFilter Field:=10, Criteria1:=4

        With wbDestino.Sheets("Fwd Industria")
            filaFinalFwd = .Cells(.Rows.Count, "AK").End(xlUp).Row + 1
            hojaOrigen.Range("W2:W" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("AK" & filaFinalFwd).PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("AA2:AA" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("AL" & filaFinalFwd).PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("AC2:AC" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("AM" & filaFinalFwd).PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("AB2:AB" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("AO" & filaFinalFwd).PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("AD2:AD" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("AP" & filaFinalFwd).PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("AE2:AE" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("AR" & filaFinalFwd).PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("F2:F" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("AS" & filaFinalFwd).PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("X2:X" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("AT" & filaFinalFwd).PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("A2:A" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("AU" & filaFinalFwd).PasteSpecial Paste:=xlPasteValues
            
            filaFinalFwd2 = .Cells(.Rows.Count, "AK").End(xlUp).Row
            .Range("AZ" & filaFinalFwd & ":AZ" & filaFinalFwd2) = "TRM"
            
            'Arrastrar fórmulas en AN y AQ (Calculamos si es venta o compra y que nominal se toma según la misma)
            
            filaFinalFwd2 = .Cells(.Rows.Count, "AK").End(xlUp).Row
            If .Range("AN23").HasFormula Then
                .Range("AN23").AutoFill Destination:=.Range("AN23:AN" & filaFinalFwd2)
            End If
            If .Range("AQ23").HasFormula Then
                .Range("AQ23").AutoFill Destination:=.Range("AQ23:AQ" & filaFinalFwd2)
            End If
            If .Range("AV23").HasFormula Then
                .Range("AV23").AutoFill Destination:=.Range("AV23:AV" & filaFinalFwd2)
            End If
            If .Range("AW23").HasFormula Then
                .Range("AW23").AutoFill Destination:=.Range("AW23:AW" & filaFinalFwd2)
            End If
            If .Range("AX23").HasFormula Then
                .Range("AX23").AutoFill Destination:=.Range("AX23:AX" & filaFinalFwd2)
            End If
            If .Range("AY23").HasFormula Then
                .Range("AY23").AutoFill Destination:=.Range("AY23:AY" & filaFinalFwd2)
            End If
        End With
    End With

    hojaOrigen.AutoFilterMode = False
    wbOrigen.Close SaveChanges:=False
    Application.ScreenUpdating = True

End Sub
Sub Importar_Swaps()
 Dim wbOrigen As Workbook, wbDestino As Workbook
    Dim rutaArchivo As String
    Dim hojaOrigen As Worksheet
    Dim celdaControl As Range
    Dim filaFinalFwd As Long
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    
    StartTime = Timer
    
    Application.Calculation = xlManual
        Application.ScreenUpdating = False
        
        Calculate

    Set wbDestino = ThisWorkbook
    rutaArchivo = wbDestino.Sheets("Parámetros").Range("C17").Value
    Set celdaControl = wbDestino.Sheets("Parámetros").Range("F19")

    If celdaControl.Value <> True Then
        MsgBox "Revisar derivados antes de importar.", vbExclamation
        Exit Sub
    End If

    If Dir(rutaArchivo) = "" Then
        MsgBox "Archivo no encontrado en la ruta especificada.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Set wbOrigen = Workbooks.Open(rutaArchivo)
    Set hojaOrigen = wbOrigen.Sheets("Fmto-415")
    
    With wbDestino.Sheets("Swaps")
        .Range("A2:BV10000").ClearContents
    End With
    
    'FILTRO POR 16 y 17 en J = Swaps CCS e IRS
    With hojaOrigen
        .AutoFilterMode = False
        .Range("J1").AutoFilter Field:=10, Criteria1:=16
        filafinalswaps = .Cells(.Rows.Count, "A").End(xlUp).Row
        With wbDestino.Sheets("Swaps")
            hojaOrigen.Range("A2:BV" & filafinalswaps).SpecialCells(xlCellTypeVisible).Copy
            .Range("A2").PasteSpecial Paste:=xlPasteValues
        End With
    End With
    
        With hojaOrigen
        .AutoFilterMode = False
        .Range("J1").AutoFilter Field:=10, Criteria1:=17
        filafinalswaps = .Cells(.Rows.Count, "A").End(xlUp).Row
        With wbDestino.Sheets("Swaps")
            hojaOrigen.Range("A2:BV" & filafinalswaps).SpecialCells(xlCellTypeVisible).Copy
            filafinalswaps2 = .Cells(.Rows.Count, "A").End(xlUp).Row
            .Range("A" & filafinalswaps2).PasteSpecial Paste:=xlPasteValues
        End With
    End With
    With wbDestino.Sheets("Swaps")
    .Range("BX3:CU10000").ClearContents
    filaFinalswaps3 = .Cells(.Rows.Count, "A").End(xlUp).Row
      If .Range("BX2:CU2").HasFormula Then
          .Range("BX2:CU2").AutoFill Destination:=.Range("BX2:CU" & filaFinalswaps3)
      End If
    End With
    
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
     
    SecondsElapsed = Round(Timer - StartTime, 2)
    ThisWorkbook.Sheets("Parámetros").Activate
    Sheets("Parámetros").Range("L9") = SecondsElapsed
    ThisWorkbook.Sheets("Parámetros").Select
    Call Importar_CSA_BMK
End Sub

'Esta macro arregla detalles de todos los derivados importados, dejandolos listos para llevar al detallado, no se hace en el archivo fuente
'origen porque hay muchos archivos que usan ese archivo y estan parametrizados como venía la fuente anteriormente, igualmente añado,
'que la corrección y ajustes realizados son eficientes y permiten agregar facilmente en la hoja parámetros nuevos datos o cambios

Sub ReemplazarDiccionarioSFC()
    Dim wsDic As Worksheet
    Dim hojasDestino As Variant
    Dim hoja As Worksheet
    Dim celda As Range
    Dim fila As Long
    Dim buscar As String, reemplazo As String
    Dim rangoBusqueda As Range

    Dim StartTime As Double
    Dim SecondsElapsed As Double
    
    StartTime = Timer
    
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Set wsDic = ThisWorkbook.Sheets("Diccionario SFC")
    hojasDestino = Array("FutLoc Industria", "FutInt Industria", "Fwd Industria", "Clasificación")

    ' Reemplazar nombre de AFPS
    For fila = 9 To wsDic.Cells(wsDic.Rows.Count, "A").End(xlUp).Row
        buscar = wsDic.Cells(fila, "A").Value
        reemplazo = wsDic.Cells(fila, "C").Value

        If buscar <> "" And reemplazo <> "" Then
            For Each hoja In ThisWorkbook.Worksheets
                If Not IsError(Application.Match(hoja.Name, hojasDestino, 0)) Then
                    If hoja.Name = "Clasificación" Then
                        Set rangoBusqueda = hoja.Range("A1:B" & hoja.Cells(hoja.Rows.Count, "A").End(xlUp).Row)
                    Else
                        Set rangoBusqueda = hoja.UsedRange
                    End If

                    For Each celda In rangoBusqueda
                        If Not IsEmpty(celda.Value) And Not celda.HasFormula Then
                            If celda.Value = buscar Then celda.Value = reemplazo
                        End If
                    Next celda
                End If
            Next hoja
        End If
    Next fila

    ' Reemplazar nombre de PORTAFOLIOS
    For fila = 4 To wsDic.Cells(wsDic.Rows.Count, "F").End(xlUp).Row
        buscar = wsDic.Cells(fila, "F").Value
        reemplazo = wsDic.Cells(fila, "E").Value

        If buscar <> "" And reemplazo <> "" Then
            For Each hoja In ThisWorkbook.Worksheets
                If Not IsError(Application.Match(hoja.Name, hojasDestino, 0)) Then
                    If hoja.Name = "Clasificación" Then
                        Set rangoBusqueda = hoja.Range("A1:B" & hoja.Cells(hoja.Rows.Count, "A").End(xlUp).Row)
                    Else
                        Set rangoBusqueda = hoja.UsedRange
                    End If

                    For Each celda In rangoBusqueda
                        If Not IsEmpty(celda.Value) And Not celda.HasFormula Then
                            If celda.Value = buscar Then celda.Value = reemplazo
                        End If
                    Next celda
                End If
            Next hoja
        End If
    Next fila

    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    SecondsElapsed = Round(Timer - StartTime, 2)
    Sheets("Parámetros").Activate
    Sheets("Parámetros").Range("L5") = SecondsElapsed
    Sheets("Parámetros").Select
End Sub



