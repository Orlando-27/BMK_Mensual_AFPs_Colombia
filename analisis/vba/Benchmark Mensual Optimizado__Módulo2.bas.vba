Attribute VB_Name = "Módulo2"
Sub Importar_activos_industria()
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
        MsgBox "Revisar Controles Activos antes de importar.", vbExclamation
        Exit Sub
    End If

    If Dir(rutaArchivo) = "" Then
        MsgBox "Archivo no encontrado en la ruta especificada.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Set wbOrigen = Workbooks.Open(rutaArchivo)
    Set hojaOrigen = wbOrigen.Sheets("Fmto-351")
    
    With wbDestino.Sheets("Clasificación")
        .Range("A3:AL100000").ClearContents
    End With

    'Empezamos con activos, aqui no hay filtro, procesamos la data dentro del archivo luego
    With hojaOrigen

    .Range("A1:CX100000").AutoFilter Field:=7, Criteria1:="<>FONDO ALTERNATIVO"

        With wbDestino.Sheets("Clasificación")
            hojaOrigen.Range("G2:G" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("A2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("C2:C" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("B2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("AB2:AB" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("C2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("CD2:CD" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("D2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("Z2:Z" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("E2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("R2:R" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("F2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("AH2:AH" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("G2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("AF2:AF" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("H2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("AI2:AI" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("I2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("AK2:AK" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("J2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("AM2:AM" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("K2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("AN2:AN" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("L2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("BB2:BB" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("M2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("BE2:BE" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("N2").PasteSpecial Paste:=xlPasteValues
            hojaOrigen.Range("AR2:AR" & hojaOrigen.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
            .Range("O2").PasteSpecial Paste:=xlPasteValues
        End With
    End With
    
    hojaOrigen.AutoFilterMode = False
    wbOrigen.Close SaveChanges:=False
    Application.ScreenUpdating = True
    
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
     
    SecondsElapsed = Round(Timer - StartTime, 2)
    Sheets("Parámetros").Activate
    Sheets("Parámetros").Range("L4") = SecondsElapsed
    
    Sheets("Parámetros").Select
End Sub

Sub Formulas_Clasificación()

    Dim StartTime As Double
    Dim SecondsElapsed As Double
    
    StartTime = Timer
    
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
        
    Calculate
    Set wbDestino = ThisWorkbook
    With wbDestino.Sheets("Clasificación")
        filaFinal = .Cells(.Rows.Count, "A").End(xlUp).Row
        .Range("P2:AS2").AutoFill Destination:=.Range("P2:AS" & filaFinal)
    End With
    SecondsElapsed = Round(Timer - StartTime, 2)
    Sheets("Parámetros").Activate
    Sheets("Parámetros").Range("L8") = SecondsElapsed
    
    Sheets("Parámetros").Select
    
End Sub

