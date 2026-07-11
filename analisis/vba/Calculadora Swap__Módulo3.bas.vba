Attribute VB_Name = "Módulo3"
'JOSE MANUEL IBAÑEZ DELGADO
'INGENIERIA FINANCIERA
'CALCULADORA SWAP - 2016 - 1

Sub CALCULARVPN()

Application.Calculation = xlManual
Application.ScreenUpdating = False

Sheets("SWAPIND").Activate

Dim StartTime As Double
Dim SecondsElapsed As Double
 
StartTime = Timer
    
Dim UFV As Variant
    UFV = Worksheets("SWAPIND").Cells(Rows.Count, 1).End(xlUp).Row

For i = 1 To UFV - 3

    Sheets("SWAPIND").Activate
    Range("a3").Select
    Range("b1") = ActiveCell.Offset(i, 0).Value
    
    Range("c1").Calculate
    
    Dim hoja As Variant
        hoja = Range("c1").Value
    
    ActiveCell.Offset(i, 1).Select
    ActiveCell.Copy
    Sheets(hoja).Select
    Range("B4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    ActiveSheet.Range("A1:AO150").Calculate
    
    'Pega VPN derecho SWAPIND
    Sheets(hoja).Select
    Range("b15").Copy
    Sheets("SWAPIND").Select
    Range("a3").Select
    ActiveCell.Offset(i, 13).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    'Pega duracion modif derecho
    Sheets(hoja).Select
    Range("b16").Copy
    Sheets("SWAPIND").Select
    Range("a3").Select
    ActiveCell.Offset(i, 45).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    'Pega VPN obligacion SWAPIND
    Sheets(hoja).Select
    Range("e15").Copy
    Sheets("SWAPIND").Select
    Range("a3").Select
    ActiveCell.Offset(i, 21).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    'Pega duracion modif obligacion
    Sheets(hoja).Select
    Range("e16").Copy
    Sheets("SWAPIND").Select
    Range("a3").Select
    ActiveCell.Offset(i, 46).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
     
    Next i

SecondsElapsed = Round(Timer - StartTime, 2)
Sheets("RUTAS").Activate
Sheets("RUTAS").Range("i24") = SecondsElapsed

Sheets("SWAPIND").Select

End Sub































'Range("A1048576").Select
'    Selection.End(xlUp).Select
'    Ultima_Fila = Selection.Row
'
'
'Range("A3").Select
'
'For FILA = 3 To Ultima_Fila
'
'Select Case Selection.Value
'
''swap CCS COP UVR
'
'Case Is = "CCS COPUVR"
'
'ActiveCell.Offset(0, 1).Select
'ActiveCell.Copy
'Sheets("SWAPCC_COPUVR").Select
'Range("B4").Select
'ActiveCell.Select
'                   Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                    :=False, Transpose:=False
'
'                            Calculate
'                            Sheets("SWAPCC_COPUVR").Select
'                            Range("b15").Select
'                            ActiveCell.Copy
'                            Sheets("SWAPIND").Select
'                             ActiveCell.Offset(0, 12).Select
'                             ActiveCell.Select
'                                               Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                                                :=False, Transpose:=False
'
'
'                            Sheets("SWAPCC_COPUVR").Select
'                            Range("e15").Select
'                            ActiveCell.Copy
'                            Sheets("SWAPIND").Select
'                             ActiveCell.Offset(0, 8).Select
'                             ActiveCell.Select
'                                               Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                                                :=False, Transpose:=False
'
'
'
'                            'vuelve a ubicacion hoja swapind
'                            Sheets("SWAPIND").Select
'                             ActiveCell.Offset(0, -21).Select
'
'Case Is = "CCS COPUSD"
'ActiveCell.Offset(0, 1).Select
'ActiveCell.Copy
'Sheets("SWAPCC_COPUSD").Select
'Range("B4").Select
'ActiveCell.Select
'                   Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                    :=False, Transpose:=False
'
'                            Calculate
'                            Sheets("SWAPCC_COPUSD").Select
'                            Range("b15").Select
'                            ActiveCell.Copy
'                            Sheets("SWAPIND").Select
'                             ActiveCell.Offset(0, 12).Select
'                             ActiveCell.Select
'                                               Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                                                :=False, Transpose:=False
'
'                            Sheets("SWAPCC_COPUSD").Select
'                            Range("e15").Select
'                            ActiveCell.Copy
'                            Sheets("SWAPIND").Select
'                             ActiveCell.Offset(0, 8).Select
'                             ActiveCell.Select
'                                               Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                                                :=False, Transpose:=False
'
'
'
'                            'vuelve a ubicacion hoja swapind
'                            Sheets("SWAPIND").Select
'                             ActiveCell.Offset(0, -21).Select
'
'
'Case Is = "IRS IBRCOP"
'ActiveCell.Offset(0, 1).Select
'ActiveCell.Copy
'Sheets("SWAPIRS_IBRoncop").Select
'Range("B4").Select
'ActiveCell.Select
'                   Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                    :=False, Transpose:=False
'
'                            Calculate
'                            Sheets("SWAPIRS_IBRoncop").Select
'                            Range("b15").Select
'                            ActiveCell.Copy
'                            Sheets("SWAPIND").Select
'                             ActiveCell.Offset(0, 12).Select
'                             ActiveCell.Select
'                                               Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                                                :=False, Transpose:=False
'
'                            Sheets("SWAPIRS_IBRoncop").Select
'                            Range("e15").Select
'                            ActiveCell.Copy
'                            Sheets("SWAPIND").Select
'                             ActiveCell.Offset(0, 8).Select
'                             ActiveCell.Select
'                                               Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                                                :=False, Transpose:=False
'
'
'                            'vuelve a ubicacion hoja swapind
'                            Sheets("SWAPIND").Select
'                             ActiveCell.Offset(0, -21).Select
'
'Case Is = "IRS LIBORUSD"
'ActiveCell.Offset(0, 1).Select
'ActiveCell.Copy
'Sheets("SWAPIRS_LIBOR").Select
'Range("B4").Select
'ActiveCell.Select
'                   Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                    :=False, Transpose:=False
'
'                            Calculate
'                            Sheets("SWAPIRS_LIBOR").Select
'                            Range("b15").Select
'                            ActiveCell.Copy
'                            Sheets("SWAPIND").Select
'                             ActiveCell.Offset(0, 12).Select
'                             ActiveCell.Select
'                                               Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                                                :=False, Transpose:=False
'
'                            Sheets("SWAPIRS_LIBOR").Select
'                            Range("e15").Select
'                            ActiveCell.Copy
'                            Sheets("SWAPIND").Select
'                             ActiveCell.Offset(0, 8).Select
'                             ActiveCell.Select
'                                               Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                                                :=False, Transpose:=False
'
'
'                            'vuelve a ubicacion hoja swapind
'                            Sheets("SWAPIND").Select
'                             ActiveCell.Offset(0, -21).Select
'
'
'Case Is = "IRS IBR+PB"
'ActiveCell.Offset(0, 1).Select
'ActiveCell.Copy
'Sheets("SWAPIRS_IBRME").Select
'Range("B4").Select
'ActiveCell.Select
'                   Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                    :=False, Transpose:=False
'
'                            Calculate
'                            Sheets("SWAPIRS_IBRME").Select
'                            Range("b15").Select
'                            ActiveCell.Copy
'                            Sheets("SWAPIND").Select
'                             ActiveCell.Offset(0, 12).Select
'                             ActiveCell.Select
'                                               Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                                                :=False, Transpose:=False
'
'                            Sheets("SWAPIRS_IBRME").Select
'                            Range("e15").Select
'                            ActiveCell.Copy
'                            Sheets("SWAPIND").Select
'                             ActiveCell.Offset(0, 8).Select
'                             ActiveCell.Select
'                                               Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                                                :=False, Transpose:=False
'
'
'                            'vuelve a ubicacion hoja swapind
'                            Sheets("SWAPIND").Select
'                             ActiveCell.Offset(0, -21).Select
'
'Case Is = "CCS COPEUR"
'ActiveCell.Offset(0, 1).Select
'ActiveCell.Copy
'Sheets("SWAPIRS_IBRME").Select
'Range("B4").Select
'ActiveCell.Select
'                   Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                    :=False, Transpose:=False
'
'                            Calculate
'                            Sheets("SWAPIRS_IBRME").Select
'                            Range("b15").Select
'                            ActiveCell.Copy
'                            Sheets("SWAPIND").Select
'                             ActiveCell.Offset(0, 12).Select
'                             ActiveCell.Select
'                                               Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                                                :=False, Transpose:=False
'
'                            Sheets("SWAPIRS_IBRME").Select
'                            Range("e15").Select
'                            ActiveCell.Copy
'                            Sheets("SWAPIND").Select
'                             ActiveCell.Offset(0, 8).Select
'                             ActiveCell.Select
'                                               Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                                                :=False, Transpose:=False
'
'
'                            'vuelve a ubicacion hoja swapind
'                            Sheets("SWAPIND").Select
'                             ActiveCell.Offset(0, -21).Select
'
'Case Is = "CCS UVREUR"
'
'ActiveCell.Offset(0, 1).Select
'ActiveCell.Copy
'Sheets("SWAPCC_COPUVR").Select
'Range("B4").Select
'ActiveCell.Select
'                   Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                    :=False, Transpose:=False
'
'                            Calculate
'                            Sheets("SWAPCC_COPUVR").Select
'                            Range("b15").Select
'                            ActiveCell.Copy
'                            Sheets("SWAPIND").Select
'                             ActiveCell.Offset(0, 12).Select
'                             ActiveCell.Select
'                                               Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                                                :=False, Transpose:=False
'
'
'                            Sheets("SWAPCC_COPUVR").Select
'                            Range("e15").Select
'                            ActiveCell.Copy
'                            Sheets("SWAPIND").Select
'                             ActiveCell.Offset(0, 8).Select
'                             ActiveCell.Select
'                                               Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                                                :=False, Transpose:=False
'
'
'
'                            'vuelve a ubicacion hoja swapind
'                            Sheets("SWAPIND").Select
'                             ActiveCell.Offset(0, -21).Select
'
'Case Is = "CCS IBRLIB"
'
'ActiveCell.Offset(0, 1).Select
'ActiveCell.Copy
'Sheets("SWAPCC_IBRLIBOR").Select
'Range("B4").Select
'ActiveCell.Select
'                   Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                    :=False, Transpose:=False
'
'                            Calculate
'                            Sheets("SWAPCC_IBRLIBOR").Select
'                            Range("b15").Select
'                            ActiveCell.Copy
'                            Sheets("SWAPIND").Select
'                             ActiveCell.Offset(0, 12).Select
'                             ActiveCell.Select
'                                               Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                                                :=False, Transpose:=False
'
'
'                            Sheets("SWAPCC_IBRLIBOR").Select
'                            Range("e15").Select
'                            ActiveCell.Copy
'                            Sheets("SWAPIND").Select
'                             ActiveCell.Offset(0, 8).Select
'                             ActiveCell.Select
'                                               Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                                                :=False, Transpose:=False
'
'
'
'                            'vuelve a ubicacion hoja swapind
'                            Sheets("SWAPIND").Select
'                             ActiveCell.Offset(0, -21).Select
'
'
'
'
'Case Is = "CCS UVRUSD"
'
'ActiveCell.Offset(0, 1).Select
'ActiveCell.Copy
'Sheets("SWAPCC_COPUVR").Select
'Range("B4").Select
'ActiveCell.Select
'                   Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                    :=False, Transpose:=False
'
'                            Calculate
'                            Sheets("SWAPCC_COPUVR").Select
'                            Range("b15").Select
'                            ActiveCell.Copy
'                            Sheets("SWAPIND").Select
'                             ActiveCell.Offset(0, 12).Select
'                             ActiveCell.Select
'                                               Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                                                :=False, Transpose:=False
'
'
'                            Sheets("SWAPCC_COPUVR").Select
'                            Range("e15").Select
'                            ActiveCell.Copy
'                            Sheets("SWAPIND").Select
'                             ActiveCell.Offset(0, 8).Select
'                             ActiveCell.Select
'                                               Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                                                :=False, Transpose:=False
'
'Case Is = "IRS SOFUSD"
'
'ActiveCell.Offset(0, 1).Select
'ActiveCell.Copy
'Sheets("SWAPIRS_SOF").Select
'Range("B4").Select
'ActiveCell.Select
'                   Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                    :=False, Transpose:=False
'
'                            Calculate
'                            Sheets("SWAPIRS_SOF").Select
'                            Range("b15").Select
'                            ActiveCell.Copy
'                            Sheets("SWAPIND").Select
'                             ActiveCell.Offset(0, 12).Select
'                             ActiveCell.Select
'                                               Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                                                :=False, Transpose:=False
'
'
'                            Sheets("SWAPIRS_SOF").Select
'                            Range("e15").Select
'                            ActiveCell.Copy
'                            Sheets("SWAPIND").Select
'                             ActiveCell.Offset(0, 8).Select
'                             ActiveCell.Select
'                                               Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                                                :=False, Transpose:=False
'
'
'
'                            'vuelve a ubicacion hoja swapind
'                            Sheets("SWAPIND").Select
'                             ActiveCell.Offset(0, -21).Select
'
'
'End Select
'
'
'
'
'
'Sheets("SWAPIND").Select
' ActiveCell.Offset(1, 0).Select
'
'
'Next
'
'
'
'
'
'
'
'
'End Sub








