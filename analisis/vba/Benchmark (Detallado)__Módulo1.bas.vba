Attribute VB_Name = "Módulo1"
Sub Benchmark()

Sheets("Procedimiento").Calculate


Application.ScreenUpdating = False
'Borra la informacion de la hoja colfondos
IMPORTACION_CURVAS
        Dim Limit As Integer
        Limit = Range("Limit")
        Sheets("Colfondos").Select
        Range("a" + Format(Limit, 0), Range("bb" + Format(1045876, 0))).ClearContents
        
'Actualizar Hoja Vector de precios de la respectiva macro


Sheets("VECTOR DE PRECIOS").Select
Range("A2:P1048576").ClearContents

     Windows("INICIO MACRO MULTIFONDOS.xlsm").Activate
        Sheets("IND").Select
        Range("A1:D1").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        
      Windows("Benchmark (Detallado).xlsb").Activate
        Sheets("VECTOR DE PRECIOS").Select
        Range("A1").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        Windows("INICIO MACRO MULTIFONDOS.xlsm").Activate
        Sheets("IND").Select
        Range("O3:R3").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        
      Windows("Benchmark (Detallado).xlsb").Activate
        Sheets("VECTOR DE PRECIOS").Select
        Range("G3").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
         Windows("INICIO MACRO MULTIFONDOS.xlsm").Activate
        Sheets("IND2").Select
        Range("A1:E1").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        
      Windows("Benchmark (Detallado).xlsb").Activate
        Sheets("VECTOR DE PRECIOS").Select
        Range("L2").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        Sheets("VECTOR DE PRECIOS").Select
        Dim UFV As Variant
        UFV = Worksheets("VECTOR DE PRECIOS").Cells(Rows.Count, 1).End(xlUp).Row
        Range("E3:E" + Format(UFV, 0)).FormulaR1C1 = "=RC[-4]&RC[-1]"
        Range("E3:E" + Format(UFV, 0)).Select
        Selection.Calculate
      
'Actualizar Hoja Divisas de la respectiva macro
    Sheets("Fut Colf").Select
    Range("A1:Q5000").ClearContents
    Range("R4:BE5000").ClearContents
    
    Windows("divisas.xlsm").Activate
    Sheets("577").Select
    Range("A2:Q5000").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("Benchmark (Detallado).xlsb").Activate
    Sheets("Fut Colf").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Dim LastRow9 As Long
    LastRow = Range("C" & Rows.Count).End(xlUp).Row
    Range("R3:BE3").AutoFill Destination:=Range("R3:BE" & LastRow), Type:=xlFillDefault

    Sheets("Fwd Colfondos").Select
    Range("A22:T1048576").ClearContents
    
    Windows("divisas.xlsm").Activate
    Sheets("DPO").Select
    Range("A5:Q5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("Benchmark (Detallado).xlsb").Activate
    Sheets("Fwd Colfondos").Select
    Range("A22").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Windows("divisas.xlsm").Activate
    Sheets("DPO").Select
    Range("t5:u5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("Benchmark (Detallado).xlsb").Activate
    Sheets("Fwd Colfondos").Select
    Range("r22").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
' QUITA LOS ESPACIO DE LA MONEDA
    Sheets("Fwd Colfondos").Select
    
    Dim CF As Variant
        CF = Worksheets("Fwd Colfondos").Cells(Rows.Count, 19).End(xlUp).Row
     
    Range("T22:T" + Format(CF, 0)).FormulaR1C1 = "=+TRIM(RC[-18])"
    Range("T22:T" + Format(CF, 0)).Calculate
    Range("T22:T" + Format(CF, 0)).Copy
    Range("B22").PasteSpecial xlPasteValues
    Range("T22:T" + Format(CF, 0)).ClearContents
    Range("T22:T" + Format(CF, 0)).FormulaR1C1 = "=IF(ISERROR(VLOOKUP(TEXT(RC[-3],""##""),Parametros!R118C2:R1048576C2,1,FALSE)),""NA"",""V"")"
    
'Actualizar Hoja Opciones

    Sheets("Opt Colfondos").Select
    Range("A1:X10000").ClearContents

    Windows("divisas.xlsm").Activate
    Sheets("572").Activate
    Range("A2").Select
    ActiveCell.End(xlDown).Offset(2, 0).Select
    Range(ActiveCell, ActiveCell.Offset(0, 23)).Select
    Range(Selection, Selection.End(xlDown)).Copy
    
    Windows("Benchmark (Detallado).xlsb").Activate
    Sheets("Opt Colfondos").Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    If Cells(4, 1) <> "Opc No" Then
        Range("A1:X10000").ClearContents
    End If
    
    Range("AI1").Calculate
    Range(Cells(4, 26), Cells(4, 33)).Select
    Selection.Copy
    opcionesN = Cells(1, 35)
    Range(ActiveCell, ActiveCell.Offset(opcionesN - 1, 8)).Select
    ActiveSheet.Paste
    
'Actualizar Hoja de Opciones de la industria

    Sheets("Opt Ind").Select
    Range("Z1:AI10000").ClearContents
    
    R_OpInd = Range("R_OpInd").Value
    N_OpInd = Range("N_OpInd").Value
    OptInd = R_OpInd & N_OpInd
    ArchivoBMK = ActiveWorkbook.Name
    
    Workbooks.Open (OptInd)
    Sheets("Hoja1").Select
    Range("A:J").Copy
    
    Workbooks(ArchivoBMK).Activate
    Range("Z1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    Workbooks(N_OpInd).Close SaveChanges:=False
    
    
'Actualizar Hoja Swaps de la respectiva macro
   
   Sheets("SWAP").Select
    Range("A:u").ClearContents
    
     Windows("divisas.xlsm").Activate
    Sheets("T695").Select
    Range("A2:u2").Select
    Range(Selection, Selection.End(xlDown)).Select
    
    Selection.Copy
    
     Windows("Benchmark (Detallado).xlsb").Activate
    Sheets("SWAP").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
'Pega lo rotulos  de las formulas
    
    Sheets("SWAP").Select

    Range("W1").FormulaR1C1 = "PORTAFOLIO"
    Range("X1").FormulaR1C1 = "TIPO DE SWAP"
    Range("Y1").FormulaR1C1 = "VPN DERECHO"
    Range("Z1").FormulaR1C1 = "VPN OBLIGACIÓN"
    Range("AA1").FormulaR1C1 = "Ubicación derecho"
    Range("AB1").FormulaR1C1 = "Ubicación obligación"
    Range("AC1").FormulaR1C1 = "RIESGO DERECHO"
    Range("AD1").FormulaR1C1 = "RIESGO OBLIGACIÓN"
    Range("AE1").FormulaR1C1 = "Valor tasa facial derecho"
    Range("AF1").FormulaR1C1 = "Valor tasa facial obligación"
    Range("AG1").FormulaR1C1 = "VPN  SWAP TF DERECHO"
    Range("AH1").FormulaR1C1 = "VPN  SWAP TF OBLIGACION"
    Range("AI1").FormulaR1C1 = "Frecuencia Derecho"
    Range("AJ1").FormulaR1C1 = "Frecuencia Obligación"
    Range("AK1").FormulaR1C1 = "Contraparte Derecho"
    Range("AL1").FormulaR1C1 = "Contraparte Obligacion"
    Range("Am1").FormulaR1C1 = "CONTROL"
    Range("AN1").FormulaR1C1 = "CONTROL"
    
    'Pega las formulas hasta el ultimo rango
    
    Range("W3").FormulaR1C1 = "=if(len(trim(rc[-3]))=4,LEFT(RC[-3],2),LEFT(RC[-3],3))"
    'Range("W3").FormulaR1C1 = "=LEFT(RC[-3],2)"
    Range("X3").FormulaR1C1 = "=IF(RC[-19]=RC[-11],""IRS"",""CCS"")"
    Range("Y3").FormulaR1C1 = "=IF(RC[-1]=""IRS"",(RC[-20]*VLOOKUP(RC[-17],'VECTOR DE PRECIOS'!R1C19:R13C20,2,FALSE))/(1+IF(RC[-17]=""COP"",SUMIFS(Curvas!C[-20],Curvas!C[-22],""FWTCOP"",Curvas!C[-21],RC[-23]-Fecha_Valoración),SUMIFS(Curvas!C[-20],Curvas!C[-22],""LIBBTS"",Curvas!C[-21],RC[-23]-Fecha_Valoración))%*((RC[-23]-Fecha_Valoración)/360))+RC[-15],RC[-15])"
    Range("Z3").FormulaR1C1 = "=IF(RC[-2]=""IRS"",-((RC[-21]*VLOOKUP(RC[-18],'VECTOR DE PRECIOS'!R1C19:R13C20,2,FALSE))/(1+IF(RC[-18]=""COP"",SUMIFS(Curvas!C[-21],Curvas!C[-23],""FWTCOP"",Curvas!C[-22],RC[-24]-Fecha_Valoración),SUMIFS(Curvas!C[-21],Curvas!C[-23],""LIBBTS"",Curvas!C[-22],RC[-24]-Fecha_Valoración))%*((RC[-24]-Fecha_Valoración)/360))+RC[-8]),-RC[-8])"
    Range("AA3").FormulaR1C1 = "=IF(OR(RC[-19]=""cop"",RC[-19]=""uvr""),""NACIONAL"",""INTERNACIONAL"")"
    Range("AB3").FormulaR1C1 = "=IF(OR(RC[-12]=""cop"",RC[-12]=""uvr""),""NACIONAL"",""INTERNACIONAL"")"
    Range("AC3").FormulaR1C1 = "=IF(MID(RC[-26],6,2)=""TF"",""TASA FIJA"",IF(MID(RC[-26],6,2)=""TV"",""VARIABLE"",""INFLACIÓN""))"
    Range("AD3").FormulaR1C1 = "=IF(MID(RC[-19],6,2)=""TF"",""TASA FIJA"",IF(MID(RC[-19],6,2)=""TV"",""VARIABLE"",""INFLACIÓN""))"
    Range("AE3").FormulaR1C1 = "=IFERROR(IF(ISNUMBER(MID(RC[-25],1,5)*1),MID(RC[-25],1,5)*1,IF(ISNUMBER(MID(RC[-25],6,4)*1)=FALSE,MID(RC[-25],6,1)*1,IF(ISNUMBER(MID(RC[-25],6,4)*1)=TRUE,MID(RC[-25],6,4),0)))/100,0)"
    Range("AF3").FormulaR1C1 = "=IFERROR(IF(ISNUMBER(MID(RC[-18],1,5)*1),MID(RC[-18],1,5)*1,IF(ISNUMBER(MID(RC[-18],6,4)*1)=FALSE,MID(RC[-18],6,1)*1,IF(ISNUMBER(MID(RC[-18],6,4)*1)=TRUE,MID(RC[-18],6,4),0)))/100,0)"
    Range("AG3").FormulaR1C1 = "=IF(AND(RC[-9]=""IRS"",RC[-30]=""SWAP TF""),RC[-8]+RC[-7],0)"
    Range("AH3").FormulaR1C1 = "=IF(AND(RC[-10]=""IRS"",RC[-23]=""SWAP TF""),RC[-9]+RC[-8],0)"
    Range("AI3").FormulaR1C1 = "=VLOOKUP(RC[-28],Parametros!R1C5:R7C6,2,FALSE)"
    Range("AJ3").FormulaR1C1 = "=VLOOKUP(RC[-21],Parametros!R1C5:R7C6,2,FALSE)"
    Range("AK3").FormulaR1C1 = "=+VLOOKUP(RC4,[divisas.xlsm]DICCIONARIO!C5:C6,2,0)"
    Range("AL3").FormulaR1C1 = "=+VLOOKUP(RC12,[divisas.xlsm]DICCIONARIO!C5:C6,2,0)"
    Range("Am3").FormulaR1C1 = "=VLOOKUP(RC[-35],Colfondos!C[-35]:C[-27],9,FALSE)-RC[-14]"
    Range("An3").FormulaR1C1 = "=VLOOKUP(RC[-28],Colfondos!C[-36]:C[-28],9,FALSE)-RC[-14]"
    
    'Detecta la ultima celda usada y arrastra las formulas
    
    Dim CU0 As Variant
        CU0 = Worksheets("SWAP").Cells(Rows.Count, 1).End(xlUp).Row
        Range("w3:aN3").Copy
        Range("w3:aN" + Format(CU0, 0)).PasteSpecial Paste:=xlPasteFormulas
        
        Sheets("SWAP").Range("A:AN").Calculate
        
    'Detecta la ultima celda usada
    
     Dim CU1 As Variant
        CU1 = Worksheets("Colfondos").Cells(Rows.Count, 1).End(xlUp).Row
            
    '''''''Pega la información de la hoja swap en la hoja colfondos DERECHO
    
    'fondo
    Sheets("SWAP").Select
    Range("w3:w" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("a" + Format(CU1 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
    'Nemo e Isin
    Sheets("SWAP").Select
    Range("c3:c" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("b" + Format(CU1 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    Sheets("Colfondos").Range("c" + Format(CU1 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    Sheets("Colfondos").Range("f" + Format(CU1 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    Sheets("Colfondos").Range("m" + Format(CU1 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    Sheets("Colfondos").Range("y" + Format(CU1 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
    'Titulo del derecho
    Sheets("SWAP").Select
    Range("d3:d" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("d" + Format(CU1 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
    'Contraparte del derecho
    Sheets("SWAP").Select
    Range("ak3:ak" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("e" + Format(CU1 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
    'Titulo del derecho
    '''''Range("u3:u" + Format(CU0, 0)).Copy
    '''''Sheets("Colfondos").Range("e" + Format(CU1 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
     'Titulo del derecho
    Sheets("SWAP").Select
    Range("a3:a" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("g" + Format(CU1 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    Sheets("Colfondos").Range("h" + Format(CU1 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
    Sheets("SWAP").Select
    Range("B3:B" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("i" + Format(CU1 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
    'Moneda del derecho
    Sheets("SWAP").Select
    Range("h3:h" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("j" + Format(CU1 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    Sheets("Colfondos").Range("n" + Format(CU1 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    Sheets("Colfondos").Range("z" + Format(CU1 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
    'Nominal del derecho
    Sheets("SWAP").Select
    Range("e3:e" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("k" + Format(CU1 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
    'Valor Mdo del derecho
    Sheets("SWAP").Select
    Range("y3:y" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("l" + Format(CU1 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
    'Tasa del derecho
    Sheets("SWAP").Select
    Range("f3:f" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("o" + Format(CU1 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
    'Ubicación del derecho
    Sheets("SWAP").Select
    Range("aa3:aa" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("p" + Format(CU1 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    Sheets("Colfondos").Range("x" + Format(CU1 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
    'Clasificación del derecho
    Dim CU4 As Variant
        CU4 = Worksheets("Colfondos").Cells(Rows.Count, 1).End(xlUp).Row
        
     Sheets("Colfondos").Select
     Range("Q" + Format(CU1 + 1, 0), Range("Q" + Format(CU4, 0))).FormulaR1C1 = "SWAP"
     Sheets("SWAP").Select
    'Riesgo del derecho
    Range("ac3:ac" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("r" + Format(CU1 + 1, 0)).PasteSpecial Paste:=xlPasteValues
     
    'Tasa Facial del derecho
    Sheets("SWAP").Select
    Range("x3:x" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("s" + Format(CU1 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    Sheets("Colfondos").Range("aa" + Format(CU1 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
    'Valor Tasa Facial del derecho
    Sheets("SWAP").Select
    Range("ae3:ae" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("t" + Format(CU1 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
    'Tasa de descuento del derecho
    Sheets("SWAP").Select
    Range("ag3:ag" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("u" + Format(CU1 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
    'Tasa de descuento del derecho
    Sheets("SWAP").Select
    Range("ai3:ai" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("v" + Format(CU1 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
    'Clase
    Sheets("Colfondos").Select
    Range("w" + Format(CU1 + 1, 0), Range("W" + Format(CU4 + 1, 0))).FormulaR1C1 = "DERIVADOS"
    Sheets("SWAP").Select
    ''''''Pega la información de la hoja swap en la hoja colfondos DERECHO
    
    'Detecta la ultima celda usada
    
     Dim CU2 As Variant
        CU2 = Worksheets("Colfondos").Cells(Rows.Count, 1).End(xlUp).Row
                
    'fondo
    Sheets("SWAP").Select
    Range("w3:w" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("a" + Format(CU2 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
    'Nemo e Isin
    Sheets("SWAP").Select
    Range("K3:K" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("b" + Format(CU2 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    Sheets("Colfondos").Range("c" + Format(CU2 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    Sheets("Colfondos").Range("f" + Format(CU2 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    Sheets("Colfondos").Range("m" + Format(CU2 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    Sheets("Colfondos").Range("y" + Format(CU2 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
    'Titulo del OBLIGACION
    Sheets("SWAP").Select
    Range("L3:L" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("d" + Format(CU2 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
    'Titulo del OBLIGACION
    Sheets("SWAP").Select
    Range("al3:aL" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("e" + Format(CU2 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
    'EMISOR del OBLIGACION
    'Range("u3:u" + Format(CU0, 0)).Copy
    'Sheets("Colfondos").Range("e" + Format(CU2 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
     'FECHAS del OBLIGACION
    Sheets("SWAP").Select
    Range("a3:a" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("g" + Format(CU2 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    Sheets("Colfondos").Range("h" + Format(CU2 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
    Sheets("SWAP").Select
    Range("B3:B" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("i" + Format(CU2 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
    'Moneda del OBLIGACION
    Sheets("SWAP").Select
    Range("P3:P" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("j" + Format(CU2 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    Sheets("Colfondos").Range("n" + Format(CU2 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    Sheets("Colfondos").Range("z" + Format(CU2 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
    'Nominal del OBLIGACION
    Sheets("SWAP").Select
    Range("M3:M" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("k" + Format(CU2 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
    'Valor Mdo del OBLIGACION
    Sheets("SWAP").Select
    Range("Z3:Z" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("l" + Format(CU2 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
    'Tasa del OBLIGACION
    Sheets("SWAP").Select
    Range("N3:N" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("o" + Format(CU2 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
    'Ubicación del OBLIGACION
    Sheets("SWAP").Select
    Range("AB3:AB" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("p" + Format(CU2 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    Sheets("Colfondos").Range("x" + Format(CU2 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
    'Clasificación del OBLIGACION
     Dim CU5 As Variant
        CU5 = Worksheets("Colfondos").Cells(Rows.Count, 1).End(xlUp).Row
    
     Sheets("Colfondos").Select
     Range("Q" + Format(CU2 + 1, 0), Range("Q" + Format(CU5 + 1, 0))).FormulaR1C1 = "SWAP"
     Sheets("SWAP").Select
    'Riesgo del OBLIGACION
    Range("AD3:AD" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("r" + Format(CU2 + 1, 0)).PasteSpecial Paste:=xlPasteValues
     
    'Tasa Facial del OBLIGACION
    Sheets("SWAP").Select
    Range("x3:x" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("s" + Format(CU2 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    Sheets("Colfondos").Range("aa" + Format(CU2 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
    'Valor Tasa Facial del OBLIGACION
    Sheets("SWAP").Select
    Range("AF3:AF" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("t" + Format(CU2 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
    'Tasa de descuento del OBLIGACION
    Sheets("SWAP").Select
    Range("AH3:AH" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("u" + Format(CU2 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
    'Tasa de descuento del OBLIGACION
    Sheets("SWAP").Select
    Range("AJ3:AJ" + Format(CU0, 0)).Copy
    Sheets("Colfondos").Range("v" + Format(CU2 + 1, 0)).PasteSpecial Paste:=xlPasteValues
    
    'Clase
     Sheets("Colfondos").Select
     Range("w" + Format(CU2 + 1, 0), Range("W" + Format(CU5 + 1, 0))).FormulaR1C1 = "DERIVADOS"
     Sheets("SWAP").Select
     
''''' Pega la informacion de los titulos de la Macro_multifondos
     
     Dim CU3 As Variant
        CU3 = Worksheets("Colfondos").Cells(Rows.Count, 1).End(xlUp).Row
        
        Windows("INICIO MACRO MULTIFONDOS.xlsm").Activate
        Sheets("PORTAFOLIOS").Select
        Range("A2:AA2").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        
      Windows("Benchmark (Detallado).xlsb").Activate
        Sheets("Colfondos").Select
        Range("A" + Format(CU3 + 1, 0)).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                
''''''Pega las formulas desde la participacion hasta el sector

Sheets("Colfondos").Select
Range("AB" + Format(Limit - 1, 0), Range("BO" + Format(Limit - 1, 0))).Copy

Dim CU6 As Variant
        CU6 = Worksheets("Colfondos").Cells(Rows.Count, 1).End(xlUp).Row
        Range("AB" + Format(Limit - 1, 0), Range("AB" + Format(CU6, 0))).PasteSpecial Paste:=xlPasteFormulas


       Calculate
     
       Sheets("Procedimiento").Select
       
       Respuesta = MsgBox("¿Dese actualizar el archivo de Notas Estructuradas", vbYesNo, "Notas Estructuradas")
       
       If Respuesta = vbYes Then
       
                   Dim N_NE, R_NE As String
            
                   N_NE = Range("N_NE")
                   R_NE = Range("R_NE")
            
                   Sheets("NE").Select
                   Range("A1:W5000").ClearContents
            
                   archivo = R_NE + "\" + N_NE + ".CSV"
            
               With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + archivo, Destination:=Range("$A$1"))
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
                    .TextFileSemicolonDelimiter = False
                    .TextFileCommaDelimiter = True
                    .TextFileSpaceDelimiter = False
                    .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
                    .TextFileTrailingMinusNumbers = True
                    .Refresh BackgroundQuery:=False
               
               'Range("A30:U30").ClearContents
               'Range("A31:U31").ClearContents
               
                             
               End With
               
               Sheets("NE").Select
               Dim notascel As Variant
               notascel = Worksheets("NE").Cells(Rows.Count, 1).End(xlUp).Row
               Range("A" + Format(notascel + 3, 0)).Select
               
              Dim N_NE_BMK As Variant
            N_NE_BMK = Range("N_NE_BMK")
                   archivo = R_NE + "\" + N_NE_BMK + ".CSV"
               
               With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + archivo, Destination:=Range("a" + Format(notascel + 2, "00")))
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
                    .TextFileSemicolonDelimiter = False
                    .TextFileCommaDelimiter = True
                    .TextFileSpaceDelimiter = False
                    .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
                    .TextFileTrailingMinusNumbers = True
                    .Refresh BackgroundQuery:=False
               
               'Range("A30:U30").ClearContents
               'Range("A31:U31").ClearContents
               
               
               
               End With


Sheets("VECTOR DE PRECIOS").Select
Range("T1:T14").Calculate
                
'Dim UFV1 As Variant
'UFV1 = Worksheets("VECTOR DE PRECIOS").Cells(Rows.Count, 7).End(xlUp).Row
                    
'Range("S17:V19").Copy
'Range("G" + Format(UFV1 + 1, 0)).Select
'Selection.PasteSpecial xlPasteValues
Sheets("Formato de Apuestas PO-CS").Select
     
   End If
   
        Sheets("Procedimiento").Select
        Range("Fecha_Valoración").Select
        Selection.Copy
        
        Sheets("NE").Select
        Range("C2").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        Range("C2").Select
        Selection.Copy
        
        Range("C2").End(xlDown).Offset(3).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
   
 Call futuros
 Call traer_sensibilidad
 Call ReemplazarPalabraNotas
    Sheets("Procedimiento").Select

    'ActiveWorkbook.Save

End Sub

Sub IMPORTACION_CURVAS()
    
    Application.ScreenUpdating = True
    Worksheets("Procedimiento").Range("A:F").Calculate
    IMP_CURVAS_HOY
    'IMP_CURVAS_ANT
    Importar_indicadores
    Limpiar_CONEXIONES
    Limpiar_Datos_externos
    Sheets("Curvas").Select
    Range("AA:AP").Calculate
    Sheets("Procedimiento").Select
    Application.ScreenUpdating = False
    
End Sub

Sub IMP_CURVAS_HOY()

Dim Ruta_PF As String
Dim N_Fwd_USDCOP_T As String
'Dim N_Fwd_USDCOP_T_1 As String
Dim N_Fwd_USDCAD_T As String
Dim N_Fwd_USDBRL_T As String
Dim N_Fwd_USDMXN_T As String
Dim N_Fwd_USDJPY_T As String
Dim N_Fwd_EURUSD_T As String
Dim N_SWAP_LBRUSD_T As String
Dim N_SWAP_LBRCOP_T As String
Dim N_Fwd_EURCOP_T As String
Dim N_Fwd_EUR2COP_T As String
Dim N_Fwd_GBPUSD_T As String
Dim N_Fwd_USDCLP_T As String
Dim N_Fwd_USDCHF_T As String
Dim N_Fwd_USDKRW_T As String


Fecha_Valoración_ANT = Range("Fecha_Valoración_ANT")

N_Fwd_USDCOP_T = Range("N_Fwd_USDCOP_T")
'N_Fwd_USDCOP_T_1 = Range("N_Fwd_USDCOP_T_1")
N_Fwd_USDCAD_T = Range("N_Fwd_USDCAD_T")
N_Fwd_USDBRL_T = Range("N_Fwd_USDBRL_T")
N_Fwd_USDMXN_T = Range("N_Fwd_USDMXN_T")
N_Fwd_USDJPY_T = Range("N_Fwd_USDJPY_T")
N_Fwd_EURUSD_T = Range("N_Fwd_EURUSD_T")
N_SWAP_LBRUSD_T = Range("N_SWAP_LBRUSD_T")
N_SWAP_LBRCOP_T = Range("N_SWAP_LBRCOP_T")
N_Fwd_EURCOP_T = Range("N_Fwd_EURCOP_T")
N_Fwd_EUR2COP_T = Range("N_Fwd_EUR2COP_T")
N_Fwd_GBPUSD_T = Range("N_Fwd_GBPUSD_T")
N_Fwd_USDCLP_T = Range("N_Fwd_USDCLP_T")
N_Fwd_USDCHF_T = Range("N_Fwd_USDCHF_T")
N_Fwd_USDKRW_T = Range("N_Fwd_USDKRW_T")

Ruta_PF = Range("Ruta_PF")

Sheets("Curvas").Activate
Sheets("Curvas").Range("a:aa").Select
Sheets("Curvas").Range("a:aa").ClearContents

    Sheets("Curvas").Range("a1").FormulaR1C1 = "CURVAS FORWARD ACTUAL"
    Sheets("Curvas").Range("a2").FormulaR1C1 = "Orden"
    Sheets("Curvas").Range("b2").FormulaR1C1 = "Fecha_Operacion"
    Sheets("Curvas").Range("c2").FormulaR1C1 = "Nombre_Indicador"
    Sheets("Curvas").Range("d2").FormulaR1C1 = "Plazo_en_Dias"
    Sheets("Curvas").Range("e2").FormulaR1C1 = "Valor"
    
    Sheets("Curvas").Range("a3").FormulaR1C1 = N_Fwd_USDCOP_T

Ruta_Nombre = Ruta_PF + "\" + N_Fwd_USDCOP_T + ".txt"

With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + Ruta_Nombre, Destination:=Range("$d$4"))
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

Sheets("Curvas").Range("a4").FormulaR1C1 = 1
Sheets("Curvas").Range("b4").FormulaR1C1 = "=Fecha_Valoración_ANT"
Sheets("Curvas").Range("c4").FormulaR1C1 = "FWPCOP"
Sheets("Curvas").Range("a5").FormulaR1C1 = "=R[-1]C+1"
Sheets("Curvas").Range("b5").FormulaR1C1 = "=Fecha_Valoración_ANT"
Sheets("Curvas").Range("c5").FormulaR1C1 = "FWPCOP"

    Range("a5:c5").Select
    Selection.AutoFill Destination:=Range("a5", "c" + Format(Sheets("Curvas").Cells(Rows.Count, 4).End(xlUp).Row, "00"))
    


COUNTULT = Sheets("Curvas").Cells(Rows.Count, 1).End(xlUp).Row
Range("a" + Format(COUNTULT + 1, "0")).FormulaR1C1 = N_Fwd_USDCAD_T
    
Ruta_Nombre = Ruta_PF + "\" + N_Fwd_USDCAD_T + ".txt"

With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + Ruta_Nombre, Destination:=Range("d" + Format(COUNTULT + 2, "00")))
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

COUNTULT1 = Sheets("Curvas").Cells(Rows.Count, 4).End(xlUp).Row
Sheets("Curvas").Range("f" + Format(COUNTULT + 2, "0"), "f" + Format(COUNTULT1, "0")).FormulaR1C1 = "=RC[-1]/10000"
Sheets("Curvas").Range("f" + Format(COUNTULT + 2, "0"), "f" + Format(COUNTULT1, "0")).Copy
Sheets("Curvas").Range("e" + Format(COUNTULT + 2, "0"), "e" + Format(COUNTULT1, "0")).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

Sheets("Curvas").Range("f:g").ClearContents
Sheets("Curvas").Range("a" + Format(COUNTULT + 2, "0")).FormulaR1C1 = 1
Sheets("Curvas").Range("b" + Format(COUNTULT + 2, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
Sheets("Curvas").Range("c" + Format(COUNTULT + 2, "0")).FormulaR1C1 = "FWPCAD"
Sheets("Curvas").Range("a" + Format(COUNTULT + 3, "0")).FormulaR1C1 = "=R[-1]C+1"
Sheets("Curvas").Range("b" + Format(COUNTULT + 3, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
Sheets("Curvas").Range("c" + Format(COUNTULT + 3, "0")).FormulaR1C1 = "FWPCAD"

Range("a" + Format(COUNTULT + 3, "0"), "c" + Format(COUNTULT + 3, "0")).Select
Selection.AutoFill Destination:=Range("a" + Format(COUNTULT + 3, "0"), "c" + Format(COUNTULT1, "00"))

Sheets("Curvas").Range("a" + Format(COUNTULT1 + 1, "00")).FormulaR1C1 = N_Fwd_USDBRL_T
Ruta_Nombre = Ruta_PF + "\" + N_Fwd_USDBRL_T + ".txt"

With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + Ruta_Nombre, Destination:=Range("d" + Format(COUNTULT1 + 2, "00")))
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
Sheets("Curvas").Range("f:g").ClearContents
COUNTULT2 = Sheets("Curvas").Cells(Rows.Count, 4).End(xlUp).Row

Sheets("Curvas").Range("a" + Format(COUNTULT1 + 2, "0")).FormulaR1C1 = 1
Sheets("Curvas").Range("b" + Format(COUNTULT1 + 2, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
Sheets("Curvas").Range("c" + Format(COUNTULT1 + 2, "0")).FormulaR1C1 = "FWPBRL"
Sheets("Curvas").Range("a" + Format(COUNTULT1 + 3, "0")).FormulaR1C1 = "=R[-1]C+1"
Sheets("Curvas").Range("b" + Format(COUNTULT1 + 3, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
Sheets("Curvas").Range("c" + Format(COUNTULT1 + 3, "0")).FormulaR1C1 = "FWPBRL"
Sheets("Curvas").Range("a" + Format(COUNTULT1 + 3, "0"), "c" + Format(COUNTULT1 + 3, "0")).Select
Selection.AutoFill Destination:=Range("a" + Format(COUNTULT1 + 3, "0"), "c" + Format(COUNTULT2, "00"))

Sheets("Curvas").Range("a" + Format(COUNTULT2 + 1, "0")).FormulaR1C1 = N_Fwd_USDMXN_T

Ruta_Nombre = Ruta_PF + "\" + N_Fwd_USDMXN_T + ".txt"

With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + Ruta_Nombre, Destination:=Range("d" + Format(COUNTULT2 + 2, "00")))
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

Sheets("Curvas").Range("a" + Format(COUNTULT2 + 2, "0")).FormulaR1C1 = 1
Sheets("Curvas").Range("b" + Format(COUNTULT2 + 2, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
Sheets("Curvas").Range("c" + Format(COUNTULT2 + 2, "0")).FormulaR1C1 = "FWPMXN"
Sheets("Curvas").Range("a" + Format(COUNTULT2 + 3, "0")).FormulaR1C1 = "=R[-1]C+1"
Sheets("Curvas").Range("b" + Format(COUNTULT2 + 3, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
Sheets("Curvas").Range("c" + Format(COUNTULT2 + 3, "0")).FormulaR1C1 = "FWPMXN"
Sheets("Curvas").Range("a" + Format(COUNTULT2 + 3, "0"), "c" + Format(COUNTULT2 + 3, "0")).Select
COUNTULT3 = Sheets("Curvas").Cells(Rows.Count, 4).End(xlUp).Row
Selection.AutoFill Destination:=Range("a" + Format(COUNTULT2 + 3, "0"), "c" + Format(COUNTULT3, "0"))
Sheets("Curvas").Range("f:g").ClearContents

Sheets("Curvas").Range("A" + Format(COUNTULT3 + 1, "0")).FormulaR1C1 = N_Fwd_USDJPY_T
Ruta_Nombre = Ruta_PF + "\" + N_Fwd_USDJPY_T + ".txt"

With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + Ruta_Nombre, Destination:=Range("d" + Format(COUNTULT3 + 2, "00")))
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

COUNTULT4 = Sheets("Curvas").Cells(Rows.Count, 4).End(xlUp).Row

Sheets("Curvas").Range("a" + Format(COUNTULT3 + 2, "0")).FormulaR1C1 = 1
Sheets("Curvas").Range("b" + Format(COUNTULT3 + 2, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
Sheets("Curvas").Range("c" + Format(COUNTULT3 + 2, "0")).FormulaR1C1 = "FWPJPY"
Sheets("Curvas").Range("a" + Format(COUNTULT3 + 3, "0")).FormulaR1C1 = "=R[-1]C+1"
Sheets("Curvas").Range("b" + Format(COUNTULT3 + 3, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
Sheets("Curvas").Range("c" + Format(COUNTULT3 + 3, "0")).FormulaR1C1 = "FWPJPY"

Sheets("Curvas").Range("a" + Format(COUNTULT3 + 3, "0"), "c" + Format(COUNTULT3 + 3, "0")).Select
Selection.AutoFill Destination:=Range("a" + Format(COUNTULT3 + 3, "0"), "c" + Format(COUNTULT4, "0"))

Sheets("Curvas").Range("f" + Format(COUNTULT3 + 2, "0"), "f" + Format(COUNTULT4, "0")).FormulaR1C1 = "=RC[-1]/100"
Sheets("Curvas").Range("f" + Format(COUNTULT3 + 2, "0"), "f" + Format(COUNTULT4, "0")).Copy
Sheets("Curvas").Range("e" + Format(COUNTULT3 + 2, "0"), "e" + Format(COUNTULT4, "0")).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Sheets("Curvas").Range("f:g").ClearContents


Sheets("Curvas").Range("a" + Format(COUNTULT4 + 1, "0")).FormulaR1C1 = N_Fwd_EURUSD_T
Ruta_Nombre = Ruta_PF + "\" + N_Fwd_EURUSD_T + ".txt"

With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + Ruta_Nombre, Destination:=Range("d" + Format(COUNTULT4 + 2, "00")))
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

COUNTULT5 = Sheets("Curvas").Cells(Rows.Count, 4).End(xlUp).Row
Sheets("Curvas").Range("a" + Format(COUNTULT4 + 2, "0")).FormulaR1C1 = 1
Sheets("Curvas").Range("b" + Format(COUNTULT4 + 2, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
Sheets("Curvas").Range("c" + Format(COUNTULT4 + 2, "0")).FormulaR1C1 = "FWPEUR"
Sheets("Curvas").Range("a" + Format(COUNTULT4 + 3, "0")).FormulaR1C1 = "=R[-1]C+1"
Sheets("Curvas").Range("b" + Format(COUNTULT4 + 3, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
Sheets("Curvas").Range("c" + Format(COUNTULT4 + 3, "0")).FormulaR1C1 = "FWPEUR"
Sheets("Curvas").Range("a" + Format(COUNTULT4 + 3, "0"), "c" + Format(COUNTULT4 + 3, "0")).Select
Selection.AutoFill Destination:=Range("a" + Format(COUNTULT4 + 3, "0"), "c" + Format(COUNTULT5, "0"))

Sheets("Curvas").Range("f" + Format(COUNTULT4 + 2, "0"), "f" + Format(COUNTULT5, "0")).FormulaR1C1 = "=RC[-1]/10000"
Sheets("Curvas").Range("f" + Format(COUNTULT4 + 2, "0"), "f" + Format(COUNTULT5, "0")).Copy
Sheets("Curvas").Range("e" + Format(COUNTULT4 + 2, "0"), "e" + Format(COUNTULT5, "0")).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Sheets("Curvas").Range("f:g").ClearContents

Sheets("Curvas").Range("a" + Format(COUNTULT5 + 1, "0")).FormulaR1C1 = N_SWAP_LBRUSD_T
Ruta_Nombre = Ruta_PF + "\" + N_SWAP_LBRUSD_T + ".txt"

With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + Ruta_Nombre, Destination:=Range("d" + Format(COUNTULT5 + 2, "00")))
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

COUNTULT6 = Sheets("Curvas").Cells(Rows.Count, 4).End(xlUp).Row
Sheets("Curvas").Range("a" + Format(COUNTULT5 + 2, "0")).FormulaR1C1 = 1
Sheets("Curvas").Range("b" + Format(COUNTULT5 + 2, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
Sheets("Curvas").Range("c" + Format(COUNTULT5 + 2, "0")).FormulaR1C1 = "LIBBTS"
Sheets("Curvas").Range("a" + Format(COUNTULT5 + 3, "0")).FormulaR1C1 = "=R[-1]C+1"
Sheets("Curvas").Range("b" + Format(COUNTULT5 + 3, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
Sheets("Curvas").Range("c" + Format(COUNTULT5 + 3, "0")).FormulaR1C1 = "LIBBTS"
Sheets("Curvas").Range("a" + Format(COUNTULT5 + 3, "0"), "c" + Format(COUNTULT5 + 3, "0")).Select
Selection.AutoFill Destination:=Range("a" + Format(COUNTULT5 + 3, "0"), "c" + Format(COUNTULT6, "0"))

Sheets("Curvas").Range("f" + Format(COUNTULT5 + 2, "0"), "f" + Format(COUNTULT6, "0")).FormulaR1C1 = "=RC[-1]*100"
Sheets("Curvas").Range("f" + Format(COUNTULT5 + 2, "0"), "f" + Format(COUNTULT6, "0")).Copy
Sheets("Curvas").Range("e" + Format(COUNTULT5 + 2, "0"), "e" + Format(COUNTULT6, "0")).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Sheets("Curvas").Range("f:g").ClearContents


Sheets("Curvas").Range("a" + Format(COUNTULT6 + 1, "0")).FormulaR1C1 = N_SWAP_LBRCOP_T
Ruta_Nombre = Ruta_PF + "\" + N_SWAP_LBRCOP_T + ".txt"

With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + Ruta_Nombre, Destination:=Range("d" + Format(COUNTULT6 + 2, "00")))
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

COUNTULT7 = Sheets("Curvas").Cells(Rows.Count, 4).End(xlUp).Row
Sheets("Curvas").Range("a" + Format(COUNTULT6 + 2, "0")).FormulaR1C1 = 1
Sheets("Curvas").Range("b" + Format(COUNTULT6 + 2, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
Sheets("Curvas").Range("c" + Format(COUNTULT6 + 2, "0")).FormulaR1C1 = "FWTCOP"
Sheets("Curvas").Range("a" + Format(COUNTULT6 + 3, "0")).FormulaR1C1 = "=R[-1]C+1"
Sheets("Curvas").Range("b" + Format(COUNTULT6 + 3, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
Sheets("Curvas").Range("c" + Format(COUNTULT6 + 3, "0")).FormulaR1C1 = "FWTCOP"
Sheets("Curvas").Range("a" + Format(COUNTULT6 + 3, "0"), "c" + Format(COUNTULT6 + 3, "0")).Select
Selection.AutoFill Destination:=Range("a" + Format(COUNTULT6 + 3, "0"), "c" + Format(COUNTULT7, "0"))

Sheets("Curvas").Range("f" + Format(COUNTULT6 + 2, "0"), "f" + Format(COUNTULT7, "0")).FormulaR1C1 = "=RC[-1]*100"
Sheets("Curvas").Range("f" + Format(COUNTULT6 + 2, "0"), "f" + Format(COUNTULT7, "0")).Copy
Sheets("Curvas").Range("e" + Format(COUNTULT6 + 2, "0"), "e" + Format(COUNTULT7, "0")).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Sheets("Curvas").Range("f:g").ClearContents


Sheets("Curvas").Range("a" + Format(COUNTULT7 + 1, "0")).FormulaR1C1 = N_Fwd_EURCOP_T
Ruta_Nombre = Ruta_PF + "\" + N_Fwd_EURCOP_T + ".csv"

With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + Ruta_Nombre, Destination:=Range("d" + Format(COUNTULT7 + 1, "00")))
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
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
End With

COUNTULT8 = Sheets("Curvas").Cells(Rows.Count, 4).End(xlUp).Row
Sheets("Curvas").Range("a" + Format(COUNTULT7 + 2, "0")).FormulaR1C1 = 1
Sheets("Curvas").Range("b" + Format(COUNTULT7 + 2, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
Sheets("Curvas").Range("c" + Format(COUNTULT7 + 2, "0")).FormulaR1C1 = "FWTEURCOP"
Sheets("Curvas").Range("a" + Format(COUNTULT7 + 3, "0")).FormulaR1C1 = "=R[-1]C+1"
Sheets("Curvas").Range("b" + Format(COUNTULT7 + 3, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
Sheets("Curvas").Range("c" + Format(COUNTULT7 + 3, "0")).FormulaR1C1 = "FWTEURCOP"
Sheets("Curvas").Range("a" + Format(COUNTULT7 + 3, "0"), "c" + Format(COUNTULT7 + 3, "0")).Select
Selection.AutoFill Destination:=Range("a" + Format(COUNTULT7 + 3, "0"), "c" + Format(COUNTULT8, "0"))

Sheets("Curvas").Range("f" + Format(COUNTULT7 + 2, "0"), "f" + Format(COUNTULT8, "0")).FormulaR1C1 = "=+RC[5]"
Sheets("Curvas").Range("f" + Format(COUNTULT7 + 2, "0"), "f" + Format(COUNTULT8, "0")).Copy
Sheets("Curvas").Range("e" + Format(COUNTULT7 + 2, "0"), "e" + Format(COUNTULT8, "0")).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Sheets("Curvas").Range("f:m").ClearContents

Sheets("Curvas").Range("a" + Format(COUNTULT8 + 1, "0")).FormulaR1C1 = N_Fwd_EUR2COP_T
Ruta_Nombre = Ruta_PF + "\" + N_Fwd_EUR2COP_T + ".txt"

With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + Ruta_Nombre, Destination:=Range("d" + Format(COUNTULT8 + 2, "00")))
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
        .TextFileSemicolonDelimiter = True
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = True
        .TextFileColumnDataTypes = Array(1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
End With

COUNTULT9 = Sheets("Curvas").Cells(Rows.Count, 4).End(xlUp).Row
Sheets("Curvas").Range("a" + Format(COUNTULT8 + 2, "0")).FormulaR1C1 = 1
Sheets("Curvas").Range("b" + Format(COUNTULT8 + 2, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
Sheets("Curvas").Range("c" + Format(COUNTULT8 + 2, "0")).FormulaR1C1 = "Fwd_EUR2COP"
Sheets("Curvas").Range("a" + Format(COUNTULT8 + 3, "0")).FormulaR1C1 = "=R[-1]C+1"
Sheets("Curvas").Range("b" + Format(COUNTULT8 + 3, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
Sheets("Curvas").Range("c" + Format(COUNTULT8 + 3, "0")).FormulaR1C1 = "Fwd_EUR2COP"
Sheets("Curvas").Range("a" + Format(COUNTULT8 + 3, "0"), "c" + Format(COUNTULT8 + 3, "0")).Select
Selection.AutoFill Destination:=Range("a" + Format(COUNTULT8 + 3, "0"), "c" + Format(COUNTULT9, "0"))

Sheets("Curvas").Range("f" + Format(COUNTULT8 + 2, "0"), "f" + Format(COUNTULT9, "0")).FormulaR1C1 = "=RC[-1]"
Sheets("Curvas").Range("f" + Format(COUNTULT8 + 2, "0"), "f" + Format(COUNTULT9, "0")).Copy
Sheets("Curvas").Range("e" + Format(COUNTULT8 + 2, "0"), "e" + Format(COUNTULT9, "0")).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Sheets("Curvas").Range("f:g").ClearContents


Sheets("Curvas").Range("a" + Format(COUNTULT9 + 1, "0")).FormulaR1C1 = N_Fwd_GBPUSD_T
Ruta_Nombre = Ruta_PF + "\" + N_Fwd_GBPUSD_T + ".txt"

With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + Ruta_Nombre, Destination:=Range("d" + Format(COUNTULT9 + 2, "00")))
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

COUNTULT10 = Sheets("Curvas").Cells(Rows.Count, 4).End(xlUp).Row
Sheets("Curvas").Range("a" + Format(COUNTULT9 + 2, "0")).FormulaR1C1 = 1
Sheets("Curvas").Range("b" + Format(COUNTULT9 + 2, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
Sheets("Curvas").Range("c" + Format(COUNTULT9 + 2, "0")).FormulaR1C1 = "Fwd_GBPUSD_Diaria"
Sheets("Curvas").Range("a" + Format(COUNTULT9 + 3, "0")).FormulaR1C1 = "=R[-1]C+1"
Sheets("Curvas").Range("b" + Format(COUNTULT9 + 3, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
Sheets("Curvas").Range("c" + Format(COUNTULT9 + 3, "0")).FormulaR1C1 = "Fwd_GBPUSD_Diaria"
Sheets("Curvas").Range("a" + Format(COUNTULT9 + 3, "0"), "c" + Format(COUNTULT9 + 3, "0")).Select
Selection.AutoFill Destination:=Range("a" + Format(COUNTULT9 + 3, "0"), "c" + Format(COUNTULT10, "0"))

Sheets("Curvas").Range("F" + Format(COUNTULT9 + 2, "0"), "F" + Format(COUNTULT10, "0")).FormulaR1C1 = "=RC[-1]/10000"
Sheets("Curvas").Range("F" + Format(COUNTULT9 + 2, "0"), "F" + Format(COUNTULT10, "0")).Copy
Sheets("Curvas").Range("e" + Format(COUNTULT9 + 2, "0"), "e" + Format(COUNTULT10, "0")).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Sheets("Curvas").Range("f:g").ClearContents

'' CURVA USDCLP ''

Sheets("Curvas").Range("a" + Format(COUNTULT10 + 1, "0")).FormulaR1C1 = N_Fwd_USDCLP_T
Ruta_Nombre = Ruta_PF + "\" + N_Fwd_USDCLP_T + ".txt"

With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + Ruta_Nombre, Destination:=Range("d" + Format(COUNTULT10 + 2, "00")))
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

COUNTULT11 = Sheets("Curvas").Cells(Rows.Count, 4).End(xlUp).Row
Sheets("Curvas").Range("a" + Format(COUNTULT10 + 2, "0")).FormulaR1C1 = 1
Sheets("Curvas").Range("b" + Format(COUNTULT10 + 2, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
Sheets("Curvas").Range("c" + Format(COUNTULT10 + 2, "0")).FormulaR1C1 = "FWPCLP"
Sheets("Curvas").Range("a" + Format(COUNTULT10 + 3, "0")).FormulaR1C1 = "=R[-1]C+1"
Sheets("Curvas").Range("b" + Format(COUNTULT10 + 3, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
Sheets("Curvas").Range("c" + Format(COUNTULT10 + 3, "0")).FormulaR1C1 = "FWPCLP"
Sheets("Curvas").Range("a" + Format(COUNTULT10 + 3, "0"), "c" + Format(COUNTULT10 + 3, "0")).Select
Selection.AutoFill Destination:=Range("a" + Format(COUNTULT10 + 3, "0"), "c" + Format(COUNTULT11, "0"))

'Sheets("Curvas").Range("e" + Format(COUNTULT10 + 2, "0"), "e" + Format(COUNTULT11, "0")).FormulaR1C1 = "=RC[1]/100"
Sheets("Curvas").Range("F" + Format(COUNTULT10 + 2, "0"), "F" + Format(COUNTULT11, "0")).FormulaR1C1 = "=RC[-1]/100"
Sheets("Curvas").Range("G" + Format(COUNTULT10 + 2, "0"), "G" + Format(COUNTULT11, "0")).FormulaR1C1 = "=RC[-1]/100"
Sheets("Curvas").Range("F" + Format(COUNTULT10 + 2, "0"), "F" + Format(COUNTULT11, "0")).Copy
Sheets("Curvas").Range("e" + Format(COUNTULT10 + 2, "0"), "e" + Format(COUNTULT11, "0")).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Sheets("Curvas").Range("f:g").ClearContents

'' CURVA USDCHF ''

Sheets("Curvas").Range("A" + Format(COUNTULT11 + 1, "0")).FormulaR1C1 = N_Fwd_USDCHF_T
Ruta_Nombre = Ruta_PF + "\" + N_Fwd_USDCHF_T + ".txt"

With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + Ruta_Nombre, Destination:=Range("d" + Format(COUNTULT11 + 2, "00")))
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

COUNTULT12 = Sheets("Curvas").Cells(Rows.Count, 4).End(xlUp).Row

Sheets("Curvas").Range("a" + Format(COUNTULT11 + 2, "0")).FormulaR1C1 = 1
Sheets("Curvas").Range("b" + Format(COUNTULT11 + 2, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
Sheets("Curvas").Range("c" + Format(COUNTULT11 + 2, "0")).FormulaR1C1 = "FWPCHF"
Sheets("Curvas").Range("a" + Format(COUNTULT11 + 3, "0")).FormulaR1C1 = "=R[-1]C+1"
Sheets("Curvas").Range("b" + Format(COUNTULT11 + 3, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
Sheets("Curvas").Range("c" + Format(COUNTULT11 + 3, "0")).FormulaR1C1 = "FWPCHF"

Sheets("Curvas").Range("a" + Format(COUNTULT11 + 3, "0"), "c" + Format(COUNTULT11 + 3, "0")).Select
Selection.AutoFill Destination:=Range("a" + Format(COUNTULT11 + 3, "0"), "c" + Format(COUNTULT12, "0"))

Sheets("Curvas").Range("f" + Format(COUNTULT11 + 2, "0"), "f" + Format(COUNTULT12, "0")).FormulaR1C1 = "=RC[-1]/10000"
Sheets("Curvas").Range("f" + Format(COUNTULT11 + 2, "0"), "f" + Format(COUNTULT12, "0")).Copy
Sheets("Curvas").Range("e" + Format(COUNTULT11 + 2, "0"), "e" + Format(COUNTULT12, "0")).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Sheets("Curvas").Range("f:g").ClearContents


'' CURVA USDKRW ''

'Sheets("Curvas").Range("A" + Format(COUNTULT12 + 1, "0")).FormulaR1C1 = N_Fwd_USDKRW_T
'Ruta_Nombre = Ruta_PF + "\" + N_Fwd_USDKRW_T + ".txt"
'
'With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + Ruta_Nombre, Destination:=Range("d" + Format(COUNTULT12 + 2, "00")))
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
'End With
'
'COUNTULT13 = Sheets("Curvas").Cells(Rows.Count, 4).End(xlUp).Row
'
'Sheets("Curvas").Range("a" + Format(COUNTULT12 + 2, "0")).FormulaR1C1 = 1
'Sheets("Curvas").Range("b" + Format(COUNTULT12 + 2, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
'Sheets("Curvas").Range("c" + Format(COUNTULT12 + 2, "0")).FormulaR1C1 = "FWPKRW"
'Sheets("Curvas").Range("a" + Format(COUNTULT12 + 3, "0")).FormulaR1C1 = "=R[-1]C+1"
'Sheets("Curvas").Range("b" + Format(COUNTULT12 + 3, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
'Sheets("Curvas").Range("c" + Format(COUNTULT12 + 3, "0")).FormulaR1C1 = "FWPKRW"
'
'Sheets("Curvas").Range("a" + Format(COUNTULT12 + 3, "0"), "c" + Format(COUNTULT12 + 3, "0")).Select
'Selection.AutoFill Destination:=Range("a" + Format(COUNTULT12 + 3, "0"), "c" + Format(COUNTULT13, "0"))
'
'Sheets("Curvas").Range("f" + Format(COUNTULT12 + 2, "0"), "f" + Format(COUNTULT13, "0")).FormulaR1C1 = "=RC[-1]/100"
'Sheets("Curvas").Range("f" + Format(COUNTULT12 + 2, "0"), "f" + Format(COUNTULT13, "0")).Copy
'Sheets("Curvas").Range("e" + Format(COUNTULT12 + 2, "0"), "e" + Format(COUNTULT13, "0")).Select
'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'Sheets("Curvas").Range("f:g").ClearContents

End Sub

Sub IMP_CURVAS_ANT()

'
'Dim Ruta_PF As String
'Dim N_Fwd_USDCOP_T As String
'Dim N_Fwd_USDCOP_T_1 As String
'Dim N_Fwd_USDCAD_T As String
'Dim N_Fwd_USDBRL_T As String
'Dim N_Fwd_USDMXN_T As String
'Dim N_Fwd_USDJPY_T As String
'Dim N_Fwd_EURUSD_T As String
'Dim N_SWAP_LBRUSD_T As String
'Dim N_SWAP_LBRCOP_T As String
'Dim N_Fwd_EURCOP_T As String
'Dim N_Fwd_EUR2COP_T As String
'Dim N_Fwd_GBPUSD_T As String
'
'
'Fecha_Valoración_ANT = Range("Fecha_Valoración_ANT")
'
'N_Fwd_USDCOP_T = Range("N_Fwd_USDCOP_T")
'N_Fwd_USDCOP_T_1 = Range("N_Fwd_USDCOP_T_1")
'N_Fwd_USDCAD_T = Range("N_Fwd_USDCAD_T")
'N_Fwd_USDCAD_T_1 = Range("N_Fwd_USDCAD_T_1")
'N_Fwd_USDBRL_T = Range("N_Fwd_USDBRL_T")
'N_Fwd_USDBRL_T_1 = Range("N_Fwd_USDBRL_T_1")
'N_Fwd_USDMXN_T = Range("N_Fwd_USDMXN_T")
'N_Fwd_USDMXN_T_1 = Range("N_Fwd_USDMXN_T_1")
'N_Fwd_USDJPY_T = Range("N_Fwd_USDJPY_T")
'N_Fwd_USDJPY_T_1 = Range("N_Fwd_USDJPY_T_1")
'N_Fwd_EURUSD_T = Range("N_Fwd_EURUSD_T")
'N_Fwd_EURUSD_T_1 = Range("N_Fwd_EURUSD_T_1")
'N_SWAP_LBRUSD_T = Range("N_SWAP_LBRUSD_T")
'N_SWAP_LBRCOP_T_1 = Range("N_SWAP_LBRCOP_T_1")
'N_SWAP_LBRUSD_T_1 = Range("N_SWAP_LBRUSD_T_1")
'N_Fwd_EURCOP_T_1 = Range("N_Fwd_EURCOP_T_1")
'N_Fwd_EUR2COP_T_1 = Range("N_Fwd_EUR2COP_T_1")
'N_Fwd_GBPUSD_T_1 = Range("N_Fwd_GBPUSD_T_1")
'
'Ruta_PF = Range("Ruta_PF")
'
'Sheets("Curvas").Activate
'Sheets("Curvas").Range("H:aa").Select
'Sheets("Curvas").Range("H:aa").ClearContents
'
'    Sheets("Curvas").Range("H1").FormulaR1C1 = "CURVAS FORWARD ACTUAL"
'    Sheets("Curvas").Range("H2").FormulaR1C1 = "Orden"
'    Sheets("Curvas").Range("I2").FormulaR1C1 = "Fecha_Valoración_ANT_ANT"
'    Sheets("Curvas").Range("J2").FormulaR1C1 = "Nombre_Indicador"
'    Sheets("Curvas").Range("K2").FormulaR1C1 = "Plazo_en_Dias"
'    Sheets("Curvas").Range("L2").FormulaR1C1 = "Valor"
'
'    Sheets("Curvas").Range("H3").FormulaR1C1 = N_Fwd_USDCOP_T_1
'
'Ruta_Nombre = Ruta_PF + "\" + N_Fwd_USDCOP_T_1 + ".txt"
'
'With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + Ruta_Nombre, Destination:=Range("$K$4"))
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
'End With
'
'Sheets("Curvas").Range("H4").FormulaR1C1 = 1
'Sheets("Curvas").Range("I4").FormulaR1C1 = "=Fecha_Valoración_ANT"
'Sheets("Curvas").Range("J4").FormulaR1C1 = "FWPCOP"
'Sheets("Curvas").Range("H5").FormulaR1C1 = "=R[-1]C+1"
'Sheets("Curvas").Range("I5").FormulaR1C1 = "=Fecha_Valoración_ANT"
'Sheets("Curvas").Range("J5").FormulaR1C1 = "FWPCOP"
'
'    Range("H5:J5").Select
'    Selection.AutoFill Destination:=Range("H5", "J" + Format(Sheets("Curvas").Cells(Rows.Count, 11).End(xlUp).Row, "00"))
'
'
'
'COUNTULT = Sheets("Curvas").Cells(Rows.Count, 8).End(xlUp).Row
'Range("H" + Format(COUNTULT + 1, "0")).FormulaR1C1 = N_Fwd_USDCAD_T_1
'
'Ruta_Nombre = Ruta_PF + "\" + N_Fwd_USDCAD_T_1 + ".txt"
'
'With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + Ruta_Nombre, Destination:=Range("K" + Format(COUNTULT + 2, "00")))
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
'End With
'
'COUNTULT1 = Sheets("Curvas").Cells(Rows.Count, 11).End(xlUp).Row
'Sheets("Curvas").Range("M" + Format(COUNTULT + 2, "0"), "M" + Format(COUNTULT1, "0")).FormulaR1C1 = "=RC[-1]/10000"
'Sheets("Curvas").Range("M" + Format(COUNTULT + 2, "0"), "M" + Format(COUNTULT1, "0")).Copy
'Sheets("Curvas").Range("L" + Format(COUNTULT + 2, "0"), "L" + Format(COUNTULT1, "0")).Select
'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'
'Sheets("Curvas").Range("M:N").ClearContents
'Sheets("Curvas").Range("H" + Format(COUNTULT + 2, "0")).FormulaR1C1 = 1
'Sheets("Curvas").Range("I" + Format(COUNTULT + 2, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
'Sheets("Curvas").Range("J" + Format(COUNTULT + 2, "0")).FormulaR1C1 = "FWPCAD"
'Sheets("Curvas").Range("H" + Format(COUNTULT + 3, "0")).FormulaR1C1 = "=R[-1]C+1"
'Sheets("Curvas").Range("I" + Format(COUNTULT + 3, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
'Sheets("Curvas").Range("J" + Format(COUNTULT + 3, "0")).FormulaR1C1 = "FWPCAD"
'
'Range("H" + Format(COUNTULT + 3, "0"), "J" + Format(COUNTULT + 3, "0")).Select
'Selection.AutoFill Destination:=Range("H" + Format(COUNTULT + 3, "0"), "J" + Format(COUNTULT1, "00"))
'
'Sheets("Curvas").Range("H" + Format(COUNTULT1 + 1, "00")).FormulaR1C1 = N_Fwd_USDBRL_T_1
'Ruta_Nombre = Ruta_PF + "\" + N_Fwd_USDBRL_T_1 + ".txt"
'
'With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + Ruta_Nombre, Destination:=Range("K" + Format(COUNTULT1 + 2, "00")))
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
'End With
'Sheets("Curvas").Range("M:N").ClearContents
'COUNTULT2 = Sheets("Curvas").Cells(Rows.Count, 11).End(xlUp).Row
'
'Sheets("Curvas").Range("H" + Format(COUNTULT1 + 2, "0")).FormulaR1C1 = 1
'Sheets("Curvas").Range("I" + Format(COUNTULT1 + 2, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
'Sheets("Curvas").Range("J" + Format(COUNTULT1 + 2, "0")).FormulaR1C1 = "FWPBRL"
'Sheets("Curvas").Range("H" + Format(COUNTULT1 + 3, "0")).FormulaR1C1 = "=R[-1]C+1"
'Sheets("Curvas").Range("I" + Format(COUNTULT1 + 3, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
'Sheets("Curvas").Range("J" + Format(COUNTULT1 + 3, "0")).FormulaR1C1 = "FWPBRL"
'Sheets("Curvas").Range("H" + Format(COUNTULT1 + 3, "0"), "J" + Format(COUNTULT1 + 3, "0")).Select
'Selection.AutoFill Destination:=Range("H" + Format(COUNTULT1 + 3, "0"), "J" + Format(COUNTULT2, "00"))
'
'Sheets("Curvas").Range("H" + Format(COUNTULT2 + 1, "0")).FormulaR1C1 = N_Fwd_USDMXN_T_1
'
'Ruta_Nombre = Ruta_PF + "\" + N_Fwd_USDMXN_T_1 + ".txt"
'
'With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + Ruta_Nombre, Destination:=Range("K" + Format(COUNTULT2 + 2, "00")))
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
'End With
'
'Sheets("Curvas").Range("H" + Format(COUNTULT2 + 2, "0")).FormulaR1C1 = 1
'Sheets("Curvas").Range("I" + Format(COUNTULT2 + 2, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
'Sheets("Curvas").Range("J" + Format(COUNTULT2 + 2, "0")).FormulaR1C1 = "FWPMXN"
'Sheets("Curvas").Range("H" + Format(COUNTULT2 + 3, "0")).FormulaR1C1 = "=R[-1]C+1"
'Sheets("Curvas").Range("I" + Format(COUNTULT2 + 3, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
'Sheets("Curvas").Range("J" + Format(COUNTULT2 + 3, "0")).FormulaR1C1 = "FWPMXN"
'Sheets("Curvas").Range("H" + Format(COUNTULT2 + 3, "0"), "J" + Format(COUNTULT2 + 3, "0")).Select
'COUNTULT3 = Sheets("Curvas").Cells(Rows.Count, 11).End(xlUp).Row
'Selection.AutoFill Destination:=Range("H" + Format(COUNTULT2 + 3, "0"), "J" + Format(COUNTULT3, "0"))
'Sheets("Curvas").Range("M:N").ClearContents
'
'Sheets("Curvas").Range("H" + Format(COUNTULT3 + 1, "0")).FormulaR1C1 = N_Fwd_USDJPY_T_1
'Ruta_Nombre = Ruta_PF + "\" + N_Fwd_USDJPY_T_1 + ".txt"
'
'With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + Ruta_Nombre, Destination:=Range("K" + Format(COUNTULT3 + 2, "00")))
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
'End With
'
'COUNTULT4 = Sheets("Curvas").Cells(Rows.Count, 11).End(xlUp).Row
'
'Sheets("Curvas").Range("H" + Format(COUNTULT3 + 2, "0")).FormulaR1C1 = 1
'Sheets("Curvas").Range("I" + Format(COUNTULT3 + 2, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
'Sheets("Curvas").Range("J" + Format(COUNTULT3 + 2, "0")).FormulaR1C1 = "FWPJPY"
'Sheets("Curvas").Range("H" + Format(COUNTULT3 + 3, "0")).FormulaR1C1 = "=R[-1]C+1"
'Sheets("Curvas").Range("I" + Format(COUNTULT3 + 3, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
'Sheets("Curvas").Range("J" + Format(COUNTULT3 + 3, "0")).FormulaR1C1 = "FWPJPY"
'
'Sheets("Curvas").Range("H" + Format(COUNTULT3 + 3, "0"), "J" + Format(COUNTULT3 + 3, "0")).Select
'Selection.AutoFill Destination:=Range("H" + Format(COUNTULT3 + 3, "0"), "J" + Format(COUNTULT4, "0"))
'
'Sheets("Curvas").Range("M" + Format(COUNTULT3 + 2, "0"), "M" + Format(COUNTULT4, "0")).FormulaR1C1 = "=RC[-1]/100"
'Sheets("Curvas").Range("M" + Format(COUNTULT3 + 2, "0"), "M" + Format(COUNTULT4, "0")).Copy
'Sheets("Curvas").Range("L" + Format(COUNTULT3 + 2, "0"), "L" + Format(COUNTULT4, "0")).Select
'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'Sheets("Curvas").Range("M:N").ClearContents
'
'
'Sheets("Curvas").Range("H" + Format(COUNTULT4 + 1, "0")).FormulaR1C1 = N_Fwd_EURUSD_T_1
'Ruta_Nombre = Ruta_PF + "\" + N_Fwd_EURUSD_T_1 + ".txt"
'
'With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + Ruta_Nombre, Destination:=Range("K" + Format(COUNTULT4 + 2, "00")))
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
'End With
'
'COUNTULT5 = Sheets("Curvas").Cells(Rows.Count, 11).End(xlUp).Row
'Sheets("Curvas").Range("H" + Format(COUNTULT4 + 2, "0")).FormulaR1C1 = 1
'Sheets("Curvas").Range("I" + Format(COUNTULT4 + 2, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
'Sheets("Curvas").Range("J" + Format(COUNTULT4 + 2, "0")).FormulaR1C1 = "FWPEUR"
'Sheets("Curvas").Range("H" + Format(COUNTULT4 + 3, "0")).FormulaR1C1 = "=R[-1]C+1"
'Sheets("Curvas").Range("I" + Format(COUNTULT4 + 3, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
'Sheets("Curvas").Range("J" + Format(COUNTULT4 + 3, "0")).FormulaR1C1 = "FWPEUR"
'Sheets("Curvas").Range("H" + Format(COUNTULT4 + 3, "0"), "J" + Format(COUNTULT4 + 3, "0")).Select
'Selection.AutoFill Destination:=Range("H" + Format(COUNTULT4 + 3, "0"), "J" + Format(COUNTULT5, "0"))
'
'Sheets("Curvas").Range("M" + Format(COUNTULT4 + 2, "0"), "M" + Format(COUNTULT5, "0")).FormulaR1C1 = "=RC[-1]/10000"
'Sheets("Curvas").Range("M" + Format(COUNTULT4 + 2, "0"), "M" + Format(COUNTULT5, "0")).Copy
'Sheets("Curvas").Range("L" + Format(COUNTULT4 + 2, "0"), "L" + Format(COUNTULT5, "0")).Select
'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'Sheets("Curvas").Range("M:N").ClearContents
'
'Sheets("Curvas").Range("H" + Format(COUNTULT5 + 1, "0")).FormulaR1C1 = N_SWAP_LBRUSD_T_1
'Ruta_Nombre = Ruta_PF + "\" + N_SWAP_LBRUSD_T_1 + ".txt"
'
'With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + Ruta_Nombre, Destination:=Range("K" + Format(COUNTULT5 + 2, "00")))
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
'End With
'
'COUNTULT6 = Sheets("Curvas").Cells(Rows.Count, 11).End(xlUp).Row
'Sheets("Curvas").Range("H" + Format(COUNTULT5 + 2, "0")).FormulaR1C1 = 1
'Sheets("Curvas").Range("I" + Format(COUNTULT5 + 2, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
'Sheets("Curvas").Range("J" + Format(COUNTULT5 + 2, "0")).FormulaR1C1 = "LIBBTS"
'Sheets("Curvas").Range("H" + Format(COUNTULT5 + 3, "0")).FormulaR1C1 = "=R[-1]C+1"
'Sheets("Curvas").Range("I" + Format(COUNTULT5 + 3, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
'Sheets("Curvas").Range("J" + Format(COUNTULT5 + 3, "0")).FormulaR1C1 = "LIBBTS"
'Sheets("Curvas").Range("H" + Format(COUNTULT5 + 3, "0"), "J" + Format(COUNTULT5 + 3, "0")).Select
'Selection.AutoFill Destination:=Range("H" + Format(COUNTULT5 + 3, "0"), "J" + Format(COUNTULT6, "0"))
'
'Sheets("Curvas").Range("M" + Format(COUNTULT5 + 2, "0"), "M" + Format(COUNTULT6, "0")).FormulaR1C1 = "=RC[-1]*100"
'Sheets("Curvas").Range("M" + Format(COUNTULT5 + 2, "0"), "M" + Format(COUNTULT6, "0")).Copy
'Sheets("Curvas").Range("L" + Format(COUNTULT5 + 2, "0"), "L" + Format(COUNTULT6, "0")).Select
'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'Sheets("Curvas").Range("M:N").ClearContents
'
'
'Sheets("Curvas").Range("H" + Format(COUNTULT6 + 1, "0")).FormulaR1C1 = N_SWAP_LBRCOP_T_1
'Ruta_Nombre = Ruta_PF + "\" + N_SWAP_LBRCOP_T_1 + ".txt"
'
'With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + Ruta_Nombre, Destination:=Range("K" + Format(COUNTULT6 + 2, "00")))
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
'End With
'
'COUNTULT7 = Sheets("Curvas").Cells(Rows.Count, 11).End(xlUp).Row
'Sheets("Curvas").Range("H" + Format(COUNTULT6 + 2, "0")).FormulaR1C1 = 1
'Sheets("Curvas").Range("I" + Format(COUNTULT6 + 2, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
'Sheets("Curvas").Range("J" + Format(COUNTULT6 + 2, "0")).FormulaR1C1 = "FWTCOP"
'Sheets("Curvas").Range("H" + Format(COUNTULT6 + 3, "0")).FormulaR1C1 = "=R[-1]C+1"
'Sheets("Curvas").Range("I" + Format(COUNTULT6 + 3, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
'Sheets("Curvas").Range("J" + Format(COUNTULT6 + 3, "0")).FormulaR1C1 = "FWTCOP"
'Sheets("Curvas").Range("H" + Format(COUNTULT6 + 3, "0"), "J" + Format(COUNTULT6 + 3, "0")).Select
'Selection.AutoFill Destination:=Range("H" + Format(COUNTULT6 + 3, "0"), "J" + Format(COUNTULT7, "0"))
'
'Sheets("Curvas").Range("M" + Format(COUNTULT6 + 2, "0"), "M" + Format(COUNTULT7, "0")).FormulaR1C1 = "=RC[-1]*100"
'Sheets("Curvas").Range("M" + Format(COUNTULT6 + 2, "0"), "M" + Format(COUNTULT7, "0")).Copy
'Sheets("Curvas").Range("L" + Format(COUNTULT6 + 2, "0"), "L" + Format(COUNTULT7, "0")).Select
'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'Sheets("Curvas").Range("M:N").ClearContents
'
'
'Sheets("Curvas").Range("H" + Format(COUNTULT7 + 1, "0")).FormulaR1C1 = N_Fwd_EURCOP_T_1
'Ruta_Nombre = Ruta_PF + "\" + N_Fwd_EURCOP_T_1 + ".csv"
'
'With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + Ruta_Nombre, Destination:=Range("K" + Format(COUNTULT7 + 1, "00")))
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
'        .TextFileConsecutiveDelimiter = False
'        .TextFileTabDelimiter = True
'        .TextFileSemicolonDelimiter = True
'        .TextFileCommaDelimiter = True
'        .TextFileSpaceDelimiter = False
'        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1)
'        .TextFileTrailingMinusNumbers = True
'        .Refresh BackgroundQuery:=False
'End With
'
'COUNTULT8 = Sheets("Curvas").Cells(Rows.Count, 11).End(xlUp).Row
'Sheets("Curvas").Range("H" + Format(COUNTULT7 + 2, "0")).FormulaR1C1 = 1
'Sheets("Curvas").Range("I" + Format(COUNTULT7 + 2, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
'Sheets("Curvas").Range("J" + Format(COUNTULT7 + 2, "0")).FormulaR1C1 = "FWTEURCOP"
'Sheets("Curvas").Range("H" + Format(COUNTULT7 + 3, "0")).FormulaR1C1 = "=R[-1]C+1"
'Sheets("Curvas").Range("I" + Format(COUNTULT7 + 3, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
'Sheets("Curvas").Range("J" + Format(COUNTULT7 + 3, "0")).FormulaR1C1 = "FWTEURCOP"
'Sheets("Curvas").Range("H" + Format(COUNTULT7 + 3, "0"), "J" + Format(COUNTULT7 + 3, "0")).Select
'Selection.AutoFill Destination:=Range("H" + Format(COUNTULT7 + 3, "0"), "J" + Format(COUNTULT8, "0"))
'
'Sheets("Curvas").Range("M" + Format(COUNTULT7 + 2, "0"), "M" + Format(COUNTULT8, "0")).FormulaR1C1 = "=+RC[5]"
'Sheets("Curvas").Range("M" + Format(COUNTULT7 + 2, "0"), "M" + Format(COUNTULT8, "0")).Copy
'Sheets("Curvas").Range("L" + Format(COUNTULT7 + 2, "0"), "L" + Format(COUNTULT8, "0")).Select
'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'Sheets("Curvas").Range("M:AA").ClearContents
'
'Sheets("Curvas").Range("H" + Format(COUNTULT8 + 1, "0")).FormulaR1C1 = N_Fwd_EUR2COP_T_1
'Ruta_Nombre = Ruta_PF + "\" + N_Fwd_EUR2COP_T_1 + ".txt"
'
'With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + Ruta_Nombre, Destination:=Range("K" + Format(COUNTULT8 + 2, "00")))
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
'        .TextFileSemicolonDelimiter = True
'        .TextFileCommaDelimiter = True
'        .TextFileSpaceDelimiter = True
'        .TextFileColumnDataTypes = Array(1, 1, 1, 1)
'        .TextFileTrailingMinusNumbers = True
'        .Refresh BackgroundQuery:=False
'End With
'
'COUNTULT9 = Sheets("Curvas").Cells(Rows.Count, 11).End(xlUp).Row
'Sheets("Curvas").Range("H" + Format(COUNTULT8 + 2, "0")).FormulaR1C1 = 1
'Sheets("Curvas").Range("I" + Format(COUNTULT8 + 2, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
'Sheets("Curvas").Range("J" + Format(COUNTULT8 + 2, "0")).FormulaR1C1 = "Fwd_EUR2COP"
'Sheets("Curvas").Range("H" + Format(COUNTULT8 + 3, "0")).FormulaR1C1 = "=R[-1]C+1"
'Sheets("Curvas").Range("I" + Format(COUNTULT8 + 3, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
'Sheets("Curvas").Range("J" + Format(COUNTULT8 + 3, "0")).FormulaR1C1 = "Fwd_EUR2COP"
'Sheets("Curvas").Range("H" + Format(COUNTULT8 + 3, "0"), "J" + Format(COUNTULT8 + 3, "0")).Select
'Selection.AutoFill Destination:=Range("H" + Format(COUNTULT8 + 3, "0"), "J" + Format(COUNTULT9, "0"))
'
''Sheets("Curvas").Range("f" + Format(COUNTULT8 + 2, "0"), "f" + Format(COUNTULT9, "0")).FormulaR1C1 = "=RC[1]"
'Sheets("Curvas").Range("M" + Format(COUNTULT8 + 2, "0"), "M" + Format(COUNTULT9, "0")).Copy
'Sheets("Curvas").Range("L" + Format(COUNTULT8 + 2, "0"), "L" + Format(COUNTULT9, "0")).Select
'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'Sheets("Curvas").Range("M:N").ClearContents
'
'
'Sheets("Curvas").Range("H" + Format(COUNTULT9 + 1, "0")).FormulaR1C1 = N_Fwd_GBPUSD_T_1
'Ruta_Nombre = Ruta_PF + "\" + N_Fwd_GBPUSD_T_1 + ".txt"
'
'With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + Ruta_Nombre, Destination:=Range("K" + Format(COUNTULT9 + 2, "00")))
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
'End With
'
'COUNTULT10 = Sheets("Curvas").Cells(Rows.Count, 11).End(xlUp).Row
'Sheets("Curvas").Range("H" + Format(COUNTULT9 + 2, "0")).FormulaR1C1 = 1
'Sheets("Curvas").Range("I" + Format(COUNTULT9 + 2, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
'Sheets("Curvas").Range("J" + Format(COUNTULT9 + 2, "0")).FormulaR1C1 = "Fwd_GBPUSD_Diaria"
'Sheets("Curvas").Range("H" + Format(COUNTULT9 + 3, "0")).FormulaR1C1 = "=R[-1]C+1"
'Sheets("Curvas").Range("I" + Format(COUNTULT9 + 3, "0")).FormulaR1C1 = "=Fecha_Valoración_ANT"
'Sheets("Curvas").Range("J" + Format(COUNTULT9 + 3, "0")).FormulaR1C1 = "Fwd_GBPUSD_Diaria"
'Sheets("Curvas").Range("H" + Format(COUNTULT9 + 3, "0"), "J" + Format(COUNTULT9 + 3, "0")).Select
'Selection.AutoFill Destination:=Range("H" + Format(COUNTULT9 + 3, "0"), "J" + Format(COUNTULT10, "0"))
'
'Sheets("Curvas").Range("L" + Format(COUNTULT9 + 2, "0"), "L" + Format(COUNTULT10, "0")).FormulaR1C1 = "=RC[1]/10000"
'Sheets("Curvas").Range("L" + Format(COUNTULT9 + 2, "0"), "L" + Format(COUNTULT10, "0")).Copy
'Sheets("Curvas").Range("L" + Format(COUNTULT9 + 2, "0"), "L" + Format(COUNTULT10, "0")).Select
'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'Sheets("Curvas").Range("M:N").ClearContents
'
End Sub


Sub Importar_indicadores()


'Dim Fecha_Valoración, Fecha_Valoración_T_1, Fecha_Valoración_T_2 As String

Application.ScreenUpdating = False
Fecha_Valoración = Range("Fecha_Valoración")
Fecha_Valoración_T_1 = Fecha_Valoración - 1
Fecha_Valoración_T_2 = Fecha_Valoración - 2

N_IND_1 = Range("N_IND_1")
N_IND_2 = Range("N_IND_2")
N_IND_3 = Range("N_IND_3")
R_IND_1 = Range("R_IND_1")


Worksheets("Curvas").Select
Range("R:AA").ClearContents

archivo = R_IND_1 & "\" & N_IND_1 & ".CSV"

    With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + archivo, Destination:=Range("$R$1"))
       
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

COUNTULT = Worksheets("Curvas").Cells(Rows.Count, 18).End(xlUp).Row
Range("T1:T" + Format(COUNTULT, "0")) = Fecha_Valoración

archivo = R_IND_1 & "\" & N_IND_2 & ".CSV"

    With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + archivo, Destination:=Range("$R$" + Format(COUNTULT + 1, "0")))
       
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

COUNTULT1 = Worksheets("Curvas").Cells(Rows.Count, 18).End(xlUp).Row
Range("T" + Format(COUNTULT + 1, "0") & ":" & "T" + Format(COUNTULT1, "0")) = Fecha_Valoración_T_1

archivo = R_IND_1 & "\" & N_IND_3 & ".CSV"

    With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + archivo, Destination:=Range("$R$" + Format(COUNTULT1 + 1, "0")))
       
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
COUNTULT2 = Worksheets("Curvas").Cells(Rows.Count, 18).End(xlUp).Row
Range("T" + Format(COUNTULT1 + 1, "0") & ":" & "T" + Format(COUNTULT2, "0")) = Fecha_Valoración_T_2

Range("X:Z").Select
Selection.ClearContents
End Sub

Sub Limpiar_CONEXIONES()

'  Delete Additional Connections
        If ActiveWorkbook.Connections.Count > 0 Then
        For i = 1 To ActiveWorkbook.Connections.Count
        ActiveWorkbook.Connections.Item(1).Delete
        Next i
        End If
End Sub


Sub Limpiar_Datos_externos()
    Dim nm As Name
    Dim hoja

    For Each nm In ThisWorkbook.Names
        posicion_busqueda = InStr(1, nm.Name, "DatosExternos")
        If posicion_busqueda <> 0 Then
        nm.Delete
        End If
    Next nm

End Sub

Sub Generar_PDF_POCS_Relativo()

Call Generar_Carpeta_Diaria
Dim RUTA_GUARDAR, PORTAFOLIO As String
Dim objWord, objDoc, objWordDoc As Object

CONTROL_1 = Range("CONTROL_1")
PORTAFOLIO = Sheets("Formato de Apuestas PO-CS").Range("PORTAFOLIO")

RUTA_COLF_DIA = Sheets("Procedimiento").Range("E38").Value
ARCH_COLF_DIA = Sheets("Procedimiento").Range("D38").Value
BMK = ThisWorkbook.Name

Workbooks.Open Filename:=RUTA_COLF_DIA & ARCH_COLF_DIA
Workbooks(BMK).Activate

For i = 1 To 5
    If i = X Then
        Sheets("Formato de Apuestas PO-CS").Select
        Columns("G:K").Select
            Selection.EntireColumn.Hidden = True
        Columns("O:S").Select
            Selection.EntireColumn.Hidden = True
        Columns("U:U").Select
            Selection.EntireColumn.Hidden = False
    
            Sheets("Parametros").Select
        Range("b21").Select
        ActiveCell.Offset(i, 0).Select
        PORTAFOLIO = Format(ActiveCell.Value, "#0")
        
        Sheets("Formato de Apuestas PO-CS").Select
        Range(Format("B9", 0)) = PORTAFOLIO
        Sheets("Procedimiento").Range("c27:c29").Calculate
        Calculate
        CONTROL_1 = Range("CONTROL_1")
        
        If CONTROL_1 = True Then
        
            RUTA_GUARDAR = Worksheets("Procedimiento").Range("RUTA_GUARDAR")
            
            ActiveWindow.View = xlNormalView
            
            Sheets("Formato de Apuestas PO-CS").Activate
            Range("$B$1:$V$137").Select
            
            Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
            Set objWord = CreateObject("Word.Application")
            Set objDoc = objWord.Documents.Add
            
            With objDoc.PageSetup
                .LineNumbering.Active = False
                .Orientation = wdOrientPortrait
                .TopMargin = 1.27
                .BottomMargin = 1.27
                .LeftMargin = 1.27
                .RightMargin = 1.27
            End With
            
            objWord.Visible = True
            objWord.Selection.Paste
            objWord.Selection.TypeParagraph
            
            Sheets("Formato de Apuestas PO-CS").Activate
            Range("$W$1:$AO$84").Select
            
            Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
            
            objWord.Visible = True
            objWord.Selection.Paste
            objWord.Selection.TypeParagraph
           
            'tablas colf diario
            Workbooks(ARCH_COLF_DIA).Activate
            
            Sheets("Resumen diario").Select
            Range("Resumen_rojosazules_" & PORTAFOLIO).Select
            
            Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
            
            objWord.Visible = True
            objWord.Selection.Paste
            objWord.Selection.TypeParagraph
'            ''tabla vs ind
'            Range("Resumen_ind_rojosazules_" & PORTAFOLIO).Select
'
'            Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
'
'            objWord.Visible = True
'            objWord.Selection.Paste
'            objWord.Selection.TypeParagraph
            
            Workbooks(BMK).Activate
            
            objDoc.SaveAs ("C:\IT\MyFirstSave8.docx")
            objDoc.Close
            
            Set objWordDoc = objWord.Documents.Open("C:\IT\MyFirstSave8.docx")
            
            objWordDoc.ExportAsFixedFormat RUTA_GUARDAR, 17
            
            objWordDoc.Close
        
        'codigo incial de generación de pdf.
        '        RUTA_GUARDAR = Worksheets("Procedimiento").Range("RUTA_GUARDAR")
        '            Sheets("Formato de Apuestas PO-CS").Select
        '            ActiveSheet.PageSetup.PrintArea = "$B$1:$V$113,$W$1:$AO$84"
        '            ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        '                IgnorePrintAreas:=False
        '
        '            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        '               RUTA_GUARDAR, Quality:= _
        '                xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        '                OpenAfterPublish:=False
        
        Else
          
                Dim CU1_pr As Variant
                CU1_pr = Worksheets("Controles").Cells(Rows.Count, 3).End(xlUp).Row
        
        
                Sheets("Controles").Select
                Range("C" + Format(CU1_pr + 1, 0)).FormulaR1C1 = "Error generando PDF de " & PORTAFOLIO
        End If
    
    Else
    
        Sheets("Formato de Apuestas PO-CS").Select
        Columns("G:K").Select
            Selection.EntireColumn.Hidden = False
        Columns("O:S").Select
            Selection.EntireColumn.Hidden = False
        Columns("U:U").Select
            Selection.EntireColumn.Hidden = True
    
        Sheets("Parametros").Select
        Range("b21").Select
        ActiveCell.Offset(i, 0).Select
        PORTAFOLIO = Format(ActiveCell.Value, "#0")
        
        Sheets("Formato de Apuestas PO-CS").Select
        Range(Format("B9", 0)) = PORTAFOLIO
        Sheets("Procedimiento").Range("c27:c29").Calculate
        Calculate
        CONTROL_1 = Range("CONTROL_1")
        
        If CONTROL_1 = True Then
        
            RUTA_GUARDAR = Worksheets("Procedimiento").Range("RUTA_GUARDAR")
            
            ActiveWindow.View = xlNormalView
            
            Sheets("Formato de Apuestas PO-CS").Activate
            Range("$B$1:$V$109").Select
            
            Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
            Set objWord = CreateObject("Word.Application")
            Set objDoc = objWord.Documents.Add
            
            With objDoc.PageSetup
                .LineNumbering.Active = False
                .Orientation = wdOrientPortrait
                .TopMargin = 1.27
                .BottomMargin = 1.27
                .LeftMargin = 1.27
                .RightMargin = 1.27
            End With
            
            objWord.Visible = True
            objWord.Selection.Paste
            objWord.Selection.TypeParagraph
            
            Sheets("Formato de Apuestas PO-CS").Activate
            Range("$W$1:$AO$84").Select
            
            Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
            
            objWord.Visible = True
            objWord.Selection.Paste
            objWord.Selection.TypeParagraph
            
            Sheets("Formato de Apuestas PO-CS").Activate
            Range("$b$111:$k$138").Select
            
            Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
            
            objWord.Visible = True
            objWord.Selection.Paste
            objWord.Selection.TypeParagraph
            
            Sheets("Formato de Apuestas PO-CS").Activate
            Range("$FK$5:$FY$127").Select
            
            Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
            
            objWord.Visible = True
            objWord.Selection.Paste
            objWord.Selection.TypeParagraph
            
            
            Sheets("Formato de Apuestas PO-CS").Activate
            Range("$FK$128:$FY$209").Select
            
            Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
            
            objWord.Visible = True
            objWord.Selection.Paste
            objWord.Selection.TypeParagraph
            
            Sheets("Formato de Apuestas PO-CS").Activate
            Range("$ex$117:$Fi$144").Select
            
            Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
            
            objWord.Visible = True
            objWord.Selection.Paste
            objWord.Selection.TypeParagraph
            
            'tablas colf diario
            Workbooks(ARCH_COLF_DIA).Activate
            '' tabla vs bmks
            Sheets("Resumen diario").Select
            Range("Resumen_rojosazules_" & PORTAFOLIO).Select
            
            Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
            
            objWord.Visible = True
            objWord.Selection.Paste
            objWord.Selection.TypeParagraph
'            ''tabla vs ind
'            Range("Resumen_ind_rojosazules_" & PORTAFOLIO).Select
'
'            Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
'
'            objWord.Visible = True
'            objWord.Selection.Paste
'            objWord.Selection.TypeParagraph
            
            Workbooks(BMK).Activate
            
            objDoc.SaveAs ("C:\IT\MyFirstSave8.docx")
            objDoc.Close
            
            Set objWordDoc = objWord.Documents.Open("C:\IT\MyFirstSave8.docx")
            
            objWordDoc.ExportAsFixedFormat RUTA_GUARDAR, 17
            
            objWordDoc.Close
        
        'codigo incial de generación de pdf.
        '        RUTA_GUARDAR = Worksheets("Procedimiento").Range("RUTA_GUARDAR")
        '            Sheets("Formato de Apuestas PO-CS").Select
        '            ActiveSheet.PageSetup.PrintArea = "$B$1:$V$113,$W$1:$AO$84"
        '            ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        '                IgnorePrintAreas:=False
        ''
        '            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        '               RUTA_GUARDAR, Quality:= _
        '                xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        '                OpenAfterPublish:=False
        Else
          
                Dim CU1 As Variant
                CU1 = Worksheets("Controles").Cells(Rows.Count, 3).End(xlUp).Row
        
        
                Sheets("Controles").Select
                Range("C" + Format(CU1 + 1, 0)).FormulaR1C1 = "Error generando PDF de " & PORTAFOLIO
        End If
    End If
 Next
 
    Columns("U:U").Select
        Selection.EntireColumn.Hidden = False

Workbooks(ARCH_COLF_DIA).Close SaveChanges:=False


Call Generar_PDF_hoja_Voluntarias
Call Generar_PDF_VOLUNTARIAS

  
End Sub

Sub Generar_PDF_VOLUNTARIAS()


Dim RUTA_GUARDAR_VOL As String

Dim CONTROL_1 As String


RUTA_COLF_DIA = Sheets("Procedimiento").Range("E39").Value
ARCH_COLF_DIA = Sheets("Procedimiento").Range("D39").Value
BMK = ThisWorkbook.Name

Workbooks.Open Filename:=RUTA_COLF_DIA & ARCH_COLF_DIA
Workbooks(BMK).Activate

PORTAFOLIO2 = Range("PORTAFOLIO2")

For i = 6 To 44

    Sheets("Parametros").Select
    Range("B21").Select
    ActiveCell.Offset(i, 0).Select
    PORTAFOLIO2 = ActiveCell.Value
    
    Sheets("Formato de Apuestas").Select
    Range("B9") = PORTAFOLIO2
    
    Sheets("Procedimiento").Range("D51:D52").Calculate
    Calculate
    CONTROL_1 = Range("CONTROL_1")
    
    If CONTROL_1 = True Then
    
        If i = 7 Or i = 8 Or i = 9 Or i = 10 Or i = 11 Then
        
                
        
                RUTA_GUARDAR = Worksheets("Procedimiento").Range("RUTA_GUARDAR_VOL")
                
                ActiveWindow.View = xlNormalView
                
                Sheets("Formato de Apuestas").Activate
                Range("$A$1:$J$98").Select
                
                Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
                Set objWord = CreateObject("Word.Application")
                Set objDoc = objWord.Documents.Add
                
                With objDoc.PageSetup
                    .LineNumbering.Active = False
                    .Orientation = wdOrientPortrait
                    .TopMargin = 1.27
                    .BottomMargin = 1.27
                    .LeftMargin = 1.27
                    .RightMargin = 1.27
                End With
                
                objWord.Visible = True
                objWord.Selection.Paste
                objWord.Selection.TypeParagraph
                
                'tablas colf diario
                Workbooks(ARCH_COLF_DIA).Activate
                
                Sheets("Resumen").Select
                Range("rojosazules_" & PORTAFOLIO2).Select
                
                Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
                
                objWord.Visible = True
                objWord.Selection.Paste
                objWord.Selection.TypeParagraph
    '            ''tabla vs ind
    '            Range("Resumen_ind_rojosazules_" & PORTAFOLIO).Select
    '
    '            Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    '
    '            objWord.Visible = True
    '            objWord.Selection.Paste
    '            objWord.Selection.TypeParagraph
                
                Workbooks(BMK).Activate
                
                objDoc.SaveAs ("C:\IT\MyFirstSave8.docx")
                objDoc.Close
                
                Set objWordDoc = objWord.Documents.Open("C:\IT\MyFirstSave8.docx")
                
                objWordDoc.ExportAsFixedFormat RUTA_GUARDAR, 17
                
                objWordDoc.Close
        Else
                'codigo incial de generación de pdf.
                RUTA_GUARDAR = Worksheets("Procedimiento").Range("RUTA_GUARDAR_VOL")
                Sheets("Formato de Apuestas").Select
                ActiveSheet.PageSetup.PrintArea = "$A$1:$J$98"
                '   ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
                    IgnorePrintAreas:=False
                        
                ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
                    RUTA_GUARDAR, Quality:= _
                    xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
                    OpenAfterPublish:=False

        End If
    Else
     
        Dim CU As Variant
        CU = Worksheets("Controles").Cells(Rows.Count, 2).End(xlUp).Row
            
            
        Sheets("Controles").Select
        Range("B" + Format(CU + 1, 0)).FormulaR1C1 = "Error generando PDF de " & PORTAFOLIO2
    End If

Next

Workbooks(ARCH_COLF_DIA).Close SaveChanges:=False

End Sub

Sub Generar_PDFCP()

Dim RUTA_GUARDAR As String

Dim CONTROL_2 As String

'CONTROL_2 = Range("CONTROL_2")
PORTAFOLIO = Range("PORTAFOLIO")

Sheets("Controles").Range("b23:b50").ClearContents

Sheets("Parametros").Select
Range("C50").Select
PORTAFOLIO = ActiveCell.Value

Sheets("Formato de Apuestas PO CS").Select
Range("B12") = PORTAFOLIO
Sheets("Procedimiento").Range("c24:c25").Calculate
Calculate
CONTROL_2 = Range("CONTROL_2")

If CONTROL_2 = True Then

RUTA_GUARDAR = Worksheets("Procedimiento").Range("RUTA_GUARDAR")
        Sheets("Formato de Apuestas PO CS").Select
        ActiveSheet.PageSetup.PrintArea = "$A$1:$l$79,$n$1:$X$80,$CU$1:$DF$107"
 '       ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
            IgnorePrintAreas:=False
            
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
           RUTA_GUARDAR, Quality:= _
            xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
 End If
  
            
End Sub



Private Sub btnMensaje_Click()
 For Each hoja In ActiveWorkbook.Sheets
 If hoja.AutoFilterMode Then
 hoja.Range(“A1”).AutoFilter
 'Me.btnMensaje.Caption = "Actvar Autofiltros"
Else
 hoja.Range("A1").AutoFilter
 'Me.btnMensaje.Caption = "Desactivar Autofiltros"
End If
 Next hoja
 End Sub


Sub Guardar()

ActiveWorkbook.Save
Sheets("Procedimiento").Select
Dim RUTA_PUBLICACION As String
Dim N_PUBLICACION As String


RUTA_PUBLICACION = Range("RUTA_PUBLICACION")
N_PUBLICACION = Range("N_PUBLICACION")

ThisWorkbook.SaveAs Filename:=RUTA_PUBLICACION & "\" & N_PUBLICACION & ".xlsb", FileFormat:=xlExcel12, CreateBackup:=False
Call Copiar_paste
End Sub

Sub Copiar_paste()


    
    Sheets("Controles").Select
    
    Dim wSheet As Worksheet
    For Each wSheet In Worksheets
    
       
         Range("a1:ca1048576").Select
         Selection.Copy
         Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
         ActiveSheet.Next.Select
         
         If ActiveSheet.Name = "Formato de Apuestas" Then
         
         ActiveSheet.Next.Select
                          
         ElseIf ActiveSheet.Visible = False Then
         ActiveSheet.Next.Select
                   
         End If
         
         If ActiveSheet.Name = "Formato de Apuestas PO-CS" Then
         
         ActiveSheet.Next.Select
         
         ElseIf ActiveSheet.Visible = False Then
         ActiveSheet.Next.Select
         
         End If
         
          If ActiveSheet.Name = "Formato de Apuestas" Then
          ActiveSheet.Next.Select
          
         ElseIf ActiveSheet.Visible = False Then
         
         ActiveSheet.Next.Select
         End If

         
         On Error GoTo Terminar2:
         
         
    
        
    Next wSheet
    
   
Terminar2:
Sheets("Procedimiento").Select
End Sub

Sub GenerarCarpeta()

 
      Dim C_M As String, C_P As String, N_M2 As String, N_P As String
      
     
      
              
        Sheets("Parametros").Range("T14:T20").Calculate
       
        C_M = Range("C_M")
        C_P = Range("C_P")
        N_M2 = Range("N_M2")
      
          
     'Verificamos si la carpeta existe ya...
     If Dir(C_M, vbDirectory) <> "" Then
     'Comprueba que la carpeta no existe para crearla.
     If Dir(C_M & "\" & N_M2, vbDirectory) = "" Then MkDir C_M & "\" & N_M2
     'MkDir se emplea para crear un directorio/carpeta.
     'Si no se especifica la unidad de disco, el directorio/carpeta se crea en la unidad actual.
     'Comprueba que la carpeta no existe para crearla.
     If Dir(C_P & "\" & N_M2, vbDirectory) = "" Then
     If Dir(C_P & "\" & N_M2, vbDirectory) = "" Then MkDir C_P & "\" & "Consolidado"
    
     'Comprueba que la carpeta no existe para crearla.
     'If Dir(C_P, vbDirectory) <> "" Then
    'If Dir(C_P & "\" & N_M, vbDirectory) = "" Then MkDir C_P & "\" & N_M
     'MkDir se emplea para crear un directorio/carpeta.
     'Si no se especifica la unidad de disco, el directorio/carpeta se crea en la unidad actual.
     
     End If
     End If
 
      For i = 1 To 20
      
           
        Sheets("Parametros").Select
        Range("b21").Select
        ActiveCell.Offset(i, 0).Select
        Sheets("Parametros").Range("S13") = Format(ActiveCell.Value, "#0")
        
        
        Sheets("Parametros").Range("T14:T20").Calculate
     
      N_M2 = Range("N_M2")
      N_P = Range("N_P")
     
     
     If Dir(C_P, vbDirectory) <> "" Then
     'Comprueba que la carpeta no existe para crearla.
     If Dir(C_P & "\" & N_P, vbDirectory) = "" Then MkDir C_P & "\" & N_P
     'MkDir se emplea para crear un directorio/carpeta.
     'Si no se especifica la unidad de disco, el directorio/carpeta se crea en la unidad actual.
     
     End If
   
     
 Next
End Sub



Sub Generar_Carpeta_Diaria()

 
      Dim C_D As String, N_D As String
      
     
      
              
        Sheets("Parametros").Range("T23:T29").Calculate
       
        C_D = Range("C_D")
        N_D = Range("N_D")
      
          
     'Verificamos si la carpeta existe ya...
     If Dir(C_D, vbDirectory) <> "" Then
     'Comprueba que la carpeta no existe para crearla.
     If Dir(C_D & "\" & N_D, vbDirectory) = "" Then MkDir C_D & "\" & N_D
     'MkDir se emplea para crear un directorio/carpeta.
     'Si no se especifica la unidad de disco, el directorio/carpeta se crea en la unidad actual.
     
    
          End If
 
     
End Sub

Sub PDF()

Generar_Carpeta_Diaria
MsgBox "Se generó la carpeta del dia de hoy con exito"
Generar_PDF
'Generar_PDFCP
Generar_PDF2
Generar_PDFVoluntarias
Sheets("Controles").Select
Range("E23").Select

End Sub


Sub Copiar_paste2()


    
    Sheets("Procedimiento").Select
    
    Dim wSheet As Worksheet
    For Each wSheet In Worksheets
    
       
         Range("a1:ca1048576").Select
         Selection.Copy
         Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
         ActiveSheet.Next.Select
         
                 
         On Error GoTo Terminar2:
         
         
    
        
    Next wSheet
    
   
Terminar2:
Sheets("Procedimiento").Select
End Sub
Sub Generar_PDF_hoja_Voluntarias()

Dim RUTA_GUARDAR As String

Dim CONTROL_2 As String

CONTROL_2 = Range("CONTROL_2")
CONTROL_VOLUNTARIAS = Range("CONTROL_VOLUNTARIAS")


Sheets("Voluntarias").Select
RUTA_GUARDAR = Worksheets("Voluntarias").Range("infVol")

If CONTROL_2 = True And CONTROL_VOLUNTARIAS = True Then

        Sheets("Voluntarias").Select
        ActiveSheet.PageSetup.PrintArea = "$A$1:$k$87"

        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
           RUTA_GUARDAR, Quality:= _
            xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
            
            Else
            
            MsgBox ("Revisar la hoja voluntarias")
            
 End If
  
            
End Sub













