Attribute VB_Name = "Module1"
Sub cargardatos()
Dim oProgress As New frm_lcf_ProgressBar
oProgress.Initialize 22, 2, "Progreso"
oProgress.Show 0
Dim Ultimac, Ultimaf, fila, Columna, m, k, Fondop
Dim SALIDA As String, tipopo As String
Dim excelapp As Excel.Application
Dim excelwrk As Excel.Workbook
Worksheets("T597").Activate
Module4.copia 'Hace una copia de respaldo del archivo 597 ahi es donde se va a trabajar
oProgress.Increase 1
Module4.portafolio
oProgress.Increase 1
Module2.Ventas
oProgress.Increase 1
Module2.Compras
oProgress.Increase 1
Module3.PORTAFOLIOS
oProgress.Increase 1
Módulo7.Copy
Unload oProgress
Sheets("CONTROL").Select
ActiveWorkbook.RefreshAll

CxC
Call Limpiar_Datos_externos2
Call HOJA_IND
Call ReposNegativos
Call Limpiar_Datos_externos2
'Call guardar
End Sub

    
Sub CxC()
    
    '30/03/2015
    'Ing Financiero. Jhonatan Franco
    'Ing Financiero. Fabio Escala
    'Analisis Cuantitativo
    
    'Este modulo anexado a la  macro "Inicio Macro MultiFondos" es para tener de manera simultanea a la carga de datos
    'los titulos en transito para cada uno de los portafolio con su moneda respectiva. La formula BuscarV trae la moneda
    'del diccionario. El diccionario esta ubicado en la hoja "Resumen CV" donde tambien se encuentra el cuadro que relaciona
    'las Compras y Venta para cada uno de los portafolios que van en Cxc CXP y T620 de cada uno de los Benchmark.
   
    
    Columns("W:W").Select
    Selection.ClearContents

    Sheets("T620").Select
    Range("A1048576").Select
    Selection.End(xlUp).Select
    ULTIMA_FILA = Selection.Row
       
    
    Range("Y2").FormulaR1C1 = "Clase de Moneda"
    Range("Y4").FormulaR1C1 = "=VLOOKUP(RC[-21],'Resumen CV'!C[6]:C[7],2,FALSE)"
    Range("Y4").Select
    Selection.Copy
    
    Range("Y4:Y" + Format(ULTIMA_FILA, "0")).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Calculate
     
End Sub

Sub ReposNegativos()

Sheets("597").Select
Range("A1048576").Select
Selection.End(xlUp).Select
ULTIMA_FILA = Selection.Row
Range("C3").Select

For i = 4 To ULTIMA_FILA
    If Trim(Cells(i, 3).Value) = "REPO PASIVO" Then
        Cells(i, 14).Value = (Cells(i, 14).Value) * -1
        Cells(i, 18).Value = (Cells(i, 18).Value) * -1
        Cells(i, 21).Value = (Cells(i, 21).Value) * -1
        Cells(i, 23).Value = (Cells(i, 23).Value) * -1
        Cells(i, 25).Value = (Cells(i, 25).Value) * -1
    End If

Next


Sheets("T597").Select
Range("A1048576").Select
Selection.End(xlUp).Select
ULTIMA_FILA = Selection.Row
Range("C3").Select

For i = 4 To ULTIMA_FILA
    If Trim(Cells(i, 3).Value) = "REPO PASIVO" Then
        Cells(i, 14).Value = (Cells(i, 14).Value) * -1
        Cells(i, 18).Value = (Cells(i, 18).Value) * -1
        Cells(i, 21).Value = (Cells(i, 21).Value) * -1
        Cells(i, 23).Value = (Cells(i, 23).Value) * -1
        Cells(i, 25).Value = (Cells(i, 25).Value) * -1
    End If

Next

Sheets("PORTAFOLIOS").Select
Range("A1048576").Select
Selection.End(xlUp).Select
ULTIMA_FILA = Selection.Row
Range("C3").Select

For i = 4 To ULTIMA_FILA
    If Trim(Cells(i, 6).Value) = "REPO PASIVO" Then
        Cells(i, 11).Value = (Cells(i, 11).Value) * -1
        Cells(i, 12).Value = (Cells(i, 12).Value) * -1
    End If
    If Cells(i, 3) = "COC04CB00137" Then
        Cells(i, 9) = 46992
    End If
Next
End Sub
