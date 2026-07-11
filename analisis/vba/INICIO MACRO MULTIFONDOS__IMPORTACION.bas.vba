Attribute VB_Name = "IMPORTACION"
'Jose Orlando Bobadilla Fuentes
'Ingeniero Financiero
'Noviembre de 2015
Dim Carpeta_597 As String
Dim Carpeta_620 As String
Dim Carpeta_Ind As String

Dim Año As String
Dim Año_Completo As String
Dim Mes As String
Dim dia As String



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

Sub Importar_Informes()



Application.ScreenUpdating = False
Worksheets("597").Select
Range("A:AZ").ClearContents
Worksheets("620").Select
Range("A:AZ").ClearContents

Dim FECHAVALORACION As String
Dim N_597 As String
Dim R_597 As String
Dim N_620 As String
Dim R_620 As String
Dim Archivo As String
Dim N_IND As String
Dim R_IND As String
Dim N_LIMITES As String
Dim R_LIMITES As String
Dim N_VARIABLE As String
Dim R_VARIABLE As String
Dim N_BLOOMBERG As String
Dim R_BLOOMBERG As String






Fecha_Actualizacion = Range("Fecha_Actualizacion")
N_597 = Range("N_597")
R_597 = Range("R_597")
N_620 = Range("N_620")
R_620 = Range("R_620")
N_IND = Range("N_IND")
R_IND = Range("R_IND")
N_LIMITES = Range("N_LIMITES")
R_LIMITES = Range("R_LIMITES")
N_VARIABLE = Range("N_VARIABLE")
R_VARIABLE = Range("R_VARIABLE")
N_BLOOMBERG = Range("N_BLOOMBERG")
R_BLOOMBERG = Range("R_BLOOMBERG")

   
    
      Worksheets("597").Activate
     

'descarga el archivo 597 vigente
 Archivo = R_597 & "\" & N_597 & ".CSV"

    With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + Archivo, Destination:=Range("$A$1"))
       
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
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
        
        Columns("AM:AM").Copy
        Columns("H:H").PasteSpecial Paste:=xlPasteValues
        
    End With
    
 'Range("2:2").Rows.Delete
 Range("AU:AU").Columns.Delete
 Range("A2:AT2").AutoFilter
    ActiveSheet.Range("A2:AT2").AutoFilter Field:=3, Criteria1:="=AODF5 AH BBH USD              "
        
     Dim CU597 As Variant
        CU597 = Worksheets("597").Cells(Rows.Count, 1).End(xlUp).Row
        
    Range("A3").Select
    Range("A4:AT" + Format(CU597, 0)).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.EntireRow.Delete
    Selection.AutoFilter
    
''descarga el archivo 620 vigente
Worksheets("620").Select
Archivo2 = R_620 & "\" & N_620 & ".CSV"

With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + Archivo2, Destination:=Range("$A$1"))
        
        
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
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
   Sheets("620").Select
   Range("AA:AC").Select
   Selection.Delete
   If Range("V2") = "   Valor CVA/DVA" Then
      Columns("U:V").Select
 Selection.Delete Shift:=xlToLeft

                             End If
 Worksheets("CONTROL").Select

 Range("B1:F37").Calculate
 Range("c3").Calculate
 
'descarga el archivo IND vigente
    
    Worksheets("IND2").Select
    Range("A1:M1000000").Select
    Selection.ClearContents
    
    Archivo4 = R_IND & "\" & N_IND & ".CSV"

With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + Archivo4, Destination:=Range("$A$1"))

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
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = True
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
   Worksheets("PROCEDIMIENTO").Select
  
End Sub

Sub Importar()

    Limpiar_Datos_externos2
    Importar_Informes
    Limpiar_Datos_externos2
    'Importar_CSA
    CSA_balance
    Limpiar_Datos_externos2
    
    Calculate
    

End Sub
Sub Limpiar_Datos_externos2()
'  Delete Additional Connections
        If ActiveWorkbook.Connections.Count > 0 Then
        For i = 1 To ActiveWorkbook.Connections.Count
        ActiveWorkbook.Connections.Item(1).Delete
        Next i
        End If
End Sub

Sub HOJA_IND()


    Worksheets("VECTOR DE PRECIOS").Select
    ActiveWorkbook.RefreshAll
    
    Dim COUNTULT7 As String
    Range("U3").FormulaR1C1 = "=VLOOKUP(VLOOKUP(RC[-2],C[-17]:C[-15],3,FALSE),C[3]:C[4],2,FALSE)"
   
    COUNTULT7 = Worksheets("VECTOR DE PRECIOS").Cells(Rows.Count, 19).End(xlUp).Row
    Range("U3").Copy
    
    Range(("U3:U" + Format(COUNTULT7 - 1, "0"))).Select
    Selection.PasteSpecial xlPasteFormulas
    Calculate
    
           
    Worksheets("IND").Select
    Range("A2:D1048576").Select
    Selection.ClearContents
    
    Range("o3:R1048576").Select
    Selection.ClearContents
    
    

    Dim N_VARIABLE As String
    Dim R_VARIABLE As String
    Dim N_BLOOMBERG As String
    Dim R_BLOOMBERG As String
    Dim N_VECTOR As String
    Dim R_VECTOR As String
    
    N_VARIABLE = Range("N_VARIABLE")
    R_VARIABLE = Range("R_VARIABLE")
    N_BLOOMBERG = Range("N_BLOOMBERG")
    R_BLOOMBERG = Range("R_BLOOMBERG")
    N_VECTOR = Range("N_VECTOR")
    R_VECTOR = Range("R_VECTOR")
    
    Dim workbookorigen As Workbook
    Dim wsOrigen As Excel.Worksheet, _
    wsDestino As Excel.Worksheet, _
    rngOrigen As Excel.Range, _
    rngDestino As Excel.Range, _
    VECTOR As String
    
    
    narchivo = 1

 Worksheets("Procedimiento").Activate
      
 ''trae la informacion del Vector de precios
 


      

      Set workbookorigen = Workbooks.Open(R_VECTOR & "\" & N_VECTOR & ".xlsm", , , , "INVERSIONES", "INVERSIONES")
      
          
          ThisWorkbook.Activate
               
      Set wsOrigen = workbookorigen.Worksheets("Vector")
      Set wsDestino = Worksheets("IND")
      
      Const celdaorigen = "A3:D3"
      
      
             
      Set rngOrigen = wsOrigen.Range(celdaorigen)
      wsOrigen.Activate
      rngOrigen.Select
      Range(Selection, Selection.End(xlDown)).Copy


       Windows("INICIO MACRO MULTIFONDOS.xlsm").Activate
       Sheets("IND").Select
       Range("a2").PasteSpecial xlPasteValues


                  
          
        Windows(N_VECTOR & ".xlsm").Activate
        Sheets("Vector").Select
        Range("F3").Select
        Range(Selection, Selection.End(xlDown)).Copy
    
       Windows("INICIO MACRO MULTIFONDOS.xlsm").Activate
       Sheets("IND").Select
       Range("O3").PasteSpecial xlPasteValues
       
       
       
       Windows(N_VECTOR & ".xlsm").Activate
        Sheets("Vector").Select
        Range("i3:K3").Select
        Range(Selection, Selection.End(xlDown)).Copy
    
       Windows("INICIO MACRO MULTIFONDOS.xlsm").Activate
       Sheets("IND").Select
       Range("p3").PasteSpecial xlPasteValues
             
         Call Limpiar_Datos_externos2
          workbookorigen.Close
          Call Limpiar_Datos_externos
                    
          
         ''TRAE LA INFORMACION DE LA HOJA vector de precios
          
          
          Dim COUNTULT3 As String
    Worksheets("VECTOR DE PRECIOS").Select
    ActiveWorkbook.RefreshAll
    COUNTULT3 = Worksheets("VECTOR DE PRECIOS").Cells(Rows.Count, 19).End(xlUp).Row
    
    
    Range(("S3:T" + Format(COUNTULT3 - 1, "0"))).Copy
    ''Range(Selection, Selection.End(xlDown)).Copy
    
    Dim COUNTULT4 As String
    Dim COUNTULT5 As String
    Worksheets("IND").Select
    COUNTULT4 = Worksheets("IND").Cells(Rows.Count, 1).End(xlUp).Row
    Range(("A" + Format(COUNTULT4 + 1, "0"))).PasteSpecial xlPasteValues
    COUNTULT5 = Worksheets("IND").Cells(Rows.Count, 2).End(xlUp).Row
    Range(("C" + Format(COUNTULT4 + 1, "0"))).FormulaR1C1 = "PORFIN"
    Range(("C" + Format(COUNTULT4 + 1, "0"))).Copy
    Range((("C" + Format(COUNTULT4 + 1, "0"))), ("C" + Format(COUNTULT5, "0"))).Select
     Selection.PasteSpecial xlPasteValues
    
    
    Worksheets("VECTOR DE PRECIOS").Select
    Range(("U3:U" + Format(COUNTULT3 - 1, "0"))).Copy
    Worksheets("IND").Select
    Range(("D" + Format(COUNTULT4 + 1, "0"))).PasteSpecial xlPasteValues
    
    
    

End Sub

Sub GENERAR_TXT_VECTORPRECIOS()
Application.ScreenUpdating = False
    R_EXPO_VECTOR = Range("R_EXPO_VECTOR")
    N_EXPO_VECTOR = Range("N_EXPO_VECTOR")
    Mes = Range("MES")
'CREA UNA HOJA LLAMADA RECOPILACION
Sheets.Add
ActiveSheet.Name = "VECTOR_TXT"

Sheets("IND").Select
Dim COUNTULT As String
    COUNTULT = Worksheets("IND").Cells(Rows.Count, 1).End(xlUp).Row

'CONCATENA LAS COLUMNAS SEPARADAS POR ;
Range("U1").FormulaR1C1 = "=+RC[-20]&"";""&RC[-19]&"";""&RC[-18]&"";""&RC[-17]"
Range("U1").Copy Destination:=Range("U1:U" + Format(COUNTULT, "0"))

Calculate

Worksheets("VECTOR_TXT").Range("A1:A" + Format(COUNTULT, "0")).Value = Worksheets("IND").Range("U1:U" + Format(COUNTULT, "0")).Value

  Sheets("VECTOR_TXT").Select
  ActiveSheet.Move
  Range("A1").Select

'GUARDA LA INFORMACION CONCATENADA EN FORMATO TXT
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:=R_EXPO_VECTOR & "\" & N_EXPO_VECTOR & ".txt", _
        FileFormat:=xlText, CreateBackup:=False
    ActiveWorkbook.Close
    Application.ScreenUpdating = True
'LIMPIA LA INFORMACION DE LA CELDA O
    Worksheets("IND").Select
    Range("U:U").Select
    Selection.ClearContents
    Worksheets("PROCEDIMIENTO").Select

    
End Sub



Sub GENERAR_TXT_DURACION()
Application.ScreenUpdating = False
    R_EXPO_DUR = Range("R_EXPO_DUR")
    N_EXPO_DUR = Range("N_EXPO_DUR")
    Mes = Range("MES")
'CREA UNA HOJA LLAMADA RECOPILACION
Sheets.Add
ActiveSheet.Name = "VECTOR_DUR"

Sheets("IND").Select
Dim COUNTULT As String
    COUNTULT = Worksheets("IND").Cells(Rows.Count, 1).End(xlUp).Row

'CONCATENA LAS COLUMNAS SEPARADAS POR ;
Range("V1").FormulaR1C1 = "=+RC[-7]&"";""&RC[-6]&"";""&RC[-5]"
Range("V1").Copy Destination:=Range("V1:V" + Format(COUNTULT, "0"))

Calculate

Worksheets("VECTOR_DUR").Range("A1:A" + Format(COUNTULT, "0")).Value = Worksheets("IND").Range("V1:V" + Format(COUNTULT, "0")).Value

  Sheets("VECTOR_DUR").Select
  ActiveSheet.Move
  Range("A1").Select

'GUARDA LA INFORMACION CONCATENADA EN FORMATO TXT
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:=R_EXPO_DUR & "\" & N_EXPO_DUR & ".txt", _
        FileFormat:=xlText, CreateBackup:=False
    ActiveWorkbook.Close
    Application.ScreenUpdating = True
'LIMPIA LA INFORMACION DE LA CELDA O
    Worksheets("IND").Select
    Range("V:V").Select
    Selection.ClearContents
    Worksheets("PROCEDIMIENTO").Select
Call guardar
    
End Sub

Sub guardar()
Dim R_GUARDAR As String
Dim Nom_GUARDAR As String
Dim extension As String
Dim N_MACRO As String
Dim R_MACRO As String

Calculate
Sheets("PROCEDIMIENTO").Select
R_GUARDAR = Range("R_GUARDAR")
Nom_GUARDAR = Range("Nom_GUARDAR")
N_MACRO = Range("N_MACRO")
R_MACRO = Range("R_MACRO")

Sheets("CONTROL").Select
Range("B19").Select

If Range("B19") = "GUARDAR" Then

ThisWorkbook.Save
ThisWorkbook.SaveAs Filename:=R_GUARDAR & "\" & Nom_GUARDAR & ".xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False


Dim workbookorigen As Workbook
Set workbookorigen = Workbooks.Open(R_MACRO & "\" & N_MACRO)

          ThisWorkbook.Activate
Windows(Nom_GUARDAR & ".xlsm").Activate
ThisWorkbook.Close
End If
End Sub
