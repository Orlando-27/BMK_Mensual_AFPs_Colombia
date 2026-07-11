Attribute VB_Name = "Módulo3"
Sub Generar_Benchmark()

    ' Desfiltar todas las hojas
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.AutoFilterMode Then
            ws.AutoFilterMode = False
        End If
    Next ws

    Dim N_PUBLICACION As String
    Dim RUTA_PUBLICACION As String

    N_PUBLICACION = Range("N_PUBLICACION")
    RUTA_PUBLICACION = Range("RUTA_PUBLICACION")

    Sheets("Procedimiento").Select
    Range("Fecha_Valoración").Copy
    Range("Fecha_Valoración").PasteSpecial Paste:=xlPasteValues

    Copiar_paste

    Sheets("Procedimiento").Select

    ActiveWorkbook.SaveAs Filename:=RUTA_PUBLICACION & "\" & N_PUBLICACION & ".xlsb", FileFormat:=xlExcel12, CreateBackup:=False

    Windows(N_PUBLICACION & ".xlsb").Activate

'''''''''''VECTOR DURACION SWAPS ''''''''''''''''''''''
    
    Call Vector_duraciones_SWAPs

    Call Generar_Portafolios
    
End Sub

Sub Generar_Portafolios()

    Dim N_IND As String
    Dim R_IND As String
    Dim PORTAFOLIO As String
    Dim N_PUBLICACION As String

    N_PUBLICACION = Range("N_PUBLICACION")

    For i = 1 To 5

        Windows(N_PUBLICACION & ".xlsb").Activate
        Sheets("Parametros").Select
        Range("B21").Offset(i, 0).Select

        PORTAFOLIO = Format(ActiveCell.Value, "#0")
        Sheets("Parametros").Range("S21") = PORTAFOLIO

        Sheets("Formato de Apuestas").Range("B9") = PORTAFOLIO
        Sheets("Formato de Apuestas PO-CS").Range("B9") = PORTAFOLIO

        Calculate

        N_IND = Range("N_IND")
        R_IND = Range("R_IND")

        Calculate

        Sheets(Array("Procedimiento", "Colfondos", "Benchmark", "Formato de Apuestas PO-CS", "Formato de Apuestas", "SWAP", "Fwd Colfondos", "Fwd Industria", "FutLoc Industria", "FutInt Industria", "Curvas", "NE", "VECTOR DE PRECIOS", "Parametros", "Fut Colf")).Copy
        Copiar_paste2

        ActiveWorkbook.SaveAs Filename:=R_IND & "\" & N_IND & ".xlsb", FileFormat:=xlExcel12, CreateBackup:=False
        ActiveWorkbook.Close

    Next i

    Call Generar_Portafolios_Voluntarias
    
    Sheets("Procedimiento").Select

End Sub

Sub Generar_Portafolios_Voluntarias()


    Dim N_IND_VOL As String
    Dim R_IND_VOL As String
    Dim N_PUBLICACION As String
    
    N_PUBLICACION = Range("N_PUBLICACION")
  
For i = 6 To 45

 Windows(N_PUBLICACION & ".xlsb").Activate
    Sheets("Parametros").Select
    Range("b21").Select
    ActiveCell.Offset(i, 0).Select
    
'    If ActiveCell.Value = "PO" Or ActiveCell.Value = "PC" Or ActiveCell.Value = "PR" Or ActiveCell.Value = "PM" Or ActiveCell.Value = "CS" Or ActiveCell.Value = "CP" Then
    PORTAFOLIO2 = Format(ActiveCell.Value, "#0")
'    PORTAFOLIO = Format(ActiveCell.Value, "#0")
    Sheets("Parametros").Range("S21") = Format(ActiveCell.Value, "#0")
    
    Sheets("Formato de Apuestas").Select
    Range(Format("B9", 0)) = PORTAFOLIO2
    
'    Sheets("Formato de Apuestas PO CS").Select
'    Range(Format("B12", 0)) = PORTAFOLIO
'    Sheets("Procedimiento").Range("c27:c29").Calculate
'    Calculate
'
'    Else
'    PORTAFOLIO2 = Format(ActiveCell.Value, "#0")
'    Sheets("Parametros").Range("S21") = Format(ActiveCell.Value, "#0")
'
'    Sheets("Formato de Apuestas").Select
'    Range(Format("B9", 0)) = PORTAFOLIO2
Calculate
'    Calculate
'
'    End If
    
     N_IND_VOL = Range("N_IND_VOL")
    R_IND_VOL = Range("R_IND_VOL")
    
    Calculate
    
    'If CONTROL_3 = True Then
    
    If i = 6 Then
    
    Sheets(Array("Procedimiento", "Colfondos", "Formato de Apuestas", "Formato de Apuestas CP", "SWAP", "Fwd Colfondos", "Curvas", "NE", "VECTOR DE PRECIOS", "Parametros")).Copy
    Copiar_paste2
    
    Else
    
    Sheets(Array("Procedimiento", "Colfondos", "Benchmark", "Formato de Apuestas PO-CS", "Formato de Apuestas", "SWAP", "Fwd Colfondos", "Fwd Industria", "Curvas", "NE", "VECTOR DE PRECIOS", "Parametros")).Copy
    Copiar_paste2
    
    End If
    
    'Sheets(Array("Procedimiento", "Colfondos", "Benchmark", "Formato de Apuestas", "Formato de Apuestas PO CS", "SWAP", "Fwd Colfondos", "VECTOR DE PRECIOS", "Parametros")).Select
   ' Sheets(Array("Procedimiento", "Colfondos", "Benchmark", "Formato de Apuestas", "Formato de Apuestas PO CS", "SWAP", "Fwd Colfondos", "VECTOR DE PRECIOS", "Parametros")).Copy
    
    
    ActiveWorkbook.SaveAs Filename:=R_IND_VOL & "\" & N_IND_VOL & ".xlsb", FileFormat:=xlExcel12, CreateBackup:=False
    ActiveWorkbook.Close
    
    'Else
    
    'Dim CU As Variant
        'CU = Worksheets("Controles").Cells(Rows.Count, 5).End(xlUp).Row
        
        
        'Sheets("Controles").Select
       ' Range("E" + Format(CU + 1, 0)).FormulaR1C1 = "Error generando portafolio de " & PORTAFOLIO
    
    'End If
Next
    
Sheets("Procedimiento").Select

End Sub

Sub Borrar_Exceles()

    Dim fso As Object
    Dim rutaPortafolios As String
    Dim nombreArchivo As String
    Dim i As Integer
    Dim filaRegistro As Integer
    Dim hojaProc As Worksheet
    Dim hojaParam As Worksheet

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set hojaProc = Sheets("Procedimiento")
    Set hojaParam = Sheets("Parametros")
    filaRegistro = 91 ' Comenzar en la fila 91

    ' Limpiar registros anteriores en C91:D200
    hojaProc.Range("C91:D200").ClearContents

    For i = 6 To 40
        ' Pegar valor como texto en C88 (POR_BORR_EXC)
        hojaProc.Range("POR_BORR_EXC").Value = "'" & hojaParam.Range("B21").Offset(i, 0).Value
        ' Forzar recálculo de la hoja
        hojaProc.Calculate
        DoEvents
        Application.Wait Now + TimeValue("0:00:01") ' Espera 1 segundo

        rutaPortafolios = hojaProc.Range("R_BORR").Value
        nombreArchivo = hojaProc.Range("N_BORR").Value & ".xlsb"

        If fso.FileExists(rutaPortafolios & "\" & nombreArchivo) Then
            On Error Resume Next
            fso.DeleteFile rutaPortafolios & "\" & nombreArchivo
            If Err.Number <> 0 Then
                hojaProc.Cells(filaRegistro, 3).Value = "'" & hojaProc.Range("N_BORR").Value
                hojaProc.Cells(filaRegistro, 4).Value = "Error al borrar"
                filaRegistro = filaRegistro + 1
                Err.Clear
            End If
            On Error GoTo 0
        Else
            hojaProc.Cells(filaRegistro, 3).Value = "'" & hojaProc.Range("N_BORR").Value
            hojaProc.Cells(filaRegistro, 4).Value = "No existe"
            filaRegistro = filaRegistro + 1
        End If
    Next i

End Sub

