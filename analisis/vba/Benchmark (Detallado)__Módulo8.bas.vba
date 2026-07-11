Attribute VB_Name = "Módulo8"
Sub ReemplazarPalabraEnHojasEspecificas()
    Dim ws As Worksheet
    Dim buscarTexto As String
    Dim reemplazarTexto As String
    
    ' Nombres de las hojas en las que deseas realizar el reemplazo
    Dim hojasObjetivo As Variant
    hojasObjetivo = Array("FutLoc Industria", "FutInt Industria", "Fwd Industria")
    
   
    
    For Each hojaNombre In hojasObjetivo
        ' Verificar si la hoja objetivo existe en el libro
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(hojaNombre)
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            ' Usar el método Replace para reemplazar el texto en la hoja objetivo
            ws.Cells.Replace What:="FONDO DE PENSIONES OBLIGATORIAS PORVENIR CONSERVADOR", Replacement:="PC", _
                LookAt:=xlPart, MatchCase:=False
                
            ' Usar el método Replace para reemplazar el texto en la hoja objetivo
            ws.Cells.Replace What:="FONDO DE PENSIONES OBLIGATORIAS PORVENIR MAYOR RIESGO", Replacement:="PM", _
                LookAt:=xlPart, MatchCase:=False

      ' Usar el método Replace para reemplazar el texto en la hoja objetivo
            ws.Cells.Replace What:="FONDO DE PENSIONES OBLIGATORIAS PORVENIR RETIRO PROGRAMADO", Replacement:="PR", _
                LookAt:=xlPart, MatchCase:=False
                
     ' Usar el método Replace para reemplazar el texto en la hoja objetivo
            ws.Cells.Replace What:="FONDO DE CESANTIAS SKANDIA", Replacement:="CS", _
                LookAt:=xlPart, MatchCase:=False

    ' Usar el método Replace para reemplazar el texto en la hoja objetivo
            ws.Cells.Replace What:="FONDO DE CESANTIAS", Replacement:="CS", _
                LookAt:=xlPart, MatchCase:=False

      ' Usar el método Replace para reemplazar el texto en la hoja objetivo
            ws.Cells.Replace What:="FONDO DE PENSIONES MODERADO", Replacement:="PO", _
                LookAt:=xlPart, MatchCase:=False
    
     ' Usar el método Replace para reemplazar el texto en la hoja objetivo
            ws.Cells.Replace What:="FONDO DE PENSIONES OBLIGATORIAS PROTECCIÓN CONSERVADOR", Replacement:="PC", _
                LookAt:=xlPart, MatchCase:=False
    
    ' Usar el método Replace para reemplazar el texto en la hoja objetivo
            ws.Cells.Replace What:="FONDO DE PENSIONES OBLIGATORIAS PROTECCION MAYOR RIESGO", Replacement:="PM", _
                LookAt:=xlPart, MatchCase:=False
    
    ' Usar el método Replace para reemplazar el texto en la hoja objetivo
            ws.Cells.Replace What:="FONDO DE PENSIONES OBLIGATORIAS PROTECCIÓN RETIRO PROGRAMADO", Replacement:="PR", _
                LookAt:=xlPart, MatchCase:=False

   

    ' Usar el método Replace para reemplazar el texto en la hoja objetivo
            ws.Cells.Replace What:="FONDO DE PENSIONES OBLIGATORIAS SKANDIA MODERADO", Replacement:="PO", _
                LookAt:=xlPart, MatchCase:=False

    ' Usar el método Replace para reemplazar el texto en la hoja objetivo
            ws.Cells.Replace What:="FONDO DE PENSIONES OBLIGATORIAS SKANDIA   CONSERVADOR", Replacement:="PC", _
                LookAt:=xlPart, MatchCase:=False
    
    
    ' Usar el método Replace para reemplazar el texto en la hoja objetivo
            ws.Cells.Replace What:="FONDO DE PENSIONES OBLIGATORIAS SKANDIA  MAYOR RIESGO", Replacement:="PM", _
                LookAt:=xlPart, MatchCase:=False

    ' Usar el método Replace para reemplazar el texto en la hoja objetivo
            ws.Cells.Replace What:="FONDO DE PENSIONES OBLIGATORIAS  SKANDIA RETIRO PROGRAMADO", Replacement:="PR", _
                LookAt:=xlPart, MatchCase:=False
                
    
    ' Usar el método Replace para reemplazar el texto en la hoja objetivo
            ws.Cells.Replace What:="FONDO DE PENSIONES OBLIGATORIO MODERADO", Replacement:="PO", _
                LookAt:=xlPart, MatchCase:=False
    
    
    
    ' Usar el método Replace para reemplazar el texto en la hoja objetivo
            ws.Cells.Replace What:="PROTECCION", Replacement:="PROTECCIÓN", _
                LookAt:=xlPart, MatchCase:=False
    
    ' Usar el método Replace para reemplazar el texto en la hoja objetivo
            ws.Cells.Replace What:="SKANDIA PENSIONES Y CESANTÍAS S.A.", Replacement:="OLD MUTUAL", _
                LookAt:=xlPart, MatchCase:=False
    
    
    
        
    


        Else
            MsgBox "La hoja '" & hojaNombre & "' no se encontró en el libro.", vbExclamation
        End If
    Next hojaNombre
    
End Sub



