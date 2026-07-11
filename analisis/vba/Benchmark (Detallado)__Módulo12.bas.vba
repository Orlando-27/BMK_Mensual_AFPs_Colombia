Attribute VB_Name = "Módulo12"

Sub Copiarvalores()

    Dim valorOrigen As Variant
    
    ' Tomar el valor de la celda FU105 de la hoja activa
    valorOrigen = ActiveSheet.Range("PORTAFOLION").Value
    
    ' Pegar el valor en el rango llamado "PORTAFOLIO"
    Range("PORTAFOLIO").Value = valorOrigen

End Sub

