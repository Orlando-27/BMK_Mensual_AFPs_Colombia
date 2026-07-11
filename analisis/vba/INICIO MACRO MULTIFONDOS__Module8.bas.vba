Attribute VB_Name = "Module8"
Sub borrar()
    Sheets("PORTAFOLIOS").Select
    
    Range("A2:AZ1048576").ClearContents
    
    Sheets("PROCEDIMIENTO").Select
End Sub
