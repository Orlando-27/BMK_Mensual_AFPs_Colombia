Attribute VB_Name = "Módulo2"
'JOSE MANUEL IBAÑEZ DELGADO
'INGENIERIA FINANCIERA
'CALCULADORA SWAP - 2016 - 1


Sub vpn_1()

Dim StartTime As Double
Dim SecondsElapsed As Double
 
StartTime2 = Timer


Range("b3").Select
'ActiveCell.FormulaR1C1 = "=+TODAY()-85"
Application.ScreenUpdating = False
Call EXPORTARDATOS
Calculate
Call CALCULARVPN
Call PEGADO
Calculate
Call Limpiar_Datos_externos

SecondsElapsed = Round(Timer - StartTime2, 2)
Sheets("RUTAS").Activate
Sheets("RUTAS").Range("i25") = SecondsElapsed


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
