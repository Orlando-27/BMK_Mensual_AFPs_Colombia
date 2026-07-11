Attribute VB_Name = "Módulo2"
'Formulas para el calculo de la duración de los SWAPS

Function Interp(ByVal plazos As Excel.Range, ByVal curva As Excel.Range, ByVal T As Currency)
    Dim i(1) As Integer
    Dim inf_x, sup_x, inf_y, sup_y As Double
    If T < plazos.Rows(1) Then
        Interp = curva.Rows(1)
    ElseIf T > plazos.Rows(plazos.Rows.Count) Then
        Interp = curva.Rows(plazos.Rows.Count)
    Else
        For i(1) = 2 To plazos.Rows.Count
            If plazos.Rows(i(1) - 1) <= T And plazos.Rows(i(1)) >= T Then
                inf_x = plazos.Rows(i(1) - 1)
                inf_y = curva.Rows(i(1) - 1)
                sup_x = plazos.Rows(i(1))
                sup_y = curva.Rows(i(1))
            End If
        Next
        Interp = (T - inf_x) * (sup_y - inf_y) / (sup_x - inf_x) + inf_y
    End If
End Function

Function Tiempo(modalidad)

Dim letra As Variant

    letra = Mid(modalidad, 1, 1)
    
    If letra = "A" Then
        Tiempo = 12
    ElseIf letra = "S" Then
        Tiempo = 6
    ElseIf letra = "T" Then
        Tiempo = 3
    ElseIf letra = "M" Then
        Tiempo = 1
    ElseIf letra = "P" Then
        Tiempo = 0
    End If
    
End Function

Function SWAPTF(vencimiento, CUPON, modalidad, datum As Date, MONEDA As String, NOMINAL As Variant, VPN0DUR1 As Integer)
 
 Dim T As Variant
 Dim i As Variant
 Dim n As Variant
 Dim par As Variant
 Dim yld As Variant
 Dim yld1 As Variant
 Dim yld2 As Variant
 Dim fechas() As Variant
 Dim valpneto As Variant
 Dim valpneto1 As Variant
 Dim valpneto2 As Variant
        
        
    If vencimiento < datum Then
        SWAPTF = "Vencido"
        Exit Function
    End If
    
    Dim origen As Date
      
    T = Tiempo(modalidad)
    origen = EmisionInicial(vencimiento, modalidad, datum)
    n = Numcupones(origen, vencimiento, modalidad)

    par = NOMINAL
    If MONEDA = "USD" Then
        Set TASA1 = Range("Dias_Libor")
        Set TASA2 = Range("Tasa_Libor")
    Else
        Set TASA1 = Range("Dias_IBR") 'vector de dias
        Set TASA2 = Range("Tasa_IBR")   'vectos de tasas
    End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If T = 0 Then
        
        ReDim fechas(1, 9)
        
        fechas(1, 0) = datum  'fecha valoracion
        fechas(1, 1) = (vencimiento - fechas(1, 0)) / 365 'factor de conversiòn dias.
        
        yld = Interp(TASA1, TASA2, fechas(1, 1) * 365)  'tasa de descuento flujo
        yld1 = yld + (0.01 / 100) 'tasa de descuento + 1pbs
        yld2 = yld - (0.01 / 100)   'tasa de descuento - 1pbs
        fechas(1, 2) = par * (1 + CUPON)    'flujo
        fechas(1, 3) = 1 / ((1 + yld) ^ fechas(1, 1))   'factor de valor presente tasa de descuento
        fechas(1, 4) = 1 / ((1 + yld1) ^ fechas(1, 1))  'factor de valor presente tasa de descuento + 1pbs
        fechas(1, 5) = 1 / ((1 + yld2) ^ fechas(1, 1))  'factor de valor presente tasa de descuento - 1pbs
        fechas(1, 6) = yld 'tasa de descuento
        
        fechas(1, 7) = fechas(1, 2) * fechas(1, 3)  'flujo por factor de valor presente tasa de descuento
        'VPN 1PB+
        fechas(1, 8) = fechas(1, 2) * fechas(1, 4)   'flujo por factor de valor presente tasa de descuento +1pbs
        'VPN 1PB-
        fechas(1, 9) = fechas(1, 2) * fechas(1, 5)  'flujo por factor de valor presente tasa de descuento -1pbs

        valpneto = fechas(1, 7)
        valpneto1 = fechas(1, 8)
        valpneto2 = fechas(1, 9)

            If VPN0DUR1 = 1 Then
                SWAPTF = (valpneto2 - valpneto1) / (2 * valpneto * (0.01 / 100))
            Else
                SWAPTF = valpneto
            End If
    Else
        ReDim fechas(n, 10)

        For i = 0 To n
            fechas(i, 0) = DateAdd("m", T * i, origen)  'fecha primer flujo
            fechas(i, 1) = (fechas(i, 0) - datum) / 365 'factor de conversiòn perdiodicidad al año.
            fechas(i, 2) = 12 / T   'num flujos en el año

            fechas(i, 3) = par * CUPON / fechas(i, 2)   'monto del flujo

            yld = Interp(TASA1, TASA2, fechas(i, 1) * 365)  'tasa de descuento flujo
            yld1 = yld + (0.01 / 100) 'tasa de descuento + 1pbs
            yld2 = yld - (0.01 / 100)   'tasa de descuento - 1pbs
            fechas(i, 4) = 1 / ((1 + yld) ^ fechas(i, 1))   'factor de valor presente tasa de descuento
            fechas(i, 6) = 1 / ((1 + yld1) ^ fechas(i, 1))  'factor de valor presente tasa de descuento + 1pbs
            fechas(i, 7) = 1 / ((1 + yld2) ^ fechas(i, 1))  'factor de valor presente tasa de descuento - 1pbs
            fechas(i, 10) = yld 'tasa de descuento
        Next i
    
        fechas(n, 3) = par + fechas(n, 3)   'nominal + ultimo flujo
    
        For i = 0 To n
            'VPN
            fechas(i, 5) = fechas(i, 3) * fechas(i, 4)  'flujo por factor de valor presente tasa de descuento
            'VPN 1PB+
            fechas(i, 8) = fechas(i, 3) * fechas(i, 6)   'flujo por factor de valor presente tasa de descuento +1pbs
            'VPN 1PB-
            fechas(i, 9) = fechas(i, 3) * fechas(i, 7)  'flujo por factor de valor presente tasa de descuento -1pbs
    
            valpneto = valpneto + fechas(i, 5)
            valpneto1 = valpneto1 + fechas(i, 8)
            valpneto2 = valpneto2 + fechas(i, 9)
        Next i
           If VPN0DUR1 = 1 Then
            SWAPTF = (valpneto2 - valpneto1) / (2 * valpneto * (0.01 / 100))
           Else
            SWAPTF = valpneto
    
           End If

    End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
'    ReDim fechas(n, 10)
'
'    For i = 0 To n
'        fechas(i, 0) = DateAdd("m", T * i, origen)  'fecha primer flujo
'        fechas(i, 1) = (fechas(i, 0) - datum) / 365 'factor de conversiòn perdiodicidad al año.
'        fechas(i, 2) = 12 / T   'num flujos en el año
'
'        fechas(i, 3) = par * CUPON / fechas(i, 2)   'monto del flujo
'
'        yld = Interp(TASA1, TASA2, fechas(i, 1) * 365)  'tasa de descuento flujo
'        yld1 = yld + (0.01 / 100) 'tasa de descuento + 1pbs
'        yld2 = yld - (0.01 / 100)   'tasa de descuento - 1pbs
'        fechas(i, 4) = 1 / ((1 + yld) ^ fechas(i, 1))   'factor de valor presente tasa de descuento
'        fechas(i, 6) = 1 / ((1 + yld1) ^ fechas(i, 1))  'factor de valor presente tasa de descuento + 1pbs
'        fechas(i, 7) = 1 / ((1 + yld2) ^ fechas(i, 1))  'factor de valor presente tasa de descuento - 1pbs
'        fechas(i, 10) = yld 'tasa de descuento
'    Next i
'
'    fechas(n, 3) = par + fechas(n, 3)   'nominal + ultimo flujo
'
'    For i = 0 To n
'        'VPN
'        fechas(i, 5) = fechas(i, 3) * fechas(i, 4)  'flujo por factor de valor presente tasa de descuento
'        'VPN 1PB+
'        fechas(i, 8) = fechas(i, 3) * fechas(i, 6)   'flujo por factor de valor presente tasa de descuento +1pbs
'        'VPN 1PB-
'        fechas(i, 9) = fechas(i, 3) * fechas(i, 7)  'flujo por factor de valor presente tasa de descuento -1pbs
'
'        valpneto = valpneto + fechas(i, 5)
'        valpneto1 = valpneto1 + fechas(i, 8)
'        valpneto2 = valpneto2 + fechas(i, 9)
'    Next i
'       If VPN0DUR1 = 1 Then
'        SWAPTF = (valpneto2 - valpneto1) / (2 * valpneto * (0.01 / 100))
'       Else
'        SWAPTF = valpneto
'
'       End If
End Function


Function EmisionInicial(vencimiento, modalidad, datum As Date) As Date

    T = Tiempo(modalidad)
    
    If T = 0 Then
        
        EmisionInicial = datum
        
    Else
    
        Fecha = vencimiento
        
        Do Until Fecha <= datum
            Fecha = DateAdd("m", -T * (i + 1), vencimiento)
            i = i + 1
        Loop
            
        EmisionInicial = DateAdd("m", T, Fecha)
        
    End If
       
End Function
Function Numcupones(origen, vencimiento, modalidad)

    T = Tiempo(modalidad)
    
    If T = 0 Then
        Exit Function
    End If
    
    Fecha = origen
    Do Until Fecha >= vencimiento
        Fecha = DateAdd("m", T * (n + 1), origen)
        n = n + 1
    Loop
    
    Numcupones = n
    
End Function




Function SWAPLIBOR()

SWAPLIBOR = 1 / 365

End Function
Function SWAPIPC(vencimiento, CUPON, modalidad, datum As Date, MONEDA As String, NOMINAL As Variant, VPN0DUR1 As Integer)
 Dim T As Variant
 
         Dim i As Variant
         Dim n As Variant
        Dim par As Variant
        Dim yld As Variant
        Dim yld1 As Variant
        Dim yld2 As Variant
        Dim fechas() As Variant
        Dim valpneto As Variant
        Dim valpneto1 As Variant
        Dim valpneto2 As Variant
        
        
        
        
        
    If vencimiento < datum Then
        SWAPIPC = "Vencido"
        Exit Function
    End If
    
    Dim origen As Date
      
    T = Tiempo(modalidad)
    origen = EmisionInicial(vencimiento, modalidad, datum)
    n = Numcupones(origen, vencimiento, modalidad)

    par = NOMINAL
    cuponipc = (1 + Range("IPC")) * (1 + CUPON) - 1
  
    ReDim fechas(n, 10)
    
    For i = 0 To n
        fechas(i, 0) = DateAdd("m", T * i, origen)
        fechas(i, 1) = (fechas(i, 0) - datum) / 365
        fechas(i, 2) = 12 / T
        
        fechas(i, 3) = par * ((1 + cuponipc) ^ ((fechas(i, 0) - DateAdd("m", -T, fechas(i, 0))) / 365) - 1)
        
        yld = Range("IPC")
        yld1 = yld + (0.01 / 100)
        yld2 = yld - (0.01 / 100)
        fechas(i, 4) = 1 / ((1 + yld) ^ fechas(i, 1))
        fechas(i, 6) = 1 / ((1 + yld1) ^ fechas(i, 1))
        fechas(i, 7) = 1 / ((1 + yld2) ^ fechas(i, 1))
    
    Next i
    
    fechas(n, 3) = par + fechas(n, 3)
    
    For i = 0 To n
        'VPN
        fechas(i, 5) = fechas(i, 3) * fechas(i, 4)
       'VPN 1PB+
       fechas(i, 8) = fechas(i, 3) * fechas(i, 6)
        'VPN 1PB-
        fechas(i, 9) = fechas(i, 3) * fechas(i, 7)
        
        valpneto = valpneto + fechas(i, 5)
        valpneto1 = valpneto1 + fechas(i, 8)
        valpneto2 = valpneto2 + fechas(i, 9)
    Next i
       If VPN0DUR1 = 1 Then
        SWAPIPC = (((valpneto2 - valpneto1) / 2) / valpneto) * 10000
       Else
        SWAPIPC = valpneto
       End If
End Function

Function SWAPDTF(vencimiento, CUPON As Variant, modalidad, datum As Date, MONEDA As String, NOMINAL As Variant, base As Integer, VPN0DUR1 As Integer)
 Dim T As Variant
         Dim i As Variant
         Dim n As Variant
        Dim par As Variant
        Dim yld As Variant
        Dim yld1 As Variant
        Dim yld2 As Variant
        Dim fechas() As Variant
        Dim valpneto As Variant
        Dim valpneto1 As Variant
        Dim valpneto2 As Variant
        
       If base = 0 Then
        
        base1 = 365
       Else
        base1 = 360
       End If
    If vencimiento < datum Then
        SWAPDTF = "Vencido"
        Exit Function
    End If
    
    Dim origen As Date
      
    T = Tiempo(modalidad)
    origen = EmisionInicial(vencimiento, modalidad, datum)
    n = Numcupones(origen, vencimiento, modalidad)

    par = NOMINAL
    
    cuponDTF = ((1 + Range("DTF")) * (1 + CUPON)) - 1
    
    ReDim fechas(n, 10)
    
    For i = 0 To n
        fechas(i, 0) = DateAdd("m", T * i, origen)
        fechas(i, 1) = (fechas(i, 0) - datum) / 365
        fechas(i, 2) = 12 / T
        'PRIMER CUPON
       If i = 0 Then
        
       
            CUPONini = ((1 + Interp(Range("fecha"), Range("DTFSERIE"), DateAdd("m", -T, fechas(i, 0)))) * (1 + CUPON)) - 1
          If base = 0 Then
             fechas(i, 3) = par * ((1 + CUPONini) ^ ((fechas(i, 0) - DateAdd("m", -T, fechas(i, 0))) / 365) - 1)
          Else
            fechas(i, 3) = par * ((1 + CUPONini) ^ (1 / 4) - 1)
          End If
       Else
         If base = 0 Then
            fechas(i, 3) = par * ((1 + cuponDTF) ^ ((fechas(i, 0) - DateAdd("m", -T, fechas(i, 0))) / 365) - 1)
         Else
            fechas(i, 3) = par * ((1 + cuponDTF) ^ (1 / 4) - 1)
         End If
       End If
        yld = Range("DTF")
        yld1 = yld + (0.01 / 100)
        yld2 = yld - (0.01 / 100)
        fechas(i, 4) = 1 / ((1 + yld) ^ fechas(i, 1))
        fechas(i, 6) = 1 / ((1 + yld1) ^ fechas(i, 1))
        fechas(i, 7) = 1 / ((1 + yld2) ^ fechas(i, 1))
    
    Next i
    
    fechas(n, 3) = par + fechas(n, 3)
    
    For i = 0 To n
        'VPN
        fechas(i, 5) = fechas(i, 3) * fechas(i, 4)
       'VPN 1PB+
       fechas(i, 8) = fechas(i, 3) * fechas(i, 6)
        'VPN 1PB-
        fechas(i, 9) = fechas(i, 3) * fechas(i, 7)
        
        valpneto = valpneto + fechas(i, 5)
        valpneto1 = valpneto1 + fechas(i, 8)
        valpneto2 = valpneto2 + fechas(i, 9)
    Next i
       If VPN0DUR1 = 1 Then
        SWAPDTF = (((valpneto2 - valpneto1) / 2) / valpneto) * 10000
       Else
        SWAPDTF = valpneto
    
       End If
End Function

Function SWAPUVR(vencimiento, CUPON, modalidad, datum As Date, MONEDA As String, NOMINAL As Variant, VPN0DUR1 As Integer)
 Dim T As Variant
         Dim i As Variant
         Dim n As Variant
        Dim par As Variant
        Dim yld As Variant
        Dim yld1 As Variant
        Dim yld2 As Variant
        Dim fechas() As Variant
        Dim valpneto As Variant
        Dim valpneto1 As Variant
        Dim valpneto2 As Variant
        
        
        
        
        
    If vencimiento < datum Then
        SWAPUVR = "Vencido"
        Exit Function
    End If
    
    Dim origen As Date
      
    T = Tiempo(modalidad)
    origen = EmisionInicial(vencimiento, modalidad, datum)
    n = Numcupones(origen, vencimiento, modalidad)

    par = NOMINAL
    cuponUVR = CUPON
  
    ReDim fechas(n, 10)
    
    For i = 0 To n
        fechas(i, 0) = DateAdd("m", T * i, origen)
        fechas(i, 1) = (fechas(i, 0) - datum) / 365
        fechas(i, 2) = 12 / T
        
        fechas(i, 3) = par * ((1 + cuponUVR) ^ ((fechas(i, 0) - DateAdd("m", -T, fechas(i, 0))) / 365) - 1)
        
        yld = Interp(Range("Dias_UVR"), Range("Tasa_UVR"), fechas(i, 1) * 365)
        yld1 = yld + (0.01 / 100)
        yld2 = yld - (0.01 / 100)
        fechas(i, 4) = 1 / ((1 + yld) ^ fechas(i, 1))
        fechas(i, 6) = 1 / ((1 + yld1) ^ fechas(i, 1))
        fechas(i, 7) = 1 / ((1 + yld2) ^ fechas(i, 1))
    
    Next i
    
    fechas(n, 3) = par + fechas(n, 3)
    
    For i = 0 To n
        'VPN
        fechas(i, 5) = fechas(i, 3) * fechas(i, 4)
       'VPN 1PB+
       fechas(i, 8) = fechas(i, 3) * fechas(i, 6)
        'VPN 1PB-
        fechas(i, 9) = fechas(i, 3) * fechas(i, 7)
        
        valpneto = valpneto + fechas(i, 5)
        valpneto1 = valpneto1 + fechas(i, 8)
        valpneto2 = valpneto2 + fechas(i, 9)
    Next i
       If VPN0DUR1 = 1 Then
        SWAPUVR = (((valpneto2 - valpneto1) / 2) / valpneto) * 10000
       Else
        SWAPUVR = valpneto
       End If
End Function

Function SWAPIBRPV()

SWAPIBRPV = 1 / 365

End Function
Function SWAPIBRPF(vencimiento, CUPON, modalidad, datum As Date, MONEDA As String, NOMINAL As Variant, VPN0DUR1 As Integer)
 Dim T As Variant
         Dim i As Variant
         Dim n As Variant
        Dim par As Variant
        Dim yld As Variant
        Dim yld1 As Variant
        Dim yld2 As Variant
        Dim fechas() As Variant
        Dim valpneto As Variant
        Dim valpneto1 As Variant
        Dim valpneto2 As Variant
        
        
        
        
        
    If vencimiento < datum Then
        SWAPIBRPF = "Vencido"
        Exit Function
    End If
    
    Dim origen As Date
      
    T = Tiempo(modalidad)
    origen = EmisionInicial(vencimiento, modalidad, datum)
    n = Numcupones(origen, vencimiento, modalidad)

    par = NOMINAL
    
         Set TASA1 = Range("IBRF1")
        Set TASA2 = Range("IBRF2")
    
  
    ReDim fechas(n, 10)
    
    For i = 0 To n
        fechas(i, 0) = DateAdd("m", T * i, origen)
        fechas(i, 1) = (fechas(i, 0) - datum) / 365
        fechas(i, 2) = 12 / T
        
        fechas(i, 3) = par * CUPON / fechas(i, 2)
        
        yld = Interp(TASA1, TASA2, fechas(i, 1) * 365)
        yld1 = yld + (0.01 / 100)
        yld2 = yld - (0.01 / 100)
        fechas(i, 4) = 1 / ((1 + yld) ^ fechas(i, 1))
        fechas(i, 6) = 1 / ((1 + yld1) ^ fechas(i, 1))
        fechas(i, 7) = 1 / ((1 + yld2) ^ fechas(i, 1))
        fechas(i, 10) = yld
    Next i
    
    fechas(n, 3) = par + fechas(n, 3)
    
    For i = 0 To n
        'VPN
        fechas(i, 5) = fechas(i, 3) * fechas(i, 4)
       'VPN 1PB+
       fechas(i, 8) = fechas(i, 3) * fechas(i, 6)
        'VPN 1PB-
        fechas(i, 9) = fechas(i, 3) * fechas(i, 7)
        
        valpneto = valpneto + fechas(i, 5)
        valpneto1 = valpneto1 + fechas(i, 8)
        valpneto2 = valpneto2 + fechas(i, 9)
    Next i
       If VPN0DUR1 = 1 Then
        SWAPIBRPF = (valpneto2 - valpneto1) / (2 * valpneto * (0.01 / 100))
       Else
        SWAPIBRPF = valpneto
        
       End If
End Function







