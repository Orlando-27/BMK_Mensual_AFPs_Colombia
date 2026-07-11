Attribute VB_Name = "Módulo11"
Sub Enviar_Email_BMK_PO()

Sheets("Procedimiento").Calculate

    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String
    Dim FechaBMK As Date
    Dim fechaIND As Date
    Dim A_1, A_2, A_3, A_4, A_5 As Variant
    
    A_1 = Sheets("Parametros").Range("C84")
    A_2 = Sheets("Parametros").Range("C85")
    A_3 = Sheets("Parametros").Range("C86")
    A_4 = Sheets("Parametros").Range("C87")
    A_5 = Sheets("Parametros").Range("C88")
    FechaBMK = Sheets("Procedimiento").Range("D7").Value
    fechaIND = Range("fechaIND").Value
    
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    strbody = "<div style='font-family:Roboto, Arial; font-size:10pt; color:#0A2A66;'>" & _
              "Buenos días.<br><br>" & _
              "A continuación, se encuentra el benchmark valorado a la fecha " & Format(FechaBMK, "dd/mm/yyyy") & _
              ", con información de la industria al " & Format(fechaIND, "dd/mm/yyyy") & ".<br><br>" & _
              "En el siguiente hipervinculo encontrarán la herramienta:<br><br>" & _
              "<a href=""M:\COB\Gerencia Estrategia\Analisis Cuantitativo\GERENCIA DE ANALISIS\ANALISIS FINANCIERO\ANALISIS CUANTITATIVO\Benchmark\Publicados"" target=""_blank"">" & _
              "<img src=""M:\COB\Gerencia Estrategia\Analisis Cuantitativo\GERENCIA DE ANALISIS\ANALISIS FINANCIERO\ANALISIS CUANTITATIVO\Benchmark\Publicados\Imagencorreo.png"" style=""width:5px; height:auto; border:0;"">" & _
              "</a>" & _
              "</div>"
              
    On Error Resume Next
    
    With OutMail
        .Display
        .To = Sheets("Parametros").Range("C78").Value
        .CC = Sheets("Parametros").Range("C79").Value
        .BCC = Sheets("Parametros").Range("C80").Value
        .Subject = Sheets("Parametros").Range("C81").Value
        .HTMLBody = strbody & "<br>" & .HTMLBody
        .HTMLBody = .HTMLBody
        .Attachments.Add A_1
        .Attachments.Add A_2
        .Attachments.Add A_3
        .Attachments.Add A_4
        .Attachments.Add A_5
        .Send
        
    End With
    
    On Error GoTo 0
    Set OutMail = Nothing
    Set OutApp = Nothing
    Call Enviar_Email_BMK_PV
    Call Enviar_Email_Caldas_Manizales
      
End Sub


Sub Enviar_Email_BMK_PV()

    Sheets("Procedimiento").Calculate

    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String
    Dim DateBMK As Date
    DateBMK = Sheets("Procedimiento").Range("D7").Value

    ' Cargar archivos (excepto Caldas y Manizales)
    Dim A_12, A_13, A_14, A_15, A_16, A_17, A_18, A_19, A_20, A_21, A_22, A_23, A_24, A_26 As Variant
    Dim A_27, A_28, A_29, A_30, A_31, A_32, A_33, A_34, A_35, A_36, A_37, A_38, A_39, A_40 As Variant
    Dim A_41, A_42, A_43, A_44, A_45, A_46, A_47, A_48, A_49, A_50, A_51, A_52 As Variant

    A_12 = Sheets("Parametros").Range("C102")
    A_13 = Sheets("Parametros").Range("C103")
    A_14 = Sheets("Parametros").Range("C104")
    A_15 = Sheets("Parametros").Range("C105")
    A_16 = Sheets("Parametros").Range("C106")
    A_17 = Sheets("Parametros").Range("C107")
    A_18 = Sheets("Parametros").Range("C108")
    A_19 = Sheets("Parametros").Range("C109")
    A_20 = Sheets("Parametros").Range("C110")
    A_21 = Sheets("Parametros").Range("C111")
    A_22 = Sheets("Parametros").Range("C112")
    A_23 = Sheets("Parametros").Range("C113")
    A_24 = Sheets("Parametros").Range("C114")
    A_26 = Sheets("Parametros").Range("C116")

    A_27 = Sheets("Parametros").Range("E166")
    A_28 = Sheets("Parametros").Range("E167")
    A_29 = Sheets("Parametros").Range("E168")
    A_30 = Sheets("Parametros").Range("E169")
    A_31 = Sheets("Parametros").Range("E170")
    A_32 = Sheets("Parametros").Range("E171")
    A_33 = Sheets("Parametros").Range("E172")
    A_34 = Sheets("Parametros").Range("E173")
    A_35 = Sheets("Parametros").Range("E174")
    A_36 = Sheets("Parametros").Range("E175")
    A_37 = Sheets("Parametros").Range("E176")
    A_38 = Sheets("Parametros").Range("E177")
    A_39 = Sheets("Parametros").Range("E178")
    A_40 = Sheets("Parametros").Range("E179")
    A_41 = Sheets("Parametros").Range("E180")
    A_42 = Sheets("Parametros").Range("E181")
    A_43 = Sheets("Parametros").Range("E182")
    A_44 = Sheets("Parametros").Range("E183")
    A_45 = Sheets("Parametros").Range("E184")
    A_46 = Sheets("Parametros").Range("E185")
    A_47 = Sheets("Parametros").Range("E186")
    A_48 = Sheets("Parametros").Range("E187")
    A_49 = Sheets("Parametros").Range("E188")
    A_50 = Sheets("Parametros").Range("E189")
    A_51 = Sheets("Parametros").Range("E190")
    A_52 = Sheets("Parametros").Range("E191")

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    strbody = "<div style='font-family:Roboto, Arial; font-size:10pt; color:#0A2A66;'>" & _
              "Buenos días.<br><br>" & _
              "A continuación, se encuentra el benchmark valorado a la fecha " & Format(DateBMK, "dd/mm/yyyy.") & "<br><br>" & _
              "En el siguiente hipervinculo encontrarán la herramienta:<br><br>" & _
              "<a href=""M:\COB\Gerencia Estrategia\Analisis Cuantitativo\GERENCIA DE ANALISIS\ANALISIS FINANCIERO\ANALISIS CUANTITATIVO\Benchmark\Publicados"" target=""_blank"">" & _
              "<img src=""M:\COB\Gerencia Estrategia\Analisis Cuantitativo\GERENCIA DE ANALISIS\ANALISIS FINANCIERO\ANALISIS CUANTITATIVO\Benchmark\Publicados\Imagencorreo.png"" style=""width:5px; height:auto; border:0;"">" & _
              "</a>" & _
              "</div>"
              
    On Error Resume Next

    With OutMail
        .Display
        .To = Sheets("Parametros").Range("C96").Value
        .CC = Sheets("Parametros").Range("C97").Value
        .BCC = Sheets("Parametros").Range("C98").Value
        .Subject = Sheets("Parametros").Range("C99").Value
        .HTMLBody = strbody & "<br>" & .HTMLBody

        ' Adjuntar archivos (excepto Caldas y Manizales)
        .Attachments.Add A_12
        .Attachments.Add A_13
        .Attachments.Add A_14
        .Attachments.Add A_15
        .Attachments.Add A_16
        .Attachments.Add A_17
        .Attachments.Add A_18
        .Attachments.Add A_19
        .Attachments.Add A_20
        .Attachments.Add A_21
        .Attachments.Add A_22
        .Attachments.Add A_23
        .Attachments.Add A_24
        .Attachments.Add A_26
        .Attachments.Add A_27
        .Attachments.Add A_28
        .Attachments.Add A_29
        .Attachments.Add A_30
        .Attachments.Add A_31
        .Attachments.Add A_32
        .Attachments.Add A_33
        .Attachments.Add A_34
        .Attachments.Add A_35
        .Attachments.Add A_36
        .Attachments.Add A_37
        .Attachments.Add A_38
        .Attachments.Add A_39
        .Attachments.Add A_40
        .Attachments.Add A_41
        .Attachments.Add A_42
        .Attachments.Add A_43
        .Attachments.Add A_44
        .Attachments.Add A_45
        .Attachments.Add A_46
        .Attachments.Add A_47
        .Attachments.Add A_48
        .Attachments.Add A_49
        .Attachments.Add A_50
        .Attachments.Add A_51
        .Attachments.Add A_52
        .Send
    End With

    On Error GoTo 0
    Set OutMail = Nothing
    Set OutApp = Nothing

End Sub

Sub Enviar_Email_Caldas_Manizales()

    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String
    Dim DateBMK As Date
    DateBMK = Sheets("Procedimiento").Range("D7").Value

    Dim A_25, A_53 As Variant
    A_25 = Sheets("Parametros").Range("C115")
    A_53 = Sheets("Parametros").Range("E192")

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    strbody = "<div style='font-family:Roboto, Arial; font-size:10pt; color:#0A2A66;'>" & _
              "Buenos días.<br><br>" & _
              "A continuación, se encuentra el benchmark valorado a la fecha " & Format(DateBMK, "dd/mm/yyyy.") & "<br><br>" & _
              "En el siguiente hipervinculo encontrarán la herramienta:<br><br>" & _
              "<a href=""M:\COB\Gerencia Estrategia\Analisis Cuantitativo\GERENCIA DE ANALISIS\ANALISIS FINANCIERO\ANALISIS CUANTITATIVO\Benchmark\Publicados"" target=""_blank"">" & _
              "<img src=""M:\COB\Gerencia Estrategia\Analisis Cuantitativo\GERENCIA DE ANALISIS\ANALISIS FINANCIERO\ANALISIS CUANTITATIVO\Benchmark\Publicados\Imagencorreo.png"" style=""width:5px; height:auto; border:0;"">" & _
              "</a>" & _
              "</div>"

    On Error Resume Next
    With OutMail
        .Display
        .To = Sheets("Parametros").Range("C96").Value
        .CC = Sheets("Parametros").Range("C97").Value
        .BCC = Sheets("Parametros").Range("C98").Value
        .Subject = Sheets("Parametros").Range("D99").Value
        .HTMLBody = strbody & "<br>" & .HTMLBody
        .Attachments.Add A_25
        .Attachments.Add A_53
        .Send
    End With

    On Error GoTo 0
    Set OutMail = Nothing
    Set OutApp = Nothing

End Sub


Sub Enviar_Email_BMK_PVant()

Sheets("Procedimiento").Calculate

    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String
    Dim FechaBMK As Date
    Dim fechaIND As Date
    Dim A_12, A_13, A_14, A_15, A_16, A_17, A_18, A_19, A_20, A_21, A_22, A_23, A_24, A_25_A_26, A_27, A_28, A_29, A_30, A_31, A_32, A_33, A_34, A_35, A_36, A_37, A_38, A_39, A_40, A_41, A_42, A_43, A_44, A_45, A_46, A_47, A_48, A_49, A_50, A_51, A_52, A_53 As Variant
    DateBMK = Sheets("Procedimiento").Range("D7").Value
    
    A_12 = Sheets("Parametros").Range("C102")
    A_13 = Sheets("Parametros").Range("C103")
    A_14 = Sheets("Parametros").Range("C104")
    A_15 = Sheets("Parametros").Range("C105")
    A_16 = Sheets("Parametros").Range("C106")
    A_17 = Sheets("Parametros").Range("C107")
    A_18 = Sheets("Parametros").Range("C108")
    A_19 = Sheets("Parametros").Range("C109")
    A_20 = Sheets("Parametros").Range("C110")
    A_21 = Sheets("Parametros").Range("C111")
    A_22 = Sheets("Parametros").Range("C112")
    A_23 = Sheets("Parametros").Range("C113")
    A_24 = Sheets("Parametros").Range("C114")
    A_25 = Sheets("Parametros").Range("C115")
    A_26 = Sheets("Parametros").Range("C116")
    A_27 = Sheets("Parametros").Range("E166")
    A_28 = Sheets("Parametros").Range("E167")
    A_29 = Sheets("Parametros").Range("E168")
    A_30 = Sheets("Parametros").Range("E169")
    A_31 = Sheets("Parametros").Range("E170")
    A_32 = Sheets("Parametros").Range("E171")
    A_33 = Sheets("Parametros").Range("E172")
    A_34 = Sheets("Parametros").Range("E173")
    A_35 = Sheets("Parametros").Range("E174")
    A_36 = Sheets("Parametros").Range("E175")
    A_37 = Sheets("Parametros").Range("E176")
    A_38 = Sheets("Parametros").Range("E177")
    A_39 = Sheets("Parametros").Range("E178")
    A_40 = Sheets("Parametros").Range("E179")
    A_41 = Sheets("Parametros").Range("E180")
    A_42 = Sheets("Parametros").Range("E181")
    A_43 = Sheets("Parametros").Range("E182")
    A_44 = Sheets("Parametros").Range("E183")
    A_45 = Sheets("Parametros").Range("E184")
    A_46 = Sheets("Parametros").Range("E185")
    A_47 = Sheets("Parametros").Range("E186")
    A_48 = Sheets("Parametros").Range("E187")
    A_49 = Sheets("Parametros").Range("E188")
    A_50 = Sheets("Parametros").Range("E189")
    A_51 = Sheets("Parametros").Range("E190")
    A_52 = Sheets("Parametros").Range("E191")
    A_53 = Sheets("Parametros").Range("E192")

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
     strbody = "Buenos días.<br><br>" & _
              "A continuación, se encuentra el benchmark valorado a la fecha " & Format(DateBMK, "dd/mm/yyyy.") & "<br><br>" & _
              "En el siguiente hipervínculo encontrarán la herramienta:<br><br>" & _
              "<a href='M:\COB\Gerencia Estrategia\Analisis Cuantitativo\GERENCIA DE ANALISIS\ANALISIS FINANCIERO\ANALISIS CUANTITATIVO\Benchmark\Publicados'target'_BLANK>" & _
              "</FONT><img src='M:\COB\Gerencia Estrategia\Analisis Cuantitativo\GERENCIA DE ANALISIS\ANALISIS FINANCIERO\ANALISIS CUANTITATIVO\Benchmark\Imagen publicados\Publicados.png'><a/><br/>" & _
              "<br><br>" & _
              "Cordialmente, <br>"
            
    On Error Resume Next
    
    With OutMail
        .Display
        .To = Sheets("Parametros").Range("C96").Value
        .CC = Sheets("Parametros").Range("C97").Value
        .BCC = Sheets("Parametros").Range("C98").Value
        .Subject = Sheets("Parametros").Range("C99").Value
        .HTMLBody = strbody & "<br>" & .HTMLBody
        .HTMLBody = .HTMLBody
        .Attachments.Add A_12
        .Attachments.Add A_13
        .Attachments.Add A_14
        .Attachments.Add A_15
        .Attachments.Add A_16
        .Attachments.Add A_17
        .Attachments.Add A_18
        .Attachments.Add A_19
        .Attachments.Add A_20
        .Attachments.Add A_21
        .Attachments.Add A_22
        .Attachments.Add A_23
        .Attachments.Add A_24
        .Attachments.Add A_26
        .Attachments.Add A_27
        .Attachments.Add A_28
        .Attachments.Add A_29
        .Attachments.Add A_30
        .Attachments.Add A_31
        .Attachments.Add A_32
        .Attachments.Add A_33
        .Attachments.Add A_34
        .Attachments.Add A_35
        .Attachments.Add A_36
        .Attachments.Add A_37
        .Attachments.Add A_38
        .Attachments.Add A_39
        .Attachments.Add A_40
        .Attachments.Add A_41
        .Attachments.Add A_42
        .Attachments.Add A_43
        .Attachments.Add A_44
        .Attachments.Add A_45
        .Attachments.Add A_46
        .Attachments.Add A_47
        .Attachments.Add A_48
        .Attachments.Add A_49
        .Attachments.Add A_50
        .Attachments.Add A_51
        .Attachments.Add A_52
        .Send
        
    End With
    
    On Error GoTo 0
    Set OutMail = Nothing
    Set OutApp = Nothing
    'Manizales y Caldas
       Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
     strbody = "Buenos días.<br><br>" & _
              "A continuación, se encuentra el benchmark valorado a la fecha " & Format(DateBMK, "dd/mm/yyyy.") & "<br><br>" & _
              "En el siguiente hipervínculo encontrarán la herramienta:<br><br>" & _
              "<a href='M:\COB\Gerencia Estrategia\Analisis Cuantitativo\GERENCIA DE ANALISIS\ANALISIS FINANCIERO\ANALISIS CUANTITATIVO\Benchmark\Publicados'target'_BLANK>" & _
              "</FONT><img src='M:\COB\Gerencia Estrategia\Analisis Cuantitativo\GERENCIA DE ANALISIS\ANALISIS FINANCIERO\ANALISIS CUANTITATIVO\Benchmark\Imagen publicados\Publicados.png'><a/><br/>" & _
              "<br><br>" & _
              "Cordialmente, <br>"
            
    On Error Resume Next
    
    With OutMail
        .Display
        .To = Sheets("Parametros").Range("C96").Value
        .CC = Sheets("Parametros").Range("C97").Value
        .BCC = Sheets("Parametros").Range("C98").Value
        .Subject = Sheets("Parametros").Range("D99").Value
        .HTMLBody = strbody & "<br>" & .HTMLBody
        .HTMLBody = .HTMLBody
        .Attachments.Add A_25
        .Attachments.Add A_53
        .Send
        
    End With
    
    On Error GoTo 0
    Set OutMail = Nothing
    Set OutApp = Nothing
    
End Sub

