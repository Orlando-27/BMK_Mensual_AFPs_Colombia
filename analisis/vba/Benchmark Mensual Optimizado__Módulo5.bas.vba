Attribute VB_Name = "Módulo5"
Sub CrearCarpetaYCopiarArchivos()

    Dim rutaBase As String
    Dim rutaFinal As String
    Dim anio As String
    Dim mes As String
    Dim fecha As Date
    Dim archivos As Variant
    Dim i As Long
    Dim nombreArchivo As String
    Dim fso As Object
    Dim respuesta As VbMsgBoxResult
    
    ' Ruta base
    rutaBase = "M:\COB\Gerencia Estrategia\Analisis Cuantitativo\GERENCIA DE ANALISIS\ANALISIS FINANCIERO\ANALISIS CUANTITATIVO\Benchmark\Actualización Mensual\"
    
    ' Leer fecha
    fecha = Sheets("Parámetros").Range("C5").Value
    
    ' Año y mes
    anio = Year(fecha)
    mes = UCase(Format(fecha, "mmmm"))
    
    ' Ruta final
    rutaFinal = rutaBase & anio & "\" & mes
    
    ' Crear FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Crear carpeta año si no existe
    If Not fso.FolderExists(rutaBase & anio) Then
        fso.CreateFolder rutaBase & anio
    End If
    
    ' Crear carpeta mes si no existe
    If Not fso.FolderExists(rutaFinal) Then
        fso.CreateFolder rutaFinal
    End If
    
    ' Lista de archivos
    archivos = Array( _
    "M:\COB\Gerencia Estrategia\Analisis Cuantitativo\GERENCIA DE ANALISIS\ANALISIS FINANCIERO\ANALISIS CUANTITATIVO\Benchmark\Actualización Mensual\Benchmark Mensual Optimizado.xlsb", _
    "M:\COB\Gerencia Estrategia\Analisis Cuantitativo\GERENCIA DE ANALISIS\ANALISIS FINANCIERO\ANALISIS CUANTITATIVO\Benchmark\Calculadora Swap.xlsb", _
    "M:\COB\Gerencia Estrategia\Analisis Cuantitativo\GERENCIA DE ANALISIS\ANALISIS FINANCIERO\ANALISIS CUANTITATIVO\Benchmark\Sensibilidad_Fwds.xlsb", _
    "M:\COB\Gerencia Estrategia\Analisis Cuantitativo\GERENCIA DE ANALISIS\ANALISIS FINANCIERO\ANALISIS CUANTITATIVO\Benchmark\INICIO MACRO MULTIFONDOS.xlsm", _
    "M:\COB\Gerencia Estrategia\Analisis Cuantitativo\GERENCIA DE ANALISIS\ANALISIS FINANCIERO\ANALISIS CUANTITATIVO\Benchmark\Benchmark (Detallado).xlsb")
    
    ' Recorrer archivos
    For i = LBound(archivos) To UBound(archivos)
        
        If fso.FileExists(archivos(i)) Then
            
            nombreArchivo = fso.GetFileName(archivos(i))
            
            ' Si el archivo YA existe en destino
            If fso.FileExists(rutaFinal & "\" & nombreArchivo) Then
                
                respuesta = MsgBox( _
                    "El archivo ya existe:" & vbCrLf & nombreArchivo & vbCrLf & vbCrLf & _
                    "¿Desea reemplazarlo?", _
                    vbYesNo + vbQuestion, _
                    "Confirmar reemplazo")
                
                If respuesta = vbYes Then
                    fso.CopyFile archivos(i), rutaFinal & "\" & nombreArchivo, True
                Else
                    ' NO hace nada y sigue
                End If
                
            Else
                ' Si no existe ? copia directo
                fso.CopyFile archivos(i), rutaFinal & "\" & nombreArchivo
            End If
            
        Else
            MsgBox "No se encontró: " & archivos(i), vbExclamation
        End If
        
    Next i
    
    MsgBox "? Proceso completado correctamente", vbInformation

End Sub

