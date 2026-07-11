Attribute VB_Name = "Módulo6"
Sub futuros()
celdas = Sheets("Procedimiento").Cells(36, 7)
rangobusca1 = Sheets("Procedimiento").Cells(36, 8)

Worksheets("Curvas_2").Activate
 With Range(celdas)                                    'set range to copy from / to.
            .Formula = rangobusca1 'refers to a workbook, sheet and first cell.
            '.Formula = "='M:\COB\PYG\2017\MARZO\28-MARZO\[Plantilla Performance 28-03-2017 Valores.xlsb]620 Contable'!A2"                                                    'It will put the relative references into the target sheet correctly.
            .Value = .Value                                     'changes formula to value.
        End With

End Sub

