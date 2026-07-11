Attribute VB_Name = "Módulo6"
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    ActiveSheet.Range("$A$3:$AD$101").AutoFilter Field:=8, Criteria1:="IBR"
    ActiveWindow.SmallScroll Down:=-3
End Sub
