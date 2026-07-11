Attribute VB_Name = "Module5"
Sub limpiar()
    Range("A2:V2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
End Sub
