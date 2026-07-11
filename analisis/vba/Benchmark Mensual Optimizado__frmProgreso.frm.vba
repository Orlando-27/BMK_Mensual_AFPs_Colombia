Attribute VB_Name = "frmProgreso"
Attribute VB_Base = "0{5F62F24A-08F9-4F74-9961-E4172D321AC4}{EACDC83E-D25B-497F-BD24-16877A60D8A6}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub lblProgreso_Click()

End Sub
Private Sub UserForm_Activate()
    Me.Repaint
End Sub

Public Sub ActualizarProgreso(paso As Integer, totalPasos As Integer, mensaje As String)
    Dim porcentaje As Double
    porcentaje = paso / totalPasos
    lblProgreso.Width = porcentaje * Me.Width
    Me.Caption = mensaje
    Me.Repaint
End Sub

