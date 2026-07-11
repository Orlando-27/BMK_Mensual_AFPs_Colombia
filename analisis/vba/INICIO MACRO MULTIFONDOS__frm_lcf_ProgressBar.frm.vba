Attribute VB_Name = "frm_lcf_ProgressBar"
Attribute VB_Base = "0{9581044F-373C-4501-B564-BBED039952EE}{D2D0EA21-497C-4B3C-B44F-03D6DCEC188C}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

'>==============================================================================
'> /////////////////////////////////////////////////////////////////////////////
'>
'>  (c) Copyright 2007, Luis Carlos Flores López
'>  Under creative commons Attribution-NonCommercial-ShareAlike
'>  http://creativecommons.org/licenses/
'>
'>  File                : frm_lcf_ProgressBar.frm
'>  Author              : Luis Carlos Flores López
'>  Author web page     : www.Xperimentos.com
'>
'>  Form web page       : www.Xperimentos.com
'>
'>  Date                : 18/05/2001
'>  Last update date    : 10/06/2007
'>
'>  Language            : Visual Basic 6.0 for applications
'>  Made with           : Microsoft Excel
'>  Operating System    : Windows (Workstation & Server)
'>
'>  Description         : Progress Bar Form shows a windows form progress bar.
'>                        First of all you must initialize the progress bar,
'>                        use the "Initialize" method.
'>                        Shows the progress bar and use the increase method
'>                        to change the progress bar status.
'>                        Finally you should unload the windows form.
'>                        See below example for more details
'>
'>  Example:
'>
'      '>-----------------------------------------------------------------
'      Public Sub Example()
'          Dim oProgress As New frm_lcf_ProgressBar
'          Dim style As Integer
'          Dim windowCaption As String
'          Dim endRow As Long
'          Dim i As Long
'          style = 2                 ' Progress bar style (1 / 2).
'          windowCaption = "Example" ' Progress bar window caption.
'          endRow = 100000           ' Max value
'
'          ' Progress bar initialization
'          oProgress.Initialize endRow, style, windowCaption
'          oProgress.Show 0          ' Shows the progress bar window
'
'          For i = 0 To endRow - 4   ' Dummy loop for this example
'              '--
'              ' <<Do something, put here your code>>
'              '--
'              oProgress.Increase    ' Increases 1 unit the progress bar
'          Next
'          oProgress.Increase 4      ' Increases 4 units the progress bar
'          Unload oProgress          ' Unload progress bar window
'      End Sub
'      '>-----------------------------------------------------------------
'>
'>  History:
'>  18/05/2001        First version
'>  12/01/2004        New visual style
'>  22/05/2005        Some additions
'>  10/06/2007        Instructions was added
'>
'>  Version 0.5, 10/06/2007
'>
'> /////////////////////////////////////////////////////////////////////////////
'>==============================================================================
Option Explicit


'>==============================================================================
' Local Variables
Private mIntCurrent As Long
Private mIntMax As Long
Private Const c_MAX_LENG = 196
Private mIntType As Integer


'>==============================================================================
' Initialize the progress bar
' Params:
'       - vMax     - Max value = 100%. When the 100% of process has done
'       - vType    - visual style (1,2)
'       - vCaption - Window caption
Public Sub Initialize(vMax As Long, Optional vType As Integer = 1, Optional vCaption As String = "")
    mIntType = vType
    txtBar.Left = 6
    txtBar.Top = 6
    txtBar.Width = 198
    txtBar.Height = 26
    
    If vCaption <> "" Then
        Me.Caption = vCaption
    End If
    
    Select Case mIntType
        Case 1
            cmdProgressBar.Visible = True
            txt_ProgressBar.Visible = False
            txt_Blue.Visible = False
            txt_Grey.Visible = False
            
            cmdProgressBar.Left = 7
            cmdProgressBar.Top = 7
            cmdProgressBar.Height = 24
            cmdProgressBar.Width = 0
            cmdProgressBar.Caption = "0%"
        Case 2
            cmdProgressBar.Visible = False
            txt_ProgressBar.Visible = True
            txt_Blue.Visible = True
            txt_Grey.Visible = True
            
            txt_ProgressBar.Left = 7
            txt_ProgressBar.Top = 7.5
            txt_ProgressBar.Height = 23
            txt_ProgressBar.Width = 0

            txt_Blue.Left = 90
            txt_Blue.Top = 12
            txt_Blue.Height = 16
            txt_Blue.Width = 0
            txt_Blue.Caption = "0%"
            
            txt_Grey.Left = 90
            txt_Grey.Top = 12
            txt_Grey.Height = 16
            txt_Grey.Width = 36
            txt_Grey.Caption = "0%"
    End Select
    
    mIntCurrent = 0
    mIntMax = vMax
    
    DoEvents
End Sub
'>==============================================================================


'>==============================================================================
' Increase one unit the progress bar
Public Sub Increase(Optional vIncrease As Long = 1)
    Dim tmpBarValue As Integer
    mIntCurrent = mIntCurrent + vIncrease
    
    tmpBarValue = CInt(mIntCurrent * (c_MAX_LENG / mIntMax))
    Select Case mIntType
        Case 1
            cmdProgressBar.Width = tmpBarValue
            cmdProgressBar.Caption = Trim(CStr(CInt(tmpBarValue * (100 / c_MAX_LENG)))) & "%"
        Case 2
            txt_ProgressBar.Width = tmpBarValue
            If txt_ProgressBar.Width + txt_ProgressBar.Left >= txt_Blue.Left And txt_Blue.Width < txt_Grey.Width Then
                txt_Blue.Width = txt_ProgressBar.Width - 90 + 7
            End If
            txt_Blue.Caption = Trim(CStr(CInt(tmpBarValue * (100 / c_MAX_LENG)))) & "%"
            txt_Grey.Caption = Trim(CStr(CInt(tmpBarValue * (100 / c_MAX_LENG)))) & "%"
    End Select
    DoEvents
End Sub
'>==============================================================================



'>==============================================================================
' On from terminate (close or unload)
Private Sub UserForm_Terminate()
    'gBlnStopProcess = True
End Sub
'>==============================================================================
