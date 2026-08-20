VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NumericTextBoxCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents NumericTB As MSForms.TextBox
Attribute NumericTB.VB_VarHelpID = -1

Private Sub NumericTB_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 48 To 57
        Case 44
        Case Else
            KeyAscii = 0
    End Select
End Sub
