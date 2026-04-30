VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CU_Form 
   Caption         =   "Ambiguous CU"
   ClientHeight    =   2220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3435
   OleObjectBlob   =   "CU_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CU_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Public IsCancelled As Boolean

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        IsCancelled = True
        Me.Hide
        Cancel = True
    End If
End Sub

Public Sub Initialize(size As String)
    Me.Label1.caption = "Size " & size & " Primary"
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.height) - (0.5 * Me.height)
End Sub

Private Sub CommandButton1_Click()
    If Me.OptionButton1.visible And Not Me.OptionButton1.Value And Not Me.OptionButton2.Value Then
        MsgBox "Please select whether it's top or side ties."
    End If
    
    Me.Hide
End Sub
