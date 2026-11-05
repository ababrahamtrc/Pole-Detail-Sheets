VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SIDesDoc_Form 
   Caption         =   "Job Type"
   ClientHeight    =   2295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3135
   OleObjectBlob   =   "SIDesDoc_FOrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SIDesDoc_FOrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public jobType As String

Private Sub CommandButton1_Click()
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "OptionButton" Then
            If ctrl.value = True Then
                jobType = ctrl.caption
                Exit For
            End If
        End If
    Next ctrl
    
    Me.Hide
End Sub
