VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar_Form 
   Caption         =   "Import Data"
   ClientHeight    =   1290
   ClientLeft      =   195
   ClientTop       =   960
   ClientWidth     =   14040
   OleObjectBlob   =   "ProgressBar_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressBar_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Initialize()
  Me.StartUpPosition = 0
  Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
  Me.Top = Application.Top + (0.5 * Application.height) - (0.5 * Me.height)
End Sub
