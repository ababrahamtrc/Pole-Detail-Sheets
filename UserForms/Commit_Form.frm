VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Commit_Form 
   Caption         =   "Commit"
   ClientHeight    =   3510
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "Commit_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Commit_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public currentVersion As String
Public newVersion As String
Public newChanges As String

Public Sub Initialize(currentVersion_ As String)
    currentVersion = currentVersion_
    Me.version.text = currentVersion

    newVersion = ""
    changes = ""

    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.height) - (0.5 * Me.height)
End Sub

Private Sub commit_Click()
    If Me.version.text <> "" And Me.version.text <> currentVersion And Me.changes.text <> "" Then
        newVersion = Me.version.text
        newChanges = Me.changes.text
        Me.Hide
    Else
        MsgBox "Need a new version number and changes."
    End If
End Sub
