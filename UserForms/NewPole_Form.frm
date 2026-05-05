VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewPole_Form 
   Caption         =   "New Pole"
   ClientHeight    =   2565
   ClientLeft      =   90
   ClientTop       =   390
   ClientWidth     =   3480
   OleObjectBlob   =   "NewPole_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewPole_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Initialize()
    ComboBox1.list = Array("35", "40", "45", "50", "55", "60", "65", "70")
    ComboBox2.list = Array("2", "3", "4")
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.height) - (0.5 * Me.height)
End Sub

Private Sub CommandButton1_Click()
    If ComboBox1.Value = "" Or ComboBox2.Value = "" Or TextBox1.Value = "" Or TextBox2.Value = "" Then
        MsgBox "Fill in all the required fields first."
        Exit Sub
    End If
    
    If SheetExists(TextBox1.Value) Then
        MsgBox "There's already a pole with that number."
        Exit Sub
    End If
    
    Dim Project As Project: Set Project = New Project
    Dim pole As pole: Set pole = New pole
    Project.extractFromSheets
    
    pole.poleNumber = TextBox1.Value
    pole.height = ComboBox1.Value
    pole.Class = ComboBox2.Value
    pole.species = TextBox2.Value
    
    Call pole.createSheet(Project)
    
    Me.Hide
End Sub
