VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GenerateSinglePDS_Form 
   Caption         =   "Generate Single PDS"
   ClientHeight    =   2385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3780
   OleObjectBlob   =   "GenerateSinglePDS_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GenerateSinglePDS_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Project As Project

Public Sub Initialize()
    On Error Resume Next
    
    ComboBox2.list = Array("4", "8", "12")
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.height) - (0.5 * Me.height)
    
    Set Project = New Project
    Call Project.extractImportDataFormat
    
    Dim poleNumbers() As Variant: ReDim poleNumbers(0 To Project.poles.count - 1)
    
    For i = 0 To Project.poles.count - 1
        poleNumbers(i) = Project.poles(i + 1).poleNumber
    Next i
    
    ComboBox1.list = poleNumbers
    ComboBox1.ListIndex = 0
    ComboBox2.ListIndex = 0
End Sub

Private Sub CommandButton1_Click()
    Dim pole As pole: Set pole = Project.findPole(GenerateSinglePDS_Form.ComboBox1.Value)
    
    If Not pole Is Nothing Then
        If Not Utilities.SheetExists(pole.poleNumber) Then
            Application.EnableEvents = False
            Application.ScreenUpdating = False
            Application.DisplayAlerts = False
        
            Project.extractImportDataFormat
            Call pole.createSheet(Project, GenerateSinglePDS_Form.ComboBox2.Value)
            
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            Application.DisplayAlerts = True
            Me.Hide
        Else
            MsgBox "Pole Already Exists."
        End If
    End If
End Sub
