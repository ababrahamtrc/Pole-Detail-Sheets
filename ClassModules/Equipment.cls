VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Equipment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public componentType As String
Public equipmentId As String
Public height As Integer 'inches
Public bottomHeight As Integer 'inches
Public owner As String
Public size As String
Public Bonded As String

Public Function getDisplayHeight() As String
    displayHeight = Utilities.inchesToFeetInches(Me.height)
    If Me.bottomHeight > 0 And Me.bottomHeight <> Me.height Then
        displayHeight = displayHeight & "(" & Utilities.inchesToFeetInches(Me.bottomHeight) & ")"
    End If
    
    getDisplayHeight = displayHeight
End Function
