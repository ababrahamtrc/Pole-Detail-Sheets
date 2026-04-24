VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Guy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public size As String
Public height As Integer
Public owner As String
Public id As String
Public componentType As String
Public proposedSize As String
Public proposedLead As String
Public proposedDirection As String

Public Function getDisplayHeight() As String
    Dim displayHeight As String
    displayHeight = Utilities.inchesToFeetInches(Me.height)
    
    getDisplayHeight = displayHeight
End Function
