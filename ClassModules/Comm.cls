VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Comm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public owner As String
Public height As String
Public modification As String
Public midspans As String
Public index As Integer
Public Boxed As Boolean
Public drops As Boolean
Public UpgradeGuy As Boolean
Public MoveAnchor As Boolean
Public PreviousAnchor As Integer
Public NextAnchor As Integer
Public Mainline As Boolean
Public Ignored As Boolean
Public NotAttached As Boolean
Public Bracket As Boolean
Public InstallAnchor As Boolean
Public reasonForMovement As String
Public reasonForAnchorMovement As String

Public Function Equals(other As Comm)
    If other Is Nothing Then
        Equals = False
    Else
        Equals = (Me.owner = other.owner) And _
                 (Me.height = other.height) And _
                 (Me.modification = other.modification)
    End If
End Function

