VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public location As String
Public code As String
Public qty As Integer
Public action As String
Public parentCU As String
Public parentInstance As Integer
Public childCode As String
Public childQty As Integer

Public Function Equals(other As cu)
    If other Is Nothing Then
        Equals = False
    Else
        Equals = (Me.code = other.code) And _
                 (Me.action = other.action) And _
                 (Me.parentCU = other.parentCU) And _
                 (Me.parentInstance = other.parentInstance) And _
                 (Me.childCode = other.childCode)
    End If
End Function

