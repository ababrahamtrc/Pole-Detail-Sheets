VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Span"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public distance As Integer
Public angle As Integer
Public direction As String
Public spanId As String
Public spanSlot As Integer
Public houseNumber As String
Public otherPole As String
Public treeWork As Boolean
Public wires As Collection
Public utilWires As Collection
Public commWires  As Collection
Public utilComponents As Collection
Public commComponents  As Collection
Public spanGuys As Collection
Public lowestPowerWire As Wire

Private Sub Class_Initialize()
    Set wires = New Collection
    Set utilComponents = New Collection
    Set commComponents = New Collection
    Set utilWires = New Collection
    Set commWires = New Collection
    
    Set lowestPowerWires = New scripting.Dictionary
    For i = 1 To 12
        lowestPowerWires.Add i, Nothing
    Next i
End Sub

Public Sub setLowestPowerWire()
    Dim Wire As Wire
    For Each Wire In Me.utilWires
        If Wire.componentType <> "SPG" And Wire.componentType <> "PRI" Then
            If lowestPowerWire Is Nothing Then
                If IsNumeric(Wire.midspans(Me.spanSlot)) And Wire.midspans(Me.spanSlot) > 0 Then Set lowestPowerWire = Wire
            Else
                If lowestPowerWire.midspans(midspanSlot) < Wire.midspans(Me.spanSlot) And Wire.midspans(Me.spanSlot) > 0 Then Set lowestPowerWire = Wire
            End If
        End If
    Next Wire
End Sub

Public Function getToSpan() As String
    If Me.otherPole <> "" Then
        getToSpan = Me.otherPole & "(" & Me.angle & ")"
    Else
        getToSpan = Me.direction & "(" & Me.angle & ")"
    End If
End Function
