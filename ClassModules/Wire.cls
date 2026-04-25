VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Wire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public midspans As Scripting.Dictionary
Public tensions As Scripting.Dictionary
Public mrMoves As Scripting.Dictionary
Public adjacentHeights As Scripting.Dictionary
Public owner As String
Public size As String
Public height As Integer
Public modification As Integer
Public diameter As String
Public componentType As String
Public index As Integer
Public crossArm As String
Public trace As String 'for katapult import
Public wepHeight As Integer 'for spanguys
Public phase As Integer 'for primary

Private Sub Class_Initialize()
    Set midspans = New Scripting.Dictionary
    Set tensions = New Scripting.Dictionary
    Set mrMoves = New Scripting.Dictionary
    Set adjacentHeights = New Scripting.Dictionary
End Sub

Public Function getDisplayMidspan(Span As Span) As String
    Dim displayMidspan As String
    If Me.componentType = "PRI" Then
        displayMidspan = "PRI"
    ElseIf IsNumeric(Me.midspans(Span.spanSlot)) Then
        displayMidspan = Utilities.inchesToFeetInches(Me.midspans(Span.spanSlot))
    Else
        displayMidspan = Me.midspans(Span.spanSlot)
    End If
    
    If Span.houseNumber <> "" Then
        If displayMidspan = "0'0""" Or displayMidspan = "N/A" Then
            displayMidspan = Span.houseNumber
        Else
            displayMidspan = displayMidspan & "(" & Span.houseNumber & ")"
        End If
    End If
        
    If displayMidspan = "" Then displayMidspan = "0'0"""
    
    getDisplayMidspan = displayMidspan
End Function

Public Function getDisplayHeight(Optional category As Boolean) As String
    Dim displayHeight As String
    displayHeight = Utilities.inchesToFeetInches(Me.height)
    
    If category Then
        If Me.componentType = "DROP" Then
            displayHeight = displayHeight & " (DROP)"
        ElseIf Me.componentType = "MSG" Then
            displayHeight = displayHeight & " (MSG)"
        End If
    End If
    
    getDisplayHeight = displayHeight
End Function

