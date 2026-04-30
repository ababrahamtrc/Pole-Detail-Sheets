VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TensionCalculator_Form 
   Caption         =   "Secondary Tension Calculator"
   ClientHeight    =   3090
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6135
   OleObjectBlob   =   "TensionCalculator_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TensionCalculator_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Public Sub Initialize()
    On Error Resume Next
    
    ComboBox1.list = Array("4 TX")
    ComboBox2.list = Array("1'-0""", "1'-6""", "2'-0""", "2'-6""", "3'-0""", "3'-6""", "4'-0""", "4'-6""")
End Sub

Private Sub CommandButton1_Click()
    wireSize = ComboBox1.Value
    wireSag = ComboBox2.Value
    spanLength = TextBox1.Value
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.height) - (0.5 * Me.height)
    
    Dim xOffset As Double: xOffset = 0
    Dim yIntercept As Double: yIntercept = 0
    Dim slope As Double: slope = 0
    
    Select Case wireSize
        Case "4 TX"
            xOffset = 50
            Select Case wireSag
                Case "4'-6"""
                    yIntercept = 100
                    slope = 5.55555555556
                Case "4'-0"""
                    yIntercept = 116.666666666667
                    slope = 5.8
                Case "3'-6"""
                    yIntercept = 125
                    slope = 6.3333333333333
                Case "3'-0"""
                    yIntercept = 145
                    slope = 6.6666666666667
                Case "2'-6"""
                    yIntercept = 180
                    slope = 7.25
                Case "2'-0"""
                    yIntercept = 230
                    slope = 7.3
                Case "1'-6"""
                    yIntercept = 280
                    slope = 7.75
                Case "1'-0"""
                    yIntercept = 370
                    slope = 10
            End Select
    End Select
    
    If IsNumeric(spanLength) Then
        If spanLength > 0 Then
            MsgBox Round(((spanLength - xOffset) * slope) + yIntercept)
        End If
    Else
        MsgBox "Span length must be a valid number"
    End If
End Sub
