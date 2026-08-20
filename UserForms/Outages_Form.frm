VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Outages_Form 
   Caption         =   "Outages"
   ClientHeight    =   7365
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13695
   OleObjectBlob   =   "Outages_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Outages_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private locationTLMs As Scripting.Dictionary
Dim TextBoxInstances As Collection
Public finished As Boolean

Sub Initialize(ByRef locationTLMs_ As Scripting.Dictionary)
    Set locationTLMs = locationTLMs_
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.top = Application.top + (0.5 * Application.height) - (0.5 * Me.height)
    
    If locationTLMs.count > 0 Then
        LOC1.caption = "LOC " & locationTLMs.keys(0)
    End If
    
    Dim rowUnitHeight As Integer: rowUnitHeight = 30
    Dim colUnitHeight As Integer: colUnitHeight = 125
    
    For i = 1 To locationTLMs.count
        location = locationTLMs.keys(i - 1)
        Set tlms = locationTLMs(location)
        
        Dim row As Integer: row = ((i - 1) Mod 10)
        Dim col As Integer: col = ((i - 1) \ 10)
        
        Dim newBox As MSForms.TextBox: Set newBox = Me.Controls.Add("Forms.TextBox.1", "TextBox" & i, True)
        With newBox
            .Left = TextBox1.Left + (col * colUnitHeight)
            .top = TextBox1.top + (row * rowUnitHeight)
            .Width = TextBox1.Width
            .height = TextBox1.height
            .text = Utilities.JoinCollection(tlms, ",")
        End With
        
        Dim newLabel1 As MSForms.Label: Set newLabel1 = Me.Controls.Add("Forms.Label.1", "TLM" & i, True)
        With newLabel1
            .Left = TLM1.Left + (col * colUnitHeight)
            .top = TLM1.top + (row * rowUnitHeight)
            .Width = TLM1.Width
            .height = TLM1.height
            .caption = "TLM"
        End With
        
        Dim newLabel2 As MSForms.Label: Set newLabel2 = Me.Controls.Add("Forms.Label.1", "LOC" & i, True)
        With newLabel2
            .Left = LOC1.Left + (col * colUnitHeight)
            .top = LOC1.top + (row * rowUnitHeight)
            .Width = LOC1.Width
            .height = LOC1.height
            .caption = "LOC " & location
        End With
    Next i
    
    TextBox1.Locked = True
    TextBox1.visible = False
    TLM1.visible = False
    LOC1.visible = False
    
    Dim ctrl As control
    Dim obj As NumericTextBoxCls
    
    Set TextBoxInstances = New Collection
    
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is MSForms.TextBox Then
            Set obj = New NumericTextBoxCls
            Set obj.NumericTB = ctrl
            TextBoxInstances.Add obj
        End If
    Next ctrl
End Sub

Private Sub CommandButton1_Click()
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is MSForms.TextBox Then
            If ctrl.visible = True Then
                If Len(Utilities.OnlyNumbers(ctrl.text)) < 10 Then MsgBox "TLMs must be at least 10 digits long": Exit Sub
                tlms = Split(ctrl.text, ",")
                For Each tlm In tlms
                    If Len(tlm) < 10 Then MsgBox "TLMs must be at least 10 digits long": Exit Sub
                Next tlm
            End If
        End If
    Next ctrl
    
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is MSForms.TextBox Then
            If ctrl.visible = True Then
                index = Utilities.OnlyNumbers(ctrl.name)
                location = Replace(Me.Controls("LOC" & index).caption, "LOC ", "")
                tlms = Split(ctrl.text, ",")
                Dim tlmsUsed As Scripting.Dictionary: Set tlmsUsed = New Scripting.Dictionary
                Set locationTLMs(location) = New Collection
                For Each tlm In tlms
                    If Not tlmsUsed.exists(tlm) Then
                        tlmsUsed.Add tlm, Nothing
                        locationTLMs(location).Add tlm
                    End If
                Next tlm
            End If
        End If
    Next ctrl
    
    finished = True
    Me.Hide
End Sub
