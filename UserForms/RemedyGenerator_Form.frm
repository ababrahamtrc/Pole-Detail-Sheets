VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RemedyGenerator_Form 
   Caption         =   "Midspan Remedy Calculator"
   ClientHeight    =   15420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15765
   OleObjectBlob   =   "RemedyGenerator_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RemedyGenerator_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Dim sheet As Worksheet
Dim objects As Collection
Dim applicant As Wire
Dim wires As Collection
Dim weps As scripting.Dictionary
Dim clearanceMidspans As scripting.Dictionary
Dim OGClearanceMidspans As scripting.Dictionary
Dim spanNumber As Integer
Dim done As Boolean
Dim boltHoles As Collection
Dim uniqueHeights As scripting.Dictionary
Dim moddedMidspans As scripting.Dictionary
Dim automateMidspan As scripting.Dictionary
Dim isClosing As Boolean
Dim enterDone As Boolean
Dim powers As Collection
Dim OGPowers As Collection
Dim overlashJob As Boolean

Const POLE_COLOR As Long = &H40C0&     'RGB(192, 64, 0)
Const APPLICANT_WIRE_COLOR As Long = &HFFFF00    'RGB(0, 255, 255)
Const COMM_WIRE_COLOR As Long = &HFF0000  'RGB(0, 0, 255)
Const CLEARANCE_WIRE_COLOR As Long = &HFF00& 'RGB(0, 255, 0)
Const STREETLIGHT_COLOR As Long = &HFF0000  'RGB(0, 0, 255)
Const STREETLIGHT_DRIPLOOP_COLOR As Long = &HFFFF00  'RGB(0, 255, 255)
Const BOLTHOLE_COLOR As Long = &HFF&  'RGB(255, 0, 0)
Const VIOLATION_COLOR As Long = &HFF&  'RGB(0, 0, 255)

Const BOLTHOLE_TEXT_COLOR As Long = &H0&      'RGB(0, 0, 0)
Const COMM_WIRE_TEXT_COLOR As Long = &HFFFFFF  'RGB(255, 255, 255)
Const APPLICANT_WIRE_TEXT_COLOR As Long = &H0& 'RGB(0, 0, 0)
Const CLEARANCE_WIRE_TEXT_COLOR As Long = &H0& 'RGB(0, 0, 0)

Public Sub Initialize(sheet_ As Worksheet, powers_ As Collection, OGPowers_ As Collection, clearanceMidspans_ As scripting.Dictionary, OGClearanceMidspans_ As scripting.Dictionary, weps_ As scripting.Dictionary, applicant_ As Wire, wires_ As Collection, IgnoreBolt_ As Boolean, Bonded_ As Boolean, overlash As Boolean)
    
    On Error Resume Next
    
    enterDone = True
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.height) - (0.5 * Me.height)
    
    Set sheet = sheet_
    Set objects = New Collection
    Set applicant = applicant_
    Set automateMidspan = New scripting.Dictionary
    Set wires = wires_
    Set weps = weps_
    Set clearanceMidspans = clearanceMidspans_
    Set OGClearanceMidspans = OGClearanceMidspans_
    
    Pole1.BackColor = POLE_COLOR
    Pole2.BackColor = POLE_COLOR
    
    Set moddedMidspans = New scripting.Dictionary
    For Each midspanSlot In applicant.midspans
        Dim difference As Double
        difference = applicant.modification - applicant.height
        If applicant.adjacentHeights.Exists(midspanSlot) Then
            If applicant.adjacentHeights(midspanSlot).count > 0 Then
                difference = difference + applicant.adjacentHeights(midspanSlot)(2) - applicant.adjacentHeights(midspanSlot)(1)
            End If
        End If
        moddedMidspans.Add midspanSlot, applicant.midspans(midspanSlot) + Int(difference / 2)
    Next midspanSlot
    
    Set powers = powers_
    LWSTPWR.Value = Utilities.inchesToFeetInches(powers(1))
    STLTBRKT.Value = Utilities.inchesToFeetInches(powers(2))
    STLTDL.Value = Utilities.inchesToFeetInches(powers(3))
    If powers(2) > 0 Then TrimDL.visible = True
    
    Set OGPowers = OGPowers_
    OGLWSTPWR.Value = Utilities.inchesToFeetInches(OGPowers(1))
    OGSTLTBRKT.Value = Utilities.inchesToFeetInches(OGPowers(2))
    OGSTLTDL.Value = Utilities.inchesToFeetInches(OGPowers(3))
    
    Call createLabel(LWSTPWR.Value, UtilityMidspanClearance.name, CLEARANCE_WIRE_COLOR, CLEARANCE_WIRE_TEXT_COLOR, "SZV Clearance Att. Ht. " & LWSTPWRClearance & " Lowest Clearance Midspan In Span " & UtilityMidspanClearance.text, True)
    Call createBoundryLabel(LWSTPWR.Value, "Lowest Power", CLEARANCE_WIRE_COLOR, 6, " ")
    Call createBoundryLabel(STLTBRKT.Value, "Lowest Streetlight", STREETLIGHT_COLOR, 6, " ")
    Call createBoundryLabel(STLTDL.Value, "Lowest Streetlight DL", STREETLIGHT_DRIPLOOP_COLOR, 6, " ")
    
    Set boltHoles = New Collection
    Set uniqueHeights = New scripting.Dictionary
    For i = 1 To wires.count
        If Not uniqueHeights.Exists(wires(i).height) Then
            Call createBoundryLabel(Utilities.inchesToFeetInches(wires(i).height), i, BOLTHOLE_COLOR, textColor:=BOLTHOLE_TEXT_COLOR)
            uniqueHeights.Add wires(i).height, Nothing
        End If
    Next i
    
    If overlash Then
        Call createBoundryLabel(Utilities.inchesToFeetInches(applicant.height), i, BOLTHOLE_COLOR, textColor:=BOLTHOLE_TEXT_COLOR)
        uniqueHeights.Add applicant.height, Nothing
    End If

    For Each wep In weps
        Span.AddItem weps(wep)
        automateMidspan.Add wep, False
    Next wep
    
    If IgnoreBolt_ Then IgnoreBolt.Value = True
    If Bonded_ Then Bonded.Value = True
    
    Call IgnoreBolt_Change
    
    If Span.ListCount > 0 Then
        spanNumber = 0
        Span.ListIndex = 0
    End If
End Sub

Private Sub spanFill()
    On Error Resume Next
    
    Application.EnableEvents = False
    
    done = False
    
    If clearanceMidspans.Exists(spanNumber) Then
        UtilityMidspan.text = Utilities.inchesToFeetInches(clearanceMidspans(spanNumber))
    End If
    If OGClearanceMidspans.Exists(spanNumber) Then
        OGUtilityMidspan.text = Utilities.inchesToFeetInches(OGClearanceMidspans(spanNumber))
    End If
    
    If Not applicant Is Nothing Then
        If applicant.midspans.Exists(spanNumber) Then
            Height0.Value = Utilities.inchesToFeetInches(applicant.height)
            Mod0.Value = Utilities.inchesToFeetInches(applicant.modification)
            Midspan0.Value = Utilities.inchesToFeetInches(applicant.midspans(spanNumber))
            MidspanMod0.Value = Utilities.inchesToFeetInches(moddedMidspans(spanNumber))
            If applicant.adjacentHeights.Exists(spanNumber) Then
                If applicant.adjacentHeights(spanNumber).count > 0 Then
                    OtherOwner0.caption = "Applicant"
                    OtherHeight0.text = Utilities.inchesToFeetInches(applicant.adjacentHeights(spanNumber)(1))
                    OtherMod0.text = Utilities.inchesToFeetInches(applicant.adjacentHeights(spanNumber)(2))
                    Call makeOtherVisibile(0)
                End If
            End If
            Call makeVisibile(0)
        End If
    End If
    
    Dim i As Integer: i = 1
    For Each Wire In wires
        If Wire.midspans.Exists(spanNumber) Then
            Controls("Owner" & i).caption = Wire.owner
            Controls("Height" & i).Value = Utilities.inchesToFeetInches(Wire.height)
            Controls("Mod" & i).Value = Utilities.inchesToFeetInches(Wire.modification)
            Controls("Midspan" & i).Value = Utilities.inchesToFeetInches(Wire.midspans(spanNumber))
            If Wire.adjacentHeights.Exists(spanNumber) Then
                If Wire.adjacentHeights(spanNumber).count > 0 Then
                    Controls("OtherOwner" & i).caption = Wire.owner
                    Controls("OtherHeight" & i).text = Utilities.inchesToFeetInches(Wire.adjacentHeights(spanNumber)(1))
                    Controls("OtherMod" & i).text = Utilities.inchesToFeetInches(Wire.adjacentHeights(spanNumber)(2))
                    Call makeOtherVisibile(i)
                End If
            End If
            Wire.index = i
            Call makeVisibile(i)
            i = i + 1
        Else
            Wire.index = 0
        End If
    Next Wire
    
    Call Mod0_Change
    
    If Not applicant Is Nothing And moddedMidspans.Exists(spanNumber) Then
        If MidspanMod0.Value <> moddedMidspans(spanNumber) Then MidspanMod0.Value = Utilities.inchesToFeetInches(moddedMidspans(spanNumber))
    End If
    
    Call Mod1_Change
    Call Mod2_Change
    Call Mod3_Change
    Call Mod4_Change
    Call Mod5_Change
    Call Mod6_Change
    Call Mod7_Change
    Call Mod8_Change
    Call Mod9_Change
    Call Mod10_Change
    Call Mod11_Change
    Call Mod12_Change
    
    Call LWSTPWR_Change
    Call STLTBRKT_Change
    Call STLTDL_Change
    
    Dim modi As MSForms.TextBox
    For i = 1 To 12
        Call createLabel(Controls("Mod" & i).text, Controls("Mod" & i).name, COMM_WIRE_COLOR, COMM_WIRE_TEXT_COLOR, "Att. Ht.  " & Controls("Mod" & i).text & "  " & Controls("Owner" & i).caption & "  MS  " & Controls("MidspanMod" & i).text)
    Next i
    
    Call createLabel(Mod0.text, Mod0.name, APPLICANT_WIRE_COLOR, APPLICANT_WIRE_TEXT_COLOR, "Att. Ht.  " & Mod0.text & "  " & Owner0.caption & "  MS  " & MidspanMod0.text)
    
    Application.EnableEvents = True
    
    Previous = AutoApplicant.Value
    AutoApplicant.Value = automateMidspan(spanNumber)
    If AutoApplicant.Value Then Call automateApplicantMidspan
    
    done = True
    
    Call checkViolations
End Sub

Private Sub makeVisibile(i As Integer)
    
    On Error Resume Next
    
    Controls("Owner" & i).visible = True
    Controls("Height" & i).visible = True
    Controls("Mod" & i).visible = True
    Controls("Midspan" & i).visible = True
    Controls("MidspanMod" & i).visible = True
    Controls("FootUp" & i).visible = True
    Controls("FootDown" & i).visible = True
    Controls("InchUp" & i).visible = True
    Controls("InchDown" & i).visible = True
    Controls("HeightText" & i).visible = True
    Controls("ModText" & i).visible = True
    Controls("MidspanModText" & i).visible = True
    Controls("MidspanText" & i).visible = True
End Sub

Private Sub makeOtherVisibile(i As Integer)
    
    On Error Resume Next
    
    Controls("OtherOwner" & i).visible = True
    Controls("OtherHeight" & i).visible = True
    Controls("OtherMod" & i).visible = True
    Controls("OtherHeightText" & i).visible = True
    Controls("OtherModText" & i).visible = True
End Sub


Private Sub spanClear()
    
    On Error Resume Next
    
    done = False

    Application.EnableEvents = False

    For i = 0 To 12
        Controls("Owner" & i).visible = False
        Controls("Height" & i).visible = False
        If Controls("Height" & i).text <> "" Then Controls("Height" & i).text = ""
        Controls("Mod" & i).visible = False
        If Controls("Mod" & i).text <> "" Then Controls("Mod" & i).text = ""
        Controls("Midspan" & i).visible = False
        If Controls("Midspan" & i).text <> "" Then Controls("Midspan" & i).text = ""
        Controls("MidspanMod" & i).visible = False
        If Controls("MidspanMod" & i).text <> "" Then Controls("MidspanMod" & i).text = ""
        Controls("FootUp" & i).visible = False
        Controls("FootDown" & i).visible = False
        Controls("InchUp" & i).visible = False
        Controls("InchDown" & i).visible = False
        Controls("HeightText" & i).visible = False
        Controls("ModText" & i).visible = False
        Controls("MidspanModText" & i).visible = False
        Controls("MidspanText" & i).visible = False
        Controls("OtherOwner" & i).visible = False
        Controls("OtherHeight" & i).visible = False
        Controls("OtherMod" & i).visible = False
        Controls("OtherHeightText" & i).visible = False
        Controls("OtherModText" & i).visible = False
    Next i
    
    For Each Object In objects
        Me.Controls.Remove Object.name
    Next Object
    
    Set objects = New Collection
    
    If Controls("UtilityMidspan").text <> "" Then Controls("UtilityMidspan").text = ""
    
    Application.EnableEvents = True
    
    done = True
End Sub

Private Sub createLabel(ByVal height As String, ByVal name As String, ByVal color As Long, ByVal textColor As Long, ByVal midspan As String, Optional exclude As Boolean)
    
    On Error Resume Next
    
    inches = Utilities.convertToInches(height)
    If Not inches > 0 Then Exit Sub
    
    Dim lbl As MSForms.Label
    Set lbl = Controls.Add("Forms.Label.1", name & "Label", True)
    
    With lbl
        .Width = Pole2.Left - Pole1.Left - Pole2.Width
        .height = 10
        .BackColor = color
        .caption = midspan
        .TextAlign = fmTextAlignCenter
        .ForeColor = textColor
        .Font.size = 8
        .BorderStyle = fmBorderStyleSingle
        .SpecialEffect = fmSpecialEffectFlat
        .BorderColor = RGB(0, 0, 0)
    End With
    
    lbl.Move Pole1.Left + Pole1.Width, Pole1.Top + Pole1.height - inches
    
    If Not exclude Then objects.Add lbl
End Sub

Private Sub updateLabel(ByVal height As String, ByVal name As String, ByVal midspan As String, Optional index As Integer)
    
    On Error Resume Next
    
    If Not LabelExistsInUserForm(name & "Label") Then Exit Sub
    Set lbl = Controls(name & "Label")
    
    inches = Utilities.convertToInches(height)
    
    If inches > 0 Then
        For Each Wire In wires
            If Wire.index = index And index <> 0 And Wire.modification <> inches Then
                Wire.modification = inches
                Exit For
            End If
        Next Wire
    
        If lbl.caption <> midspan Then lbl.caption = midspan
        If Pole1.Top <> Pole1.Top + Pole1.height - inches Then
            lbl.Top = Pole1.Top + Pole1.height - inches
        End If
    End If
End Sub

Private Sub createBoundryLabel(ByVal height As String, ByVal name As String, ByVal color As Long, Optional size As Integer, Optional caption As String, Optional textColor As Long)
    
    On Error Resume Next
    
    inches = Utilities.convertToInches(height)
    If Not inches > 0 Then Exit Sub
    
    Dim lbl As MSForms.Label
    Set lbl = Controls.Add("Forms.Label.1", name & "Boundry", True)
    
    With lbl
        .Width = Pole1.Width
        .height = IIf(size <> 0, size, 8)
        .BackColor = color
        .TextAlign = fmTextAlignCenter
        .ForeColor = IIf(textColor > 0, textColor, RGB(0, 0, 0))
        .Font.size = IIf(size <> 0, size, 6)
        .caption = IIf(caption <> "", caption, height)
        .BorderStyle = fmBorderStyleSingle
        .SpecialEffect = fmSpecialEffectFlat
        .BorderColor = RGB(0, 0, 0)
    End With
    
    lbl.Move Pole1.Left, Pole1.Top + Pole1.height - inches
    If Not boltHoles Is Nothing Then boltHoles.Add lbl
    
    
End Sub

Private Sub updateBoundryLabel(ByVal height As String, ByVal name As String, Optional blank)
    
    On Error Resume Next
    
    If Not LabelExistsInUserForm(name & "Boundry") Then Exit Sub
    Set lbl = Controls(name & "Boundry")
    
    inches = Utilities.convertToInches(height)
    
    If inches > 0 Then
        If Not blank And lbl.caption <> height Then lbl.caption = height
        lbl.Move Pole1.Left, Pole1.Top + Pole1.height - inches
    End If
End Sub

Private Sub updateMidspan(ByVal i As Integer)
    
    On Error Resume Next
    
    If Utilities.convertToInches(Controls("Mod" & i).text) > 0 And Utilities.convertToInches(Controls("Height" & i).text) > 0 Then
        Dim newMidspan As String
        difference = Utilities.convertToInches(Controls("Mod" & i).text) - Utilities.convertToInches(Controls("Height" & i).text)
        If Controls("OtherHeight" & i).visible And Utilities.convertToInches(Controls("OtherHeight" & i).text) > 0 Then
            difference = difference + Utilities.convertToInches(Controls("OtherMod" & i).text) - Utilities.convertToInches(Controls("OtherHeight" & i).text)
        End If
        newMidspan = Utilities.inchesToFeetInches(Utilities.convertToInches(Controls("Midspan" & i).text) + Int(difference / 2))
        If Controls("MidspanMod" & i).text <> newMidspan Then
            Controls("MidspanMod" & i).text = newMidspan
        End If
    End If
End Sub

Private Function checkViolations(Optional report As Boolean) As String
    
    On Error Resume Next
    
    violationsReport = ""
    
    utilityMidspanClearanceInches = Utilities.convertToInches(UtilityMidspanClearance.text)
    utilityClearanceInches = Utilities.convertToInches(LWSTPWR.text)
    stltbrktClearanceInches = Utilities.convertToInches(STLTBRKT.text)
    STLTDLClearanceInches = Utilities.convertToInches(STLTDL.text)
    
    If Not utilityMidspanClearanceInches > 0 Or IgnoreClearance.Value Then utilityMidspanClearanceInches = 9999
    If Not utilityClearanceInches > 0 Or IgnoreClearance.Value Then utilityClearanceInches = 9999
    If Not stltbrktClearanceInches > 0 Then stltbrktClearanceInches = 9999
    If Not STLTDLClearanceInches > 0 Then stltbrktClearanceInches = 9999
    
    applicantSpan = Mod0.visible
    For i = IIf(applicantSpan, 0, 1) To 12
        If Not Controls("Mod" & i).visible Then Exit For
        inchesOne = Utilities.convertToInches(Controls("Mod" & i).text)
        MidspanOne = Utilities.convertToInches(Controls("MidspanMod" & i).text)
        violates = False
        inchesTwo = 0
        MidspanTwo = 0
        For j = IIf(applicantSpan, 0, 1) To 12
            If i <> j Then
                inchesTwo = Utilities.convertToInches(Controls("Mod" & j).text)
                MidspanTwo = Utilities.convertToInches(Controls("MidspanMod" & j).text)
                If inchesTwo > 0 And ((inchesOne >= inchesTwo) And ((MidspanOne - MidspanTwo) < 6)) Or ((inchesOne < inchesTwo) And ((MidspanTwo - MidspanOne) < 6)) Then
                    violates = True
                    violationsReport = violationsReport & "Wire" & i + IIf(applicantSpan, 1, 0) & " 6"" midspan separation with Wire" & j + IIf(applicantSpan, 1, 0) & vbLf
                End If
                If inchesTwo > 0 And Abs(inchesOne - inchesTwo) < 12 Then
                    violates = True
                    violationsReport = violationsReport & "Wire" & i + IIf(applicantSpan, 1, 0) & " 12"" separation with Wire" & j + IIf(applicantSpan, 1, 0) & vbLf
                End If
                If Not report And violates Then Exit For
            End If
        Next j
        If inchesOne < 186 Then
            violates = True
            violationsReport = violationsReport & "Wire" & i + IIf(applicantSpan, 1, 0) & " 15'6"" attach height violation" & vbLf
        ElseIf MidspanOne < 186 Then
            violates = True
            violationsReport = violationsReport & "Wire" & i + IIf(applicantSpan, 1, 0) & " 15'6"" midspan violation" & vbLf
        End If
        If Abs(inchesOne - utilityClearanceInches) < 40 Then
            violates = True
            violationsReport = violationsReport & "Wire" & i + IIf(applicantSpan, 1, 0) & " 40"" safety zone violation" & vbLf
        End If
        If Abs(inchesOne - stltbrktClearanceInches) < 40 And Not Bonded.Value Then
            violates = True
            violationsReport = violationsReport & "Wire" & i + IIf(applicantSpan, 1, 0) & " 40"" unbonded streetlight separation" & vbLf
        End If
        If Abs(inchesOne - stltbrktClearanceInches) < 4 And Bonded.Value Then
            violates = True
            violationsReport = violationsReport & "Wire" & i + IIf(applicantSpan, 1, 0) & " 4"" bonded streetlight separation" & vbLf
        End If
        If Abs(inchesOne - STLTDLClearanceInches) < 12 Then
            violates = True
            violationsReport = violationsReport & "Wire" & i + IIf(applicantSpan, 1, 0) & " 12"" streetlight driploop separation" & vbLf
        End If
        If MidspanOne > utilityMidspanClearanceInches Then
            violates = True
            violationsReport = violationsReport & "Wire" & i + IIf(applicantSpan, 1, 0) & " 30"" midspan separation vioaltion" & vbLf
        End If
        If Not IgnoreBolt.Value Then
            boltHoleViolation = False
            For Each heightInches In uniqueHeights
                If heightInches <> inchesOne Then
                    If Abs(inchesOne - heightInches) < 4 Then
                        boltHoleViolation = True
                    End If
                Else
                    boltHoleViolation = False
                    Exit For
                End If
            Next heightInches
            If boltHoleViolation Then
                violates = True
                violationsReport = violationsReport & "Wire" & i + IIf(applicantSpan, 1, 0) & " 4"" bolt hole separation violation" & vbLf
            End If
        End If
        
        If i > 0 Then
            For Each Wire In wires
                If Wire.index = i Then
                    For Each otherWire In wires
                        If otherWire.index = 0 And Wire.owner <> otherWire.owner Then
                            inchesOne = Wire.modification
                            inchesTwo = otherWire.modification
                            If Abs(inchesOne - inchesTwo) < 12 Then
                                violationsReport = violationsReport & "Wire" & i + IIf(applicantSpan, 1, 0) & " 12"" separation violation with wire in another span" & vbLf
                                violates = True
                            End If
                        End If
                    Next otherWire
                    If Not applicant Is Nothing Then
                        inchesOne = Wire.modification
                        inchesTwo = applicant.modification
                        If Abs(inchesOne - inchesTwo) < 12 Then
                                violationsReport = violationsReport & "Wire" & i + IIf(applicantSpan, 1, 0) & " 12"" separation violation with wire in another span" & vbLf
                                violates = True
                        End If
                    End If
                End If
            Next Wire
        Else
            If Not applicant Is Nothing And applicantSpan Then
                For Each otherWire In wires
                    If otherWire.index = 0 Then
                        inchesOne = applicant.modification
                        inchesTwo = otherWire.modification
                        If Abs(inchesOne - inchesTwo) < 12 Then
                            violates = True
                            violationsReport = violationsReport & "Wire1 12"" separation with wire in another span " & vbLf
                        End If
                    End If
                Next otherWire
            End If
        End If
        
        If Not report Then
            If violates Then
                Call updateLabelViolation("Mod" & i & "Label", True)
            Else
                Call updateLabelViolation("Mod" & i & "Label", False)
            End If
        End If
    Next i
    
    If violationsReport = "" Then violationsReport = "No violations found"
    
    checkViolations = violationsReport
End Function

Private Sub updateLabelViolation(name As String, violates As Boolean)
    
    On Error Resume Next
    
    If Not LabelExistsInUserForm(name) Then Exit Sub
    Set lbl = Controls(name)
    If violates Then
        If lbl.BackColor <> VIOLATION_COLOR And violates Then lbl.BackColor = VIOLATION_COLOR
    Else
        If name = "Mod0Label" Then
            If lbl.BackColor <> APPLICANT_WIRE_COLOR Then lbl.BackColor = APPLICANT_WIRE_COLOR
        Else
            If lbl.BackColor <> COMM_WIRE_COLOR Then lbl.BackColor = COMM_WIRE_COLOR
        End If
    End If
End Sub

Private Function LabelExistsInUserForm(labelName As String) As Boolean
    
    On Error Resume Next
    
    LabelExistsInUserForm = False

    For Each ctrl In Controls
        If TypeOf ctrl Is MSForms.Label Then
            If ctrl.name = labelName Then
                LabelExistsInUserForm = True
                Exit Function
            End If
        End If
    Next ctrl
End Function

Private Sub AutoApplicant_Change()
    
    On Error Resume Next

    If AutoApplicant.Value Then
        automateMidspan(spanNumber) = True
        Call automateApplicantMidspan
    Else
        automateMidspan(spanNumber) = False
        If applicant.midspans.Exists(spanNumber) Then
            Dim difference As Double
            difference = applicant.modification - applicant.height
            If applicant.adjacentHeights.Exists(spanNumber) Then
                If applicant.adjacentHeights(spanNumber).count > 0 Then
                    difference = difference + applicant.adjacentHeights(spanNumber)(2) - applicant.adjacentHeights(spanNumber)(1)
                End If
            End If
            MidspanMod0.text = Utilities.inchesToFeetInches(applicant.midspans(spanNumber) + Int(difference / 2))
        End If
    End If
End Sub

Private Sub automateApplicantMidspan()
    
   On Error Resume Next
    
    Dim heightBelow As Integer: heightBelow = 0
    Dim midspanbelow As Integer: midspanbelow = 0
    Dim heightAbove As Integer: heightAbove = 9999
    Dim midspanAbove As Integer: midspanAbove = 9999
    Dim applicantHeight As Integer: applicantHeight = Utilities.convertToInches(Controls("Mod0").text)

    For i = 1 To 12
        If Not Controls("MidspanMod" & i).visible Then Exit For
        attachHeight = Utilities.convertToInches(Controls("Mod" & i).text)
        midspanHeight = Utilities.convertToInches(Controls("MidspanMod" & i).text)
        If attachHeight <= applicantHeight And attachHeight >= heightBelow And midspanHeight >= midspanbelow Then
            heightBelow = attachHeight
            midspanbelow = midspanHeight
        ElseIf attachHeight > applicantHeight And attachHeight <= heightAbove And midspanHeight <= midspanAbove Then
            heightAbove = attachHeight
            midspanAbove = midspanHeight
        End If
    Next i
    If heightBelow > 0 And midspanbelow > 0 Then
        MidspanMod0.text = Utilities.inchesToFeetInches(midspanbelow + 6)
    ElseIf heightAbove < 9999 And midspanAbove < 9999 Then
        MidspanMod0.text = Utilities.inchesToFeetInches(midspanAbove - 6)
    End If
End Sub

Private Sub FootDown0_Click()
    Mod0.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod0.text) - 12)
End Sub

Private Sub FootUp0_Click()
    Mod0.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod0.text) + 12)
End Sub

Private Sub IgnoreBolt_Change()
    If done Then Call checkViolations
    For Each Object In boltHoles
        Object.visible = Not IgnoreBolt.Value
    Next Object
End Sub

Private Sub IgnoreClearance_Change()
    If done Then Call checkViolations
End Sub

Private Sub InchDown0_Click()
    Mod0.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod0.text) - 1)
End Sub

Private Sub InchUp0_Click()
    Mod0.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod0.text) + 1)
End Sub

Private Sub FootDown1_Click()
    Mod1.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod1.text) - 12)
End Sub

Private Sub FootUp1_Click()
    Mod1.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod1.text) + 12)
End Sub

Private Sub InchDown1_Click()
    Mod1.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod1.text) - 1)
End Sub

Private Sub InchUp1_Click()
    Mod1.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod1.text) + 1)
End Sub

Private Sub FootDown2_Click()
    Mod2.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod2.text) - 12)
End Sub

Private Sub FootUp2_Click()
    Mod2.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod2.text) + 12)
End Sub

Private Sub InchDown2_Click()
    Mod2.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod2.text) - 1)
End Sub

Private Sub InchUp2_Click()
    Mod2.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod2.text) + 1)
End Sub

Private Sub FootDown3_Click()
    Mod3.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod3.text) - 12)
End Sub

Private Sub FootUp3_Click()
    Mod3.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod3.text) + 12)
End Sub

Private Sub InchDown3_Click()
    Mod3.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod3.text) - 1)
End Sub

Private Sub InchUp3_Click()
    Mod3.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod3.text) + 1)
End Sub

Private Sub FootDown4_Click()
    Mod4.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod4.text) - 12)
End Sub

Private Sub FootUp4_Click()
    Mod4.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod4.text) + 12)
End Sub

Private Sub InchDown4_Click()
    Mod4.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod4.text) - 1)
End Sub

Private Sub InchUp4_Click()
    Mod4.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod4.text) + 1)
End Sub

Private Sub FootDown5_Click()
    Mod5.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod5.text) - 12)
End Sub

Private Sub FootUp5_Click()
    Mod5.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod5.text) + 12)
End Sub

Private Sub InchDown5_Click()
    Mod5.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod5.text) - 1)
End Sub

Private Sub InchUp5_Click()
    Mod5.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod5.text) + 1)
End Sub

Private Sub FootDown6_Click()
    Mod6.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod6.text) - 12)
End Sub

Private Sub FootUp6_Click()
    Mod6.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod6.text) + 12)
End Sub

Private Sub InchDown6_Click()
    Mod6.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod6.text) - 1)
End Sub

Private Sub InchUp6_Click()
    Mod6.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod6.text) + 1)
End Sub

Private Sub FootDown7_Click()
    Mod7.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod7.text) - 12)
End Sub

Private Sub FootUp7_Click()
    Mod7.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod7.text) + 12)
End Sub

Private Sub InchDown7_Click()
    Mod7.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod7.text) - 1)
End Sub

Private Sub InchUp7_Click()
    Mod7.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod7.text) + 1)
End Sub

Private Sub FootDown8_Click()
    Mod8.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod8.text) - 12)
End Sub

Private Sub FootUp8_Click()
    Mod8.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod8.text) + 12)
End Sub

Private Sub InchDown8_Click()
    Mod8.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod8.text) - 1)
End Sub

Private Sub InchUp8_Click()
    Mod8.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod8.text) + 1)
End Sub

Private Sub FootDown9_Click()
    Mod9.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod9.text) - 12)
End Sub

Private Sub FootUp9_Click()
    Mod9.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod9.text) + 12)
End Sub

Private Sub InchDown9_Click()
    Mod9.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod9.text) - 1)
End Sub

Private Sub InchUp9_Click()
    Mod9.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod9.text) + 1)
End Sub

Private Sub FootDown10_Click()
    Mod10.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod10.text) - 12)
End Sub

Private Sub FootUp10_Click()
    Mod10.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod10.text) + 12)
End Sub

Private Sub InchDown10_Click()
    Mod10.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod10.text) - 1)
End Sub

Private Sub InchUp10_Click()
    Mod10.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod10.text) + 1)
End Sub

Private Sub FootDown11_Click()
    Mod11.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod11.text) - 12)
End Sub

Private Sub FootUp11_Click()
    Mod11.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod11.text) + 12)
End Sub

Private Sub InchDown11_Click()
    Mod11.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod11.text) - 1)
End Sub

Private Sub InchUp11_Click()
    Mod11.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod11.text) + 1)
End Sub

Private Sub FootDown12_Click()
    Mod12.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod12.text) - 12)
End Sub

Private Sub FootUp12_Click()
    Mod12.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod12.text) + 12)
End Sub

Private Sub InchDown12_Click()
    Mod12.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod12.text) - 1)
End Sub

Private Sub InchUp12_Click()
    Mod12.text = Utilities.inchesToFeetInches(Utilities.convertToInches(Mod12.text) + 1)
End Sub

Private Sub FootDown13_Click()
    STLTBRKT.text = Utilities.inchesToFeetInches(Utilities.convertToInches(STLTBRKT.text) - 12)
End Sub

Private Sub FootUp13_Click()
    STLTBRKT.text = Utilities.inchesToFeetInches(Utilities.convertToInches(STLTBRKT.text) + 12)
End Sub

Private Sub InchDown13_Click()
    STLTBRKT.text = Utilities.inchesToFeetInches(Utilities.convertToInches(STLTBRKT.text) - 1)
End Sub

Private Sub InchUp13_Click()
    STLTBRKT.text = Utilities.inchesToFeetInches(Utilities.convertToInches(STLTBRKT.text) + 1)
End Sub

Private Sub FootDown14_Click()
    STLTDL.text = Utilities.inchesToFeetInches(Utilities.convertToInches(STLTDL.text) - 12)
End Sub

Private Sub FootUp14_Click()
    STLTDL.text = Utilities.inchesToFeetInches(Utilities.convertToInches(STLTDL.text) + 12)
End Sub

Private Sub InchDown14_Click()
    STLTDL.text = Utilities.inchesToFeetInches(Utilities.convertToInches(STLTDL.text) - 1)
End Sub

Private Sub InchUp14_Click()
    STLTDL.text = Utilities.inchesToFeetInches(Utilities.convertToInches(STLTDL.text) + 1)
End Sub

Private Sub FootDown15_Click()
    LWSTPWR.text = Utilities.inchesToFeetInches(Utilities.convertToInches(LWSTPWR.text) - 12)
End Sub

Private Sub FootUp15_Click()
    LWSTPWR.text = Utilities.inchesToFeetInches(Utilities.convertToInches(LWSTPWR.text) + 12)
End Sub

Private Sub InchDown15_Click()
    LWSTPWR.text = Utilities.inchesToFeetInches(Utilities.convertToInches(LWSTPWR.text) - 1)
End Sub

Private Sub InchUp15_Click()
    LWSTPWR.text = Utilities.inchesToFeetInches(Utilities.convertToInches(LWSTPWR.text) + 1)
End Sub

Private Sub FootDown16_Click()
    UtilityMidspan.text = Utilities.inchesToFeetInches(Utilities.convertToInches(UtilityMidspan.text) - 12)
End Sub

Private Sub FootUp16_Click()
    UtilityMidspan.text = Utilities.inchesToFeetInches(Utilities.convertToInches(UtilityMidspan.text) + 12)
End Sub

Private Sub InchDown16_Click()
    UtilityMidspan.text = Utilities.inchesToFeetInches(Utilities.convertToInches(UtilityMidspan.text) - 1)
End Sub

Private Sub InchUp16_Click()
    UtilityMidspan.text = Utilities.inchesToFeetInches(Utilities.convertToInches(UtilityMidspan.text) + 1)
End Sub

Private Sub LWSTPWR_Change()
    If Not Me.ActiveControl Is Me.LWSTPWR Then
        LWSTPWRClearance.text = Utilities.inchesToFeetInches(Utilities.convertToInches(LWSTPWR.text) - 40)
    End If
End Sub

Private Sub LWSTPWR_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    If keyCode = vbKeyReturn Then
        LWSTPWRClearance.text = Utilities.inchesToFeetInches(Utilities.convertToInches(LWSTPWR.text) - 40)
    End If
End Sub

Private Sub LWSTPWR_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not isClosing Then
        LWSTPWRClearance.text = Utilities.inchesToFeetInches(Utilities.convertToInches(LWSTPWR.text) - 40)
    End If
End Sub

Private Sub LWSTPWRClearance_Change()
    If done Then Call updateLabel(LWSTPWR.text, UtilityMidspanClearance.name, "SZV Clearance Att. Ht. " & LWSTPWRClearance & " Lowest Clearance Midspan In Span " & UtilityMidspanClearance.text)
    If done Then Call checkViolations
    If done Then Call updateBoundryLabel(LWSTPWR.text, "Lowest Power", True)
End Sub

Private Sub Span_Change()
    Dim foundWep As Integer: foundWep = 0
    For Each wep In weps
        If weps(wep) = Span.text Then
            foundWep = wep
            Exit For
        End If
    Next wep

    If foundWep > 0 Then
        SpanText.caption = "Span " & Span.ListIndex + 1
        spanNumber = foundWep
        Call spanClear
        Call spanFill
    End If
End Sub

Private Sub SpanDown_Click()
    If Span.ListCount > 0 Then
        If Span.ListIndex = 0 Then
            Span.ListIndex = Span.ListCount - 1
        Else
            Span.ListIndex = Span.ListIndex - 1
        End If
    End If
End Sub

Private Sub SpanUp_Click()
    If Span.ListCount > 0 Then
        If Span.ListIndex = Span.ListCount - 1 Then
            Span.ListIndex = 0
        Else
            Span.ListIndex = Span.ListIndex + 1
        End If
    End If
End Sub

Private Sub STLTBRKT_Change()
    If Not Me.ActiveControl Is Me.STLTBRKT Then
        STLTBRKTClearance.text = Utilities.inchesToFeetInches(Utilities.convertToInches(STLTBRKT.text) - IIf(Bonded.Value, 4, 40))
        If TrimDL Then STLTDL.Value = Utilities.inchesToFeetInches(Utilities.convertToInches(STLTBRKT) - 1)
    End If
End Sub

Private Sub STLTBRKT_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    If keyCode = vbKeyReturn Then
        STLTBRKTClearance.text = Utilities.inchesToFeetInches(Utilities.convertToInches(STLTBRKT.text) - IIf(Bonded.Value, 4, 40))
    End If
End Sub

Private Sub STLTBRKT_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not isClosing Then
        STLTBRKTClearance.text = Utilities.inchesToFeetInches(Utilities.convertToInches(STLTBRKT.text) - IIf(Bonded.Value, 4, 40))
    End If
End Sub

Private Sub Bonded_Click()
    STLTBRKTClearance.text = Utilities.inchesToFeetInches(Utilities.convertToInches(STLTBRKT.text) - IIf(Bonded.Value, 4, 40))
End Sub

Private Sub STLTBRKTClearance_Change()
    If done Then Call checkViolations
    If done Then Call updateBoundryLabel(STLTBRKT.text, "Lowest Streetlight", True)
End Sub

Private Sub STLTDL_Change()
    If Not Me.ActiveControl Is Me.STLTDL Then
        STLTDLClearance.text = Utilities.inchesToFeetInches(Utilities.convertToInches(STLTDL.text) - 12)
    End If
    If TrimDL Then STLTDL.Value = Utilities.inchesToFeetInches(Utilities.convertToInches(STLTBRKT) - 1)
End Sub

Private Sub STLTDL_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    If keyCode = vbKeyReturn Then
        STLTDLClearance.text = Utilities.inchesToFeetInches(Utilities.convertToInches(STLTDL.text) - 12)
    End If
End Sub

Private Sub STLTDL_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not isClosing Then
        STLTDLClearance.text = Utilities.inchesToFeetInches(Utilities.convertToInches(STLTDL.text) - 12)
    End If
End Sub

Private Sub STLTDLClearance_Change()
    If done Then Call checkViolations
    If done Then Call updateBoundryLabel(STLTDL.text, "Lowest Streetlight DL", True)
End Sub

Private Sub Mod0_Change()
    If Not Me.ActiveControl Is Me.Mod0 Then
        inches = Utilities.convertToInches(Mod0.text)
        If inches > 0 And applicant.modification <> inches Then
            applicant.modification = Utilities.convertToInches(Mod0.text)
            Dim difference As Double
            For Each midspanSlot In moddedMidspans
                If applicant.midspans.Exists(midspanSlot) Then
                    difference = applicant.modification - applicant.height
                    If applicant.adjacentHeights.Exists(midspanSlot) Then
                        If applicant.adjacentHeights(midspanSlot).count > 0 Then
                            difference = difference + applicant.adjacentHeights(midspanSlot)(2) - applicant.adjacentHeights(midspanSlot)(1)
                        End If
                    End If
                    newMidspan = applicant.midspans(midspanSlot) + Int(difference / 2)
                    moddedMidspans(midspanSlot) = newMidspan
                    If spanNumber = midspanSlot Then MidspanMod0.text = Utilities.inchesToFeetInches(newMidspan)
                End If
            Next midspanSlot
        End If
        If done Then Call updateLabel(Mod0.text, Mod0.name, "Att. Ht.  " & Mod0.text & "  " & Owner0.caption & "  MS  " & MidspanMod0.text)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod0_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    If keyCode = vbKeyReturn Then
        enterDone = False
        inches = Utilities.convertToInches(Mod0.text)
         If inches > 0 And applicant.modification <> inches Then
            applicant.modification = Utilities.convertToInches(Mod0.text)
            Dim difference As Double
            For Each midspanSlot In moddedMidspans
                difference = applicant.modification - applicant.height
                If applicant.adjacentHeights.Exists(midspanSlot) Then
                    If applicant.adjacentHeights(midspanSlot).count > 0 Then
                        difference = difference + applicant.adjacentHeights(midspanSlot)(2) - applicant.adjacentHeights(midspanSlot)(1)
                    End If
                End If
                newMidspan = applicant.midspans(midspanSlot) + Int(difference / 2)
                moddedMidspans(midspanSlot) = newMidspan
                If spanNumber = midspanSlot Then MidspanMod0.text = Utilities.inchesToFeetInches(newMidspan)
            Next midspanSlot
        End If
        If done Then Call updateLabel(Mod0.text, Mod0.name, "Att. Ht.  " & Mod0.text & "  " & Owner0.caption & "  MS  " & MidspanMod0.text)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod0_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not isClosing And enterDone Then
        inches = Utilities.convertToInches(Mod0.text)
         If inches > 0 And applicant.modification <> inches Then
            applicant.modification = inches
            Dim difference As Double
            For Each midspanSlot In moddedMidspans
                difference = applicant.modification - applicant.height
                If applicant.adjacentHeights.Exists(midspanSlot) Then
                    If applicant.adjacentHeights(midspanSlot).count > 0 Then
                        difference = difference + applicant.adjacentHeights(midspanSlot)(2) - applicant.adjacentHeights(midspanSlot)(1)
                    End If
                End If
                newMidspan = applicant.midspans(midspanSlot) + Int(difference / 2)
                moddedMidspans(midspanSlot) = newMidspan
                If spanNumber = midspanSlot Then MidspanMod0.text = Utilities.inchesToFeetInches(newMidspan)
            Next midspanSlot
        End If
        If done Then Call updateLabel(Mod0.text, Mod0.name, "Att. Ht.  " & Mod0.text & "  " & Owner0.caption & "  MS  " & MidspanMod0.text)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
    enterDone = True
End Sub

Private Sub MidspanMod0_Change()
    If Not Me.ActiveControl Is Me.MidspanMod0 And Utilities.convertToInches(MidspanMod0.text) <> moddedMidspans(spanNumber) Then
        Call updateLabel(Mod0.text, Mod0.name, "Att. Ht.  " & Mod0.text & "  " & Owner0.caption & "  MS  " & MidspanMod0.text)
        If done Then Call checkViolations
        If done And moddedMidspans.Exists(spanNumber) And Utilities.convertToInches(MidspanMod0.text) > 0 Then moddedMidspans(spanNumber) = Utilities.convertToInches(MidspanMod0.text)
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub MidspanMod0_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    If keyCode = vbKeyReturn And Utilities.convertToInches(MidspanMod0.text) <> moddedMidspans(spanNumber) Then
        enterDone = False
        Call updateLabel(Mod0.text, Mod0.name, "Att. Ht.  " & Mod0.text & "  " & Owner0.caption & "  MS  " & MidspanMod0.text)
        If done Then Call checkViolations
        If done And moddedMidspans.Exists(spanNumber) And Utilities.convertToInches(MidspanMod0.text) > 0 Then moddedMidspans(spanNumber) = Utilities.convertToInches(MidspanMod0.text)
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub MidspanMod0_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not isClosing And enterDone And Utilities.convertToInches(MidspanMod0.text) <> moddedMidspans(spanNumber) Then
        Call updateLabel(Mod0.text, Mod0.name, "Att. Ht.  " & Mod0.text & "  " & Owner0.caption & "  MS  " & MidspanMod0.text)
        If done Then Call checkViolations
        If done And moddedMidspans.Exists(spanNumber) And Utilities.convertToInches(MidspanMod0.text) > 0 Then moddedMidspans(spanNumber) = Utilities.convertToInches(MidspanMod0.text)
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
    enterDone = True
End Sub

Private Sub Mod1_Change()
    If Not Me.ActiveControl Is Me.Mod1 Then
        Call updateMidspan(1)
        If done Then Call updateLabel(Mod1.text, Mod1.name, "Att. Ht.  " & Mod1.text & "  " & Owner1.caption & "  MS  " & MidspanMod1.text, 1)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod1_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    If keyCode = vbKeyReturn Then
        Call updateMidspan(1)
        If done Then Call updateLabel(Mod1.text, Mod1.name, "Att. Ht.  " & Mod1.text & "  " & Owner1.caption & "  MS  " & MidspanMod1.text, 1)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not isClosing Then
        Call updateMidspan(1)
        If done Then Call updateLabel(Mod1.text, Mod1.name, "Att. Ht.  " & Mod1.text & "  " & Owner1.caption & "  MS  " & MidspanMod1.text, 1)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod2_Change()
    If Not Me.ActiveControl Is Me.Mod2 Then
        Call updateMidspan(2)
        If done Then Call updateLabel(Mod2.text, Mod2.name, "Att. Ht.  " & Mod2.text & "  " & Owner2.caption & "  MS  " & MidspanMod2.text, 2)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod2_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    If keyCode = vbKeyReturn Then
        Call updateMidspan(2)
        If done Then Call updateLabel(Mod2.text, Mod2.name, "Att. Ht.  " & Mod2.text & "  " & Owner2.caption & "  MS  " & MidspanMod2.text, 2)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not isClosing Then
        Call updateMidspan(2)
        If done Then Call updateLabel(Mod2.text, Mod2.name, "Att. Ht.  " & Mod2.text & "  " & Owner2.caption & "  MS  " & MidspanMod2.text, 2)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod3_Change()
    If Not Me.ActiveControl Is Me.Mod3 Then
        Call updateMidspan(3)
        If done Then Call updateLabel(Mod3.text, Mod3.name, "Att. Ht.  " & Mod3.text & "  " & Owner3.caption & "  MS  " & MidspanMod3.text, 3)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod3_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    If keyCode = vbKeyReturn Then
        Call updateMidspan(3)
        If done Then Call updateLabel(Mod3.text, Mod3.name, "Att. Ht.  " & Mod3.text & "  " & Owner3.caption & "  MS  " & MidspanMod3.text, 3)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not isClosing Then
        Call updateMidspan(3)
        If done Then Call updateLabel(Mod3.text, Mod3.name, "Att. Ht.  " & Mod3.text & "  " & Owner3.caption & "  MS  " & MidspanMod3.text, 3)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod4_Change()
    If Not Me.ActiveControl Is Me.Mod4 Then
        Call updateMidspan(4)
        If done Then Call updateLabel(Mod4.text, Mod4.name, "Att. Ht.  " & Mod4.text & "  " & Owner4.caption & "  MS  " & MidspanMod4.text, 4)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod4_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    If keyCode = vbKeyReturn Then
        Call updateMidspan(4)
        If done Then Call updateLabel(Mod4.text, Mod4.name, "Att. Ht.  " & Mod4.text & "  " & Owner4.caption & "  MS  " & MidspanMod4.text, 4)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not isClosing Then
        Call updateMidspan(4)
        If done Then Call updateLabel(Mod4.text, Mod4.name, "Att. Ht.  " & Mod4.text & "  " & Owner4.caption & "  MS  " & MidspanMod4.text, 4)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod5_Change()
    If Not Me.ActiveControl Is Me.Mod5 Then
        Call updateMidspan(5)
        If done Then Call updateLabel(Mod5.text, Mod5.name, "Att. Ht.  " & Mod5.text & "  " & Owner5.caption & "  MS  " & MidspanMod5.text, 5)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod5_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    If keyCode = vbKeyReturn Then
        Call updateMidspan(5)
        If done Then Call updateLabel(Mod5.text, Mod5.name, "Att. Ht.  " & Mod5.text & "  " & Owner5.caption & "  MS  " & MidspanMod5.text, 5)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod5_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not isClosing Then
        Call updateMidspan(5)
        If done Then Call updateLabel(Mod5.text, Mod5.name, "Att. Ht.  " & Mod5.text & "  " & Owner5.caption & "  MS  " & MidspanMod5.text, 5)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod6_Change()
    If Not Me.ActiveControl Is Me.Mod6 Then
        Call updateMidspan(6)
        If done Then Call updateLabel(Mod6.text, Mod6.name, "Att. Ht.  " & Mod6.text & "  " & Owner6.caption & "  MS  " & MidspanMod6.text, 6)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod6_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    If keyCode = vbKeyReturn Then
        Call updateMidspan(6)
        If done Then Call updateLabel(Mod6.text, Mod6.name, "Att. Ht.  " & Mod6.text & "  " & Owner6.caption & "  MS  " & MidspanMod6.text, 6)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod6_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not isClosing Then
        Call updateMidspan(6)
        If done Then Call updateLabel(Mod6.text, Mod6.name, "Att. Ht.  " & Mod6.text & "  " & Owner6.caption & "  MS  " & MidspanMod6.text, 6)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod7_Change()
    If Not Me.ActiveControl Is Me.Mod7 Then
        Call updateMidspan(7)
        If done Then Call updateLabel(Mod7.text, Mod7.name, "Att. Ht.  " & Mod7.text & "  " & Owner7.caption & "  MS  " & MidspanMod7.text, 7)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod7_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    If keyCode = vbKeyReturn Then
        Call updateMidspan(7)
        If done Then Call updateLabel(Mod7.text, Mod7.name, "Att. Ht.  " & Mod7.text & "  " & Owner7.caption & "  MS  " & MidspanMod7.text, 7)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod7_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not isClosing Then
        Call updateMidspan(7)
        If done Then Call updateLabel(Mod7.text, Mod7.name, "Att. Ht.  " & Mod7.text & "  " & Owner7.caption & "  MS  " & MidspanMod7.text, 7)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod8_Change()
    If Not Me.ActiveControl Is Me.Mod8 Then
        Call updateMidspan(8)
        If done Then Call updateLabel(Mod8.text, Mod8.name, "Att. Ht.  " & Mod8.text & "  " & Owner8.caption & "  MS  " & MidspanMod8.text, 8)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod8_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    If keyCode = vbKeyReturn Then
        Call updateMidspan(8)
        If done Then Call updateLabel(Mod8.text, Mod8.name, "Att. Ht.  " & Mod8.text & "  " & Owner8.caption & "  MS  " & MidspanMod8.text, 8)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod8_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not isClosing Then
        Call updateMidspan(8)
        If done Then Call updateLabel(Mod8.text, Mod8.name, "Att. Ht.  " & Mod8.text & "  " & Owner8.caption & "  MS  " & MidspanMod8.text, 8)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod9_Change()
    If Not Me.ActiveControl Is Me.Mod9 Then
        Call updateMidspan(9)
        If done Then Call updateLabel(Mod9.text, Mod9.name, "Att. Ht.  " & Mod9.text & "  " & Owner9.caption & "  MS  " & MidspanMod9.text, 9)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod9_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    If keyCode = vbKeyReturn Then
        Call updateMidspan(9)
        If done Then Call updateLabel(Mod9.text, Mod9.name, "Att. Ht.  " & Mod9.text & "  " & Owner9.caption & "  MS  " & MidspanMod9.text, 9)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod9_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not isClosing Then
        Call updateMidspan(9)
        If done Then Call updateLabel(Mod9.text, Mod9.name, "Att. Ht.  " & Mod9.text & "  " & Owner9.caption & "  MS  " & MidspanMod9.text, 9)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod10_Change()
    If Not Me.ActiveControl Is Me.Mod10 Then
        Call updateMidspan(10)
        If done Then Call updateLabel(Mod10.text, Mod10.name, "Att. Ht.  " & Mod10.text & "  " & Owner10.caption & "  MS  " & MidspanMod10.text, 10)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod10_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    If keyCode = vbKeyReturn Then
        Call updateMidspan(10)
        If done Then Call updateLabel(Mod10.text, Mod10.name, "Att. Ht.  " & Mod10.text & "  " & Owner10.caption & "  MS  " & MidspanMod10.text, 10)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod10_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not isClosing Then
        Call updateMidspan(10)
        If done Then Call updateLabel(Mod10.text, Mod10.name, "Att. Ht.  " & Mod10.text & "  " & Owner10.caption & "  MS  " & MidspanMod10.text, 10)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod11_Change()
    If Not Me.ActiveControl Is Me.Mod11 Then
        Call updateMidspan(11)
        If done Then Call updateLabel(Mod11.text, Mod11.name, "Att. Ht.  " & Mod11.text & "  " & Owner11.caption & "  MS  " & MidspanMod11.text, 11)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod11_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    If keyCode = vbKeyReturn Then
        Call updateMidspan(11)
        If done Then Call updateLabel(Mod11.text, Mod11.name, "Att. Ht.  " & Mod11.text & "  " & Owner11.caption & "  MS  " & MidspanMod11.text, 11)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod11_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not isClosing Then
        Call updateMidspan(11)
        If done Then Call updateLabel(Mod11.text, Mod11.name, "Att. Ht.  " & Mod11.text & "  " & Owner11.caption & "  MS  " & MidspanMod11.text, 11)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod12_Change()
    If Not Me.ActiveControl Is Me.Mod12 Then
        Call updateMidspan(12)
        If done Then Call updateLabel(Mod12.text, Mod12.name, "Att. Ht.  " & Mod12.text & "  " & Owner12.caption & "  MS  " & MidspanMod12.text, 12)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod12_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    If keyCode = vbKeyReturn Then
        Call updateMidspan(12)
        If done Then Call updateLabel(Mod12.text, Mod12.name, "Att. Ht.  " & Mod12.text & "  " & Owner12.caption & "  MS  " & MidspanMod12.text, 12)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub Mod12_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not isClosing Then
        Call updateMidspan(12)
        If done Then Call updateLabel(Mod12.text, Mod12.name, "Att. Ht.  " & Mod12.text & "  " & Owner12.caption & "  MS  " & MidspanMod12.text, 12)
        If done Then Call checkViolations
        If done And AutoApplicant.Value Then Call automateApplicantMidspan
    End If
End Sub

Private Sub TrimDL_Click()
    If TrimDL Then STLTDL.Value = Utilities.inchesToFeetInches(Utilities.convertToInches(STLTBRKT) - 1)
    If Not TrimDL Then STLTDL.Value = OGSTLTDL.Value
End Sub

Private Sub UpdateMods_Click()
    
    On Error Resume Next
    
    Dim comms As Collection: Set comms = New Collection
    comms.Add "COMM1"
    comms.Add "COMM2"
    comms.Add "COMM3"
    comms.Add "COMM4"
    comms.Add "COMM5"
    comms.Add "COMM6"
    comms.Add "COMM7"
    comms.Add "COMM8"
    
    answer = MsgBox("Do you wish to overwrite existing values on excel pole detail sheet?", vbYesNoCancel + vbQuestion, "Confirmation")
    If answer = vbYes Then
        Call RemedyGen.clearModCells(sheet)
        If InStr(sheet.Range("PROPOSEDHEIGHT"), "OL") = 0 Then
            If sheet.Range("CMRF1") <> Utilities.inchesToFeetInches(applicant.modification) Then sheet.Range("CMRF1") = Utilities.inchesToFeetInches(applicant.modification)
        Else
            If sheet.Range("CMRF1") <> "OL @ " & Utilities.inchesToFeetInches(applicant.modification) Then sheet.Range("CMRF1") = "OL @ " & Utilities.inchesToFeetInches(applicant.modification)
            Set modCell = findModCell(sheet, applicant.height, applicant.owner, True)
            If Not modCell Is Nothing Then
                If Utilities.convertToInches(modCell.text) < 1 And applicant.modification <> applicant.height Then modCell.Value = Utilities.inchesToFeetInches(applicant.modification)
                If Utilities.convertToInches(modCell.text) > 0 And applicant.modification > 0 And Utilities.convertToInches(modCell.text) <> applicant.modification Then
                    If Utilities.convertToInches(applicant.modification) <> Utilities.convertToInches(applicant.height) Then
                        modCell.Value = Utilities.inchesToFeetInches(applicant.modification)
                    Else
                        modCell.Value = ""
                    End If
                End If
            End If
        End If
        Dim midspanText As String: midspanText = ""
        Set regex = CreateObject("VBScript.RegExp")
        With regex
            .Pattern = "^(.*?)(?=\()"
            .IgnoreCase = True
            .Global = False
        End With
        polesChanged = ""
        For Each midspanSlot In applicant.midspans
            If regex.test(weps(midspanSlot)) Then
                Set matches = regex.Execute(weps(midspanSlot))
                Dim direction As String: direction = ThisWorkbook.RemoveParentheses(Trim(matches(0)))
                If SheetExists(direction) Then
                    Set otherSheet = Utilities.GetPDS(direction)
                    
                    If InStr(otherSheet.Range("CMRF2").Value, "@P" & sheet.Range("POLENUM")) > 0 Then
                        answer = MsgBox("Do you want to update the value of midspans on pole " & direction, vbYesNoCancel + vbQuestion, "Confirmation")
                        If answer = vbYes Then
                            Dim lines() As String
                            lines = Split(otherSheet.Range("CMRF2").Value, vbLf)
                            For i = LBound(lines) To UBound(lines)
                                If InStr(lines(i), " @P" & sheet.Range("POLENUM")) > 0 Then
                                    lines(i) = Utilities.inchesToFeetInches(moddedMidspans(midspanSlot)) & " @P" & sheet.Range("POLENUM")
                                End If
                            Next i
                            otherSheet.Range("CMRF2").Value = Join(lines, vbLf)
                            If polesChanged <> "" Then polesChanged = polesChanged & ", "
                            polesChanged = polesChanged & direction
                        End If
                    Else
                        If midspanText <> "" Then midspanText = midspanText & vbLf
                        midspanText = midspanText & Utilities.inchesToFeetInches(moddedMidspans(midspanSlot)) & " @P" & direction
                    End If
                Else
                    If direction <> "N" And direction <> "E" And direction <> "S" And direction <> "W" And direction <> "NE" And direction <> "SE" And direction <> "SW" And direction <> "NW" Then direction = "@P" & direction
                    If Trim(Replace(sheet.Range("TOPOLE" & midspanSlot).offset(1, 0), "-", "")) <> "" Then
                        If midspanText <> "" Then midspanText = midspanText & vbLf
                        midspanText = midspanText & Utilities.inchesToFeetInches(moddedMidspans(midspanSlot)) & " " & direction
                    End If
                End If
            End If
        Next midspanSlot
        
        If sheet.Range("CMRF2") <> midspanText Then sheet.Range("CMRF2") = midspanText
        
        For Each Wire In wires
            Set modCell = findModCell(sheet, Wire.height, Wire.owner, True)
            If Not modCell Is Nothing Then
                If Wire.modification <> Wire.height Then
                    modCell.Value = Utilities.inchesToFeetInches(Wire.modification)
                Else
                    modCell.Value = Utilities.inchesToFeetInches(Wire.height)
                End If
            End If
        Next Wire
        
        If LWSTPWR.text <> "" And powers.count > 0 Then
            If LWSTPWR.text <> powers(1) Then
                newPower = ""
                If sheet.Range("LWSTPWR").Value <> "" Then newPower = Split(sheet.Range("LWSTPWR").Value, vbLf)(0)
                If newPower <> "" Then newPower = newPower & vbLf
                sheet.Range("LWSTPWR").Value = newPower & LWSTPWR.text
            End If
        End If
        
        If STLTBRKT.text <> "" And STLTBRKT.text <> "N/A" And powers.count > 1 Then
            If STLTBRKT.text <> powers(2) Then
                newPower = ""
                If sheet.Range("STLTBRKT").Value <> "" Then newPower = Split(sheet.Range("STLTBRKT").Value, vbLf)(0)
                If newPower <> "" Then newPower = newPower & vbLf
                sheet.Range("STLTBRKT").Value = newPower & STLTBRKT.text
            End If
        End If
        
        If STLTDL.text <> "" And STLTDL.text <> "N/A" And powers.count > 2 Then
            If STLTDL.text <> powers(3) Then
                newPower = ""
                If sheet.Range("STLTDL").Value <> "" Then newPower = Split(sheet.Range("STLTDL").Value, vbLf)(0)
                If newPower <> "" Then newPower = newPower & vbLf
                sheet.Range("STLTDL").Value = newPower & STLTDL.text
            End If
        End If
        
        For Each clearanceMidspan In weps
            If clearanceMidspan > 0 Then
                If clearanceMidspans.Exists(clearanceMidspan) Then
                    If sheet.Range("CMMIDSPAN" & clearanceMidspan).offset(-2, 0).Value <> "" Then
                        If Split(sheet.Range("CMMIDSPAN" & clearanceMidspan).offset(-2, 0).Value, vbLf)(0) <> Utilities.inchesToFeetInches(clearanceMidspans(clearanceMidspan) - 30) Then
                            newPower = ""
                            newPower = Split(sheet.Range("CMMIDSPAN" & clearanceMidspan).offset(-2, 0).Value, vbLf)(0)
                            If newPower <> "" Then newPower = newPower & vbLf
                            sheet.Range("CMMIDSPAN" & clearanceMidspan).offset(-2, 0).Value = newPower & Utilities.inchesToFeetInches(clearanceMidspans(clearanceMidspan) - 30)
                        End If
                    End If
                End If
            End If
        Next clearanceMidspan
        
        Unload Me
    End If
End Sub
Private Sub UtilityMidspan_Change()
    If Not Me.ActiveControl Is Me.UtilityMidspan Then
        UtilityMidspanClearance.text = Utilities.inchesToFeetInches(Utilities.convertToInches(UtilityMidspan.text) - 30)
    End If
End Sub

Private Sub UtilityMidspan_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    If keyCode = vbKeyReturn Then
        UtilityMidspanClearance.text = Utilities.inchesToFeetInches(Utilities.convertToInches(UtilityMidspan.text) - 30)
    End If
End Sub

Private Sub UtilityMidspan_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not isClosing Then
        UtilityMidspanClearance.text = Utilities.inchesToFeetInches(Utilities.convertToInches(UtilityMidspan.text) - 30)
    End If
End Sub

Private Sub UtilityMidspanClearance_Change()
    Call updateLabel(LWSTPWR.text, UtilityMidspanClearance.name, "SZV Clearance Att. Ht. " & LWSTPWRClearance & " Lowest Clearance Midspan In Span " & UtilityMidspanClearance.text)
    If done Then clearanceMidspans(spanNumber) = Utilities.convertToInches(UtilityMidspan.Value)
    If done Then Call checkViolations
End Sub

Private Sub Violations_Click()
    MsgBox checkViolations(True)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    isClosing = True
    
    Set sheet = Nothing
    Set objects = Nothing
    Set applicant = Nothing
    Set wires = Nothing
    Set weps = Nothing
    Set clearanceMidspans = Nothing
    Set boltHoles = Nothing
    Set uniqueHeights = Nothing
    Set moddedMidspans = Nothing
    Set automateMidspan = Nothing
End Sub

