VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NJUNS_Form 
   Caption         =   "NJUNS Movements Generator"
   ClientHeight    =   7290
   ClientLeft      =   270
   ClientTop       =   1035
   ClientWidth     =   24390
   OleObjectBlob   =   "NJUNS_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NJUNS_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private companies As Collection
Private heights As Collection
Private modifications As Collection
Private dgaprevious As Collection
Private dgaafter As Collection
Private rpguy As Collection
Private MoveAnchor As Collection
Private drops As Collection
Private mainlines As Collection
Private Ignored As Collection
Private NotAttached As Collection
Private Boxed As Collection
Private Bracket As Collection
Private InstallAnchor As Collection
Private reasonForMovement As Collection
Private reasonForAnchorMovement As Collection
Private sheet As Worksheet

Private comms As Collection
Private ticketType As String

Sub Initialize()
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.height) - (0.5 * Me.height)
    Set companies = New Collection
    companies.Add CommCompany1
    companies.Add CommCompany2
    companies.Add CommCompany3
    companies.Add CommCompany4
    companies.Add CommCompany5
    companies.Add CommCompany6
    companies.Add CommCompany7
    companies.Add CommCompany8
    
    Set heights = New Collection
    heights.Add CommHeight1
    heights.Add CommHeight2
    heights.Add CommHeight3
    heights.Add CommHeight4
    heights.Add CommHeight5
    heights.Add CommHeight6
    heights.Add CommHeight7
    heights.Add CommHeight8
    
    Set modifications = New Collection
    modifications.Add CommModification1
    modifications.Add CommModification2
    modifications.Add CommModification3
    modifications.Add CommModification4
    modifications.Add CommModification5
    modifications.Add CommModification6
    modifications.Add CommModification7
    modifications.Add CommModification8
    
    Set MoveAnchor = New Collection
    MoveAnchor.Add CommMoveAnchor1
    MoveAnchor.Add CommMoveAnchor2
    MoveAnchor.Add CommMoveAnchor3
    MoveAnchor.Add CommMoveAnchor4
    MoveAnchor.Add CommMoveAnchor5
    MoveAnchor.Add CommMoveAnchor6
    MoveAnchor.Add CommMoveAnchor7
    MoveAnchor.Add CommMoveAnchor8
    
    Set rpguy = New Collection
    rpguy.Add CommReplaceAnchor1
    rpguy.Add CommReplaceAnchor2
    rpguy.Add CommReplaceAnchor3
    rpguy.Add CommReplaceAnchor4
    rpguy.Add CommReplaceAnchor5
    rpguy.Add CommReplaceAnchor6
    rpguy.Add CommReplaceAnchor7
    rpguy.Add CommReplaceAnchor8
    
    Set drops = New Collection
    drops.Add CommDrops1
    drops.Add CommDrops2
    drops.Add CommDrops3
    drops.Add CommDrops4
    drops.Add CommDrops5
    drops.Add CommDrops6
    drops.Add CommDrops7
    drops.Add CommDrops8
    
    Set dgaprevious = New Collection
    dgaprevious.Add CommDGAPrevious1
    dgaprevious.Add CommDGAPrevious2
    dgaprevious.Add CommDGAPrevious3
    dgaprevious.Add CommDGAPrevious4
    dgaprevious.Add CommDGAPrevious5
    dgaprevious.Add CommDGAPrevious6
    dgaprevious.Add CommDGAPrevious7
    dgaprevious.Add CommDGAPrevious8
    
    Set dgaafter = New Collection
    dgaafter.Add CommDGAAfter1
    dgaafter.Add CommDGAAfter2
    dgaafter.Add CommDGAAfter3
    dgaafter.Add CommDGAAfter4
    dgaafter.Add CommDGAAfter5
    dgaafter.Add CommDGAAfter6
    dgaafter.Add CommDGAAfter7
    dgaafter.Add CommDGAAfter8
    
    Set mainlines = New Collection
    mainlines.Add Mainline1
    mainlines.Add Mainline2
    mainlines.Add Mainline3
    mainlines.Add Mainline4
    mainlines.Add Mainline5
    mainlines.Add Mainline6
    mainlines.Add Mainline7
    mainlines.Add Mainline8
    
    Set Ignored = New Collection
    Ignored.Add Ignore1
    Ignored.Add Ignore2
    Ignored.Add Ignore3
    Ignored.Add Ignore4
    Ignored.Add Ignore5
    Ignored.Add Ignore6
    Ignored.Add Ignore7
    Ignored.Add Ignore8
    
    Set NotAttached = New Collection
    NotAttached.Add NotAttached1
    NotAttached.Add NotAttached2
    NotAttached.Add NotAttached3
    NotAttached.Add NotAttached4
    NotAttached.Add NotAttached5
    NotAttached.Add NotAttached6
    NotAttached.Add NotAttached7
    NotAttached.Add NotAttached8
    
    Set Boxed = New Collection
    Boxed.Add Boxed1
    Boxed.Add Boxed2
    Boxed.Add Boxed3
    Boxed.Add Boxed4
    Boxed.Add Boxed5
    Boxed.Add Boxed6
    Boxed.Add Boxed7
    Boxed.Add Boxed8
    
    Set Bracket = New Collection
    Bracket.Add Bracket1
    Bracket.Add Bracket2
    Bracket.Add Bracket3
    Bracket.Add Bracket4
    Bracket.Add Bracket5
    Bracket.Add Bracket6
    Bracket.Add Bracket7
    Bracket.Add Bracket8
    
    Set InstallAnchor = New Collection
    InstallAnchor.Add InstallAnchor1
    InstallAnchor.Add InstallAnchor2
    InstallAnchor.Add InstallAnchor3
    InstallAnchor.Add InstallAnchor4
    InstallAnchor.Add InstallAnchor5
    InstallAnchor.Add InstallAnchor6
    InstallAnchor.Add InstallAnchor7
    InstallAnchor.Add InstallAnchor8
    
    Set reasonForMovement = New Collection
    reasonForMovement.Add RFM1
    reasonForMovement.Add RFM2
    reasonForMovement.Add RFM3
    reasonForMovement.Add RFM4
    reasonForMovement.Add RFM5
    reasonForMovement.Add RFM6
    reasonForMovement.Add RFM7
    reasonForMovement.Add RFM8
    
    Dim rfmArray(13) As String
    rfmArray(0) = "OTHER"
    rfmArray(1) = "correct multiple separation violations. "
    rfmArray(2) = "correct 15'6"" midspan ground clearance violation. "
    rfmArray(3) = "correct 12"" comm at pole separation violation. "
    rfmArray(4) = "correct 6"" comm midspan separation violation. "
    rfmArray(5) = "allow comms below to correct 15'6"" midspan ground clearance violation. "
    rfmArray(6) = "correct 40"" safety zone violation. "
    rfmArray(7) = "allow comms above to correct 40"" safety zone violation. "
    rfmArray(8) = "correct 30"" midspan separation violation. "
    rfmArray(9) = "allow comms above to correct 30"" midspan separation violation. "
    rfmArray(10) = "correct 40"" bottom streetlight bracket separation violation. "
    rfmArray(11) = "allow comms above to correct 40"" bottom streetlight bracket separation violation. "
    rfmArray(12) = "correct 12"" streetlight driploop separation violation. "
    rfmArray(13) = "allow comms above to correct 12"" streetlight driploop separation violation. "
    
    For i = 1 To 8
        reasonForMovement(i).list = rfmArray
        reasonForMovement(i).ListIndex = 0
    Next i
    
    Set reasonForAnchorMovement = New Collection
    reasonForAnchorMovement.Add RFAM1
    reasonForAnchorMovement.Add RFAM2
    reasonForAnchorMovement.Add RFAM3
    reasonForAnchorMovement.Add RFAM4
    reasonForAnchorMovement.Add RFAM5
    reasonForAnchorMovement.Add RFAM6
    reasonForAnchorMovement.Add RFAM7
    reasonForAnchorMovement.Add RFAM8
    
    Dim rfamArray(2) As String
    rfamArray(0) = "OTHER"
    rfamArray(1) = "make room for additional anchor. "
    rfamArray(2) = "correct pole loading failure. "
    
    For i = 1 To 8
        reasonForAnchorMovement(i).list = rfamArray
        reasonForAnchorMovement(i).ListIndex = 0
    Next i
    
    Set sheet = Application.ActiveSheet()
    Set comms = New Collection
    
    Call InitComms
End Sub

Private Sub InitComms()
    On Error Resume Next

    Dim pole As pole: Set pole = New pole
    Call pole.extractFromSheet(sheet)

    If Not sheet.Range("CEPOLE") Is Nothing Then
        If sheet.Range("CEPOLE").Value = True Then
            TextBox1.Value = "Consumers Energy"
        Else
            If Not sheet.Range("OTHERPOLEOWNER") Is Nothing Then
                TextBox1.Value = sheet.Range("OTHERPOLEOWNER").Value
            Else
                TextBox1.Value = "Unknown"
            End If
        End If
    End If
    
    For Each Comm In pole.commWires
        If Comm.modification > pole.applicant.modification Then CheckBox2.Value = False
    Next Comm
    
    If pole.applicant.height > 0 And Not pole.overlash Then CheckBox1.Value = True
    
    Dim commCount As Integer: commCount = 0
    For i = 1 To 8
        If commCount > 8 Then Exit For
        If Not sheet.Range("COMM" & i) Is Nothing Then
            If sheet.Range("COMM" & i).Value <> "COMM #" & i Then
                For j = 0 To 7
                    If commCount > 8 Then Exit For
                    modificationString = sheet.Range("COMM" & i).offset(2 + (j * 2), 0).offset(0, 1)
                    If Len(modificationString) > 0 Then
                        If Not IsNumeric(Left(modificationString, 1)) Then
                            modificationString = ""
                        Else
                            modificationString = Utilities.inchesToFeetInches(Utilities.convertToInches(modificationString))
                        End If
                    Else
                        modificationString = ""
                    End If
                    
                    If j = 0 Or modificationString <> "" Then
                        commCount = commCount + 1
                        modifications(commCount).Value = modificationString
                        companies(commCount).Value = sheet.Range("COMM" & i).Value
                        
                        Dim anchorDistance As String: anchorDistance = ""
                        For Each Anchor In pole.anchors
                            test1 = UCase(Anchor.owner)
                            test2 = UCase(companies(commCount).Value)
                            test3 = test1 = test2
                            If Trim(UCase(Anchor.owner)) = Trim(UCase(companies(commCount).Value)) Then
                                If anchorDistance <> "" Then
                                    anchorDistance = ""
                                    Exit For
                                End If
                                anchorDistance = Anchor.distance & "'"
                            End If
                        Next Anchor
                        If anchorDistance <> "" Then dgaprevious(i).text = anchorDistance
                        
                        heights(commCount).Value = Utilities.inchesToFeetInches(Utilities.convertToInches(sheet.Range("COMM" & i).offset(2 + (j * 2), 0)))
                        
                        If modifications(commCount).Value = "" Then modifications(commCount).Value = heights(commCount).Value
                    End If
                Next j
            End If
        End If
    Next i
End Sub

Private Function reset()
    Dim temp As Variant
    For Each temp In companies
        temp.Value = ""
    Next temp
    
    For Each temp In heights
        temp.Value = ""
    Next temp
    
    For Each temp In modifications
        temp.Value = ""
    Next temp
    
    For Each temp In MoveAnchor
        temp.Value = False
    Next temp
    
    For Each temp In rpguy
        temp.Value = False
    Next temp
    
    For Each temp In drops
        temp.Value = False
    Next temp
    
    For Each temp In dgaprevious
        temp.Value = ""
    Next temp
    
    For Each temp In dgaafter
        temp.Value = ""
    Next temp
    
    For Each temp In mainlines
        temp.Value = True
    Next temp
    
    For Each temp In Ignored
        temp.Value = False
    Next temp
    
    For Each temp In NotAttached
        temp.Value = False
    Next temp
        
    For Each temp In Boxed
        temp.Value = False
    Next temp
    
    For Each temp In Bracket
        temp.Value = False
    Next temp
    
    For Each temp In InstallAnchor
        temp.Value = False
    Next temp
    
    TextBox1.Value = ""
    
End Function

Private Sub CommandButton1_Click()
    
    On Error Resume Next
    
    Call FillComms

    Dim sortedcomms As Collection
    Set sortedcomms = sortComms(comms)
    Set comms = New Collection
    Dim NJUNSType As String
    NJUNSType = ticketType
    Dim applicant As Boolean
    applicant = CheckBox1.Value
    Dim ApplyAbove As Boolean
    If applicant Then
        ApplyAbove = CheckBox2.Value
    Else
        ApplyAbove = False
    End If
    
    Dim commheight As Integer
    Dim commmodification As Integer
    Dim movements As Collection
    Set movements = New Collection
    Dim Comm As Comm
    For Each Comm In sortedcomms
        commheight = Utilities.convertToInches(Comm.height)
        commmodification = Utilities.convertToInches(Comm.modification)
        If Comm.NotAttached Then
            movements.Add "Attach"
        ElseIf commmodification = -1 Or commmodification = commheight Then
            movements.Add "Nothing"
        ElseIf commheight > commmodification Then
            movements.Add "Lower"
        Else
            movements.Add "Raise"
        End If
    Next Comm
          
    NJUNSString = moveComms(1, sortedcomms, movements, applicant, ApplyAbove)
    
    NJUNSStringCondensed = ""
    previousCompany = ""
    If InStr(NJUNSString, vbCrLf & vbCrLf) > 0 Then
        steps = Split(NJUNSString, vbCrLf & vbCrLf)
        For i = 0 To UBound(steps) - 1
            lines = Split(steps(i), vbCrLf)
            company = lines(0)
            If previousCompany = company Then
                If NJUNSStringCondensed <> "" Then NJUNSStringCondensed = NJUNSStringCondensed & " "
                NJUNSStringCondensed = NJUNSStringCondensed & lines(1)
            Else
                If NJUNSStringCondensed <> "" Then NJUNSStringCondensed = NJUNSStringCondensed & vbCrLf & vbCrLf
                NJUNSStringCondensed = NJUNSStringCondensed & steps(i)
            End If
            previousCompany = company
        Next i
    End If
    
    Do While Right(NJUNSStringCondensed, 1) = Chr(10) Or Right(NJUNSStringCondensed, 1) = Chr(13)
        NJUNSStringCondensed = Left(NJUNSStringCondensed, Len(NJUNSStringCondensed) - 1)
    Loop
    
    If CE.Value And NJUNSStringCondensed <> "" Then
        NJUNSStringCondensed = "Consumers to complete required work." & vbCrLf & vbCrLf & NJUNSStringCondensed
    End If
    
    Dim DataObj As DataObject
    Set DataObj = New DataObject
    DataObj.SetText NJUNSStringCondensed
    DataObj.PutInClipboard
    
    If Not sheet.Range("NJUNS") Is Nothing Then
        If Trim(sheet.Range("NJUNS").Value) = "" Then sheet.Range("NJUNS").Value = NJUNSStringCondensed
    End If
    
    MsgBox NJUNSStringCondensed
End Sub

Private Sub CommandButton2_Click()
    
    On Error Resume Next
    
    Call FillComms
    Dim sortedcomms As Collection
    Set sortedcomms = sortComms(comms)
    Set comms = New Collection
    
    Dim NJUNSType As String
    NJUNSType = ticketType
    Dim applicant As Boolean
    applicant = CheckBox1.Value
    Dim ApplyAbove As Boolean
    If applicant Then
        ApplyAbove = CheckBox2.Value
    Else
        ApplyAbove = False
    End If
    
    NJUNSString = topPole(sortedcomms, applicant, ApplyAbove)
    
    NJUNSStringCondensed = ""
    previousCompany = ""
    If InStr(NJUNSString, vbCrLf & vbCrLf) > 0 Then
        steps = Split(NJUNSString, vbCrLf & vbCrLf)
        For i = 0 To UBound(steps)
            lines = Split(steps(i), vbCrLf)
            company = lines(0)
            If previousCompany = company Then
                If NJUNSStringCondensed <> "" Then NJUNSStringCondensed = NJUNSStringCondensed & " "
                NJUNSStringCondensed = NJUNSStringCondensed & lines(1)
            Else
                If NJUNSStringCondensed <> "" Then NJUNSStringCondensed = NJUNSStringCondensed & vbCrLf & vbCrLf
                NJUNSStringCondensed = NJUNSStringCondensed & steps(i)
            End If
            previousCompany = company
        Next i
    End If
    
    Do While Right(NJUNSStringCondensed, 1) = Chr(10) Or Right(NJUNSStringCondensed, 1) = Chr(13)
        NJUNSStringCondensed = Left(NJUNSStringCondensed, Len(NJUNSStringCondensed) - 1)
    Loop
    Dim DataObj As DataObject
    Set DataObj = New DataObject
    DataObj.SetText NJUNSStringCondensed
    DataObj.PutInClipboard
    
    If Not sheet.Range("NJUNS") Is Nothing Then
        If Trim(sheet.Range("NJUNS").Value) = "" Then sheet.Range("NJUNS").Value = NJUNSStringCondensed
    End If
    
    MsgBox NJUNSStringCondensed
    
End Sub

Private Sub CommandButton3_click()

    On Error Resume Next

    NJUNSString = TextBox1.Value & vbCrLf
    NJUNSString = NJUNSString & "Safety zone violation was identified at pole location, verify NESC clearance requirements." & vbCrLf
    NJUNSString = NJUNSString & "Low clearance was identified at pole location in [DIRECTION] span, verify NESC clearance requirements." & vbCrLf
    NJUNSString = NJUNSString & "Pole is tagged for replacement." & vbCrLf
    NJUNSString = NJUNSString & "[DELETE THIS LINE AND ALL LINES THAT DON'T APPLY TO THIS POLE]" & vbCrLf

    Dim DataObj As DataObject
    Set DataObj = New DataObject
    DataObj.SetText NJUNSString
    DataObj.PutInClipboard
    
    If Not sheet.Range("NJUNS") Is Nothing Then
        If Trim(sheet.Range("NJUNS").Value) = "" Then sheet.Range("NJUNS").Value = NJUNSString
    End If
    
    MsgBox NJUNSString

End Sub

Private Sub FillComms()
    Dim i As Integer
    For i = 1 To 8
        Dim Comm As Comm
        Set Comm = New Comm
        If companies(i).Value <> "" And heights(i).Value <> "" Then
            If companyCount(companies, i) > 1 Then Comm.orientation = findOrientation(heights, companies, i)
            Comm.owner = companies(i).Value
            Comm.height = heights(i).Value
            Comm.modification = Trim(modifications(i).Value)
            Comm.reasonForMovement = Replace(reasonForMovement(i).Value, "OTHER", "[REASON FOR WORK]. ")
            Comm.reasonForAnchorMovement = Replace(reasonForAnchorMovement(i).Value, "OTHER", "[REASON FOR WORK]. ")
            If Comm.modification = "" Then Comm.modification = Comm.height
            Comm.drops = drops(i).Value
            Comm.MoveAnchor = MoveAnchor(i).Value
            Comm.InstallAnchor = InstallAnchor(i).Value
            If Comm.MoveAnchor Then
                If OnlyNumbers(dgaprevious(i).Value) = "" Then dgaprevious(i).Value = 0
                If OnlyNumbers(dgaafter(i).Value) = "" Then dgaafter(i).Value = 0
                If InStr(dgaprevious(i).Value, "'") = 0 Then dgaprevious(i).Value = Utilities.OnlyNumbers(dgaprevious(i).Value) & "'"
                If InStr(dgaafter(i).Value, "'") = 0 Then dgaafter(i).Value = Utilities.OnlyNumbers(dgaafter(i).Value) & "'"
                Comm.PreviousAnchor = Utilities.OnlyNumbers(dgaprevious(i).Value)
                Comm.NextAnchor = Utilities.OnlyNumbers(dgaafter(i).Value)
            ElseIf Comm.InstallAnchor Then
                Comm.NextAnchor = Utilities.OnlyNumbers(dgaafter(i).Value)
            End If
            Comm.UpgradeGuy = rpguy(i).Value
            Comm.Mainline = mainlines(i).Value
            Comm.Ignored = Ignored(i).Value
            Comm.NotAttached = NotAttached(i).Value
            Comm.Boxed = Boxed(i).Value
            Comm.Bracket = Bracket(i).Value
            comms.Add Comm
        End If
    Next i
End Sub

Public Function companyCount(companies As Collection, i As Integer) As Integer
    count = 0
    For j = 1 To companies.count
        If companies(j).Value = companies(i).Value Then count = count + 1
    Next j
    companyCount = count
End Function

Public Function findOrientation(heights As Collection, companies As Collection, i As Integer) As String
    orientation = ""
    highestCompany = True
    lowestCompany = True
    duplicateHeight = False
    For j = 1 To 8
        If i <> j And companies(i) = companies(j) Then
            If Utilities.convertToInches(heights(j).Value) > Utilities.convertToInches(heights(i).Value) Then
                highestCompany = False
            ElseIf Utilities.convertToInches(heights(j).Value) < Utilities.convertToInches(heights(i).Value) Then
                lowestCompany = False
            ElseIf Utilities.convertToInches(heights(j).Value) = Utilities.convertToInches(heights(i).Value) Then
                duplicateHeight = True
            End If
        End If
    Next j
    
    If highestCompany Then
        orientation = "top"
    ElseIf lowestCompany Then
        orientation = "bottom"
    Else
        orientation = "middle"
    End If
    
    If duplicateHeight Then orientation = "one of the " & orientation
    
    findOrientation = orientation
End Function

Private Sub CheckBox1_change()
    If CheckBox1 Then
        CheckBox2.visible = True
    Else
        CheckBox2.visible = False
    End If
End Sub

Private Sub CommMoveAnchor1_change()
    DGAPrevious1.visible = CommMoveAnchor1
    DGAAfter1.visible = CommMoveAnchor1 Or InstallAnchor1
    CommDGAPrevious1.visible = CommMoveAnchor1
    CommDGAAfter1.visible = CommMoveAnchor1 Or InstallAnchor1
    RFAM1.visible = CommMoveAnchor1
    RFAML1.visible = CommMoveAnchor1
End Sub

Private Sub CommMoveAnchor2_change()
    DGAPrevious2.visible = CommMoveAnchor2
    DGAAfter2.visible = CommMoveAnchor2 Or InstallAnchor2
    CommDGAPrevious2.visible = CommMoveAnchor2
    CommDGAAfter2.visible = CommMoveAnchor2 Or InstallAnchor2
    RFAM2.visible = CommMoveAnchor2
    RFAML2.visible = CommMoveAnchor2
End Sub

Private Sub CommMoveAnchor3_change()
    DGAPrevious3.visible = CommMoveAnchor3
    DGAAfter3.visible = CommMoveAnchor3 Or InstallAnchor3
    CommDGAPrevious3.visible = CommMoveAnchor3
    CommDGAAfter3.visible = CommMoveAnchor3 Or InstallAnchor3
    RFAM3.visible = CommMoveAnchor3
    RFAML3.visible = CommMoveAnchor3
End Sub

Private Sub CommMoveAnchor4_change()
    DGAPrevious4.visible = CommMoveAnchor4
    DGAAfter4.visible = CommMoveAnchor4 Or InstallAnchor4
    CommDGAPrevious4.visible = CommMoveAnchor4
    CommDGAAfter4.visible = CommMoveAnchor4 Or InstallAnchor4
    RFAM4.visible = CommMoveAnchor4
    RFAML4.visible = CommMoveAnchor4
End Sub

Private Sub CommMoveAnchor5_change()
    DGAPrevious5.visible = CommMoveAnchor5
    DGAAfter5.visible = CommMoveAnchor5 Or InstallAnchor5
    CommDGAPrevious5.visible = CommMoveAnchor5
    CommDGAAfter5.visible = CommMoveAnchor5 Or InstallAnchor5
    RFAM5.visible = CommMoveAnchor5
    RFAML5.visible = CommMoveAnchor5
End Sub

Private Sub CommMoveAnchor6_change()
    DGAPrevious6.visible = CommMoveAnchor6
    DGAAfter6.visible = CommMoveAnchor6 Or InstallAnchor6
    CommDGAPrevious6.visible = CommMoveAnchor6
    CommDGAAfter6.visible = CommMoveAnchor6 Or InstallAnchor6
    RFAM6.visible = CommMoveAnchor6
    RFAML6.visible = CommMoveAnchor6
End Sub

Private Sub CommMoveAnchor7_change()
    DGAPrevious7.visible = CommMoveAnchor7
    DGAAfter7.visible = CommMoveAnchor7 Or InstallAnchor7
    CommDGAPrevious7.visible = CommMoveAnchor7
    CommDGAAfter7.visible = CommMoveAnchor7 Or InstallAnchor7
    RFAM7.visible = CommMoveAnchor7
    RFAML7.visible = CommMoveAnchor7
End Sub

Private Sub CommMoveAnchor8_change()
    DGAPrevious8.visible = CommMoveAnchor8
    DGAAfter8.visible = CommMoveAnchor8 Or InstallAnchor8
    CommDGAPrevious8.visible = CommMoveAnchor8
    CommDGAAfter8.visible = CommMoveAnchor8 Or InstallAnchor8
    RFAM8.visible = CommMoveAnchor8
    RFAML8.visible = CommMoveAnchor8
End Sub

Private Sub InstallAnchor1_Click()
    DGAAfter1.visible = CommMoveAnchor1 Or InstallAnchor1
    CommDGAAfter1.visible = CommMoveAnchor1 Or InstallAnchor1
End Sub

Private Sub InstallAnchor2_Click()
    DGAAfter2.visible = CommMoveAnchor2 Or InstallAnchor2
    CommDGAAfter2.visible = CommMoveAnchor2 Or InstallAnchor2
End Sub

Private Sub InstallAnchor3_Click()
    DGAAfter3.visible = CommMoveAnchor3 Or InstallAnchor3
    CommDGAAfter3.visible = CommMoveAnchor3 Or InstallAnchor3
End Sub

Private Sub InstallAnchor4_Click()
    DGAAfter4.visible = CommMoveAnchor4 Or InstallAnchor4
    CommDGAAfter4.visible = CommMoveAnchor4 Or InstallAnchor4
End Sub

Private Sub InstallAnchor5_Click()
    DGAAfter5.visible = CommMoveAnchor5 Or InstallAnchor5
    CommDGAAfter5.visible = CommMoveAnchor5 Or InstallAnchor5
End Sub

Private Sub InstallAnchor6_Click()
    DGAAfter6.visible = CommMoveAnchor6 Or InstallAnchor6
    CommDGAAfter6.visible = CommMoveAnchor6 Or InstallAnchor6
End Sub

Private Sub InstallAnchor7_Click()
    DGAAfter7.visible = CommMoveAnchor7 Or InstallAnchor7
    CommDGAAfter7.visible = CommMoveAnchor7 Or InstallAnchor7
End Sub

Private Sub InstallAnchor8_Click()
    DGAAfter8.visible = CommMoveAnchor8 Or InstallAnchor8
    CommDGAAfter8.visible = CommMoveAnchor8 Or InstallAnchor8
End Sub
