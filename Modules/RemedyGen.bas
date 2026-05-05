Attribute VB_Name = "RemedyGen"
Sub RemedyGenerator()
    On Error Resume Next
    
    Call LogMessage.SendLogMessage("RemedyGenerator")
    
    Dim modificationCells As Collection
    Dim otherModificationCells As scripting.Dictionary
    
    Dim wires As Collection: Set wires = New Collection
    Dim sheet As Worksheet: Set sheet = ThisWorkbook.ActiveSheet
    Dim spans As Integer: spans = 0
    Dim weps As scripting.Dictionary: Set weps = New scripting.Dictionary
    Dim i, j As Integer
    Dim midspans As scripting.Dictionary
    Dim overlash As Boolean
    
    If sheet.name = "4 Spans" Or sheet.name = "8 Spans" Or sheet.name = "12 Spans" Or sheet.Cells(2, 2).Value <> "Notification:" Then
        MsgBox "You need to have a pole detail sheet active to run this script."
        Exit Sub
    End If
    
    Dim comms As Collection: Set comms = New Collection
    comms.Add "COMM1"
    comms.Add "COMM2"
    comms.Add "COMM3"
    comms.Add "COMM4"
    comms.Add "COMM5"
    comms.Add "COMM6"
    comms.Add "COMM7"
    comms.Add "COMM8"
    
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Pattern = "^(.*?)(?=\()"
        .IgnoreCase = True
        .Global = False
    End With
    
    Dim found As Boolean
    For i = 1 To 12
        found = False
        For Each name In sheet.names
            If name.name = "'" & sheet.name & "'" & "!" & "TOPOLE" & i Then
                found = True
                spans = i
                Exit For
            End If
        Next name
        If Not found Then Exit For
    Next i
    
    Dim applicant As Wire: Set applicant = Nothing
    Set adjacentHeights = New scripting.Dictionary
    Set midspans = New scripting.Dictionary
    Set applicant = New Wire
    applicant.owner = Trim(UCase(sheet.Range("APPLICANT")))
    applicant.height = Utilities.convertToInches(sheet.Range("PROPOSEDHEIGHT"))
    If Utilities.convertToInches(sheet.Range("CMRF1")) > 0 Then
        applicant.modification = Utilities.convertToInches(sheet.Range("CMRF1"))
    Else
        applicant.modification = applicant.height
    End If
    For i = 1 To spans
        If Trim(Replace(sheet.Range("TOPOLE" & i).offset(1, 0).text, "-", "")) <> "" Then
            Dim midspan As Integer: midspan = Utilities.convertToInches(sheet.Range("TOPOLE" & i).offset(1, 0).text)
            midspans.Add i, midspan
            If Not weps.Exists(i) Then weps.Add i, sheet.Range("TOPOLE" & i).text
            If regex.test(weps(i)) Then
                Set matches = regex.Execute(weps(i))
                Dim otherSheetName As String: otherSheetName = matches(0)
                If SheetExists(otherSheetName) Then
                    Dim otherSheet As Worksheet: Set otherSheet = Utilities.GetPDS(otherSheetName)
                    Dim adjacentHeight As Integer: adjacentHeight = Utilities.convertToInches(otherSheet.Range("PROPOSEDHEIGHT"))
                    Dim adjacentModification As Integer: adjacentModification = Utilities.convertToInches(otherSheet.Range("CMRF1"))
                    If InStr(adjacentHeight, "OL") = 0 And adjacentHeight <> adjacentModification And Utilities.convertToInches(adjacentHeight) > 0 And Utilities.convertToInches(adjacentModification) > 0 Then
                        adjacentHeights.Add i, New Collection
                        adjacentHeights(i).Add Utilities.convertToInches(adjacentHeight)
                        adjacentHeights(i).Add Utilities.convertToInches(adjacentModification)
                    ElseIf InStr(adjacentHeight, "OL") > 0 And Utilities.convertToInches(adjacentModification) > 0 Then
                        adjacentHeights.Add i, New Collection
                        adjacentHeights(i).Add Utilities.convertToInches(adjacentHeight)
                        adjacentHeights(i).AddUtilities.convertToInches (adjacentModification)
                    End If
                End If
            End If
        End If
    Next i
    
    Set applicant.midspans = midspans
    Set applicant.adjacentHeights = adjacentHeights
    
    overlash = False
    Set modificationCells = New Collection
    Set otherModificationCells = New scripting.Dictionary
    For i = 0 To 50
        If sheet.Range("CMOWNER").offset(i, 0).Interior.color = 16777215 Then Exit For
        If sheet.Range("CMOWNER").offset(i, 0) = "" Then Exit For
        If sheet.Range("CMOWNER").offset(i, 0).text = "Clearance Requirment" Then Exit For
        addComm = True
        Dim owner As String: owner = Trim(UCase(Replace(Replace(sheet.Range("CMOWNER").offset(i, 0).text, " MSG", ""), " SVC", "")))
        Dim height As Integer: height = Utilities.convertToInches(sheet.Range("CMHEIGHT").offset(i, 0).text)
        
        If InStr(sheet.Range("CMOWNER").offset(i, 0).text, "SVC") > 0 Then
            For j = 0 To 50
                If sheet.Range("CMOWNER").offset(j, 0).Interior.color = 16777215 Then Exit For
                If sheet.Range("CMOWNER").offset(j, 0) = "" Then Exit For
                If sheet.Range("CMOWNER").offset(j, 0).text = "Clearance Requirment" Then Exit For
                If Replace(Trim(sheet.Range("CMOWNER").offset(j, 0).text), " MSG", "") = owner Then
                    addComm = False
                    Exit For
                End If
            Next j
        End If
        If addComm Then
            Set midspans = New scripting.Dictionary
            Set adjacentHeights = New scripting.Dictionary
            For j = 1 To spans
                midspan = Utilities.convertToInches(sheet.Range("CMMIDSPAN" & j).offset(i, 0).text)
                If midspan > 0 Then
                    midspans.Add j, midspan
                    If Not weps.Exists(j) Then weps.Add j, sheet.Range("TOPOLE" & j).Value
                    If regex.test(weps(j)) Then
                        Set matches = regex.Execute(weps(j))
                        otherSheetName = matches(0)
                        If SheetExists(otherSheetName) Then
                            Set otherSheet = Utilities.GetPDS(otherSheetName)
                            If Not otherModificationCells.Exists(otherSheet.name) Then otherModificationCells.Add otherSheet.name, New Collection
                            adjacentHeight = getHeight(otherSheet, owner, midspan, sheet.Range("POLENUM").text)
                            Set modCell = findModCell(otherSheet, adjacentHeight, owner, False, otherModificationCells(otherSheet.name))
                            adjacentModification = Utilities.convertToInches(modCell.Value)
                            If adjacentModification < 1 Then adjacentModification = adjacentHeight
                            
                            If adjacentModification > 0 Then
                                adjacentHeights.Add j, New Collection
                                adjacentHeights(j).Add adjacentHeight
                                adjacentHeights(j).Add adjacentModification
                            End If
                        End If
                    End If
                End If
            Next j
            
            If owner = applicant.owner Then
                If overlash Then MsgBox "Warning, Remedy generator won't work correctly on an overlash job where the applicant is attached at 2 different heights"
                overlash = True
                applicant.height = height
                If Utilities.convertToInches(sheet.Range("CMRF1")) > 0 Then
                    applicant.modification = Utilities.convertToInches(sheet.Range("CMRF1"))
                Else
                    applicant.modification = Utilities.convertToInches(findModCell(sheet, height, owner).text)
                    If applicant.modification < 1 Then applicant.modification = height
                End If
                
                For Each key In midspans.keys
                    If adjacentHeights.Exists(key) Then
                        Set applicant.adjacentHeights.item(key) = adjacentHeights.item(key)
                        adjacentHeights.Remove key
                    End If
                    If applicant.midspans.Exists(key) Then
                        If applicant.midspans(key) > midspans.item(key) Then applicant.midspans.item(key) = midspans.item(key)
                    Else
                        applicant.midspans.Add key, midspans.item(key)
                        If Not adjacentHeights.Exists(key) Then
                            If regex.test(weps(key)) Then
                                Set matches = regex.Execute(weps(key))
                                otherSheetName = matches(0)
                                If SheetExists(otherSheetName) Then
                                    Set otherSheet = Utilities.GetPDS(otherSheetName)
                                    midspan = key
                                    adjacentHeight = getHeight(otherSheet, owner, midspan, sheet.Range("POLENUM").text)
                                    adjacentModification = otherSheet.Range("CMRF1")
                                    If Len(adjacentModification) > 0 Then
                                        If Not IsNumeric(Left(adjacentModification, 1)) Then adjacentModification = adjacentHeight
                                    Else
                                        adjacentModification = adjacentHeight
                                    End If
                                    If adjacentHeight <> adjacentModification And Utilities.convertToInches(adjacentModification) > 0 Then
                                        applicant.adjacentHeights.Add key, New Collection
                                        applicant.adjacentHeights(key).Add Utilities.convertToInches(adjacentHeight)
                                        applicant.adjacentHeights(key).Add Utilities.convertToInches(adjacentModification)
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    midspans.Remove key
                Next key
            End If
            
            
            If midspans.count > 0 Then
                Dim Wire As Wire: Set Wire = New Wire
                
                overlash = False
                If Wire.owner = applicant.owner Then overlash = True
                
                Wire.owner = owner
                Wire.height = height
                Dim modification As String
                Set modCell = findModCell(sheet, Wire.height, Wire.owner, False, modificationCells)
                Wire.modification = Utilities.convertToInches(modCell.Value)
                If Wire.modification < 1 Then Wire.modification = Wire.height
                
                Duplicate = False
                For Each otherWire In wires
                    If Wire.owner = otherWire.owner And Wire.height = otherWire.height Then
                        Duplicate = True
                        For Each key In otherWire.midspans
                            If midspans.Exists(key) Then Duplicate = False
                        Next key
                        If Duplicate Then
                            For Each key In midspans
                                otherWire.midspans.Add key, midspans(key)
                            Next key
                            For Each key In adjacentHeights
                                otherWire.adjacentHeights.Add key, adjacentHeights(key)
                            Next key
                            Exit For
                        End If
                    End If
                Next otherWire
                
                If Not Duplicate Then
                    Set Wire.midspans = midspans
                    Set Wire.adjacentHeights = adjacentHeights
                    wires.Add Wire
                End If
            End If
        End If
    Next i
      
    Dim powers As Collection: Set powers = New Collection
    Dim OGPowers As Collection: Set OGPowers = New Collection
    
    If sheet.Range("LWSTPWR") = "" Then
        powers.Add "N/A"
        OGPowers.Add "N/A"
    Else
        parts = Split(sheet.Range("LWSTPWR"), vbLf)
        OGPowers.Add Utilities.convertToInches(CStr(parts(0)))
        If UBound(parts) > 0 Then
            If Utilities.convertToInches(CStr(parts(1))) > 0 Then
                powers.Add Utilities.convertToInches(parts(1))
            Else
                powers.Add Utilities.convertToInches(parts(0))
            End If
        Else
            powers.Add Utilities.convertToInches(parts(0))
        End If
    End If
    
    If sheet.Range("STLTBRKT") = "" Then
        powers.Add "N/A"
        OGPowers.Add "N/A"
    Else
        parts = Split(sheet.Range("STLTBRKT"), vbLf)
        OGPowers.Add Utilities.convertToInches(parts(0))
        If UBound(parts) > 0 Then
            If Utilities.convertToInches(parts(1)) > 0 Then
                powers.Add Utilities.convertToInches(parts(1))
            Else
                powers.Add Utilities.convertToInches(parts(0))
            End If
        Else
            powers.Add Utilities.convertToInches(parts(0))
        End If
    End If
    
    If sheet.Range("STLTDL") = "" Then
        powers.Add "N/A"
        OGPowers.Add "N/A"
    Else
        parts = Split(sheet.Range("STLTDL"), vbLf)
        OGPowers.Add Utilities.convertToInches(parts(0))
        If UBound(parts) > 0 Then
            If Utilities.convertToInches(parts(1)) > 0 Then
                powers.Add Utilities.convertToInches(parts(1))
                Else
                powers.Add Utilities.convertToInches(parts(0))
            End If
        Else
            powers.Add Utilities.convertToInches(parts(0))
        End If
    End If
    
    Dim clearanceMidspans As scripting.Dictionary: Set clearanceMidspans = New scripting.Dictionary
    Dim OGClearanceMidspans As scripting.Dictionary: Set OGClearanceMidspans = New scripting.Dictionary
    
    For Each wep In weps
        If sheet.Range("CMMIDSPAN" & wep).offset(-2, 0).text <> "" Then
            parts = Split(sheet.Range("CMMIDSPAN" & wep).offset(-2, 0).text, vbLf)
            OGClearanceMidspans.Add wep, Utilities.convertToInches(parts(0)) + 30
            If UBound(parts) > 0 Then
                If Utilities.convertToInches(parts(1)) > 0 Then
                    clearanceMidspans.Add wep, Utilities.convertToInches(parts(1)) + 30
                Else
                    clearanceMidspans.Add wep, Utilities.convertToInches(parts(0)) + 30
                End If
            Else
                clearanceMidspans.Add wep, Utilities.convertToInches(parts(0)) + 30
            End If
        Else
            OGClearanceMidspans.Add wep, -1
            clearanceMidspans.Add wep, -1
        End If
    Next wep
    
    Dim IgnoreBolt As Boolean
    If InStr(sheet.Range("NJUNSTICKET"), "PT") > 0 Or sheet.Range("REPLACEPOLE").Value = True Then IgnoreBolt = True
    
    Dim Bonded As Boolean
    If sheet.Range("BONDED") = "Yes" Then
        Bonded = True
    Else
        If InStr(sheet.Range("ALTONE"), "BOND") > 0 Then Bonded = True
    End If
    
    Unload RemedyGenerator_Form
    Call RemedyGenerator_Form.Initialize(sheet, powers, OGPowers, clearanceMidspans, OGClearanceMidspans, weps, applicant, wires, IgnoreBolt, Bonded, overlash)
    RemedyGenerator_Form.Show vbModeless
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Private Function getMod(sheet As Worksheet, height As Integer, owner As String, comms As Collection) As Integer
    If height > 0 Then
        For Each commCellName In comms
            If InStr(sheet.Range(commCellName).text, "COMM #") > 0 Then Exit For
            If InStr(1, sheet.Range(commCellName).text, owner, vbTextCompare) > 0 Then
                Set startingPoint = sheet.Range(commCellName).offset(1, 0)
                For i = 1 To 100 Step 2
                    If startingPoint.offset(i, 0).text = "" Then Exit For
                    If Utilities.convertToInches(startingPoint.offset(i, 0).text) = height And Utilities.convertToInches(startingPoint.offset(i, 1).text) > 0 Then
                        getMod = Utilities.convertToInches(startingPoint.offset(i, 0).offset(0, 1).text)
                        Exit Function
                    End If
                Next i
            End If
        Next commCellName
    End If
    
    getMod = height
End Function

Private Function getHeight(sheet As Worksheet, owner As String, midspan As Integer, poleNumber As String) As Integer
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Pattern = "^(.*?)(?=\()"
        .IgnoreCase = True
        .Global = False
    End With
    
    Dim Span As Integer: Span = 0
    For i = 1 To 12
        If Utilities.RangeExists(sheet, "TOPOLE" & i) Then
            If regex.test(sheet.Range("TOPOLE" & i)) Then
                Set matches = regex.Execute(sheet.Range("TOPOLE" & i))
                If matches(0) = poleNumber Then
                    Span = i
                    Exit For
                End If
            End If
        End If
        If Span <> 0 Then Exit For
    Next i
    
    If Span <> 0 Then
        For i = 0 To 100
            If sheet.Range("CMMIDSPAN" & Span).offset(i + 1, 0).Interior.color <> 16312794 Then Exit For
            If Utilities.convertToInches(sheet.Range("CMMIDSPAN" & Span).offset(i, 0).text) = midspan And InStr(sheet.Range("CMOWNER").offset(i, 0).text, owner) > 0 Then
                getHeight = Utilities.convertToInches(sheet.Range("CMHEIGHT").offset(i, 0).text)
                Exit Function
            End If
        Next i
    End If
    
    getHeight = -1
End Function

Public Sub calculateProposedMidspans(Optional poleSheet As Worksheet)

    On Error Resume Next
    
    Dim sheet As Worksheet, Sheet2 As Worksheet

    If poleSheet Is Nothing Then
        Call LogMessage.SendLogMessage("calculateProposedMidspans")
    
        Set sheet = ThisWorkbook.ActiveSheet()
        If sheet.name = "4 Spans" Or sheet.name = "8 Spans" Or sheet.name = "12 Spans" Or sheet.Cells(2, 2).Value <> "Notification:" Then
            MsgBox "You need to have a pole detail sheet active to run this script."
            Exit Sub
        End If
    Else
        Set sheet = poleSheet
    End If
    
    PROPOSEDHEIGHT = sheet.Range("PROPOSEDHEIGHT").Value
    finalHeight = sheet.Range("CMRF1").Value
    
    If Utilities.convertToInches(PROPOSEDHEIGHT) < 1 Then
        If poleSheet Is Nothing Then
            MsgBox "Please put a Proposed Height first"
        End If
        Exit Sub
    End If

    If Utilities.convertToInches(finalHeight) < 1 Then
        If poleSheet Is Nothing Then
            MsgBox "Please put a New Attacher Height first"
        End If
        Exit Sub
    End If
    
    difference = Utilities.convertToInches(finalHeight) - Utilities.convertToInches(PROPOSEDHEIGHT)
    midspanChange = difference / 2
    
    updatedMidspans = ""
    polesChanged = ""
    
    For i = 1 To 12
        midspanChange2 = 0
        updatedMidspan = ""
        midspan = ""
        direction = ""
        Set toPoleCell = Nothing
        If Utilities.RangeExists(sheet, "TOPOLE" & i) Then Set toPoleCell = sheet.Range("TOPOLE" & i)
        If toPoleCell Is Nothing Then Exit For
        midspan = toPoleCell.offset(1, 0).Value
        If Utilities.convertToInches(midspan) > 0 Then
            parenth = InStr(toPoleCell.Value, "(")
            If parenth > 0 Then
                pole = Left(toPoleCell.Value, parenth - 1)
                If Utilities.SheetExists(pole) Then
                    For Each oSheet In ThisWorkbook.sheets
                        If ThisWorkbook.RemoveParentheses(oSheet.name) = pole Then
                            Set Sheet2 = oSheet
                            Exit For
                        End If
                    Next oSheet
                    
                    direction = " @P" & pole
                    
                    If Utilities.convertToInches(Sheet2.Range("CMRF1").Value) > 0 Then
                        proposedHeight2 = Sheet2.Range("PROPOSEDHEIGHT").Value
                        finalHeight2 = Sheet2.Range("CMRF1").Value
                        difference2 = Utilities.convertToInches(finalHeight2) - Utilities.convertToInches(proposedHeight2)
                        midspanChange2 = difference2 / 2
                        
                        If InStr(Sheet2.Range("CMRF2").Value, " @P" & sheet.Range("POLENUM")) > 0 Then
                            If poleSheet Is Nothing Then
                                answer = MsgBox("Do you want to update the value of midspans on pole " & pole, vbYesNoCancel + vbQuestion, "Confirmation")
                            Else
                                answer = vbNo
                            End If
                            
                            If answer = vbYes Then
                                Dim lines() As String
                                lines = Split(Sheet2.Range("CMRF2").Value, vbLf)
                                For j = LBound(lines) To UBound(lines)
                                    If InStr(lines(j), " @P" & sheet.Range("POLENUM")) > 0 Then
                                        lines(j) = Utilities.inchesToFeetInches(Utilities.convertToInches(midspan) + Int(midspanChange + midspanChange2)) & " @P" & sheet.Range("POLENUM")
                                    End If
                                Next j
                                Sheet2.Range("CMRF2").Value = Join(lines, vbLf)
                                If polesChanged <> "" Then polesChanged = polesChanged & ", "
                                polesChanged = polesChanged & pole
                            End If
                            midspan = ""
                        End If
                    End If
                Else
                    direction = " " & pole
                End If
                If midspan <> "" Then
                    updatedMidspan = Utilities.inchesToFeetInches(Utilities.convertToInches(midspan) + Int(midspanChange + midspanChange2))
                    If updatedMidspans <> "" Then updatedMidspans = updatedMidspans & vbLf
                    updatedMidspans = updatedMidspans & updatedMidspan & direction
                End If
            End If

        End If
    Next i
    
    If updatedMidspans <> "" And polesChanged <> "" Then
        sheet.Range("CMRF2").Value = updatedMidspans
        If poleSheet Is Nothing Then
            MsgBox "Updated the midspans on this pole" & vbLf & "Poles: " & polesChanged & " have also had their midspans updated on a different page"
        End If
    ElseIf updatedMidspans <> "" Then
        sheet.Range("CMRF2").Value = updatedMidspans
        If poleSheet Is Nothing Then
            MsgBox "Updated the midspans on this pole"
        End If
    ElseIf polesChanged <> "" Then
        If poleSheet Is Nothing Then
            MsgBox "Poles: " & polesChanged & " have had their midspans updated on a different page"
        End If
    End If
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Public Function findModCell(sheet As Worksheet, height As Integer, owner As String, Optional findEmpty As Boolean, Optional modificationCells As Collection) As Range
    
    On Error Resume Next
    Dim commCellName As String
    Dim cell As Range
    
    If height > 0 Then
        For i = 1 To 8
            commCellName = "COMM" & i
            If InStr(sheet.Range(commCellName).text, "COMM #") > 0 Then Exit For
            If InStr(1, sheet.Range(commCellName).text, owner, vbTextCompare) > 0 Then
                Set startingPoint = sheet.Range(commCellName).offset(1, 0)
                For j = 1 To 100 Step 2
                    If startingPoint.offset(j, 0).text = "" Then Exit For
                    If Utilities.convertToInches(startingPoint.offset(j, 0)) = height Then
                        If (findEmpty And startingPoint.offset(j, 0).offset(0, 1) = "") Or Not findEmpty Then
                            Unique = True
                            If Not modificationCells Is Nothing Then
                                For Each cell In modificationCells
                                    If cell.address = startingPoint.offset(j, 0).offset(0, 1).address Then
                                        Unique = False
                                        Exit For
                                    End If
                                Next cell
                            End If
                            If Unique Then
                                Set findModCell = startingPoint.offset(j, 0).offset(0, 1)
                                If Not modificationCells Is Nothing Then modificationCells.Add startingPoint.offset(j, 0).offset(0, 1)
                                Exit Function
                            End If
                        End If
                    End If
                Next j
            End If
        Next i
    End If
    
    Set findModCell = Nothing
End Function

Public Sub clearModCells(sheet As Worksheet)
    For i = 1 To 8
        commCellName = "COMM" & i
        If InStr(sheet.Range(commCellName).text, "COMM #") > 0 Then Exit For
        Set startingPoint = sheet.Range(commCellName).offset(1, 0)
        For j = 1 To 100 Step 2
            If startingPoint.offset(j, 0).text = "" Then Exit For
            startingPoint.offset(j, 0).offset(0, 1) = ""
        Next j
    Next i
End Sub
