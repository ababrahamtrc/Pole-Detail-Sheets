Attribute VB_Name = "ExportStudio"
Sub ExportStudioJson()
    Dim json As Scripting.Dictionary: Set json = New Scripting.Dictionary
    Dim project As project: Set project = New project
    Call project.extractFromSheets
    
    json("Notification") = project.notification
    json("Permit") = project.Permit
    json("Applicant") = project.applicant
    json("Data Collection Date") = project.dataCollectionDate
    json("Inspected By") = project.fielder
    Set json("Poles") = New Scripting.Dictionary
    
    Dim pole As pole
    For Each pole In project.poles
        If Not json("Poles").exists(pole.poleNumber) Then
            Dim jsonPole As Scripting.Dictionary: Set jsonPole = New Scripting.Dictionary
            jsonPole("Pole Number") = pole.poleNumber
            jsonPole("Location") = pole.location
            jsonPole("Pole CEID") = pole.existingCEID
            jsonPole("Pole Old CEID") = pole.gisCEID
            jsonPole("Pole Height") = CStr(Application.WorksheetFunction.MRound(Abs(pole.height) / 12, 5))
            If jsonPole("Pole Height") = "0" Then jsonPole("Pole Height") = ""
            jsonPole("Pole Class") = pole.Class
            jsonPole("Pole Species") = pole.species
            jsonPole("Pole Owner") = pole.owner
            jsonPole("Framing Type") = getFramingType(pole)
            jsonPole("NJUNS") = IIf(InStr(pole.njunsTicket, "PT") > 0, pole.njunsTicket, "")
            jsonPole("New Attachment - Attach Height") = pole.sheet.Range("CMRF1").value
            jsonPole("New Attachment - Midspans") = pole.sheet.Range("CMRF2").value
            jsonPole("New Attachment - Downguys") = pole.sheet.Range("CMRF3").value
            jsonPole("Ok to Attach Before Work") = pole.sheet.Range("SUMSHEET1").value
            jsonPole("Work Required Before Attach") = pole.sheet.Range("SUMSHEET2").value
            jsonPole("CE to Correct Violations") = pole.sheet.Range("SUMSHEET3").value
            jsonPole("Comm to Correct Violations") = pole.sheet.Range("SUMSHEET4").value
            jsonPole("CE Maintenance Work Required") = pole.sheet.Range("SUMSHEET5").value
            jsonPole("CE Make Ready Work Required") = pole.sheet.Range("SUMSHEET6").value
            jsonPole("Com. Make Ready Work Required") = pole.sheet.Range("SUMSHEET7").value
            jsonPole("Attachment Denied") = pole.sheet.Range("SUMSHEET8").value
            jsonPole("Foriegn Pole") = pole.sheet.Range("SUMSHEET9").value
            jsonPole("PAAPMS") = pole.sheet.Range("SUMSHEET10").value
            jsonPole("Alt Designs") = Replace(pole.sheet.Range("SUMSHEET11").value, "-", "")
            jsonPole("HVD Structure #") = pole.hvd
            jsonPole("PPRR Required") = pole.pprr
            jsonPole("Tree Trimming Required") = pole.treeWork
            jsonPole("Outage Management Required") = pole.outage
            jsonPole("Site Coordination Required") = pole.sheet.Range("SITECOOR")
            jsonPole("Pole Denial Reason") = IIf(pole.denied, pole.sheet.Range("SUMSHEET14"), "")
            jsonPole("As-Is %") = FixPercentage(pole.sheet.Range("ASIS"))
            jsonPole("New App %") = FixPercentage(pole.sheet.Range("NEWAPP"))
            jsonPole("Existing Height of Lowest Power Attachment") = Utilities.inchesToFeetInches(pole.lowestPower)
            jsonPole("Existing Height of Streetlight from bottom of bracket") = Utilities.inchesToFeetInches(pole.slBottomBracketHeight)
            jsonPole("Equipment") = getEquipment(pole.sheet.Range("EQUIPMENT"))
            jsonPole("Clearance Requirement") = getClearance(pole)
            jsonPole("Room to Guy") = Not (pole.sheet.Range("ROOMTOGUY") = "NO")
            jsonPole("Guying Comments") = IIf(pole.sheet.Range("GUYCOMMENT").value <> "", pole.sheet.Range("GUYCOMMENT").value, "")
            jsonPole("CE Work - Violations - Emergency Work") = pole.sheet.Range("EW").offSet(0, 1).value <> ""
            jsonPole("CE Work - Violations - Pole Loading") = pole.sheet.Range("PL").offSet(0, 1).value <> ""
            jsonPole("CE Work - Violations - Safety Zone Violation") = pole.sheet.Range("SZV").offSet(0, 1).value <> "" Or pole.sheet.Range("SZV").offSet(1, 0).value <> ""
            jsonPole("CE Work - Violations - Midspan Seperation") = pole.sheet.Range("MSS").offSet(0, 1).value <> "" Or pole.sheet.Range("MSS").offSet(1, 0).value <> ""
            jsonPole("CE Work - Violations - Deterioration") = pole.sheet.Range("DET").offSet(0, 1).value <> ""
            jsonPole("CE Work - Violations - Missing or Broken CE Strain Ins") = pole.sheet.Range("MBCSI").offSet(0, 2).value <> ""
            jsonPole("CE Work - Violations - Missing or Broken CE Guy Marker") = pole.sheet.Range("MBCGM").offSet(0, 2).value <> ""
            jsonPole("CE Work - Violations - Missing or Broken Guys") = pole.sheet.Range("CEMBG").offSet(0, 2).value <> ""
            jsonPole("CE Work - Violations - Missing or Broken STL Molding") = pole.sheet.Range("MBSM").offSet(0, 1).value = "YES"
            jsonPole("CE Work - Violations - UnBonded Streetlight") = pole.sheet.Range("BONDED").value = "NO"
            jsonPole("CE Work - Violations - Other") = pole.sheet.Range("CEOTHER").offSet(0, 1).value <> "" Or pole.sheet.Range("CEOTHER").offSet(1, 0).value <> ""
            jsonPole("Communication Work - Violations - Separation") = pole.sheet.Range("SEP").offSet(0, 1).value <> "" Or pole.sheet.Range("SEP").offSet(1, 0).value <> ""
            jsonPole("Communication Work - Violations - Clearance") = pole.sheet.Range("CLEARANCE").offSet(0, 1).value <> "" Or pole.sheet.Range("CLEARANCE").offSet(1, 0).value <> ""
            jsonPole("Communication Work - Violations - Illegal Equipment") = pole.sheet.Range("IE").offSet(0, 1).value <> "" Or pole.sheet.Range("IE").offSet(1, 0).value <> ""
            jsonPole("Communication Work - Violations - Illegal Anchor") = pole.sheet.Range("IA").offSet(0, 1).value <> "" Or pole.sheet.Range("IA").offSet(1, 0).value <> ""
            jsonPole("Communication Work - Violations - Missing or Broken CE Strain Ins") = pole.sheet.Range("MBSI").offSet(0, 1).value <> "" Or pole.sheet.Range("MBSI").offSet(1, 0).value <> ""
            jsonPole("Communication Work - Violations - Missing or Broken CE Guy Marker") = pole.sheet.Range("MBGM").offSet(0, 1).value <> "" Or pole.sheet.Range("MBGM").offSet(1, 0).value <> ""
            jsonPole("Communication Work - Violations - Missing or Broken Guys") = pole.sheet.Range("COMMBG").offSet(0, 1).value <> "" Or pole.sheet.Range("COMMBG").offSet(1, 0).value <> ""
            jsonPole("Communication Work - Violations - Other") = pole.sheet.Range("COMOTHER").offSet(0, 1).value <> "" Or pole.sheet.Range("COMOTHER").offSet(1, 0).value <> ""
            
            Set results = determineEngineeringCost(pole)
            If InStr(pole.sheet.Range("SUMSHEET11").value, "2") > 0 And InStr(pole.sheet.Range("SUMSHEET11").value, "3") > 0 Then
                jsonPole("Alt 2 Cost") = "$" & results(2)
                jsonPole("Alt 3 Cost") = "$" & results(2)
                jsonPole("Total Alt Cost") = "$" & Round(CDbl(results(2)) * 2, 2)
            ElseIf InStr(pole.sheet.Range("SUMSHEET11").value, "2") > 0 Then
                jsonPole("Alt 2 Cost") = "$" & results(2)
                jsonPole("Alt 3 Cost") = "$0.00"
                jsonPole("Total Alt Cost") = "$" & results(2)
            ElseIf InStr(pole.sheet.Range("SUMSHEET11").value, "3") > 0 Then
                jsonPole("Alt 2 Cost") = "$0.00"
                jsonPole("Alt 3 Cost") = "$" & results(2)
                jsonPole("Total Alt Cost") = "$" & results(2)
            Else
                jsonPole("Alt 2 Cost") = "N/A"
                jsonPole("Alt 3 Cost") = "N/A"
                jsonPole("Total Alt Cost") = "N/A"
            End If
            
            Dim comBundles As Collection: Set comBundles = getCommBundles(pole)
            For i = 1 To comBundles.count
                jsonPole("Existing Attacher " & i & " - Owner") = comBundles(i)(1)
                jsonPole("Existing Attacher " & i & " - Attach Height") = comBundles(i)(2)
                jsonPole("Existing Attacher " & i & " - Attachment Modification") = comBundles(i)(3)
                jsonPole("Existing Attacher " & i & " - Midspans") = comBundles(i)(4)
            Next i
            
            json("Poles").Add pole.poleNumber, jsonPole
        End If
    Next pole
    
    jsonText = JsonConverter.ConvertToJson(json, Whitespace:=2)
    filePath = Environ("USERPROFILE") & "\Downloads\" & project.notification & " - Spida Studio.json"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.CreateTextFile(filePath, True, False)
    
    file.Write jsonText
    file.Close
    
    MsgBox "Finished Exporting Studio JSON"
End Sub

Private Function getFramingType(pole As pole) As String
    If pole.tangent Then
        getFramingType = "Tangent"
    ElseIf pole.angle Then
        getFramingType = "Angle"
    ElseIf pole.corner Then
        getFramingType = "Corner"
    ElseIf pole.deadend Then
        getFramingType = "Dead End"
    ElseIf pole.buck Then
        getFramingType = "Buck"
    Else
        getFramingType = "Sec Only"
    End If
End Function

Private Function FixPercentage(ByVal inputString As Variant) As String
    inputString = CStr(inputString)
    inputString = Replace(inputString, vbLf, " ")
    inputString = Replace(inputString, vbNewLine, " ")
    inputString = Replace(inputString, vbCrLf, " ")
    inputString = Replace(inputString, "\", "\\")
    inputString = Replace(inputString, """", "\""")
    inputString = Replace(inputString, vbTab, "")
    inputString = Trim(inputString)
    If InStr(inputString, ".") > 0 Then
        inputString = Replace(inputString, ".", "")
        If Len(inputString) = 2 Then inputString = inputString & "0"
        If IsNumeric(inputString) Then inputString = CStr(CDbl(inputString)) & "%"
    End If
    If inputString <> "" And IsNumeric(inputString) Then
        inputString = inputString & "%"
    End If
    FixPercentage = inputString
End Function

Private Function getEquipment(inputString As Variant) As String
    inputString = CStr(inputString)
    If InStr(LCase(inputString), "x") Or InStr(LCase(inputString), "tran") > 0 Then
        getEquipment = "Transformer"
    ElseIf InStr(LCase(inputString), "cap") > 0 Then
        getEquipment = "Cap Bank"
    ElseIf InStr(LCase(inputString), "iso") > 0 Then
        getEquipment = "Isolator"
    ElseIf InStr(LCase(inputString), "auto") > 0 Then
        getEquipment = "Autobooster"
    ElseIf InStr(LCase(inputString), "reg") > 0 Then
        getEquipment = "Regulator"
    ElseIf InStr(LCase(inputString), "rec") > 0 Then
        getEquipment = "Recloser"
    ElseIf InStr(LCase(inputString), "asc") > 0 Or InStr(LCase(inputString), "aerial") > 0 Then
        getEquipment = "Aerial Spacer Cable"
    ElseIf InStr(LCase(inputString), "fuse") > 0 Then
        getEquipment = "Fuse"
    Else
        getEquipment = ""
    End If
End Function

Private Function getClearance(pole As pole) As String
    Dim Span As Span
    Dim highestClearance As Integer: highestClearance = 186
    For Each Span In pole.spans
        If Utilities.convertToInches(Span.clearance) > highestClearance Then lowestClearance = Span.clearance
    Next Span
    
    getClearance = Utilities.inchesToFeetInches(highestClearance)
End Function

Private Function getCommBundles(pole As pole) As Collection
    Dim results As Collection: Set results = New Collection
    Dim cell As Range
    Dim Span As Span
    Dim wire As wire
    For i = 1 To 6
        Dim commOwner As String
        Set cell = pole.sheet.Range("COMM" & i)
        Dim commBundle As Collection: Set commBundle = New Collection
        commOwner = cell.value
        If InStr(commOwner, "COMM #") > 0 Then Exit For
        
        Dim heights As Scripting.Dictionary: Set heights = New Scripting.Dictionary
        Dim heightBundle As String: heightBundle = ""
        For j = 1 To 100 Step 2
            If cell.offSet(1 + j, 0).value = "" Then Exit For
            If Not heights.exists(cell.offSet(1 + j, 0).value) Then heights.Add cell.offSet(1 + j, 0).value, Nothing
        Next j
        If heights.count > 0 Then heightBundle = Utilities.JoinCollection(heights, ", ")
        
        Dim modifications As Scripting.Dictionary: Set modifications = New Scripting.Dictionary
        Dim modificationBundle As String: modificationBundle = ""
        For j = 1 To 100 Step 2
            If cell.offSet(1 + j, 0).offSet(0, 1).value = "" Then Exit For
            If Not modifications.exists(cell.offSet(1 + j, 0).offSet(0, 1).value) Then
                If ThisWorkbook.RemoveParentheses(cell.offSet(1 + j, 0).value) <> ThisWorkbook.RemoveParentheses(cell.offSet(1 + j, 0).offSet(0, 1).value) Then
                    modifications.Add cell.offSet(1 + j, 0).offSet(0, 1).value, Nothing
                End If
            End If
        Next j
        If modifications.count > 0 And Utilities.JoinCollection(heights, ", ") <> Utilities.JoinCollection(modifications, ", ") Then modificationBundle = Utilities.JoinCollection(modifications, ", ")
        
        Dim midspans As Collection: Set midspans = New Collection
        Dim midspanBundle As String: midspanBundle = ""
        For Each Span In pole.spans
            For Each wire In Span.commWires
                If wire.owner = commOwner Then
                    midspans.Add Utilities.inchesToFeetInches(wire.midspans(Span.spanSlot)) & " " & Span.direction
                End If
            Next wire
        Next Span
        If midspans.count > 0 Then midspanBundle = Utilities.JoinCollection(midspans, ", ")
        
        commBundle.Add commOwner
        commBundle.Add heightBundle
        commBundle.Add modificationBundle
        commBundle.Add midspanBundle
        
        results.Add commBundle
    Next i
    
    Set getCommBundles = results
End Function
