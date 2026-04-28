Attribute VB_Name = "AHK"
Public drawAdjacentPoles, drawServices, drawTrees, drawGuys, drawStreetlights, drawTransformers As Boolean
Public primaryLevel As Integer
Public finished As scripting.Dictionary

Sub generateAHKJson()
    Dim json As scripting.Dictionary
    Set json = New scripting.Dictionary
    
    Dim generatedPoles As scripting.Dictionary
    Set generatedPoles = New scripting.Dictionary
    
    Set finished = New scripting.Dictionary
    
    Dim result As scripting.Dictionary
    Dim startingSheet As Worksheet
    Dim sheet As Worksheet
    
    Call LogMessage.SendLogMessage("generateAHKJson")
    
    folderPath = Replace(ThisWorkbook.sheets("Control").Range("AHKPATH").Value, "AHK Folder Path: ", "")
    If Dir(folderPath, vbDirectory) = "" Or Trim(folderPath) = "" Then
        Dim fileDiag As FileDialog
        Set fileDiag = Application.FileDialog(msoFileDialogFolderPicker)
        With fileDiag
            .AllowMultiSelect = False
            .Title = "Select a folder "
            If .Show = -1 Then folderPath = Application.FileDialog(msoFileDialogFilePicker).SelectedItems.item(1) & Application.PathSeparator Else Exit Sub
        End With
        ThisWorkbook.sheets("Control").Range("AHKPATH").Value = "AHK Folder Path: " & folderPath
    End If
    drawAdjacentPoles = ThisWorkbook.sheets("Control").Range("AHKAP").Value = True
    drawServices = ThisWorkbook.sheets("Control").Range("AHKA").Value = True
    drawTrees = ThisWorkbook.sheets("Control").Range("AHKT").Value = True
    drawGuys = ThisWorkbook.sheets("Control").Range("AHKG").Value = True
    drawStreetlights = ThisWorkbook.sheets("Control").Range("AHKS").Value = True
    drawTransformers = ThisWorkbook.sheets("Control").Range("AHKX").Value = True
    
    json("drawAdjacentPoles") = drawAdjacentPoles
    json("drawServices") = drawServices
    json("drawTrees") = drawTrees
    json("drawGuys") = drawGuys
    json("drawStreetlights") = drawStreetlights
    json("drawTransformers") = drawTransformers
    
    Do While True
        primaryLevel = 1
        Set startingSheet = Nothing
        For Each sheet In ThisWorkbook.sheets
            If sheet.name <> "4 Spans" And sheet.name <> "8 Spans" And sheet.name <> "12 Spans" Then
                If sheet.Cells(2, 2).Value = "Notification:" Then
                    If Not generatedPoles.Exists(CStr(sheet.Range("POLENUM"))) Then
                        Set result = getConnectedPoles(sheet)
                        If result.count < 2 Then
                            Set startingSheet = sheet
                            Exit For
                        End If
                    End If
                End If
            End If
        Next sheet
        If startingSheet Is Nothing Then Exit Do
        Call generateJson(json, generatedPoles, startingSheet, starting:=True)
    Loop
    
    Dim jsonText As String
    jsonText = JsonConverter.ConvertToJson(json, Whitespace:=2)
    
    Dim filePath As String
    filePath = folderPath & "\AHK.json"
    
    Dim fso As Object
    Dim file As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.CreateTextFile(filePath, True, False)
    
    file.Write jsonText
    file.Close
End Sub

Private Function getServices(sheet As Worksheet) As scripting.Dictionary
    Dim connectedServices As scripting.Dictionary: Set connectedServices = New scripting.Dictionary
    
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Pattern = "^(.*?)(?=\()"
        .IgnoreCase = True
        .Global = False
    End With

    count = 0
    For i = 1 To 12
        For Each name In sheet.names
            If name.name = "'" & sheet.name & "'" & "!TOPOLE" & i Then
                If Trim(Replace(sheet.Range("TOPOLE" & i), "-", "")) <> "" Then
                    If regex.test(sheet.Range("TOPOLE" & i)) Then
                        Set matches = regex.Execute(sheet.Range("TOPOLE" & i))
                        If Not SheetExists(matches(0)) Then
                            For j = 0 To 100
                                If sheet.Range("UTTYPE").offset(j, 0).Interior.color <> 16312794 Then Exit For
                                If InStr(sheet.Range("UTTYPE").offset(j, 0), "SVC") > 0 Then
                                    If Replace(sheet.Range("UTMIDSPAN" & i).offset(j, 0), "-", "") <> "" Then
                                        addressCellValue = sheet.Range("UTMIDSPAN" & i).offset(j, 0).Value
                                        addressCellValue = Trim(Replace(Replace(addressCellValue, "-", ""), "0'0""", ""))
                                        If InStr(addressCellValue, "(") > 0 Then
                                            leftPar = InStr(addressCellValue, "(") + 1
                                            address = Mid(addressCellValue, leftPar, Len(addressCellValue) - leftPar)
                                        Else
                                            address = addressCellValue
                                        End If
                                        Dim result As Collection: Set result = New Collection
                                        result.Add i
                                        result.Add address
                                        connectedServices.Add count, result
                                        count = count + 1
                                        Exit For
                                    End If
                                End If
                            Next j
                        End If
                    End If
                End If
            End If
        Next name
    Next i
    
    Set getServices = connectedServices
End Function

Private Function getConnectedOtherPoles(ByRef generatedPoles As scripting.Dictionary, sheet As Worksheet) As scripting.Dictionary
    Dim connectedPoles As scripting.Dictionary: Set connectedPoles = New scripting.Dictionary
    
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Pattern = "^(.*?)(?=\()"
        .IgnoreCase = True
        .Global = False
    End With

    count = 0
 
    For i = 1 To 12
        For Each name In sheet.names
            If name.name = "'" & sheet.name & "'" & "!TOPOLE" & i Then
                If Trim(Replace(sheet.Range("TOPOLE" & i), "-", "")) <> "" Then
                    If regex.test(sheet.Range("TOPOLE" & i)) Then
                        Set matches = regex.Execute(sheet.Range("TOPOLE" & i))
                        If Not SheetExists(matches(0)) Then
                            If sheet.Range("DL") <> "" Then
                                For j = 0 To 100
                                    If sheet.Range("UTTYPE").offset(j, 0).Interior.color <> 16312794 Then Exit For
                                    If InStr(sheet.Range("UTTYPE").offset(j, 0), "SVC") = 0 Then
                                        If Replace(sheet.Range("UTMIDSPAN" & i).offset(j, 0), "-", "") <> "" Then
                                            connectedPoles.Add count, i
                                            count = count + 1
                                            Exit For
                                        End If
                                    End If
                                Next j
                            End If
                        End If
                    End If
                End If
            End If
        Next name
    Next i
    
    Set getConnectedOtherPoles = connectedPoles
End Function

Private Function getConnectedPoles(sheet As Worksheet) As scripting.Dictionary
    Dim connectedPoles As scripting.Dictionary: Set connectedPoles = New scripting.Dictionary
    
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Pattern = "^(.*?)(?=\()"
        .IgnoreCase = True
        .Global = False
    End With

    For i = 1 To 12
        For Each name In sheet.names
            If name.name = "'" & sheet.name & "'" & "!TOPOLE" & i Then
                If Trim(Replace(sheet.Range("TOPOLE" & i), "-", "")) <> "" Then
                    If regex.test(sheet.Range("TOPOLE" & i)) Then
                        Set matches = regex.Execute(sheet.Range("TOPOLE" & i))
                        If SheetExists(matches(0)) Then
                            'For j = 0 To 100
                               ' If sheet.Range("UTTYPE").Offset(j, 0).Interior.color <> 16312794 Then Exit For
                                'If InStr(sheet.Range("UTTYPE").Offset(j, 0), uttype) > 0 Then
                                   ' If Replace(sheet.Range("UTMIDSPAN" & i).Offset(j, 0), "-", "") <> "" Then
                                        connectedPoles.Add matches(0), i
                                        'Exit For
                                   ' End If
                               ' End If
                            'Next j
                        End If
                    End If
                End If
            End If
        Next name
    Next i
    
    Set getConnectedPoles = connectedPoles
End Function

Private Function getItems(sheet As Worksheet, ByVal i As String, uttype As String) As Collection
    Dim items As Collection: Set items = New Collection
    
    For j = 0 To 100
        If sheet.Range("UTTYPE").offset(j, 0).Interior.color <> 16312794 Then Exit For
        If InStr(sheet.Range("UTTYPE").offset(j, 0), uttype) > 0 Then
            If Replace(sheet.Range("UTMIDSPAN" & i).offset(j, 0), "-", "") <> "" Then
                Dim item As scripting.Dictionary: Set item = New scripting.Dictionary
                item("size") = OnlyNumbers(sheet.Range("UTSIZE").offset(j, 0), True)
                item("type") = uttype
                If uttype = "PRI" Then
                    If InStr(sheet.Range("UTSIZE").offset(j, 0), "Ø") > 1 Then
                        item("phase") = Mid(sheet.Range("UTSIZE").offset(j, 0), InStr(sheet.Range("UTSIZE").offset(j, 0), "Ø") - 1, 1)
                        If CStr(item("phase")) = "1" Then
                            item("phase") = "Z"
                        ElseIf CStr(item("phase")) = "2" Then
                            item("phase") = "XZ"
                        Else
                            item("phase") = "3"
                        End If
                        item("size") = Left(item("size"), Len(item("size")) - 1)
                        neutralSpanCount = 0
                        Dim neutrals As Collection: Set neutrals = New Collection
                        neutralSize = ""
                        secondarySpanCount = 0
                        Dim secondaries As Collection: Set secondaries = New Collection
                        neutralShareHeight = False
                        For k = 0 To 100
                            If sheet.Range("UTTYPE").offset(k, 0).Interior.color <> 16312794 Then Exit For
                            If InStr(sheet.Range("UTTYPE").offset(k, 0), "NEUT") > 0 Then
                                neutrals.Add k
                                If Replace(sheet.Range("UTMIDSPAN" & i).offset(k, 0), "-", "") <> "" Then
                                    neutralSize = Utilities.OnlyNumbers(sheet.Range("UTSIZE").offset(k, 0), True)
                                    neutralSpanCount = neutralSpanCount + 1
                                    netrualHeight = Utilities.convertToInches(sheet.Range("UTHEIGHT").offset(k, 0))
                                    primaryHeight = Utilities.convertToInches(sheet.Range("UTHEIGHT").offset(j, 0))
                                    If (Abs(primaryHeight - netrualHeight) < 18) Then
                                        neutralShareHeight = True
                                    End If
                                End If
                            ElseIf InStr(sheet.Range("UTTYPE").offset(k, 0), "SEC") > 0 Or InStr(sheet.Range("UTTYPE").offset(k, 0), "OW") > 0 Then
                                secondaries.Add k
                                If Replace(sheet.Range("UTMIDSPAN" & i).offset(k, 0), "-", "") <> "" Then
                                    secondarySpanCount = secondarySpanCount + 1
                                End If
                            End If
                        Next k
                        If (neutralShareHeight) Then
                            item("config") = "N"
                            item("neutralSize") = neutralSize
                        ElseIf (secondarySpanCount > 0) Then
                            item("config") = "SN"
                        ElseIf (neutralSpanCount > 0) Then
                            item("config") = "NB"
                            item("neutralSize") = neutralSize
                        ElseIf (neutrals.count > 0 And secondaries.count = 0) Then
                            item("config") = "NB"
                            item("neutralSize") = neutralSize
                        ElseIf (secondaries.count > 0 And neutrals.count = 0) Then
                            item("config") = "SN"
                        ElseIf (secondaries.count = 0 And neutrals.count = 0) Then
                            item("config") = "N"
                            item("neutralSize") = item("size")
                        Else
                            closestSecAngleDif = 360
                            closestNeutAngleDif = 360
                            Set distanceAngle = getDistanceAngle(sheet, i)
                            angle = distanceAngle(2)
                            For Each secondary In secondaries
                                Set results = getClosestAngle(sheet, secondary, angle)
                                If closestSecAngleDif > results(1) Then closestSecAngleDif = results(1)
                            Next secondary
                            
                            For Each neutral In neutrals
                                Set results = getClosestAngle(sheet, neutral, angle)
                                If closestNeutAngleDif > results(1) Then
                                    closestNeutAngleDif = results(1)
                                    neutralSize = results(2)
                                End If
                            Next neutral

                            If closestSecAngleDif <= closestNeutAngleDif And closestSecAngleDif < 30 Then
                                item("config") = "SN"
                            ElseIf closestNeutAngleDif <= closestSecAngleDif And closestNeutAngleDif < 30 Then
                                item("config") = "NB"
                                item("neutralSize") = neutralSize
                            Else
                                item("config") = "N"
                                item("neutralSize") = item("size")
                            End If
                        End If
                    End If
                ElseIf uttype = "SEC" Then
                    size = sheet.Range("UTSIZE").offset(j, 0)
                    If InStr(size, "TX") > 0 Then
                        item("size") = item("size") & "TX"
                    ElseIf InStr(size, "DX") > 0 Then
                        item("size") = item("size") & "DX"
                    ElseIf InStr(size, "QX") > 0 Then
                        item("size") = item("size") & "QX"
                    ElseIf InStr(size, "AWAC") > 0 Then
                        item("size") = item("size") & "AWAC"
                    End If
                End If
                items.Add item
            End If
        End If
    Next j
    Set getItems = items
End Function

Private Function getClosestAngle(sheet As Worksheet, ByVal k As Integer, ByVal angle As Integer) As Collection
    Set closestAngle = New Collection
    
    smallestAngleDif = 360
    size = ""
    
    For i = 1 To 12
        For Each name In sheet.names
            If name.name = "'" & sheet.name & "'" & "!TOPOLE" & i Then
                If Trim(Replace(sheet.Range("TOPOLE" & i), "-", "")) <> "" Then
                    If Replace(sheet.Range("UTMIDSPAN" & i).offset(k, 0), "-", "") <> "" Then
                        Set results = getDistanceAngle(sheet, i)
                        angleDif = Abs(angle - results(2))
                        If smallestAngleDif > angleDif Then
                            smallestAngleDif = angleDif
                            size = OnlyNumbers(sheet.Range("UTSIZE").offset(k, 0), True)
                        End If
                    End If
                End If
            End If
        Next name
    Next i
    
    closestAngle.Add angleDif
    closestAngle.Add size
    Set getClosestAngle = closestAngle
End Function

Private Sub generateJson(ByRef json As scripting.Dictionary, ByRef generatedPoles As scripting.Dictionary, sheet As Worksheet, Optional starting As Boolean = False)
    Dim pole As scripting.Dictionary: Set pole = New scripting.Dictionary
    
    locationAdjacent = False
    drawConductors = False
    
    If sheet.Range("DL") = "" Then
        Set regex = CreateObject("VBScript.RegExp")
        With regex
            .Pattern = "^(.*?)(?=\()"
            .IgnoreCase = True
            .Global = False
        End With
        For i = 1 To 12
            nameExists = sheet.Evaluate("ISREF(" & "TOPOLE" & i & ")")
            If nameExists Then
                If Trim(Replace(sheet.Range("TOPOLE" & i), "-", "")) <> "" Then
                    If regex.test(sheet.Range("TOPOLE" & i)) Then
                        Set matches = regex.Execute(sheet.Range("TOPOLE" & i))
                        If SheetExists(matches(0)) Then
                            Dim oSheet As Worksheet
                            For Each oSheet In ThisWorkbook.sheets
                                If ThisWorkbook.RemoveParentheses(oSheet.name) = matches(0) Then
                                    If oSheet.Range("DL") <> "" Then
                                        locationAdjacent = True
                                        Exit For
                                    End If
                                End If
                            Next oSheet
                        End If
                    End If
                End If
            End If
        Next i
    Else
        locationAdjacent = True
    End If
    
    If ThisWorkbook.sheets("Control").Range("AHKC").Value Then
        If ThisWorkbook.sheets("Control").OptionButtons("Option Button 1").Value = 1 Then
            If locationAdjacent Then drawConductors = True
        Else
            drawConductors = True
        End If
    End If
    
    pole("ceid") = sheet.Range("CEID")
    If Trim(sheet.Range("HEIGHT")) <> "" Then
        pole("height") = Utilities.OnlyNumbers(ThisWorkbook.RemoveParentheses(sheet.Range("HEIGHT")))
    End If
    pole("class") = sheet.Range("class")
    pole("location") = CStr(sheet.Range("DL"))
    If starting Then pole("starting") = True
    If drawTrees Then pole("tree") = (sheet.Range("TREE") = True Or sheet.Range("TREE2") = "YES")
    If drawTransformers Then pole("transformer") = (InStr(sheet.Range("EQUIPMENT"), "XFMR") > 0 Or InStr(sheet.Range("INVENTORY"), "TRANSFORMER") > 0)
    If drawStreetlights Then pole("streetlight") = sheet.Range("STLTBRKT").Value <> "" And sheet.Range("STLTBRKT").Value <> "N/A"
    If sheet.Range("REPLACEPOLE") Then pole("replace") = True
    nameExists = sheet.Evaluate("ISREF(" & "HVD" & ")")
    If nameExists Then
        If Trim(sheet.Range("HVD").Value) <> "" Then
            pole("hvd") = sheet.Range("HVD").Value
        End If
    End If
    
    Dim guys As Collection: Set guys = New Collection
    Dim services As Collection: Set services = New Collection
    Dim connections As Collection: Set connections = New Collection
    Dim jobConnections As Collection: Set jobConnections = New Collection
    
    Dim poleNumber As String
    poleNumber = sheet.Range("POLENUM")
    generatedPoles.Add poleNumber, Nothing
    
    Dim connectedServices As scripting.Dictionary: Set connectedServices = getServices(sheet)
    Dim connectedOtherPoles As scripting.Dictionary: Set connectedOtherPoles = getConnectedOtherPoles(generatedPoles, sheet)
    Dim connectedPoles As scripting.Dictionary: Set connectedPoles = getConnectedPoles(sheet)
    
    If drawGuys Then
        Dim guy As scripting.Dictionary
        For i = 0 To 15
            If Trim(sheet.Range("ANCHOROWNER").offset(i, 0).Value) = "" Then Exit For
            If Trim(sheet.Range("ANCHOROWNER").offset(i, 0).Value) = "New App" Then Exit For
            If UCase(Trim(sheet.Range("ANCHOROWNER").offset(i, 0).Value)) = "CONSUMERS ENERGY" Then
                Set guy = New scripting.Dictionary
                guy("angle") = OnlyNumbers(sheet.Range("ANCHORDIRECTION").offset(i, 0).Value)
                guys.Add guy
            End If
        Next i
    End If
    
    If drawServices And locationAdjacent Then
        Dim service As scripting.Dictionary
        Dim connectedService As Variant
        For Each connectedService In connectedServices
            Set service = New scripting.Dictionary
            Set distanceAngle = getDistanceAngle(sheet, connectedServices(connectedService)(1))
            service("distance") = distanceAngle(1)
            service("angle") = distanceAngle(2)
            service("address") = connectedServices(connectedService)(2)
            services.Add service
        Next connectedService
    End If
    
    Dim connection As scripting.Dictionary
    Dim connectedPole As Variant
    If drawAdjacentPoles Then
        For Each connectedPole In connectedOtherPoles
            Set connection = New scripting.Dictionary
            Set distanceAngle = getDistanceAngle(sheet, connectedOtherPoles(connectedPole))
            connection("distance") = distanceAngle(1)
            connection("angle") = distanceAngle(2)
            If drawGuys Then connection("spanGuys") = getItems(sheet, connectedOtherPoles(connectedPole), "SPG").count
            If drawConductors Then
                Set connection("secondaries") = getItems(sheet, connectedOtherPoles(connectedPole), "SEC")
                Set openWires = getItems(sheet, connectedOtherPoles(connectedPole), "OW")
                openWireString = ""
                For Each OPENWIRE In openWires
                    openWireString = openWireString & OPENWIRE("size") & "-"
                Next OPENWIRE
                If openWireString <> "" Then
                    openWireString = Left(openWireString, Len(openWireString) - 1)
                    Set item = New scripting.Dictionary
                    item("size") = openWireString
                    item("type") = "OW"
                    connection("secondaries").Add item
                End If
                If primaryLevel <= connection("secondaries").count Then primaryLevel = connection("secondaries").count + 1
                Set connection("primaries") = getItems(sheet, connectedOtherPoles(connectedPole), "PRI")
            End If
            connections.Add connection
        Next connectedPole
    End If
    
    For Each connectedPole In connectedPoles
        If Not generatedPoles.Exists(CStr(connectedPole)) Then
            If SheetExists(connectedPole) Then
                For Each oSheet In ThisWorkbook.sheets
                    If ThisWorkbook.RemoveParentheses(oSheet.name) = connectedPole Then
                        Set connection = New scripting.Dictionary
                        connection("poleNumber") = connectedPole
                        Set distanceAngle = getDistanceAngle(sheet, connectedPoles(connectedPole))
                        connection("distance") = distanceAngle(1)
                        connection("angle") = distanceAngle(2)
                        If drawGuys Then connection("spanGuys") = getItems(sheet, connectedPoles(connectedPole), "SPG").count
                        If drawConductors Then
                            Set connection("secondaries") = getItems(sheet, connectedPoles(connectedPole), "SEC")
                            Set openWires = getItems(sheet, connectedPoles(connectedPole), "OW")
                            openWireString = ""
                            For Each OPENWIRE In openWires
                                openWireString = openWireString & OPENWIRE("size") & "-"
                            Next OPENWIRE
                            If openWireString <> "" Then
                                openWireString = Left(openWireString, Len(openWireString) - 1)
                                Set item = New scripting.Dictionary
                                item("size") = openWireString
                                item("type") = "OW"
                                connection("secondaries").Add item
                            End If
                            Set connection("primaries") = getItems(sheet, connectedPoles(connectedPole), "PRI")
                            If primaryLevel <= connection("secondaries").count Then primaryLevel = connection("secondaries").count + 1
                        End If
                        jobConnections.Add connection
                        Call generateJson(json, generatedPoles, oSheet)
                        Exit For
                    End If
                Next oSheet
            End If
        Else
            If finished.Exists(CStr(connectedPole)) Then
                Set connection = New scripting.Dictionary
                connection("poleNumber") = "skip"
                Set distanceAngle = getDistanceAngle(sheet, connectedPoles(connectedPole))
                connection("distance") = distanceAngle(1)
                connection("angle") = distanceAngle(2)
                If drawGuys Then connection("spanGuys") = getItems(sheet, connectedPoles(connectedPole), "SPG").count
                If drawConductors Then
                    Set connection("secondaries") = getItems(sheet, connectedPoles(connectedPole), "SEC")
                    Set openWires = getItems(sheet, connectedPoles(connectedPole), "OW")
                    openWireString = ""
                    For Each OPENWIRE In openWires
                        openWireString = openWireString & OPENWIRE("size") & "-"
                    Next OPENWIRE
                    If openWireString <> "" Then
                        openWireString = Left(openWireString, Len(openWireString) - 1)
                        Set item = New scripting.Dictionary
                        item("size") = openWireString
                        item("type") = "OW"
                        connection("secondaries").Add item
                    End If
                    Set connection("primaries") = getItems(sheet, connectedPoles(connectedPole), "PRI")
                    If primaryLevel <= connection("secondaries").count Then primaryLevel = connection("secondaries").count + 1
                End If
                connections.Add connection
            End If
        End If
    Next connectedPole
    
    finished.Add poleNumber, Nothing
    If starting Then pole("primaryLevel") = primaryLevel
    Set pole("guys") = guys
    Set pole("services") = services
    For Each connection In jobConnections
        connections.Add connection
    Next connection
    Set pole("connections") = connections
    Set json(poleNumber) = pole
    
End Sub

Private Function getDistanceAngle(sheet As Worksheet, ByVal i As String) As Collection
    distance = OnlyNumbers(sheet.Range("SPAN" & i).Value)
    angle = OnlyNumbers(Mid(sheet.Range("TOPOLE" & i).Value, InStr(sheet.Range("TOPOLE" & i).Value, "(")))
    
    Dim distanceAngle As Collection: Set distanceAngle = New Collection
    distanceAngle.Add distance
    distanceAngle.Add angle
    
    Set getDistanceAngle = distanceAngle
End Function

