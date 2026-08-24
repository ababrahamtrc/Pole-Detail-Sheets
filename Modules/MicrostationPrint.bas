Attribute VB_Name = "MicrostationPrint"
Sub GenerateMicrostationPrintFiles()
    Dim restartMicrostationNeeded As Boolean
    
    Call LogMessage.SendLogMessage("GenerateMicrostationPrint")
    
    restartMicrostationNeeded = False
    
    If Environ$("USERNAME") <> "aabraham" Then
        restartMicrostationNeeded = injectHotkey()
        If Not ForceInjectModuleToBentley("PrintGenerator") Then Exit Sub
        If Not ForceInjectUserFormToBentley("PrintOptions") Then Exit Sub
        Call FixJsonConverter("Private Function json_ParseObject(json_String As String, ByRef json_Index As Long) As Dictionary", "Private Function json_ParseObject(json_String As String, ByRef json_Index As Long) As Object")
        Call FixJsonConverter("Set json_ParseObject = New Dictionary", vbTab & "Set json_ParseObject = CreateObject(""Scripting.Dictionary"")")
        If Not ForceInjectModuleToBentley("JsonConverter") Then Exit Sub
    End If
    
    If Not generateJSON Then Exit Sub
    
    If restartMicrostationNeeded Then
        MsgBox "Restart Open Map Utilities if open for changes to take effect."
    End If
    
    MsgBox "Press '9' to generate the print on open map utilities. Script will have to be rerun for future generations."
End Sub

Sub FixJsonConverter(searchStr As String, replaceStr As String)
    Dim vbaProject As Object
    Dim vbaModule As Object
    Dim targetModuleName As String
    Dim i As Long
    Dim currentLine As String
    
    targetModuleName = "JsonConverter"

    Set vbaProject = Application.VBE.ActiveVBProject
    Set vbaModule = vbaProject.VBComponents(targetModuleName).CodeModule

    For i = 1 To vbaModule.CountOfLines
        currentLine = Trim(vbaModule.lines(i, 1))
        
        If currentLine = Trim(searchStr) Then
            vbaModule.ReplaceLine i, replaceStr
            Exit Sub
        End If
    Next i
End Sub

Sub ClipboardGISTokenURL()
    Dim url As String

    url = "javascript:(function(){if(window.__CE_TOKEN_SEARCHER__){alert('Token searcher already active.\n\nPan/zoom to trigger requests.');return;}window.__CE_TOKEN_SEARCHER__=true;window.__CE_LAST_TOKEN__=null;function emit(t,src){if(!t||t===window.__CE_LAST_TOKEN__)return;window.__CE_LAST_TOKEN__=t;var payload={token:t,captured_at:new Date().toISOString(),source:src};console.log('[CE TOKEN]',payload);try{navigator.clipboard.writeText(JSON.stringify(payload,null,2));}catch(e)" & _
        "{}var blob=new Blob([JSON.stringify(payload,null,2)],{type:'application/json'});var a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='arcgis_token.json';a.click();}var xsend=XMLHttpRequest.prototype.send;XMLHttpRequest.prototype.send=function(b){try{if(typeof b==='string'&&b.indexOf('token=')>-1){var p=new URLSearchParams(b);emit(p.get('token'),'XMLHttpRequest.send');}}catch(e){}return xsend.apply(this,arguments);};var ffetch=window.fetch;window.fetch=function(){try{var u=arguments[0];if(typeof u==='string'&&u.indexOf('to" & _
        "ken=')>-1){var q=u.split('?%27)[1];if(q){var p2=new URLSearchParams(q);emit(p2.get(%27token%27),%27fetch(url)%27);}}if(arguments[1]&&arguments[1].body&&typeof arguments[1].body===%27string%27&&arguments[1].body.indexOf(%27token=%27)>-1){var p3=new URLSearchParams(arguments[1].body);emit(p3.get(%27token%27),%27fetch(body)%27);}}catch(e){}return ffetch.apply(this,arguments);};alert(%27ArcGIS token searcher armed.\n\nPan, zoom, or toggle a layer to capture token.%27);})();"

    Dim DataObj As DataObject: Set DataObj = New DataObject
    DataObj.SetText url
    DataObj.PutInClipboard

    MsgBox "Bookmark url copied to clipboard." & vbLf & vbLf & "Go to browser and right click over bookmarks bars. Select 'Open Bookmark Manager' or 'Manage Favorites' depending on the browser." & vbLf & vbLf & "Then select 'Add new Bookmark' or 'Add Favorite', paste the copied url from this button under the URL section. For the Name, enter 'GIS Token'" & vbLf & vbLf & "Activate this bookmark when on the GIS website to automatically download GIS token JSONs to your downloads folder to use for the print and outages automation."

End Sub

Function generateJSON() As Boolean
    Dim json As Scripting.Dictionary: Set json = New Scripting.Dictionary
    Dim project As project: Set project = New project
    Call project.extractFromSheets
    
    Dim jsonPole As Scripting.Dictionary
    Dim jsonPoles As Object
    Dim jsonStreetlight As Scripting.Dictionary
    Dim pole As pole
    Dim otherPole As pole
    Dim Service As wire
    Dim Span As Span
    Dim anchor As anchor
    Dim wire As wire
    Dim otherWire As wire
    Dim dx As Double
    Dim dy As Double
    Dim x As Double
    Dim y As Double
    Dim x2 As Double
    Dim y2 As Double
    Dim spgCount As Integer
    Dim radAngle As Double
    
    Dim closestPole As Object
    
    Dim owWires As Collection
    Dim secWires As Collection
    Dim priWires As Collection
    
    Dim jsonText As String
    Dim filePath As String
    Dim fso As Object
    Dim file As Object
    
    Dim token As String: token = GetToken
    If Not testToken(token) Then
        MsgBox "Invalid token, get an up to date one from GIS."
        generateJSON = False
        Exit Function
    End If
    
    LoadingBar_Form2.Show vbModeless
    Call LoadingBar_Form2.InitProgress(1, True, 5)
    Call LoadingBar_Form2.UpdateProgress("Generating Print JSON", "Extracting Pole Detail Sheet Information", True)
    
    Dim poleCollections As Collection: Set poleCollections = findPoleGroups(project.poles)
    Dim poleCollection As Collection
    
    Dim usedPoles As Scripting.Dictionary: Set usedPoles = New Scripting.Dictionary
    
    Dim groupNumber As Integer: groupNumber = 1
    Dim lowestLatitude As Double
    Dim lowestLongitude As Double
    Dim highestLatitude As Double
    Dim highestLongitude As Double
    
    Dim priRegex As Object
    Set priRegex = CreateObject("VBScript.RegExp")
    
    priRegex.Pattern = "\s*(\d*)'[ OF]*\s*(\d)PH\s*(.*)PRI\s*\/\s*[\d' ]*[ OF]*\dPH(.*)PRI\s*(.*)"
    priRegex.Global = True
    priRegex.IgnoreCase = True
    
    Dim neutRegex As Object
    Set neutRegex = CreateObject("VBScript.RegExp")
    
    neutRegex.Pattern = "\s*(\d*)'[ OF]*\s*(.*)NEUT\s*\/\s*[\d' ]*[OF]*(.*)NEUT\s*(.*)"
    neutRegex.Global = True
    neutRegex.IgnoreCase = True
    
    Dim primaryReconductoringDict As Scripting.Dictionary
    Set primaryReconductoringDict = New Scripting.Dictionary
    Dim crewNotes As String
    
    For Each pole In project.poles
        If LoadingBar_Form2.gTotal = 0 Then Exit Function
        If pole.location <> "" Then
            crewNotes = pole.alt1
            Call Utilities.applyStandardAbbreviations(crewNotes)
            lines = Split(crewNotes, vbLf)
            For Each line In lines
                If InStr(line, "PRI") > 0 Then
                    If priRegex.test(line) Then
                        Set matches = priRegex.Execute(line)
                        Dim primaryReconductored As Scripting.Dictionary: Set primaryReconductored = New Scripting.Dictionary
                        primaryReconductored("distance") = CInt(matches(0).SubMatches(0))
                        primaryReconductored("phase") = CInt(matches(0).SubMatches(1))
                        primaryReconductored("size1") = Utilities.OnlyNumbers(matches(0).SubMatches(2), True)
                        primaryReconductored("size2") = Utilities.OnlyNumbers(matches(0).SubMatches(3), True)
                        For Each Span In pole.spans
                            If Round(Span.distance, 0) = primaryReconductored("distance") Then
                                correctSize = False
                                For Each wire In Span.utilWires
                                    If wire.componentType = "PRI" And InStr(wire.size, primaryReconductored("size1")) > 0 Then
                                        If wire.phase = primaryReconductored("phase") Then correctSize = True: Exit For
                                    End If
                                Next wire
                                If correctSize Then
                                    lengthCount = 0
                                    For Each otherSpan In pole.spans
                                        If Round(otherSpan.distance, 0) = primaryReconductored("distance") Then lengthCount = lengthCount + 1
                                    Next otherSpan
                                
                                    Set otherPole = Utilities.getPole(Span.otherPole)
                                    If Not otherPole Is Nothing Then
                                        If lengthCount = 1 Or Left(Span.direction, 1) = Left(matches(0).SubMatches(4), 1) Then
                                            primaryReconductored("otherPole") = Span.otherPole
                                            
                                            For Each line2 In lines
                                                If neutRegex.test(line2) Then
                                                    Set matches = neutRegex.Execute(line2)
                                                    If primaryReconductored("distance") = CInt(matches(0).SubMatches(0)) Then
                                                        primaryReconductored("neutSize") = matches(0).SubMatches(2)
                                                    End If
                                                End If
                                            Next line2
                                            
                                            If Not primaryReconductoringDict.exists(pole.poleNumber) Then Set primaryReconductoringDict(pole.poleNumber) = New Scripting.Dictionary
                                            If Not primaryReconductoringDict(pole.poleNumber).exists(Span.spanSlot) Then Set primaryReconductoringDict(pole.poleNumber)(Span.spanSlot) = New Collection
                                            primaryReconductoringDict(pole.poleNumber)(Span.spanSlot).Add primaryReconductored
                                            For Each otherSpan In otherPole.spans
                                                If pole.poleNumber = otherSpan.otherPole Then
                                                    If Not primaryReconductoringDict.exists(otherPole.poleNumber) Then Set primaryReconductoringDict(otherPole.poleNumber) = New Scripting.Dictionary
                                                    If Not primaryReconductoringDict(otherPole.poleNumber).exists(otherSpan.spanSlot) Then Set primaryReconductoringDict(otherPole.poleNumber)(otherSpan.spanSlot) = New Collection
                                                    primaryReconductoringDict(otherPole.poleNumber)(otherSpan.spanSlot).Add primaryReconductored
                                                    Exit For
                                                End If
                                            Next otherSpan
                                        End If
                                    End If
                                End If
                            End If
                        Next Span
                    End If
                End If
            Next line
        End If
    Next pole
    
    Dim secRegex As Object
    Set secRegex = CreateObject("VBScript.RegExp")
    
    secRegex.Pattern = "\s*(\d*)'[ OF]*\s*(.*)SEC\s*\/\s*(.*)SEC\s*(.*)"
    secRegex.Global = True
    secRegex.IgnoreCase = True
    
    Dim secReconductoringDict As Scripting.Dictionary
    Set secReconductoringDict = New Scripting.Dictionary
    
    For Each pole In project.poles
        If LoadingBar_Form2.gTotal = 0 Then Exit Function
        If pole.location <> "" Then
            crewNotes = pole.alt1
            Call Utilities.applyStandardAbbreviations(crewNotes)
            lines = Split(crewNotes, vbLf)
            For Each line In lines
                If InStr(line, "SEC") > 0 Then
                    If secRegex.test(line) Then
                        Set matches = secRegex.Execute(line)
                        Set secReconductored = New Scripting.Dictionary
                        secReconductored("distance") = CInt(matches(0).SubMatches(0))
                        secReconductored("size1") = Replace(matches(0).SubMatches(1), " ", "")
                        secReconductored("size2") = Replace(matches(0).SubMatches(2), " ", "")
                        For Each Span In pole.spans
                            If Round(Span.distance, 0) = secReconductored("distance") Then
                                correctSize = False
                                For Each wire In Span.utilWires
                                    If wire.componentType = "SEC" And Replace(wire.size, " ", "") = secReconductored("size1") Then correctSize = True: Exit For
                                Next wire
                                If correctSize Then
                                    lengthCount = 0
                                    For Each otherSpan In pole.spans
                                        If Round(otherSpan.distance, 0) = secReconductored("distance") Then lengthCount = lengthCount + 1
                                    Next otherSpan
                                
                                    Set otherPole = Utilities.getPole(Span.otherPole)
                                    If Not otherPole Is Nothing Then
                                        If lengthCount = 1 Or Left(Span.direction, 1) = Left(matches(0).SubMatches(3), 1) Then
                                            secReconductored("otherPole") = Span.otherPole
                                            
                                            If Not secReconductoringDict.exists(pole.poleNumber) Then Set secReconductoringDict(pole.poleNumber) = New Scripting.Dictionary
                                            If Not secReconductoringDict(pole.poleNumber).exists(Span.spanSlot) Then Set secReconductoringDict(pole.poleNumber)(Span.spanSlot) = New Collection
                                            secReconductoringDict(pole.poleNumber)(Span.spanSlot).Add secReconductored
                                            For Each otherSpan In otherPole.spans
                                                If pole.poleNumber = otherSpan.otherPole Then
                                                    If Not secReconductoringDict.exists(otherPole.poleNumber) Then Set secReconductoringDict(otherPole.poleNumber) = New Scripting.Dictionary
                                                    If Not secReconductoringDict(otherPole.poleNumber).exists(otherSpan.spanSlot) Then Set secReconductoringDict(otherPole.poleNumber)(otherSpan.spanSlot) = New Collection
                                                    secReconductoringDict(otherPole.poleNumber)(otherSpan.spanSlot).Add secReconductored
                                                    Exit For
                                                End If
                                            Next otherSpan
                                        End If
                                    End If
                                End If
                            End If
                        Next Span
                    End If
                End If
            Next line
        End If
    Next pole
    
    
    Dim owRegex As Object
    Set owRegex = CreateObject("VBScript.RegExp")
    
    owRegex.Pattern = "\s*(\d*)'[ OF]*\s*(\d[-\d]*)\s*O[PEN]*\s*W[IRE]*\s*\/\s*(.*)SEC\s*(.*)"
    owRegex.Global = True
    owRegex.IgnoreCase = True
    
    Dim owReconductoringDict As Scripting.Dictionary
    Set owReconductoringDict = New Scripting.Dictionary
    
    For Each pole In project.poles
        If LoadingBar_Form2.gTotal = 0 Then Exit Function
        If pole.location <> "" Then
            crewNotes = pole.alt1
            Call Utilities.applyStandardAbbreviations(crewNotes)
            lines = Split(crewNotes, vbLf)
            For Each line In lines
                If InStr(line, "SEC") > 0 Then
                    If owRegex.test(line) Then
                        Set matches = owRegex.Execute(line)
                        Set owReconductored = New Scripting.Dictionary
                        owReconductored("distance") = CInt(matches(0).SubMatches(0))
                        owReconductored("size1") = Trim(matches(0).SubMatches(1))
                        owReconductored("size2") = Trim(matches(0).SubMatches(2))
                        For Each Span In pole.spans
                            If Round(Span.distance, 0) = owReconductored("distance") Then
                                correctSize = False
                                For Each wire In Span.utilWires
                                    If wire.componentType = "OW" Then correctSize = True: Exit For
                                Next wire
                                If correctSize Then
                                    lengthCount = 0
                                    For Each otherSpan In pole.spans
                                        If Round(otherSpan.distance, 0) = owReconductored("distance") Then lengthCount = lengthCount + 1
                                    Next otherSpan
                                
                                    Set otherPole = Utilities.getPole(Span.otherPole)
                                    If Not otherPole Is Nothing Then
                                        If lengthCount = 1 Or Left(Span.direction, 1) = Left(matches(0).SubMatches(3), 1) Then
                                            owReconductored("otherPole") = Span.otherPole
                                            
                                            If Not owReconductoringDict.exists(pole.poleNumber) Then Set owReconductoringDict(pole.poleNumber) = New Scripting.Dictionary
                                            If Not owReconductoringDict(pole.poleNumber).exists(Span.spanSlot) Then Set owReconductoringDict(pole.poleNumber)(Span.spanSlot) = New Collection
                                            owReconductoringDict(pole.poleNumber)(Span.spanSlot).Add owReconductored
                                            For Each otherSpan In otherPole.spans
                                                If pole.poleNumber = otherSpan.otherPole Then
                                                    If Not owReconductoringDict.exists(otherPole.poleNumber) Then Set owReconductoringDict(otherPole.poleNumber) = New Scripting.Dictionary
                                                    If Not owReconductoringDict(otherPole.poleNumber).exists(otherSpan.spanSlot) Then Set owReconductoringDict(otherPole.poleNumber)(otherSpan.spanSlot) = New Collection
                                                    owReconductoringDict(otherPole.poleNumber)(otherSpan.spanSlot).Add owReconductored
                                                    Exit For
                                                End If
                                            Next otherSpan
                                        End If
                                    End If
                                End If
                            End If
                        Next Span
                    End If
                End If
            Next line
        End If
    Next pole
    
    Set json("groups") = New Scripting.Dictionary
    
    For Each poleCollection In poleCollections
        If LoadingBar_Form2.gTotal = 0 Then Exit Function
        lowestLatitude = 0
        lowestLongitude = 0
        highestLatitude = 0
        highestLongitude = 0
        For i = 1 To poleCollection.count
            If LoadingBar_Form2.gTotal = 0 Then Exit Function
            Set pole = poleCollection(i)
            If lowestLatitude = 0 Or pole.latitude < lowestLatitude Then lowestLatitude = pole.latitude
            If lowestLongitude = 0 Or pole.longitude < lowestLongitude Then lowestLongitude = pole.longitude
            If highestLatitude = 0 Or pole.latitude > highestLatitude Then highestLatitude = pole.latitude
            If highestLongitude = 0 Or pole.longitude > highestLongitude Then highestLongitude = pole.longitude
            pole.groupNumber = groupNumber
        Next i
        
        Dim jsonGroup As Scripting.Dictionary: Set jsonGroup = New Scripting.Dictionary
        results = LatLonToMI2253(lowestLatitude, lowestLongitude)
        jsonGroup("minX") = results(0)
        jsonGroup("minY") = results(1)
        
        results = LatLonToMI2253(highestLatitude, highestLongitude)
        jsonGroup("maxX") = results(0)
        jsonGroup("maxY") = results(1)
        
        Dim mergeGroup As Boolean: mergeGroup = False
        For Each group In json("groups")
            If LoadingBar_Form2.gTotal = 0 Then Exit Function
            If json("groups")(group)("maxX") < jsonGroup("minX") Then
                dx = jsonGroup("minX") - json("groups")(group)("maxX")
            ElseIf jsonGroup("maxX") < json("groups")(group)("minX") Then
                dx = json("groups")(group)("minX") - jsonGroup("maxX")
            Else
                dx = 0
            End If
            
            If json("groups")(group)("maxY") < jsonGroup("minY") Then
                dy = jsonGroup("minY") - json("groups")(group)("maxY")
            ElseIf jsonGroup("maxY") < json("groups")(group)("minY") Then
                dy = json("groups")(group)("minY") - jsonGroup("maxY")
            Else
                dy = 0
            End If
            
            If (dx * dx + dy * dy) <= 500 ^ 2 Then
                mergeGroup = True
                If json("groups")(group)("maxX") < jsonGroup("maxX") Then json("groups")(group)("maxX") = jsonGroup("maxX")
                If json("groups")(group)("maxY") < jsonGroup("maxY") Then json("groups")(group)("maxY") = jsonGroup("maxY")
                If json("groups")(group)("minX") > jsonGroup("minX") Then json("groups")(group)("minX") = jsonGroup("minX")
                If json("groups")(group)("minY") > jsonGroup("minY") Then json("groups")(group)("minY") = jsonGroup("minY")
                For i = 1 To poleCollection.count
                    Set pole = poleCollection(i)
                    pole.groupNumber = group
                Next i
            End If
        Next group
        
        If Not mergeGroup Then
            Set json("groups")(groupNumber) = jsonGroup
            groupNumber = groupNumber + 1
        End If
    Next poleCollection
    
    Set json("poles") = New Collection
    Set json("streetlights") = New Collection
    Set json("spanguys") = New Collection
    Set json("openWires") = New Collection
    Set json("secWires") = New Collection
    Set json("priWires") = New Collection
    Set json("services") = New Collection
    
    Call LoadingBar_Form2.InitProgress(poleCollections.count)
    
    groupNumber = 1
    For Each poleCollection In poleCollections
        If LoadingBar_Form2.gTotal = 0 Then Exit Function
        Call LoadingBar_Form2.UpdateProgress("Generating Print JSON", "Generating Pole Information for JSON", True)
    
        Set jsonPoles = getElectricJson(poleCollection, 3, token)
        Dim priLayers As Scripting.Dictionary: Set priLayers = New Scripting.Dictionary
        For Each pole In poleCollection
            If LoadingBar_Form2.gTotal = 0 Then Exit Function
            Set jsonPole = New Scripting.Dictionary
            jsonPole("ceid") = pole.existingCEID
            results = LatLonToMI2253(pole.latitude, pole.longitude)
            jsonPole("x") = pole.x
            jsonPole("y") = pole.y
            jsonPole("hvd") = ""
            jsonPole("height") = pole.height / 12
            jsonPole("class") = pole.Class
            jsonPole("crewNotes") = pole.alt1
            jsonPole("location") = pole.location
            jsonPole("replace") = pole.ReplacePole
            jsonPole("newHeight") = pole.newHeight / 12
            jsonPole("newClass") = pole.newClass
            jsonPole("tree") = pole.treeWork
            jsonPole("topped") = pole.topped
            jsonPole("group") = pole.groupNumber
            jsonPole("skipSpan") = pole.skipSpan
            
            For Each jsonFeature In jsonPoles("features")
                If pole.existingCEID = jsonFeature("attributes")("CE_TAG") Or pole.gisCEID = jsonFeature("attributes")("CE_TAG") Then
                    If Not IsNull(jsonFeature("attributes")("HVD_TAG")) Then jsonPole("hvd") = jsonFeature("attributes")("HVD_TAG")
                End If
            Next jsonFeature
            
            If pole.locationAdjacent Then
                For Each Service In pole.services
                    For Each midspan In Service.midspans
                        Set jsonService = New Scripting.Dictionary
                        Set Span = pole.spans(midspan)
                        jsonService("adjacent") = False
                        jsonService("x") = pole.x
                        jsonService("y") = pole.y
                        jsonService("ug") = False
                        jsonService("distance") = Span.distance
                        jsonService("angle") = Span.angle
                        jsonService("address") = Span.houseNumber
                        json("services").Add jsonService
                    Next midspan
                Next Service
            End If
            
            Set jsonPole("guys") = New Collection
            For Each anchor In pole.anchors
                If anchor.owner = "Consumers Energy" Then
                    Set jsonGuy = New Scripting.Dictionary
                    jsonGuy("angle") = anchor.angle
                    jsonGuy("count") = anchor.ceCount
                    jsonGuy("replace") = pole.ReplacePole
                    jsonPole("guys").Add jsonGuy
                End If
            Next anchor
            
            json("poles").Add jsonPole
            
            If pole.slBottomBracketHeight > 0 Then
                Set jsonStreetlight = New Scripting.Dictionary
                jsonStreetlight("adjacent") = False
                jsonStreetlight("x") = results(0)
                jsonStreetlight("y") = results(1)
                json("streetlights").Add jsonStreetlight
            End If
            
            If pole.locationAdjacent Then
                For Each Span In pole.spans
                    bottomFound = False
                
                    If Utilities.SheetExists(Span.otherPole) Then
                        Set otherPole = project.findPole(Span.otherPole)
                        otherResults = LatLonToMI2253(otherPole.latitude, otherPole.longitude)
                        x2 = otherResults(0)
                        y2 = otherResults(1)
                    ElseIf pole.locationAdjacent Then
                        radAngle = (90 - Span.angle) * (3.14159265358979 / 180)
                        x2 = results(0) + (Span.distance * Cos(radAngle))
                        y2 = results(1) + (Span.distance * Sin(radAngle))
                    End If
                    
                    If pole.location <> "" Or Utilities.SheetExists(Span.otherPole) Then
                        If Not usedPoles.exists(Span.otherPole) Then
                            spgCount = 0
                            Dim layerNumber As Integer: layerNumber = 1
                            Dim openWireDone As Boolean
                            
                            For i = Span.utilWires.count To 1 Step -1
                                Set wire = Span.utilWires(i)
                                If wire.componentType = "SPG" Then
                                    spgCount = spgCount + 1
                                    bottomFound = True
                                End If
                            Next i
                            
                            Dim priCount As Integer: priCount = 1
                            Dim nonPriLayerNumber As Integer: nonPriLayerNumber = 1
                            openWireDone = False
                            For i = Span.utilWires.count To 1 Step -1
                                Set wire = Span.utilWires(i)
                                
                                If wire.componentType = "SEC" Then
                                    Set jsonsecWire = New Scripting.Dictionary
                                    jsonsecWire("length") = Round(Span.distance, 0)
                                    jsonsecWire("angle") = Span.angle
                                    jsonsecWire("layer") = layerNumber
                                    jsonsecWire("size") = Replace(wire.size, " ", "")
                                    jsonsecWire("size2") = ""
                                     
                                    jsonsecWire("startDeadend") = True
                                    For Each otherSpan In pole.spans
                                        If Span.spanSlot <> otherSpan.spanSlot Then
                                            For Each otherWire In otherSpan.utilWires
                                                If otherWire.componentType = "SEC" Then
                                                    If Utilities.OnlyNumbers(otherWire.size) = Utilities.OnlyNumbers(CStr(jsonsecWire("size"))) Then jsonsecWire("startDeadend") = False: Exit For
                                                End If
                                            Next otherWire
                                        End If
                                    Next otherSpan
                                    If Utilities.SheetExists(Span.otherPole) Then
                                        jsonsecWire("endDeadend") = True
                                        For Each otherSpan In otherPole.spans
                                            If otherSpan.otherPole <> pole.poleNumber Then
                                                For Each otherWire In otherSpan.utilWires
                                                    If otherWire.componentType = "SEC" Then
                                                        If Utilities.OnlyNumbers(otherWire.size) = Utilities.OnlyNumbers(CStr(jsonsecWire("size"))) Then jsonsecWire("endDeadend") = False: Exit For
                                                    End If
                                                Next otherWire
                                            End If
                                        Next otherSpan
                                    Else
                                        jsonsecWire("endDeadend") = False
                                    End If
                                    
                                    If secReconductoringDict.exists(pole.poleNumber) Then
                                        If secReconductoringDict(pole.poleNumber).exists(Span.spanSlot) Then
                                            For Each secReconductor In secReconductoringDict(pole.poleNumber)(Span.spanSlot)
                                                If secReconductor("size1") = jsonsecWire("size") And secReconductor("distance") = jsonsecWire("length") Then
                                                    If secReconductoringDict(pole.poleNumber).count = 1 Then jsonsecWire("startDeadend") = True
                                                    If secReconductoringDict(secReconductor("otherPole")).count = 1 Then jsonsecWire("endDeadend") = True
                                                    jsonsecWire("size2") = secReconductor("size2")
                                                    Exit For
                                                End If
                                            Next secReconductor
                                        End If
                                    End If
                                    
                                    jsonsecWire("x1") = pole.x
                                    jsonsecWire("y1") = pole.y
                                    jsonsecWire("x2") = x2
                                    jsonsecWire("y2") = y2
                                    
                                    If Not bottomFound Then
                                        bottomFound = True
                                        jsonsecWire("bottom") = bottomFound
                                    Else
                                        jsonsecWire("bottom") = False
                                    End If
                                    topFound = True
                                    For j = i - 1 To 1 Step -1
                                        Set otherWire = Span.utilWires(j)
                                        If otherWire.componentType = "SEC" Or otherWire.componentType = "OW" Or otherWire.componentType = "PRI" Then topFound = False
                                    Next j
                                    jsonsecWire("top") = topFound
                                     
                                    jsonsecWire("group") = pole.groupNumber
                                    json("secWires").Add jsonsecWire
                                    layerNumber = layerNumber + 1
                                    nonPriLayerNumber = layerNumber
                                End If
                                
                                If wire.componentType = "OW" And Not openWireDone Then
                                    openWireDone = True
                                    Set owWires1 = New Collection
                                    Set owWires2 = New Collection
                                    
                                    For j = i To 1 Step -1
                                        Set otherWire = Span.utilWires(j)
                                        If otherWire.componentType = "OW" Then
                                            If owWires2.count = 0 Then
                                                owWires1.Add otherWire
                                            Else
                                                owWires2.Add otherWire
                                            End If
                                            If owWires1.count > 5 Then
                                                owWires2.Add owWires1(4)
                                                owWires2.Add owWires1(5)
                                                Call owWires1.Remove(5)
                                                Call owWires1.Remove(4)
                                            End If
                                        End If
                                    Next j
                                    
                                    Set owWiresCollections = New Collection
                                    owWiresCollections.Add owWires1
                                    owWiresCollections.Add owWires2
                                    
                                    For Each owWires In owWiresCollections
                                        If owWires.count > 0 Then
                                            Set jsonOpenwire = New Scripting.Dictionary
                                            jsonOpenwire("length") = Round(Span.distance, 0)
                                            jsonOpenwire("angle") = Span.angle
                                            jsonOpenwire("layer") = layerNumber
                                            
                                            openWireString = ""
                                            For Each owWire In owWires
                                                openWireString = openWireString & Utilities.OnlyNumbers(ThisWorkbook.RemoveParentheses(owWire.size)) & "-"
                                            Next owWire
                                            
                                            If openWireString <> "" Then
                                                openWireString = Left(openWireString, Len(openWireString) - 1)
                                                jsonOpenwire("size") = openWireString
                                                jsonOpenwire("size2") = ""
                                                If owWires.count = 3 And Len(openWireString) = 5 Then
                                                    char1 = Mid(openWireString, 1, 1)
                                                    char2 = Mid(openWireString, 3, 1)
                                                    char3 = Mid(openWireString, 5, 1)
                                                    
                                                    If char1 <> char3 Then
                                                        If char1 <> char2 Then
                                                            Mid(openWireString, 1, 1) = char2
                                                            Mid(openWireString, 3, 1) = char1
                                                        ElseIf char3 <> char2 Then
                                                            Mid(openWireString, 5, 1) = char2
                                                            Mid(openWireString, 3, 1) = char3
                                                        End If
                                                    End If
                                                End If
                                                
                                                jsonOpenwire("startDeadend") = True
                                                For Each otherSpan In pole.spans
                                                    If Span.spanSlot <> otherSpan.spanSlot Then
                                                        For Each otherWire In otherSpan.utilWires
                                                            If otherWire.componentType = "OW" Then
                                                                If InStr(jsonOpenwire("size"), Utilities.OnlyNumbers(otherWire.size)) > 0 Then jsonOpenwire("startDeadend") = False: Exit For
                                                            End If
                                                        Next otherWire
                                                    End If
                                                Next otherSpan
                                                If Utilities.SheetExists(Span.otherPole) Then
                                                    jsonOpenwire("endDeadend") = True
                                                    For Each otherSpan In otherPole.spans
                                                        If otherSpan.otherPole <> pole.poleNumber Then
                                                            For Each otherWire In otherSpan.utilWires
                                                                If otherWire.componentType = "OW" Then
                                                                    If InStr(jsonOpenwire("size"), Utilities.OnlyNumbers(otherWire.size)) > 0 Then jsonOpenwire("endDeadend") = False: Exit For
                                                                End If
                                                            Next otherWire
                                                        End If
                                                    Next otherSpan
                                                Else
                                                    jsonOpenwire("endDeadend") = False
                                                End If
                                            End If
                                            
                                            jsonOpenwire("x1") = results(0)
                                            jsonOpenwire("y1") = results(1)
                                            jsonOpenwire("x2") = x2
                                            jsonOpenwire("y2") = y2
                                            
                                            If owReconductoringDict.exists(pole.poleNumber) Then
                                                If owReconductoringDict(pole.poleNumber).exists(Span.spanSlot) Then
                                                    For Each owReconductor In owReconductoringDict(pole.poleNumber)(Span.spanSlot)
                                                        If owReconductor("size1") = jsonOpenwire("size") And owReconductor("distance") = jsonOpenwire("length") Then
                                                            If owReconductoringDict(pole.poleNumber).count = 1 Then jsonOpenwire("startDeadend") = True
                                                            If owReconductoringDict(owReconductor("otherPole")).count = 1 Then jsonOpenwire("endDeadend") = True
                                                            jsonOpenwire("size2") = Replace(owReconductor("size2"), " ", "")
                                                            Exit For
                                                        End If
                                                    Next owReconductor
                                                End If
                                            End If
                                    
                                            If Not bottomFound Then
                                                bottomFound = True
                                                jsonOpenwire("bottom") = bottomFound
                                            Else
                                                jsonOpenwire("bottom") = False
                                            End If
                                            topFound = True
                                            For j = i - 1 To 1 Step -1
                                                Set otherWire = Span.utilWires(j)
                                                If otherWire.componentType = "SEC" Or otherWire.componentType = "PRI" Then topFound = False
                                            Next j
                                            jsonOpenwire("top") = topFound
                                            
                                            jsonOpenwire("group") = pole.groupNumber
                                            json("openWires").Add jsonOpenwire
                                            layerNumber = layerNumber + 1
                                            nonPriLayerNumber = layerNumber
                                        End If
                                    Next owWires
                                End If
                                
                                If wire.componentType = "PRI" Then
                                    Set jsonPriWire = New Scripting.Dictionary
                                    
                                    jsonPriWire("size") = Utilities.OnlyNumbers(wire.size, True)
                                    jsonPriWire("size2") = ""
                                    jsonPriWire("phase") = wire.phase
                                    If Not priLayers.exists(priCount) Then priLayers(priCount) = 1
                                    If nonPriLayerNumber = layerNumber Then layerNumber = layerNumber + 1
                                    If priLayers(priCount) < layerNumber Then priLayers(priCount) = layerNumber
                                    
                                    jsonPriWire("startDeadend") = True
                                    For Each otherSpan In pole.spans
                                        If Span.spanSlot <> otherSpan.spanSlot Then
                                            For Each otherWire In otherSpan.utilWires
                                                If otherWire.componentType = "PRI" Then
                                                    If Utilities.OnlyNumbers(otherWire.size, True) = jsonPriWire("size") Then
                                                        jsonPriWire("startDeadend") = False: Exit For
                                                    End If
                                                End If
                                            Next otherWire
                                        End If
                                    Next otherSpan
                                    If Utilities.SheetExists(Span.otherPole) Then
                                        jsonPriWire("endDeadend") = True
                                        For Each otherSpan In otherPole.spans
                                            If otherSpan.otherPole <> pole.poleNumber Then
                                                For Each otherWire In otherSpan.utilWires
                                                    If otherWire.componentType = "PRI" Then
                                                        If Utilities.OnlyNumbers(otherWire.size, True) = jsonPriWire("size") Then jsonPriWire("endDeadend") = False: Exit For
                                                    End If
                                                Next otherWire
                                            End If
                                        Next otherSpan
                                    Else
                                        jsonPriWire("endDeadend") = False
                                    End If
                                    
                                    jsonPriWire("length") = Round(Span.distance, 0)
                                    jsonPriWire("angle") = Span.angle
                                    jsonPriWire("layer") = layerNumber
                                    jsonPriWire("priLayer") = priCount
                                    
                                    Dim config As String: config = ""
                                    For Each otherWire In Span.utilWires
                                        If otherWire.componentType = "NEUT" Then
                                            jsonPriWire("neutSize") = Utilities.OnlyNumbers(otherWire.size, True)
                                            If wire.height - otherWire.height < 18 Then
                                                config = "N"
                                            Else
                                                config = "NB"
                                            End If
                                            Exit For
                                        End If
                                        If otherWire.componentType = "OW" Or otherWire.componentType = "SEC" Then config = "SN"
                                    Next otherWire
                                    
                                    
                                    If config = "" Then
                                        For Each otherSpan In pole.spans
                                            If Span.spanSlot <> otherSpan.spanSlot Then
                                                angleDif = Abs(Span.angle - otherSpan.angle)
                                                If angleDif >= 315 Then angleDif = 360 - angleDif
                                                If angleDif < 45 Then
                                                    For Each otherWire In otherSpan.utilWires
                                                        If otherWire.componentType = "NEUT" Then
                                                            jsonPriWire("neutSize") = Utilities.OnlyNumbers(otherWire.size, True)
                                                            If wire.height - otherWire.height < 18 Then
                                                                config = "N"
                                                            Else
                                                                config = "NB"
                                                            End If
                                                            Exit For
                                                        End If
                                                        If otherWire.componentType = "OW" Or otherWire.componentType = "SEC" Then config = "SN"
                                                    Next otherWire
                                                End If
                                            End If
                                        Next otherSpan
                                    End If
                                    
                                    jsonPriWire("neutSize2") = ""
                                    If primaryReconductoringDict.exists(pole.poleNumber) Then
                                        If primaryReconductoringDict(pole.poleNumber).exists(Span.spanSlot) Then
                                            For Each priReconductor In primaryReconductoringDict(pole.poleNumber)(Span.spanSlot)
                                                If priReconductor("size1") = jsonPriWire("size") And priReconductor("distance") = jsonPriWire("length") Then
                                                    If primaryReconductoringDict(pole.poleNumber).count = 1 Then jsonPriWire("startDeadend") = True
                                                    If primaryReconductoringDict(priReconductor("otherPole")).count = 1 Then jsonPriWire("endDeadend") = True
                                                    jsonPriWire("size2") = Utilities.OnlyNumbers(CStr(priReconductor("size2")), True)
                                                    If priReconductor.exists("neutSize") Then jsonPriWire("neutSize2") = Utilities.OnlyNumbers(CStr(priReconductor("neutSize")), True)
                                                    Exit For
                                                End If
                                            Next priReconductor
                                        End If
                                    End If
                                    
                                    jsonPriWire("configuration") = config
                                    
                                    jsonPriWire("x1") = results(0)
                                    jsonPriWire("y1") = results(1)
                                    jsonPriWire("x2") = x2
                                    jsonPriWire("y2") = y2
                                    
                                    If Not bottomFound Then
                                        bottomFound = True
                                        jsonPriWire("bottom") = bottomFound
                                    Else
                                        jsonPriWire("bottom") = False
                                    End If
                                    topFound = True
                                    For j = i - 1 To 1 Step -1
                                        Set otherWire = Span.utilWires(j)
                                        If otherWire.componentType = "SEC" Or otherWire.componentType = "OW" Or otherWire.componentType = "PRI" Then topFound = False
                                    Next j
                                    jsonPriWire("top") = topFound
                                    
                                    For Each otherSpan In pole.spans
                                        If Span.spanSlot <> otherSpan.spanSlot Then
                                            If Abs(Span.angle - otherSpan.angle) < 15 Then
                                                If otherSpan.otherPole <> "" Then
                                                    If otherSpan.utilWires.count > 0 Then
                                                        If project.findPole(otherSpan.otherPole).skipSpan Then
                                                            jsonPriWire("length") = ""
                                                            Exit For
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Next otherSpan
                                    
                                    
                                    jsonPriWire("group") = pole.groupNumber
                                    json("priWires").Add jsonPriWire
                                    layerNumber = layerNumber + 1
                                    priCount = priCount + 1
                                    End If
                            Next i
                            
                            
                            If spgCount > 0 Then
                                Set jsonSpanguy = New Scripting.Dictionary
                                jsonSpanguy("length") = Round(Span.distance, 0)
                                jsonSpanguy("count") = spgCount
                                jsonSpanguy("angle") = Span.angle
                                
                                jsonSpanguy("x1") = results(0)
                                jsonSpanguy("y1") = results(1)
                                jsonSpanguy("x2") = x2
                                jsonSpanguy("y2") = y2
                                jsonSpanguy("top") = spgCount = Span.utilWires.count
                                
                                json("spanguys").Add jsonSpanguy
                            End If
                        End If
                    End If
                Next Span
                
                usedPoles.Add pole.poleNumber, Nothing
            End If
        Next pole
        
        For Each jsonPriWire In json("priWires")
            If jsonPriWire("group") = groupNumber Then jsonPriWire("layer") = priLayers(jsonPriWire("priLayer"))
        Next jsonPriWire
        groupNumber = groupNumber + 1
    Next poleCollection
    
    Call LoadingBar_Form2.InitProgress(6 * poleCollections.count)
    
    Dim layers As Collection: Set layers = New Collection
    layers.Add 7  ' Major Roads
    layers.Add 8  ' Minor Roads
    layers.Add 9  ' Pavement Lines
    layers.Add 10 ' Private Drive
    layers.Add 11 ' Raildroad
    layers.Add 15 ' Waterway
    
    Set json("roads") = New Collection
    For Each layer In layers
        If LoadingBar_Form2.gTotal = 0 Then Exit Function
        For Each poleCollection In poleCollections
            If LoadingBar_Form2.gTotal = 0 Then Exit Function
            Call LoadingBar_Form2.UpdateProgress("Generating Print JSON", "Extracting and Generating ROW Information for JSON", True)
             
            ' Get ROW
            Set mapjson = getROWJSON(poleCollection, CInt(layer), token)
            If Not mapjson Is Nothing Then
                If mapjson.exists("features") Then
                    For Each jsonFeature In mapjson("features")
                        If jsonFeature.exists("geometry") Then
                            If jsonFeature("geometry").exists("paths") Then
                                For Each jsonPath In jsonFeature("geometry")("paths")
                                    Set path = New Collection
                                    For Each jsonLine In jsonPath
                                        Set line = New Collection
                                        line.Add jsonLine(1)
                                        line.Add jsonLine(2)
                                        path.Add line
                                    Next jsonLine
                                    Dim exists As Boolean
                                    For Each road In json("roads")
                                        exists = False
                                        If path.count = road.count And path.count > 0 Then
                                            exists = True
                                            For j = 1 To road.count
                                                If path(j)(1) <> road(j)(1) Then
                                                    exists = False
                                                    Exit For
                                                End If
                                                If path(j)(2) <> road(j)(2) Then
                                                    exists = False
                                                    Exit For
                                                End If
                                            Next j
                                            If exists Then Exit For
                                        End If
                                    Next road
                                    If Not exists Then json("roads").Add path
                                Next jsonPath
                            End If
                        End If
                    Next jsonFeature
                End If
            End If
            ' End ROW
            
        Next poleCollection
    Next layer
    
    Call LoadingBar_Form2.InitProgress(poleCollections.count)
    
    Set json("centerlines") = New Collection
    For Each poleCollection In poleCollections
        If LoadingBar_Form2.gTotal = 0 Then Exit Function
        Call LoadingBar_Form2.UpdateProgress("Generating Print JSON", "Extracting and Generating Centerline Information for JSON", True)
        ' Get Centerlines
        Set mapjson = getROWJSON(poleCollection, 13, token)
        If Not mapjson Is Nothing Then
            If mapjson.exists("features") Then
                For Each jsonFeature In mapjson("features")
                    If jsonFeature.exists("geometry") Then
                        If jsonFeature("geometry").exists("paths") Then
                            For Each jsonPath In jsonFeature("geometry")("paths")
                                Set path = New Collection
                                For Each jsonLine In jsonPath
                                    Set line = New Collection
                                    line.Add jsonLine(1)
                                    line.Add jsonLine(2)
                                    path.Add line
                                Next jsonLine
                                exists = False
                                For Each centerline In json("centerlines")
                                    exists = False
                                    If path.count = centerline.count And path.count > 0 Then
                                        exists = True
                                        For j = 1 To centerline.count
                                            If path(j)(1) <> centerline(j)(1) Then
                                                exists = False
                                                Exit For
                                            End If
                                            If path(j)(2) <> centerline(j)(2) Then
                                                exists = False
                                                Exit For
                                            End If
                                        Next j
                                        If exists Then Exit For
                                    End If
                                Next centerline
                                If Not exists Then json("centerlines").Add path
                            Next jsonPath
                        End If
                    End If
                Next jsonFeature
            End If
        End If
        ' End Centerlines
            
    Next poleCollection
    
    Call LoadingBar_Form2.InitProgress(project.poles.count)
    For Each poleCollection In poleCollections
        If LoadingBar_Form2.gTotal = 0 Then Exit Function
        For Each pole In poleCollection
            Call LoadingBar_Form2.UpdateProgress("Generating Print JSON", "Extracting and Generating Adjacent Pole Information for JSON", True)
            If LoadingBar_Form2.gTotal = 0 Then Exit Function
            ' Get Other Poles
            If pole.location <> "" Then
                For Each Span In pole.spans
                    Dim notService As Boolean: notService = False
                    For Each wire In Span.utilWires
                        If wire.componentType <> "SVC" Then notService = True
                    Next wire
                    If notService And Span.otherPole = "" Then
                        Set otherPole = New pole
                        otherPole.latitude = pole.latitude
                        otherPole.longitude = pole.longitude
                        otherPole.transformerSizes = -1
                        otherPole.slBottomBracketHeight = -1
                        
                    
                        Set jsonPole = New Scripting.Dictionary
                        radAngle = (90 - Span.angle) * (3.14159265358979 / 180)
                        x = pole.x + (Span.distance * Cos(radAngle))
                        y = pole.y + (Span.distance * Sin(radAngle))
                        otherPole.x = x
                        otherPole.y = y
                        jsonPole("x") = otherPole.x
                        jsonPole("y") = otherPole.y
                        
                        Set jsonPoles = getOtherPole(x, y, 20, token)
                        
                        Dim closestFeature As Object
                        Dim closestDistance As Double
                        
                        results = getClosestPole(jsonPoles, x, y)
                        Set closestFeature = results(0)
                        closestDistance = results(1)
                        
                        Dim distance As Integer: distance = 30
                        Do While closestFeature Is Nothing
                            Set jsonPoles = getOtherPole(x, y, distance, token)
                            If jsonPoles("features").count = 1 Then
                                results = getClosestPole(jsonPoles, x, y)
                                Set closestFeature = results(0)
                                closestDistance = results(1)
                                Exit Do
                            End If
                            distance = distance + 10
                            If distance > 40 Then Exit Do
                        Loop
                        
                        If Not closestFeature Is Nothing Then
                            If closestFeature("attributes")("CE_TAG") <> pole.existingCEID And closestFeature("attributes")("CE_TAG") <> pole.gisCEID Then jsonPole("ceid") = closestFeature("attributes")("CE_TAG")
                            If closestFeature("attributes")("OWNER") <> "Consumers Energy" Then jsonPole("ceid") = "FOREIGN"
                            jsonPole("height") = closestFeature("attributes")("HEIGHT")
                            jsonPole("class") = closestFeature("attributes")("CLASS")
                            jsonPole("hvd") = closestFeature("attributes")("HVD_TAG")
                            jsonPole("x2") = closestFeature("geometry")("x")
                            jsonPole("y2") = closestFeature("geometry")("y")
                        End If
                        
                        Dim duplicatePole As Boolean: duplicatePole = False
                        For Each jsonPole2 In json("poles")
                            If jsonPole("x2") <> "" And ((jsonPole2("x") = jsonPole("x") And jsonPole2("y") = jsonPole("y")) Or (jsonPole2("x2") = jsonPole("x2") And jsonPole2("y2") = jsonPole("y2"))) Then
                                duplicatePole = True
                                Exit For
                            End If
                            If Utilities.isCEID(jsonPole("ceid")) And jsonPole("ceid") = jsonPole2("ceid") Then
                                duplicatePole = True
                                Exit For
                            End If
                        Next jsonPole2
                        If Not duplicatePole Then
                            json("poles").Add jsonPole
                            otherPole.existingCEID = jsonPole("ceid")
                            otherPole.gisCEID = jsonPole("ceid")
                            otherPole.poleNumber = "NEWGISPOLE"
                            poleCollection.Add otherPole
                        End If
                    End If
                Next Span
            End If
        Next pole
    Next poleCollection
    
    Dim phases As Scripting.Dictionary: Set phases = New Scripting.Dictionary
    phases(1) = "Z"
    phases(2) = "Y"
    phases(3) = "YZ"
    phases(4) = "X"
    phases(5) = "XZ"
    phases(6) = "XY"
    phases(7) = "3P"
    
    Dim lowSideVoltages As Scripting.Dictionary: Set lowSideVoltages = New Scripting.Dictionary
    lowSideVoltages(10) = "120"
    lowSideVoltages(20) = "120/208"
    lowSideVoltages(30) = "120/240"
    lowSideVoltages(40) = "240"
    lowSideVoltages(45) = "240/480"
    lowSideVoltages(50) = "277/480"
    lowSideVoltages(60) = "480"
    lowSideVoltages(70) = "600"
    
    Set json("transformers") = New Collection
    Set json("capacitors") = New Collection
    Set json("regulators") = New Collection
    Set json("isolators") = New Collection
    
    Set json("fuses") = New Collection
    Set json("reclosers") = New Collection
    Set json("sectionalizers") = New Collection
    Set json("switches") = New Collection
    Set json("sensors") = New Collection
    
    Set json("risers") = New Collection
    
    Call LoadingBar_Form2.InitProgress(poleCollections.count * 2)
    
    For Each poleCollection In poleCollections
        If LoadingBar_Form2.gTotal = 0 Then Exit Function
        Call LoadingBar_Form2.UpdateProgress("Generating Print JSON", "Extracting Equipment GIS Information", True)
        Set jsonPoles = getElectricJson(poleCollection, 3, token)
        If LoadingBar_Form2.gTotal = 0 Then Exit Function
        
        Set jsonFuses = getElectricJson(poleCollection, 60, token)
        If LoadingBar_Form2.gTotal = 0 Then Exit Function
        Set jsonReclosers = getElectricJson(poleCollection, 70, token)
        If LoadingBar_Form2.gTotal = 0 Then Exit Function
        Set jsonSectionalizers = getElectricJson(poleCollection, 71, token)
        If LoadingBar_Form2.gTotal = 0 Then Exit Function
        Set jsonSwitches = getElectricJson(poleCollection, 65, token)
        If LoadingBar_Form2.gTotal = 0 Then Exit Function
        Set jsonSensors = getElectricJson(poleCollection, 13, token, "SENSORMODEL+IS+NOT+NULL")
        If LoadingBar_Form2.gTotal = 0 Then Exit Function
        Set jsonPriRisers = getElectricJson(poleCollection, 4, token)
        If LoadingBar_Form2.gTotal = 0 Then Exit Function
        Set jsonSecRisers = getElectricJson(poleCollection, 5, token)
        If LoadingBar_Form2.gTotal = 0 Then Exit Function
        
        Set jsonSecOH = getElectricJson(poleCollection, 32, token)
        If LoadingBar_Form2.gTotal = 0 Then Exit Function
        Set jsonSecUG = getElectricJson(poleCollection, 33, token)
        If LoadingBar_Form2.gTotal = 0 Then Exit Function
        Set jsonServicePoints = getElectricJson(poleCollection, 26, token)
        If LoadingBar_Form2.gTotal = 0 Then Exit Function
        
        Set jsonTransformerConnections = getElectricJson(poleCollection, 30, token)
        If LoadingBar_Form2.gTotal = 0 Then Exit Function
        Set jsonTransformers = getElectricJson(poleCollection, 27, token)
        If LoadingBar_Form2.gTotal = 0 Then Exit Function
        Set jsonStreetlights = getElectricJson(poleCollection, 2, token)
        If LoadingBar_Form2.gTotal = 0 Then Exit Function
        Set jsonCapacitors = getElectricJson(poleCollection, 87, token)
        If LoadingBar_Form2.gTotal = 0 Then Exit Function
        Set jsonRegulators = getElectricJson(poleCollection, 77, token)
        If LoadingBar_Form2.gTotal = 0 Then Exit Function
        Set jsonIsolators = getElectricJson(poleCollection, 82, token)
        If LoadingBar_Form2.gTotal = 0 Then Exit Function
        
        Call LoadingBar_Form2.UpdateProgress("Generating Print JSON", "Generating Equipment Information for JSON", True)
        For Each jsonFeature In jsonFuses("features")
            If LoadingBar_Form2.gTotal = 0 Then Exit Function
            x = jsonFeature("geometry")("x")
            y = jsonFeature("geometry")("y")
            Set closestPole = isClosestPoleOnJob(jsonPoles, poleCollection, x, y)
            If Not closestPole Is Nothing Then
                Set jsonFuse = New Scripting.Dictionary
                
                If closestPole.poleNumber = "NEWGISPOLE" Then
                    jsonFuse("adjacent") = True
                Else
                    jsonFuse("adjacent") = False
                End If
                
                jsonFuse("x") = x
                jsonFuse("y") = y
                jsonFuse("lcp") = jsonFeature("attributes")("LCP")
                jsonFuse("size") = jsonFeature("attributes")("RATEDCURRENT")
                jsonFuse("open") = jsonFeature("attributes")("SWITCHSTATUS") = 0
                jsonFuse("rotation") = jsonFeature("attributes")("SYMBOLROTATION")
                
                If jsonFuse("lcp") <> "" Then json("fuses").Add jsonFuse
            End If
        Next jsonFeature
        
        For Each jsonFeature In jsonReclosers("features")
            If LoadingBar_Form2.gTotal = 0 Then Exit Function
            x = jsonFeature("geometry")("x")
            y = jsonFeature("geometry")("y")
            
            Set closestPole = isClosestPoleOnJob(jsonPoles, poleCollection, x, y)
            If Not closestPole Is Nothing Then
                Set jsonRecloser = New Scripting.Dictionary
                
                If closestPole.poleNumber = "NEWGISPOLE" Then
                    jsonRecloser("adjacent") = True
                Else
                    jsonRecloser("adjacent") = False
                End If
                
                jsonRecloser("x") = x
                jsonRecloser("y") = y
                jsonRecloser("lcp") = jsonFeature("attributes")("LCP")
                jsonRecloser("size") = jsonFeature("attributes")("RATEDCURRENT")
                jsonRecloser("open") = jsonFeature("attributes")("SWITCHSTATUS") = 0
                jsonRecloser("rotation") = jsonFeature("attributes")("SYMBOLROTATION")
                
                If jsonRecloser("lcp") <> "" Then json("reclosers").Add jsonRecloser
            End If
        Next jsonFeature
    
        For Each jsonFeature In jsonSectionalizers("features")
            If LoadingBar_Form2.gTotal = 0 Then Exit Function
            x = jsonFeature("geometry")("x")
            y = jsonFeature("geometry")("y")
            
            Set closestPole = isClosestPoleOnJob(jsonPoles, poleCollection, x, y)
            If Not closestPole Is Nothing Then
                Set jsonSectionalizer = New Scripting.Dictionary
                
                If closestPole.poleNumber = "NEWGISPOLE" Then
                    jsonSectionalizer("adjacent") = True
                Else
                    jsonSectionalizer("adjacent") = False
                End If
                
                jsonSectionalizer("x") = x
                jsonSectionalizer("y") = y
                jsonSectionalizer("lcp") = jsonFeature("attributes")("LCP")
                jsonSectionalizer("size") = jsonFeature("attributes")("RATEDCURRENT")
                jsonSectionalizer("open") = jsonFeature("attributes")("SWITCHSTATUS") = 0
                jsonSectionalizer("rotation") = jsonFeature("attributes")("SYMBOLROTATION")
                
                If jsonSectionalizer("lcp") <> "" Then json("sectionalizers").Add jsonSectionalizer
            End If
        Next jsonFeature
        
        For Each jsonFeature In jsonSwitches("features")
            If LoadingBar_Form2.gTotal = 0 Then Exit Function
            x = jsonFeature("geometry")("x")
            y = jsonFeature("geometry")("y")
            
            Set closestPole = isClosestPoleOnJob(jsonPoles, poleCollection, x, y)
            If Not closestPole Is Nothing Then
                Set jsonSwitch = New Scripting.Dictionary
                
                If closestPole.poleNumber = "NEWGISPOLE" Then
                    jsonSwitch("adjacent") = True
                Else
                    jsonSwitch("adjacent") = False
                End If
                
                jsonSwitch("x") = x
                jsonSwitch("y") = y
                jsonSwitch("lcp") = jsonFeature("attributes")("LCP")
                jsonSwitch("size") = jsonFeature("attributes")("RATEDCURRENT")
                jsonSwitch("open") = jsonFeature("attributes")("SWITCHSTATUS") = 0
                jsonSwitch("rotation") = jsonFeature("attributes")("SYMBOLROTATION")
                
                If jsonSwitch("lcp") <> "" Then
                    equipmentId = jsonFeature("attributes")("EQUIPMENTID")
                    If equipmentId = "S_BLADE" Or equipmentId = "S_LINK" Or equipmentId = "SB_300A_BLADE" Or equipmentId = "SL_LINK_100A" Or equipmentId = "SL_LINK_200A" Then
                        json("fuses").Add jsonSwitch
                    Else
                        json("switches").Add jsonSwitch
                    End If
                End If
            End If
        Next jsonFeature
        
        For Each jsonFeature In jsonSensors("features")
            If LoadingBar_Form2.gTotal = 0 Then Exit Function
            x = jsonFeature("geometry")("x")
            y = jsonFeature("geometry")("y")

            Set closestPole = isClosestPoleOnJob(jsonPoles, poleCollection, x, y)
            If Not closestPole Is Nothing Then
                Set jsonSensor = New Scripting.Dictionary
                
                If closestPole.poleNumber = "NEWGISPOLE" Then
                    jsonSensor("adjacent") = True
                Else
                    jsonSensor("adjacent") = False
                End If
                
                jsonSensor("x") = x
                jsonSensor("y") = y
                jsonSensor("lcp") = jsonFeature("attributes")("LCP")
                jsonSensor("rotation") = jsonFeature("attributes")("SYMBOLROTATION")
                
                If jsonFeature("attributes")("SUBTYPECD") = "13" Then
                    jsonSensor("power") = True
                    json("sensors").Add jsonSensor
                ElseIf jsonFeature("attributes")("SUBTYPECD") = "14" Then
                    jsonSensor("power") = False
                    json("sensors").Add jsonSensor
                End If
            End If
        Next jsonFeature
        
        For Each jsonFeature In jsonPriRisers("features")
            If LoadingBar_Form2.gTotal = 0 Then Exit Function
            x = jsonFeature("geometry")("x")
            y = jsonFeature("geometry")("y")
            
            Set closestPole = isClosestPoleOnJob(jsonPoles, poleCollection, x, y)
            If Not closestPole Is Nothing Then
                Set jsonRiser = New Scripting.Dictionary
                
                If closestPole.poleNumber = "NEWGISPOLE" Then
                    jsonRiser("adjacent") = True
                Else
                    jsonRiser("adjacent") = False
                End If
                
                jsonRiser("x") = closestPole.x
                jsonRiser("y") = closestPole.y
                jsonRiser("lcp") = jsonFeature("attributes")("LCP")
                jsonRiser("rotation") = jsonFeature("attributes")("SYMBOLROTATION")
                jsonRiser("type") = "Primary"
                
                If jsonRiser("lcp") <> "" Then
                    For i = json("fuses").count To 1 Step -1
                        Set jsonFuse = json("fuses")(i)
                        If jsonFuse("lcp") = jsonRiser("lcp") Then
                            jsonRiser("size") = jsonFuse("size")
                            Call json("fuses").Remove(i)
                        End If
                    Next i
                End If
                
                json("risers").Add jsonRiser
            End If
        Next jsonFeature
        
        For Each jsonFeature In jsonSecRisers("features")
            If LoadingBar_Form2.gTotal = 0 Then Exit Function
            x = jsonFeature("geometry")("x")
            y = jsonFeature("geometry")("y")
            
            Set closestPole = isClosestPoleOnJob(jsonPoles, poleCollection, x, y)
            If Not closestPole Is Nothing Then
                For Each jsonSecUGFeature In jsonSecUG("features")
                    For Each jsonSecUGPath In jsonSecUGFeature("geometry")("paths")
                        If jsonSecUGPath(1)(1) = x And jsonSecUGPath(1)(2) = y Then
                            x2 = jsonSecUGPath(2)(1)
                            y2 = jsonSecUGPath(2)(2)
                            Exit For
                        End If
                        If jsonSecUGPath(2)(1) = x And jsonSecUGPath(2)(2) = y Then
                            x2 = jsonSecUGPath(1)(1)
                            y2 = jsonSecUGPath(1)(2)
                            Exit For
                        End If
                    Next jsonSecUGPath
                Next jsonSecUGFeature
                
                Dim serviceRiser As Boolean: serviceRiser = False
                If x2 And y2 <> 0 Then
                    For Each jsonServicePointFeature In jsonServicePoints("features")
                        If LoadingBar_Form2.gTotal = 0 Then Exit Function
                        If jsonServicePointFeature("geometry")("x") = x2 And jsonServicePointFeature("geometry")("y") = y2 Then
                            If closestPole.poleNumber = "NEWGISPOLE" Or closestPole.locationAdjacent Then
                                Set jsonService = New Scripting.Dictionary
                                
                                If closestPole.poleNumber = "NEWGISPOLE" Then
                                    jsonService("adjacent") = True
                                Else
                                    jsonService("adjacent") = False
                                End If
                                
                                jsonService("x") = closestPole.x
                                jsonService("y") = closestPole.y
                                jsonService("ug") = True
                                
                                dx = x2 - x
                                dy = y2 - y
                                distance = Sqr((dx ^ 2) + (dy ^ 2))
                                
                                PI = 4 * Atn(1)
                                radAngle = Atn2(dy, dx)
                                angle = radAngle * (180 / PI)
                                angle = 90 - angle
                                If angle < 0 Then angle = angle + 360
                                
                                jsonService("distance") = distance
                                jsonService("angle") = angle
                                
                                street = jsonServicePointFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.STREET")
                                If Not IsNull(street) Then
                                    streetnumber = Trim(Left(street, InStr(street, " ")))
                                    jsonService("address") = streetnumber
                                End If
                                
                                json("services").Add jsonService
                            End If
                            serviceRiser = True
                            Exit For
                        End If
                    Next jsonServicePointFeature
                End If
                
                If Not serviceRiser Then
                    Set jsonRiser = New Scripting.Dictionary
                    
                    If closestPole.poleNumber = "NEWGISPOLE" Then
                        jsonRiser("adjacent") = True
                    Else
                        jsonRiser("adjacent") = False
                    End If
                    
                    jsonRiser("x") = closestPole.x
                    jsonRiser("y") = closestPole.y
                    jsonRiser("rotation") = jsonFeature("attributes")("SYMBOLROTATION")
                    jsonRiser("type") = "Secondary"
                    json("risers").Add jsonRiser
                End If
            End If
        Next jsonFeature
        
        For Each jsonServicePointFeature In jsonServicePoints("features")
            If LoadingBar_Form2.gTotal = 0 Then Exit Function
            x = jsonServicePointFeature("geometry")("x")
            y = jsonServicePointFeature("geometry")("y")
            y2 = 0
            x2 = 0
            For Each jsonSecOHFeature In jsonSecOH("features")
                For Each jsonSecOHPath In jsonSecOHFeature("geometry")("paths")
                    If jsonSecOHPath(1)(1) = x And jsonSecOHPath(1)(2) = y Then
                        x2 = jsonSecOHPath(2)(1)
                        y2 = jsonSecOHPath(2)(2)
                        Exit For
                    End If
                    If jsonSecOHPath(2)(1) = x And jsonSecOHPath(2)(2) = y Then
                        x2 = jsonSecOHPath(1)(1)
                        y2 = jsonSecOHPath(1)(2)
                        Exit For
                    End If
                Next jsonSecOHPath
            Next jsonSecOHFeature
            If x2 <> 0 And y2 <> 0 Then
                Set closestPole = isClosestPoleOnJob(jsonPoles, poleCollection, x2, y2)
                If Not closestPole Is Nothing Then
                    If closestPole.poleNumber = "NEWGISPOLE" Then
                        Set jsonService = New Scripting.Dictionary
                        
                        If closestPole.poleNumber = "NEWGISPOLE" Then
                            jsonService("adjacent") = True
                        Else
                            jsonService("adjacent") = False
                        End If
                
                        jsonService("x") = closestPole.x
                        jsonService("y") = closestPole.y
                        jsonService("ug") = False
                        
                        dx = x - x2
                        dy = y - y2
                        distance = Sqr((dx ^ 2) + (dy ^ 2))
                        
                        PI = 4 * Atn(1)
                        radAngle = Atn2(dy, dx)
                        angle = radAngle * (180 / PI)
                        angle = 90 - angle
                        If angle < 0 Then angle = angle + 360
                        
                        jsonService("distance") = distance
                        jsonService("angle") = angle
                         
                        If Not IsNull(jsonServicePointFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.STREET")) Then
                            street = CStr(jsonServicePointFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.STREET"))
                            If InStr(street, " ") > 0 Then
                                streetnumber = Trim(Left(street, InStr(street, " ")))
                            Else
                                streetnumber = ""
                            End If
                        Else
                            streetnumber = ""
                        End If
                        
                        jsonService("address") = streetnumber
                        
                        json("services").Add jsonService
                    End If
                End If
            End If
        Next jsonServicePointFeature
        
        For Each jsonFeature In jsonStreetlights("features")
            If LoadingBar_Form2.gTotal = 0 Then Exit Function
            x = jsonFeature("geometry")("x")
            y = jsonFeature("geometry")("y")

            Set closestPole = isClosestPoleOnJob(jsonPoles, poleCollection, x, y)
            If Not closestPole Is Nothing Then
                If closestPole.poleNumber = "NEWGISPOLE" Or closestPole.slBottomBracketHeight < 1 Then
                    Set jsonStreetlight = New Scripting.Dictionary
                    jsonStreetlight("adjacent") = True
                    
                    jsonStreetlight("x") = closestPole.x
                    jsonStreetlight("y") = closestPole.y
                    json("streetlights").Add jsonStreetlight
                End If
            End If
        Next jsonFeature
        
        For Each jsonFeature In jsonCapacitors("features")
            If LoadingBar_Form2.gTotal = 0 Then Exit Function
            x = jsonFeature("geometry")("x")
            y = jsonFeature("geometry")("y")

            Set closestPole = isClosestPoleOnJob(jsonPoles, poleCollection, x, y)
            If Not closestPole Is Nothing Then
                Set jsonCapacitor = New Scripting.Dictionary
                
                If closestPole.poleNumber = "NEWGISPOLE" Then
                    jsonCapacitor("adjacent") = True
                Else
                    jsonCapacitor("adjacent") = False
                End If
                
                jsonCapacitor("x") = closestPole.x
                jsonCapacitor("y") = closestPole.y
                jsonCapacitor("lcp") = jsonFeature("attributes")("LCP")
                jsonCapacitor("rotation") = jsonFeature("attributes")("SYMBOLROTATION")
                jsonCapacitor("size") = jsonFeature("attributes")("TOTALKVAR")
                
                If jsonFeature("attributes")("SWITCHTYPE") = "Unswitched" Then
                    jsonCapacitor("switched") = False
                Else
                    jsonCapacitor("switched") = True
                End If
                
                If jsonCapacitor("lcp") <> "" Then json("capacitors").Add jsonCapacitor
            End If
        Next jsonFeature
        
        For Each jsonFeature In jsonRegulators("features")
            If LoadingBar_Form2.gTotal = 0 Then Exit Function
            x = jsonFeature("geometry")("x")
            y = jsonFeature("geometry")("y")

            Set closestPole = isClosestPoleOnJob(jsonPoles, poleCollection, x, y)
            If Not closestPole Is Nothing Then
                Set jsonRegulator = New Scripting.Dictionary
                
                If closestPole.poleNumber = "NEWGISPOLE" Then
                    jsonRegulator("adjacent") = True
                Else
                    jsonRegulator("adjacent") = False
                End If
                
                jsonRegulator("x") = closestPole.x
                jsonRegulator("y") = closestPole.y
                jsonRegulator("lcp") = jsonFeature("attributes")("LCP")
                jsonRegulator("rotation") = jsonFeature("attributes")("SYMBOLROTATION")
                jsonRegulator("size") = jsonFeature("attributes")("RATEDKVA")
                jsonRegulator("phase") = phases(jsonFeature("attributes")("PHASEDESIGNATION"))
                
                jsonRegulator("auto") = False
                jsonRegulator("fixed") = False
                If jsonFeature("attributes")("SUBTYPECD") = "5" Then
                    jsonRegulator("fixed") = True
                ElseIf jsonFeature("attributes")("SUBTYPECD") = "8" Then
                    jsonRegulator("auto") = True
                    jsonRegulator("size") = jsonFeature("attributes")("RATEDCURRENT")
                End If

                If jsonRegulator("lcp") <> "" Then json("regulators").Add jsonRegulator
            End If
        Next jsonFeature
        
        For Each jsonFeature In jsonIsolators("features")
            If LoadingBar_Form2.gTotal = 0 Then Exit Function
            x = jsonFeature("geometry")("x")
            y = jsonFeature("geometry")("y")

            Set closestPole = isClosestPoleOnJob(jsonPoles, poleCollection, x, y)
            If Not closestPole Is Nothing Then
                Set jsonIsolator = New Scripting.Dictionary
                
                If closestPole.poleNumber = "NEWGISPOLE" Then
                    jsonIsolator("adjacent") = True
                Else
                    jsonIsolator("adjacent") = False
                End If
                
                jsonIsolator("x") = closestPole.x
                jsonIsolator("y") = closestPole.y
                jsonIsolator("lcp") = jsonFeature("attributes")("LCP")
                jsonIsolator("rotation") = jsonFeature("attributes")("SYMBOLROTATION")
                jsonIsolator("size") = jsonFeature("attributes")("RATEDKVA")
                jsonIsolator("phase") = phases(jsonFeature("attributes")("PHASEDESIGNATION"))

                If jsonIsolator("lcp") <> "" Then json("isolators").Add jsonIsolator
            End If
        Next jsonFeature
        
        For Each pole In poleCollection
            If LoadingBar_Form2.gTotal = 0 Then Exit Function
            ' Get Transformers
            If pole.transformerSizes <> 0 Then
                Set jsonXFMR = New Scripting.Dictionary
                If pole.poleNumber = "NEWGISPOLE" Then
                    jsonXFMR("adjacent") = True
                Else
                    jsonXFMR("adjacent") = False
                End If
                jsonXFMR("size") = ""
                replaceSection = False
                crewNotes = pole.alt1
                Call Utilities.applyStandardAbbreviations(crewNotes)
                lines = Split(crewNotes, vbLf)
                For Each line In lines
                    line = Trim(UCase(line))
                    If line = "REPLACE" Then replaceSection = True
                    If line = "TRANSFER" Then Exit For
                    If replaceSection Then
                        If InStr(line, "XFMR") > 0 Then
                            If InStr(line, "/") > 0 Then line = Split(line, "/")(1)
                            If InStr(line, "KVA") > 0 Then jsonXFMR("size2") = Left(line, InStr(line, "KVA") - 1)
                            If IsNumeric(jsonXFMR("size2")) Then jsonXFMR("size2") = CInt(jsonXFMR("size2"))
                        End If
                    End If
                Next line
                For Each jsonFeature In jsonPoles("features")
                    If LoadingBar_Form2.gTotal = 0 Then Exit Function
                    If pole.existingCEID = jsonFeature("attributes")("CE_TAG") Or pole.gisCEID = jsonFeature("attributes")("CE_TAG") Then
                        
                        x = jsonFeature("geometry")("x")
                        y = jsonFeature("geometry")("y")
                        x2 = 0
                        y2 = 0
                        For Each jsonTransformerConnectionFeature In jsonTransformerConnections("features")
                            For Each jsonTransformerConnectionPath In jsonTransformerConnectionFeature("geometry")("paths")
                                If jsonTransformerConnectionPath(1)(1) = x And jsonTransformerConnectionPath(1)(2) = y Then
                                    x2 = jsonTransformerConnectionPath(2)(1)
                                    y2 = jsonTransformerConnectionPath(2)(2)
                                    Exit For
                                End If
                                If jsonTransformerConnectionPath(2)(1) = x And jsonTransformerConnectionPath(2)(2) = y Then
                                    x2 = jsonTransformerConnectionPath(1)(1)
                                    y2 = jsonTransformerConnectionPath(1)(2)
                                    Exit For
                                End If
                            Next jsonTransformerConnectionPath
                            If x2 And y2 <> 0 Then
                                For Each jsonTransformer In jsonTransformers("features")
                                    If jsonTransformer("geometry")("x") = x2 And jsonTransformer("geometry")("y") = y2 Then
                                        jsonXFMR("x") = x
                                        jsonXFMR("y") = y
                                        jsonXFMR("phase") = phases(jsonTransformer("attributes")("PHASEDESIGNATION"))
                                        If pole.transformerSizes > 0 Then
                                            jsonXFMR("size") = pole.transformerSizes
                                        Else
                                            jsonXFMR("size") = CInt(Utilities.OnlyNumbers(CStr(jsonTransformer("attributes")("RATEDKVA"))))
                                        End If
                                        jsonXFMR("TLM") = Mid(jsonTransformer("attributes")("TLM"), Len(jsonTransformer("attributes")("TLM")) - 3)
                                        jsonXFMR("lowSideVoltage") = lowSideVoltages(jsonTransformer("attributes")("LOWSIDEVOLTAGE"))
                                        json("transformers").Add jsonXFMR
                                        Exit For
                                    End If
                                Next jsonTransformer
                            Exit For
                            End If
                        Next jsonTransformerConnectionFeature

                        closestDistance = -1
                        Dim closestTransformer As Object
                        If jsonXFMR.count = 0 Then
                            For Each jsonTransformer In jsonTransformers("features")
                                x2 = jsonTransformer("geometry")("x")
                                y2 = jsonTransformer("geometry")("y")
                                distance = Sqr((x2 - x) ^ 2 + (y2 - y) ^ 2)
                                If closestDistance = -1 Or distance < closestDistance Then
                                    closestDistance = distance
                                    Set closestTransformer = jsonTransformer
                                End If
                            Next jsonTransformer
                            
                            If Not closestTransformer Is Nothing Then
                                x2 = closestTransformer("geometry")("x")
                                y2 = closestTransformer("geometry")("y")
                                
                                closestDistance = -1
                                For Each jsonFeature2 In jsonPoles("features")
                                    x3 = jsonFeature2("geometry")("x")
                                    y3 = jsonFeature2("geometry")("y")
                                    distance = Sqr((x3 - x2) ^ 2 + (y3 - y2) ^ 2)
                                    If closestDistance = -1 Or distance < closestDistance Then
                                        closestDistance = distance
                                        Set closestPole = jsonFeature2
                                    End If
                                Next jsonFeature2
                            
                            
                                If Not closestPole Is Nothing Then
                                    If pole.existingCEID = closestPole("attributes")("CE_TAG") Or pole.gisCEID = closestPole("attributes")("CE_TAG") Then
                                        jsonXFMR("x") = x
                                        jsonXFMR("y") = y
                                        jsonXFMR("phase") = phases(closestTransformer("attributes")("PHASEDESIGNATION"))
                                        If pole.transformerSizes > 0 Then
                                            jsonXFMR("size") = pole.transformerSizes
                                        Else
                                            jsonXFMR("size") = CInt(Utilities.OnlyNumbers(CStr(closestTransformer("attributes")("RATEDKVA"))))
                                        End If
                                        jsonXFMR("TLM") = Mid(closestTransformer("attributes")("TLM"), Len(closestTransformer("attributes")("TLM")) - 3)
                                        jsonXFMR("lowSideVoltage") = lowSideVoltages(closestTransformer("attributes")("LOWSIDEVOLTAGE"))
                                        json("transformers").Add jsonXFMR
                                    End If
                                End If
                            End If
                        End If
                        
                        Exit For
                    End If
                Next jsonFeature
            End If
            ' End Transformers
            
        Next pole
    Next poleCollection
    
    jsonText = JsonConverter.ConvertToJson(json, Whitespace:=2)
    filePath = "C:\ProgramData\Bentley\OpenUtilities Map Connect Edition\Configuration\WorkSpaces\ConsumersEnergy\Standards\vba\print.json"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.CreateTextFile(filePath, True, False)
    
    file.Write jsonText
    file.Close
    
    generateJSON = True
    Call LoadingBar_Form2.FinishProgress
End Function

Function getClosestPole(jsonPoles As Object, x As Double, y As Double, Optional attributes As Scripting.Dictionary, Optional ignoreZeros As Boolean) As Variant()
    Dim closestFeature As Object: Set closestFeature = Nothing
    Dim closestDistance As Double: closestDistance = -1
    Dim jsonFeature As Object
    For Each jsonFeature In jsonPoles("features")
        x2 = jsonFeature("geometry")("x")
        y2 = jsonFeature("geometry")("y")
        
        Dim attMatch As Boolean: attMatch = True
        If Not attributes Is Nothing Then
            For Each att In attributes
                If Not jsonFeature("attributes").exists(att) Then attMatch = False: Exit For
                Dim attValueMatch As Boolean: attValueMatch = False
                For Each attValue In attributes(att)
                    If jsonFeature("attributes")(att) = attValue Then attValueMatch = True: Exit For
                Next attValue
                If Not attValueMatch Then Exit For
            Next att
        End If
        
        If attMatch Then
            distance = Sqr((x2 - x) ^ 2 + (y2 - y) ^ 2)
            If closestDistance = -1 Or distance < closestDistance Then
                If Not ignoreZeros Or distance <> 0 Then
                    Set closestFeature = jsonFeature
                    closestDistance = distance
                End If
            End If
        End If
    Next jsonFeature
    
    getClosestPole = Array(closestFeature, closestDistance)
End Function
Function isClosestPoleOnJob(jsonPoles As Object, poleCollection As Collection, x As Double, y As Double) As pole
    results = getClosestPole(jsonPoles, x, y)
    Set closestFeature = results(0)
    
    Set isClosestPoleOnJob = Nothing
    
    Dim pole As pole
    If Not IsNull(closestFeature("attributes")("CE_TAG")) Then
        If Utilities.isCEID(CStr(closestFeature("attributes")("CE_TAG"))) Then
            For Each pole In poleCollection
                If closestFeature("attributes")("CE_TAG") = pole.existingCEID Or closestFeature("attributes")("CE_TAG") = pole.gisCEID Then
                    Set isClosestPoleOnJob = pole
                    Exit For
                End If
            Next pole
        End If
    End If
    
    For Each pole In poleCollection
        If Not Utilities.isCEID(pole.existingCEID) Then
            dx = pole.x - x
            dy = pole.y - y
            distance = Sqr((dx ^ 2) + (dy ^ 2))
            If distance < 40 Then
                Set isClosestPoleOnJob = pole
                Exit For
            End If
        End If
    Next pole
    
End Function
            
Function findPoleGroups(poles As Collection) As Collection
    Dim pole As pole
    Dim found As Scripting.Dictionary: Set found = New Scripting.Dictionary
    Dim poleGroups As Collection: Set poleGroups = New Collection
    
    Dim poleGroup As Collection
    While found.count <> poles.count
        Set poleGroup = New Collection
        For Each pole In poles
            If Not found.exists(pole.poleNumber) Then Exit For
        Next pole
        found.Add pole.poleNumber, pole
        If pole.latitude <> 0 And pole.longitude <> 0 Then
            poleGroup.Add pole
        
            Call getAllConnectedPoles(pole, found, poleGroup, poles)
        End If
        
        poleGroups.Add poleGroup
    Wend
    
    Set findPoleGroups = poleGroups
End Function

Sub getAllConnectedPoles(pole As pole, found As Scripting.Dictionary, poleGroup As Collection, poles As Collection)
    Dim Span As Span
    Dim otherPole As pole
    For Each Span In pole.spans
        If Span.otherPole <> "" Then
            If Not found.exists(Span.otherPole) Then
                For Each otherPole In poles
                    If otherPole.poleNumber = Span.otherPole Then
                        If otherPole.latitude <> 0 And otherPole.longitude <> 0 Then
                            found.Add otherPole.poleNumber, otherPole
                            poleGroup.Add otherPole
                            Call getAllConnectedPoles(otherPole, found, poleGroup, poles)
                            Exit For
                        End If
                    End If
                Next otherPole
            End If
        End If
    Next Span
End Sub

Function GetToken() As String
    Dim downloadsPath As String
    downloadsPath = Environ("USERPROFILE") & "\Downloads\"
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fso.GetFolder(downloadsPath)
    
    Dim file As Object
    Dim newestFile As Object
    Dim latestDate As Date
    latestDate = #1/1/1980#

    For Each file In folder.Files
        If LCase(file.name) Like "arcgis_token*.json" Then
            If LCase(file.name) = "arcgis_token.json" Or file.name Like "*(*).json" Then
                If file.DateLastModified > latestDate Then
                    latestDate = file.DateLastModified
                    Set newestFile = file
                End If
            End If
        End If
    Next file

    If newestFile Is Nothing Then
        MsgBox "No arcgis_token.json file found in Downloads.", vbExclamation
        GetToken = ""
        Exit Function
    End If
    
    For Each file In folder.Files
        If LCase(file.name) Like "arcgis_token*.json" Then
            If LCase(file.name) = "arcgis_token.json" Or file.name Like "*(*).json" Then
                If file.path <> newestFile.path Then
                    file.Delete True
                End If
            End If
        End If
    Next file

    Dim TextStream As Object
    Set TextStream = newestFile.OpenAsTextStream(1, -2)
    
    Dim jsonRaw As String
    jsonRaw = TextStream.ReadAll
    TextStream.Close
    
    Dim jsonParsed As Object
    Set jsonParsed = JsonConverter.ParseJson(jsonRaw)
    
    GetToken = jsonParsed("token")
    
    Dim finalPath As String
    finalPath = downloadsPath & "arcgis_token.json"
    
    If newestFile.path <> finalPath Then
        If fso.FileExists(finalPath) Then fso.DeleteFile finalPath, True
        newestFile.name = "arcgis_token.json"
    End If
End Function


Function injectHotkey() As Boolean
    Dim fileNum As Integer
    Dim fileText As String
    Dim lineBreakPos As Long
    Dim i As Integer
    Dim filePath As String
    
    Dim filePath1 As String: filePath1 = "C:\ProgramData\Bentley\OpenUtilities Map Connect Edition\Configuration\Organization\USER_APPSETTINGS_DFLTS\Consumers_KeyboardShortcutsSeed.xml"
    Dim filePath2 As String: filePath2 = "C:\Users\" & Environ$("USERNAME") & "\AppData\Local\Bentley\OpenUtilitiesMap\10.0.0\prefs\Personal.KeyboardShortcuts.xml"
    
    If Dir(filePath2) <> "" Then
        filePath = filePath2
    ElseIf Dir(filePath1) <> "" Then
        filePath = filePath1
    Else
        MsgBox "Failed to find keyboard shortcut file. Can't inject hotkey."
        injectHotkey = False
        Exit Function
    End If
    
    Dim keyinShortcut As String: keyinShortcut = "" & _
    vbTab & "<KeyboardShortcut ScanCode=""0x0a"" Comment=""9"">" & vbLf & _
    vbTab & vbTab & "<Label>Generate Print</Label>" & vbLf & _
    vbTab & vbTab & "<Keyin>vba run GeneratePrint</Keyin>" & vbLf & _
    vbTab & "</KeyboardShortcut>" & vbLf
    
    fileNum = FreeFile()
    Open filePath For Input As #fileNum
    fileText = Input$(LOF(fileNum), fileNum)
    Close #fileNum
    
    If InStr(1, fileText, Trim(keyinShortcut), vbTextCompare) > 0 Then
        injectHotkey = False
        Exit Function
    End If
    
    lineBreakPos = 1
    For i = 1 To 2
        lineBreakPos = InStr(lineBreakPos, fileText, vbCrLf)
        If lineBreakPos = 0 Then
            Debug.Print "File has fewer than 3 lines. Cannot inject.", vbExclamation
            injectHotkey = False
            Exit Function
        End If
        lineBreakPos = lineBreakPos + 2
    Next i
    
    fileText = Left(fileText, lineBreakPos - 1) & keyinShortcut & Mid(fileText, lineBreakPos)
    
    fileNum = FreeFile()
    Open filePath For Output As #fileNum
    Print #fileNum, fileText;
    Close #fileNum

    injectHotkey = True
End Function

Function ForceInjectModuleToBentley(strModuleName As String) As Boolean
    Dim BentleyConnector As Object
    Dim BentleyEngine As Object
    Dim targetProject As Object
    Dim strTempBasPath As String
    Dim strMVBAProjectPath As String
    Dim strProjectNameOnly As String
    Dim fileFound As Boolean
    
    strTempBasPath = "C:\ProgramData\Bentley\OpenUtilities Map Connect Edition\Configuration\WorkSpaces\ConsumersEnergy\Standards\vba\temp.bas"
    strMVBAProjectPath = "C:\ProgramData\Bentley\OpenUtilities Map Connect Edition\Configuration\WorkSpaces\ConsumersEnergy\Standards\vba\CECADReferences.mvba"
    strProjectNameOnly = CreateObject("Scripting.FileSystemObject").GetBaseName(strMVBAProjectPath)

    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents(strModuleName).Export strTempBasPath
    If Err.Number <> 0 Then
        MsgBox "Failed to export Excel module. Go to Excel Trust Center and enable 'Trust access to the VBA project object model'.", vbCritical
        ForceInjectModuleToBentley = False
        Exit Function
    End If
    On Error GoTo 0
    
    On Error Resume Next
    Set BentleyConnector = GetObject(, "MicroStationDGN.ApplicationObjectConnector")
    If BentleyConnector Is Nothing Then
        MsgBox "Please make sure Bentley OpenUtilities is open and running a DGN file first.", vbExclamation
        ForceInjectModuleToBentley = False
        Exit Function
    End If
    Set BentleyEngine = BentleyConnector.Application
    On Error GoTo 0
    
    BentleyEngine.CadInputQueue.SendKeyin "VBA LOAD """ & strMVBAProjectPath & """"
    DoEvents
    
    On Error Resume Next
    For Each targetProject In BentleyEngine.VBE.VBProjects
        If UCase(targetProject.name) = UCase(strProjectNameOnly) Then
            targetProject.VBComponents.Remove targetProject.VBComponents(strModuleName)
            targetProject.VBComponents.Import strTempBasPath
            fileFound = True
            Exit For
        End If
    Next targetProject
    On Error GoTo 0
    
    If fileFound Then
        BentleyEngine.CadInputQueue.SendKeyin "VBA SAVE " & strProjectNameOnly
    Else
        MsgBox "VBA target project matching '" & strProjectNameOnly & "' was not found active in Bentley's workspace memory.", vbCritical
    End If
    
    Set BentleyEngine = Nothing
    Set BentleyConnector = Nothing
    
    ForceInjectModuleToBentley = True
End Function

Function ForceInjectUserFormToBentley(strModuleName As String) As Boolean
    Dim BentleyConnector As Object
    Dim BentleyEngine As Object
    Dim targetProject As Object
    Dim strTempFrmPath As String
    Dim strTempFrxPath As String
    Dim strMVBAProjectPath As String
    Dim strProjectNameOnly As String
    Dim fileFound As Boolean
    Dim fso As Object
    
    strTempFrmPath = "C:\ProgramData\Bentley\OpenUtilities Map Connect Edition\Configuration\WorkSpaces\ConsumersEnergy\Standards\vba\" & strModuleName & ".frm"
    strTempFrxPath = "C:\ProgramData\Bentley\OpenUtilities Map Connect Edition\Configuration\WorkSpaces\ConsumersEnergy\Standards\vba\" & strModuleName & ".frx"
    
    strMVBAProjectPath = "C:\ProgramData\Bentley\OpenUtilities Map Connect Edition\Configuration\WorkSpaces\ConsumersEnergy\Standards\vba\CECADReferences.mvba"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    strProjectNameOnly = fso.GetBaseName(strMVBAProjectPath)

    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents(strModuleName).Export strTempFrmPath
    If Err.Number <> 0 Then
        MsgBox "Failed to export Excel form. Go to Excel Trust Center and enable 'Trust access to the VBA project object model'.", vbCritical
        ForceInjectUserFormToBentley = False
        Exit Function
    End If
    On Error GoTo 0
    
    On Error Resume Next
    Set BentleyConnector = GetObject(, "MicroStationDGN.ApplicationObjectConnector")
    If BentleyConnector Is Nothing Then
        MsgBox "Please make sure Bentley OpenUtilities is open and running a DGN file first.", vbExclamation
        ForceInjectUserFormToBentley = False
        Exit Function
    End If
    Set BentleyEngine = BentleyConnector.Application
    On Error GoTo 0
    
    BentleyEngine.CadInputQueue.SendKeyin "VBA LOAD """ & strMVBAProjectPath & """"
    DoEvents
    
    On Error Resume Next
    For Each targetProject In BentleyEngine.VBE.VBProjects
        If UCase(targetProject.name) = UCase(strProjectNameOnly) Then
            targetProject.VBComponents.Remove targetProject.VBComponents(strModuleName)
            targetProject.VBComponents.Import strTempFrmPath
            fileFound = True
            Exit For
        End If
    Next targetProject
    On Error GoTo 0
    
    If fileFound Then
        BentleyEngine.CadInputQueue.SendKeyin "VBA SAVE " & strProjectNameOnly
    Else
        MsgBox "VBA target project matching '" & strProjectNameOnly & "' was not found active in Bentley's workspace memory.", vbCritical
    End If
    
    On Error Resume Next
    If fso.FileExists(strTempFrmPath) Then fso.DeleteFile strTempFrmPath
    If fso.FileExists(strTempFrxPath) Then fso.DeleteFile strTempFrxPath
    On Error GoTo 0
    
    Set fso = Nothing
    Set BentleyEngine = Nothing
    Set BentleyConnector = Nothing
    
    ForceInjectUserFormToBentley = True
End Function
