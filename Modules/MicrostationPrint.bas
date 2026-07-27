Attribute VB_Name = "MicrostationPrint"
Sub GenerateMicrostationPrintFiles()
    Dim restartMicrostationNeeded As Boolean
    
    restartMicrostationNeeded = injectHotkey()
    'Call ForceInjectModuleToBentley
    If Not generateJSON Then Exit Sub
    
    If restartMicrostationNeeded Then
        MsgBox "Restart Open Map Utilities if open for changes to take effect."
    End If
    
    MsgBox "Press '9' to generate the print on open map utilities. Script will have to be rerun for future generations."
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
    Dim service As wire
    Dim Span As Span
    Dim anchor As anchor
    Dim wire As wire
    Dim x As Double
    Dim y As Double
    Dim x2 As Double
    Dim y2 As Double
    Dim spgCount As Integer
    Dim radAngle As Double
    
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
    
    Dim usedPoles As Scripting.Dictionary: Set usedPoles = New Scripting.Dictionary
    
    Dim groupNumber As Integer: groupNumber = 1
    Dim lowestLatitude As Double
    Dim lowestLongitude As Double
    Dim highestLatitude As Double
    Dim highestLongitude As Double
    Dim poleCollections As Collection: Set poleCollections = findPoleGroups(project.poles)
    Dim poleCollection As Collection
    
    Set json("groups") = New Scripting.Dictionary
    For Each poleCollection In poleCollections
        lowestLatitude = 0
        lowestLongitude = 0
        highestLatitude = 0
        highestLongitude = 0
        For i = 1 To poleCollection.count
            Set pole = project.findPole(poleCollection(i).poleNumber)
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
        
        Set json("groups")(groupNumber) = jsonGroup
        groupNumber = groupNumber + 1
    Next poleCollection
    
    Set json("poles") = New Collection
    Set json("streetlights") = New Collection
    Set json("spanguys") = New Collection
    Set json("openWires") = New Collection
    Set json("secWires") = New Collection
    Set json("priWires") = New Collection
    For Each pole In project.poles
        Set jsonPole = New Scripting.Dictionary
        jsonPole("ceid") = pole.existingCEID
        results = LatLonToMI2253(pole.latitude, pole.longitude)
        jsonPole("x") = results(0)
        jsonPole("y") = results(1)
        jsonPole("height") = pole.height / 12
        jsonPole("crewNotes") = pole.alt1
        jsonPole("class") = pole.Class
        jsonPole("location") = pole.location
        jsonPole("replace") = pole.replacePole
        jsonPole("newHeight") = pole.newHeight / 12
        jsonPole("newClass") = pole.newClass
        jsonPole("tree") = pole.treeWork
        jsonPole("topped") = pole.topped
        jsonPole("group") = pole.groupNumber
        jsonPole("skipSpan") = pole.skipSpan
        
        Set jsonPole("services") = New Collection
        If pole.locationAdjacent Then
            For Each service In pole.services
                For Each midspan In service.midspans
                    Set jsonService = New Scripting.Dictionary
                    Set Span = pole.spans(midspan)
                    jsonService("angle") = Span.angle
                    jsonService("distance") = Span.distance
                    jsonService("address") = Span.houseNumber
                    jsonPole("services").Add jsonService
                Next midspan
            Next service
        End If
        
        Set jsonPole("guys") = New Collection
        For Each anchor In pole.anchors
            If anchor.owner = "Consumers Energy" Then
                Set jsonGuy = New Scripting.Dictionary
                jsonGuy("angle") = anchor.angle
                jsonGuy("count") = anchor.ceCount
                jsonPole("guys").Add jsonGuy
            End If
        Next anchor
        
        json("poles").Add jsonPole
        
        If pole.slBottomBracketHeight <> 0 Then
            Set jsonStreetlight = New Scripting.Dictionary
            jsonStreetlight("x") = results(0)
            jsonStreetlight("y") = results(1)
            json("streetlights").Add jsonStreetlight
        End If
        
        If pole.locationAdjacent Then
            For Each Span In pole.spans
                If pole.location <> "" Or Utilities.SheetExists(Span.otherPole) Then
                    If Not usedPoles.exists(Span.otherPole) Then
                        Set owWires = MicrostationUtilities.getItems(pole, Span.spanSlot, "OW")
                        Set secWires = MicrostationUtilities.getItems(pole, Span.spanSlot, "SEC")
                        Set priWires = MicrostationUtilities.getItems(pole, Span.spanSlot, "PRI")
                        spgCount = 0
                        For i = Span.utilWires.count To 1 Step -1
                            Set wire = Span.utilWires(i)
                            If wire.componentType = "SPG" Then spgCount = spgCount + 1
                        Next i
                        
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
                    
                        
                        If spgCount > 0 Then
                            Set jsonSpanguy = New Scripting.Dictionary
                            jsonSpanguy("length") = Utilities.OnlyNumbers(Span.distance)
                            jsonSpanguy("count") = spgCount
                            jsonSpanguy("angle") = Span.angle
                            jsonSpanguy("x1") = results(0)
                            jsonSpanguy("y1") = results(1)
                            jsonSpanguy("x2") = x2
                            jsonSpanguy("y2") = y2
        
                            json("spanguys").Add jsonSpanguy
                        End If
                        
                        layer = 0
                        
                        For Each secWire In secWires
                            Set jsonSecWire = New Scripting.Dictionary
                            jsonSecWire("length") = Utilities.OnlyNumbers(Span.distance)
                            jsonSecWire("angle") = Span.angle
                            jsonSecWire("layer") = layer
                            jsonSecWire("size") = secWire("size")
                            
                            jsonSecWire("x1") = results(0)
                            jsonSecWire("y1") = results(1)
                            jsonSecWire("x2") = x2
                            jsonSecWire("y2") = y2
                           
                            jsonSecWire("label") = layer = 0 And spgCount = 0
                            
                            jsonSecWire("group") = pole.groupNumber
                            json("secWires").Add jsonSecWire
                            layer = layer + 1
                        Next secWire
                        
                        If owWires.count > 0 Then
                            Set jsonOpenwire = New Scripting.Dictionary
                            jsonOpenwire("length") = Utilities.OnlyNumbers(Span.distance)
                            jsonOpenwire("angle") = Span.angle
                            jsonOpenwire("layer") = layer
                            
                            openWireString = ""
                            For Each owWire In owWires
                                openWireString = openWireString & owWire("size") & "-"
                            Next owWire
                            
                            If openWireString <> "" Then
                                openWireString = Left(openWireString, Len(openWireString) - 1)
                                jsonOpenwire("size") = openWireString
                            End If
                            
                            jsonOpenwire("x1") = results(0)
                            jsonOpenwire("y1") = results(1)
                            jsonOpenwire("x2") = x2
                            jsonOpenwire("y2") = y2
                            
                            jsonOpenwire("label") = layer = 0 And spgCount = 0
                            
                            jsonOpenwire("group") = pole.groupNumber
                            json("openWires").Add jsonOpenwire
                            layer = layer + 1
                        End If
                        
                        For Each priWire In priWires
                            Set jsonPriWire = New Scripting.Dictionary
                            jsonPriWire("length") = Utilities.OnlyNumbers(Span.distance)
                            jsonPriWire("angle") = Span.angle
                            jsonPriWire("layer") = 1
                            
                            jsonPriWire("size") = priWire("size")
                            jsonPriWire("phase") = priWire("phase")
                            jsonPriWire("configuration") = priWire("config")
                            
                            jsonPriWire("x1") = results(0)
                            jsonPriWire("y1") = results(1)
                            jsonPriWire("x2") = x2
                            jsonPriWire("y2") = y2
                            
                            jsonPriWire("label") = layer = 0 And spgCount = 0
                            
                            jsonPriWire("group") = pole.groupNumber
                            json("priWires").Add jsonPriWire
                            layer = layer + 1
                        Next priWire
                    End If
                End If
            Next Span
            usedPoles.Add pole.poleNumber, Nothing
        End If
    Next pole
    
    Dim layers As Collection: Set layers = New Collection
    layers.Add 7  ' Major Roads
    layers.Add 8  ' Minor Roads
    layers.Add 9  ' Pavement Lines
    layers.Add 10 ' Private Drive
    layers.Add 11 ' Raildroad
    layers.Add 15 ' Waterway
    
    Set json("roads") = New Collection
    For Each layer In layers
        For Each poleCollection In poleCollections
            
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
    
    Set json("centerlines") = New Collection
    For Each poleCollection In poleCollections
            
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
    
    For Each poleCollection In poleCollections
        For Each pole In poleCollection
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
                        results = LatLonToMI2253(pole.latitude, pole.longitude)
                        radAngle = (90 - Span.angle) * (3.14159265358979 / 180)
                        x = results(0) + (Span.distance * Cos(radAngle))
                        y = results(1) + (Span.distance * Sin(radAngle))
                        jsonPole("x") = x
                        jsonPole("y") = y
                        
                        Set jsonPoles = getOtherPole(x, y, 20, token)
                        
                        
                        Dim closestFeature As Object
                        Dim closestDistance As Double
                        
                        results = getClosestPole(jsonPoles, x, y)
                        Set closestFeature = results(0)
                        closestDistance = results(1)
                        
                        If Not closestFeature Is Nothing Then
                            If closestFeature("attributes")("CE_TAG") <> pole.existingCEID And closestFeature("attributes")("CE_TAG") <> pole.gisCEID Then jsonPole("ceid") = closestFeature("attributes")("CE_TAG")
                            jsonPole("height") = closestFeature("attributes")("HEIGHT")
                            jsonPole("class") = closestFeature("attributes")("CLASS")
                            jsonPole("x2") = closestFeature("geometry")("x")
                            jsonPole("y2") = closestFeature("geometry")("y")
                        End If
                        
                        Dim duplicatePole As Boolean
                        For Each jsonPole2 In json("poles")
                            If (jsonPole2("x") = jsonPole("x") And jsonPole2("y") = jsonPole("y")) Or (jsonPole2("x2") = jsonPole("x2") And jsonPole2("y2") = jsonPole("y2")) Then
                                duplicatedPole = True
                                Exit For
                            End If
                        Next jsonPole2
                        If Not duplicatePole Then
                            json("poles").Add jsonPole
                            otherPole.gisCEID = jsonPole("ceid")
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
    
    Set json("fuses") = New Collection
    Set json("reclosers") = New Collection
    Set json("sectionalizers") = New Collection
    Set json("switches") = New Collection
    Set json("capacitors") = New Collection
    Set json("regulators") = New Collection
    
    Set json("risers") = New Collection
    For Each poleCollection In poleCollections
        Set jsonPoles = getElectricJson(poleCollection, 3, token)
        Set jsonTransformerConnections = getElectricJson(poleCollection, 30, token)
        Set jsonTransformers = getElectricJson(poleCollection, 27, token)
        Set jsonFuses = getElectricJson(poleCollection, 60, token)
        Set jsonReclosers = getElectricJson(poleCollection, 70, token)
        Set jsonSectionalizers = getElectricJson(poleCollection, 71, token)
        Set jsonSwitches = getElectricJson(poleCollection, 65, token)
        Set jsonPriRisers = getElectricJson(poleCollection, 4, token)
        Set jsonSecRisers = getElectricJson(poleCollection, 5, token)
        For Each jsonFeature In jsonFuses("features")
            x = jsonFeature("geometry")("x")
            y = jsonFeature("geometry")("y")
            
            If isClosestPoleOnJob(jsonPoles, poleCollection, x, y) Then
                Set jsonFuse = New Scripting.Dictionary
                jsonFuse("x") = x
                jsonFuse("y") = y
                jsonFuse("lcp") = jsonFeature("attributes")("LCP")
                jsonFuse("size") = jsonFeature("attributes")("RATEDCURRENT")
                jsonFuse("open") = jsonFeature("attributes")("SWITCHSTATUS") = 0
                jsonFuse("rotation") = jsonFeature("attributes")("SYMBOLROTATION")
                json("fuses").Add jsonFuse
            End If
        Next jsonFeature
        
        For Each jsonFeature In jsonReclosers("features")
            x = jsonFeature("geometry")("x")
            y = jsonFeature("geometry")("y")
            
            If isClosestPoleOnJob(jsonPoles, poleCollection, x, y) Then
                Set jsonRecloser = New Scripting.Dictionary
                jsonRecloser("x") = x
                jsonRecloser("y") = y
                jsonRecloser("lcp") = jsonFeature("attributes")("LCP")
                jsonRecloser("size") = jsonFeature("attributes")("RATEDCURRENT")
                jsonRecloser("open") = jsonFeature("attributes")("SWITCHSTATUS") = 0
                jsonRecloser("rotation") = jsonFeature("attributes")("SYMBOLROTATION")
                json("reclosers").Add jsonRecloser
            End If
        Next jsonFeature
    
        For Each jsonFeature In jsonSectionalizers("features")
            x = jsonFeature("geometry")("x")
            y = jsonFeature("geometry")("y")
            
            If isClosestPoleOnJob(jsonPoles, poleCollection, x, y) Then
                Set jsonSectionalizer = New Scripting.Dictionary
                jsonSectionalizer("x") = x
                jsonSectionalizer("y") = y
                jsonSectionalizer("lcp") = jsonFeature("attributes")("LCP")
                jsonSectionalizer("size") = jsonFeature("attributes")("RATEDCURRENT")
                jsonSectionalizer("open") = jsonFeature("attributes")("SWITCHSTATUS") = 0
                jsonSectionalizer("rotation") = jsonFeature("attributes")("SYMBOLROTATION")
                json("sectionalizers").Add jsonSectionalizer
            End If
        Next jsonFeature
        
        For Each jsonFeature In jsonSwitches("features")
            x = jsonFeature("geometry")("x")
            y = jsonFeature("geometry")("y")
            
            If isClosestPoleOnJob(jsonPoles, poleCollection, x, y) Then
                Set jsonSwitch = New Scripting.Dictionary
                jsonSwitch("x") = x
                jsonSwitch("y") = y
                jsonSwitch("lcp") = jsonFeature("attributes")("LCP")
                jsonSwitch("size") = jsonFeature("attributes")("RATEDCURRENT")
                jsonSwitch("open") = jsonFeature("attributes")("SWITCHSTATUS") = 0
                jsonSwitch("rotation") = jsonFeature("attributes")("SYMBOLROTATION")
                
                equipmentId = jsonFeature("attributes")("EQUIPMENTID")
                If equipmentId = "S_BLADE" Or equipmentId = "S_LINK" Or equipmentId = "SB_300A_BLADE" Or equipmentId = "SL_LINK_100A" Or equipmentId = "SL_LINK_200A" Then
                    json("fuses").Add jsonSwitch
                Else
                    json("switches").Add jsonSwitch
                End If
            End If
        Next jsonFeature
        
        For Each jsonFeature In jsonPriRisers("features")
            x = jsonFeature("geometry")("x")
            y = jsonFeature("geometry")("y")
            
            If isClosestPoleOnJob(jsonPoles, poleCollection, x, y) Then
                Set jsonRiser = New Scripting.Dictionary
                jsonRiser("x") = x
                jsonRiser("y") = y
                jsonRiser("lcp") = jsonFeature("attributes")("LCP")
                jsonRiser("rotation") = jsonFeature("attributes")("SYMBOLROTATION")
                jsonRiser("type") = "Primary"
                
                json("risers").Add jsonRiser
            End If
        Next jsonFeature
        
        For Each jsonFeature In jsonSecRisers("features")
            x = jsonFeature("geometry")("x")
            y = jsonFeature("geometry")("y")
            
            If isClosestPoleOnJob(jsonPoles, poleCollection, x, y) Then
                Set jsonRiser = New Scripting.Dictionary
                jsonRiser("x") = x
                jsonRiser("y") = y
                jsonRiser("lcp") = jsonFeature("attributes")("LCP")
                jsonRiser("rotation") = jsonFeature("attributes")("SYMBOLROTATION")
                jsonRiser("type") = "Secondary"
                
                json("risers").Add jsonRiser
            End If
        Next jsonFeature
        
        For Each pole In poleCollection
            ' Get Transformers
            If pole.transformerSizes <> 0 Then
                For Each jsonFeature In jsonPoles("features")
                    If pole.existingCEID = jsonFeature("attributes")("CE_TAG") Or pole.gisCEID = jsonFeature("attributes")("CE_TAG") Then
                        Set jsonXFMR = New Scripting.Dictionary
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
                                Dim closestPole As Object
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
End Function

Function getClosestPole(jsonPoles As Object, x As Double, y As Double) As Variant()
    Dim closestFeature As Object: Set closestFeature = Nothing
    Dim closestDistance As Double: closestDistance = -1
    Dim jsonFeature As Object
    For Each jsonFeature In jsonPoles("features")
        x2 = jsonFeature("geometry")("x")
        y2 = jsonFeature("geometry")("y")
        
        distance = Sqr((x2 - x) ^ 2 + (y2 - y) ^ 2)
        If closestDistance = -1 Or distance < closestDistance Then
            Set closestFeature = jsonFeature
            closestDistance = distance
        End If
    Next jsonFeature
    
    getClosestPole = Array(closestFeature, closestDistance)
End Function
Function isClosestPoleOnJob(jsonPoles As Object, poleCollection As Collection, x As Double, y As Double) As Boolean
    results = getClosestPole(jsonPoles, x, y)
    Set closestFeature = results(0)
    
    isClosestPoleOnJob = False
    
    Dim pole As pole
    For Each pole In poleCollection
        If closestFeature("attributes")("CE_TAG") = pole.existingCEID Or closestFeature("attributes")("CE_TAG") = pole.gisCEID Then
            isClosestPoleOnJob = True
            Exit For
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
        poleGroup.Add pole
        
        Call getAllConnectedPoles(pole, found, poleGroup)
        
        poleGroups.Add poleGroup
    Wend
    
    Set findPoleGroups = poleGroups
End Function

Sub getAllConnectedPoles(pole As pole, found As Scripting.Dictionary, poleGroup As Collection)
    Dim Span As Span
    Dim otherPole As pole
    For Each Span In pole.spans
        If Span.otherPole <> "" Then
            If Not found.exists(Span.otherPole) Then
                Set otherPole = Utilities.getPole(Span.otherPole)
                found.Add otherPole.poleNumber, otherPole
                poleGroup.Add otherPole
                Call getAllConnectedPoles(otherPole, found, poleGroup)
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
    
    ' First Pass: Identify the newest file matching the pattern
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
    
    ' Error handling if no matching token file exists
    If newestFile Is Nothing Then
        MsgBox "No arcgis_token.json file found in Downloads.", vbExclamation
        GetToken = ""
        Exit Function
    End If
    
    ' Second Pass: Delete all older files, skipping the newest one
    For Each file In folder.Files
        If LCase(file.name) Like "arcgis_token*.json" Then
            If LCase(file.name) = "arcgis_token.json" Or file.name Like "*(*).json" Then
                ' If it's not the newest file, delete it permanently
                If file.path <> newestFile.path Then
                    file.Delete True ' True forces deletion of read-only files if necessary
                End If
            End If
        End If
    Next file
    
    ' Read and parse the remaining newest file
    Dim TextStream As Object
    Set TextStream = newestFile.OpenAsTextStream(1, -2) ' 1=ForReading, -2=TristateUseDefault
    
    Dim jsonRaw As String
    jsonRaw = TextStream.ReadAll
    TextStream.Close
    
    Dim jsonParsed As Object
    Set jsonParsed = JsonConverter.ParseJson(jsonRaw)
    
    GetToken = jsonParsed("token")
    
    Dim finalPath As String
    finalPath = downloadsPath & "arcgis_token.json"
    
    If newestFile.path <> finalPath Then
        ' Check if a clean file somehow exists to prevent a runtime collision error
        If fso.FileExists(finalPath) Then fso.DeleteFile finalPath, True
        newestFile.name = "arcgis_token.json"
    End If
End Function

Function testToken(token As String) As Boolean
    Dim url As String: url = "https://gis.consumersenergy.com/mapping/rest/services/Landbase/Landbase_Grids_Boundaries_PUB/MapServer/1/query?where=1=1&outFields=*&returnGeometry=true&geometryType=esriGeometryEnvelope&geometry=&spatialRel=esriSpatialRelIntersects&inSR=2253&outSR=2253+&f=json&token=" & token
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", url, False
    http.setRequestHeader "Authorization", "Bearer " & token
    http.setRequestHeader "Content-Type", "application/json"
    http.send
    If InStr(http.responseText, "Invalid Token") > 0 Then
        testToken = False
    Else
        testToken = True
    End If
End Function

Function injectHotkey() As Boolean
    Dim fileNum As Integer
    Dim fileText As String
    Dim lineBreakPos As Long
    Dim i As Integer
    
    Dim filePath As String: filePath = "C:\ProgramData\Bentley\OpenUtilities Map Connect Edition\Configuration\Organization\USER_APPSETTINGS_DFLTS\Consumers_KeyboardShortcutsSeed.xml"
    
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

Sub ForceInjectModuleToBentley()
    Dim BentleyConnector As Object
    Dim BentleyEngine As Object
    Dim targetProject As Object
    Dim strModuleName As String
    Dim strTempBasPath As String
    Dim strMVBAProjectPath As String
    Dim strProjectNameOnly As String
    Dim fileFound As Boolean
    

    strModuleName = "Test"
    strTempBasPath = "C:\ProgramData\Bentley\OpenUtilities Map Connect Edition\Configuration\WorkSpaces\ConsumersEnergy\Standards\vba\test2.bas"
    strMVBAProjectPath = "C:\ProgramData\Bentley\OpenUtilities Map Connect Edition\Configuration\WorkSpaces\ConsumersEnergy\Standards\vba\CECADReferences.mvba"
    strProjectNameOnly = CreateObject("Scripting.FileSystemObject").GetBaseName(strMVBAProjectPath)

    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents(strModuleName).Export strTempBasPath
    If Err.Number <> 0 Then
        MsgBox "Failed to export Excel module. Go to Excel Trust Center and enable 'Trust access to the VBA project object model'.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    On Error Resume Next
    Set BentleyConnector = GetObject(, "MicroStationDGN.ApplicationObjectConnector")
    If BentleyConnector Is Nothing Then
        MsgBox "Please make sure Bentley OpenUtilities is open and running a DGN file first.", vbExclamation
        Exit Sub
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
        MsgBox "Successfully injected and saved '" & strModuleName & "' directly inside Bentley VBA tree!", vbInformation
    Else
        MsgBox "VBA target project matching '" & strProjectNameOnly & "' was not found active in Bentley's workspace memory.", vbCritical
    End If
    
    Set BentleyEngine = Nothing
    Set BentleyConnector = Nothing
End Sub
