Attribute VB_Name = "UtilitiesSpidaCalc"
Public Function InitProjectFromSpidaJson(ByVal json As Object) As project
    Dim project As project: Set project = New project
    
    project.Notification = Trim(json("forms")(1)("fields")("Notification Number"))
    project.permit = Trim(json("forms")(1)("fields")("Permit Number"))
    project.applicant = Trim(json("forms")(1)("fields")("Applicant"))
    project.county = Trim(json("address")("county"))
    project.township = Trim(json("address")("city"))
    project.jobLocation = project.township & ", " & Trim(json("address")("state")) & " " & Trim(json("address")("zip_code"))
    project.fielder = json("engineer")
    
    Dim missingCounter As Integer: missingCounter = 1
    Dim wire As wire
    Dim pole As pole
    Dim lead As Variant
    Dim keys() As Variant
    Dim poleDetails As Scripting.Dictionary
    For Each lead In json("leads")
        Dim jsonPole As Variant
        Set poleDetails = Nothing
        For Each jsonPole In lead("locations")
            For Each form In jsonPole("forms")
                If form("title") = "Pole Details" Then
                    Set poleDetails = form
                    Exit For
                End If
            Next form
            Set pole = New pole
            pole.poleNumber = Trim(Replace(poleDetails("fields")("Pole Number"), "-", ""))
            If pole.poleNumber = "" Then
                pole.poleNumber = "NoPoleNum" & missingCounter
                missingCounter = missingCounter + 1
            End If
            pole.gisCEID = poleDetails("fields")("CE ID Tag")
            pole.existingCEID = poleDetails("fields")("New CE ID Tag")
            pole.address = Trim(jsonPole("address")("number")) & " " & Trim(jsonPole("address")("street")) & ", " & Trim(jsonPole("address")("city")) & ", " & Trim(jsonPole("address")("state")) & " " & Trim(jsonPole("address")("zip_code"))
            pole.notes = jsonPole("comments")
            pole.collectedDate = json("date")
            pole.latitude = jsonPole("geographicCoordinate")("coordinates")(2)
            pole.longitude = jsonPole("geographicCoordinate")("coordinates")(1)
            pole.Bonded = poleDetails("fields")("Bonded Streetlight")
            
            pole.classVerified = poleDetails("fields")("Class Verified") = "true"
            pole.heightVerified = poleDetails("fields")("Height Verified") = "true"
            pole.hammerTestFailed = poleDetails("fields")("Hammer Test") = "Failed (Bad Pole)"
            pole.visualCheckFailed = poleDetails("fields")("Visual Check") = "Deteriorated (Looks Bad)"
            
            Set existing = Nothing
            Set proposed = Nothing
            Set Remedy = Nothing
            
            For Each design In jsonPole("designs")
                If design("label") = "Existing" Then
                    Set existing = design("structure")
                End If
                If design("label") = "Proposed" Then
                    Set proposed = design("structure")
                End If
                If design("label") = "Remedy" Then
                    Set Remedy = design("structure")
                End If
            Next design
            
            Set wiresToSpansDict = New Scripting.Dictionary
            
            If Not existing Is Nothing Then
                industry = existing("pole")("owner")("industry")
                pole.owner = existing("pole")("owner")("id")
                
                pole.species = getSpidaCalcNameMapping(existing("pole")("clientItem")("species"))
                pole.Class = existing("pole")("clientItem")("classOfPole")
                pole.height = CInt(existing("pole")("clientItem")("height")("value") * 3.28084)
                pole.glc = Round(CDbl(existing("pole")("glc")("value") * 3.28084 * 12), 2)
                pole.agl = existing("pole")("agl")("value") * 39.3701
            
                For Each wireEndPoint In existing("wireEndPoints")
                    Dim Span As Span: Set Span = New Span
                    
                    If wireEndPoint("connectionId") <> "" Then
                        Span.spanId = wireEndPoint("connectionId")
                    Else
                        Span.spanId = pole.poleNumber & wireEndPoint("id")
                    End If
                    
                    Span.spanSlot = pole.spans.count + 1
                    
                    Span.distance = wireEndPoint("distance")("value") * 39.3701 / 12
                    Span.angle = wireEndPoint("direction")
                    Span.houseNumber = wireEndPoint("comments")
                    
                    pole.spans.Add Span
                    
                    For Each wep In wireEndPoint("wires")
                        Set wiresToSpansDict(pole.poleNumber & wep) = Span
                    Next wep
                    
                    For Each wep In wireEndPoint("spanGuys")
                        Set wiresToSpansDict(pole.poleNumber & wep) = Span
                    Next wep
                Next wireEndPoint
                
                Dim Equipment As Equipment
                
                For Each jsonEquipment In existing("equipments")
                    Set Equipment = New Equipment
                
                    Equipment.componentType = getSpidaCalcNameMapping(jsonEquipment("clientItem")("type"))
                    Equipment.size = getSpidaCalcNameMapping(jsonEquipment("clientItem")("size"))
                    Equipment.height = jsonEquipment("attachmentHeight")("value") * 39.3701
                    Equipment.bottomHeight = jsonEquipment("bottomHeight")("value") * 39.3701
                    
                    
                    If Equipment.size <> "NOT BONDED" Then pole.equipments.Add Equipment
                Next jsonEquipment
                
                Dim anchor As anchor
                For Each jsonAnchor In existing("anchors")
                    Set anchor = New anchor
                    
                    anchor.distance = jsonAnchor("distance")("value") * 39.3701 / 12
                    anchor.angle = jsonAnchor("direction")
                    anchor.owner = UCase(jsonAnchor("owner")("id"))
                    
                    pole.anchors.Add anchor
                Next jsonAnchor
                
                Dim guy As guy
                For Each jsonGuy In existing("guys")
                    Set guy = New guy
                    
                    guy.size = getSpidaCalcNameMapping(getSpidaCalcNameMapping(jsonGuy("clientItem")("size")))
                    guy.owner = UCase(jsonGuy("owner")("id"))
                    guy.height = jsonGuy("attachmentHeight")("value") * 39.3701
                    
                    pole.guys.Add guy
                Next jsonGuy
                
                For Each jsonSpanguy In existing("spanGuys")
                    Set wire = New wire
                    Set Span = wiresToSpansDict(pole.poleNumber & jsonSpanguy("id"))
                    
                    wire.size = getSpidaCalcNameMapping(jsonSpanguy("clientItem")("size"))
                    wire.owner = UCase(jsonSpanguy("owner")("id"))
                    wire.height = jsonSpanguy("attachmentHeight")("value") * 39.3701
                    wire.wepHeight = jsonSpanguy("height")("value") * 39.3701
                    wire.midspans.Add Span.spanSlot, jsonSpanguy("midspanHeight")("value") * 39.3701
                    
                    If jsonSpanguy("owner")("industry") = "UTILITY" Then
                        wire.componentType = "SPG"
                        Span.utilWires.Add wire
                    Else
                        wire.componentType = "MSG"
                        Span.commWires.Add wire
                    End If
                Next jsonSpanguy
                
                For Each jsonWire In existing("wires")
                    Set wire = New wire
                    Set Span = wiresToSpansDict(pole.poleNumber & jsonWire("id"))
                    
                    wire.owner = UCase(jsonWire("owner")("id"))
                    wire.height = jsonWire("attachmentHeight")("value") * 39.3701
                    wire.midspans.Add Span.spanSlot, jsonWire("midspanHeight")("value") * 39.3701
                    wire.componentType = getSpidaCalcNameMapping(jsonWire("usageGroup"))
                    If wire.componentType = "DROP" Then
                        wire.size = "DROP"
                    Else
                        wire.size = getSpidaCalcNameMapping(jsonWire("clientItem")("size"))
                    End If
                    
                    If jsonWire("owner")("industry") = "UTILITY" Then
                        Span.utilWires.Add wire
                    Else
                        Span.commWires.Add wire
                    End If
                Next jsonWire
            End If
        For Each Span In pole.spans
            Call updateOpenwireMidspans(Span)
        Next Span
        project.poles.Add pole
        Next jsonPole
    Next lead
    
    If project.Notification = "" Then
        project.Notification = InputBox("Enter the Notification:", "User Input")
    End If
    
    If project.permit = "" Then
        project.permit = InputBox("Enter the Permit:", "User Input")
    End If
    
    If project.county = "" Then
        project.county = InputBox("Enter the county and please be exact with no typos, future scripts will care about this:", "User Input")
    End If
    
    If project.fielder = "" Then
        project.fielder = InputBox("Enter the fielder:", "User Input")
    End If
    
    Set InitProjectFromSpidaJson = project
End Function

Private Sub updateOpenwireMidspans(Span As Span)
    Dim openWires As Collection: Set openWires = New Collection
    Dim midspans As Scripting.Dictionary: Set midspans = New Scripting.Dictionary
    Dim midspan As Integer, lowest As Integer
    
    On Error GoTo 0
    
    Dim wire As wire
    Dim wire2 As wire
    
    For Each wire In Span.utilWires
        If wire.componentType = "OW" Then
            midspan = wire.midspans(Span.spanSlot)
            If midspan = 0 Then
                openWires.Add wire
            ElseIf Not midspans.exists(midspan) Then
                near = False
                For Each wire2 In Span.utilWires
                    If wire2.componentType = "OW" Then
                        If Abs(wire.height - wire2.height) < 16 And wire.height <> wire2.height Then near = True
                    End If
                Next wire2
                If near Then
                    midspans.Add midspan, Nothing
                    openWires.Add wire
                End If
            Else
                openWires.Add wire
            End If
        End If
    Next wire
    
    If openWires.count > 0 Then
    
        Set openWires = Utilities.sortComponents(openWires)
    
        If midspans.count = 1 Then
            keys = midspans.keys
            lowest = keys(0)
        ElseIf openWires.count > 0 Then
            lowest = openWires(openWires.count).midspans(Span.spanSlot)
        End If
        
        If midspans.count = 1 And lowest > 0 Then
            For i = 1 To openWires.count
                openWires(i).midspans(Span.spanSlot) = lowest + ((openWires.count - i) * 8)
            Next i
        End If
    End If
End Sub
