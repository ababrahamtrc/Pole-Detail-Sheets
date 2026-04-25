Attribute VB_Name = "UtilitiesSpidaCalc"
Public Function InitProjectFromSpidaJson(ByVal json As Object) As Project
    Dim Project As Project: Set Project = New Project
    
    Project.Notification = Trim(json("forms")(1)("fields")("Notification Number"))
    Project.permit = Trim(json("forms")(1)("fields")("Permit Number"))
    Project.applicant = Trim(json("forms")(1)("fields")("Applicant"))
    Project.county = Trim(json("address")("county"))
    Project.township = Trim(json("address")("city"))
    Project.jobLocation = Project.township & ", " & Trim(json("address")("state")) & " " & Trim(json("address")("zip_code"))
    Project.fielder = json("engineer")
    
    Dim missingCounter As Integer: missingCounter = 1
    Dim Wire As Wire
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
                    Dim span As span: Set span = New span
                    
                    If wireEndPoint("connectionId") <> "" Then
                        span.spanId = wireEndPoint("connectionId")
                    Else
                        span.spanId = pole.poleNumber & wireEndPoint("id")
                    End If
                    
                    span.spanSlot = pole.spans.count + 1
                    
                    span.distance = wireEndPoint("distance")("value") * 39.3701 / 12
                    span.angle = wireEndPoint("direction")
                    span.houseNumber = wireEndPoint("comments")
                    
                    pole.spans.Add span
                    
                    For Each wep In wireEndPoint("wires")
                        Set wiresToSpansDict(pole.poleNumber & wep) = span
                    Next wep
                    
                    For Each wep In wireEndPoint("spanGuys")
                        Set wiresToSpansDict(pole.poleNumber & wep) = span
                    Next wep
                Next wireEndPoint
                
                Dim equipment As equipment
                
                For Each jsonEquipment In existing("equipments")
                    Set equipment = New equipment
                
                    equipment.componentType = getSpidaCalcNameMapping(jsonEquipment("clientItem")("type"))
                    equipment.size = getSpidaCalcNameMapping(jsonEquipment("clientItem")("size"))
                    equipment.height = jsonEquipment("attachmentHeight")("value") * 39.3701
                    equipment.bottomHeight = jsonEquipment("bottomHeight")("value") * 39.3701
                    
                    
                    If equipment.size <> "NOT BONDED" Then pole.equipments.Add equipment
                Next jsonEquipment
                
                Dim Anchor As Anchor
                For Each jsonAnchor In existing("anchors")
                    Set Anchor = New Anchor
                    
                    Anchor.distance = jsonAnchor("distance")("value") * 39.3701 / 12
                    Anchor.angle = jsonAnchor("direction")
                    Anchor.owner = UCase(jsonAnchor("owner")("id"))
                    
                    pole.anchors.Add Anchor
                Next jsonAnchor
                
                Dim guy As guy
                For Each jsonGuy In existing("guys")
                    Set guy = New guy
                    
                    guy.size = getSpidaCalcNameMapping(getSpidaCalcNameMapping(jsonGuy("clientItem")("size")))
                    guy.owner = UCase(jsonGuy("owner")("id"))
                    guy.height = jsonGuy("attachmentHeight")("value") * 39.3701
                    
                    pole.guys.Add guy
                Next jsonGuy
                
                For Each jsonSpanGuy In existing("spanGuys")
                    Set Wire = New Wire
                    Set span = wiresToSpansDict(pole.poleNumber & jsonSpanGuy("id"))
                    
                    Wire.size = getSpidaCalcNameMapping(jsonSpanGuy("clientItem")("size"))
                    Wire.owner = UCase(jsonSpanGuy("owner")("id"))
                    Wire.height = jsonSpanGuy("attachmentHeight")("value") * 39.3701
                    Wire.wepHeight = jsonSpanGuy("height")("value") * 39.3701
                    Wire.midspans.Add span.spanSlot, jsonSpanGuy("midspanHeight")("value") * 39.3701
                    
                    If jsonSpanGuy("owner")("industry") = "UTILITY" Then
                        Wire.componentType = "SPG"
                        span.utilWires.Add Wire
                    Else
                        Wire.componentType = "MSG"
                        span.commWires.Add Wire
                    End If
                Next jsonSpanGuy
                
                For Each jsonWire In existing("wires")
                    Set Wire = New Wire
                    Set span = wiresToSpansDict(pole.poleNumber & jsonWire("id"))
                    
                    Wire.owner = UCase(jsonWire("owner")("id"))
                    Wire.height = jsonWire("attachmentHeight")("value") * 39.3701
                    Wire.midspans.Add span.spanSlot, jsonWire("midspanHeight")("value") * 39.3701
                    Wire.componentType = getSpidaCalcNameMapping(jsonWire("usageGroup"))
                    If Wire.componentType = "DROP" Then
                        Wire.size = "DROP"
                    Else
                        Wire.size = getSpidaCalcNameMapping(jsonWire("clientItem")("size"))
                    End If
                    
                    If jsonWire("owner")("industry") = "UTILITY" Then
                        span.utilWires.Add Wire
                    Else
                        span.commWires.Add Wire
                    End If
                Next jsonWire
            End If
        For Each span In pole.spans
            Call updateOpenwireMidspans(span)
        Next span
        Project.poles.Add pole
        Next jsonPole
    Next lead
    
    If Project.county = "" Then
        Project.county = InputBox("Enter the county and please be exact with no typos, future scripts will care about this:", "User Input")
    End If
    
    If Project.fielder = "" Then
        Project.fielder = InputBox("Enter the fielder:", "User Input")
    End If
    
    Set InitProjectFromSpidaJson = Project
End Function

Private Sub updateOpenwireMidspans(span As span)
    Dim openWires As Collection: Set openWires = New Collection
    Dim midspans As Scripting.Dictionary: Set midspans = New Scripting.Dictionary
    Dim midspan As Integer, lowest As Integer
    
    On Error GoTo 0
    
    Dim Wire As Wire
    Dim wire2 As Wire
    
    For Each Wire In span.utilWires
        If Wire.componentType = "OW" Then
            midspan = Wire.midspans(span.spanSlot)
            If midspan = 0 Then
                openWires.Add Wire
            ElseIf Not midspans.Exists(midspan) Then
                near = False
                For Each wire2 In span.utilWires
                    If wire2.componentType = "OW" Then
                        If Abs(Wire.height - wire2.height) < 16 And Wire.height <> wire2.height Then near = True
                    End If
                Next wire2
                If near Then
                    midspans.Add midspan, Nothing
                    openWires.Add Wire
                End If
            Else
                openWires.Add Wire
            End If
        End If
    Next Wire
    
    If openWires.count > 0 Then
    
        Set openWires = Utilities.sortComponents(openWires)
    
        If midspans.count = 1 Then
            keys = midspans.keys
            lowest = keys(0)
        ElseIf openWires.count > 0 Then
            lowest = openWires(openWires.count).midspans(span.spanSlot)
        End If
        
        If midspans.count = 1 And lowest > 0 Then
            For i = 1 To openWires.count
                openWires(i).midspans(span.spanSlot) = lowest + ((openWires.count - i) * 8)
            Next i
        End If
    End If
End Sub
