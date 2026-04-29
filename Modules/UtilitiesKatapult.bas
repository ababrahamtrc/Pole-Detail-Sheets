Attribute VB_Name = "UtilitiesKatapult"
Option Explicit

Public Function InitProjectFromKatapultJson(ByVal json As Object) As Project
    Dim Project As Project: Set Project = New Project
    Dim pole As pole
    Dim jsonNode As Object, jsonPhotoData As Object, jsonWire As Object, jsonArm As Object, jsonInsulator As Object, jsonPoleTag As Object, jsonConnection As Object, jsonGuy As Object
    
    Dim Wire As Wire
    Dim Arm As Arm
    Dim Insulator As Insulator
    Dim equipment As equipment
    Dim guy As guy
    
    Dim nodeKey, photoKey, armKey, insulatorKey, wireKey, equipmentKey, poleTagKey, connectionKey, poleTopKey, guyKey As Variant
    Dim nodeType As String
    
    If json.Exists("metadata") Then
        If json("metadata").Exists("TRC_tracking_ID") Then Project.Notification = Trim(json("metadata")("TRC_tracking_ID"))
        If json("metadata").Exists("communication_tracking_ID") Then Project.permit = Trim(json("metadata")("communication_tracking_ID"))
        If json("metadata").Exists("communication_tracking_ID") Then
            Project.township = Trim(json("metadata")("job_city"))
            Project.jobLocation = Project.township & ", MI"
        End If
        If json("metadata").Exists("communication_company") Then Project.applicant = Trim(json("metadata")("communication_company"))
        If json("metadata").Exists("CE_MKR_fielder") Then Project.fielder = Trim(json("metadata")("CE_MKR_fielder"))
    End If

    Dim nodeKeys As scripting.Dictionary: Set nodeKeys = New scripting.Dictionary
    Dim insulators As scripting.Dictionary: Set insulators = New scripting.Dictionary
    Dim wires As scripting.Dictionary: Set wires = New scripting.Dictionary
    
    If json.Exists("nodes") Then
        For Each nodeKey In json("nodes").keys
            Set jsonNode = json("nodes")(nodeKey)
            If jsonNode.Exists("attributes") Then
                If jsonNode("attributes").Exists("node_type") Then
                    nodeType = getFirstValueJson(jsonNode("attributes")("node_type"))
                    If nodeType = "pole" Then
                        Set pole = New pole
                        pole.classVerified = False
                        pole.heightVerified = False
                        
                        If jsonNode("attributes").Exists("scid") Then pole.poleNumber = getFirstValueJson(jsonNode("attributes")("scid"))
                        If jsonNode("attributes").Exists("hammer_test") Then pole.hammerTestFailed = getFirstValueJson(jsonNode("attributes")("hammer_test")) = "Hammer Fail"
                        If jsonNode("attributes").Exists("visual_test") Then pole.visualCheckFailed = getFirstValueJson(jsonNode("attributes")("visual_test")) = "Visual Fail"
                        If jsonNode("attributes").Exists("pole_height") Then pole.height = getFirstValueJson(jsonNode("attributes")("pole_height"))
                        If jsonNode("attributes").Exists("branded_height") Then pole.heightVerified = getFirstValueJson(jsonNode("attributes")("branded_height"))
                        If jsonNode("attributes").Exists("pole_species") Then pole.species = getKatapultNameMapping(getFirstValueJson(jsonNode("attributes")("pole_species")))
                        If jsonNode("attributes").Exists("pole_class") Then pole.Class = getFirstValueJson(jsonNode("attributes")("pole_class"))
                        If jsonNode("attributes").Exists("branded_class") Then pole.classVerified = getFirstValueJson(jsonNode("attributes")("branded_class"))
                        If jsonNode("attributes").Exists("existing_CE_ID_tag") Then pole.existingCEID = getFirstValueJson(jsonNode("attributes")("existing_CE_ID_tag"))
                        If jsonNode("attributes").Exists("measured_groundline_circumference") Then pole.glc = getFirstValueJson(jsonNode("attributes")("measured_groundline_circumference"))
                        If jsonNode("attributes").Exists("CE_MKR_tree_trimming") Then pole.treeWork = getFirstValueJson(jsonNode("attributes")("CE_MKR_tree_trimming"))
                        If Utilities.OnlyNumbers(pole.glc) = "" And pole.height <> "" Then pole.glc = autoGLC(pole.height, pole.species, pole.Class)
                        If jsonNode("attributes").Exists("note") Then
                            Dim key As Variant
                            For Each key In jsonNode("attributes")("note")
                                If pole.notes <> "" Then pole.notes = pole.notes & vbLf
                                pole.notes = pole.notes & jsonNode("attributes")("note")(key)
                            Next key
                        End If
                        
                        If jsonNode("attributes").Exists("county") And Project.county = "" Then Project.county = Replace(getFirstValueJson(jsonNode("attributes")("county")), " County", "")
                        If jsonNode("attributes").Exists("address") Then pole.address = getFirstValueJson(jsonNode("attributes")("address"))
                        
                        If jsonNode.Exists("photos") Then
                            photoKey = getMainPhoto(jsonNode)
                            If json("photos").Exists(photoKey) Then
                                If json("photos")(photoKey).Exists("date_taken") Then
                                    pole.collectedDate = DateAdd("s", json("photos")(photoKey)("date_taken"), #1/1/1970#)
                                End If
                                If json("photos")(photoKey).Exists("photofirst_data") Then
                                    Set jsonPhotoData = json("photos")(photoKey)("photofirst_data")
                                    If jsonPhotoData.Exists("arm") Then
                                        For Each armKey In jsonPhotoData("arm")
                                            Set jsonArm = jsonPhotoData("arm")(armKey)
                                            If jsonArm.Exists("arm_spec") Then
                                                Set Arm = New Arm
                                                Arm.armSpec = jsonArm("arm_spec")
                                                If jsonArm.Exists("_children") Then
                                                    If jsonArm("_children").Exists("insulator") Then
                                                        For Each insulatorKey In jsonArm("_children")("insulator")
                                                            Set jsonInsulator = jsonArm("_children")("insulator")(insulatorKey)
                                                            Set Insulator = New Insulator
                                                            If jsonInsulator.Exists("insulator_spec") Then
                                                                Insulator.insulatorSpec = jsonInsulator("insulator_spec")
                                                                If jsonInsulator.Exists("_children") Then
                                                                    If jsonInsulator("_children").Exists("wire") Then
                                                                        For Each wireKey In jsonInsulator("_children")("wire")
                                                                            Set jsonWire = jsonInsulator("_children")("wire")(wireKey)
                                                                            Set Wire = New Wire
                                                                            Wire.height = getMesManHeight(jsonArm)
                                                                            Wire.modification = Wire.height
                                                                            If jsonWire.Exists("mr_move") Then Wire.modification = Wire.modification + CInt(jsonWire("mr_move"))
                                                                            Wire.trace = jsonWire("_trace")
                                                                            Wire.owner = UCase(json("traces")("trace_data")(Wire.trace)("company"))
                                                                            Wire.componentType = getKatapultNameMapping(json("traces")("trace_data")(Wire.trace)("cable_type"))
                                                                            If Wire.componentType = "NEUT" Then Wire.crossArm = Arm.armSpec
                                                                            Set Insulator.Wire = Wire
                                                                            pole.wires.Add Wire
                                                                            Call splitUtilCommWires(Wire, pole)
                                                                        Next wireKey
                                                                    End If
                                                                End If
                                                                Arm.insulators.Add Insulator
                                                            End If
                                                        Next insulatorKey
                                                    End If
                                                End If
                                            End If
                                        Next armKey
                                    End If
                                    If jsonPhotoData.Exists("equipment") Then
                                        For Each equipmentKey In jsonPhotoData("equipment")
                                            Set equipment = UtilitiesKatapult.InitEquipmentFromKatapultJson(jsonPhotoData("equipment"), CStr(equipmentKey))
                                            If Not equipment Is Nothing Then
                                                If equipment.Bonded = "YES" Or equipment.Bonded = "NO" Then pole.Bonded = equipment.Bonded
                                                pole.equipments.Add equipment
                                            End If
                                        Next equipmentKey
                                    End If
                                    If jsonPhotoData.Exists("insulator") Then
                                        For Each insulatorKey In jsonPhotoData("insulator")
                                            Set jsonInsulator = jsonPhotoData("insulator")(insulatorKey)
                                            Set Insulator = New Insulator
                                            If jsonInsulator.Exists("insulator_spec") Then
                                                Insulator.insulatorSpec = jsonInsulator("insulator_spec")
                                                If jsonInsulator.Exists("_children") Then
                                                    If jsonInsulator("_children").Exists("wire") Then
                                                        For Each wireKey In jsonInsulator("_children")("wire")
                                                            Set jsonWire = jsonInsulator("_children")("wire")(wireKey)
                                                            Set Wire = New Wire
                                                            Wire.height = getMesManHeight(jsonInsulator)
                                                            Wire.modification = Wire.height
                                                            If jsonWire.Exists("mr_move") Then Wire.modification = Wire.modification + CInt(jsonWire("mr_move"))
                                                            Wire.trace = jsonWire("_trace")
                                                            Wire.owner = UCase(json("traces")("trace_data")(Wire.trace)("company"))
                                                            Wire.componentType = getKatapultNameMapping(json("traces")("trace_data")(Wire.trace)("cable_type"))
                                                            Set Insulator.Wire = Wire
                                                            pole.wires.Add Wire
                                                            Call splitUtilCommWires(Wire, pole)
                                                        Next wireKey
                                                    End If
                                                End If
                                                pole.insulators.Add Insulator
                                            End If
                                        Next insulatorKey
                                    End If
                                    If jsonPhotoData.Exists("pole_top") Then
                                        For Each poleTopKey In jsonPhotoData("pole_top")
                                            pole.agl = getMesManHeight(jsonPhotoData("pole_top")(poleTopKey))
                                        Next poleTopKey
                                    End If
                                    If jsonPhotoData.Exists("wire") Then
                                        For Each wireKey In jsonPhotoData("wire")
                                            Set jsonWire = jsonPhotoData("wire")(wireKey)
                                            Set Wire = New Wire
                                            Wire.height = getMesManHeight(jsonWire)
                                            Wire.modification = Wire.height
                                            If jsonWire.Exists("mr_move") Then Wire.modification = Wire.modification + CInt(jsonWire("mr_move"))
                                            If jsonWire.Exists("_trace") Then
                                                Wire.trace = jsonWire("_trace")
                                                If json("traces")("trace_data")(Wire.trace).Exists("company") Then
                                                    Wire.owner = UCase(json("traces")("trace_data")(Wire.trace)("company"))
                                                End If
                                                Wire.componentType = getKatapultNameMapping(json("traces")("trace_data")(Wire.trace)("cable_type"))
                                                pole.wires.Add Wire
                                                Call splitUtilCommWires(Wire, pole)
                                            End If
                                        Next wireKey
                                    End If
                                    
                                    If jsonPhotoData.Exists("guying") Then
                                        For Each guyKey In jsonPhotoData("guying")
                                            Set jsonGuy = jsonPhotoData("guying")(guyKey)
                                            If jsonGuy("guying_type") = "Proposed Down Guy" Then
                                                Set guy = New guy
                                                
                                                If jsonGuy.Exists("proposed_size") Then guy.proposedSize = jsonGuy("proposed_size")
                                                If jsonGuy.Exists("proposed_lead") Then guy.proposedLead = jsonGuy("proposed_lead")
                                                If jsonGuy.Exists("proposed_direction") Then guy.proposedDirection = jsonGuy("proposed_direction")
                                                
                                                pole.applicantGuys.Add guy
                                            End If
                                        Next guyKey
                                    End If
                                End If
                            End If
                        End If
                        If jsonNode.Exists("latitude") Then pole.latitude = jsonNode("latitude")
                        If jsonNode.Exists("longitude") Then pole.longitude = jsonNode("longitude")
                            
                        If jsonNode("attributes").Exists("pole_tag") Then
                            For Each poleTagKey In jsonNode("attributes")("pole_tag")
                                Set jsonPoleTag = jsonNode("attributes")("pole_tag")(poleTagKey)
                                If jsonPoleTag.Exists("tagtext") Then pole.gisCEID = jsonPoleTag("tagtext")
                                If jsonPoleTag.Exists("company") Then pole.owner = jsonPoleTag("company")
                            Next poleTagKey
                        End If
                        
                        Project.poles.Add pole
                        nodeKeys.Add nodeKey, pole
                    End If
                End If
            End If
        Next nodeKey
    End If
        
    Dim latitude As Double
    Dim longitude As Double
    Dim Span As Span
    For Each connectionKey In json("connections")
        Set jsonConnection = json("connections")(connectionKey)
        
        Dim nodeId1 As String: nodeId1 = jsonConnection("node_id_1")
        Dim nodeId2 As String: nodeId2 = jsonConnection("node_id_2")
        Dim nodeType1 As String: nodeType1 = getFirstValueJson(json("nodes")(nodeId1)("attributes")("node_type"))
        Dim nodeType2 As String: nodeType2 = getFirstValueJson(json("nodes")(nodeId2)("attributes")("node_type"))

        Call addConnections(json, CStr(connectionKey), jsonConnection, nodeKeys, nodeType2, nodeId1, nodeId2)
        Call addConnections(json, CStr(connectionKey), jsonConnection, nodeKeys, nodeType1, nodeId2, nodeId1)
    Next connectionKey

    For Each pole In Project.poles
        Call pole.setLineStructureTypes
    Next pole

    If Project.county = "" Then
        Project.county = InputBox("Enter the county and please be exact with no typos, future scripts will care about this:", "User Input")
    End If
    
    If Project.fielder = "" Then
        Project.fielder = InputBox("Enter the fielder:", "User Input")
    End If
    
    Set InitProjectFromKatapultJson = Project
End Function

Private Sub addConnections(ByVal json As Object, connectionKey As String, ByVal jsonConnection As Object, nodeKeys As scripting.Dictionary, nodeType As String, nodeId1 As String, nodeId2 As String)
    Dim pole As pole
    Dim otherPole As pole
    Dim latitude As Double
    Dim longitude As Double
    Dim Span As Span
    Dim trace As String
    Dim Wire As Wire
    Dim Anchor As Anchor
    Dim guy As guy
    Dim jsonGuy As Object
    Dim otherGuy As Variant
    Dim result As Variant
    Dim jsonSection, jsonPhoto, jsonWire, jsonWire2, jsonNode, jsonPhotoData, jsonAttributes As Object
    Dim jsonAnchor As Object
    Dim guyKey As Variant
    Dim highest As Boolean
    Dim address As String
    Dim sectionKey, photoKey, photoKey2, wireKey, wireKey2 As Variant
    Dim height As Integer
    Dim owner As String
    Dim componentType As String
    
    If nodeKeys.Exists(nodeId1) Then
        If (nodeType = "pole") Or (nodeType = "building") Or (nodeType = "other pole") Then
            Set pole = nodeKeys(nodeId1)
            latitude = json("nodes")(nodeId2)("latitude")
            longitude = json("nodes")(nodeId2)("longitude")
            result = DistanceAngleFromLatLong(pole.latitude, pole.longitude, latitude, longitude)
            Set Span = New Span
            Span.distance = result(0)
            Span.angle = result(1)
            Span.spanId = connectionKey
            Span.spanSlot = pole.spans.count + 1
            If jsonConnection.Exists("sections") Then
                Dim section As Variant
                For Each section In jsonConnection("sections")
                    If jsonConnection("sections")(section).Exists("multi_attributes") Then
                        If jsonConnection("sections")(section)("multi_attributes").Exists("CE_MKR_tree_trimming") Then
                            Span.treeWork = getFirstValueJson(jsonConnection("sections")(section)("multi_attributes")("CE_MKR_tree_trimming"))
                        End If
                    End If
                Next section
            End If
            'If Not Span.treeWork Then Span.treeWork = getFirstValueJson(jsonConnection("attributes")("CE_MKR_tree_trimming"))
            
            If nodeType = "pole" Then
                Set otherPole = nodeKeys(nodeId2)
                Span.otherPole = otherPole.poleNumber
            End If
            
            If nodeType = "building" Then
                Set jsonNode = json("nodes")(nodeId2)
                If jsonNode.Exists("attributes") Then
                    Set jsonAttributes = jsonNode("attributes")
                    If jsonAttributes.Exists("address") Then
                        address = getFirstValueJson(jsonNode("attributes")("address"))
                        Span.houseNumber = Left(address, InStr(address, " "))
                        If Span.houseNumber = "" Then Span.houseNumber = address
                    End If
                End If
            End If
        
            pole.spans.Add Span
               
            If jsonConnection.Exists("sections") Then
                For Each sectionKey In jsonConnection("sections")
                    Set jsonSection = jsonConnection("sections")(sectionKey)
                    If jsonSection.Exists("photos") Then
                    
                        photoKey = getMainPhoto(jsonSection)
                        
                        If json.Exists("photos") Then
                            If json("photos").Exists(photoKey) Then
                                Set jsonPhoto = json("photos")(photoKey)
                                If jsonPhoto.Exists("photofirst_data") Then
                                    If jsonPhoto("photofirst_data").Exists("wire") Then
                                        For Each wireKey In jsonPhoto("photofirst_data")("wire")
                                            Set jsonWire = jsonPhoto("photofirst_data")("wire")(wireKey)
                                            trace = jsonWire("_trace")
                                            Set Wire = pole.findWireByTrace(trace, getKatapultNameMapping(jsonWire("wire_spec")))
                                            If Not Wire Is Nothing Then
                                                If Wire.size <> "" And Wire.size <> getKatapultNameMapping(jsonWire("wire_spec")) And Wire.size <> "DROP" Then
                                                    height = Wire.height
                                                    owner = Wire.owner
                                                    componentType = Wire.componentType
                                                    Set Wire = New Wire
                                                    Wire.height = height
                                                    Wire.trace = trace
                                                    Wire.owner = owner
                                                    Wire.componentType = componentType
                                                    
                                                    pole.wires.Add Wire
                                                    Call splitUtilCommWires(Wire, pole)
                                                End If
                                                
                                                If Wire.componentType = "PROPOSED" Then
                                                    If jsonWire.Exists("diameter") Then
                                                        If Wire.diameter = "" Then
                                                            Wire.diameter = jsonWire("diameter")
                                                        ElseIf InStr(Wire.diameter, jsonWire("diameter")) = 0 Then
                                                            Wire.diameter = Wire.diameter & ", " & jsonWire("diameter")
                                                        End If
                                                    End If
                                                    If jsonWire.Exists("tension") Then Wire.tensions.Add Span.spanSlot, jsonWire("tension")
                                                    If jsonWire.Exists("mr_move") Then Wire.mrMoves.Add Span.spanSlot, jsonWire("mr_move")
                                                End If
                                                
                                                Wire.size = getKatapultNameMapping(jsonWire("wire_spec"))
                                                If Wire.componentType = "SEC" And isOpenWire(Wire.size) Then Wire.componentType = "OW"
                                                
                                                If Wire.componentType = "SPG" And nodeType = "pole" Then
                                                    Set jsonNode = json("nodes")(nodeId2)
                                                    If jsonNode.Exists("photos") Then
                                                        photoKey2 = getMainPhoto(jsonNode)
                                                        If json("photos").Exists(photoKey2) Then
                                                            If json("photos")(photoKey2).Exists("photofirst_data") Then
                                                                Set jsonPhotoData = json("photos")(photoKey2)("photofirst_data")
                                                                If jsonPhotoData.Exists("wire") Then
                                                                    For Each wireKey2 In jsonPhotoData("wire")
                                                                        Set jsonWire2 = jsonPhotoData("wire")(wireKey2)
                                                                        If jsonWire2("_trace") = Wire.trace Then
                                                                            Wire.wepHeight = getMesManHeight(jsonWire2)
                                                                            Exit For
                                                                        End If
                                                                    Next wireKey2
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                                
                                                If Wire.midspans.Exists(Span.spanSlot) Then
                                                    If Wire.midspans(Span.spanSlot) < 1 Or getMesManHeight(jsonWire) < Wire.midspans(Span.spanSlot) Then
                                                        Call Wire.midspans.Remove(Span.spanSlot)
                                                        If Wire.crossArm <> "" And getMesManHeight(jsonWire) = 0 Then
                                                            Wire.midspans.Add Span.spanSlot, "XARM"
                                                        Else
                                                            Wire.midspans.Add Span.spanSlot, getMesManHeight(jsonWire)
                                                        End If
                                                    End If
                                                Else
                                                    If Wire.crossArm <> "" And getMesManHeight(jsonWire) = 0 Then
                                                        Wire.midspans.Add Span.spanSlot, "XARM"
                                                    Else
                                                        Wire.midspans.Add Span.spanSlot, getMesManHeight(jsonWire)
                                                    End If
                                                    Span.wires.Add Wire
                                                    Call splitUtilCommWires(Wire, Span)
                                                End If
                                            End If
                                        Next wireKey
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next sectionKey
            End If
        ElseIf (nodeType = "existing anchor") Then
            Set pole = nodeKeys(nodeId1)
            
            Set Anchor = New Anchor
            Set jsonNode = json("nodes")(nodeId1)
            Set jsonAnchor = json("nodes")(nodeId2)
            
            Dim lat1 As Double, lat2 As Double, long1 As Double, long2 As Double
            
            lat1 = jsonNode("latitude")
            long1 = jsonNode("longitude")
            lat2 = jsonAnchor("latitude")
            long2 = jsonAnchor("longitude")
            
            result = DistanceAngleFromLatLong(lat1, long1, lat2, long2)
            Anchor.distance = result(0)
            Anchor.angle = result(1)
            
            Dim anchorOwnerSet As Boolean
            If jsonNode.Exists("attributes") Then
                If jsonNode("attributes").Exists("company") Then
                    Anchor.owner = getFirstValueJson(jsonNode("attributes")("company"))
                    anchorOwnerSet = True
                End If
            End If
            
            photoKey = getMainPhoto(jsonNode)
            
            If photoKey <> "" Then
                Set jsonPhoto = json("photos")(photoKey)
                If jsonPhoto.Exists("photofirst_data") Then
                    If jsonPhoto("photofirst_data").Exists("guying") Then
                        For Each guyKey In jsonPhoto("photofirst_data")("guying")
                            Set jsonGuy = jsonPhoto("photofirst_data")("guying")(guyKey)
                            If jsonGuy("anchor_id") = nodeId2 Then
                                Set guy = New guy
                                guy.height = getMesManHeight(jsonGuy)
                                trace = jsonGuy("_trace")
                                guy.owner = UCase(json("traces")("trace_data")(trace)("company"))
                                If Not anchorOwnerSet Then
                                    If Anchor.owner <> "" And Anchor.owner <> guy.owner Then
                                        highest = True
                                        For Each otherGuy In pole.guys
                                            If otherGuy.height > guy.height Then
                                                highest = False
                                            End If
                                        Next otherGuy
                                        If highest Then Anchor.owner = guy.owner
                                    ElseIf Anchor.owner = "" Then
                                        Anchor.owner = guy.owner
                                    End If
                                End If
                                guy.id = trace
                                If jsonGuy.Exists("down_guy_spec") Then
                                    guy.size = getKatapultNameMapping(jsonGuy("down_guy_spec"))
                                ElseIf jsonGuy.Exists("wire_spec") Then
                                    guy.size = getKatapultNameMapping(jsonGuy("wire_spec"))
                                End If
                                guy.componentType = "DG"
                                pole.guys.Add guy
                            End If
                        Next guyKey
                        pole.anchors.Add Anchor
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Function InitEquipmentFromKatapultJson(ByVal equipments As Object, equipmentKey As String) As equipment
    Dim otherEquipmentKey As Variant
    
    Dim equipment As equipment: Set equipment = New equipment
    Dim json, otherJson As Object
    
    Dim measurementType, otherMeasurementType As String
    Dim trace As String
    
    Set json = equipments(equipmentKey)
    Dim katapultComponentType As String
    katapultComponentType = LCase(json("equipment_type"))
    equipment.componentType = getKatapultNameMapping(katapultComponentType)
    
    equipment.equipmentId = equipmentKey
    measurementType = ""
    If json.Exists("measurement_of") Then measurementType = json("measurement_of")
    If json.Exists("CE_MKR_bonded_STL") Then
        If json("CE_MKR_bonded_STL") = "Bonde" Then
            equipment.Bonded = "YES"
        ElseIf json("CE_MKR_bonded_STL") = "Not Bonded" Then
            equipment.Bonded = "NO"
        End If
    End If
        
    If equipment.componentType = "DL" Then
        equipment.height = getMesManHeight(json)
        equipment.size = getKatapultNameMapping(json("drip_loop_spec"))
        equipment.bottomHeight = equipment.height
    ElseIf equipment.componentType = "RISER" Then
        equipment.height = getMesManHeight(json)
        equipment.size = getKatapultNameMapping(json("riser_spec"))
    ElseIf InStr(measurementType, "bottom") > 0 And (equipment.componentType = "SL" Or equipment.componentType = "XFMR") Then
        trace = json("_trace")
    
        equipment.bottomHeight = getMesManHeight(json)
        equipment.size = getKatapultNameMapping(json(katapultComponentType & "_spec"))
        For Each otherEquipmentKey In equipments
            If otherEquipmentKey <> equipmentKey Then
                Set otherJson = equipments(otherEquipmentKey)
                If otherJson("_trace") = trace Then
                    otherMeasurementType = ""
                    If otherJson.Exists("measurement_of") Then otherMeasurementType = otherJson("measurement_of")
                    If InStr(otherMeasurementType, "top") Then
                        equipment.height = getMesManHeight(otherJson)
                    End If
                End If
            End If
        Next otherEquipmentKey
    ElseIf InStr(measurementType, "top") > 0 And (equipment.componentType <> "SL" And equipment.componentType <> "XFMR") Then
        trace = json("_trace")
        
        equipment.height = getMesManHeight(json)
        equipment.size = getKatapultNameMapping(json(katapultComponentType & "_spec"))
    
        For Each otherEquipmentKey In equipments
            If otherEquipmentKey <> equipmentKey Then
                Set otherJson = equipments(otherEquipmentKey)
                If otherJson("_trace") = trace Then
                    otherMeasurementType = ""
                    If otherJson.Exists("measurement_of") Then otherMeasurementType = otherJson("measurement_of")
                    If InStr(otherMeasurementType, "bottom") Then
                        equipment.bottomHeight = getMesManHeight(otherJson)
                    End If
                End If
            End If
        Next otherEquipmentKey
    ElseIf equipment.componentType <> "SL" And equipment.componentType <> "XFMR" And equipment.componentType <> "CAPACITOR" And equipment.componentType <> "RISER" And equipment.componentType <> "RECLOSER" And equipment.componentType <> "REGULATOR" Then
        equipment.height = getMesManHeight(json)
        equipment.size = equipment.componentType
        equipment.owner = json("company")
    End If
    
    If Not equipment Is Nothing Then
        If equipment.height = 0 And equipment.bottomHeight = 0 Then
            Set equipment = Nothing
        End If
    End If
    
    Set InitEquipmentFromKatapultJson = equipment
    
End Function

Private Function getMainPhoto(ByVal json As Object) As String 'json as node id for pole, returns photoKey
    Dim photoKey As Variant
    Dim jsonPhoto As Object
    
    If json.Exists("photos") Then
        For Each photoKey In json("photos")
            Set jsonPhoto = json("photos")(photoKey)
            If jsonPhoto.Exists("association") Then
                If jsonPhoto("association") = "main" Then
                    getMainPhoto = CStr(photoKey)
                    Exit For
                End If
            End If
        Next photoKey
    Else
        getMainPhoto = ""
    End If
End Function

Private Function getFirstValueJson(ByVal json As Object) As String
    Dim key As Variant
    
    For Each key In json
        getFirstValueJson = json(key)
        Exit Function
    Next key
    getFirstValueJson = ""
End Function

Private Function getMesManHeight(ByVal json As Object) As Double
    If json.Exists("_measured_height") Then
        If Not IsNumeric(json("_measured_height")) Then
            getMesManHeight = 0
        Else
            getMesManHeight = json("_measured_height")
        End If
    ElseIf json.Exists("manual_height") Then
        If Not IsNumeric(json("manual_height")) Then
            If InStr(json("manual_height"), "'") And InStr(json("manual_height"), """") > 0 Then
                getMesManHeight = Utilities.convertToInches(json("manual_height"))
            Else
                getMesManHeight = 0
            End If
        Else
            getMesManHeight = json("manual_height") * 12
        End If
    Else
        getMesManHeight = 0
    End If
End Function

Private Sub splitUtilCommWires(Wire As Wire, poleOrSpan As Object)
    If Wire.componentType = "PRI" Or Wire.componentType = "NEUT" Or Wire.componentType = "SEC" Or Wire.componentType = "OW" Or Wire.componentType = "TRAFFIC" Or Wire.componentType = "SVC" Or (Wire.componentType = "SPG" And Wire.owner = "CONSUMERS ENERGY") Then
        poleOrSpan.utilWires.Add Wire
    ElseIf Wire.componentType = "COM" Or Wire.componentType = "MSG" Or Wire.componentType = "DROP" Or Wire.componentType = "PROPOSED" Or (Wire.componentType = "SPG" And Wire.owner <> "Consumers Energy") Then
        If Wire.componentType = "DROP" Then Wire.size = "DROP"
        If Wire.componentType = "SPG" Then
            Wire.componentType = "MSG"
        End If
        poleOrSpan.commComponents.Add Wire
        poleOrSpan.commWires.Add Wire
    End If
End Sub

Private Function DistanceAngleFromLatLong(lat1 As Double, long1 As Double, lat2 As Double, long2 As Double)
    Const PI As Double = 3.14159265358979
    Const R As Double = 20903520  ' Earth radius in ft
    
    Dim phi1 As Double, phi2 As Double
    Dim dPhi As Double, dLambda As Double
    Dim a As Double, c As Double
    Dim distance As Double
    Dim y As Double, x As Double
    Dim bearing As Double
    
    phi1 = lat1 * PI / 180
    phi2 = lat2 * PI / 180
    dPhi = (lat2 - lat1) * PI / 180
    dLambda = (long2 - long1) * PI / 180
    
    a = (Sin(dPhi / 2) * Sin(dPhi / 2)) + (Cos(phi1) * Cos(phi2) * Sin(dLambda / 2) * Sin(dLambda / 2))
    
    c = 2 * Atn2(Sqr(a), Sqr(1 - a))
    
    distance = R * c
    
    y = Sin(dLambda) * Cos(phi2)
    x = Cos(phi1) * Sin(phi2) - Sin(phi1) * Cos(phi2) * Cos(dLambda)
    
    bearing = Atn2(y, x) * 180 / PI
    bearing = (bearing + 360) Mod 360
    
    DistanceAngleFromLatLong = Array(distance, bearing)
End Function

Private Function Atn2(y As Double, x As Double) As Double
    If x = 0 Then
        If y > 0 Then Atn2 = 1.57079632679 Else Atn2 = -1.57079632679
    Else
        Atn2 = Atn(y / x)
        If x < 0 Then Atn2 = Atn2 + 3.14159265359
    End If
End Function

Public Function isOpenWire(size As String) As Boolean
    If size = "4 ACSR" Then isOpenWire = True: Exit Function
    If size = "2 ACSR" Then isOpenWire = True: Exit Function
    If size = "6 CU" Then isOpenWire = True: Exit Function
End Function

