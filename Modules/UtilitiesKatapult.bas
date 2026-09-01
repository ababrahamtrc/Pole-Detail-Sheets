Attribute VB_Name = "UtilitiesKatapult"
Option Explicit

Public Function InitProjectFromKatapultJson(ByVal json As Object) As project
    Dim project As project: Set project = New project
    Dim pole As pole
    Dim jsonNode As Object, jsonPhotoData As Object, jsonWire As Object, jsonArm As Object, jsonInsulator As Object, jsonPoleTag As Object, jsonConnection As Object, jsonGuy As Object
    
    Dim wire As wire
    Dim Arm As Arm
    Dim Insulator As Insulator
    Dim Equipment As Equipment
    Dim guy As guy
    
    Dim nodeKey, photoKey, armKey, insulatorKey, wireKey, equipmentKey, poleTagKey, connectionKey, poleTopKey, guyKey As Variant
    Dim nodeType As String
    
    If json.exists("metadata") Then
        If json("metadata").exists("TRC_tracking_ID") Then project.Notification = Trim(json("metadata")("TRC_tracking_ID"))
        If json("metadata").exists("communication_tracking_ID") Then project.permit = Trim(json("metadata")("communication_tracking_ID"))
        If json("metadata").exists("communication_tracking_ID") Then
            project.township = Trim(json("metadata")("job_city"))
            project.jobLocation = project.township & ", MI"
        End If
        If json("metadata").exists("communication_company") Then project.applicant = Trim(json("metadata")("communication_company"))
        If json("metadata").exists("CE_MKR_fielder") Then project.fielder = Trim(json("metadata")("CE_MKR_fielder"))
    End If

    Dim nodeKeys As Scripting.Dictionary: Set nodeKeys = New Scripting.Dictionary
    Dim insulators As Scripting.Dictionary: Set insulators = New Scripting.Dictionary
    Dim wires As Scripting.Dictionary: Set wires = New Scripting.Dictionary
    
    If json.exists("nodes") Then
        For Each nodeKey In json("nodes").keys
            Set jsonNode = json("nodes")(nodeKey)
            If jsonNode.exists("attributes") Then
                If jsonNode("attributes").exists("node_type") Then
                    nodeType = getFirstValueJson(jsonNode("attributes")("node_type"))
                    If nodeType = "pole" Then
                        Set pole = New pole
                        pole.classVerified = False
                        pole.heightVerified = False
                        
                        If jsonNode("attributes").exists("scid") Then pole.poleNumber = getFirstValueJson(jsonNode("attributes")("scid"))
                        If jsonNode("attributes").exists("hammer_test") Then pole.hammerTestFailed = getFirstValueJson(jsonNode("attributes")("hammer_test")) = "Hammer Fail"
                        If jsonNode("attributes").exists("visual_test") Then pole.visualCheckFailed = getFirstValueJson(jsonNode("attributes")("visual_test")) = "Visual Fail"
                        If jsonNode("attributes").exists("pole_height") Then pole.height = getFirstValueJson(jsonNode("attributes")("pole_height"))
                        If jsonNode("attributes").exists("branded_height") Then pole.heightVerified = getFirstValueJson(jsonNode("attributes")("branded_height"))
                        If jsonNode("attributes").exists("pole_species") Then pole.species = getKatapultNameMapping(getFirstValueJson(jsonNode("attributes")("pole_species")))
                        If jsonNode("attributes").exists("pole_class") Then pole.Class = getFirstValueJson(jsonNode("attributes")("pole_class"))
                        If jsonNode("attributes").exists("branded_class") Then pole.classVerified = getFirstValueJson(jsonNode("attributes")("branded_class"))
                        If jsonNode("attributes").exists("existing_CE_ID_tag") Then pole.existingCEID = getFirstValueJson(jsonNode("attributes")("existing_CE_ID_tag"))
                        If jsonNode("attributes").exists("measured_groundline_circumference") Then pole.glc = getFirstValueJson(jsonNode("attributes")("measured_groundline_circumference"))
                        If jsonNode("attributes").exists("tree_trimming_direction") Then
                            Dim treeKey As Variant
                            For Each treeKey In jsonNode("attributes")("tree_trimming_direction")
                                If pole.notes <> "" Then pole.notes = pole.notes & vbLf
                                pole.notes = pole.notes & "Tree work required " & jsonNode("attributes")("tree_trimming_direction")(treeKey)
                            Next treeKey
                        End If
                        
                        If Utilities.OnlyNumbers(pole.glc) = "" And pole.height <> "" Then pole.glc = autoGLC(pole.height, pole.species, pole.Class)
                        If jsonNode("attributes").exists("note") Then
                            Dim key As Variant
                            For Each key In jsonNode("attributes")("note")
                                If pole.notes <> "" Then pole.notes = pole.notes & vbLf
                                pole.notes = pole.notes & jsonNode("attributes")("note")(key)
                            Next key
                        End If
                        
                        If jsonNode("attributes").exists("county") And project.county = "" Then project.county = Replace(getFirstValueJson(jsonNode("attributes")("county")), " County", "")
                        If jsonNode("attributes").exists("address") Then pole.address = getFirstValueJson(jsonNode("attributes")("address"))
                        
                        If jsonNode.exists("photos") Then
                            photoKey = getMainPhoto(jsonNode)
                            If json("photos").exists(photoKey) Then
                                If json("photos")(photoKey).exists("date_taken") Then
                                    pole.collectedDate = DateAdd("s", json("photos")(photoKey)("date_taken"), #1/1/1970#)
                                End If
                                If json("photos")(photoKey).exists("photofirst_data") Then
                                    Set jsonPhotoData = json("photos")(photoKey)("photofirst_data")
                                    If jsonPhotoData.exists("arm") Then
                                        For Each armKey In jsonPhotoData("arm")
                                            Set jsonArm = jsonPhotoData("arm")(armKey)
                                            If jsonArm.exists("arm_spec") Then
                                                Set Arm = New Arm
                                                Arm.armSpec = jsonArm("arm_spec")
                                                If jsonArm.exists("_children") Then
                                                    If jsonArm("_children").exists("equipment") Then
                                                        For Each equipmentKey In jsonArm("_children")("equipment")
                                                            Set Equipment = UtilitiesKatapult.InitEquipmentFromKatapultJson(jsonArm("_children")("equipment"), CStr(equipmentKey), True)
                                                            Equipment.height = getMesManHeight(jsonArm)
                                                            If Not Equipment Is Nothing Then
                                                                If Equipment.Bonded = "YES" Or Equipment.Bonded = "NO" Then pole.Bonded = Equipment.Bonded
                                                                pole.equipments.Add Equipment
                                                            End If
                                                        Next equipmentKey
                                                    End If
                                                    If jsonArm("_children").exists("wire") Then
                                                       For Each wireKey In jsonArm("_children")("wire")
                                                            Set jsonWire = jsonArm("_children")("wire")(wireKey)
                                                            Set wire = New wire
                                                            wire.height = getMesManHeight(jsonArm)
                                                            wire.modification = wire.height
                                                            If jsonWire.exists("mr_move") Then wire.modification = wire.modification + CInt(jsonWire("mr_move"))
                                                            wire.trace = jsonWire("_trace")
                                                            wire.owner = UCase(json("traces")("trace_data")(wire.trace)("company"))
                                                            wire.componentType = getKatapultNameMapping(json("traces")("trace_data")(wire.trace)("cable_type"))
                                                            If wire.componentType = "NEUT" Then wire.crossArm = Arm.armSpec
                                                            Set Insulator.wire = wire
                                                            pole.wires.Add wire
                                                            Call splitUtilCommWires(wire, pole)
                                                        Next wireKey
                                                    End If
                                                    If jsonArm("_children").exists("insulator") Then
                                                        For Each insulatorKey In jsonArm("_children")("insulator")
                                                            Set jsonInsulator = jsonArm("_children")("insulator")(insulatorKey)
                                                            Set Insulator = New Insulator
                                                            If jsonInsulator.exists("insulator_spec") Then
                                                                Insulator.insulatorSpec = jsonInsulator("insulator_spec")
                                                                If jsonInsulator.exists("_children") Then
                                                                    If jsonInsulator("_children").exists("wire") Then
                                                                        For Each wireKey In jsonInsulator("_children")("wire")
                                                                            Set jsonWire = jsonInsulator("_children")("wire")(wireKey)
                                                                            Set wire = New wire
                                                                            wire.height = getMesManHeight(jsonArm)
                                                                            wire.modification = wire.height
                                                                            If jsonWire.exists("mr_move") Then wire.modification = wire.modification + CInt(jsonWire("mr_move"))
                                                                            wire.trace = jsonWire("_trace")
                                                                            wire.owner = UCase(json("traces")("trace_data")(wire.trace)("company"))
                                                                            wire.componentType = getKatapultNameMapping(json("traces")("trace_data")(wire.trace)("cable_type"))
                                                                            If wire.componentType = "NEUT" Then wire.crossArm = Arm.armSpec
                                                                            Set Insulator.wire = wire
                                                                            pole.wires.Add wire
                                                                            Call splitUtilCommWires(wire, pole)
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
                                    If jsonPhotoData.exists("equipment") Then
                                        For Each equipmentKey In jsonPhotoData("equipment")
                                            Set Equipment = UtilitiesKatapult.InitEquipmentFromKatapultJson(jsonPhotoData("equipment"), CStr(equipmentKey))
                                            If Not Equipment Is Nothing Then
                                                If Equipment.Bonded = "YES" Or Equipment.Bonded = "NO" Then pole.Bonded = Equipment.Bonded
                                                pole.equipments.Add Equipment
                                            End If
                                        Next equipmentKey
                                    End If
                                    If jsonPhotoData.exists("insulator") Then
                                        For Each insulatorKey In jsonPhotoData("insulator")
                                            Set jsonInsulator = jsonPhotoData("insulator")(insulatorKey)
                                            Set Insulator = New Insulator
                                            If jsonInsulator.exists("insulator_spec") Then
                                                Insulator.insulatorSpec = jsonInsulator("insulator_spec")
                                                If jsonInsulator.exists("_children") Then
                                                    If jsonInsulator("_children").exists("wire") Then
                                                        For Each wireKey In jsonInsulator("_children")("wire")
                                                            Set jsonWire = jsonInsulator("_children")("wire")(wireKey)
                                                            Set wire = New wire
                                                            wire.height = getMesManHeight(jsonInsulator)
                                                            wire.modification = wire.height
                                                            If jsonWire.exists("mr_move") Then wire.modification = wire.modification + CInt(jsonWire("mr_move"))
                                                            wire.trace = jsonWire("_trace")
                                                            wire.owner = UCase(json("traces")("trace_data")(wire.trace)("company"))
                                                            wire.componentType = getKatapultNameMapping(json("traces")("trace_data")(wire.trace)("cable_type"))
                                                            Set Insulator.wire = wire
                                                            pole.wires.Add wire
                                                            Call splitUtilCommWires(wire, pole)
                                                        Next wireKey
                                                    End If
                                                End If
                                                pole.insulators.Add Insulator
                                            End If
                                        Next insulatorKey
                                    End If
                                    If jsonPhotoData.exists("pole_top") Then
                                        For Each poleTopKey In jsonPhotoData("pole_top")
                                            pole.agl = getMesManHeight(jsonPhotoData("pole_top")(poleTopKey))
                                        Next poleTopKey
                                    End If
                                    If jsonPhotoData.exists("wire") Then
                                        For Each wireKey In jsonPhotoData("wire")
                                            Set jsonWire = jsonPhotoData("wire")(wireKey)
                                            Set wire = New wire
                                            wire.height = getMesManHeight(jsonWire)
                                            wire.modification = wire.height
                                            If jsonWire.exists("mr_move") Then wire.modification = wire.modification + CInt(jsonWire("mr_move"))
                                            If jsonWire.exists("_trace") Then
                                                wire.trace = jsonWire("_trace")
                                                If json("traces")("trace_data")(wire.trace).exists("company") Then
                                                    wire.owner = UCase(json("traces")("trace_data")(wire.trace)("company"))
                                                End If
                                                wire.componentType = getKatapultNameMapping(json("traces")("trace_data")(wire.trace)("cable_type"))
                                                pole.wires.Add wire
                                                Call splitUtilCommWires(wire, pole)
                                            End If
                                        Next wireKey
                                    End If
                                    
                                    If jsonPhotoData.exists("guying") Then
                                        For Each guyKey In jsonPhotoData("guying")
                                            Set jsonGuy = jsonPhotoData("guying")(guyKey)
                                            If jsonGuy("guying_type") = "Proposed Down Guy" Then
                                                Set guy = New guy
                                                
                                                If jsonGuy.exists("proposed_size") Then guy.proposedSize = jsonGuy("proposed_size")
                                                If jsonGuy.exists("proposed_lead") Then guy.proposedLead = jsonGuy("proposed_lead")
                                                If jsonGuy.exists("proposed_direction") Then guy.proposedDirection = jsonGuy("proposed_direction")
                                                
                                                pole.applicantGuys.Add guy
                                            End If
                                        Next guyKey
                                    End If
                                End If
                            End If
                        End If
                        If jsonNode.exists("latitude") Then pole.latitude = jsonNode("latitude")
                        If jsonNode.exists("longitude") Then pole.longitude = jsonNode("longitude")
                            
                        If jsonNode("attributes").exists("pole_tag") Then
                            For Each poleTagKey In jsonNode("attributes")("pole_tag")
                                Set jsonPoleTag = jsonNode("attributes")("pole_tag")(poleTagKey)
                                If jsonPoleTag.exists("tagtext") Then pole.gisCEID = jsonPoleTag("tagtext")
                                If jsonPoleTag.exists("company") Then pole.owner = jsonPoleTag("company")
                            Next poleTagKey
                        End If
                        
                        project.poles.Add pole
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

    For Each pole In project.poles
        Call pole.setLineStructureTypes
    Next pole

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
    
    Set InitProjectFromKatapultJson = project
End Function

Private Sub addConnections(ByVal json As Object, connectionKey As String, ByVal jsonConnection As Object, nodeKeys As Scripting.Dictionary, nodeType As String, nodeId1 As String, nodeId2 As String)
    Dim pole As pole
    Dim otherPole As pole
    Dim latitude As Double
    Dim longitude As Double
    Dim Span As Span
    Dim trace As String
    Dim wire As wire
    Dim anchor As anchor
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
    
    If nodeKeys.exists(nodeId1) Then
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
            If jsonConnection.exists("sections") Then
                Dim section As Variant
                For Each section In jsonConnection("sections")
                    If jsonConnection("sections")(section).exists("multi_attributes") Then
                        If jsonConnection("sections")(section)("multi_attributes").exists("CE_MKR_tree_trimming") Then
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
                If jsonNode.exists("attributes") Then
                    Set jsonAttributes = jsonNode("attributes")
                    If jsonAttributes.exists("address") Then
                        address = getFirstValueJson(jsonNode("attributes")("address"))
                        Span.houseNumber = Left(address, InStr(address, " "))
                        If Span.houseNumber = "" Then Span.houseNumber = address
                    End If
                End If
            End If
        
            pole.spans.Add Span
               
            If jsonConnection.exists("sections") Then
                For Each sectionKey In jsonConnection("sections")
                    Set jsonSection = jsonConnection("sections")(sectionKey)
                    If jsonSection.exists("photos") Then
                    
                        photoKey = getMainPhoto(jsonSection)
                        
                        If json.exists("photos") Then
                            If json("photos").exists(photoKey) Then
                                Set jsonPhoto = json("photos")(photoKey)
                                If jsonPhoto.exists("photofirst_data") Then
                                    If jsonPhoto("photofirst_data").exists("wire") Then
                                        For Each wireKey In jsonPhoto("photofirst_data")("wire")
                                            Set jsonWire = jsonPhoto("photofirst_data")("wire")(wireKey)
                                            trace = jsonWire("_trace")
                                            Set wire = pole.findWireByTrace(trace, getKatapultNameMapping(jsonWire("wire_spec")))
                                            If Not wire Is Nothing Then
                                                If wire.size <> "" And wire.size <> getKatapultNameMapping(jsonWire("wire_spec")) And wire.size <> "DROP" Then
                                                    height = wire.height
                                                    owner = wire.owner
                                                    componentType = wire.componentType
                                                    Set wire = New wire
                                                    wire.height = height
                                                    wire.trace = trace
                                                    wire.owner = owner
                                                    wire.componentType = componentType
                                                    
                                                    pole.wires.Add wire
                                                    Call splitUtilCommWires(wire, pole)
                                                End If
                                                
                                                If wire.componentType = "PROPOSED" Then
                                                    If jsonWire.exists("diameter") Then
                                                        If wire.diameter = "" Then
                                                            wire.diameter = jsonWire("diameter")
                                                        ElseIf InStr(wire.diameter, jsonWire("diameter")) = 0 Then
                                                            wire.diameter = wire.diameter & ", " & jsonWire("diameter")
                                                        End If
                                                    End If
                                                    If jsonWire.exists("tension") Then wire.tensions.Add Span.spanSlot, jsonWire("tension")
                                                    If jsonWire.exists("mr_move") Then wire.mrMoves.Add Span.spanSlot, jsonWire("mr_move")
                                                End If
                                                
                                                wire.size = getKatapultNameMapping(jsonWire("wire_spec"))
                                                If wire.componentType = "SEC" And isOpenWire(wire.size) Then wire.componentType = "OW"
                                                
                                                If wire.componentType = "SPG" Then
                                                    Set jsonNode = json("nodes")(nodeId2)
                                                    If jsonNode.exists("photos") Then
                                                        photoKey2 = getMainPhoto(jsonNode)
                                                        If json("photos").exists(photoKey2) Then
                                                            If json("photos")(photoKey2).exists("photofirst_data") Then
                                                                Set jsonPhotoData = json("photos")(photoKey2)("photofirst_data")
                                                                If jsonPhotoData.exists("wire") Then
                                                                    For Each wireKey2 In jsonPhotoData("wire")
                                                                        Set jsonWire2 = jsonPhotoData("wire")(wireKey2)
                                                                        If jsonWire2("_trace") = wire.trace Then
                                                                            wire.wepHeight = getMesManHeight(jsonWire2)
                                                                            Exit For
                                                                        End If
                                                                    Next wireKey2
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                                
                                                If wire.midspans.exists(Span.spanSlot) Then
                                                    If wire.midspans(Span.spanSlot) < 1 Or getMesManHeight(jsonWire) < wire.midspans(Span.spanSlot) Then
                                                        Call wire.midspans.Remove(Span.spanSlot)
                                                        If wire.crossArm <> "" And getMesManHeight(jsonWire) = 0 Then
                                                            wire.midspans.Add Span.spanSlot, "XARM"
                                                        Else
                                                            wire.midspans.Add Span.spanSlot, getMesManHeight(jsonWire)
                                                        End If
                                                    End If
                                                Else
                                                    If wire.crossArm <> "" And getMesManHeight(jsonWire) = 0 Then
                                                        wire.midspans.Add Span.spanSlot, "XARM"
                                                    Else
                                                        wire.midspans.Add Span.spanSlot, getMesManHeight(jsonWire)
                                                    End If
                                                    Span.wires.Add wire
                                                    Call splitUtilCommWires(wire, Span)
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
            
            Set anchor = New anchor
            Set jsonNode = json("nodes")(nodeId1)
            Set jsonAnchor = json("nodes")(nodeId2)
            
            Dim lat1 As Double, lat2 As Double, long1 As Double, long2 As Double
            
            lat1 = jsonNode("latitude")
            long1 = jsonNode("longitude")
            lat2 = jsonAnchor("latitude")
            long2 = jsonAnchor("longitude")
            
            result = DistanceAngleFromLatLong(lat1, long1, lat2, long2)
            anchor.distance = result(0)
            anchor.angle = result(1)
            
            Dim anchorOwnerSet As Boolean
            If jsonNode.exists("attributes") Then
                If jsonNode("attributes").exists("company") Then
                    anchor.owner = getFirstValueJson(jsonNode("attributes")("company"))
                    anchorOwnerSet = True
                End If
            End If
            
            photoKey = getMainPhoto(jsonNode)
            
            If photoKey <> "" Then
                Set jsonPhoto = json("photos")(photoKey)
                If jsonPhoto.exists("photofirst_data") Then
                    If jsonPhoto("photofirst_data").exists("guying") Then
                        For Each guyKey In jsonPhoto("photofirst_data")("guying")
                            Set jsonGuy = jsonPhoto("photofirst_data")("guying")(guyKey)
                            If jsonGuy("anchor_id") = nodeId2 Then
                                Set guy = New guy
                                guy.height = getMesManHeight(jsonGuy)
                                trace = jsonGuy("_trace")
                                guy.owner = UCase(json("traces")("trace_data")(trace)("company"))
                                
                                If guy.owner = "CONSUMERS ENERGY" Then anchor.ceCount = anchor.ceCount + 1
                                
                                If Not anchorOwnerSet Then
                                    If anchor.owner <> "" And anchor.owner <> guy.owner Then
                                        highest = True
                                        For Each otherGuy In pole.guys
                                            If otherGuy.height > guy.height Then
                                                highest = False
                                            End If
                                        Next otherGuy
                                        If highest Then anchor.owner = guy.owner
                                    ElseIf anchor.owner = "" Then
                                        anchor.owner = guy.owner
                                    End If
                                End If
                                guy.id = trace
                                If jsonGuy.exists("down_guy_spec") Then
                                    guy.size = getKatapultNameMapping(jsonGuy("down_guy_spec"))
                                ElseIf jsonGuy.exists("wire_spec") Then
                                    guy.size = getKatapultNameMapping(jsonGuy("wire_spec"))
                                End If
                                guy.componentType = "DG"
                                pole.guys.Add guy
                            End If
                        Next guyKey
                        pole.anchors.Add anchor
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Function InitEquipmentFromKatapultJson(ByVal equipments As Object, equipmentKey As String, Optional Arm As Boolean) As Equipment
    Dim otherEquipmentKey As Variant
    
    Dim Equipment As Equipment: Set Equipment = New Equipment
    Dim json, otherJson As Object
    
    Dim measurementType, otherMeasurementType As String
    Dim trace As String
    
    Set json = equipments(equipmentKey)
    Dim katapultComponentType As String
    katapultComponentType = LCase(json("equipment_type"))
    Equipment.componentType = getKatapultNameMapping(katapultComponentType)
    
    Equipment.equipmentId = equipmentKey
    measurementType = ""
    If json.exists("measurement_of") Then measurementType = json("measurement_of")
    If json.exists("CE_MKR_bonded_STL") Then
        If json("CE_MKR_bonded_STL") = "Bonde" Then
            Equipment.Bonded = "YES"
        ElseIf json("CE_MKR_bonded_STL") = "Not Bonded" Then
            Equipment.Bonded = "NO"
        End If
    End If
        
    If Equipment.componentType = "DL" Then
        Equipment.height = getMesManHeight(json)
        Equipment.size = getKatapultNameMapping(json("drip_loop_spec"))
        Equipment.bottomHeight = Equipment.height
    ElseIf Equipment.componentType = "RISER" Then
        Equipment.height = getMesManHeight(json)
        Equipment.size = getKatapultNameMapping(json("riser_spec"))
    ElseIf InStr(measurementType, "bottom") > 0 And (Equipment.componentType = "SL" Or Equipment.componentType = "XFMR") Then
        trace = json("_trace")
    
        Equipment.bottomHeight = getMesManHeight(json)
        Equipment.size = getKatapultNameMapping(json(katapultComponentType & "_spec"))
        For Each otherEquipmentKey In equipments
            If otherEquipmentKey <> equipmentKey Then
                Set otherJson = equipments(otherEquipmentKey)
                If otherJson("_trace") = trace Then
                    otherMeasurementType = ""
                    If otherJson.exists("measurement_of") Then otherMeasurementType = otherJson("measurement_of")
                    If InStr(otherMeasurementType, "top") Then
                        Equipment.height = getMesManHeight(otherJson)
                    End If
                End If
            End If
        Next otherEquipmentKey
    ElseIf InStr(measurementType, "top") > 0 And (Equipment.componentType <> "SL" And Equipment.componentType <> "XFMR") Then
        trace = json("_trace")
        
        Equipment.height = getMesManHeight(json)
        Equipment.size = getKatapultNameMapping(json(katapultComponentType & "_spec"))
    
        For Each otherEquipmentKey In equipments
            If otherEquipmentKey <> equipmentKey Then
                Set otherJson = equipments(otherEquipmentKey)
                If otherJson("_trace") = trace Then
                    otherMeasurementType = ""
                    If otherJson.exists("measurement_of") Then otherMeasurementType = otherJson("measurement_of")
                    If InStr(otherMeasurementType, "bottom") Then
                        Equipment.bottomHeight = getMesManHeight(otherJson)
                    End If
                End If
            End If
        Next otherEquipmentKey
    ElseIf Equipment.componentType <> "SL" And Equipment.componentType <> "XFMR" And Equipment.componentType <> "CAPACITOR" And Equipment.componentType <> "RISER" And Equipment.componentType <> "RECLOSER" And Equipment.componentType <> "REGULATOR" Then
        Equipment.height = getMesManHeight(json)
        Equipment.size = Equipment.componentType
        Equipment.owner = json("company")
    End If
    
    If Not Equipment Is Nothing And Not Arm Then
        If Equipment.height = 0 And Equipment.bottomHeight = 0 Then
            Set Equipment = Nothing
        End If
    End If
    
    Set InitEquipmentFromKatapultJson = Equipment
    
End Function

Private Function getMainPhoto(ByVal json As Object) As String 'json as node id for pole, returns photoKey
    Dim photoKey As Variant
    Dim jsonPhoto As Object
    
    If json.exists("photos") Then
        For Each photoKey In json("photos")
            Set jsonPhoto = json("photos")(photoKey)
            If jsonPhoto.exists("association") Then
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
    If json.exists("_measured_height") Then
        If Not IsNumeric(json("_measured_height")) Then
            getMesManHeight = 0
        Else
            getMesManHeight = json("_measured_height")
        End If
    ElseIf json.exists("_manual_height") Then
        If Not IsNumeric(json("_manual_height")) Then
            getMesManHeight = 0
        Else
            getMesManHeight = json("_manual_height")
        End If
    End If
End Function

Private Sub splitUtilCommWires(wire As wire, poleOrSpan As Object)
    If wire.componentType = "PRI" Or wire.componentType = "NEUT" Or wire.componentType = "SEC" Or wire.componentType = "OW" Or wire.componentType = "TRAFFIC" Or wire.componentType = "SVC" Or (wire.componentType = "SPG" And wire.owner = "CONSUMERS ENERGY") Then
        poleOrSpan.utilWires.Add wire
    ElseIf wire.componentType = "COM" Or wire.componentType = "MSG" Or wire.componentType = "DROP" Or wire.componentType = "PROPOSED" Or (wire.componentType = "SPG" And wire.owner <> "Consumers Energy") Then
        If wire.componentType = "DROP" Then wire.size = "DROP"
        If wire.componentType = "SPG" Then
            wire.componentType = "MSG"
        End If
        poleOrSpan.commComponents.Add wire
        poleOrSpan.commWires.Add wire
    End If
End Sub

Private Function DistanceAngleFromLatLong(lat1 As Double, long1 As Double, lat2 As Double, long2 As Double)
    Const PI As Double = 3.14159265358979
    Const R As Double = 20903520  ' Earth radius in ft
    
    Dim phi1 As Double, phi2 As Double
    Dim dPhi As Double, dLambda As Double
    Dim A As Double, C As Double
    Dim distance As Double
    Dim y As Double, x As Double
    Dim bearing As Double
    
    phi1 = lat1 * PI / 180
    phi2 = lat2 * PI / 180
    dPhi = (lat2 - lat1) * PI / 180
    dLambda = (long2 - long1) * PI / 180
    
    A = (Sin(dPhi / 2) * Sin(dPhi / 2)) + (Cos(phi1) * Cos(phi2) * Sin(dLambda / 2) * Sin(dLambda / 2))
    
    C = 2 * Atn2(Sqr(A), Sqr(1 - A))
    
    distance = R * C
    
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

