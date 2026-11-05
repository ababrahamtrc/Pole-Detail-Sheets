VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Project"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public notification As String
Public Permit As String
Public applicant As String
Public jobLocation As String
Public county As String
Public township As String
Public fielder As String
Public dataCollectionDate As String
Public poles As Collection
Public mode As String

Private Sub Class_Initialize()
    Set poles = New Collection
    Me.mode = ThisWorkbook.sheets("Control").Range("MODE").text
End Sub

Public Sub extractFromSheets(Optional printMode As Boolean)
    Dim sheet As Worksheet
    Dim pole As pole

    For Each sheet In ThisWorkbook.sheets
        If sheet.name <> "4 Spans" And sheet.name <> "8 Spans" And sheet.name <> "12 Spans" And sheet.cells(2, 2).value = "Notification:" Then
            Set pole = New pole
            Call pole.extractFromSheet(sheet)
            Me.poles.Add pole
            
            If Me.township = "" Then Me.township = sheet.Range("TWP")
            If Me.county = "" Then Me.county = sheet.Range("COUNTY")
            If Me.applicant = "" Then Me.applicant = sheet.Range("APPLICANT")
            If Me.notification = "" Then Me.notification = sheet.Range("NOTIFICATION")
            If Me.Permit = "" Then Me.Permit = sheet.Range("PERMIT")
            If Me.fielder = "" Then Me.fielder = sheet.Range("INSPECTEDBY")
            If Me.dataCollectionDate = "" Then Me.dataCollectionDate = sheet.Range("DATE").value
        End If
    Next sheet
    
    If printMode Then
        Dim Span As Span
        Dim otherPole As pole
        Dim otherSpan As Span
        For Each pole In Me.poles
            If pole.latitude = 0 Or pole.longitude = 0 Then
                For Each Span In pole.spans
                    Set otherPole = Me.findPole(Span.otherPole)
                    If Not otherPole Is Nothing Then
                        For Each otherSpan In otherPole.spans
                            If otherSpan.otherPole = pole.poleNumber Then
                                If otherPole.latitude <> 0 And otherPole.longitude <> 0 Then
                                    r = 20902231#
                                    PI = 4# * Atn(1#)
                                    
                                    radLat = otherPole.latitude * PI / 180#
                                    radLon = otherPole.longitude * PI / 180#
                                    radAngle = otherSpan.angle * PI / 180#
                                    
                                    dLat = (otherSpan.distance * Cos(radAngle)) / r
                                    dLon = (otherSpan.distance * Sin(radAngle)) / (r * Cos(radLat))
                                    
                                    pole.latitude = otherPole.latitude + (dLat * 180# / PI)
                                    pole.longitude = otherPole.longitude + (dLon * 180# / PI)
                                    results = LatLonToMI2253(pole.latitude, pole.longitude)
                                    pole.x = results(0)
                                    pole.y = results(1)
                                End If
                            End If
                        Next otherSpan
                    End If
                Next Span
            End If
        Next pole
        
        Dim otherSpan2 As Span
        Dim otherSpan3 As Span
        Dim otherPole2 As pole
        Dim otherPole3 As pole
        Dim wire As wire
        Dim angleOffsetLimit: angleOffsetLimit = 15
        For Each pole In Me.poles
            If pole.primaries.count = 0 Then
                For Each Span In pole.spans
                    Set otherPole = Me.findPole(Span.otherPole)
                    If Not otherPole Is Nothing Then
                        If otherPole.primaries.count > 0 Then
                            For Each otherSpan In otherPole.spans
                                For Each wire In otherSpan.utilWires
                                    If wire.componentType = "PRI" Then
                                        angle1 = Span.angle
                                        angle2 = otherSpan.angle
                                        If angle1 > angle2 Then
                                            angleDifference = angle1 - angle2
                                        Else
                                            angleDifference = angle2 - angle1
                                        End If
                                        If Abs(angleDifference - 180) < angleOffsetLimit Then
                                            pole.skipSpan = True
                                            Exit For
                                        End If
                                    End If
                                Next wire
                                If pole.skipSpan Then Exit For
                            Next otherSpan
                        Else
                            For Each otherSpan In otherPole.spans
                                angle1 = Span.angle
                                angle2 = otherSpan.angle
                                If Abs(angle1 - angle2) < angleOffsetLimit Then
                                    Set otherPole2 = Me.findPole(otherSpan.otherPole)
                                    If Not otherPole2 Is Nothing Then
                                        If otherPole2.primaries.count > 0 Then
                                            For Each otherSpan2 In otherPole2.spans
                                                For Each wire In otherSpan2.utilWires
                                                    If wire.componentType = "PRI" Then
                                                        angle1 = otherSpan.angle
                                                        angle2 = otherSpan2.angle
                                                        If angle1 > angle2 Then
                                                            angleDifference = angle1 - angle2
                                                        Else
                                                            angleDifference = angle2 - angle1
                                                        End If
                                                        If Abs(angleDifference - 180) < angleOffsetLimit Then
                                                            pole.skipSpan = True
                                                            Exit For
                                                        End If
                                                    End If
                                                Next wire
                                                If pole.skipSpan Then Exit For
                                            Next otherSpan2
                                        Else
                                            For Each otherSpan2 In otherPole2.spans
                                                angle1 = otherSpan.angle
                                                angle2 = otherSpan2.angle
                                                If Abs(angle1 - angle2) < angleOffsetLimit Then
                                                    Set otherPole3 = Me.findPole(otherSpan2.otherPole)
                                                    If Not otherPole3 Is Nothing Then
                                                        If otherPole3.primaries.count > 0 Then
                                                            For Each otherSpan3 In otherPole3.spans
                                                                For Each wire In otherSpan3.utilWires
                                                                    If wire.componentType = "PRI" Then
                                                                        angle1 = otherSpan2.angle
                                                                        angle2 = otherSpan3.angle
                                                                        If angle1 > angle2 Then
                                                                            angleDifference = angle1 - angle2
                                                                        Else
                                                                            angleDifference = angle2 - angle1
                                                                        End If
                                                                        If Abs(angleDifference - 180) < angleOffsetLimit Then
                                                                            pole.skipSpan = True
                                                                            Exit For
                                                                        End If
                                                                    End If
                                                                Next wire
                                                                If pole.skipSpan Then Exit For
                                                            Next otherSpan3
                                                        End If
                                                    End If
                                                End If
                                                If pole.skipSpan Then Exit For
                                            Next otherSpan2
                                        End If
                                    End If
                                End If
                                If pole.skipSpan Then Exit For
                            Next otherSpan
                        End If
                    End If
                    If pole.skipSpan Then Exit For
                Next Span
            End If
        Next pole
    End If
End Sub

Public Sub createAndFillSheets()
    Dim pole As pole
    Dim totalPoles As Integer
    Dim poleCount As Integer
    
    For Each pole In Me.poles
        If Not Utilities.SheetExists(pole.poleNumber) Then totalPoles = totalPoles + 1
    Next pole
    
    ProgressBar_Form.Show vbModeless
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim sheetsCreated As Scripting.Dictionary: Set sheetsCreated = New Scripting.Dictionary
    
    For i = 1 To Me.poles.count
        Set pole = Me.poles(i)
        If Not Utilities.SheetExists(pole.poleNumber) Then
            sheetsCreated.Add pole.poleNumber, Nothing
            poleCount = poleCount + 1
            ProgressBar_Form.Label1.caption = "Creating Pole Detail Sheet " & poleCount & " of " & totalPoles & " ... please wait..."
            ProgressBar_Form.Repaint
            Call pole.createSheet(Me)
        End If
    Next i
    
    
    Dim sheet As Worksheet
    Dim Span As Span
    For Each sheet In ThisWorkbook.sheets
        If Utilities.IsPDS(sheet) And sheetsCreated.exists(sheet.name) Then
            Call RemedyGen.calculateProposedMidspans(sheet)
            Set pole = Me.findPole(sheet.name)
            For Each Span In pole.spans
                If pole.applicant.mrMoves.exists(Span.spanSlot) And Span.otherPole <> "" Then
                    Dim lines() As String
                    lines = Split(sheet.Range("CMRF2").value, vbLf)
                    For j = LBound(lines) To UBound(lines)
                        If InStr(lines(j), " @P" & Span.otherPole) > 0 Then
                            lines(j) = Utilities.inchesToFeetInches(Utilities.convertToInches(lines(j)) + CInt(pole.applicant.mrMoves(Span.spanSlot))) & " @P" & Span.otherPole
                        End If
                    Next j
                    sheet.Range("CMRF2").value = Join(lines, vbLf)
                End If
            Next Span
        End If
    Next sheet
    
    ProgressBar_Form.Hide
    
    ThisWorkbook.sheets("Control").Activate
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Public Function findPole(text As String) As pole
    Dim pole As pole
    For Each pole In Me.poles
        If pole.poleNumber = text Then
            Set findPole = pole
            Exit Function
        End If
    Next pole
    
    Set findPole = Nothing
End Function

Public Sub extractImportDataFormat()
    Set jobWs = ThisWorkbook.sheets("Job Info")
    jobWsLastRow = jobWs.cells(jobWs.Rows.count, "A").End(xlUp).ROW

    Set colWs = ThisWorkbook.sheets("Collection")
    colWsLastRow = colWs.cells(colWs.Rows.count, "A").End(xlUp).ROW
    
    Set spanWs = ThisWorkbook.sheets("Span")
    spanWsLastRow = spanWs.cells(spanWs.Rows.count, "A").End(xlUp).ROW
    spanWsRow = 2
    
    Set spanComWs = ThisWorkbook.sheets("Span.Communication")
    spanComWsLastRow = spanComWs.cells(spanComWs.Rows.count, "A").End(xlUp).ROW
    spanComWsRow = 2
    
    Set spanPowWs = ThisWorkbook.sheets("Span.Power Circuit")
    spanPowWsLastRow = spanPowWs.cells(spanPowWs.Rows.count, "A").End(xlUp).ROW
    spanPowWsRow = 2
    
    Set anchorWs = ThisWorkbook.sheets("Anchor")
    anchorWsLastRow = anchorWs.cells(anchorWs.Rows.count, "A").End(xlUp).ROW
    anchorWsRow = 2
    
    Set guyWs = ThisWorkbook.sheets("Anchor.Guys")
    guyWsLastRow = guyWs.cells(guyWs.Rows.count, "A").End(xlUp).ROW
    guyWsRow = 2
    
    Set equipmentWs = ThisWorkbook.sheets("Equipment")
    equipmentWsLastRow = equipmentWs.cells(equipmentWs.Rows.count, "A").End(xlUp).ROW
    equipmentWsRow = 2
    
    Me.jobLocation = jobWs.cells(2, 2)
    Me.Permit = jobWs.cells(2, 3)
    Me.applicant = jobWs.cells(2, 4)
    Me.notification = jobWs.cells(2, 5)
    Me.county = jobWs.cells(2, 6)
    Me.township = jobWs.cells(2, 7)
    Me.fielder = jobWs.cells(2, 8)
    
    For colWsRow = 2 To colWsLastRow
        Dim pole As pole: Set pole = New pole
        
        pole.poleNumber = colWs.cells(colWsRow, 1)
        pole.gisCEID = colWs.cells(colWsRow, 2)
        pole.existingCEID = colWs.cells(colWsRow, 3)
        If pole.existingCEID = "" Then pole.existingCEID = pole.gisCEID
        parts = Split(colWs.cells(colWsRow, 4), " > ")
        pole.species = parts(0)
        pole.Class = parts(1)
        pole.height = parts(2)
        pole.classVerified = colWs.cells(colWsRow, 5).value
        pole.heightVerified = colWs.cells(colWsRow, 6).value
        pole.owner = colWs.cells(colWsRow, 7).value
        pole.glc = colWs.cells(colWsRow, 8).value
        pole.address = colWs.cells(colWsRow, 9).value
        pole.hammerTestFailed = colWs.cells(colWsRow, 10).value
        pole.visualCheckFailed = colWs.cells(colWsRow, 11).value
        pole.notes = colWs.cells(colWsRow, 12).value
        pole.collectedDate = colWs.cells(colWsRow, 13).value
        pole.latitude = colWs.cells(colWsRow, 14).value
        pole.longitude = colWs.cells(colWsRow, 15).value
        pole.agl = colWs.cells(colWsRow, 16).value
        pole.Bonded = colWs.cells(colWsRow, 17).value
        pole.treeWork = colWs.cells(colWsRow, 18).value
        
        'Span section
        Dim slotCount As Integer: slotCount = 0
        For spanWsRow = spanWsRow To spanWsLastRow
            If spanWs.cells(spanWsRow, 1) <> pole.poleNumber Then Exit For
            
            Dim Span As Span: Set Span = New Span
            
            Span.spanId = spanWs.cells(spanWsRow, 2)
            Span.distance = spanWs.cells(spanWsRow, 3)
            Span.angle = spanWs.cells(spanWsRow, 4)
            Span.direction = Utilities.degreesToDirection(Span.angle)
            Span.houseNumber = spanWs.cells(spanWsRow, 5)
            Span.treeWork = spanWs.cells(spanWsRow, 6)
            
            For j = 2 To spanWsLastRow
                If j <> spanWsRow Then
                    If spanWs.cells(j, 2) = Span.spanId Then
                        Span.otherPole = spanWs.cells(j, 1)
                        Exit For
                    End If
                End If
            Next j
            
            slotCount = slotCount + 1
            Span.spanSlot = slotCount
            
            pole.spans.Add Span
        Next spanWsRow
        
        'Communication wire section
        Dim wireDict As Scripting.Dictionary: Set wireDict = New Scripting.Dictionary
        For spanComWsRow = spanComWsRow To spanComWsLastRow
            If spanComWs.cells(spanComWsRow, 1) <> pole.poleNumber Then Exit For
            
            Set Span = pole.findSpanBySpanId(spanComWs.cells(spanComWsRow, 2))
            size = spanComWs.cells(spanComWsRow, 3)
            owner = spanComWs.cells(spanComWsRow, 4)
            height = spanComWs.cells(spanComWsRow, 5)
            componentType = spanComWs.cells(spanComWsRow, 7)
            modification = spanComWs.cells(spanComWsRow, 8)
            wepHeight = spanComWs.cells(spanComWsRow, 12)
            key = size & owner & height & componentType
            
            Dim wire As wire
            If Not wireDict.exists(key) Then
                Set wire = New wire
            
                wire.size = size
                wire.owner = owner
                wire.height = height
                wire.componentType = componentType
                wire.modification = modification
                wire.wepHeight = wepHeight
                
                wireDict.Add key, wire
            Else
                Continue = True
                Do While Continue
                    If wireDict(key).midspans.exists(Span.spanSlot) Then
                        key = key & "A"
                        If Not wireDict.exists(key) Then
                            Set wire = New wire
                
                            wire.size = size
                            wire.owner = owner
                            wire.height = height
                            wire.componentType = componentType
                            wire.modification = modification
                            wire.wepHeight = wepHeight
                            
                            wireDict.Add key, wire
                            
                            Continue = False
                        End If
                    Else
                        Continue = False
                    End If
                Loop
            End If
            
            Set wire = wireDict(key)
            wire.midspans.Add Span.spanSlot, spanComWs.cells(spanComWsRow, 6)
            If spanComWs.cells(spanComWsRow, 9) <> "" Then wire.diameter = spanComWs.cells(spanComWsRow, 9)
            If spanComWs.cells(spanComWsRow, 10) <> "" Then wire.tensions.Add Span.spanSlot, spanComWs.cells(spanComWsRow, 10)
            If spanComWs.cells(spanComWsRow, 11) <> "" Then wire.mrMoves.Add Span.spanSlot, spanComWs.cells(spanComWsRow, 11)
            
            Span.wires.Add wire
            Span.commWires.Add wire
        Next spanComWsRow
        
        For Each key In wireDict
            Set wire = wireDict(key)
        
            If wire.componentType <> "PROPOSED" Then
                pole.wires.Add wire
                pole.commWires.Add wire
                pole.commComponents.Add wire
            Else
                Set pole.applicant = wire
            End If
        Next key
        
        'Utility Wire Section
        Set wireDict = New Scripting.Dictionary
        For spanPowWsRow = spanPowWsRow To spanPowWsLastRow
            If spanPowWs.cells(spanPowWsRow, 1) <> pole.poleNumber Then Exit For
            
            Set Span = pole.findSpanBySpanId(spanPowWs.cells(spanPowWsRow, 2))
            size = spanPowWs.cells(spanPowWsRow, 3)
            height = spanPowWs.cells(spanPowWsRow, 4)
            componentType = spanPowWs.cells(spanPowWsRow, 6)
            wepHeight = spanPowWs.cells(spanPowWsRow, 7)
            key = size & height & componentType
            
            If Not wireDict.exists(key) Then
                Set wire = New wire
            
                wire.size = size
                wire.height = height
                wire.componentType = componentType
                wire.wepHeight = wepHeight
                
                wireDict.Add key, wire
            Else
                Continue = True
                Do While Continue
                    If wireDict(key).midspans.exists(Span.spanSlot) Then
                        key = key & "A"
                        If Not wireDict.exists(key) Then
                            Set wire = New wire
                
                            wire.size = size
                            wire.height = height
                            wire.componentType = componentType
                            wire.wepHeight = wepHeight
                            
                            wireDict.Add key, wire
                            
                            Continue = False
                        End If
                    Else
                        Continue = False
                    End If
                Loop
            End If
            
            Set wire = wireDict(key)
            wire.midspans.Add Span.spanSlot, spanPowWs.cells(spanPowWsRow, 5)
            
            Span.wires.Add wire
            Span.utilWires.Add wire
        Next spanPowWsRow
        
        For Each key In wireDict
            Set wire = wireDict(key)
        
            pole.wires.Add wire
            pole.utilWires.Add wire
            
            If wire.componentType <> "PRI" Then pole.utilComponents.Add wire
        Next key
        
        'Anchor section
        For anchorWsRow = anchorWsRow To anchorWsLastRow
            If anchorWs.cells(anchorWsRow, 1) <> pole.poleNumber Then Exit For
        
            Dim anchor As anchor: Set anchor = New anchor
            
            anchor.distance = anchorWs.cells(anchorWsRow, 2)
            anchor.angle = anchorWs.cells(anchorWsRow, 3)
            anchor.direction = Utilities.degreesToDirection(anchor.angle)
            anchor.owner = anchorWs.cells(anchorWsRow, 4)
            anchor.ceCount = anchorWs.cells(anchorWsRow, 5)
            
            pole.anchors.Add anchor
        Next anchorWsRow
        
        'Guy section
        For guyWsRow = guyWsRow To guyWsLastRow
            If guyWs.cells(guyWsRow, 1) <> pole.poleNumber Then Exit For
        
            Dim guy As guy: Set guy = New guy
            
            If guyWs.cells(guyWsRow, 4) <> "" Then
                guy.size = guyWs.cells(guyWsRow, 2)
                parts = Split(guyWs.cells(guyWsRow, 3), " > ")
                ownerType = parts(0)
                guy.owner = parts(1)
                guy.height = guyWs.cells(guyWsRow, 4)
                guy.componentType = "DG"
                
                pole.guys.Add guy
                If ownerType = "UTILITY" Then
                    pole.utilComponents.Add guy
                Else
                    pole.commComponents.Add guy
                End If
            ElseIf guyWs.cells(guyWsRow, 5) <> "" Then
                guy.proposedSize = guyWs.cells(guyWsRow, 5)
                guy.proposedLead = guyWs.cells(guyWsRow, 6)
                guy.proposedDirection = guyWs.cells(guyWsRow, 7)
                
                pole.applicantGuys.Add guy
            End If
        Next guyWsRow

        'Equipment section
        For equipmentWsRow = equipmentWsRow To equipmentWsLastRow
            If equipmentWs.cells(equipmentWsRow, 1) <> pole.poleNumber Then Exit For
            
            Dim Equipment As Equipment: Set Equipment = New Equipment
            
            parts = Split(equipmentWs.cells(equipmentWsRow, 2), " > ")
            Equipment.componentType = parts(0)
            Equipment.size = parts(1)
            Equipment.height = equipmentWs.cells(equipmentWsRow, 3)
            Equipment.bottomHeight = equipmentWs.cells(equipmentWsRow, 4)
            Equipment.owner = equipmentWs.cells(equipmentWsRow, 5)
            
            pole.equipments.Add Equipment
            If Equipment.owner = "" Or UCase(Equipment.owner) = "CONSUMERS ENERGY" Then
                pole.utilComponents.Add Equipment
            Else
                pole.commComponents.Add Equipment
            End If
        Next equipmentWsRow
        
        Me.poles.Add pole
    Next colWsRow
End Sub

Public Sub fillImportDataFormat()
    Dim pole As pole
    Dim Span As Span
    Dim wire As wire
    Dim anchor As anchor
    Dim guy As guy
    Dim Equipment As Equipment
    Dim primaryBundles As Scripting.Dictionary
    
    Set colWs = ThisWorkbook.sheets("Collection")
    colWs.cells(1, 1).value = "ID"
    colWs.cells(1, 2).value = "GIS CE ID Tag"
    colWs.cells(1, 3).value = "Existing CE ID Tag"
    colWs.cells(1, 4).value = "Type"
    colWs.cells(1, 5).value = "Pole Class Verified"
    colWs.cells(1, 6).value = "Pole Height Verified"
    colWs.cells(1, 7).value = "Owner"
    colWs.cells(1, 8).value = "Circumference ("")"
    colWs.cells(1, 9).value = "Pole Address"
    colWs.cells(1, 10).value = "Hammer Test Failed"
    colWs.cells(1, 11).value = "Visual Test Failed"
    colWs.cells(1, 12).value = "Note"
    colWs.cells(1, 13).value = "collectedAt"
    colWs.cells(1, 14).value = "Location.latitude"
    colWs.cells(1, 15).value = "Location.longitude"
    colWs.cells(1, 16).value = "Tip"
    colWs.cells(1, 17).value = "Bonded Streetlight"
    colWs.cells(1, 18).value = "Tree Work"
    
    Set jobInfoWs = ThisWorkbook.sheets("Job Info")
    jobInfoWs.cells(1, 1).value = "ID"
    jobInfoWs.cells(1, 5).value = "Notification"
    jobInfoWs.cells(1, 3).value = "Permit"
    jobInfoWs.cells(1, 4).value = "Applicant"
    jobInfoWs.cells(1, 2).value = "Job Location"
    jobInfoWs.cells(1, 6).value = "County"
    jobInfoWs.cells(1, 7).value = "Township"
    jobInfoWs.cells(1, 8).value = "Fielder"
    
    Set spanWs = ThisWorkbook.sheets("Span")
    spanWs.cells(1, 1).value = "ID"
    spanWs.cells(1, 2).value = "Span Id"
    spanWs.cells(1, 3).value = "Span Length"
    spanWs.cells(1, 4).value = "Span Length.bearing"
    spanWs.cells(1, 5).value = "House"
    spanWs.cells(1, 6).value = "Tree Work"
    
    Set equipmentWs = ThisWorkbook.sheets("Equipment")
    equipmentWs.cells(1, 1).value = "ID"
    equipmentWs.cells(1, 2).value = "Type"
    equipmentWs.cells(1, 3).value = "Attachment Height"
    equipmentWs.cells(1, 4).value = "Bottom Height"
    equipmentWs.cells(1, 5).value = "Owner"
    
    Set anchorWs = ThisWorkbook.sheets("Anchor")
    anchorWs.cells(1, 1).value = "ID"
    anchorWs.cells(1, 2).value = "Lead Length"
    anchorWs.cells(1, 3).value = "Lead Length.bearing"
    anchorWs.cells(1, 4).value = "Owner"
    anchorWs.cells(1, 5).value = "ceCount"
    
    Set guysWs = ThisWorkbook.sheets("Anchor.Guys")
    guysWs.cells(1, 1).value = "ID"
    guysWs.cells(1, 2).value = "Size"
    guysWs.cells(1, 3).value = "Owner"
    guysWs.cells(1, 4).value = "Attachment Height"
    guysWs.cells(1, 5).value = "Proposed Size"
    guysWs.cells(1, 6).value = "Proposed Lead"
    guysWs.cells(1, 7).value = "Proposed Direction"
    
    Set spanPowWs = ThisWorkbook.sheets("Span.Power Circuit")
    spanPowWs.cells(1, 1).value = "ID"
    spanPowWs.cells(1, 2).value = "Span Id"
    spanPowWs.cells(1, 3).value = "Size"
    spanPowWs.cells(1, 4).value = "Attachment Height"
    spanPowWs.cells(1, 5).value = "Mid Span Height"
    spanPowWs.cells(1, 6).value = "Type"
    spanPowWs.cells(1, 7).value = "WEP Height"
    
    Set spanComWs = ThisWorkbook.sheets("Span.Communication")
    spanComWs.cells(1, 1).value = "ID"
    spanComWs.cells(1, 2).value = "Span Id"
    spanComWs.cells(1, 3).value = "Size"
    spanComWs.cells(1, 4).value = "Owner"
    spanComWs.cells(1, 5).value = "Attachment Height"
    spanComWs.cells(1, 6).value = "Mid Span Height"
    spanComWs.cells(1, 7).value = "Type"
    spanComWs.cells(1, 8).value = "Modification"
    spanComWs.cells(1, 9).value = "Diameter"
    spanComWs.cells(1, 10).value = "Tension"
    spanComWs.cells(1, 11).value = "MrMove"
    spanComWs.cells(1, 12).value = "WEP Height"
    
    Dim spanRow As Integer: spanRow = 2
    Dim spanPowRow As Integer: spanPowRow = 2
    Dim spanComRow As Integer: spanComRow = 2
    Dim equipmentRow As Integer: equipmentRow = 2
    Dim colRow As Integer: colRow = 2
    Dim anchorRow As Integer: anchorRow = 2
    Dim guysRow As Integer: guysRow = 2
    For Each pole In poles
        Do While Left$(pole.poleNumber, 1) = "0" And Len(pole.poleNumber) > 1
            pole.poleNumber = Mid$(pole.poleNumber, 2)
        Loop
        colWs.cells(colRow, 1).value = pole.poleNumber
        colWs.cells(colRow, 2).NumberFormat = "@"
        colWs.cells(colRow, 2).value = pole.gisCEID
        colWs.cells(colRow, 3).NumberFormat = "@"
        colWs.cells(colRow, 3).value = pole.existingCEID
        colWs.cells(colRow, 4).value = pole.species & " > " & pole.Class & " > " & pole.height
        colWs.cells(colRow, 5).value = pole.classVerified
        colWs.cells(colRow, 6).value = pole.heightVerified
        colWs.cells(colRow, 7).value = pole.owner
        colWs.cells(colRow, 8).value = pole.glc
        colWs.cells(colRow, 9).value = pole.address
        colWs.cells(colRow, 10).value = pole.hammerTestFailed
        colWs.cells(colRow, 11).value = pole.visualCheckFailed
        colWs.cells(colRow, 12).value = pole.notes
        colWs.cells(colRow, 13).value = pole.collectedDate
        colWs.cells(colRow, 14).value = pole.latitude
        colWs.cells(colRow, 15).value = pole.longitude
        colWs.cells(colRow, 16).value = pole.agl
        colWs.cells(colRow, 17).value = pole.Bonded
        colWs.cells(colRow, 18).value = pole.treeWork
        
        jobInfoWs.cells(colRow, 1).value = pole.poleNumber
        jobInfoWs.cells(colRow, 2).value = jobLocation
        jobInfoWs.cells(colRow, 3).value = Permit
        jobInfoWs.cells(colRow, 4).value = applicant
        jobInfoWs.cells(colRow, 5).value = notification
        jobInfoWs.cells(colRow, 6).value = county
        jobInfoWs.cells(colRow, 7).value = township
        jobInfoWs.cells(colRow, 8).value = fielder
        
        For Each Span In pole.spans
            spanWs.cells(spanRow, 1).value = pole.poleNumber
            spanWs.cells(spanRow, 2).value = Span.spanId
            spanWs.cells(spanRow, 3).value = Span.distance
            spanWs.cells(spanRow, 4).value = Span.angle
            spanWs.cells(spanRow, 5).value = Span.houseNumber
            spanWs.cells(spanRow, 6).value = Span.treeWork
            
            spanRow = spanRow + 1
            Set primaryBundles = New Scripting.Dictionary
            For Each wire In Span.utilWires
                spanPowWs.cells(spanPowRow, 1).value = pole.poleNumber
                spanPowWs.cells(spanPowRow, 2).value = Span.spanId
                spanPowWs.cells(spanPowRow, 3).value = wire.size
                spanPowWs.cells(spanPowRow, 4).value = wire.height
                spanPowWs.cells(spanPowRow, 5).value = wire.midspans(Span.spanSlot)
                spanPowWs.cells(spanPowRow, 6).value = wire.componentType
                spanPowWs.cells(spanPowRow, 7).value = wire.wepHeight
                
                spanPowRow = spanPowRow + 1
            Next wire
            
            For Each wire In Span.commWires
                spanComWs.cells(spanComRow, 1).value = pole.poleNumber
                spanComWs.cells(spanComRow, 2).value = Span.spanId
                spanComWs.cells(spanComRow, 3).value = wire.size
                spanComWs.cells(spanComRow, 4).value = wire.owner
                spanComWs.cells(spanComRow, 5).value = wire.height
                spanComWs.cells(spanComRow, 6).value = wire.midspans(Span.spanSlot)
                spanComWs.cells(spanComRow, 7).value = wire.componentType
                spanComWs.cells(spanComRow, 8).value = wire.modification
                spanComWs.cells(spanComRow, 9).value = wire.diameter
                spanComWs.cells(spanComRow, 10).value = wire.tensions(Span.spanSlot)
                spanComWs.cells(spanComRow, 11).value = wire.mrMoves(Span.spanSlot)
                spanComWs.cells(spanComRow, 12).value = wire.wepHeight
                
                spanComRow = spanComRow + 1
            Next wire
        Next Span
        
        For Each wire In pole.utilWires
            If wire.midspans.count = 0 Then
                equipmentWs.cells(equipmentRow, 1).value = pole.poleNumber
                equipmentWs.cells(equipmentRow, 2).value = wire.componentType & " > " & wire.size
                equipmentWs.cells(equipmentRow, 3).value = wire.height
                equipmentWs.cells(equipmentRow, 5).value = IIf(wire.owner = "", "Consumers Energy", wire.owner)
                
                equipmentRow = equipmentRow + 1
            End If
        Next wire
        
        For Each wire In pole.commWires
            If wire.midspans.count = 0 Then
                equipmentWs.cells(equipmentRow, 1).value = pole.poleNumber
                If wire.componentType = "DROP" Then
                    equipmentWs.cells(equipmentRow, 2).value = wire.componentType & " > " & wire.componentType
                Else
                    equipmentWs.cells(equipmentRow, 2).value = wire.componentType & " > " & wire.size
                End If
                equipmentWs.cells(equipmentRow, 3).value = wire.height
                equipmentWs.cells(equipmentRow, 5).value = IIf(wire.owner = "", "Consumers Energy", wire.owner)
                
                equipmentRow = equipmentRow + 1
            End If
        Next wire
        
        For Each anchor In pole.anchors
            anchorWs.cells(anchorRow, 1).value = pole.poleNumber
            anchorWs.cells(anchorRow, 2).value = anchor.distance
            anchorWs.cells(anchorRow, 3).value = anchor.angle
            anchorWs.cells(anchorRow, 4).value = anchor.owner
            anchorWs.cells(anchorRow, 5).value = anchor.ceCount
                
            anchorRow = anchorRow + 1
            
        Next anchor
        
        For Each guy In pole.guys
            guysWs.cells(guysRow, 1).value = pole.poleNumber
            guysWs.cells(guysRow, 2).value = guy.size
            If guy.owner = "CONSUMERS ENERGY" Then
                guysWs.cells(guysRow, 3).value = "UTILITY > " & guy.owner
            Else
                guysWs.cells(guysRow, 3).value = "COMMUNICATION > " & guy.owner
            End If
            guysWs.cells(guysRow, 4).value = guy.height
            
            guysRow = guysRow + 1
        Next guy
        
        For Each guy In pole.applicantGuys
            guysWs.cells(guysRow, 1).value = pole.poleNumber
        
            guysWs.cells(guysRow, 5).value = guy.proposedSize
            guysWs.cells(guysRow, 6).value = guy.proposedLead
            guysWs.cells(guysRow, 7).value = guy.proposedDirection
        
            guysRow = guysRow + 1
        Next guy
        
        For Each Equipment In pole.equipments
            equipmentWs.cells(equipmentRow, 1).value = pole.poleNumber
            equipmentWs.cells(equipmentRow, 2).value = Equipment.componentType & " > " & Equipment.size
            equipmentWs.cells(equipmentRow, 3).value = Equipment.height
            equipmentWs.cells(equipmentRow, 4).value = Equipment.bottomHeight
            equipmentWs.cells(equipmentRow, 5).value = IIf(Equipment.owner = "", "Consumers Energy", Equipment.owner)
            
            equipmentRow = equipmentRow + 1
        Next Equipment
    
        colRow = colRow + 1
    Next pole
End Sub

