VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Project"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Notification As String
Public permit As String
Public applicant As String
Public jobLocation As String
Public county As String
Public township As String
Public fielder As String
Public poles As Collection

Private Sub Class_Initialize()
    Set poles = New Collection
End Sub

Public Sub extractFromSheets()
    Dim sheet As Worksheet
    For Each sheet In ThisWorkbook.sheets
        If sheet.name <> "4 Spans" And sheet.name <> "8 Spans" And sheet.name <> "12 Spans" And sheet.Cells(2, 2).Value = "Notification:" Then
            Dim pole As pole: Set pole = New pole
            Call pole.extractFromSheet(sheet)
            Me.poles.Add pole
            
            If Me.township = "" Then Me.township = sheet.Range("TWP")
            If Me.county = "" Then Me.county = sheet.Range("COUNTY")
            If Me.applicant = "" Then Me.applicant = sheet.Range("APPLICANT")
            If Me.Notification = "" Then Me.Notification = sheet.Range("NOTIFICATION")
            If Me.permit = "" Then Me.permit = sheet.Range("PERMIT")
        End If
    Next sheet
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
    Dim span As span
    For Each sheet In ThisWorkbook.sheets
        If Utilities.IsPDS(sheet) And sheetsCreated.Exists(sheet.name) Then
            Call RemedyGen.calculateProposedMidspans(sheet)
            Set pole = Me.findPole(sheet.name)
            For Each span In pole.spans
                If pole.applicant.mrMoves.Exists(span.spanSlot) And span.otherPole <> "" Then
                    Dim lines() As String
                    lines = Split(sheet.Range("CMRF2").Value, vbLf)
                    For j = LBound(lines) To UBound(lines)
                        If InStr(lines(j), " @P" & span.otherPole) > 0 Then
                            lines(j) = Utilities.inchesToFeetInches(Utilities.convertToInches(lines(j)) + CInt(pole.applicant.mrMoves(span.spanSlot))) & " @P" & span.otherPole
                        End If
                    Next j
                    sheet.Range("CMRF2").Value = Join(lines, vbLf)
                End If
            Next span
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
    jobWsLastRow = jobWs.Cells(jobWs.Rows.count, "A").End(xlUp).row

    Set colWs = ThisWorkbook.sheets("Collection")
    colWsLastRow = colWs.Cells(colWs.Rows.count, "A").End(xlUp).row
    
    Set spanWs = ThisWorkbook.sheets("Span")
    spanWsLastRow = spanWs.Cells(spanWs.Rows.count, "A").End(xlUp).row
    spanWsRow = 2
    
    Set spanComWs = ThisWorkbook.sheets("Span.Communication")
    spanComWsLastRow = spanComWs.Cells(spanComWs.Rows.count, "A").End(xlUp).row
    spanComWsRow = 2
    
    Set spanPowWs = ThisWorkbook.sheets("Span.Power Circuit")
    spanPowWsLastRow = spanPowWs.Cells(spanPowWs.Rows.count, "A").End(xlUp).row
    spanPowWsRow = 2
    
    Set anchorWs = ThisWorkbook.sheets("Anchor")
    anchorWsLastRow = anchorWs.Cells(anchorWs.Rows.count, "A").End(xlUp).row
    anchorWsRow = 2
    
    Set guyWs = ThisWorkbook.sheets("Anchor.Guys")
    guyWsLastRow = guyWs.Cells(guyWs.Rows.count, "A").End(xlUp).row
    guyWsRow = 2
    
    Set equipmentWs = ThisWorkbook.sheets("Equipment")
    equipmentWsLastRow = equipmentWs.Cells(equipmentWs.Rows.count, "A").End(xlUp).row
    equipmentWsRow = 2
    
    Me.jobLocation = jobWs.Cells(2, 2)
    Me.permit = jobWs.Cells(2, 3)
    Me.applicant = jobWs.Cells(2, 4)
    Me.Notification = jobWs.Cells(2, 5)
    Me.county = jobWs.Cells(2, 6)
    Me.township = jobWs.Cells(2, 7)
    Me.fielder = jobWs.Cells(2, 8)
    
    For colWsRow = 2 To colWsLastRow
        Dim pole As pole: Set pole = New pole
        
        pole.poleNumber = colWs.Cells(colWsRow, 1)
        pole.gisCEID = colWs.Cells(colWsRow, 2)
        pole.existingCEID = colWs.Cells(colWsRow, 3)
        If pole.existingCEID = "" Then pole.existingCEID = pole.gisCEID
        parts = Split(colWs.Cells(colWsRow, 4), " > ")
        pole.species = parts(0)
        pole.Class = parts(1)
        pole.height = parts(2)
        pole.classVerified = colWs.Cells(colWsRow, 5).Value
        pole.heightVerified = colWs.Cells(colWsRow, 6).Value
        pole.owner = colWs.Cells(colWsRow, 7).Value
        pole.glc = colWs.Cells(colWsRow, 8).Value
        pole.address = colWs.Cells(colWsRow, 9).Value
        pole.hammerTestFailed = colWs.Cells(colWsRow, 10).Value
        pole.visualCheckFailed = colWs.Cells(colWsRow, 11).Value
        pole.notes = colWs.Cells(colWsRow, 12).Value
        pole.collectedDate = colWs.Cells(colWsRow, 13).Value
        pole.latitude = colWs.Cells(colWsRow, 14).Value
        pole.longitude = colWs.Cells(colWsRow, 15).Value
        pole.agl = colWs.Cells(colWsRow, 16).Value
        pole.Bonded = colWs.Cells(colWsRow, 17).Value
        pole.treeWork = colWs.Cells(colWsRow, 18).Value
        
        'Span section
        Dim slotCount As Integer: slotCount = 0
        For spanWsRow = spanWsRow To spanWsLastRow
            If spanWs.Cells(spanWsRow, 1) <> pole.poleNumber Then Exit For
            
            Dim span As span: Set span = New span
            
            span.spanId = spanWs.Cells(spanWsRow, 2)
            span.distance = spanWs.Cells(spanWsRow, 3)
            span.angle = spanWs.Cells(spanWsRow, 4)
            span.direction = Utilities.degreesToDirection(span.angle)
            span.houseNumber = spanWs.Cells(spanWsRow, 5)
            span.treeWork = spanWs.Cells(spanWsRow, 6)
            
            For j = 2 To spanWsLastRow
                If j <> spanWsRow Then
                    If spanWs.Cells(j, 2) = span.spanId Then
                        span.otherPole = spanWs.Cells(j, 1)
                        Exit For
                    End If
                End If
            Next j
            
            slotCount = slotCount + 1
            span.spanSlot = slotCount
            
            pole.spans.Add span
        Next spanWsRow
        
        'Communication wire section
        Dim wireDict As Scripting.Dictionary: Set wireDict = New Scripting.Dictionary
        For spanComWsRow = spanComWsRow To spanComWsLastRow
            If spanComWs.Cells(spanComWsRow, 1) <> pole.poleNumber Then Exit For
            
            Set span = pole.findSpanBySpanId(spanComWs.Cells(spanComWsRow, 2))
            size = spanComWs.Cells(spanComWsRow, 3)
            owner = spanComWs.Cells(spanComWsRow, 4)
            height = spanComWs.Cells(spanComWsRow, 5)
            componentType = spanComWs.Cells(spanComWsRow, 7)
            modification = spanComWs.Cells(spanComWsRow, 8)
            key = size & owner & height & componentType
            
            Dim Wire As Wire
            If Not wireDict.Exists(key) Then
                Set Wire = New Wire
            
                Wire.size = size
                Wire.owner = owner
                Wire.height = height
                Wire.componentType = componentType
                Wire.modification = modification
                
                wireDict.Add key, Wire
            Else
                continue = True
                Do While continue
                    If wireDict(key).midspans.Exists(span.spanSlot) Then
                        key = key & "A"
                        If Not wireDict.Exists(key) Then
                            Set Wire = New Wire
                
                            Wire.size = size
                            Wire.owner = owner
                            Wire.height = height
                            Wire.componentType = componentType
                            Wire.modification = modification
                            
                            wireDict.Add key, Wire
                            
                            continue = False
                        End If
                    Else
                        continue = False
                    End If
                Loop
            End If
            
            Set Wire = wireDict(key)
            Wire.midspans.Add span.spanSlot, spanComWs.Cells(spanComWsRow, 6)
            If spanComWs.Cells(spanComWsRow, 9) <> "" Then Wire.diameter = spanComWs.Cells(spanComWsRow, 9)
            If spanComWs.Cells(spanComWsRow, 10) <> "" Then Wire.tensions.Add span.spanSlot, spanComWs.Cells(spanComWsRow, 10)
            If spanComWs.Cells(spanComWsRow, 11) <> "" Then Wire.mrMoves.Add span.spanSlot, spanComWs.Cells(spanComWsRow, 11)
            
            span.wires.Add Wire
            span.commWires.Add Wire
        Next spanComWsRow
        
        For Each key In wireDict
            Set Wire = wireDict(key)
        
            If Wire.componentType <> "PROPOSED" Then
                pole.wires.Add Wire
                pole.commWires.Add Wire
                pole.commComponents.Add Wire
            Else
                Set pole.applicant = Wire
            End If
        Next key
        
        'Utility Wire Section
        Set wireDict = New Scripting.Dictionary
        For spanPowWsRow = spanPowWsRow To spanPowWsLastRow
            If spanPowWs.Cells(spanPowWsRow, 1) <> pole.poleNumber Then Exit For
            
            Set span = pole.findSpanBySpanId(spanPowWs.Cells(spanPowWsRow, 2))
            size = spanPowWs.Cells(spanPowWsRow, 3)
            height = spanPowWs.Cells(spanPowWsRow, 4)
            componentType = spanPowWs.Cells(spanPowWsRow, 6)
            key = size & height & componentType
            
            If Not wireDict.Exists(key) Then
                Set Wire = New Wire
            
                Wire.size = size
                Wire.height = height
                Wire.componentType = componentType
                
                wireDict.Add key, Wire
            Else
                continue = True
                Do While continue
                    If wireDict(key).midspans.Exists(span.spanSlot) Then
                        key = key & "A"
                        If Not wireDict.Exists(key) Then
                            Set Wire = New Wire
                
                            Wire.size = size
                            Wire.height = height
                            Wire.componentType = componentType
                            wireDict.Add key, Wire
                            
                            continue = False
                        End If
                    Else
                        continue = False
                    End If
                Loop
            End If
            
            Set Wire = wireDict(key)
            Wire.midspans.Add span.spanSlot, spanPowWs.Cells(spanPowWsRow, 5)
            
            span.wires.Add Wire
            span.utilWires.Add Wire
        Next spanPowWsRow
        
        For Each key In wireDict
            Set Wire = wireDict(key)
        
            pole.wires.Add Wire
            pole.utilWires.Add Wire
            
            If Wire.componentType <> "PRI" Then pole.utilComponents.Add Wire
        Next key
        
        'Anchor section
        For anchorWsRow = anchorWsRow To anchorWsLastRow
            If anchorWs.Cells(anchorWsRow, 1) <> pole.poleNumber Then Exit For
        
            Dim Anchor As Anchor: Set Anchor = New Anchor
            
            Anchor.distance = anchorWs.Cells(anchorWsRow, 2)
            Anchor.angle = anchorWs.Cells(anchorWsRow, 3)
            Anchor.direction = Utilities.degreesToDirection(Anchor.angle)
            Anchor.owner = anchorWs.Cells(anchorWsRow, 4)
            
            pole.anchors.Add Anchor
        Next anchorWsRow
        
        'Guy section
        For guyWsRow = guyWsRow To guyWsLastRow
            If guyWs.Cells(guyWsRow, 1) <> pole.poleNumber Then Exit For
        
            Dim guy As guy: Set guy = New guy
            
            If guyWs.Cells(guyWsRow, 4) <> "" Then
                guy.size = guyWs.Cells(guyWsRow, 2)
                parts = Split(guyWs.Cells(guyWsRow, 3), " > ")
                ownerType = parts(0)
                guy.owner = parts(1)
                guy.height = guyWs.Cells(guyWsRow, 4)
                guy.componentType = "DG"
                
                pole.guys.Add guy
                If ownerType = "UTILITY" Then
                    pole.utilComponents.Add guy
                Else
                    pole.commComponents.Add guy
                End If
            ElseIf guyWs.Cells(guyWsRow, 5) <> "" Then
                guy.proposedSize = guyWs.Cells(guyWsRow, 5)
                guy.proposedLead = guyWs.Cells(guyWsRow, 6)
                guy.proposedDirection = guyWs.Cells(guyWsRow, 7)
                
                pole.applicantGuys.Add guy
            End If
        Next guyWsRow

        'Equipment section
        For equipmentWsRow = equipmentWsRow To equipmentWsLastRow
            If equipmentWs.Cells(equipmentWsRow, 1) <> pole.poleNumber Then Exit For
            
            Dim equipment As equipment: Set equipment = New equipment
            
            parts = Split(equipmentWs.Cells(equipmentWsRow, 2), " > ")
            equipment.componentType = parts(0)
            equipment.size = parts(1)
            equipment.height = equipmentWs.Cells(equipmentWsRow, 3)
            equipment.bottomHeight = equipmentWs.Cells(equipmentWsRow, 4)
            equipment.owner = equipmentWs.Cells(equipmentWsRow, 5)
            
            pole.equipments.Add equipment
            If equipment.owner = "" Or UCase(equipment.owner) = "CONSUMERS ENERGY" Then
                pole.utilComponents.Add equipment
            Else
                pole.commComponents.Add equipment
            End If
        Next equipmentWsRow
        
        Me.poles.Add pole
    Next colWsRow
End Sub

Public Sub fillImportDataFormat()
    Dim pole As pole
    Dim span As span
    Dim Wire As Wire
    Dim Anchor As Anchor
    Dim guy As guy
    Dim equipment As equipment
    Dim primaryBundles As Scripting.Dictionary
    
    Set colWs = ThisWorkbook.sheets("Collection")
    colWs.Cells(1, 1).Value = "ID"
    colWs.Cells(1, 2).Value = "GIS CE ID Tag"
    colWs.Cells(1, 3).Value = "Existing CE ID Tag"
    colWs.Cells(1, 4).Value = "Type"
    colWs.Cells(1, 5).Value = "Pole Class Verified"
    colWs.Cells(1, 6).Value = "Pole Height Verified"
    colWs.Cells(1, 7).Value = "Owner"
    colWs.Cells(1, 8).Value = "Circumference ("")"
    colWs.Cells(1, 9).Value = "Pole Address"
    colWs.Cells(1, 10).Value = "Hammer Test Failed"
    colWs.Cells(1, 11).Value = "Visual Test Failed"
    colWs.Cells(1, 12).Value = "Note"
    colWs.Cells(1, 13).Value = "collectedAt"
    colWs.Cells(1, 14).Value = "Location.latitude"
    colWs.Cells(1, 15).Value = "Location.longitude"
    colWs.Cells(1, 16).Value = "Tip"
    colWs.Cells(1, 17).Value = "Bonded Streetlight"
    colWs.Cells(1, 18).Value = "Tree Work"
    
    Set jobInfoWs = ThisWorkbook.sheets("Job Info")
    jobInfoWs.Cells(1, 1).Value = "ID"
    jobInfoWs.Cells(1, 5).Value = "Notification"
    jobInfoWs.Cells(1, 3).Value = "Permit"
    jobInfoWs.Cells(1, 4).Value = "Applicant"
    jobInfoWs.Cells(1, 2).Value = "Job Location"
    jobInfoWs.Cells(1, 6).Value = "County"
    jobInfoWs.Cells(1, 7).Value = "Township"
    jobInfoWs.Cells(1, 8).Value = "Fielder"
    
    Set spanWs = ThisWorkbook.sheets("Span")
    spanWs.Cells(1, 1).Value = "ID"
    spanWs.Cells(1, 2).Value = "Span Id"
    spanWs.Cells(1, 3).Value = "Span Length"
    spanWs.Cells(1, 4).Value = "Span Length.bearing"
    spanWs.Cells(1, 5).Value = "House"
    spanWs.Cells(1, 6).Value = "Tree Work"
    
    Set equipmentWs = ThisWorkbook.sheets("Equipment")
    equipmentWs.Cells(1, 1).Value = "ID"
    equipmentWs.Cells(1, 2).Value = "Type"
    equipmentWs.Cells(1, 3).Value = "Attachment Height"
    equipmentWs.Cells(1, 4).Value = "Bottom Height"
    equipmentWs.Cells(1, 5).Value = "Owner"
    
    Set anchorWs = ThisWorkbook.sheets("Anchor")
    anchorWs.Cells(1, 1).Value = "ID"
    anchorWs.Cells(1, 2).Value = "Lead Length"
    anchorWs.Cells(1, 3).Value = "Lead Length.bearing"
    anchorWs.Cells(1, 4).Value = "Owner"
    
    Set guysWs = ThisWorkbook.sheets("Anchor.Guys")
    guysWs.Cells(1, 1).Value = "ID"
    guysWs.Cells(1, 2).Value = "Size"
    guysWs.Cells(1, 3).Value = "Owner"
    guysWs.Cells(1, 4).Value = "Attachment Height"
    guysWs.Cells(1, 5).Value = "Proposed Size"
    guysWs.Cells(1, 6).Value = "Proposed Lead"
    guysWs.Cells(1, 7).Value = "Proposed Direction"
    
    Set spanPowWs = ThisWorkbook.sheets("Span.Power Circuit")
    spanPowWs.Cells(1, 1).Value = "ID"
    spanPowWs.Cells(1, 2).Value = "Span Id"
    spanPowWs.Cells(1, 3).Value = "Size"
    spanPowWs.Cells(1, 4).Value = "Attachment Height"
    spanPowWs.Cells(1, 5).Value = "Mid Span Height"
    spanPowWs.Cells(1, 6).Value = "Type"
    spanPowWs.Cells(1, 7).Value = "WEP Height"
    
    Set spanComWs = ThisWorkbook.sheets("Span.Communication")
    spanComWs.Cells(1, 1).Value = "ID"
    spanComWs.Cells(1, 2).Value = "Span Id"
    spanComWs.Cells(1, 3).Value = "Size"
    spanComWs.Cells(1, 4).Value = "Owner"
    spanComWs.Cells(1, 5).Value = "Attachment Height"
    spanComWs.Cells(1, 6).Value = "Mid Span Height"
    spanComWs.Cells(1, 7).Value = "Type"
    spanComWs.Cells(1, 8).Value = "Modification"
    spanComWs.Cells(1, 9).Value = "Diameter"
    spanComWs.Cells(1, 10).Value = "Tension"
    spanComWs.Cells(1, 11).Value = "MrMove"
    
    Dim spanRow As Integer: spanRow = 2
    Dim spanPowRow As Integer: spanPowRow = 2
    Dim spanComRow As Integer: spanComRow = 2
    Dim equipmentRow As Integer: equipmentRow = 2
    Dim colRow As Integer: colRow = 2
    Dim anchorRow As Integer: anchorRow = 2
    Dim guysRow As Integer: guysRow = 2
    For Each pole In poles
        colWs.Cells(colRow, 1).Value = pole.poleNumber
        colWs.Cells(colRow, 2).NumberFormat = "@"
        colWs.Cells(colRow, 2).Value = pole.gisCEID
        colWs.Cells(colRow, 3).NumberFormat = "@"
        colWs.Cells(colRow, 3).Value = pole.existingCEID
        colWs.Cells(colRow, 4).Value = pole.species & " > " & pole.Class & " > " & pole.height
        colWs.Cells(colRow, 5).Value = pole.classVerified
        colWs.Cells(colRow, 6).Value = pole.heightVerified
        colWs.Cells(colRow, 7).Value = pole.owner
        colWs.Cells(colRow, 8).Value = pole.glc
        colWs.Cells(colRow, 9).Value = pole.address
        colWs.Cells(colRow, 10).Value = pole.hammerTestFailed
        colWs.Cells(colRow, 11).Value = pole.visualCheckFailed
        colWs.Cells(colRow, 12).Value = pole.notes
        colWs.Cells(colRow, 13).Value = pole.collectedDate
        colWs.Cells(colRow, 14).Value = pole.latitude
        colWs.Cells(colRow, 15).Value = pole.longitude
        colWs.Cells(colRow, 16).Value = pole.agl
        colWs.Cells(colRow, 17).Value = pole.Bonded
        colWs.Cells(colRow, 18).Value = pole.treeWork
        
        jobInfoWs.Cells(colRow, 1).Value = pole.poleNumber
        jobInfoWs.Cells(colRow, 2).Value = jobLocation
        jobInfoWs.Cells(colRow, 3).Value = permit
        jobInfoWs.Cells(colRow, 4).Value = applicant
        jobInfoWs.Cells(colRow, 5).Value = Notification
        jobInfoWs.Cells(colRow, 6).Value = county
        jobInfoWs.Cells(colRow, 7).Value = township
        jobInfoWs.Cells(colRow, 8).Value = fielder
        
        For Each span In pole.spans
            spanWs.Cells(spanRow, 1).Value = pole.poleNumber
            spanWs.Cells(spanRow, 2).Value = span.spanId
            spanWs.Cells(spanRow, 3).Value = span.distance
            spanWs.Cells(spanRow, 4).Value = span.angle
            spanWs.Cells(spanRow, 5).Value = span.houseNumber
            spanWs.Cells(spanRow, 6).Value = span.treeWork
            
            spanRow = spanRow + 1
            Set primaryBundles = New Scripting.Dictionary
            For Each Wire In span.utilWires
                spanPowWs.Cells(spanPowRow, 1).Value = pole.poleNumber
                spanPowWs.Cells(spanPowRow, 2).Value = span.spanId
                spanPowWs.Cells(spanPowRow, 3).Value = Wire.size
                spanPowWs.Cells(spanPowRow, 4).Value = Wire.height
                spanPowWs.Cells(spanPowRow, 5).Value = Wire.midspans(span.spanSlot)
                spanPowWs.Cells(spanPowRow, 6).Value = Wire.componentType
                spanPowWs.Cells(spanPowRow, 7).Value = Wire.wepHeight
                
                spanPowRow = spanPowRow + 1
            Next Wire
            
            For Each Wire In span.commWires
                spanComWs.Cells(spanComRow, 1).Value = pole.poleNumber
                spanComWs.Cells(spanComRow, 2).Value = span.spanId
                spanComWs.Cells(spanComRow, 3).Value = Wire.size
                spanComWs.Cells(spanComRow, 4).Value = Wire.owner
                spanComWs.Cells(spanComRow, 5).Value = Wire.height
                spanComWs.Cells(spanComRow, 6).Value = Wire.midspans(span.spanSlot)
                spanComWs.Cells(spanComRow, 7).Value = Wire.componentType
                spanComWs.Cells(spanComRow, 8).Value = Wire.modification
                spanComWs.Cells(spanComRow, 9).Value = Wire.diameter
                spanComWs.Cells(spanComRow, 10).Value = Wire.tensions(span.spanSlot)
                spanComWs.Cells(spanComRow, 11).Value = Wire.mrMoves(span.spanSlot)
                
                spanComRow = spanComRow + 1
            Next Wire

        Next span
        
        For Each Anchor In pole.anchors
            anchorWs.Cells(anchorRow, 1).Value = pole.poleNumber
            anchorWs.Cells(anchorRow, 2).Value = Anchor.distance
            anchorWs.Cells(anchorRow, 3).Value = Anchor.angle
            anchorWs.Cells(anchorRow, 4).Value = Anchor.owner
                
            anchorRow = anchorRow + 1
            
        Next Anchor
        
        For Each guy In pole.guys
            guysWs.Cells(guysRow, 1).Value = pole.poleNumber
            guysWs.Cells(guysRow, 2).Value = guy.size
            If guy.owner = "CONSUMERS ENERGY" Then
                guysWs.Cells(guysRow, 3).Value = "UTILITY > " & guy.owner
            Else
                guysWs.Cells(guysRow, 3).Value = "COMMUNICATION > " & guy.owner
            End If
            guysWs.Cells(guysRow, 4).Value = guy.height
            
            guysRow = guysRow + 1
        Next guy
        
        For Each guy In pole.applicantGuys
            guysWs.Cells(guysRow, 1).Value = pole.poleNumber
        
            guysWs.Cells(guysRow, 5).Value = guy.proposedSize
            guysWs.Cells(guysRow, 6).Value = guy.proposedLead
            guysWs.Cells(guysRow, 7).Value = guy.proposedDirection
        
            guysRow = guysRow + 1
        Next guy
        
        For Each equipment In pole.equipments
            equipmentWs.Cells(equipmentRow, 1).Value = pole.poleNumber
            equipmentWs.Cells(equipmentRow, 2).Value = equipment.componentType & " > " & equipment.size
            equipmentWs.Cells(equipmentRow, 3).Value = equipment.height
            equipmentWs.Cells(equipmentRow, 4).Value = equipment.bottomHeight
            equipmentWs.Cells(equipmentRow, 5).Value = IIf(equipment.owner = "", "Consumers Energy", equipment.owner)
            
            equipmentRow = equipmentRow + 1
        Next equipment
    
        colRow = colRow + 1
    Next pole
End Sub

