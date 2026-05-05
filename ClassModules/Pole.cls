VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public poleNumber As String
Public gisCEID As String
Public existingCEID As String
Public address As String
Public notes As String
Public collectedDate As String
Public Bonded As String
Public species As String
Public height As String
Public Class As String
Public owner As String
Public latitude As Double
Public longitude As Double
Public agl As Double
Public glc As String
Public spans As Collection
Public arms As Collection
Public insulators As Collection
Public wires As Collection
Public utilWires As Collection
Public commWires As Collection
Public utilComponents As Collection
Public commComponents As Collection
Public equipments As Collection
Public anchors As Collection
Public guys As Collection
Public lowestPower As Integer
Public tangent As Boolean
Public angle As Boolean
Public corner As Boolean
Public deadend As Boolean
Public buck As Boolean
Public secOnly As Boolean
Public treeWork As Boolean
Public slBottomBracketHeight As Integer
Public permit As String

Public primaries As Collection
Public neutrals  As Collection
Public secondaries  As Collection
Public openWires As Collection
Public services  As Collection

Public streeLightBB As Integer

Public applicant As Wire
Public applicantGuys As Collection
Public overlash As Boolean

Public location As String
Public Alt1 As String
Public NJUNS As String
Public njunsSteps As Collection
Public njunsTicket As String
Public ttc As String
Public newHeight As Integer
Public dSpace As Integer
Public pSpace As Integer
Public replacePole As Boolean

Public classVerified As Boolean
Public heightVerified As Boolean
Public hammerTestFailed As Boolean
Public visualCheckFailed As Boolean

Public violation40SZV As String
Public violationUtil156MSGCV As String
Public violationCom156MSGCV As String
Public violation30MSSV As String
Public violation12CSV As String
Public violation6MSSV As String
Public violation40SLSV As String
Public violation12SLDLSV As String

Private Sub Class_Initialize()
    Set spans = New Collection
    Set arms = New Collection
    Set insulators = New Collection
    Set wires = New Collection
    Set utilWires = New Collection
    Set commWires = New Collection
    Set utilComponents = New Collection
    Set commComponents = New Collection
    Set equipments = New Collection
    Set anchors = New Collection
    Set guys = New Collection
    Set applicant = New Wire
    
    Set primaries = New Collection
    Set neutrals = New Collection
    Set secondaries = New Collection
    Set openWires = New Collection
    Set services = New Collection
    
    Set applicantGuys = New Collection
    
    Set njunsSteps = New Collection
    
    lowestPower = 9999
End Sub

Public Sub extractFromSheet(sheet As Worksheet)
    Me.poleNumber = sheet.Range("POLENUM")
    Me.existingCEID = sheet.Range("CEID")
    Me.permit = sheet.Range("PERMIT")
    
    Me.applicant.owner = sheet.Range("APPLICANT")
    Me.applicant.height = Utilities.convertToInches(sheet.Range("PROPOSEDHEIGHT"))
    Me.applicant.modification = Utilities.convertToInches(sheet.Range("CMRF1"))
    
    Me.lowestPower = Utilities.convertToInches(sheet.Range("LWSTPWR"))
    Me.streeLightBB = Utilities.convertToInches(sheet.Range("STLTBRKT"))
    Me.Bonded = sheet.Range("BONDED")
    
    For i = 1 To 12
        If Not Utilities.RangeExists(sheet, "TOPOLE" & i) Then Exit For
        If sheet.Range("TOPOLE" & i).offset(1, 0).Value <> "" Then
            applicant.midspans.Add i, Utilities.convertToInches(sheet.Range("TOPOLE" & i).offset(1, 0).Value)
        End If
    Next i
    
    Me.height = Utilities.convertToInches(ThisWorkbook.RemoveParentheses(sheet.Range("HEIGHT")) & "'")
    Me.address = sheet.Range("ADDRESS")
    Me.slBottomBracketHeight = Utilities.convertToInches(sheet.Range("STLTBRKT"))
    
    Me.location = sheet.Range("DL")
    Me.Alt1 = sheet.Range("ALTONE")
    Me.ttc = sheet.Range("TTC")
    Me.NJUNS = sheet.Range("NJUNS")
    If Me.NJUNS <> "" Then Set Me.njunsSteps = Utilities.SplitByBlankLine(Me.NJUNS)
    Me.njunsTicket = sheet.Range("NJUNSTICKET")
    Me.latitude = sheet.Range("LATITUDE")
    Me.longitude = sheet.Range("LONGITUDE")
    
    Me.buck = sheet.Range("BUCK")
    
    Me.newHeight = Me.height
    Me.dSpace = 0
    Me.pSpace = 0
    
    Dim regex As Object: Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\d[0:5]\s*-\s*\d\s*\/\s*(\d[0:5])\s*-\s*\d"
    regex.Global = True
    regex.IgnoreCase = True
    
    lines = Split(Me.Alt1, vbLf)
    For Each line In lines
        line = Replace(line, " ", "")
        If InStr(line, "D=") > 0 Then Me.dSpace = Utilities.convertToInches(Mid(line, 3, Len(line) - 2))
        If InStr(line, "P=") > 0 Then Me.pSpace = Utilities.convertToInches(Mid(line, 3, Len(line) - 2))
        If regex.test(line) Then
            Me.newHeight = regex.Execute(line)(0).SubMatches(0) * 12
            Me.replacePole = True
        End If
        
    Next line
    
    If InStr(sheet.Range("PROPOSEDHEIGHT"), "OL") > 0 Then Me.overlash = True
    
    Dim Span As Span
    For i = 1 To 12
        If Not Utilities.RangeExists(sheet, "TOPOLE" & i) Then Exit For
        If sheet.Range("TOPOLE" & i) = "" Or InStr(sheet.Range("TOPOLE" & i), "(") = 0 Or InStr(sheet.Range("TOPOLE" & i), ")") = 0 Then Exit For
        Set Span = New Span
        Span.distance = Utilities.OnlyNumbers(sheet.Range("SPAN" & i))
        Span.angle = CInt(ThisWorkbook.InsideParentheses(sheet.Range("TOPOLE" & i)))
        If Utilities.SheetExists(ThisWorkbook.RemoveParentheses(sheet.Range("TOPOLE" & i))) Then
            Span.otherPole = ThisWorkbook.RemoveParentheses(sheet.Range("TOPOLE" & i))
        End If
        Span.direction = Utilities.degreesToDirection(Span.angle)
        Span.spanSlot = i
        Me.spans.Add Span
    Next i
    
    Dim comp As Component
    Dim Wire As Wire
    Dim equipment As equipment
    Dim guy As guy
    Dim modificationCells As Collection: Set modificationCells = New Collection
    For i = 0 To 100
        If sheet.Range("CMOWNER").offset(i, 0).Interior.color <> 16312794 Then Exit For
        If Utilities.convertToInches(sheet.Range("CMHEIGHT").offset(i, 0)) > 0 Then
            Set comp = New Component
            comp.height = Utilities.convertToInches(sheet.Range("CMHEIGHT").offset(i, 0).text)
            comp.owner = sheet.Range("CMOWNER").offset(i, 0).text
            comp.size = sheet.Range("CMSIZE").offset(i, 0).text
            If InStr(comp.owner, " DG") > 0 Then
                comp.componentType = "DG"
                comp.owner = Replace(comp.owner, " DG", "")
            ElseIf InStr(comp.owner, " MSG") > 0 Then
                comp.componentType = "MSG"
                comp.owner = Replace(comp.owner, " MSG", "")
            ElseIf InStr(comp.owner, " SVC") > 0 Then
                comp.componentType = "DROP"
                comp.owner = Replace(comp.owner, " SVC", "")
            End If
            
            If comp.componentType <> "DG" Then
                Set Wire = comp.getWire()
                If overlash And Wire.owner = applicant.owner Then applicant.height = Wire.height
                
                For Each Span In Me.spans
                    If Replace(sheet.Range("CMMIDSPAN" & Span.spanSlot).offset(i, 0).text, "-", "") <> "" Then
                        Wire.midspans.Add Span.spanSlot, Utilities.convertToInches(sheet.Range("CMMIDSPAN" & Span.spanSlot).offset(i, 0).text)
                        If Wire.componentType = "" Then Wire.componentType = "COM"
                        Span.commComponents.Add Wire
                        Span.commWires.Add Wire
                        Span.wires.Add Wire
                        If Span.houseNumber = "" Then Span.houseNumber = ThisWorkbook.InsideParentheses(sheet.Range("CMMIDSPAN" & Span.spanSlot).offset(i, 0))
                    End If
                Next Span
                
                Set modCell = findModCell(sheet, Wire.height, Wire.owner, False, modificationCells)
                If Not modCell Is Nothing Then Wire.modification = Utilities.convertToInches(modCell)
                
                If Wire.componentType = "DROP" Or Wire.componentType = "MSG" Or Wire.componentType = "COM" Then
                    Me.commComponents.Add Wire
                    Me.commWires.Add Wire
                    Me.wires.Add Wire
                Else
                    Set equipment = comp.getEquipment
                    equipment.bottomHeight = Utilities.convertToInches(ThisWorkbook.InsideParentheses(sheet.Range("CMHEIGHT").offset(i, 0).text))
                    Me.commComponents.Add equipment
                End If
            Else
                Set guy = comp.getDownGuy()
                Me.commComponents.Add guy
                Me.guys.Add guy
            End If
        End If
    Next i
    
    For i = 0 To 100
        If sheet.Range("UTTYPE").offset(i, 0).Interior.color <> 16312794 Then Exit For
        If Utilities.convertToInches(sheet.Range("UTHEIGHT").offset(i, 0)) > 0 Then
            Set comp = New Component
            comp.height = Utilities.convertToInches(sheet.Range("UTHEIGHT").offset(i, 0).text)
            comp.componentType = sheet.Range("UTTYPE").offset(i, 0).text
            comp.size = sheet.Range("UTSIZE").offset(i, 0).text
            
            If comp.componentType = "PRI" Or comp.componentType = "NEUT" Or comp.componentType = "SEC" Or comp.componentType = "SVC" Or comp.componentType = "OW" Or comp.componentType = "SPG" Then
                Set Wire = comp.getWire()
                
                If Wire.componentType = "PRI" And InStr(Wire.size, "Ø") > 0 Then
                    Wire.phase = Mid(comp.size, Len(Wire.size) - 1, 1)
                    Wire.size = Left(comp.size, Len(Wire.size) - 3)
                End If
                If Wire.componentType = "SPG" Then Wire.wepHeight = Utilities.convertToInches(ThisWorkbook.InsideParentheses(sheet.Range("UTHEIGHT").offset(i, 0).text))
                
                For Each Span In Me.spans
                    If Replace(sheet.Range("UTMIDSPAN" & Span.spanSlot).offset(i, 0).text, "-", "") <> "" Then
                        Wire.midspans.Add Span.spanSlot, Utilities.convertToInches(sheet.Range("UTMIDSPAN" & Span.spanSlot).offset(i, 0).text)
                        Span.utilComponents.Add Wire
                        Span.utilWires.Add Wire
                        Span.wires.Add Wire
                        If Span.houseNumber = "" Then Span.houseNumber = ThisWorkbook.InsideParentheses(sheet.Range("UTMIDSPAN" & Span.spanSlot).offset(i, 0))
                    End If
                Next Span
                
                If Wire.componentType = "PRI" Then Me.primaries.Add Wire
                If Wire.componentType = "NEUT" Then Me.neutrals.Add Wire
                If Wire.componentType = "SEC" Then Me.secondaries.Add Wire
                If Wire.componentType = "OW" Then Me.openWires.Add Wire
                If Wire.componentType = "SVC" Then Me.services.Add Wire
                
                Me.utilComponents.Add Wire
                Me.utilWires.Add Wire
                Me.wires.Add Wire
            ElseIf comp.componentType = "DG" Then
                Set guy = comp.getDownGuy()
                Me.utilComponents.Add guy
                Me.guys.Add guy
            Else
                Set equipment = comp.getEquipment
                equipment.bottomHeight = Utilities.convertToInches(ThisWorkbook.InsideParentheses(sheet.Range("UTHEIGHT").offset(i, 0).text))
                Me.utilComponents.Add equipment
                Me.equipments.Add equipment
            End If
        End If
    Next i
    
    Dim Anchor As Anchor
    For i = 0 To 100
        Set Anchor = New Anchor
        If sheet.Range("ANCHOROWNER").offset(i, 0).Value = "" Or sheet.Range("ANCHOROWNER").offset(i, 0).Value = "New APP" Then Exit For
        Anchor.owner = sheet.Range("ANCHOROWNER").offset(i, 0).Value
        Anchor.distance = Utilities.OnlyNumbers(sheet.Range("ANCHORLEAD").offset(i, 0).Value)
        Anchor.angle = Utilities.OnlyNumbers(sheet.Range("ANCHORDIRECTION").offset(i, 0).Value)
        Anchor.direction = Utilities.degreesToDirection(Anchor.angle)
        Me.anchors.Add Anchor
    Next i
    
End Sub

Public Sub createSheet(Project As Project, Optional selectedTemplate As Integer)
    Dim templateWs As Worksheet, fourSpanWs As Worksheet, eightSpanWs As Worksheet, twelveSpanWs As Worksheet
    Dim maxSpans As Integer

    Set fourSpanWs = ThisWorkbook.sheets("4 Spans")
    Set eightSpanWs = ThisWorkbook.sheets("8 Spans")
    Set twelveSpanWs = ThisWorkbook.sheets("12 Spans")
    
    If Not SheetExists(Me.poleNumber) Then
        ' Find the correct template to use
        If Me.spans.count > 8 Or selectedTemplate = 12 Then
            Set templateWs = twelveSpanWs
            maxSpans = 12
        ElseIf Me.spans.count > 4 Or selectedTemplate = 8 Then
            Set templateWs = eightSpanWs
            maxSpans = 8
        Else
            Set templateWs = fourSpanWs
            maxSpans = 4
        End If
    End If
                
    If templateWs Is Nothing Then
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        MsgBox "Error finding template file, if issue persists redownload a new version or contact Alex Abraham"
        Exit Sub
    End If
                
    Dim insertSpot As Worksheet: Set insertSpot = Nothing
    Set existingNames = New scripting.Dictionary
    
    ' copy the template
    For Each sheet In ThisWorkbook.sheets
        rpName = ThisWorkbook.RemoveParentheses(sheet.name)
        existingNames(rpName) = True
        If rpName <> "4 Spans" And rpName <> "8 Spans" And rpName <> "12 Spans" Then
            If sheet.Cells(2, 2).Value = "Notification:" Then
                If val(rpName) > val(poleNumber) Or (val(rpName) = val(poleNumber) And rpName > poleNumber) Then
                    Set insertSpot = sheet
                    Exit For
                End If
            End If
        End If
    Next sheet
    
    If insertSpot Is Nothing Then
        templateWs.Copy after:=ThisWorkbook.sheets(ThisWorkbook.sheets.count)
    Else
        templateWs.Copy Before:=insertSpot
    End If
    
    Dim pdsWs As Worksheet
    
    For Each sheet In ThisWorkbook.sheets
        If Not existingNames.Exists(ThisWorkbook.RemoveParentheses(sheet.name)) Then
            Set pdsWs = sheet
            pdsWs.visible = xlSheetVisible
            If Me.poleNumber <> "" Then pdsWs.name = Me.poleNumber
            Exit For
        End If
    Next sheet
    
    Call Me.fillSheet(pdsWs, maxSpans, Project)
    sheet.Activate
    sheet.Range("A1").Select
End Sub

Public Sub fillSheet(sheet As Worksheet, maxSpans As Integer, Project As Project)
    Dim comp As Variant

    'Header info
    sheet.Unprotect
    sheet.Range("NOTIFICATION").Value = Project.Notification
    sheet.Range("PERMIT").Value = Project.permit
    sheet.Range("APPLICANT").Value = Project.applicant
    sheet.Range("LOC").Value = Project.jobLocation
    sheet.Range("INSPECTEDBY").Value = Project.fielder
    sheet.Range("TWP").Value = Project.township
    sheet.Range("DATE").Value = Me.collectedDate
    sheet.Range("COUNTY").Value = Project.county
    sheet.Range("POLENUM").Value = Me.poleNumber
    sheet.Range("CLASS").Value = Me.Class
    sheet.Range("HEIGHT").Value = Me.height & IIf(Me.agl > 0, "(" & Utilities.inchesToFeetInches(Me.agl) & ")", "")
    sheet.Range("TYPE").Value = Me.species
    If Me.owner = "Consumers Energy" Then
        sheet.Range("CEPOLE").Value = True
    ElseIf Me.owner <> "" Then
        sheet.Range("OTHERPOLE").Value = True
        sheet.Range("OTHERPOLEOWNER").Value = Me.owner
    End If
    sheet.Range("CLASSESTIMATE").Value = IIf(Me.classVerified, "BRAND", "ESTIMATE")
    sheet.Range("GLC").Value = Me.glc
    sheet.Range("HEIGHTESTIMATE").Value = IIf(Me.heightVerified, "BRAND", "ESTIMATE")
    
    'Footer info
    sheet.Range("LATITUDE") = Me.latitude
    sheet.Range("LONGITUDE") = Me.longitude
    sheet.Range("CEID").Value = Me.existingCEID
    If Me.existingCEID <> Me.gisCEID And Utilities.isCEID(Me.gisCEID) Then
        If Me.notes <> "" Then Me.notes = vbLf & Me.notes
        Me.notes = "Old GIS CEID: " & Me.gisCEID & Me.notes
    End If
    sheet.Range("ADDRESS").Value = WorksheetFunction.Proper(Me.address)
    
    'Applicant info
    For i = 1 To 3
        If i > Me.applicantGuys.count Then Exit For
        Set guy = Me.applicantGuys(i)
        sheet.Range("NEWAPPSIZE").offset(i - 1, 0) = guy.proposedSize
        sheet.Range("NEWAPPLEAD").offset(i - 1, 0) = guy.proposedLead
        sheet.Range("NEWAPPDIR").offset(i - 1, 0) = guy.proposedDirection
    Next i
    If Me.applicant.height > 0 Then sheet.Range("PROPOSEDHEIGHT").Value = Utilities.inchesToFeetInches(Me.applicant.height)
    If Me.applicant.modification <> Me.applicant.height And Me.applicant.modification > 0 Then sheet.Range("CMRF1").Value = Utilities.inchesToFeetInches(Me.applicant.modification)
    sheet.Range("PROPOSEDDIAMETER") = Me.applicant.diameter
    
    'Span info
    Dim Span As Span
    For Each Span In Me.spans
        sheet.Range("SPAN" & Span.spanSlot).Value = CInt(Span.distance) & "'"
        sheet.Range("SPAN" & Span.spanSlot).HorizontalAlignment = xlLeft
        sheet.Range("TOPOLE" & Span.spanSlot).Value = Span.getToSpan()
        sheet.Range("TOPOLE" & Span.spanSlot).HorizontalAlignment = xlLeft
        sheet.Range("CLEARANCE" & Span.spanSlot).Value = "15'6"""
        If Me.applicant.midspans.Exists(Span.spanSlot) Then
            sheet.Range("TOPOLE" & Span.spanSlot).offset(1, 0).Value = Utilities.inchesToFeetInches(Me.applicant.midspans(Span.spanSlot))
            If Me.applicant.tensions.Exists(Span.spanSlot) Then sheet.Range("TOPOLE" & Span.spanSlot).offset(2, 0).Value = Me.applicant.tensions(Span.spanSlot)
        End If
        Call Span.setLowestPowerWire
        If Not Span.lowestPowerWire Is Nothing Then
            If Span.lowestPowerWire.componentType = "SVC" Then
                sheet.Range("CMMIDSPAN" & Span.spanSlot).offset(-2, 0).Value = Utilities.inchesToFeetInches(Span.lowestPowerWire.midspans(Span.spanSlot) - 12)
            Else
                sheet.Range("CMMIDSPAN" & Span.spanSlot).offset(-2, 0).Value = Utilities.inchesToFeetInches(Span.lowestPowerWire.midspans(Span.spanSlot) - 30)
            End If
        End If
    Next Span
    
    'Make space for anchors if needed
    If Me.anchors.count > 6 Then
        For i = 1 To Me.anchors.count - 6
            Application.CutCopyMode = False
            sheet.Rows(53).Copy
            sheet.Rows(54).Insert shift:=down
            Application.CutCopyMode = False
        Next i
    End If
    
    'Anchor info
    Dim Anchor As Anchor
    Dim anchorCount As Integer
    For Each Anchor In Me.anchors
        sheet.Range("ANCHOROWNER").offset(anchorCount, 0).Value = WorksheetFunction.Proper(Anchor.owner)
        sheet.Range("ANCHORLEAD").offset(anchorCount, 0).Value = Anchor.distance & "'"
        sheet.Range("ANCHORDIRECTION").offset(anchorCount, 0).Value = Anchor.direction & "(" & Anchor.angle & ")"
        anchorCount = anchorCount + 1
    Next Anchor
    
    'Equipment Section
    Dim equipment As equipment
    For Each equipment In Me.equipments
        If equipment.componentType = "XFMR" Or equipment.componentType = "CAPACITOR" Or equipment.componentType = "RECLOSER" Or equipment.componentType = "REGULATOR" Then
            sheet.Range("EQUIPMENT") = equipment.componentType
            sheet.Range("EQUIPMENTHEIGHT") = equipment.getDisplayHeight
            sheet.Range("EQUIPMENTSIZE") = equipment.size
            If equipment.componentType = "CO" Then sheet.Range("INVENTORY").Value = "CO"
            Exit For
        End If
    Next equipment
    
    'CMRF Section
    Call fillCMRF(sheet)
    
    'Make space for comms if needed
    If Me.commComponents.count > 11 Then
        For i = 1 To Me.commComponents.count - 11
            Application.CutCopyMode = False
            sheet.Rows(46).Copy
            sheet.Rows(47).Insert shift:=down
            Application.CutCopyMode = False
        Next i
    End If
    
    'Comm components section
    Set Me.commComponents = Utilities.sortComponents(Me.commComponents)
    For i = 0 To Me.commComponents.count - 1
        Set comp = Me.commComponents(i + 1)
        Dim category As String: category = ""
        If comp.componentType = "DROP" Then category = " SVC"
        If comp.componentType = "MSG" Then category = " MSG"
        sheet.Range("CMOWNER").offset(i, 0).Value = comp.owner & category
        
        sheet.Range("CMHEIGHT").offset(i, 0).Value = Utilities.inchesToFeetInches(comp.height)
        sheet.Range("CMSIZE").offset(i, 0).Value = comp.size
        If TypeOf comp Is Wire Then
            For j = 1 To maxSpans
                If comp.midspans.Exists(j) Then
                    Set Span = Me.findSpan(j)
                    sheet.Range("CMMIDSPAN" & j).offset(i, 0).Value = comp.getDisplayMidspan(Me.findSpan(j))
                    sheet.Range("CMMIDSPAN" & j).offset(i, 0).HorizontalAlignment = xlLeft
                Else
                    sheet.Range("CMMIDSPAN" & j).offset(i, 0).Value = "-"
                End If
            Next j
            
            If comp.modification > 0 And comp.height <> comp.modification Then
                Set modCell = findModCell(sheet, comp.height, comp.owner, True)
                If Not modCell Is Nothing Then
                    modCell.Value = Utilities.inchesToFeetInches(comp.modification)
                End If
            End If
            
        Else
            For j = 1 To maxSpans
                sheet.Range("CMMIDSPAN" & j).offset(i, 0).Value = "-"
            Next j
            If TypeOf comp Is guy Then
                sheet.Range("CMOWNER").offset(i, 0).Value = comp.owner & " DG"
            End If
        End If
    Next i
    
    'Make space for utilities if needed
    Call condensePrimaries
    If Me.utilComponents.count > 12 Then
        For i = 1 To Me.utilComponents.count - 12
            Application.CutCopyMode = False
            sheet.Rows(28).Copy
            sheet.Rows(29).Insert shift:=down
            Application.CutCopyMode = False
        Next i
    End If
    
    'Utility components section
    Set Me.utilComponents = Utilities.sortComponents(Me.utilComponents)
    Dim descCount As Integer
    For i = 0 To Me.utilComponents.count - 1
        Set comp = Me.utilComponents(i + 1)
        sheet.Range("UTTYPE").offset(i, 0).Value = comp.componentType
        sheet.Range("UTHEIGHT").offset(i, 0).Value = Utilities.inchesToFeetInches(comp.height)
        sheet.Range("UTSIZE").offset(i, 0).Value = comp.size
        If TypeOf comp Is Wire Then
            If comp.componentType = "PRI" And InStr(comp.size, "Ø") > 0 Then
                sheet.Range("DESC").offset(descCount, 0).Value = Mid(comp.size, InStr(comp.size, "Ø") - 1, 2)
                sheet.Range("DESC").offset(descCount, 0).HorizontalAlignment = xlLeft
                sheet.Range("SIZE").offset(descCount, 0).Value = Left(comp.size, Len(comp.size) - 3)
                sheet.Range("SIZE").offset(descCount, 0).HorizontalAlignment = xlLeft
                primaryDirections = ""
                For Each midspanSlot In comp.midspans
                    If primaryDirections <> "" Then primaryDirections = primaryDirections & "/"
                    primaryDirections = primaryDirections & Me.findSpan(midspanSlot).direction
                Next midspanSlot
                sheet.Range("DIRECTION").offset(descCount, 0).Value = primaryDirections
                descCount = descCount + 1
            End If
            For j = 1 To maxSpans
                If comp.midspans.Exists(j) Then
                    sheet.Range("UTMIDSPAN" & j).offset(i, 0).Value = comp.getDisplayMidspan(Me.findSpan(j))
                    sheet.Range("UTMIDSPAN" & j).offset(i, 0).HorizontalAlignment = xlLeft
                Else
                    sheet.Range("UTMIDSPAN" & j).offset(i, 0).Value = "-"
                End If
            Next j
        Else
            For j = 1 To maxSpans
                sheet.Range("UTMIDSPAN" & j).offset(i, 0).Value = "-"
            Next j
            sheet.Range("UTHEIGHT").offset(i, 0).Value = comp.getDisplayHeight()
        End If
        
        If comp.componentType = "SL" Then
            sheet.Range("STLTBRKT") = Utilities.inchesToFeetInches(comp.bottomHeight)
        ElseIf comp.size = "SL" Then
            sheet.Range("STLTDL") = Utilities.inchesToFeetInches(comp.height)
        End If
    Next i
    If sheet.Range("STLTBRKT") = "" Then sheet.Range("STLTBRKT") = "N/A"
    If sheet.Range("STLTDL") = "" Then sheet.Range("STLTDL") = "N/A"
    Call Me.setLowestPower
    sheet.Range("LWSTPWR") = Utilities.inchesToFeetInches(Me.lowestPower)
    
    'Misc info
    Call setLineStructureTypes
    If Me.Bonded <> "" Then sheet.Range("BONDED") = Me.Bonded
    sheet.Range("TANGENT") = Me.tangent
    sheet.Range("ANGLE") = Me.angle
    sheet.Range("CORNER") = Me.corner
    sheet.Range("DEADEND") = Me.deadend
    sheet.Range("BUCK") = Me.buck
    sheet.Range("SECONLY") = Me.secOnly
    If Me.treeWork Then
        If Me.notes <> "" Then Me.notes = Me.notes & vbLf
        Me.notes = Me.notes & "Tree work needed for pole."
    End If
    For Each Span In Me.spans
        If Span.treeWork Then
            If Me.notes <> "" Then Me.notes = Me.notes & vbLf
            Me.notes = Me.notes & "Tree work needed in " & Span.direction & " span."
        End If
    Next Span
    If Me.hammerTestFailed Then
        sheet.Range("DET").offset(0, 1).Value = "Hammer Test Failed"
        If Me.notes <> "" Then Me.notes = Me.notes & vbLf
        Me.notes = Me.notes & "Hammer test failed."
    End If
    If Me.visualCheckFailed Then
        If Me.notes <> "" Then Me.notes = Me.notes & vbLf
        Me.notes = Me.notes & "Fielder visual test failed."
    End If
    sheet.Range("NOTES") = Me.notes
    
    'Violations
    Call calculateViolations
    sheet.Range("SEP").offset(0, 1).Value = Me.violation40SZV
    sheet.Range("CLEARANCE").offset(0, 1).Value = Me.violationUtil156MSGCV
    sheet.Range("CLEARANCE").offset(1, 0).Value = Me.violationCom156MSGCV
    
    Dim separationViolations As Collection: Set separationViolations = New Collection
    separationViolations.Add Me.violation30MSSV
    separationViolations.Add Me.violation12CSV
    separationViolations.Add Me.violation6MSSV
    separationViolations.Add Me.violation40SLSV
    separationViolations.Add Me.violation12SLDLSV
    
    sheet.Range("SEP").offset(1, 0).Value = Utilities.JoinCollection(separationViolations, vbLf)
    
    sheet.Protect _
        Password:="", _
        DrawingObjects:=False, _
        contents:=True, _
        Scenarios:=False, _
        AllowFormattingCells:=True, _
        AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, _
        AllowInsertingColumns:=False, _
        AllowInsertingRows:=False, _
        AllowDeletingColumns:=False, _
        AllowDeletingRows:=False

End Sub

Public Sub calculateViolations()
    Me.violation40SZV = Calculate40SZV()
    Me.violationUtil156MSGCV = Calculate156MSGCV(Me.utilWires)
    Me.violationCom156MSGCV = Calculate156MSGCV(Me.commComponents)
    Me.violation30MSSV = Calculate30MSSV()
    results = CalculateCommViolations()
    violation12CSV = results(0)
    violation6MSSV = results(1)
    results = Calculate40SLSV()
    violation40SLSV = results(0)
    violation12SLDLSV = results(1)

End Sub

Private Function Calculate40SZV() As String
    Dim result As String
    Dim violators As Collection: Set violators = New Collection
    Dim comp As Variant
    
    For i = 1 To Me.commComponents.count
        Set comp = Me.commComponents(i)
        If TypeOf comp Is Wire Then
            If Me.lowestPower - 40 < comp.height Then violators.Add i
        End If
    Next i
    
    If violators.count > 0 Then
        result = "#" & Utilities.JoinCollection(violators, ",") & " < 40"" SZV"
    End If
    
    Calculate40SZV = result
End Function

Private Function Calculate156MSGCV(components As Collection) As String
    Dim result As String
    Dim violators As scripting.Dictionary: Set violators = New scripting.Dictionary
    Dim Wire As Variant
    Dim midspan As Integer
    Dim direction As String
    Dim coms As Boolean
    
    For i = 1 To components.count
        Set Wire = components(i)
        If TypeOf Wire Is Wire Then
            For Each midspanSlot In Wire.midspans
                If IsNumeric(Wire.midspans(midspanSlot)) Then midspan = Wire.midspans(midspanSlot)
                If Wire.componentType = "PRI" Or Wire.componentType = "NEUT" Or Wire.componentType = "SEC" Or Wire.componentType = "OW" Or Wire.componentType = "SVC" Then
                    If midspan > 0 And midspan < 186 Then
                        direction = Me.findSpan(midspanSlot).direction
                        If Not violators.Exists(direction) Then violators.Add direction, New Collection
                        violators(direction).Add Wire.componentType
                    End If
                ElseIf Wire.componentType <> "SPG" Then
                    If midspan > 0 And midspan < 186 Then
                        direction = Me.findSpan(midspanSlot).direction
                        If Not violators.Exists(direction) Then violators.Add direction, New Collection
                        violators(direction).Add i
                        coms = True
                    End If
                End If
            Next midspanSlot
        End If
    Next i
    
    For Each key In violators
        If result <> "" Then result = result & ", "
        If coms Then result = result & "#"
        result = result & Utilities.JoinCollection(violators(key), ",") & " (" & key & ")"
    Next key
    If result <> "" Then result = result & " < 15'6"" MIDSPAN GROUND CLEARANCE VIOLATION"
    
    Calculate156MSGCV = result
End Function

Private Function Calculate30MSSV() As String
    Dim result As String
    Dim violators As scripting.Dictionary: Set violators = New scripting.Dictionary
    Dim Wire As Variant
    Dim lowestMidspan As Integer
    Dim comMidspan As Integer
    Dim direction As String
    
    Dim Span As Span
    For Each Span In Me.spans
        If Not Span.lowestPowerWire Is Nothing Then
            lowestMidspan = Span.lowestPowerWire.midspans(Span.spanSlot)
            For i = 1 To Me.commComponents.count
                Set Wire = Me.commComponents(i)
                If TypeOf Wire Is Wire Then
                    If Wire.midspans.Exists(Span.spanSlot) Then
                        comMidspan = Wire.midspans(Span.spanSlot)
                        If lowestMidspan - 30 < comMidspan Then
                            direction = Span.direction
                            If Not violators.Exists(direction) Then violators.Add direction, New Collection
                            violators(direction).Add i
                        End If
                    End If
                End If
            Next i
        End If
    Next Span
    
    For Each key In violators
        If result <> "" Then result = result & ", "
        result = result & "#"
        result = result & Utilities.JoinCollection(violators(key), ",") & "(" & key & ")"
    Next key
    If result <> "" Then result = result & " < 30"" MIDSPAN SEPARATION VIOLATION"
    
    Calculate30MSSV = result
End Function

Private Function CalculateCommViolations() As String()
    Dim results(0 To 1) As String
    Dim separationViolators As scripting.Dictionary: Set separationViolators = New scripting.Dictionary
    Dim midspanViolators As scripting.Dictionary: Set midspanViolators = New scripting.Dictionary
    Dim wire1 As Variant
    Dim wire2 As Variant
    Dim Midspan1 As Integer
    Dim Midspan2 As Integer
    Dim direction As String
    
    For i = 1 To Me.commComponents.count
        For j = i + 1 To Me.commComponents.count
            Set wire1 = Me.commComponents(i)
            Set wire2 = Me.commComponents(j)
            If TypeOf wire1 Is Wire And TypeOf wire2 Is Wire Then
                If (Abs(wire1.height - wire2.height) < 12 And wire1.owner <> wire2.owner) Then
                    If Not separationViolators.Exists(i) Then separationViolators.Add i, Nothing
                    If Not separationViolators.Exists(j) Then separationViolators.Add j, Nothing
                ElseIf Abs(wire1.height - wire2.height) > 0 And Abs(wire1.height - wire2.height) < 4 And wire1.owner = wire2.owner And Utilities.ShareAKey(wire1.midspans, wire2.midspans) Then
                    If Not separationViolators.Exists(i) Then separationViolators.Add i, Nothing
                    If Not separationViolators.Exists(j) Then separationViolators.Add j, Nothing
                End If
                
                For Each midspanSlot In wire1.midspans
                    If wire2.midspans.Exists(midspanSlot) Then
                        Midspan1 = wire1.midspans(midspanSlot)
                        Midspan2 = wire2.midspans(midspanSlot)
                        If Midspan1 <> 0 And Midspan2 <> 0 And Abs(Midspan1 - Midspan2) < 6 Then
                            direction = Me.findSpan(midspanSlot).direction
                            If Not midspanViolators.Exists(direction) Then midspanViolators.Add direction, New Collection
                            midspanViolators(direction).Add "(#" & i & "," & j & ")"
                        End If
                    End If
                Next midspanSlot
            End If
        Next j
    Next i

    If separationViolators.count > 0 Then
        results(0) = "#" & Utilities.JoinCollection(separationViolators, ",") & " < 12"" COM POLE SEPARATION VIOLATION"
    End If
    
    For Each key In midspanViolators
        If results(1) <> "" Then results(1) = results(1) & ", "
        results(1) = results(1) & Utilities.JoinCollection(midspanViolators(key), ",") & "(" & key & ")"
    Next key
    If results(1) <> "" Then results(1) = results(1) & " < 6"" COMM MID SPAN SEPARATION VIOLATION"
    
    CalculateCommViolations = results
End Function

Public Function Calculate40SLSV() As String()
    Dim results(0 To 1) As String
    Dim sLBracket As equipment
    Dim SLDL As equipment
    Dim equipment As equipment
    Dim Wire As Variant
    
    Dim sLBracketViolators As Collection: Set sLBracketViolators = New Collection
    Dim sLDLViolators As Collection: Set sLDLViolators = New Collection
    
    For Each equipment In Me.equipments
        If equipment.componentType = "SL" Then Set sLBracket = equipment
        If equipment.size = "SL" Then Set SLDL = equipment
    Next equipment
    
    If sLBracket Is Nothing And SLDL Is Nothing Then
        Calculate40SLSV = results
        Exit Function
    End If
    
    For i = 1 To Me.commComponents.count
        Set Wire = Me.commComponents(i)
        
        If TypeOf Wire Is Wire Then
            If sLBracket.bottomHeight - 40 < Wire.height Then sLBracketViolators.Add i
            If Not SLDL Is Nothing Then
                If Abs(SLDL.height - Wire.height) < 12 Then sLDLViolators.Add i
            End If
        End If
    Next i
    
    If sLBracketViolators.count > 0 Then
        results(0) = "#" & Utilities.JoinCollection(sLBracketViolators, ",") & " < 40"" STREETLIGHT BOTTOM BRACKET SEPARATION VIOLATION (ASSUMING NOT BONDED)"
    End If
    
    If sLDLViolators.count > 0 Then
        results(1) = "#" & Utilities.JoinCollection(sLDLViolators, ",") & " < 12"" STREETLIGHT DRIP LOOP SEPARATION VIOLATION"
    End If
    
    Calculate40SLSV = results
End Function

Public Sub clearCMRF(sheet As Worksheet)
    For i = 1 To 8
        sheet.Range("COMM" & i).Value = "COMM #" & i
        sheet.Range("COMM" & i).HorizontalAlignment = xlCenter
        sheet.Range("COMM" & i).offset(2, 0).Value = ""
        sheet.Range("COMM" & i).offset(2, 0).offset(0, 1).Value = ""
    Next i
    
    Dim row As Integer: row = sheet.Range("COMM1").row + 4
    For i = 0 To 50
        sheet.Rows(row).Delete
    Next i
End Sub

Public Sub fillCMRF(sheet As Worksheet)
    Dim Wire As Wire
    Dim comms As scripting.Dictionary
    Set comms = New scripting.Dictionary
    
    Set Me.commWires = Utilities.sortComponents(Me.commWires)
    
    For Each Wire In Me.commWires
        If Not comms.Exists(Wire.owner) Then comms.Add Wire.owner, New Collection
        comms(Wire.owner).Add Wire
    Next Wire
    
    Dim commOwnerCount As Integer
    For Each commOwner In comms
        commOwnerCount = commOwnerCount + 1
        sheet.Range("COMM" & commOwnerCount).Value = commOwner
        sheet.Range("COMM" & commOwnerCount).HorizontalAlignment = xlLeft
        commCount = 0
        
        'Nondrops first
        For i = 1 To comms(commOwner).count
            Set Wire = comms(commOwner)(i)
            If Wire.componentType <> "DROP" Then
                commCount = commCount + 1
                If commCount > 1 Then
                    Set SourceRange = sheet.Range(sheet.Range("COMM" & commOwnerCount).offset(2, -1), sheet.Range("COMM" & commOwnerCount).offset(3, 1))
                    SourceRange.Copy
                    sheet.Range("COMM" & commOwnerCount).offset(2, -1).offset(((commCount - 1) * 2) - 1, 0).PasteSpecial Paste:=xlPasteAll
                End If
                sheet.Range("COMM" & commOwnerCount).offset(3, 0).offset(((commCount - 1) * 2) - 1, 0).Value = Wire.getDisplayHeight(True)
            End If
        Next i
        
        'Drops after
        For i = 1 To comms(commOwner).count
            Set Wire = comms(commOwner)(i)
            If Wire.componentType = "DROP" Then
                commCount = commCount + 1
                If commCount > 1 Then
                    Set SourceRange = sheet.Range(sheet.Range("COMM" & commOwnerCount).offset(2, -1), sheet.Range("COMM" & commOwnerCount).offset(3, 1))
                    SourceRange.Copy
                    sheet.Range("COMM" & commOwnerCount).offset(2, -1).offset(((commCount - 1) * 2) - 1, 0).PasteSpecial Paste:=xlPasteAll
                End If
                sheet.Range("COMM" & commOwnerCount).offset(3, 0).offset(((commCount - 1) * 2) - 1, 0).Value = Wire.getDisplayHeight(True)
            End If
        Next i
    Next commOwner
End Sub

Public Sub setLowestPower()
    For Each comp In Me.utilComponents
        If comp.componentType <> "SL" And comp.size <> "SL" And comp.componentType <> "DG" And comp.componentType <> "SPG" Then
            Me.lowestPower = comp.height
        End If
    Next comp
End Sub

Public Sub setLineStructureTypes()
    Dim comp As Variant
    Dim primaryAngles As Collection: Set primaryAngles = New Collection
    
    For Each comp In Me.utilComponents
        If comp.componentType = "PRI" Then
            For Each midspanSlot In comp.midspans
                primaryAngles.Add Me.findSpan(midspanSlot).angle
            Next midspanSlot
        End If
    Next comp
    
    If primaryAngles.count = 0 Then Me.secOnly = True
    If primaryAngles.count Mod 2 <> 0 Then Me.deadend = True
    
    For i = 1 To primaryAngles.count
        For j = i + 1 To primaryAngles.count
            angle1 = primaryAngles(i)
            angle2 = primaryAngles(j)
            If Abs((angle1 + angle2) - 180) <= 4 Then
                Me.tangent = True
            ElseIf Abs(Abs(angle1 - angle2) - 90) <= 4 Then
                Me.corner = True
            Else
                Me.angle = True
            End If
        Next j
    Next i
End Sub

Public Function findSpan(ByVal midspanSlot As Integer) As Span
    For Each Span In Me.spans
        If Span.spanSlot = midspanSlot Then
            Set findSpan = Span
            Exit Function
        End If
    Next Span
    Set findSpan = Nothing
End Function

Public Function findSpanBySpanId(ByVal spanId As String) As Span
    For Each Span In Me.spans
        If Span.spanId = spanId Then
            Set findSpanBySpanId = Span
            Exit Function
        End If
    Next Span
    Set findSpanBySpanId = Nothing
End Function

Public Function findWireByTrace(trace_ As String, size_ As String) As Wire
    Dim Wire As Wire
    Dim foundWire As Wire: Set foundWire = Nothing
    For Each Wire In wires
        If Wire.trace = trace_ And Wire.size = size_ Then
            Set foundWire = Wire
            Exit For
        ElseIf Wire.trace = trace_ And Wire.size = "DROP" Then
            Set foundWire = Wire
            Exit For
        ElseIf Wire.trace = trace_ Then
            Set foundWire = Wire
        End If
    Next Wire
    
    Set findWireByTrace = foundWire
End Function

Public Sub condensePrimaries()
    Dim comp As Variant
    Dim primaryBundles As scripting.Dictionary: Set primaryBundles = New scripting.Dictionary
    Dim i As Integer
    Dim Wire As Wire
    Dim bundleKey As Variant
    Dim midspanSlot As Variant
    Dim phase As Integer
    Dim primary As Wire
    
    For i = 1 To 12
        phase = 0
        For Each comp In Me.utilWires
            If comp.componentType = "PRI" And InStr(comp.size, "Ø") = 0 Then
                If comp.midspans.Exists(i) Then
                    phase = phase + 1
                    Set primary = comp
                End If
            End If
        Next comp
        If phase > 0 Then
            bundleKey = primary.size & " " & phase & "Ø" & primary.height
            If Not primaryBundles.Exists(bundleKey) Then primaryBundles.Add bundleKey, Array(primary, New Collection)
            primaryBundles(bundleKey)(1).Add i
        End If
    Next i
    
    For Each bundleKey In primaryBundles
        Set Wire = New Wire
        Set comp = primaryBundles(bundleKey)(0)
        
        Wire.size = Left(bundleKey, InStr(bundleKey, "Ø"))
        Wire.componentType = comp.componentType
        Wire.height = comp.height
        For Each midspanSlot In primaryBundles(bundleKey)(1)
            Wire.midspans.Add midspanSlot, "PRI"
        Next midspanSlot
        
        Me.utilComponents.Add Wire
    Next bundleKey
End Sub
