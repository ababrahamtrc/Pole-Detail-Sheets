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
Public x As Double
Public y As Double
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
Public permit As String
Public tlm As String
Public locationAdjacent As Boolean

Public primaries As Collection
Public neutrals  As Collection
Public secondaries  As Collection
Public openWires As Collection
Public services  As Collection

Public slBottomBracketHeight As Integer

Public applicant As wire
Public applicantGuys As Collection
Public overlash As Boolean

Public topped As Boolean
Public outage As Boolean

Public location As String
Public alt1 As String
Public alt2 As String
Public alt3 As String
Public NJUNS As String
Public njunsSteps As Collection
Public njunsTicket As String
Public ttc As String
Public newHeight As Integer
Public newClass As Integer
Public dSpace As Integer
Public pSpace As Integer

Public ReplaceAnchor As Boolean
Public InstallAnchor As Boolean
Public ReplacePole As Boolean
Public InstallPole As Boolean
Public ReplaceRiser As Boolean
Public RemoveAnchor As Boolean
Public RemovePole As Boolean

Public firstCrossStreet As String

Public groupNumber As Integer
Public skipSpan As Boolean

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

Public transformerSizes As Integer

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
    Set applicant = New wire
    
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
    If sheet.Range("PROPOSEDHEIGHT").Value <> "" Then Me.applicant.height = Utilities.convertToInches(Split(sheet.Range("PROPOSEDHEIGHT").Value, vbLf)(0))
    If sheet.Range("CMRF1").Value <> "" Then Me.applicant.modification = Utilities.convertToInches(Split(sheet.Range("CMRF1").Value, vbLf)(0))
    
    If sheet.Range("LWSTPWR").Value <> "" Then Me.lowestPower = Utilities.convertToInches(Split(sheet.Range("LWSTPWR").Value, vbLf)(0))
    If sheet.Range("STLTBRKT").Value <> "" Then Me.slBottomBracketHeight = Utilities.convertToInches(Split(sheet.Range("STLTBRKT").Value, vbLf)(0))
    Me.Bonded = sheet.Range("BONDED")
    
    For i = 1 To 12
        If Not Utilities.RangeExists(sheet, "TOPOLE" & i) Then Exit For
        If sheet.Range("TOPOLE" & i).OFFSET(1, 0).Value <> "" Then
            applicant.midspans.Add i, Utilities.convertToInches(sheet.Range("TOPOLE" & i).OFFSET(1, 0).Value)
        End If
    Next i
    
    Me.height = Utilities.convertToInches(ThisWorkbook.RemoveParentheses(sheet.Range("HEIGHT")) & "'")
    Me.Class = sheet.Range("CLASS")
    Me.address = sheet.Range("ADDRESS")
    
    
    
    Me.ReplaceAnchor = sheet.Range("REPLACEANCHOR")
    Me.InstallAnchor = sheet.Range("INSTALLANCHOR")
    Me.ReplacePole = sheet.Range("REPLACEPOLE")
    Me.InstallPole = sheet.Range("INSTALLPOLE")
    Me.ReplaceRiser = sheet.Range("REPLACERISER")
    Me.RemoveAnchor = sheet.Range("REMOVEANCHOR")
    Me.RemovePole = sheet.Range("REMOVEPOLE")
    
    If Utilities.RangeExists(sheet, "FCS") Then Me.firstCrossStreet = sheet.Range("FCS").OFFSET(0, 1)
    
    Me.location = sheet.Range("DL")
    Me.alt1 = sheet.Range("ALTONE")
    Me.alt2 = sheet.Range("ALTTWO")
    Me.alt3 = sheet.Range("ALTTHREE")
    Me.ttc = sheet.Range("TTC")
    Me.NJUNS = sheet.Range("NJUNS")
    If Me.NJUNS <> "" Then Set Me.njunsSteps = Utilities.SplitByBlankLine(Me.NJUNS)
    Me.njunsTicket = sheet.Range("NJUNSTICKET")
    Me.latitude = sheet.Range("LATITUDE")
    Me.longitude = sheet.Range("LONGITUDE")
    results = LatLonToMI2253(Me.latitude, Me.longitude)
    Me.x = results(0)
    Me.y = results(1)
    
    Me.treeWork = sheet.Range("TREE") Or sheet.Range("TREE2") = "YES"
    
    If Utilities.RangeExists(sheet, "TLM") Then Me.tlm = sheet.Range("TLM")
    Me.outage = sheet.Range("OUTAGE") Or InStr(Me.alt1, "OUTAGE REQUIRED") > 0
    
    Me.buck = sheet.Range("BUCK")
    
    Me.newHeight = Me.height
    Me.dSpace = 0
    Me.pSpace = 0
    
    Dim regex As Object: Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\d[0:5]\s*-\s*\d\s*\/\s*(\d[0:5])\s*-\s*(\d)"
    regex.Global = True
    regex.IgnoreCase = True
    
    lines = Split(Me.alt1, vbLf)
    For Each line In lines
        line = Replace(line, " ", "")
        If InStr(line, "D=") > 0 Then Me.dSpace = Utilities.convertToInches(Mid(line, 3, Len(line) - 2))
        If InStr(line, "P=") > 0 Then Me.pSpace = Utilities.convertToInches(Mid(line, 3, Len(line) - 2))
        If regex.test(line) Then
            Me.newHeight = regex.Execute(line)(0).SubMatches(0) * 12
            Me.newClass = regex.Execute(line)(0).SubMatches(1)
            Me.ReplacePole = True
        End If
        
        If InStr(line, "TOP") > 0 And InStr(line, "POLE") > 0 And InStr(line, "ABOVE") > 0 Then Me.topped = True
    Next line
    
    If InStr(sheet.Range("PROPOSEDHEIGHT"), "OL") > 0 Then Me.overlash = True
    
    Dim Span As Span
    For i = 1 To 12
        If Not Utilities.RangeExists(sheet, "TOPOLE" & i) Then Exit For
        'If sheet.Range("TOPOLE" & i) = "" Or InStr(sheet.Range("TOPOLE" & i), "(") = 0 Or InStr(sheet.Range("TOPOLE" & i), ")") = 0 Then Exit For
        If sheet.Range("TOPOLE" & i) <> "" Then
            Set Span = New Span
            Span.distance = Utilities.OnlyNumbers(sheet.Range("SPAN" & i))
            Dim tempAngle As String: tempAngle = sheet.Range("TOPOLE" & i)
            If InStr(tempAngle, "(") > 0 And InStr(tempAngle, "(") > 0 Then tempAngle = ThisWorkbook.InsideParentheses(tempAngle)
            If InStr(tempAngle, "(") > 0 Then tempAngle = Right(tempAngle, Len(tempAngle) - InStr(tempAngle, "("))
            Span.angle = CInt(tempAngle)
            If Utilities.SheetExists(ThisWorkbook.RemoveParentheses(sheet.Range("TOPOLE" & i))) Then
                Span.otherPole = ThisWorkbook.RemoveParentheses(sheet.Range("TOPOLE" & i))
            End If
            Span.direction = Utilities.degreesToDirection(Span.angle)
            Span.spanSlot = i
            Me.spans.Add Span
        End If
    Next i
    
    Dim comp As Component
    Dim wire As wire
    Dim Equipment As Equipment
    Dim guy As guy
    Dim modificationCells As Collection: Set modificationCells = New Collection
    For i = 0 To 100
        If sheet.Range("CMOWNER").OFFSET(i, 0).Interior.color <> 16312794 Then Exit For
        If Utilities.convertToInches(sheet.Range("CMHEIGHT").OFFSET(i, 0)) > 0 Then
            Set comp = New Component
            comp.height = Utilities.convertToInches(sheet.Range("CMHEIGHT").OFFSET(i, 0).text)
            comp.owner = sheet.Range("CMOWNER").OFFSET(i, 0).text
            comp.size = sheet.Range("CMSIZE").OFFSET(i, 0).text
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
                Set wire = comp.getWire()
                If overlash And wire.owner = applicant.owner Then applicant.height = wire.height
                
                For Each Span In Me.spans
                    If Replace(sheet.Range("CMMIDSPAN" & Span.spanSlot).OFFSET(i, 0).text, "-", "") <> "" Then
                        wire.midspans.Add Span.spanSlot, Utilities.convertToInches(sheet.Range("CMMIDSPAN" & Span.spanSlot).OFFSET(i, 0).text)
                        If wire.componentType = "" Then wire.componentType = "COM"
                        Span.commComponents.Add wire
                        Span.commWires.Add wire
                        Span.wires.Add wire
                        If Span.houseNumber = "" Then
                            If InStr(sheet.Range("CMMIDSPAN" & Span.spanSlot).OFFSET(i, 0), "(") > 0 Then
                                Span.houseNumber = ThisWorkbook.InsideParentheses(sheet.Range("CMMIDSPAN" & Span.spanSlot).OFFSET(i, 0))
                            Else
                                Span.houseNumber = sheet.Range("CMMIDSPAN" & Span.spanSlot).OFFSET(i, 0)
                            End If
                        End If
                    End If
                Next Span
                
                Set modCell = findModCell(sheet, wire.height, wire.owner, False, modificationCells)
                If Not modCell Is Nothing Then wire.modification = Utilities.convertToInches(modCell)
                
                If wire.componentType = "DROP" Or wire.componentType = "MSG" Or wire.componentType = "COM" Then
                    Me.commComponents.Add wire
                    Me.commWires.Add wire
                    Me.wires.Add wire
                Else
                    Set Equipment = comp.getEquipment
                    Equipment.bottomHeight = Utilities.convertToInches(ThisWorkbook.InsideParentheses(sheet.Range("CMHEIGHT").OFFSET(i, 0).text))
                    Me.commComponents.Add Equipment
                End If
            Else
                Set guy = comp.getDownGuy()
                Me.commComponents.Add guy
                Me.guys.Add guy
            End If
        End If
    Next i
    
    For i = 0 To 100
        If sheet.Range("UTTYPE").OFFSET(i, 0).Interior.color <> 16312794 Then Exit For
        If Utilities.convertToInches(ThisWorkbook.RemoveParentheses(sheet.Range("UTHEIGHT").OFFSET(i, 0))) > 0 Then
            Set comp = New Component
            comp.height = Utilities.convertToInches(sheet.Range("UTHEIGHT").OFFSET(i, 0).text)
            comp.componentType = sheet.Range("UTTYPE").OFFSET(i, 0).text
            If InStr(comp.componentType, ")") > 0 And InStr(comp.componentType, "(") > 0 Then comp.componentType = ThisWorkbook.InsideParentheses(comp.componentType)
            comp.size = sheet.Range("UTSIZE").OFFSET(i, 0).text
            If InStr(comp.size, ")") > 0 And InStr(comp.size, "(") > 0 Then comp.size = ThisWorkbook.RemoveParentheses(comp.size)
            
            If comp.componentType = "PRI" Or comp.componentType = "NEUT" Or comp.componentType = "SEC" Or comp.componentType = "SVC" Or comp.componentType = "OW" Or comp.componentType = "SPG" Then
                Set wire = comp.getWire()
                
                If wire.componentType = "PRI" And InStr(wire.size, "Ø") > 0 Then
                    wire.phase = Mid(comp.size, Len(wire.size) - 1, 1)
                    wire.size = Left(comp.size, Len(wire.size) - 3)
                    If wire.phase > 3 Then wire.phase = 3
                    If wire.phase <> 1 And wire.phase <> 2 And wire.phase <> 3 Then
                        wire.phase = 1
                    End If
                ElseIf wire.componentType = "PRI" Then
                    wire.size = comp.size
                    wire.phase = 1
                End If
                If wire.componentType = "SPG" Or wire.componentType = "MSG" Then wire.wepHeight = Utilities.convertToInches(ThisWorkbook.InsideParentheses(sheet.Range("UTHEIGHT").OFFSET(i, 0).text))
                
                For Each Span In Me.spans
                    If Replace(sheet.Range("UTMIDSPAN" & Span.spanSlot).OFFSET(i, 0).text, "-", "") <> "" Then
                        wire.midspans.Add Span.spanSlot, Utilities.convertToInches(sheet.Range("UTMIDSPAN" & Span.spanSlot).OFFSET(i, 0).text)
                        Span.utilComponents.Add wire
                        Span.utilWires.Add wire
                        Span.wires.Add wire
                        If Span.houseNumber = "" Then
                            If InStr(sheet.Range("UTMIDSPAN" & Span.spanSlot).OFFSET(i, 0), "(") > 0 Then
                                Span.houseNumber = ThisWorkbook.InsideParentheses(sheet.Range("UTMIDSPAN" & Span.spanSlot).OFFSET(i, 0))
                            Else
                                Span.houseNumber = sheet.Range("UTMIDSPAN" & Span.spanSlot).OFFSET(i, 0)
                            End If
                        End If
                    End If
                Next Span
                
                If wire.componentType = "PRI" Then Me.primaries.Add wire
                If wire.componentType = "NEUT" Then Me.neutrals.Add wire
                If wire.componentType = "SEC" Then Me.secondaries.Add wire
                If wire.componentType = "OW" Then Me.openWires.Add wire
                If wire.componentType = "SVC" Then Me.services.Add wire
                
                Me.utilComponents.Add wire
                Me.utilWires.Add wire
                Me.wires.Add wire
            ElseIf comp.componentType = "DG" Then
                Set guy = comp.getDownGuy()
                Me.utilComponents.Add guy
                Me.guys.Add guy
            Else
                Set Equipment = comp.getEquipment
                Equipment.bottomHeight = Utilities.convertToInches(ThisWorkbook.InsideParentheses(sheet.Range("UTHEIGHT").OFFSET(i, 0).text))
                Me.utilComponents.Add Equipment
                Me.equipments.Add Equipment
                If Equipment.componentType = "XFMR" Then Me.transformerSizes = Me.transformerSizes + Utilities.OnlyNumbers(Equipment.size)
            End If
        End If
    Next i
    
    Dim anchor As anchor
    For i = 0 To 100
        Set anchor = New anchor
        If sheet.Range("ANCHOROWNER").OFFSET(i, 0).Value = "" Or sheet.Range("ANCHOROWNER").OFFSET(i, 0).Value = "New APP" Then Exit For
        anchor.owner = sheet.Range("ANCHOROWNER").OFFSET(i, 0).Value
        anchor.distance = Utilities.OnlyNumbers(sheet.Range("ANCHORLEAD").OFFSET(i, 0).Value)
        anchor.angle = Utilities.OnlyNumbers(sheet.Range("ANCHORDIRECTION").OFFSET(i, 0).Value)
        anchor.direction = Utilities.degreesToDirection(anchor.angle)
        
        If InStr(anchor.owner, "(x") > 0 And InStr(anchor.owner, ")") > 0 And IsNumeric(Mid(anchor.owner, InStr(anchor.owner, "(x") + 2, 1)) Then
            anchor.ceCount = CInt(Mid(anchor.owner, InStr(anchor.owner, "(x") + 2, 1))
            anchor.owner = Right(anchor.owner, Len(anchor.owner) - InStr(anchor.owner, ")"))
        ElseIf anchor.owner = "Consumers Energy" Then
            anchor.ceCount = 1
        End If
        
        Me.anchors.Add anchor
    Next i
    
    If Me.location <> "" Then
        Me.locationAdjacent = True
    Else
        For Each Span In Me.spans
            If Utilities.SheetExists(Span.otherPole) Then
                If Utilities.GetPDS(Span.otherPole).Range("DL") <> "" Then
                    Me.locationAdjacent = True
                    Exit For
                End If
            End If
        Next Span
    End If
    
End Sub

Public Sub createSheet(project As project, Optional selectedTemplate As Integer)
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
    Set existingNames = New Scripting.Dictionary
    
    ' copy the template
    For Each sheet In ThisWorkbook.sheets
        rpName = ThisWorkbook.RemoveParentheses(sheet.name)
        existingNames(rpName) = True
        If rpName <> "4 Spans" And rpName <> "8 Spans" And rpName <> "12 Spans" Then
            If sheet.Cells(2, 2).Value = "Notification:" Then
                If Val(rpName) > Val(poleNumber) Or (Val(rpName) = Val(poleNumber) And rpName > poleNumber) Then
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
        If Not existingNames.exists(ThisWorkbook.RemoveParentheses(sheet.name)) Then
            Set pdsWs = sheet
            pdsWs.visible = xlSheetVisible
            If Me.poleNumber <> "" Then pdsWs.name = Me.poleNumber
            Exit For
        End If
    Next sheet
    
    Call Me.fillSheet(pdsWs, maxSpans, project)
    sheet.Activate
    sheet.Range("A1").Select
End Sub

Public Sub fillSheet(sheet As Worksheet, maxSpans As Integer, project As project)
    Dim comp As Variant

    'Header info
    sheet.Unprotect
    sheet.Range("NOTIFICATION").Value = project.Notification
    sheet.Range("PERMIT").Value = project.permit
    sheet.Range("APPLICANT").Value = project.applicant
    sheet.Range("LOC").Value = project.jobLocation
    sheet.Range("INSPECTEDBY").Value = project.fielder
    sheet.Range("TWP").Value = project.township
    sheet.Range("DATE").Value = Me.collectedDate
    sheet.Range("COUNTY").Value = project.county
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
        sheet.Range("NEWAPPSIZE").OFFSET(i - 1, 0) = guy.proposedSize
        sheet.Range("NEWAPPLEAD").OFFSET(i - 1, 0) = guy.proposedLead
        sheet.Range("NEWAPPDIR").OFFSET(i - 1, 0) = guy.proposedDirection
    Next i
    If Me.applicant.height > 0 Then sheet.Range("PROPOSEDHEIGHT").Value = Utilities.inchesToFeetInches(Me.applicant.height)
    If Me.applicant.modification <> Me.applicant.height And Me.applicant.modification > 0 Then sheet.Range("CMRF1").Value = Utilities.inchesToFeetInches(Me.applicant.modification)
    sheet.Range("PROPOSEDDIAMETER") = Me.applicant.diameter
    
    'Span info
    Dim Span As Span
    For Each Span In Me.spans
        sheet.Range("SPAN" & Span.spanSlot).Value = Round(Span.distance, 1) & "'"
        sheet.Range("SPAN" & Span.spanSlot).HorizontalAlignment = xlLeft
        sheet.Range("TOPOLE" & Span.spanSlot).Value = Span.getToSpan()
        sheet.Range("TOPOLE" & Span.spanSlot).HorizontalAlignment = xlLeft
        sheet.Range("CLEARANCE" & Span.spanSlot).Value = "15'6"""
        If Me.applicant.midspans.exists(Span.spanSlot) Then
            sheet.Range("TOPOLE" & Span.spanSlot).OFFSET(1, 0).Value = Utilities.inchesToFeetInches(Me.applicant.midspans(Span.spanSlot))
            If Me.applicant.tensions.exists(Span.spanSlot) Then sheet.Range("TOPOLE" & Span.spanSlot).OFFSET(2, 0).Value = Me.applicant.tensions(Span.spanSlot)
        End If
        Call Span.setLowestPowerWire
        If Not Span.lowestPowerWire Is Nothing Then
            If Span.lowestPowerWire.componentType = "SVC" Then
                sheet.Range("CMMIDSPAN" & Span.spanSlot).OFFSET(-2, 0).Value = Utilities.inchesToFeetInches(Span.lowestPowerWire.midspans(Span.spanSlot) - 12)
            Else
                sheet.Range("CMMIDSPAN" & Span.spanSlot).OFFSET(-2, 0).Value = Utilities.inchesToFeetInches(Span.lowestPowerWire.midspans(Span.spanSlot) - 30)
            End If
        End If
    Next Span
    
    'Make space for anchors if needed
    If Me.anchors.count > 6 Then
        For i = 1 To Me.anchors.count - 6
            Application.CutCopyMode = False
            sheet.Rows(53).Copy
            sheet.Rows(54).Insert shift:=DOWN
            Application.CutCopyMode = False
        Next i
    End If
    
    'Anchor info
    Dim anchor As anchor
    Dim anchorCount As Integer
    For Each anchor In Me.anchors
        sheet.Range("ANCHOROWNER").OFFSET(anchorCount, 0).Value = WorksheetFunction.Proper(anchor.owner)
        If anchor.ceCount > 1 Then sheet.Range("ANCHOROWNER").OFFSET(anchorCount, 0).Value = "(x" & anchor.ceCount & ")" & sheet.Range("ANCHOROWNER").OFFSET(anchorCount, 0).Value
        sheet.Range("ANCHORLEAD").OFFSET(anchorCount, 0).Value = anchor.distance & "'"
        sheet.Range("ANCHORDIRECTION").OFFSET(anchorCount, 0).Value = anchor.direction & "(" & anchor.angle & ")"
        anchorCount = anchorCount + 1
    Next anchor
    
    'Equipment Section
    Dim Equipment As Equipment
    For Each Equipment In Me.equipments
        If Equipment.componentType = "XFMR" Or Equipment.componentType = "CAPACITOR" Or Equipment.componentType = "RECLOSER" Or Equipment.componentType = "REGULATOR" Then
            sheet.Range("EQUIPMENT") = Equipment.componentType
            sheet.Range("EQUIPMENTHEIGHT") = Equipment.getDisplayHeight
            sheet.Range("EQUIPMENTSIZE") = Equipment.size
            If Equipment.componentType = "CO" Then sheet.Range("INVENTORY").Value = "CO"
            Exit For
        End If
    Next Equipment
    
    'CMRF Section
    Call fillCMRF(sheet)
    
    'Make space for comms if needed
    If Me.commComponents.count > 11 Then
        For i = 1 To Me.commComponents.count - 11
            Application.CutCopyMode = False
            sheet.Rows(46).Copy
            sheet.Rows(47).Insert shift:=DOWN
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
        sheet.Range("CMOWNER").OFFSET(i, 0).Value = comp.owner & category
        
        sheet.Range("CMHEIGHT").OFFSET(i, 0).Value = Utilities.inchesToFeetInches(comp.height)
        sheet.Range("CMSIZE").OFFSET(i, 0).Value = comp.size
        If TypeOf comp Is wire Then
            If comp.wepHeight > 0 Then sheet.Range("CMHEIGHT").OFFSET(i, 0).Value = sheet.Range("CMHEIGHT").OFFSET(i, 0).Value & "(" & Utilities.inchesToFeetInches(comp.wepHeight) & ")"
            For j = 1 To maxSpans
                If comp.midspans.exists(j) Then
                    Set Span = Me.findSpan(j)
                    sheet.Range("CMMIDSPAN" & j).OFFSET(i, 0).Value = comp.getDisplayMidspan(Me.findSpan(j))
                    sheet.Range("CMMIDSPAN" & j).OFFSET(i, 0).HorizontalAlignment = xlLeft
                Else
                    sheet.Range("CMMIDSPAN" & j).OFFSET(i, 0).Value = "-"
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
                sheet.Range("CMMIDSPAN" & j).OFFSET(i, 0).Value = "-"
            Next j
            If TypeOf comp Is guy Then
                sheet.Range("CMOWNER").OFFSET(i, 0).Value = comp.owner & " DG"
            End If
        End If
    Next i
    
    'Make space for utilities if needed
    Call condensePrimaries
    If Me.utilComponents.count > 12 Then
        For i = 1 To Me.utilComponents.count - 12
            Application.CutCopyMode = False
            sheet.Rows(28).Copy
            sheet.Rows(29).Insert shift:=DOWN
            Application.CutCopyMode = False
        Next i
    End If
    
    'Utility components section
    Set Me.utilComponents = Utilities.sortComponents(Me.utilComponents)
    Dim descCount As Integer
    For i = 0 To Me.utilComponents.count - 1
        Set comp = Me.utilComponents(i + 1)
        sheet.Range("UTTYPE").OFFSET(i, 0).Value = comp.componentType
        sheet.Range("UTHEIGHT").OFFSET(i, 0).Value = Utilities.inchesToFeetInches(comp.height)
        sheet.Range("UTSIZE").OFFSET(i, 0).Value = comp.size
        If TypeOf comp Is wire Then
            If comp.wepHeight > 0 Then sheet.Range("UTHEIGHT").OFFSET(i, 0).Value = sheet.Range("UTHEIGHT").OFFSET(i, 0).Value & "(" & Utilities.inchesToFeetInches(comp.wepHeight) & ")"
            If comp.componentType = "PRI" And InStr(comp.size, "Ø") > 0 Then
                sheet.Range("DESC").OFFSET(descCount, 0).Value = Mid(comp.size, InStr(comp.size, "Ø") - 1, 2)
                sheet.Range("DESC").OFFSET(descCount, 0).HorizontalAlignment = xlLeft
                sheet.Range("SIZE").OFFSET(descCount, 0).Value = Left(comp.size, Len(comp.size) - 3)
                sheet.Range("SIZE").OFFSET(descCount, 0).HorizontalAlignment = xlLeft
                primaryDirections = ""
                For Each midspanSlot In comp.midspans
                    If primaryDirections <> "" Then primaryDirections = primaryDirections & "/"
                    primaryDirections = primaryDirections & Me.findSpan(midspanSlot).direction
                Next midspanSlot
                sheet.Range("DIRECTION").OFFSET(descCount, 0).Value = primaryDirections
                descCount = descCount + 1
            End If
            For j = 1 To maxSpans
                If comp.midspans.exists(j) Then
                    sheet.Range("UTMIDSPAN" & j).OFFSET(i, 0).Value = comp.getDisplayMidspan(Me.findSpan(j))
                    sheet.Range("UTMIDSPAN" & j).OFFSET(i, 0).HorizontalAlignment = xlLeft
                Else
                    sheet.Range("UTMIDSPAN" & j).OFFSET(i, 0).Value = "-"
                End If
            Next j
        Else
            For j = 1 To maxSpans
                sheet.Range("UTMIDSPAN" & j).OFFSET(i, 0).Value = "-"
            Next j
            sheet.Range("UTHEIGHT").OFFSET(i, 0).Value = comp.getDisplayHeight()
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
        sheet.Range("DET").OFFSET(0, 1).Value = "Hammer Test Failed"
        If Me.notes <> "" Then Me.notes = Me.notes & vbLf
        Me.notes = Me.notes & "Hammer test failed."
    End If
    If Me.visualCheckFailed Then
        If Me.notes <> "" Then Me.notes = Me.notes & vbLf
        Me.notes = Me.notes & "Fielder visual test failed."
        If sheet.Range("DET").OFFSET(0, 1).Value = "" Then sheet.Range("DET").OFFSET(0, 1).Value = "Fielder Visual Test Failed"
    End If
    sheet.Range("NOTES") = Me.notes
    
    'Violations
    Call calculateViolations
    sheet.Range("SEP").OFFSET(0, 1).Value = Me.violation40SZV
    sheet.Range("CLEARANCE").OFFSET(0, 1).Value = Me.violationUtil156MSGCV
    sheet.Range("CLEARANCE").OFFSET(1, 0).Value = Me.violationCom156MSGCV
    
    Dim separationViolations As Collection: Set separationViolations = New Collection
    separationViolations.Add Me.violation30MSSV
    separationViolations.Add Me.violation12CSV
    separationViolations.Add Me.violation6MSSV
    separationViolations.Add Me.violation40SLSV
    separationViolations.Add Me.violation12SLDLSV
    
    sheet.Range("SEP").OFFSET(1, 0).Value = Utilities.JoinCollection(separationViolations, vbLf)
    
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
    
    Dim project As project: Set project = New project
    
    If project.mode <> "SYSTEM IMPROVEMENT" Then
        Me.violation30MSSV = Calculate30MSSV()
        results = CalculateCommViolations()
        violation12CSV = results(0)
        violation6MSSV = results(1)
        results = Calculate40SLSV()
        violation40SLSV = results(0)
        violation12SLDLSV = results(1)
    End If
End Sub

Private Function Calculate40SZV() As String
    Dim result As String
    Dim violators As Collection: Set violators = New Collection
    Dim comp As Variant
    
    For i = 1 To Me.commComponents.count
        Set comp = Me.commComponents(i)
        If TypeOf comp Is wire Then
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
    Dim violators As Scripting.Dictionary: Set violators = New Scripting.Dictionary
    Dim wire As Variant
    Dim midspan As Integer
    Dim direction As String
    Dim coms As Boolean
    
    For i = 1 To components.count
        Set wire = components(i)
        If TypeOf wire Is wire Then
            For Each midspanSlot In wire.midspans
                If IsNumeric(wire.midspans(midspanSlot)) Then midspan = wire.midspans(midspanSlot)
                If wire.componentType = "PRI" Or wire.componentType = "NEUT" Or wire.componentType = "SEC" Or wire.componentType = "OW" Or wire.componentType = "SVC" Then
                    If midspan > 0 And midspan < 186 Then
                        direction = Me.findSpan(midspanSlot).direction
                        If Not violators.exists(direction) Then violators.Add direction, New Collection
                        violators(direction).Add wire.componentType
                    End If
                ElseIf wire.componentType <> "SPG" Then
                    If midspan > 0 And midspan < 186 Then
                        direction = Me.findSpan(midspanSlot).direction
                        If Not violators.exists(direction) Then violators.Add direction, New Collection
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
    Dim violators As Scripting.Dictionary: Set violators = New Scripting.Dictionary
    Dim wire As Variant
    Dim lowestMidspan As Integer
    Dim comMidspan As Integer
    Dim direction As String
    
    Dim Span As Span
    For Each Span In Me.spans
        If Not Span.lowestPowerWire Is Nothing Then
            lowestMidspan = Span.lowestPowerWire.midspans(Span.spanSlot)
            For i = 1 To Me.commComponents.count
                Set wire = Me.commComponents(i)
                If TypeOf wire Is wire Then
                    If wire.midspans.exists(Span.spanSlot) Then
                        comMidspan = wire.midspans(Span.spanSlot)
                        If lowestMidspan - 30 < comMidspan Then
                            direction = Span.direction
                            If Not violators.exists(direction) Then violators.Add direction, New Collection
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
    Dim separationViolators As Scripting.Dictionary: Set separationViolators = New Scripting.Dictionary
    Dim midspanViolators As Scripting.Dictionary: Set midspanViolators = New Scripting.Dictionary
    Dim wire1 As Variant
    Dim wire2 As Variant
    Dim Midspan1 As Integer
    Dim Midspan2 As Integer
    Dim direction As String
    
    For i = 1 To Me.commComponents.count
        For j = i + 1 To Me.commComponents.count
            Set wire1 = Me.commComponents(i)
            Set wire2 = Me.commComponents(j)
            If TypeOf wire1 Is wire And TypeOf wire2 Is wire Then
                If (Abs(wire1.height - wire2.height) < 12 And wire1.owner <> wire2.owner) Then
                    If Not separationViolators.exists(i) Then separationViolators.Add i, Nothing
                    If Not separationViolators.exists(j) Then separationViolators.Add j, Nothing
                ElseIf Abs(wire1.height - wire2.height) > 0 And Abs(wire1.height - wire2.height) < 4 And wire1.owner = wire2.owner And Utilities.ShareAKey(wire1.midspans, wire2.midspans) Then
                    If Not separationViolators.exists(i) Then separationViolators.Add i, Nothing
                    If Not separationViolators.exists(j) Then separationViolators.Add j, Nothing
                End If
                
                For Each midspanSlot In wire1.midspans
                    If wire2.midspans.exists(midspanSlot) Then
                        Midspan1 = wire1.midspans(midspanSlot)
                        Midspan2 = wire2.midspans(midspanSlot)
                        If Midspan1 <> 0 And Midspan2 <> 0 And Abs(Midspan1 - Midspan2) < 6 Then
                            direction = Me.findSpan(midspanSlot).direction
                            If Not midspanViolators.exists(direction) Then midspanViolators.Add direction, New Collection
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
    Dim sLBracket As Equipment
    Dim SLDL As Equipment
    Dim Equipment As Equipment
    Dim wire As Variant
    
    Dim sLBracketViolators As Collection: Set sLBracketViolators = New Collection
    Dim sLDLViolators As Collection: Set sLDLViolators = New Collection
    
    For Each Equipment In Me.equipments
        If Equipment.componentType = "SL" Then Set sLBracket = Equipment
        If Equipment.size = "SL" Then Set SLDL = Equipment
    Next Equipment
    
    If sLBracket Is Nothing And SLDL Is Nothing Then
        Calculate40SLSV = results
        Exit Function
    End If
    
    For i = 1 To Me.commComponents.count
        Set wire = Me.commComponents(i)
        
        If TypeOf wire Is wire Then
            If sLBracket.bottomHeight - 40 < wire.height Then sLBracketViolators.Add i
            If Not SLDL Is Nothing Then
                If Abs(SLDL.height - wire.height) < 12 Then sLDLViolators.Add i
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
        sheet.Range("COMM" & i).OFFSET(2, 0).Value = ""
        sheet.Range("COMM" & i).OFFSET(2, 0).OFFSET(0, 1).Value = ""
    Next i
    
    Dim ROW As Integer: ROW = sheet.Range("COMM1").ROW + 4
    For i = 0 To 50
        sheet.Rows(ROW).Delete
    Next i
End Sub

Public Sub fillCMRF(sheet As Worksheet)
    Dim wire As wire
    Dim comms As Scripting.Dictionary
    Set comms = New Scripting.Dictionary
    
    Set Me.commWires = Utilities.sortComponents(Me.commWires)
    
    For Each wire In Me.commWires
        If Not comms.exists(wire.owner) Then comms.Add wire.owner, New Collection
        comms(wire.owner).Add wire
    Next wire
    
    Dim commOwnerCount As Integer
    For Each commOwner In comms
        commOwnerCount = commOwnerCount + 1
        sheet.Range("COMM" & commOwnerCount).Value = commOwner
        sheet.Range("COMM" & commOwnerCount).HorizontalAlignment = xlLeft
        commCount = 0
        
        'Nondrops first
        For i = 1 To comms(commOwner).count
            Set wire = comms(commOwner)(i)
            If wire.componentType <> "DROP" Then
                commCount = commCount + 1
                If commCount > 1 Then
                    Set SourceRange = sheet.Range(sheet.Range("COMM" & commOwnerCount).OFFSET(2, -1), sheet.Range("COMM" & commOwnerCount).OFFSET(3, 1))
                    SourceRange.Copy
                    sheet.Range("COMM" & commOwnerCount).OFFSET(2, -1).OFFSET(((commCount - 1) * 2) - 1, 0).PasteSpecial Paste:=xlPasteAll
                End If
                sheet.Range("COMM" & commOwnerCount).OFFSET(3, 0).OFFSET(((commCount - 1) * 2) - 1, 0).Value = wire.getDisplayHeight(True)
            End If
        Next i
        
        'Drops after
        For i = 1 To comms(commOwner).count
            Set wire = comms(commOwner)(i)
            If wire.componentType = "DROP" Then
                commCount = commCount + 1
                If commCount > 1 Then
                    Set SourceRange = sheet.Range(sheet.Range("COMM" & commOwnerCount).OFFSET(2, -1), sheet.Range("COMM" & commOwnerCount).OFFSET(3, 1))
                    SourceRange.Copy
                    sheet.Range("COMM" & commOwnerCount).OFFSET(2, -1).OFFSET(((commCount - 1) * 2) - 1, 0).PasteSpecial Paste:=xlPasteAll
                End If
                sheet.Range("COMM" & commOwnerCount).OFFSET(3, 0).OFFSET(((commCount - 1) * 2) - 1, 0).Value = wire.getDisplayHeight(True)
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

Public Function findWireByTrace(trace_ As String, size_ As String) As wire
    Dim wire As wire
    Dim foundWire As wire: Set foundWire = Nothing
    For Each wire In wires
        If wire.trace = trace_ And wire.size = size_ Then
            Set foundWire = wire
            Exit For
        ElseIf wire.trace = trace_ And wire.size = "DROP" Then
            Set foundWire = wire
            Exit For
        ElseIf wire.trace = trace_ Then
            Set foundWire = wire
        End If
    Next wire
    
    Set findWireByTrace = foundWire
End Function

Public Sub condensePrimaries()
    Dim comp As Variant
    Dim primaryBundles As Scripting.Dictionary: Set primaryBundles = New Scripting.Dictionary
    Dim i As Integer
    Dim wire As wire
    Dim bundleKey As Variant
    Dim midspanSlot As Variant
    Dim primary As wire
    Dim phase As Integer
    Dim phaseSizes As Scripting.Dictionary
    
    For i = 1 To 12
        Set phaseSizes = New Scripting.Dictionary
        For Each comp In Me.utilWires
            If comp.componentType = "PRI" And InStr(comp.size, "Ø") = 0 Then
                If comp.midspans.exists(i) Then
                    If Not phaseSizes.exists(comp.size) Then
                        Set phaseSizes(comp.size) = New Scripting.Dictionary
                        phaseSizes(comp.size)("phase") = 0
                    End If
                    phaseSizes(comp.size)("phase") = phaseSizes(comp.size)("phase") + 1
                    Set phaseSizes(comp.size)("primary") = comp
                End If
            End If
        Next comp
        
        For Each size In phaseSizes
            Set primary = phaseSizes(size)("primary")
            phase = phaseSizes(size)("phase")
            If phase > 3 Then phase = 3
            bundleKey = primary.size & " " & phase & "Ø" & primary.height
            If Not primaryBundles.exists(bundleKey) Then primaryBundles.Add bundleKey, Array(primary, New Collection)
            primaryBundles(bundleKey)(1).Add i
        Next size
    Next i
    
    For Each bundleKey In primaryBundles
        Set wire = New wire
        Set comp = primaryBundles(bundleKey)(0)
        
        wire.size = Left(bundleKey, InStr(bundleKey, "Ø"))
        wire.componentType = comp.componentType
        wire.height = comp.height
        For Each midspanSlot In primaryBundles(bundleKey)(1)
            wire.midspans.Add midspanSlot, "PRI"
        Next midspanSlot
        
        Me.utilComponents.Add wire
    Next bundleKey
End Sub
