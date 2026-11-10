Attribute VB_Name = "SAPManager"
Public Sub exportSAPTxt()
    Call LogMessage.SendLogMessage("exportSAPTxt")
    
    Dim SAP As SAP: Set SAP = New SAP
    Call SAP.extractFromSheet
    
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    
    Dim SapTxt As String
    
    SapTxt = "Job Info:" & vbLf
    SapTxt = SapTxt & "APPLICANT: " & SAP.applicant & vbLf
    SapTxt = SapTxt & "#OF POLES: " & SAP.numPoles & vbLf
    SapTxt = SapTxt & "ADDRESS: " & SAP.address & vbLf
    SapTxt = SapTxt & "DUE DATE: " & SAP.dueDate & vbLf
    SapTxt = SapTxt & "NAS/PHONE#: " & SAP.coordinatorName & Chr(32) & SAP.coordinatorNumber
    If SAP.coordinatorEmails.count > 0 Then SapTxt = SapTxt & Chr(32) & SAP.coordinatorEmails(1)
    SapTxt = SapTxt & vbLf & vbLf
    
    SapTxt = SapTxt & "Designed by: " & SAP.designerName & IIf(SAP.designerNumber <> "", ", " & SAP.designerNumber, "") & vbLf
    SapTxt = SapTxt & "Notification: " & SAP.notification & vbLf
    SapTxt = SapTxt & "Construction Measure: " & SAP.constructionMeasure & vbLf
    SapTxt = SapTxt & "Permit Number: " & SAP.Permit & vbLf
    
    If SAP.mode = "MAKE READY" Then
        SapTxt = SapTxt & vbLf
        SapTxt = SapTxt & "Alt1:" & vbLf
        SapTxt = SapTxt & "Design Document: " & SAP.alt1DesDoc & vbLf
        SapTxt = SapTxt & "Order: " & SAP.alt1Order & vbLf
        SapTxt = SapTxt & "Alt2:" & vbLf
        SapTxt = SapTxt & "Design Document: " & SAP.alt2DesDoc & vbLf
        SapTxt = SapTxt & "Order: " & SAP.alt2Order & vbLf
        SapTxt = SapTxt & "Alt3:" & vbLf
        SapTxt = SapTxt & "Design Document: " & SAP.alt3DesDoc & vbLf
        SapTxt = SapTxt & "Order: " & SAP.alt3Order & vbLf & vbLf
    ElseIf SAP.mode = "SYSTEM IMPROVEMENT" Then
        SapTxt = SapTxt & "Total Planning Hours: " & SAP.totalHours & vbLf
        SapTxt = SapTxt & "Total Cost: " & SAP.totalCost & vbLf & vbLf
        
        SapTxt = SapTxt & "Design Document: " & SAP.mainDesDoc & vbLf
        SapTxt = SapTxt & "Order: " & SAP.mainOrder & vbLf
        SapTxt = SapTxt & "LLTM Notification: " & SAP.longLeadNotification & vbLf
        SapTxt = SapTxt & "LLTM Order: " & SAP.longLeadOrder & vbLf
        SapTxt = SapTxt & "Planning Hours: " & SAP.mainHours & vbLf
        SapTxt = SapTxt & "Cost: " & SAP.mainCost & vbLf & vbLf
        
        For i = 1 To SAP.serviceNotifications.count
            SapTxt = SapTxt & "SERVICE Notification: " & SAP.serviceNotifications(i) & vbLf
            SapTxt = SapTxt & "SERVICE Design Document: " & IIf(SAP.serviceDesDocs.count >= i, SAP.serviceDesDocs(i), "") & vbLf
            SapTxt = SapTxt & "Order: " & IIf(SAP.serviceOrders.count >= i, SAP.serviceOrders(i), "") & vbLf
            SapTxt = SapTxt & "LLTM Notification: " & IIf(SAP.serviceLongLeadNotifications.count >= i, SAP.serviceLongLeadNotifications(i), "") & vbLf
            SapTxt = SapTxt & "LLTM Order: " & IIf(SAP.serviceLongLeadOrders.count >= i, SAP.serviceLongLeadOrders(i), "") & vbLf
            SapTxt = SapTxt & "Planning Hours: " & IIf(SAP.serviceHours.count >= i, SAP.serviceHours(i), "") & vbLf
            SapTxt = SapTxt & "Cost: " & IIf(SAP.serviceCosts.count >= i, SAP.serviceCosts(i), "") & vbLf & vbLf
        Next i
        
        If SAP.demoNotification <> "" Then
            SapTxt = SapTxt & "DEMO Notification: " & SAP.demoNotification & vbLf
            SapTxt = SapTxt & "DEMO Design Document: " & SAP.demoDesDoc & vbLf
            SapTxt = SapTxt & "Construction Measure: " & SAP.demoConstructionMeasure & vbLf
            SapTxt = SapTxt & "Order: " & SAP.demoOrder & vbLf
            SapTxt = SapTxt & "Planning Hours: " & SAP.demoHours & vbLf
            SapTxt = SapTxt & "Cost: " & SAP.demoCost & vbLf & vbLf
        End If
    End If
    
    If InStr(LCase(SAP.treeWork), "n/a") > 0 Or InStr(LCase(SAP.treeWork), "na") > 0 Or InStr(LCase(SAP.treeWork), "no") > 0 Then SAP.treeWork = "No"
    If InStr(LCase(SAP.treeWork), "yes") = 0 And InStr(SAP.treeWork, "no") = 0 Then SAP.treeWork = "Yes" & vbLf & SAP.treeWork
    
    SapTxt = SapTxt & "Tree Work: " & SAP.treeWork & vbLf & vbLf
    
    SapTxt = SapTxt & "391 Forms: " & vbLf
    SapTxt = SapTxt & "# of Poles: " & SAP.num391Poles & vbLf & vbLf
    
    SapTxt = SapTxt & "CA/Violations NJUNS Project Number: " & SAP.njunsProject & vbLf
    SapTxt = SapTxt & "NJUNS Locations: " & SAP.njunsLocations & vbLf & vbLf
    
    SapTxt = SapTxt & "CE Maintenance:" & vbLf
    SapTxt = SapTxt & "# of Poles: " & SAP.numMaintPoles & vbLf & vbLf
    
    SapTxt = SapTxt & "# of Denied Poles: " & SAP.numDeniedPoles & vbLf & vbLf
    
    SapTxt = SapTxt & "647 Forms: " & SAP.forms647 & vbLf & vbLf
    
    SapTxt = SapTxt & "Comm Make Ready form included: Yes" & vbLf & vbLf
    
    SapTxt = SapTxt & "Design Locations requiring permit: " & SAP.permits & vbLf
    SapTxt = SapTxt & "Design Locations requiring MISS DIG forms: " & SAP.missdigLocations & vbLf
    SapTxt = SapTxt & "Design Locations requiring MDOT 5-day notify forms: " & SAP.mdot5Day & vbLf
    SapTxt = SapTxt & "Design Locations requiring MDOT tree-permit forms: " & SAP.mdotTree & vbLf & vbLf
    
    SapTxt = SapTxt & "Design Location Numbers Staked: " & SAP.stakedLocations & vbLf
    SapTxt = SapTxt & "Staked by: " & SAP.stakedBy & vbLf
    SapTxt = SapTxt & "Date Pole Staked: " & SAP.stakedDate & vbLf & vbLf
    
    SapTxt = SapTxt & "Railroad Permit Required: " & SAP.raildroadPermitRequired & vbLf
    SapTxt = SapTxt & "Railroad Permit Cost: " & SAP.permitCost & vbLf & vbLf
    
    SapTxt = SapTxt & "ERC CHECKLIST#: " & SAP.erc & vbLf
    SapTxt = SapTxt & "PLIA: " & SAP.plia & vbLf
    SapTxt = SapTxt & "TTC: " & IIf(InStr(SAP.ttc, vbLf) > 0, vbLf & "    ", "") & Replace(SAP.ttc, vbLf, vbLf & "    ") & vbLf
    SapTxt = SapTxt & "OUTAGE REQUIRED: " & SAP.outageRequired & vbLf
    
    If SAP.mode = "MAKE READY" Then
        SapTxt = SapTxt & "LONG LEAD NOTIF: " & SAP.longLeadNotification & vbLf
        SapTxt = SapTxt & "LONG LEAD ORDER: " & SAP.longLeadOrder & vbLf & vbLf
    End If
    
    SapTxt = SapTxt & "TRS: " & IIf(InStr(SAP.trs, vbLf) > 0, vbLf & "    ", "") & Replace(SAP.trs, vbLf, vbLf & "    ") & vbLf
    SapTxt = SapTxt & "TLM: " & SAP.tlm & vbLf
    If SAP.systemOwner <> "" Then SapTxt = SapTxt & "SYSTEM OWNER: " & SAP.systemOwner & vbLf
    SapTxt = SapTxt & "SUBSTATION: " & IIf(InStr(SAP.substation, vbLf) > 0, vbLf & "    ", "") & Replace(SAP.substation, vbLf, vbLf & "    ") & vbLf
    SapTxt = SapTxt & "CIRCUIT: " & IIf(InStr(SAP.circuit, vbLf) > 0, vbLf & "    ", "") & Replace(SAP.circuit, vbLf, vbLf & "    ") & vbLf
    SapTxt = SapTxt & "VOLTAGE: " & IIf(InStr(SAP.voltage, vbLf) > 0, vbLf & "    ", "") & Replace(SAP.voltage, vbLf, vbLf & "    ") & vbLf
    SapTxt = SapTxt & "TAXCODE: " & IIf(InStr(SAP.taxCode, vbLf) > 0, vbLf & "    ", "") & Replace(SAP.taxCode, vbLf, vbLf & "    ") & vbLf
    SapTxt = SapTxt & "HDQ: " & Replace(SAP.hdq, vbLf, vbLf & "    ") & vbLf
    SapTxt = SapTxt & "COUNTY: " & SAP.county & vbLf
    SapTxt = SapTxt & "TWP: " & SAP.twp & vbLf
    SapTxt = SapTxt & "Functional Loc: " & SAP.functionalLoc & vbLf & vbLf
    
    If SAP.upstreamDevices <> "" Then
        SapTxt = SapTxt & "UPSTREAM SECTIONALIZING DEVICE: " & vbLf
        If InStr(SAP.upstreamDevices, "LOC") > 0 Then
            SAP.upstreamDevices = Replace(SAP.upstreamDevices, vbLf, vbLf & "    ")
            SAP.upstreamDevices = Replace(SAP.upstreamDevices, "    LOC", "LOC")
        End If
        SapTxt = SapTxt & SAP.upstreamDevices & vbLf
    End If
    
    Dim fileSaveName As String
    jobFolderPath = Utilities.GetJobDirectory(SAP.notification)
    If jobFolderPath <> "" And Dir(jobFolderPath & "\Design Related Materials\") <> "" Then
        fileSaveName = jobFolderPath & "\Design Related Materials\" & SAP.notification & " - SAP Worksheet.txt"
    Else
        fileSaveName = GetLocalPath(ThisWorkbook.Path) & "\" & SAP.notification & " - SAP Worksheet.txt"
        If InStr(fileSaveName, "sharepoint") > 0 Or fileSaveName = "" Then fileSaveName = Environ("USERPROFILE") & "\Downloads\" & SAP.notification & " - SAP Worksheet.txt"
    End If
    
    Set oFile = fso.CreateTextFile(fileSaveName)
    oFile.WriteLine SapTxt
    oFile.Close
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    
    MsgBox "Finished Exporting SAP Worksheet"
End Sub

Sub fillSITreeWork(SAP As SAP, project As project)
    Dim pole As pole
    Dim treeWorkString As String
    For Each pole In project.poles
        If pole.location <> "" And pole.treeWork Then
            If treeWorkString <> "" Then treeWorkString = treeWorkString & vbLf
            treeWorkString = treeWorkString & "LOC " & pole.location
        End If
    Next pole
    If treeWorkString <> "" Then
        Dim dataWs As Worksheet
        Set dataWs = ThisWorkbook.sheets("DATA")
        Set hdqcell = dataWs.cells.Find(Utilities.OnlyLetters(SAP.hdq))
        forester = "Forester"
        If Not hdqcell Is Nothing Then
            foresterName = WorksheetFunction.Proper(hdqcell.offSet(0, 3).value)
            If foresterName <> "" Then forester = forester & " (" & foresterName & ")"
        End If
        treeWorkString = treeWorkString & vbLf & "Per: " & forester
    End If
    
    If treeWorkString = "" Then
        SAP.treeWork = "N/A"
    Else
        SAP.treeWork = treeWorkString
    End If
End Sub

Public Sub fillSAPInfoFromPDS()
    Call LogMessage.SendLogMessage("fillSAPInfoFromPDS")
    
    Dim project As project: Set project = New project
    Call project.extractFromSheets
    Dim SAP As SAP: Set SAP = New SAP
    Call SAP.extractFromSheet
    
    SAP.applicant = project.applicant
    SAP.numPoles = project.poles.count
    SAP.notification = project.notification
    SAP.Permit = project.Permit
    SAP.county = project.county & IIf(getCountyCode(project.county) <> "", " - ", "") & getCountyCode(project.county)
    SAP.twp = project.township
    
    If InStr(project.Permit, "-") > 0 Then
        digits = Left(project.Permit, InStr(project.Permit, "-") - 1)
        If IsNumeric(digits) Then
            Set dataWs = ThisWorkbook.sheets("Data")
            Set cell = dataWs.cells.Find("JUC-" & digits)
            If Not cell Is Nothing Then
                SAP.coordinatorName = cell.offSet(0, 1).value
                SAP.coordinatorNumber = cell.offSet(0, 2).value
                For i = 0 To 4
                    If cell.offSet(0, 3 + i).value = "" Then Exit For
                    SAP.coordinatorEmails.Add cell.offSet(0, 3 + i).value
                Next i
            End If
        End If
    End If
    
    Dim locationCount
    Dim maintCount As Integer
    Dim deniedCount As Integer
    Dim njunsCount As Integer
    Dim treeWork As Boolean
    Dim missdigLocations As Collection: Set missdigLocations = New Collection
    Dim stakedLocations As Collection: Set stakedLocations = New Collection
    Dim outageLocations As Collection: Set outageLocations = New Collection
    Dim ttcLocations As Scripting.Dictionary: Set ttcLocations = New Scripting.Dictionary
    
    Dim pole As pole
    For Each pole In project.poles
        If SAP.address = "" Then SAP.address = pole.address
        If pole.sheet.Range("MAINT").value <> "" Or pole.sheet.Range("SUMSHEET5").value Then mainCount = maintCount + 1
        If pole.denied Then deniedCount = deniedCount + 1
        If Replace(UCase(pole.njunsTicket), "/", "") <> "NA" And pole.njunsTicket <> "" Then njunsCount = njunsCount + 1
        If pole.location <> "" Then
            locationCount = locationCount + 1
            If pole.treeWork Then treeWork = True
            If pole.ReplaceAnchor Or pole.InstallAnchor Or pole.ReplacePole Or pole.InstallPole Or pole.ReplaceRiser Or pole.RemoveAnchor Or pole.RemovePole Then missdigLocations.Add pole.location
            If pole.sheet.Range("SUMSHEET13").value Then stakedLocations.Add pole.location
            If pole.outage Then outageLocations.Add pole.location
            If pole.ttc <> "" Then
                If Not ttcLocations.exists(pole.ttc) Then Set ttcLocations(pole.ttc) = New Collection
                ttcLocations(pole.ttc).Add pole.location
            End If
        End If
    Next pole
    
    SAP.numMaintPoles = maintCount
    SAP.numDeniedPoles = deniedCount
    
    SAP.njunsLocations = njunsCount
    If njunsCount = 0 Then SAP.njunsProject = "N/A"
    
    SAP.missdigLocations = compactListString(missdigLocations)
    SAP.outageRequired = IIf(outageLocations.count > 0, "OUTAGE LOC: " & compactListString(outageLocations), "NO OUTAGES")
    
    SAP.stakedLocations = compactListString(stakedLocations)
    If SAP.stakedLocations = "N/A" Then
        SAP.stakedBy = "N/A"
        SAP.stakedDate = "N/A"
    Else
        SAP.stakedBy = "TBD"
        SAP.stakedDate = "TBD"
    End If
    
    ttcDir = "C:\Users\aabraham\OneDrive - Consumers Energy\Desktop\TTC\"
    jobDir = Utilities.GetJobDirectory(SAP.notification) & "\Design Related Materials\"
    
    If Dir(ttcDir) <> "" And Dir(jobDir) <> "" Then
        fileName = Dir(jobDir & "*.*")
        Do While fileName <> ""
            If InStr(1, fileName, "TTC", vbBinaryCompare) > 0 Then
                fullPath = jobDir & fileName
                SetAttr fullPath, vbNormal
                Kill fullPath
                fileName = Dir(targetFolder & "*.*")
            Else
                fileName = Dir
            End If
        Loop
    End If
    
    Dim ttcString As String
    For Each ttc In ttcLocations
        If ttcLocations(ttc).count > 0 Then
            If Environ$("USERNAME") = "aabraham" Then
                If Dir(ttcDir) <> "" And Dir(jobDir) <> "" Then
                    sourceFile = ttcDir & "TTC " & ttc & ".pdf"
                    destFile = jobDir & SAP.notification & " - TTC " & ttc & ".pdf"
                    FileCopy sourceFile, destFile
                End If
            End If
            
            If ttcString <> "" Then ttcString = ttcString & vbLf
            ttcString = ttcString & "TTC " & ttc & ": " & compactListString(ttcLocations(ttc))
        End If
    Next ttc
    If ttcString = "" Then ttcString = "N/A"
    SAP.ttc = ttcString
    
    If Not treeWork Then
        SAP.treeWork = "N/A"
    ElseIf SAP.mode = "SYSTEM IMPROVEMENT" And SAP.treeWork = "" And SAP.hdq <> "" Then
        Call fillSITreeWork(SAP, project)
    End If
    
    Dim docWarehouseLink As String
    Dim docWarehouseLabelCell As Range
    Dim docWarehouseCell As Range
    Dim docWarehouseTarget As Range
    
    ' Get Document Warehouse URL
    docWarehouseLink = getDocWarehouseLink(SAP.notification)
    
    ' Find "Doc Warehouse Link:" on the current Fill Sheet
    Set docWarehouseLabelCell = ActiveSheet.cells.Find( _
        What:="Doc Warehouse Link:", _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        MatchCase:=False)
    
    If Not docWarehouseLabelCell Is Nothing Then
    
        ' Cell immediately to the right of the label
        Set docWarehouseCell = docWarehouseLabelCell.offSet(0, 1)
    
        ' If the destination is merged, use the entire merged range
        If docWarehouseCell.MergeCells Then
            Set docWarehouseTarget = docWarehouseCell.MergeArea
        Else
            Set docWarehouseTarget = docWarehouseCell
        End If
    
        ' Remove existing hyperlink
        If docWarehouseTarget.Hyperlinks.count > 0 Then
            docWarehouseTarget.Hyperlinks.Delete
        End If
    
        ' Clear existing content
        docWarehouseTarget.ClearContents
    
        ' Add Document Warehouse hyperlink
        If Len(docWarehouseLink) > 0 Then
            SAP.docWarehouseLink = docWarehouseLink
        Else
            SAP.docWarehouseLink = "Link Not Found"
        End If
    End If
    
    If locationCount = 0 Then
        SAP.constructionMeasure = "N/A"
        SAP.alt1DesDoc = "N/A"
        SAP.alt1Order = "N/A"
        SAP.alt2DesDoc = "N/A"
        SAP.alt2Order = "N/A"
        SAP.alt3DesDoc = "N/A"
        SAP.alt3Order = "N/A"
        SAP.longLeadNotification = "N/A"
        SAP.longLeadOrder = "N/A"
        SAP.num391Poles = 0
        SAP.forms647 = "N/A"
        SAP.permits = "N/A"
        SAP.permitType = "N/A"
        SAP.permitCost = "N/A"
        SAP.permitLocations = "N/A"
        SAP.mdot5Day = "N/A"
        SAP.mdotTree = "N/A"
        SAP.raildroadPermitRequired = "N/A"
        SAP.railroadPermitCost = "N/A"
        SAP.erc = "N/A"
        SAP.plia = "N/A"
        SAP.upstreamDevices = "N/A"
        SAP.jobPurpose = "N/A"
        SAP.hotLineTagging = "N/A"
    End If
    
    Call SAP.fillSheet
    
    MsgBox "Finished filling in SAP sheet from PDS"
End Sub

Public Sub generateTitleBlockJson()
    Call ImportBently

    Dim SAP As SAP: Set SAP = New SAP
    Call SAP.extractFromSheet
    Dim project As project: Set project = New project
    Call project.extractFromSheets
    
    Dim json As Scripting.Dictionary: Set json = New Scripting.Dictionary
    Dim locations As Collection: Set locations = New Collection
    Dim staking As Scripting.Dictionary: Set staking = New Scripting.Dictionary
    Dim pole As pole
    For Each pole In project.poles
        If pole.location <> "" Then
            locations.Add pole.location
            If pole.staking Then staking.Add pole.location, True
        End If
    Next pole
    
    json.Add "Mode", SAP.mode
    
    json.Add "DesignerName", UCase(SAP.designerName)
    json.Add "DesignerNumber", UCase(SAP.designerNumber)
    json.Add "CoordinaterName", UCase(SAP.coordinatorName)
    json.Add "CoordinaterNumber", UCase(SAP.coordinatorNumber)
    json.Add "County", Utilities.OnlyLetters(UCase(SAP.county))
    json.Add "TWP", UCase(SAP.twp)
    json.Add "JobPurpose", UCase(SAP.jobPurpose)
    json.Add "Substation", splitValuesByLoc(SAP.substation, locations)
    json.Add "Circuit", splitValuesByLoc(SAP.circuit, locations)
    json.Add "UpstreamDevice", splitValuesByLoc(SAP.upstreamDevices, locations)
    json.Add "Staking", staking
    
    If SAP.mode = "SYSTEM IMPROVEMENT" Then
        json.Add "MainNotification", SAP.notification
        json.Add "MainDesDoc", SAP.mainDesDoc
        json.Add "DemoNotification", SAP.demoNotification
        json.Add "DemoDesDoc", SAP.demoDesDoc
        json.Add "ServiceNotifications", SAP.serviceNotifications
        json.Add "ServiceDesDocs", SAP.serviceDesDocs
    End If
    
    jsonText = JsonConverter.ConvertToJson(json, Whitespace:=2)
    filePath = "C:\ProgramData\Bentley\OpenUtilities Map Connect Edition\Configuration\WorkSpaces\ConsumersEnergy\Standards\vba\title_block.json"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.CreateTextFile(filePath, True, False)
    
    file.Write jsonText
    file.Close
    
    MsgBox "Press '8' to autofill all titleblocks. If autofill goes wrong, you can press Ctrl+Z to undo the autofill."
End Sub

Private Function splitValuesByLoc(value As String, ByVal locations As Collection) As Scripting.Dictionary
    Dim splitValue As String
    Dim results As Scripting.Dictionary: Set results = New Scripting.Dictionary
    
    lines = Split(value, vbLf)
    For Each line In lines
        If InStr(line, "LOC ") > 0 Then
            If splitValue <> "" Then
                For Each location In locations
                    results.Add location, splitValue
                Next location
            End If
            splitValue = ""
            Set locations = GetNumbers(Replace(line, "LOC ", ""))
        Else
            If splitValue <> "" Then splitValue = splitValue & vbLf
            splitValue = splitValue & UCase(line)
        End If
    Next line
    
    If splitValue <> "" Then
        For Each location In locations
            results.Add location, splitValue
        Next location
    End If
    
    Set splitValuesByLoc = results
End Function

Private Function GetNumbers(ByVal s As String) As Collection
    Dim parts() As String
    Dim result As Collection: Set result = New Collection
    Dim p As Variant
    Dim bounds() As String
    Dim i As Long, n As Long
    
    s = Replace(s, " and ", ",")
    parts = Split(s, ",")
    
    For Each p In parts
        p = Trim(p)
        
        If InStr(p, "-") > 0 Then
            bounds = Split(p, "-")
            
            For i = CLng(bounds(0)) To CLng(bounds(1))
                result.Add i
                n = n + 1
            Next i
        Else
            result.Add CLng(p)
            n = n + 1
        End If
    Next p
    
   Set GetNumbers = result
End Function

Public Sub fillSAPNoPermits()
    Dim SAP As SAP: Set SAP = New SAP
    
    If SAP.permits <> "" Then SAP.permits = "N/A"
    If SAP.permitType <> "" Then SAP.permitType = "N/A"
    If SAP.permitCost <> "" Then SAP.permitCost = "N/A"
    If SAP.permitLocations <> "" Then SAP.permitLocations = "N/A"
    
    If SAP.mdot5Day <> "" Then SAP.mdot5Day = "N/A"
    If SAP.mdotTree <> "" Then SAP.mdotTree = "N/A"
    
    If SAP.raildroadPermitRequired <> "" Then SAP.raildroadPermitRequired = "N/A"
    If SAP.railroadPermitCost <> "" Then SAP.railroadPermitCost = "N/A"
    
    Call SAP.fillSheet
End Sub

Public Sub fillSAPNoServiceNotif()
    Dim SAP As SAP: Set SAP = New SAP
    
    SAP.serviceNotifications.Add "N/A"
    SAP.serviceHouseNumbers.Add "N/A"
    SAP.serviceDesDocs.Add "N/A"
    SAP.serviceOrders.Add "N/A"
    SAP.serviceLongLeadNotifications.Add "N/A"
    SAP.serviceLongLeadOrders.Add "N/A"
    SAP.serviceCosts.Add "N/A"
    SAP.serviceHours.Add "N/A"
    
    Call SAP.fillSheet
End Sub

Public Sub fillSAPNoDemoNotif()
    Dim SAP As SAP: Set SAP = New SAP
    
    SAP.demoNotification = "N/A"
    SAP.demoConstructionMeasure = "N/A"
    SAP.demoDesDoc = "N/A"
    SAP.demoOrder = "N/A"
    SAP.demoLongLeadOrder = "N/A"
    SAP.demoLongLeadNotification = "N/A"
    SAP.demoCost = "N/A"
    SAP.demoHours = "N/A"
    
    Call SAP.fillSheet
End Sub

Public Sub fillSAPInfoFromGIS()
    Dim project As project: Set project = New project
    Call project.extractFromSheets
    Dim SAP As SAP: Set SAP = New SAP
    Call SAP.extractFromSheet
    
    Dim token As String: token = GetToken
    If Not testToken(token) Then
        MsgBox "Invalid token, get an up to date one from GIS."
        Exit Sub
    End If
    
    Dim substationLocations As Scripting.Dictionary: Set substationLocations = New Scripting.Dictionary
    Dim circuitLocations As Scripting.Dictionary: Set circuitLocations = New Scripting.Dictionary
    Dim hdqLocations As Scripting.Dictionary: Set hdqLocations = New Scripting.Dictionary
    Dim systemOwnerLocations As Scripting.Dictionary: Set systemOwnerLocations = New Scripting.Dictionary
    Dim voltageLocations As Scripting.Dictionary: Set voltageLocations = New Scripting.Dictionary
    Dim upstreamDeviceLocastions As Scripting.Dictionary: Set upstreamDeviceLocastions = New Scripting.Dictionary
    Dim trsLocations As Scripting.Dictionary: Set trsLocations = New Scripting.Dictionary
    Dim taxLocations As Scripting.Dictionary: Set taxLocations = New Scripting.Dictionary
    Dim countyLocations As Scripting.Dictionary: Set countyLocations = New Scripting.Dictionary
    Dim townLocations As Scripting.Dictionary: Set townLocations = New Scripting.Dictionary
    Dim hotLineTaggingLocations As Collection: Set hotLineTaggingLocations = New Collection
    Dim tlm As String
    Dim transformerJson As Object
    Dim circuitJson As Object
    
    Dim pole As pole
    For Each pole In project.poles
        Dim hotLineTagging As Boolean: hotLineTagging = False
        Dim poleCol As Collection: Set poleCol = New Collection
        poleCol.Add pole
        If pole.location <> "" Then
            Set circuitJson = getElectricJson(poleCol, 123, token, , 1)
            For Each jsonFeature In circuitJson("features")
                Dim substation As String, circuit As String, hdq As String, systemOwner As String, voltage As String, feederId As String, upstreamDevice As String, deviceLocation As String
                
                If Not IsNull(jsonFeature("attributes")("SUBSTATION")) Then substation = jsonFeature("attributes")("SUBSTATION") & " - " & jsonFeature("attributes")("SUBSTATIONID")
                If Not IsNull(jsonFeature("attributes")("CIRCUIT")) Then circuit = jsonFeature("attributes")("CIRCUIT") & " - " & jsonFeature("attributes")("CIRCUITID")
                If Not IsNull(jsonFeature("attributes")("HDQ")) Then hdq = jsonFeature("attributes")("HDQ") & " - " & getHDQCode(jsonFeature("attributes")("HDQ"))
                If Not IsNull(jsonFeature("attributes")("SYSTEM_OWNER")) Then systemOwner = jsonFeature("attributes")("SYSTEM_OWNER") & ", " & jsonFeature("attributes")("OFFICE_PHONE")
                If Not IsNull(jsonFeature("attributes")("VOLTAGE")) Then voltage = jsonFeature("attributes")("VOLTAGE")
                If Not IsNull(jsonFeature("attributes")("FEEDERID")) Then feederId = jsonFeature("attributes")("FEEDERID")
                
                If substation <> "" And Not substationLocations.exists(substation) Then Set substationLocations(substation) = New Collection
                If circuit <> "" And Not circuitLocations.exists(circuit) Then Set circuitLocations(circuit) = New Collection
                If hdq <> "" And Not hdqLocations.exists(hdq) Then Set hdqLocations(hdq) = New Collection
                If systemOwner <> "" And Not systemOwnerLocations.exists(systemOwner) Then Set systemOwnerLocations(systemOwner) = New Collection
                If voltage <> "" And Not voltageLocations.exists(voltage) Then Set voltageLocations(voltage) = New Collection
                If feederId <> "" Then
                    Set ud = GetUpstreamDeviceLatLong(feederId, pole.latitude, pole.longitude)
                    If Not ud Is Nothing Then
                        Dim lcp As String: lcp = Replace(ud(1)("LCP"), " ", "")
                        If lcp = "SUB" Then hotLineTagging = True
                        If (pole.ReplacePole Or pole.InstallPole) And pole.phase > 1 And Not hotLineTagging Then
                            Set ud2 = GetUpstreamDeviceLcp(feederId, lcp)
                            If Not ud2 Is Nothing Then
                                Dim lcp2 As String: lcp2 = Replace(ud(1)("LCP"), " ", "")
                                If lcp2 = "SUB" Then hotLineTagging = True
                            End If
                        End If
                        
                        Dim rating As String: rating = Replace(ud(1)("rating"), " ", "")
                        deviceLocation = getLCPLocation(pole, lcp, feederId, token)
                        upstreamDevice = vbLf & lcp & "-" & rating & IIf(deviceLocation <> "", vbLf, "") & deviceLocation
                        Do While InStr(upstreamDevice, vbLf & " ") > 0
                            upstreamDevice = Replace(upstreamDevice, vbLf & " ", vbLf)
                        Loop
                        Do While InStr(upstreamDevice, "  ") > 0
                            upstreamDevice = Replace(upstreamDevice, "  ", " ")
                        Loop
                    End If
                End If
                
                
                If upstreamDevice <> "" And Not upstreamDeviceLocastions.exists(upstreamDevice) Then Set upstreamDeviceLocastions(upstreamDevice) = New Collection
                
                If substation <> "" Then substationLocations(substation).Add pole.location
                If circuit <> "" Then circuitLocations(circuit).Add pole.location
                If hdq <> "" Then hdqLocations(hdq).Add pole.location
                If systemOwner <> "" Then systemOwnerLocations(systemOwner).Add pole.location
                If voltage <> "" Then voltageLocations(voltage).Add pole.location
                If upstreamDevice <> "" Then upstreamDeviceLocastions(upstreamDevice).Add pole.location
            Next jsonFeature
            
            Dim sectionJson As Object: Set sectionJson = getROWJSON(poleCol, 30, token, 1)
            If Not sectionJson("features") Is Nothing Then
                For Each jsonFeature In sectionJson("features")
                    Dim trs As String
                
                    If Not IsNull(jsonFeature("attributes")("SECTIONNAME")) Then trs = jsonFeature("attributes")("SECTIONNAME")
                    If trs <> "" And Not trsLocations.exists(trs) Then Set trsLocations(trs) = New Collection
                    If trs <> "" Then trsLocations(trs).Add pole.location
                Next jsonFeature
            End If
            
            Dim taxJson As Object: Set taxJson = getROWJSON(poleCol, 35, token, 1)
            If Not taxJson("features") Is Nothing Then
                For Each jsonFeature In taxJson("features")
                    Dim taxCode As String
                
                    If Not IsNull(jsonFeature("attributes")("TAXID")) Then taxCode = jsonFeature("attributes")("TAXID")
                    If taxCode <> "" And Not taxLocations.exists(taxCode) Then Set taxLocations(taxCode) = New Collection
                    If taxCode <> "" Then taxLocations(taxCode).Add pole.location
                Next jsonFeature
            End If
            
            Dim countyJson As Object: Set countyJson = getROWJSON(poleCol, 37, token, 1)
            If Not countyJson("features") Is Nothing Then
                For Each jsonFeature In countyJson("features")
                    Dim countyName As String
                    Dim countyNumber As String
                    Dim countyCombined As String
                
                    If Not IsNull(jsonFeature("attributes")("BOUNDARYNAMECD")) Then countyName = jsonFeature("attributes")("BOUNDARYNAMECD")
                    If Not IsNull(jsonFeature("attributes")("CECOUNTYNUMBER")) Then countyNumber = jsonFeature("attributes")("CECOUNTYNUMBER")
                    If countyName <> "" And countyNumber <> "" Then countyCombined = countyName & " - " & countyNumber
                    
                    If countyCombined <> "" And Not countyLocations.exists(countyCombined) Then Set countyLocations(countyCombined) = New Collection
                    If countyCombined <> "" Then countyLocations(countyCombined).Add pole.location
                Next jsonFeature
            End If
            
            Dim townJson As Object: Set townJson = getElectricJson(poleCol, 122, token, , 1)
            If Not townJson("features") Is Nothing Then
                For Each jsonFeature In townJson("features")
                    Dim townName As String
                    
                
                    If Not IsNull(jsonFeature("attributes")("BOUNDARYTYPE")) Then
                        If jsonFeature("attributes")("BOUNDARYTYPE") = 3 Or jsonFeature("attributes")("BOUNDARYTYPE") = 4 Then
                            If Not IsNull(jsonFeature("attributes")("BOUNDARYNAMECD")) Then townName = jsonFeature("attributes")("BOUNDARYNAMECD")
                        End If
                    End If
                    
                    If townName <> "" And Not townLocations.exists(townName) Then Set townLocations(townName) = New Collection
                    If townName <> "" Then townLocations(townName).Add pole.location
                Next jsonFeature
            End If

        End If
        If tlm = "" And pole.transformerSizes <> 0 Then
            Set transformerJson = getElectricJson(poleCol, 27, token, , 100)
            
            Test = getClosestPole(transformerJson, pole.x, pole.y)
            If Not Test(0) Is Nothing Then tlm = Test(0)("attributes")("TLM")
        End If
        
        If hotLineTagging Then
            hotLineTaggingLocations.Add pole.location
        End If
    Next pole
    
    On Error Resume Next
    If tlm = "" Then
        Dim poleCollections As Collection: Set poleCollections = findPoleGroups(project.poles)
        Dim poleCollection As Collection
        For Each poleCollection In poleCollections

            Dim secondaryJson As Object: Set secondaryJson = getElectricJson(poleCollection, 32, token)
            Dim secondaryUGJson As Object: Set secondaryUGJson = getElectricJson(poleCollection, 33, token)
            For Each jsonFeature In secondaryUGJson("features")
                secondaryJson("features").Add jsonFeature
            Next jsonFeature
            Set transformerJson = getElectricJson(poleCollection, 27, token)
            Dim transformerConnectorJson As Object: Set transformerConnectorJson = getElectricJson(poleCollection, 30, token)
            Dim tapPointsJson As Object: Set tapPointsJson = getElectricJson(poleCollection, 13, token)
            Dim polesJson As Object: Set polesJson = getElectricJson(poleCollection, 3, token)
            Dim serviceJson As Object: Set serviceJson = getElectricJson(poleCollection, 26, token)
            Call serviceJsons.Add(serviceJson)
            Set circuitJson = getElectricJson(poleCollection, 123, token)
            For Each pole In poleCollection
                If tlm <> "" Then Exit For
                
                If pole.secondaries.count > 0 Or pole.openWires.count > 0 Or pole.transformerSizes > 0 Then
                    results = getClosestPole(polesJson, pole.x, pole.y)
                    Dim closestPole As Object: Set closestPole = results(0)
                    Dim sourceX As Double: sourceX = 0
                    Dim sourceY As Double: sourceY = 0
                    Dim x As Double, y As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double
                    x = 0: y = 0: x1 = 0: y1 = 0: x2 = 0: y2 = 0
                    If Not closestPole Is Nothing Then
                        Dim closestPoleX As Double: closestPoleX = closestPole("geometry")("x")
                        Dim closestPoleY As Double: closestPoleY = closestPole("geometry")("y")
                    
                        For Each jsonFeature In secondaryJson("features")
                            For Each Path In jsonFeature("geometry")("paths")
                                For Each Point In Path
                                    If AreTwoPointsEqual(CDbl(Point(1)), CDbl(Point(2)), closestPoleX, closestPoleY) Then
                                        sourceX = closestPole("geometry")("x")
                                        sourceY = closestPole("geometry")("y")
                                        tlm = FindTLMFromSec(jsonFeature, secondaryJson, transformerJson, True)
                                        If tlm <> "" Then Exit For
                                    End If
                                Next Point
                                If tlm <> "" Then Exit For
                            Next Path
                            If tlm <> "" Then Exit For
                        Next jsonFeature
                        
                        If tlm <> "" Then Exit For
                        
                        Dim attributes As Scripting.Dictionary: Set attributes = New Scripting.Dictionary
                        Set attributes("SUBTYPECD") = New Collection
                        attributes("SUBTYPECD").Add 8
                        attributes("SUBTYPECD").Add 7
                        results = getClosestPole(tapPointsJson, closestPoleX, closestPoleY, attributes, True)
                        If results(1) > 0 And results(1) < secondaryOffsetMaxDistance Then
                            Set closestSource = results(0)
                            sourceX = closestSource("geometry")("x")
                            sourceY = closestSource("geometry")("y")
                            For Each jsonFeature In secondaryJson("features")
                                For Each Path In jsonFeature("geometry")("paths")
                                    For Each Point In Path
                                        If AreTwoPointsEqual(CDbl(Point(1)), CDbl(Point(2)), sourceX, sourceY) Then
                                            tlm = FindTLMFromSec(jsonFeature, secondaryJson, transformerJson, True)
                                            If tlm <> "" Then Exit For
                                        End If
                                    Next Point
                                    If tlm <> "" Then Exit For
                                Next Path
                                If tlm <> "" Then Exit For
                            Next jsonFeature
                        End If
                        
                        If tlm <> "" Then Exit For
                        
                        If sourceX = 0 And sourceY = 0 Then
                            For Each jsonFeature In secondaryJson("features")
                                For Each Path In jsonFeature("geometry")("paths")
                                    For Each Point In Path
                                        x = Point(1)
                                        y = Point(2)
                                        distance = Sqr((closestPoleX - x) ^ 2 + (closestPoleY - y) ^ 2)
                                        If distance < secondaryOffsetMaxDistance Then
                                            tlm = FindTLMFromSec(jsonFeature, secondaryJson, transformerJson, True)
                                            If tlm <> "" Then Exit For
                                        End If
                                    Next Point
                                    If tlm <> "" Then Exit For
                                Next Path
                                If tlm <> "" Then Exit For
                            Next jsonFeature
                        End If
                        
                        If tlm <> "" Then Exit For
                        
                        If sourceX = 0 And sourceY = 0 Then
                            For Each jsonFeature In secondaryJson("features")
                                For Each Path In jsonFeature("geometry")("paths")
                                    For i = 1 To Path.count - 1
                                        Set point1 = Path(i)
                                        Set point2 = Path(i + 1)
                                        result = GetClosestPointOnLine(CDbl(point1(1)), CDbl(point1(2)), CDbl(point2(1)), CDbl(point2(2)), closestPoleX, closestPoleY)
                                        x = result(0)
                                        y = result(1)
                                        distance = Sqr((closestPoleX - x) ^ 2 + (closestPoleY - y) ^ 2)
                                        If distance < secondaryOffsetMaxDistance Then
                                            tlm = FindTLMFromSec(jsonFeature, secondaryJson, transformerJson, True)
                                            If tlm <> "" Then Exit For
                                        End If
                                    Next i
                                    If tlm <> "" Then Exit For
                                Next Path
                                If tlm <> "" Then Exit For
                            Next jsonFeature
                        End If
                        
                        If tlm <> "" Then Exit For
                        
                        If sourceX <> 0 And sourceY <> 0 Then
                            For Each jsonFeature In tapPointsJson("features")
                                If jsonFeature("attributes")("SUBTYPECD") = 9 Then
                                    x = jsonFeature("geometry")("x")
                                    y = jsonFeature("geometry")("y")
                                    distance = Sqr((sourceX - x) ^ 2 + (sourceY - y) ^ 2)
                                    If distance < secondaryOffsetMaxDistance Then
                                        For Each jsonSecondaryFeature In secondaryJson("features")
                                            For Each Path In jsonSecondaryFeature("geometry")("paths")
                                                For Each Point In Path
                                                    If Point(1) = x And Point(2) = y Then
                                                        tlm = FindTLMFromSec(jsonSecondaryFeature, secondaryJson, transformerJson, True)
                                                        If tlm <> "" Then Exit For
                                                    End If
                                                Next Point
                                                If tlm <> "" Then Exit For
                                            Next Path
                                            If tlm <> "" Then Exit For
                                        Next jsonSecondaryFeature
                                    End If
                                End If
                                If tlm <> "" Then Exit For
                            Next jsonFeature
                        End If
                        
                        If tlm <> "" Then Exit For
                        
                        For Each jsonFeature In tapPointsJson("features")
                            If jsonFeature("attributes")("SUBTYPECD") = 9 Then
                                x = jsonFeature("geometry")("x")
                                y = jsonFeature("geometry")("y")
                                distance = Sqr((closestPoleX - x) ^ 2 + (closestPoleY - y) ^ 2)
                                If distance < secondaryOffsetMaxDistance Then
                                    For Each jsonSecondaryFeature In secondaryJson("features")
                                        For Each Path In jsonSecondaryFeature("geometry")("paths")
                                            For Each Point In Path
                                                If Point(1) = x And Point(2) = y Then
                                                    tlm = FindTLMFromSec(jsonSecondaryFeature, secondaryJson, transformerJson, True)
                                                    If tlm <> "" Then Exit For
                                                End If
                                            Next Point
                                            If tlm <> "" Then Exit For
                                        Next Path
                                        If tlm <> "" Then Exit For
                                    Next jsonSecondaryFeature
                                End If
                            End If
                            If tlm <> "" Then Exit For
                        Next jsonFeature
                    End If
                End If
                If tlm <> "" Then Exit For
            Next pole
        Next poleCollection
    End If
    On Error GoTo 0
    
    SAP.substation = compactLocations(substationLocations)
    SAP.circuit = compactLocations(circuitLocations)
    SAP.hdq = compactLocations(hdqLocations)
    SAP.systemOwner = compactLocations(systemOwnerLocations)
    SAP.voltage = compactLocations(voltageLocations)
    SAP.upstreamDevices = compactLocations(upstreamDeviceLocastions)
    SAP.trs = compactLocations(trsLocations)
    SAP.taxCode = compactLocations(taxLocations)
    If townLocations.count > 0 Then
        Trim(UCase(SAP.twp)) = Trim(UCase(townLocations.keys(0)))
        If project.township <> SAP.twp Then
            Dim sheet As Worksheet
            For Each sheet In ThisWorkbook.sheets
                If Utilities.IsPDS(sheet) Then
                    If Utilities.RangeExists(sheet, "TWP") Then sheet.Range("TWP") = SAP.twp
                End If
            Next sheet
        End If
    End If
    If countyLocations.count > 0 Then SAP.county = compactLocations(countyLocations)
    
    SAP.tlm = tlm
    If SAP.upstreamDevices <> "" Then SAP.hotLineTagging = compactListString(hotLineTaggingLocations)
    
    If InStr(SAP.county, vbLf) > 0 Then
        county = Utilities.OnlyLetters(Replace(Split(SAP.county, vbLf)(0), "LOC", ""))
    Else
        county = Utilities.OnlyLetters(SAP.county)
    End If
    
    
    If getCountyCode(county) <> "" And Len(tlm) >= 8 Then
        SAP.functionalLoc = "CA-" & getCountyCode(county) & "-" & Left(tlm, 8) & "-01"
    ElseIf getCountyCode(county) <> "" And SAP.trs <> "" Then
        SAP.functionalLoc = "CA-" & getCountyCode(county) & "-" & SAP.trs & "00-01"
    End If
    
    If SAP.mode = "SYSTEM IMPROVEMENT" And SAP.treeWork = "" And SAP.hdq <> "" Then
        Call fillSITreeWork(SAP, project)
    End If
    
    Call SAP.fillSheet
    
    MsgBox "Finished filling in SAP sheet from GIS"
End Sub

Function GetUpstreamDeviceLatLong( _
    ByVal feederId As String, _
    ByVal latitude As String, _
    ByVal longitude As String) As Object

    On Error Resume Next

    Dim http As Object
    Dim url As String
    Dim body As String

    url = "http://eamdashboard/ProjectWorkcenter/GetUpstreamDevice"

    body = "{""InputType"":""coordinates""," & _
           """Input1"":""" & latitude & """," & _
           """Input2"":""" & longitude & """," & _
           """FeederId"":""" & feederId & """}"

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open "POST", url, False
    http.SetAutoLogonPolicy 0

    http.SetRequestHeader "Content-Type", "application/json"
    http.SetRequestHeader "Accept", "application/json"

    http.Send body
    
    Dim upstreamDevice As Object
    Set upstreamDevice = JsonConverter.ParseJson(http.ResponseText)

    On Error GoTo 0

    Set GetUpstreamDeviceLatLong = upstreamDevice
End Function

Function GetUpstreamDeviceLcp( _
    ByVal feederId As String, _
    ByVal lcp As String) As Object

    On Error Resume Next

    Dim http As Object
    Dim url As String
    Dim body As String

    url = "http://eamdashboard/ProjectWorkcenter/GetUpstreamDevice"

    body = "{""InputType"":""lcp & feeder""," & _
           """Input1"":""" & latitude & """," & _
           """FeederId"":""" & feederId & """}"

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open "POST", url, False
    http.SetAutoLogonPolicy 0

    http.SetRequestHeader "Content-Type", "application/json"
    http.SetRequestHeader "Accept", "application/json"

    http.Send body
    
    Dim upstreamDevice As Object
    Set upstreamDevice = JsonConverter.ParseJson(http.ResponseText)

    On Error GoTo 0

    Set GetUpstreamDeviceLcp = upstreamDevice
End Function

Function getLCPLocation(pole As pole, lcp As String, feederId As String, token As String) As String
    Dim locationDescription As String
    Dim poleCollection As Collection: Set poleCollection = New Collection
    poleCollection.Add pole

    Dim jsonFuses As Object: Set jsonFuses = getElectricJson(poleCollection, 60, token, radius_:=12000)
    Dim jsonReclosers As Object: Set jsonReclosers = getElectricJson(poleCollection, 70, token, radius_:=12000)
    Dim jsonSectionalizers As Object: Set jsonSectionalizers = getElectricJson(poleCollection, 71, token, radius_:=12000)
    Dim jsonSwitches As Object: Set jsonSwitches = getElectricJson(poleCollection, 65, token, radius_:=12000)
    
    If locationDescription = "" Then locationDescription = parseFeaturesForLCP(jsonFuses, lcp, feederId)
    If locationDescription = "" Then locationDescription = parseFeaturesForLCP(jsonReclosers, lcp, feederId)
    If locationDescription = "" Then locationDescription = parseFeaturesForLCP(jsonSectionalizers, lcp, feederId)
    If locationDescription = "" Then locationDescription = parseFeaturesForLCP(jsonSwitches, lcp, feederId)
    
    If locationDescription <> "" Then
        locationDescription = SplitInHalf(locationDescription)
    End If
    
    getLCPLocation = locationDescription
End Function

Function SplitInHalf(ByVal txt As String) As String

    Dim midPos As Long
    Dim leftSpace As Long
    Dim rightSpace As Long
    Dim splitPos As Long

    midPos = Len(txt) \ 2

    leftSpace = InStrRev(txt, " ", midPos)
    rightSpace = InStr(midPos + 1, txt, " ")

    If leftSpace = 0 Then
        splitPos = rightSpace
    ElseIf rightSpace = 0 Then
        splitPos = leftSpace
    ElseIf midPos - leftSpace <= rightSpace - midPos Then
        splitPos = leftSpace
    Else
        splitPos = rightSpace
    End If

    If splitPos > 0 Then
        SplitInHalf = Left$(txt, splitPos - 1) & vbCrLf & _
                      Mid$(txt, splitPos + 1)
    Else
        SplitInHalf = txt
    End If

End Function

Function parseFeaturesForLCP(jsonObjects As Object, lcp As String, feederId As String) As String
    Dim locationDescription As String
    For Each jsonFeature In jsonObjects("features")
        If lcp = jsonFeature("attributes")("LCP") And feederId = jsonFeature("attributes")("FEEDERID") Then
            locationDescription = jsonFeature("attributes")("LOCATIONDESCRIPTION")
            Exit For
        End If
    Next jsonFeature
    
    parseFeaturesForLCP = locationDescription
End Function

Public Sub clearSAPSheet()
    vbanswer = MsgBox("Are you sure you want to clear the sap worksheet?", vbYesNoCancel)
    
    If vbanswer <> vbYes Then Exit Sub
    
    Dim SAP As SAP: Set SAP = New SAP
    Call SAP.clearSheets
End Sub

Public Sub ExportDesDoc()
    Call SendLogMessage("ExportDesDoc")

    Dim sap_ws As Worksheet
    Dim cell As Range
    Dim notification As String
    Dim hdq As String
    Dim constructionMeasure As String
    Dim functionalLocation As String
    Dim Permit As String
    
    Dim planningPlant As Scripting.Dictionary
    Set planningPlant = New Scripting.Dictionary
    planningPlant.Add "0011", "Alma - 0011"
    planningPlant.Add "0012", "Clare - 0012"
    planningPlant.Add "0013", "Greenville - 0013"
    planningPlant.Add "0015", "Gladwin - 0015"
    planningPlant.Add "0017", "Big Rapids - 0017"
    planningPlant.Add "0021", "Battle Creek - 0021"
    planningPlant.Add "0022", "Bronson - 0022"
    planningPlant.Add "0023", "Coldwater - 0023"
    planningPlant.Add "0024", "Marshall - 0024"
    planningPlant.Add "0025", "Hastings - 0025"
    planningPlant.Add "0031", "Bay City - 0031"
    planningPlant.Add "0032", "Midland - 0032"
    planningPlant.Add "0033", "Rose City - 0033"
    planningPlant.Add "0034", "Tawas - 0034"
    planningPlant.Add "0035", "Standish - 0035"
    planningPlant.Add "0038", "West Branch - 0038"
    planningPlant.Add "0044", "Groveland - 0044"
    planningPlant.Add "0051", "Flint - 0051"
    planningPlant.Add "0052", "Owosso - 0052"
    planningPlant.Add "0061", "Grand Rapids - 0061"
    planningPlant.Add "0062", "Zeeland - 0062"
    planningPlant.Add "0063", "East Kent - 0063"
    planningPlant.Add "0064", "North Kent - 0064"
    planningPlant.Add "0066", "Hamilton - 0066"
    planningPlant.Add "0081", "Jackson - 0081"
    planningPlant.Add "0082", "Adrian - 0082"
    planningPlant.Add "0083", "South Monroe - 0083"
    planningPlant.Add "0084", "Jonesville - 0084"
    planningPlant.Add "0091", "Kalamazoo - 0091"
    planningPlant.Add "0094", "Allegan - 0094"
    planningPlant.Add "0101", "Lansing - 0101"
    planningPlant.Add "0105", "Howell - 0105"
    planningPlant.Add "0121", "Muskegon - 0121"
    planningPlant.Add "0122", "Ludington - 0122"
    planningPlant.Add "0124", "Fremont - 0124"
    planningPlant.Add "0151", "Saginaw - 0151"
    planningPlant.Add "0152", "Bad Axe - 0152"
    planningPlant.Add "0161", "Traverse City - 0161"
    planningPlant.Add "0163", "Manistee - 0163"
    planningPlant.Add "0164", "Cadillac - 0164"
    planningPlant.Add "0165", "Boyne City - 0165"
    planningPlant.Add "0166", "Prudenville - 0166"
    planningPlant.Add "0167", "Benzonia - 0167"
    planningPlant.Add "0171", "Macomb - 0171"
    planningPlant.Add "0181", "Royal Oak - 0181"
    planningPlant.Add "0191", "Livonia - 0191"
    
    Dim SAP As SAP: Set SAP = New SAP
    Call SAP.extractFromSheet
    
    Dim jobType As String
    If SAP.mode = "SYSTEM IMPROVEMENT" Then
        Unload SIDesDoc_FOrm
        SIDesDoc_FOrm.Show
        jobType = SIDesDoc_FOrm.jobType
    Else
        jobType = "Make Ready"
    End If
    
    If jobType = "" Or jobType = "Other" Then
        MsgBox "Need to select a job type, not all jop types have been accounted for"
        Exit Sub
    End If
    
    Dim temp As String: temp = SAP.hdq
    Dim i As Integer
    For i = 1 To Len(temp)
        Dim currentchar As String
        currentchar = Mid(temp, i, 1)
        If IsNumeric(currentchar) Then
            hdq = hdq & currentchar
        End If
    Next i
    If Len(hdq) < 4 Then hdq = Right("0000" & hdq, 4)
    
    constructionMeasure = SAP.constructionMeasure

    
    hdq = IIf(planningPlant.exists(hdq), planningPlant(hdq), "")
    
    Dim output As String
    output = "if (!document.getElementById('Notification')) {" & vbLf
    output = output & vbTab & "alert('Press the create new DD button. You won\'t have to recopy the code.')" & vbLf
    output = output & "} else if (document.getElementById('NewDDLocCount').value == '') {" & vbLf
    output = output & vbTab & " alert('Press enter the number of locations first. You won\'t have to recopy the code.')" & vbLf
    
    If SAP.mode = "SYSTEM IMPROVEMENT" Then
        If jobType = "Pole Replacement" Then
            output = output & "} else if (document.getElementById('OrderType').value != 'ESIC - Electric System Improvement (Capital)') {" & vbLf
            output = output & vbTab & " alert('Press enter the The Order Type to ESIC first. You won\'t have to recopy the code.')" & vbLf
            output = output & "} else {" & vbLf
            output = output & vbTab & "PlannerGrpDD.value = 'LVD - Elect Dist Design'" & vbLf
            output = output & vbTab & "MainWrkCntrDD.value = 'EDLOHLNS - Electric Distribution OH Line '" & vbLf
            output = output & vbTab & "PriorityDD.value = 'Pole Replacements - D'" & vbLf
            output = output & vbTab & "MATDD.value = 'OHL - Overhead Lines'" & vbLf
        End If
        
        If jobType = "Demo" Then
            output = output & "} else if (document.getElementById('OrderType').value != 'ERET - Electric Retirement (Capital)') {" & vbLf
            output = output & vbTab & " alert('Press enter the The Order Type to ERET first. You won\'t have to recopy the code.')" & vbLf
            output = output & "} else {" & vbLf
            output = output & vbTab & "PlannerGrpDD.value = 'LVD - Elect Dist Design'" & vbLf
            output = output & vbTab & "MainWrkCntrDD.value = 'EDLOHLNS - Electric Distribution OH Line '" & vbLf
            output = output & vbTab & "PriorityDD.value = 'Planned - C'" & vbLf
            output = output & vbTab & "MATDD.value = 'OHL - Overhead Lines'" & vbLf
            output = output & vbTab & "NewDDDescription.value = " & Chr(39) & SAP.Permit & " - DEMO" & Chr(39) & vbLf
        Else
            output = output & vbTab & "NewDDDescription.value = " & Chr(39) & SAP.Permit & Chr(39) & vbLf
        End If
        
        output = output & vbTab & "for (let i = 1; i < NewLocationsDiv.childNodes.length; i++) {" & vbLf
        output = output & vbTab & vbTab & "let child = NewLocationsDiv.childNodes[i];" & vbLf
        output = output & vbTab & vbTab & "child.value = 'LOC ' + i" & vbLf
    ElseIf SAP.mode = "MAKE READY" Then
        output = output & "} else if (document.getElementById('OrderType').value != 'DSGN - Engineering Design') {" & vbLf
        output = output & vbTab & " alert('Press enter the The Order Type to DSGN first. You won\'t have to recopy the code.')" & vbLf
        output = output & "} else {" & vbLf
        output = output & vbTab & "PlannerGrpDD.value = 'MKR - Make Ready'" & vbLf
        output = output & vbTab & "MainWrkCntrDD.value = 'ENGLINES - G &/or E Basic Main Design or Line '" & vbLf
        output = output & vbTab & "PriorityDD.value = 'Elec Dist Make Ready - B'" & vbLf
        output = output & vbTab & "MATDD.value = 'EST - Estimate Only Capital  '" & vbLf
        output = output & vbTab & "NewDDDescription.value = " & Chr(39) & SAP.Permit & ": Alternative 1'" & vbLf
        output = output & vbTab & "for (let i = 1; i < NewLocationsDiv.childNodes.length; i++) {" & vbLf
        output = output & vbTab & vbTab & "let child = NewLocationsDiv.childNodes[i];" & vbLf
        output = output & vbTab & vbTab & "child.value = 'L' + i.toString().padStart(3, '0') + ' ALT1'" & vbLf
    End If
    
    If SAP.mode = "MAKE READY" Or jobType <> "Demo" Then
        output = output & vbTab & "document.getElementById('Notification').value = " & Chr(39) & SAP.notification & Chr(39) & vbLf
        output = output & vbTab & "NewDDCM.value = " & Chr(39) & SAP.constructionMeasure & Chr(39) & vbLf
    Else
        output = output & vbTab & "document.getElementById('Notification').value = " & Chr(39) & SAP.demoNotification & Chr(39) & vbLf
        output = output & vbTab & "NewDDCM.value = " & Chr(39) & SAP.demoConstructionMeasure & Chr(39) & vbLf
    End If
    output = output & vbTab & "document.getElementById('PlanningPlant').value = " & Chr(39) & hdq & Chr(39) & vbLf
    output = output & vbTab & "NewDDFuncLoc.value = " & Chr(39) & SAP.functionalLoc & Chr(39) & vbLf
    output = output & vbTab & "}" & vbLf
    output = output & "}" & vbLf
    
    'output = output & vbTab & "OrderType.value = 'DSGN - Engineering Design'" & vbLf

    Dim DataObj As DataObject
    Set DataObj = New DataObject
    DataObj.SetText output
    DataObj.PutInClipboard
    
    MsgBox "Copied code, now go to EAM Create Design Doc and press f12 for console and paste it in."
End Sub

Public Sub clipboardCMDescription()
    Call SendLogMessage("clipboardCMDescription")
    
    Dim SAP As SAP: Set SAP = New SAP
    Call SAP.extractFromSheet
    
    Dim DataObj As DataObject
    Set DataObj = New DataObject
    
    DataObj.SetText SAP.Permit & " Notification " & SAP.notification
    DataObj.PutInClipboard
End Sub

Public Sub clipboardJobPurpose()
    Call SendLogMessage("clipboardJobPurpose")
    
    Dim SAP As SAP: Set SAP = New SAP
    Call SAP.extractFromSheet
    
    Dim DataObj As DataObject
    Set DataObj = New DataObject
    
    DataObj.SetText "MAKE READY FOR " & UCase(SAP.applicant) & vbLf & SAP.notification & vbLf & SAP.Permit
    DataObj.PutInClipboard
End Sub

Public Function compactLocations(dict As Scripting.Dictionary) As String
    Dim result As String
    
    If dict.count > 1 Then
        For Each key In dict
            If dict(key).count > 0 Then
                If result <> "" Then result = result & vbLf
                result = result & "LOC " & compactListString(dict(key)) & " " & key
            End If
        Next key
    ElseIf dict.count = 0 Then
        result = ""
    Else
        result = dict.keys()(0)
    End If
    
    compactLocations = result
End Function

Public Function getCountyCode(ByVal county As String) As String
    Dim counties As Scripting.Dictionary: Set counties = New Scripting.Dictionary
    counties.Add "Alcona", 1
    counties.Add "Allegan", 2
    counties.Add "Alpena", 4
    counties.Add "Antrim", 5
    counties.Add "Arenac", 6
    counties.Add "Barry", 8
    counties.Add "Bay", 9
    counties.Add "Benzie", 10
    counties.Add "Berrien", 11
    counties.Add "Branch", 12
    counties.Add "Calhoun", 13
    counties.Add "Cass", 14
    counties.Add "Charlevoix", 15
    counties.Add "Cheboygan", 16
    counties.Add "Clare", 18
    counties.Add "Clinton", 19
    counties.Add "Crawford", 20
    counties.Add "Eaton", 23
    counties.Add "Emmet", 24
    counties.Add "Genesee", 25
    counties.Add "Gladwin", 26
    counties.Add "Grand Traverse", 28
    counties.Add "Gratiot", 29
    counties.Add "Hillsdale", 30
    counties.Add "Huron", 32
    counties.Add "Ingham", 33
    counties.Add "Ionia", 34
    counties.Add "Iosco", 35
    counties.Add "Isabelle", 37
    counties.Add "Jackson", 38
    counties.Add "Kalamazoo", 39
    counties.Add "Kalkaska", 40
    counties.Add "Kent", 41
    counties.Add "Lake", 43
    counties.Add "Leelanau", 45
    counties.Add "Lenawee", 46
    counties.Add "Livingston", 47
    counties.Add "Macomb", 50
    counties.Add "Manistee", 51
    counties.Add "Mason", 53
    counties.Add "Mecosta", 54
    counties.Add "Midland", 56
    counties.Add "Missuakee", 57
    counties.Add "Monroe", 58
    counties.Add "Montcalm", 59
    counties.Add "Montmorency", 60
    counties.Add "Muskegon", 61
    counties.Add "Newaygo", 62
    counties.Add "Oakland", 63
    counties.Add "Oceana", 64
    counties.Add "Ogemaw", 65
    counties.Add "Osceola", 67
    counties.Add "Oscoda", 68
    counties.Add "Ostego", 69
    counties.Add "Ottawa", 70
    counties.Add "Presque Isle", 71
    counties.Add "Roscommon", 72
    counties.Add "Saginaw", 73
    counties.Add "St. Clair", 74
    counties.Add "St. Joseph", 75
    counties.Add "Sanilac", 76
    counties.Add "Shiawassee", 78
    counties.Add "Tuscola", 79
    counties.Add "Van Buren", 80
    counties.Add "Washtenaw", 81
    counties.Add "Wayne", 82
    counties.Add "Wexford", 83
    
    county = WorksheetFunction.Proper(county)
    If counties.exists(county) Then
        getCountyCode = counties(county)
    Else
        getCountyCode = ""
    End If
End Function

Public Function getHDQCode(ByVal hdq As String) As String
    Dim headquarters As Scripting.Dictionary: Set headquarters = New Scripting.Dictionary
    headquarters.Add "ADR", 82
    headquarters.Add "ALM", 11
    headquarters.Add "BCK", 21
    headquarters.Add "BCY", 31
    headquarters.Add "BIG", 17
    headquarters.Add "BNC", 165
    headquarters.Add "CAD", 164
    headquarters.Add "CLR", 12
    headquarters.Add "CLD", 23
    headquarters.Add "FLT", 51
    headquarters.Add "FRE", 124
    headquarters.Add "GRA", 61
    headquarters.Add "GRE", 63
    headquarters.Add "GRN", 64
    headquarters.Add "GRW", 61
    headquarters.Add "GVL", 13
    headquarters.Add "HAM", 66
    headquarters.Add "HST", 25
    headquarters.Add "HNR", 167
    headquarters.Add "HWL", 105
    headquarters.Add "JAC", 81
    headquarters.Add "KAL", 91
    headquarters.Add "LAN", 101
    headquarters.Add "LAP", 42
    headquarters.Add "LUD", 122
    headquarters.Add "MAC", 171
    headquarters.Add "MDL", 32
    headquarters.Add "MUS", 121
    headquarters.Add "OWS", 52
    headquarters.Add "PON", 41
    headquarters.Add "ROK", 181
    headquarters.Add "SAG", 151
    headquarters.Add "SMN", 83
    headquarters.Add "TWS", 34
    headquarters.Add "TRA", 161
    headquarters.Add "WBR", 38
    headquarters.Add "WES", 191
    headquarters.Add "ADRIAN", 82
    headquarters.Add "ALMA", 11
    headquarters.Add "BATTLE CREEK", 21
    headquarters.Add "BAY CITY", 31
    headquarters.Add "BIG RAPIDS", 17
    headquarters.Add "BOYNE CITY", 165
    headquarters.Add "CADILLAC", 164
    headquarters.Add "CLARE", 12
    headquarters.Add "COLDWATER", 23
    headquarters.Add "FLINT", 51
    headquarters.Add "FREEMONT", 124
    headquarters.Add "GRAND RAPIDS", 61
    headquarters.Add "GRAND RAPIDS E", 63
    headquarters.Add "GRAND RAPIDS N", 64
    headquarters.Add "GRAND RAPIDS W", 61
    headquarters.Add "GREENVILLE", 13
    headquarters.Add "HAMILTON", 66
    headquarters.Add "HASTINGS", 25
    headquarters.Add "HONOR", 167
    headquarters.Add "BENZONIA", 167
    headquarters.Add "HOWELL", 105
    headquarters.Add "JACKSON", 81
    headquarters.Add "KALAMAZOO", 91
    headquarters.Add "LANSING", 101
    headquarters.Add "LAPEER", 42
    headquarters.Add "LUDINGTON", 122
    headquarters.Add "MACOMB", 171
    headquarters.Add "MIDLAND", 32
    headquarters.Add "MUSKEGON", 121
    headquarters.Add "OWOSSO", 52
    headquarters.Add "PONTIAC", 41
    headquarters.Add "ROYAL OAK", 181
    headquarters.Add "SAGINAW", 151
    headquarters.Add "SOUTH MONROE", 83
    headquarters.Add "TAWAS", 34
    headquarters.Add "TRAVERSE CITY", 161
    headquarters.Add "WEST BRANCH", 38
    headquarters.Add "WEST WAYNE", 191

    If headquarters.exists(hdq) Then
        getHDQCode = headquarters(hdq)
    Else
        getHDQCode = ""
    End If
End Function
