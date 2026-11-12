Attribute VB_Name = "SAPEmails"
Dim outlookApp As Object
Dim outlookMail As Object
Dim signiture As String

Public Sub EmailFielderRequest()
    Dim project As project: Set project = New project
    Dim SAP As SAP: Set SAP = New SAP
    Call project.extractFromSheets
    Call SAP.extractFromSheet
    
    On Error Resume Next
    
    Call SendLogMessage("EmailFielderRequest")
    
    Dim project As ProjectData
    Set project = New ProjectData
    project.ExtractExcel
    
    Dim locationWork As Scripting.Dictionary: Set locationWork = New Scripting.Dictionary
    
    Dim stakeNumber As Long: stakeNumber = 0
    Dim pole As pole
    For Each pole In project.poles
        If pole.PoleStakingComplete Then
            stakeNumber = stakeNumber + 1
            locationWork(CInt(pole.location)) = locationWork(CInt(pole.location)) & "Location " & IIf(pole.location = "", "[LOCATION]", pole.location) & " - M1P" & pole.Number & " ([X] Anchor(s)) CEID#" & IIf(pole.id = "", "[CEID]", pole.id) & " - " & IIf(pole.address = "", [XHOUSE#, XX RD, TOWN, MI ZIP], pole.address) & "<br>"
            locationWork(CInt(pole.location)) = locationWork(CInt(pole.location)) & "[Extend existing CE anchor SSW from 12' to 17']<br>"
            locationWork(CInt(pole.location)) = locationWork(CInt(pole.location)) & "[Install new anchor NNW at 10']"
            If Not IsNumeric(pole.location) Then
                MsgBox "Pole locations need to be labled numerical for this script, check your locations": Exit Sub
            End If
        End If
    Next pole
    
    Dim tree_ws As Worksheet
    Set tree_ws = ThisWorkbook.sheets("Tree Row")
    
    tagbody = ""
    Dim tagNumber As Long
    tagNumber = 0
    Dim cell As Range
    Dim treeLocations As Scripting.Dictionary
    Set treeLocations = New Scripting.Dictionary
    For Each cell In tree_ws.Range("L6:L" & 100)
        If LCase(Trim(cell.value)) <> "yes" And LCase(Trim(cell.value)) <> "y" And cell.offSet(0, -10).value <> "" And cell.offSet(0, 7).value = True Then
            Set pole = Nothing
            Set pole = project.Find(cell.offSet(0, -10))
            tagNumber = tagNumber + 1
            If Not IsNumeric(pole.location) Then
             '   MsgBox "Pole locations need to be labled numerical for this script, check your locations": Exit Sub
            End If
            If Not locationWork.exists(CInt(pole.location)) Then
                If Not pole Is Nothing Then
                    locationWork(CInt(pole.location)) = locationWork(CInt(pole.location)) & "Location " & cell.offSet(0, -8).value & " - M1P" & cell.offSet(0, -10).value & " " & IIf(pole.ReplacePole, "Pole Replace", "(Pole Replace)") & " CEID#" & cell.offSet(0, -9).text & " - " & IIf(pole.address = "", "[ADDRESS]", pole.address)
                Else
                    locationWork(CInt(pole.location)) = locationWork(CInt(pole.location)) & "Location " & cell.offSet(0, -8).value & " - M1P" & cell.offSet(0, -10).value & " (Pole Replace) CEID#" & cell.offSet(0, -9).text & " - " & IIf(pole.address = "", "[ADDRESS]", pole.address)
                End If
            End If
            locationWork(CInt(pole.location)) = locationWork(CInt(pole.location)) & "<br>Customer contacted needed for tree trimming at " & WorksheetFunction.Proper(cell.offSet(0, 1).value)
            If Not treeLocations.exists(Trim(cell.offSet(0, -8).value)) Then treeLocations.Add Trim(cell.offSet(0, -8).value), Nothing
        End If
    Next cell
    
    MsgBox "Please remember to attach supporting documents like pictures/prints to help the fielder." & vbLf & "The email generated is just a template and needs to be edited before sending, take care of anything surrounded by ( ) and [ ]" & vbLf & "( ) mean something might be correct and anything in [ ] is definitely wrong and just there to be easily edited"
    
    Call initEmail
    recipient = "aramseyer@trccompanies.com; Zach.williams@cmsenergy.com; TRC_CE_mailbox@trccompanies.com;"
    ccRecipient = "TRC@cmsenergy.com;"
    
    For Each leadEmail In SAP.leadEmails
        If InStr(1, recipient, leadEmail, vbTextCompare) = 0 And InStr(1, ccRecipient, leadEmail, vbTextCompare) = 0 Then
            ccRecipient = ccRecipient & leadEmail
        End If
    Next leadEmail
    
    If stakeNumber > 0 And tagNumber > 0 Then
        Subject = project.notification & " - Staking/Customer Contact Request - " & project.town
    End If
    If stakeNumber > 0 Then
        If Subject = "" Then Subject = project.notification & " - Staking Request - " & project.town
        stakingBody = "I need (" & stakeNumber & ") guy" & IIf(stakeNumber > 1, "s", "") & " staked for " & stakeNumber & " location" & IIf(stakeNumber > 1, "s", "") & " on this job. See Attached Design print for details.<br>"
        stakingBody = stakingBody & "Write Notification # and ""Proposed Anchor"" on stakes<br>"
        stakingBody = stakingBody & "Take 2 Pictures of each stake.<br><br>"
        body = body & stakingBody
    End If
    If tagNumber > 0 Then
        If Subject = "" Then Subject = project.notification & " - Customer Contact for Tree Trimming - " & project.town
        tagbody = "I need " & tagNumber & " customer" & IIf(tagNumber <> 1, "s", "") & " contacted for tree trimming for " & treeLocations.count & " location" & IIf(treeLocations.count <> 1, "s", "") & " on this job. See attached design print for details.<br><br>"
        body = body & tagbody
    End If
    
    Dim keysArray As Variant
    keysArray = locationWork.keys
    For i = LBound(keysArray) To UBound(keysArray)
        sortedKey = Application.WorksheetFunction.Small(keysArray, i + 1)
        body = body & locationWork(sortedKey) & "<br><br>"
    Next i
    
    If stakeNumber > 0 Or tagNumber > 0 Then
        body = body & "Customer contact always required if possible, or leave door tag with my name & number<br>"
        body = body & IIf(project.designerName = "", "[NAME]", project.designerName) & " - " & IIf(project.designerNumber = "", "[NUMBER]", project.designerNumber) & "<br><br>"
    End If
    
    body = body & signiture
    
    outlookMail.To = recipient
    outlookMail.cc = ccRecipient
    outlookMail.Subject = Subject
    outlookMail.htmlbody = body
    
End Sub

Public Sub EmailForesty()
    Call setValues

    On Error Resume Next

    Call SendLogMessage("EmailForesty")

    recipient = "zach.williams@cmsenergy.com;" & "[FORESTER]"
    ccRecipient = "TRC@cmsenergy.com;"
    
    For Each leadEmail In SAP.leadEmails
        If InStr(1, recipient, leadEmail, vbTextCompare) = 0 And InStr(1, ccRecipient, leadEmail, vbTextCompare) = 0 Then
            ccRecipient = ccRecipient & leadEmail
        End If
    Next leadEmail
    
    forcDesc = CopyFORCDescription("<br>")
    
    Call initEmail
    
    Subject = "FORESTRY ESTIMATE - " & notification & " " & Format(Date, "mm/dd/yyyy")
    body = "I am designing a Make Ready job in the " & twp & "MI area that requires a forestry estimate. <br><br>" _
        & "Please see Doc warehouse or KDrive for Forestry Print and Tree Work ROW Control Forms.<br>" _
        & forcDesc & "<br>" _
        & signiture

    outlookMail.To = recipient
    outlookMail.cc = ccRecipient
    outlookMail.Subject = Subject
    outlookMail.htmlbody = body
    
    MsgBox "Put a comment in smartsheets for when the forestry was submitted."
End Sub

Public Sub EmailDesignReview()
    Call SendLogMessage("EmailDesignReview")
    
    Dim SAP As SAP: Set SAP = New SAP
    Call SAP.extractFromSheet
    
    Dim data_ws As Worksheet: Set data_ws = ThisWorkbook.sheets("DATA")
    Dim emailCell As Range
    Set emailCell = data_ws.cells.Find("Design Review Email")
    If Not emailCell Is Nothing Then
        If emailCell.offSet(1, 0).value <> "" Then
            recipient = emailCell.offSet(1, 0).value
        End If
        If emailCell.offSet(2, 0).value <> "" Then
            ccRecipient = emailCell.offSet(2, 0).value & " " & additionalCC & ";"
        End If
    End If
    
    For Each leadEmail In SAP.leadEmails
        If InStr(1, recipient, leadEmail, vbTextCompare) = 0 And InStr(1, ccRecipient, leadEmail, vbTextCompare) = 0 Then
            ccRecipient = ccRecipient & leadEmail
        End If
    Next leadEmail
    
    Call initEmail
    
    Subject = SAP.notification & " - Design Review"
    body = "Important Notes:<br><br>" _
        & "I have finished design on this this " & SAP.numPoles & " pole " & Application.WorksheetFunction.Proper(SAP.mode) & " job.<br>" _
        & signiture

    outlookMail.To = recipient
    outlookMail.cc = ccRecipient
    outlookMail.Subject = Subject
    outlookMail.htmlbody = body
    
    MsgBox "Update smart sheets for parker review and number of locations."
End Sub

Public Sub EmailSIInternalReview()
    Call SendLogMessage("EmailSIInternalReview")
    
    Dim SAP As SAP: Set SAP = New SAP
    Call SAP.extractFromSheet
    Dim project As project: Set project = New project
    Call project.extractFromSheets
    
    Dim locationCount As Integer
    For Each pole In project.poles
        If pole.location <> "" Then locationCount = locationCount + 1
    Next pole
    
    If SAP.notification = "" Then MsgBox "Missing Notification Number": Exit Sub
    If SAP.mainCost = "" Then MsgBox "Missing Main Notification Cost": Exit Sub
    If SAP.constructionMeasure = "" Then MsgBox "Missing Construction Measure Number. Put N/A if there isn't one": Exit Sub
    If SAP.Permit = "" Then MsgBox "Missing Permit": Exit Sub
    If SAP.designerName = "" Then MsgBox "Missing Designer Name": Exit Sub
    If SAP.dueDate = "" Then MsgBox "Missing Due Date": Exit Sub
    If SAP.numPoles = "" Then MsgBox "Missing Number of Poles": Exit Sub
    If SAP.mainDesDoc = "" Then MsgBox "Missing Main Design Doc Number. Put N/A if there isn't one": Exit Sub
    If SAP.treeWork = "" Then MsgBox "Missing Tree Work. Put N/A if there isn't one": Exit Sub
    If SAP.stakedLocations = "" Then MsgBox "Missing Staking Locations. Put N/A are there aren't any": Exit Sub
    If SAP.stakedBy = "" Then MsgBox "Missing Staked By. Put TBD if it hasn't been staked yet, or N/A if there isn't any": Exit Sub
    If SAP.stakedDate = "" Then MsgBox "Missing Staked Date. Put TBD if it hasn't been staked yet, or N/A if there isn't any": Exit Sub
    If SAP.docWarehouseLink = "" Or SAP.docWarehouseLink = "Link Not Found" Then MsgBox "Missing Doc Warehouse Link.": Exit Sub
    If SAP.num391Poles = "" Then MsgBox "Missing Number of 391 Forms": Exit Sub
    
    Dim data_ws As Worksheet: Set data_ws = ThisWorkbook.sheets("DATA")
    Dim emailCell As Range
    Set emailCell = data_ws.cells.Find("SI Internal Review Email")
    If Not emailCell Is Nothing Then
        If emailCell.offSet(1, 0).value <> "" Then
            recipient = emailCell.offSet(1, 0).value
        End If
        If emailCell.offSet(2, 0).value <> "" Then
            ccRecipient = emailCell.offSet(2, 0).value & " " & additionalCC & ";"
        End If
    End If

    For Each leadEmail In SAP.leadEmails
        If InStr(1, recipient, leadEmail, vbTextCompare) = 0 And InStr(1, ccRecipient, leadEmail, vbTextCompare) = 0 Then
            ccRecipient = ccRecipient & leadEmail
        End If
    Next leadEmail

    If InStr(LCase(SAP.treeWork), "n/a") > 0 Or InStr(LCase(SAP.treeWork), "na") > 0 Or InStr(LCase(SAP.treeWork), "no") > 0 Then
        SAP.treeWork = "No"
    ElseIf InStr(LCase(SAP.treeWork), "yes") = 0 And InStr(SAP.treeWork, "no") = 0 Then
        SAP.treeWork = "Yes" & vbLf & SAP.treeWork
    End If
    
    Call initEmail
    
    Subject = SAP.notification & " - " & SAP.Permit & " - System Improvement - Internal Review"
    body = "Important Design Notes:<br><br>" _
        & "SAP Description:<br>" _
        & "Designed by: " & SAP.designerName & Chr(32) & SAP.designerNumber & "<br>" _
        & "Due Date: " & SAP.dueDate & "<br>" _
        & "# of CE Locations: " & locationCount & "<br>" _
        & "Total Planning Hours: " & SAP.totalHours & "<br>" _
        & "Total Cost: " & SAP.totalCost & "<br><br>" _
        & "<u>Main Notification</u><br>" _
        & "Notification: " & SAP.notification & "<br>" _
        & "Design Document: " & SAP.mainDesDoc & "<br>" _
        & "Planning Hours: " & SAP.mainHours & "<br>" _
        & "Cost: " & SAP.mainCost & "<br><br>" _
        & "LLTM Notification: " & SAP.longLeadNotification & "<br>" _
        & "LLTM Order: " & SAP.longLeadOrder & "<br><br>"
        
        For i = 1 To SAP.serviceNotifications.count
            body = body & "<u>OH Service Notification</u><br>" _
                & "Notification: " & SAP.serviceNotifications(i) & "<br>" _
                & "Design Document: " & IIf(SAP.serviceDesDocs.count >= i, SAP.serviceDesDocs(i), "") & "<br>" _
                & "Planning Hours: " & IIf(SAP.serviceHours.count >= i, SAP.serviceHours(i), "") & "<br>" _
                & "Cost: " & IIf(SAP.serviceCosts.count >= i, SAP.serviceCosts(i), "") & "<br><br>"
        Next i
        
        If SAP.demoNotification <> "" Then
            body = body & "<u>Demo Notification</u><br>" _
                & "Notification: " & SAP.demoNotification & "<br>" _
                & "Design Document: " & SAP.demoDesDoc & "<br>" _
                & "Planning Hours: " & SAP.demoHours & "<br>" _
                & "Cost: " & SAP.demoCost & "<br><br>"
        End If
        
        body = body & "Tree - Trimming: " & Replace(SAP.treeWork, vbLf, "<br>") & "<br><br>" _
            & "# 391 Forms: " & SAP.num391Poles & "<br><br>" _
            & "CA/Violations NJUNS Project Number: " & SAP.njunsProject & "<br>" _
            & "# of Locations: " & SAP.njunsLocations & "<br><br>" _
            & "Design Location Numbers Staked: " & SAP.stakedLocations & "<br>" _
            & "Staked by: " & SAP.stakedBy & "<br>" _
            & "Date Pole Staked: " & SAP.stakedDate & "<br><br>" _
            & "Doc Warehouse" & "<br>" _
            & "<a href=""" & SAP.docWarehouseLink & """>" & SAP.notification & " - Doc Warehouse Link" & "</a>" _
            & signiture _

    outlookMail.To = recipient
    outlookMail.cc = ccRecipient
    outlookMail.Subject = Subject
    outlookMail.htmlbody = body
    
    MsgBox "Update smart sheets to internal review."
End Sub

Public Sub EmailSICEReview()
    Call SendLogMessage("EmailSICEReview")
    
    Dim SAP As SAP: Set SAP = New SAP
    Call SAP.extractFromSheet
    Dim project As project: Set project = New project
    Call project.extractFromSheets
    
    Dim locationCount As Integer
    For Each pole In project.poles
        If pole.location <> "" Then locationCount = locationCount + 1
    Next pole
    
    If SAP.notification = "" Then MsgBox "Missing Notification Number": Exit Sub
    If SAP.mainCost = "" Then MsgBox "Missing Main Notification Cost": Exit Sub
    If SAP.constructionMeasure = "" Then MsgBox "Missing Construction Measure Number. Put N/A if there isn't one": Exit Sub
    If SAP.Permit = "" Then MsgBox "Missing Permit": Exit Sub
    If SAP.designerName = "" Then MsgBox "Missing Designer Name": Exit Sub
    If SAP.dueDate = "" Then MsgBox "Missing Due Date": Exit Sub
    If SAP.numPoles = "" Then MsgBox "Missing Number of Poles": Exit Sub
    If SAP.mainDesDoc = "" Then MsgBox "Missing Main Design Doc Number. Put N/A if there isn't one": Exit Sub
    If SAP.treeWork = "" Then MsgBox "Missing Tree Work. Put N/A if there isn't one": Exit Sub
    If SAP.stakedLocations = "" Then MsgBox "Missing Staking Locations. Put N/A are there aren't any": Exit Sub
    If SAP.stakedBy = "" Then MsgBox "Missing Staked By. Put TBD if it hasn't been staked yet, or N/A if there isn't any": Exit Sub
    If SAP.stakedDate = "" Then MsgBox "Missing Staked Date. Put TBD if it hasn't been staked yet, or N/A if there isn't any": Exit Sub
    If SAP.docWarehouseLink = "" Or SAP.docWarehouseLink = "Link Not Found" Then MsgBox "Missing Doc Warehouse Link.": Exit Sub
    If SAP.num391Poles = "" Then MsgBox "Missing Number of 391 Forms": Exit Sub
    
    Dim data_ws As Worksheet: Set data_ws = ThisWorkbook.sheets("DATA")
    Dim emailCell As Range
    Set emailCell = data_ws.cells.Find("SI CE Review Email")
    If Not emailCell Is Nothing Then
        If emailCell.offSet(1, 0).value <> "" Then
            recipient = emailCell.offSet(1, 0).value
        End If
        If emailCell.offSet(2, 0).value <> "" Then
            ccRecipient = emailCell.offSet(2, 0).value & " " & additionalCC & ";"
        End If
    End If

    For Each leadEmail In SAP.leadEmails
        If InStr(1, recipient, leadEmail, vbTextCompare) = 0 And InStr(1, ccRecipient, leadEmail, vbTextCompare) = 0 Then
            ccRecipient = ccRecipient & leadEmail
        End If
    Next leadEmail

    If InStr(LCase(SAP.treeWork), "n/a") > 0 Or InStr(LCase(SAP.treeWork), "na") > 0 Or InStr(LCase(SAP.treeWork), "no") > 0 Then
        SAP.treeWork = "No"
    ElseIf InStr(LCase(SAP.treeWork), "yes") = 0 And InStr(SAP.treeWork, "no") = 0 Then
        SAP.treeWork = "Yes" & vbLf & SAP.treeWork
    End If
    
    Call initEmail
    
    Subject = SAP.notification & " - " & SAP.Permit & " - System Improvement - Ready for Review"
    body = "Important Design Notes:<br><br>" _
        & "SAP Description:<br>" _
        & "Designed by: " & SAP.designerName & Chr(32) & SAP.designerNumber & "<br>" _
        & "Due Date: " & SAP.dueDate & "<br>" _
        & "# of CE Locations: " & locationCount & "<br>" _
        & "Total Planning Hours: " & SAP.totalHours & "<br>" _
        & "Total Cost: " & SAP.totalCost & "<br><br>" _
        & "<u>Main Notification</u><br>" _
        & "Notification: " & SAP.notification & "<br>" _
        & "Design Document: " & SAP.mainDesDoc & "<br>" _
        & "Planning Hours: " & SAP.mainHours & "<br>" _
        & "Cost: " & SAP.mainCost & "<br><br>" _
        & "LLTM Notification: " & SAP.longLeadNotification & "<br>" _
        & "LLTM Order: " & SAP.longLeadOrder & "<br><br>"
        
        For i = 1 To SAP.serviceNotifications.count
            body = body & "<u>OH Service Notification</u><br>" _
                & "Notification: " & SAP.serviceNotifications(i) & "<br>" _
                & "Design Document: " & IIf(SAP.serviceDesDocs.count >= i, SAP.serviceDesDocs(i), "") & "<br>" _
                & "Planning Hours: " & IIf(SAP.serviceHours.count >= i, SAP.serviceHours(i), "") & "<br>" _
                & "Cost: " & IIf(SAP.serviceCosts.count >= i, SAP.serviceCosts(i), "") & "<br><br>"
        Next i
        
        If SAP.demoNotification <> "" Then
            body = body & "<u>Demo Notification</u><br>" _
                & "Notification: " & SAP.demoNotification & "<br>" _
                & "Design Document: " & SAP.demoDesDoc & "<br>" _
                & "Planning Hours: " & SAP.demoHours & "<br>" _
                & "Cost: " & SAP.demoCost & "<br><br>"
        End If
        
        body = body & "Tree - Trimming: " & Replace(SAP.treeWork, vbLf, "<br>") & "<br><br>" _
            & "# 391 Forms: " & SAP.num391Poles & "<br><br>" _
            & "CA/Violations NJUNS Project Number: " & SAP.njunsProject & "<br>" _
            & "# of Locations: " & SAP.njunsLocations & "<br><br>" _
            & "Design Location Numbers Staked: " & SAP.stakedLocations & "<br>" _
            & "Staked by: " & SAP.stakedBy & "<br>" _
            & "Date Pole Staked: " & SAP.stakedDate & "<br><br>" _
            & "Doc Warehouse" & "<br>" _
            & "<a href=""" & SAP.docWarehouseLink & """>" & SAP.notification & " - Doc Warehouse Link" & "</a>" _
            & signiture _

    outlookMail.To = recipient
    outlookMail.cc = ccRecipient
    outlookMail.Subject = Subject
    outlookMail.htmlbody = body
    
    MsgBox "Update smart sheets to CE review."
End Sub

Public Sub EmailSIDesignComplete()
    Call SendLogMessage("EmailSIDesignComplete")
    
    Dim SAP As SAP: Set SAP = New SAP
    Call SAP.extractFromSheet
    Dim project As project: Set project = New project
    Call project.extractFromSheets
    
    Dim locationCount As Integer
    For Each pole In project.poles
        If pole.location <> "" Then locationCount = locationCount + 1
    Next pole
    
    If SAP.notification = "" Then MsgBox "Missing Notification Number": Exit Sub
    If SAP.mainCost = "" Then MsgBox "Missing Main Notification Cost": Exit Sub
    If SAP.constructionMeasure = "" Then MsgBox "Missing Construction Measure Number. Put N/A if there isn't one": Exit Sub
    If SAP.Permit = "" Then MsgBox "Missing Permit": Exit Sub
    If SAP.designerName = "" Then MsgBox "Missing Designer Name": Exit Sub
    If SAP.dueDate = "" Then MsgBox "Missing Due Date": Exit Sub
    If SAP.numPoles = "" Then MsgBox "Missing Number of Poles": Exit Sub
    If SAP.mainDesDoc = "" Then MsgBox "Missing Main Design Doc Number. Put N/A if there isn't one": Exit Sub
    If SAP.treeWork = "" Then MsgBox "Missing Tree Work. Put N/A if there isn't one": Exit Sub
    If SAP.stakedLocations = "" Then MsgBox "Missing Staking Locations. Put N/A are there aren't any": Exit Sub
    If SAP.stakedBy = "" Then MsgBox "Missing Staked By. Put TBD if it hasn't been staked yet, or N/A if there isn't any": Exit Sub
    If SAP.stakedDate = "" Then MsgBox "Missing Staked Date. Put TBD if it hasn't been staked yet, or N/A if there isn't any": Exit Sub
    If SAP.docWarehouseLink = "" Or SAP.docWarehouseLink = "Link Not Found" Then MsgBox "Missing Doc Warehouse Link.": Exit Sub
    If SAP.num391Poles = "" Then MsgBox "Missing Number of 391 Forms": Exit Sub
    
    Dim data_ws As Worksheet: Set data_ws = ThisWorkbook.sheets("DATA")
    Dim emailCell As Range
    Set emailCell = data_ws.cells.Find("SI Design Complete Email")
    If Not emailCell Is Nothing Then
        If emailCell.offSet(1, 0).value <> "" Then
            recipient = emailCell.offSet(1, 0).value
        End If
        If emailCell.offSet(2, 0).value <> "" Then
            ccRecipient = emailCell.offSet(2, 0).value & " " & additionalCC & ";"
        End If
    End If

    For Each leadEmail In SAP.leadEmails
        If InStr(1, recipient, leadEmail, vbTextCompare) = 0 And InStr(1, ccRecipient, leadEmail, vbTextCompare) = 0 Then
            ccRecipient = ccRecipient & leadEmail
        End If
    Next leadEmail

    If InStr(LCase(SAP.treeWork), "n/a") > 0 Or InStr(LCase(SAP.treeWork), "na") > 0 Or InStr(LCase(SAP.treeWork), "no") > 0 Then
        SAP.treeWork = "No"
    ElseIf InStr(LCase(SAP.treeWork), "yes") = 0 And InStr(SAP.treeWork, "no") = 0 Then
        SAP.treeWork = "Yes" & vbLf & SAP.treeWork
    End If
    
    Call initEmail
    
    Subject = SAP.notification & " - " & SAP.Permit & " - Design Completed - " & Format(Date, "mm/dd/yyyy")
    body = "The " & SAP.Permit & " System Improvement job has been completed.<br><br>" _
        & "SAP Description:<br>" _
        & "Designed by: " & SAP.designerName & Chr(32) & SAP.designerNumber & "<br>" _
        & "Due Date: " & SAP.dueDate & "<br>" _
        & "# of CE Locations: " & locationCount & "<br>" _
        & "Total Planning Hours: " & SAP.totalHours & "<br>" _
        & "Total Cost: " & SAP.totalCost & "<br><br>" _
        & "<u>Main Notification</u><br>" _
        & "Notification: " & SAP.notification & "<br>" _
        & "Design Document: " & SAP.mainDesDoc & "<br>" _
        & "Planning Hours: " & SAP.mainHours & "<br>" _
        & "Cost: " & SAP.mainCost & "<br><br>" _
        & "LLTM Notification: " & SAP.longLeadNotification & "<br>" _
        & "LLTM Order: " & SAP.longLeadOrder & "<br><br>"
        
        For i = 1 To SAP.serviceNotifications.count
            body = body & "<u>OH Service Notification</u><br>" _
                & "Notification: " & SAP.serviceNotifications(i) & "<br>" _
                & "Design Document: " & IIf(SAP.serviceDesDocs.count >= i, SAP.serviceDesDocs(i), "") & "<br>" _
                & "Planning Hours: " & IIf(SAP.serviceHours.count >= i, SAP.serviceHours(i), "") & "<br>" _
                & "Cost: " & IIf(SAP.serviceCosts.count >= i, SAP.serviceCosts(i), "") & "<br><br>"
        Next i
        
        If SAP.demoNotification <> "" Then
            body = body & "<u>Demo Notification</u><br>" _
                & "Notification: " & SAP.demoNotification & "<br>" _
                & "Design Document: " & SAP.demoDesDoc & "<br>" _
                & "Planning Hours: " & SAP.demoHours & "<br>" _
                & "Cost: " & SAP.demoCost & "<br><br>"
        End If
        
        body = body & "Tree - Trimming: " & Replace(SAP.treeWork, vbLf, "<br>") & "<br><br>" _
            & "# 391 Forms: " & SAP.num391Poles & "<br><br>" _
            & "CA/Violations NJUNS Project Number: " & SAP.njunsProject & "<br>" _
            & "# of Locations: " & SAP.njunsLocations & "<br><br>" _
            & "Design Location Numbers Staked: " & SAP.stakedLocations & "<br>" _
            & "Staked by: " & SAP.stakedBy & "<br>" _
            & "Date Pole Staked: " & SAP.stakedDate & "<br><br>" _
            & "Doc Warehouse" & "<br>" _
            & "<a href=""" & SAP.docWarehouseLink & """>" & SAP.notification & " - Doc Warehouse Link" & "</a>" _
            & signiture _

    outlookMail.To = recipient
    outlookMail.cc = ccRecipient
    outlookMail.Subject = Subject
    outlookMail.htmlbody = body
    
    MsgBox "Update smart sheets to submitted."
End Sub

Public Sub EmailMKRInternalReview()
    Call SendLogMessage("EmailMKRInternalReview")
    
    Dim SAP As SAP: Set SAP = New SAP
    Call SAP.extractFromSheet
    
    If SAP.notification = "" Then MsgBox "Missing Notification Number": Exit Sub
    If SAP.constructionMeasure = "" Then MsgBox "Missing Construction Measure Number. Put N/A if there isn't one": Exit Sub
    If SAP.Permit = "" Then MsgBox "Missing Permit": Exit Sub
    If SAP.designerName = "" Then MsgBox "Missing Designer Name": Exit Sub
    If SAP.dueDate = "" Then MsgBox "Missing Due Date": Exit Sub
    If SAP.numPoles = "" Then MsgBox "Missing Number of Poles": Exit Sub
    If SAP.applicant = "" Then MsgBox "Missing Applicant": Exit Sub
    If SAP.alt1DesDoc = "" Then MsgBox "Missing Design Doc Alt1 Number. Put N/A if there isn't one": Exit Sub
    If SAP.treeWork = "" Then MsgBox "Missing Tree Work. Put N/A if there isn't one": Exit Sub
    If SAP.numDeniedPoles = "" Then MsgBox "Missing Number of Denied Poles": Exit Sub
    If SAP.stakedLocations = "" Then MsgBox "Missing Staking Locations. Put N/A are there aren't any": Exit Sub
    If SAP.stakedBy = "" Then MsgBox "Missing Staked By. Put TBD if it hasn't been staked yet, or N/A if there isn't any": Exit Sub
    If SAP.stakedDate = "" Then MsgBox "Missing Staked Date. Put TBD if it hasn't been staked yet, or N/A if there isn't any": Exit Sub
    If SAP.num391Poles = "" Then MsgBox "Missing Number of 391 Forms": Exit Sub
    
    Dim data_ws As Worksheet: Set data_ws = ThisWorkbook.sheets("DATA")
    Dim emailCell As Range
    Set emailCell = data_ws.cells.Find("MKR Internal Review Email")
    If Not emailCell Is Nothing Then
        If emailCell.offSet(1, 0).value <> "" Then
            recipient = emailCell.offSet(1, 0).value
        End If
        If emailCell.offSet(2, 0).value <> "" Then
            ccRecipient = emailCell.offSet(2, 0).value & " " & additionalCC & ";"
        End If
    End If

    For Each leadEmail In SAP.leadEmails
        If InStr(1, recipient, leadEmail, vbTextCompare) = 0 And InStr(1, ccRecipient, leadEmail, vbTextCompare) = 0 Then
            ccRecipient = ccRecipient & leadEmail
        End If
    Next leadEmail

    If InStr(LCase(SAP.treeWork), "n/a") > 0 Or InStr(LCase(SAP.treeWork), "na") > 0 Or InStr(LCase(SAP.treeWork), "no") > 0 Then SAP.treeWork = "No"
    If InStr(LCase(SAP.treeWork), "yes") = 0 And InStr(SAP.treeWork, "no") = 0 Then SAP.treeWork = "Yes" & "<br>" & SAP.treeWork
    
    Call initEmail
    
    Subject = SAP.notification & " - " & SAP.constructionMeasure & " - " & SAP.Permit & " - Internal Review " & Format(Date, "mm/dd/yyyy")
    body = "Important Design Notes:<br><br>" _
        & "Designed by: " & SAP.designerName & Chr(32) & SAP.designerNumber & "<br><br>" _
        & "Due Date: " & SAP.dueDate & "<br>" _
        & "# Poles: " & SAP.numPoles & "<br>" _
        & "Applicant: " & SAP.applicant & "<br>" _
        & "Notification: " & SAP.notification & "<br>" _
        & "Permit: " & SAP.Permit & "<br><br>" _
        & "Alt1 Design Document: " & SAP.alt1DesDoc & "<br><br>" _
        & "Alt2 Design Document: " & SAP.alt2DesDoc & "<br><br>" _
        & "Alt3 Design Document: " & SAP.alt3DesDoc & "<br><br>" _
        & "Tree - Trimming: " & Replace(SAP.treeWork, vbLf, "<br>") & "<br><br>" _
        & "# 391 Forms: " & SAP.num391Poles & "<br><br>" _
        & "# of Denied Poles: " & SAP.numDeniedPoles & "<br><br>" _
        & "Design Location Numbers Staked: " & SAP.stakedLocations & "<br>" _
        & "Staked by: " & SAP.stakedBy & "<br>" _
        & "Date Pole Staked: " & SAP.stakedDate & "<br><br>" _
        & "Any questions please feel free to call or e-mail the designer." & "<br>" _
        & signiture _

    outlookMail.To = recipient
    outlookMail.cc = ccRecipient
    outlookMail.Subject = Subject
    outlookMail.htmlbody = body
    
    MsgBox "Update smart sheets to internal review."
End Sub

Public Sub EmailMKRCEReview()
    Call SendLogMessage("EmailMKRCEReview")
    
    Dim SAP As SAP: Set SAP = New SAP
    Call SAP.extractFromSheet
    
    If SAP.notification = "" Then MsgBox "Missing Notification Number": Exit Sub
    If SAP.constructionMeasure = "" Then MsgBox "Missing Construction Measure Number. Put N/A if there isn't one": Exit Sub
    If SAP.Permit = "" Then MsgBox "Missing Permit": Exit Sub
    If SAP.designerName = "" Then MsgBox "Missing Designer Name": Exit Sub
    If SAP.dueDate = "" Then MsgBox "Missing Due Date": Exit Sub
    If SAP.numPoles = "" Then MsgBox "Missing Number of Poles": Exit Sub
    If SAP.applicant = "" Then MsgBox "Missing Applicant": Exit Sub
    If SAP.alt1DesDoc = "" Then MsgBox "Missing Design Doc Alt1 Number. Put N/A if there isn't one": Exit Sub
    If SAP.alt2DesDoc = "" Then MsgBox "Missing Design Doc Alt2 Number. Put N/A if there isn't one": Exit Sub
    If SAP.alt3DesDoc = "" Then MsgBox "Missing Design Doc Alt3 Number. Put N/A if there isn't one": Exit Sub
    If SAP.longLeadNotification = "" Then MsgBox "Missing Long Notification Number. Generate from EAM": Exit Sub
    If SAP.longLeadOrder = "" Then MsgBox "Missing Long Order Number. Generate from EAM": Exit Sub
    If SAP.treeWork = "" Then MsgBox "Missing Tree Work. Put N/A if there isn't one": Exit Sub
    If SAP.numDeniedPoles = "" Then MsgBox "Missing Number of Denied Poles": Exit Sub
    If SAP.stakedLocations = "" Then MsgBox "Missing Staking Locations. Put N/A if there aren't any": Exit Sub
    If SAP.stakedBy = "" Then MsgBox "Missing Staked By. Put TBD if it hasn't been staked yet, or N/A if there isn't any": Exit Sub
    If SAP.stakedDate = "" Then MsgBox "Missing Staked Date. Put TBD if it hasn't been staked yet, or N/A if there isn't any": Exit Sub
    If SAP.num391Poles = "" Then MsgBox "Missing Number of 391 Forms": Exit Sub
    
    Dim data_ws As Worksheet
    Set data_ws = ThisWorkbook.sheets("DATA")
    Dim emailCell As Range
    Set emailCell = data_ws.cells.Find("MKR CE Review Email")
    If Not emailCell Is Nothing Then
        If emailCell.offSet(1, 0).value <> "" Then
            recipient = emailCell.offSet(1, 0).value
        End If
        If emailCell.offSet(2, 0).value <> "" Then
            ccRecipient = emailCell.offSet(2, 0).value & " " & additionalCC & ";"
        End If
    End If
    
    For Each leadEmail In SAP.leadEmails
        If InStr(1, recipient, leadEmail, vbTextCompare) = 0 And InStr(1, ccRecipient, leadEmail, vbTextCompare) = 0 Then
            ccRecipient = ccRecipient & leadEmail
        End If
    Next leadEmail
    
    If InStr(LCase(SAP.treeWork), "n/a") > 0 Or InStr(LCase(SAP.treeWork), "na") > 0 Or InStr(LCase(SAP.treeWork), "no") > 0 Then SAP.treeWork = "No"
    If InStr(LCase(SAP.treeWork), "yes") = 0 And InStr(SAP.treeWork, "no") = 0 Then SAP.treeWork = "Yes" & "<br>" & SAP.treeWork
    
    Call initEmail
    
    Subject = notification & " - " & SAP.constructionMeasure & " - " & SAP.Permit & " - MKR - Ready For Review"
    body = "Important Design Notes:" & "<br><br>" _
        & "Designed by: " & SAP.designerName & Chr(32) & designerNumber & "<br><br>" _
        & "Due Date: " & SAP.dueDate & "<br>" _
        & "# Poles: " & SAP.numPoles & "<br>" _
        & "Applicant: " & SAP.applicant & "<br>" _
        & "Notification: " & SAP.notification & "<br>" _
        & "Permit: " & SAP.Permit & "<br><br>" _
        & "Alt1 Design Document: " & SAP.alt1DesDoc & "<br><br>" _
        & "Alt2 Design Document: " & SAP.alt2DesDoc & "<br><br>" _
        & "Alt3 Design Document: " & SAP.alt3DesDoc & "<br><br>" _
        & "Long Lead Time Material" & "<br>" _
        & "Notification: " & SAP.longLeadNotification & "<br>" _
        & "Order: " & SAP.longLeadOrder & "<br><br>" _
        & "Tree - Trimming: " & Replace(SAP.treeWork, vbLf, "<br>") & "<br><br>" _
        & "# 391 Forms: " & SAP.num391Poles & "<br><br>" _
        & "# of Denied Poles: " & SAP.numDeniedPoles & "<br><br>" _
        & "Design Location Numbers Staked: " & SAP.stakedLocations & "<br>" _
        & "Staked by: " & SAP.stakedBy & "<br>" _
        & "Date Pole Staked: " & SAP.stakedDate & "<br><br>" _
        & "Any questions please feel free to call or e-mail the designer." & "<br>" _
        & signiture

    outlookMail.To = recipient
    outlookMail.cc = ccRecipient
    outlookMail.Subject = Subject
    outlookMail.htmlbody = body
    
    MsgBox "Update smart sheets to CE review."
End Sub

Public Sub EmailMKRDesignComplete()
    Call SendLogMessage("EmailMKRDesignComplete")
    
    Dim SAP As SAP: Set SAP = New SAP
    Call SAP.extractFromSheet
    Dim project As project: Set project = New project
    Call project.extractFromSheets
    
    If SAP.notification = "" Then MsgBox "Missing Notification Number": Exit Sub
    If SAP.constructionMeasure = "" Then MsgBox "Missing Construction Measure Number. Put N/A if there isn't one": Exit Sub
    If SAP.Permit = "" Then MsgBox "Missing Permit": Exit Sub
    If SAP.designerName = "" Then MsgBox "Missing Designer Name": Exit Sub
    If SAP.dueDate = "" Then MsgBox "Missing Due Date": Exit Sub
    If SAP.numPoles = "" Then MsgBox "Missing Number of Poles": Exit Sub
    If SAP.applicant = "" Then MsgBox "Missing Applicant": Exit Sub
    If SAP.alt1DesDoc = "" Then MsgBox "Missing Design Doc Alt1 Number. Put N/A if there isn't one": Exit Sub
    If SAP.alt1Order = "" Then MsgBox "Missing Design Order Alt1 Number. Put N/A if there isn't one": Exit Sub
    If SAP.alt2DesDoc = "" Then MsgBox "Missing Design Doc Alt2 Number. Put N/A if there isn't one": Exit Sub
    If SAP.alt2Order = "" Then MsgBox "Missing Design Order Alt2 Number. Put N/A if there isn't one": Exit Sub
    If SAP.alt3DesDoc = "" Then MsgBox "Missing Design Doc Alt3 Number. Put N/A if there isn't one": Exit Sub
    If SAP.alt3Order = "" Then MsgBox "Missing Design Order Alt3 Number. Put N/A if there isn't one": Exit Sub
    'If sap.longLeadNotification = "" Then MsgBox "Missing Long Notification Number. Generate from EAM. Put N/A if there isn't isn't one": Exit Sub
    'If sap.longLeadOrder = "" Then MsgBox "Missing Long Order Number. Generate from EAM. Put N/A if there isn't one": Exit Sub
    If SAP.treeWork = "" Then MsgBox "Missing Tree Work": Exit Sub
    If SAP.numDeniedPoles = "" Then MsgBox "Missing Number of Denied Poles": Exit Sub
    If SAP.stakedLocations = "" Then MsgBox "Missing Staking Locations. Put N/A if there aren't any": Exit Sub
    If SAP.stakedBy = "" Then MsgBox "Missing Staked By. Put N/A if there isn't any": Exit Sub
    If SAP.stakedDate = "" Then MsgBox "Missing Staked Date. Put N/A if there isn't any": Exit Sub
    If SAP.njunsProject = "" Then MsgBox "Missing NJUNS Project Number. Put N/A if there isn't one": Exit Sub
    If SAP.njunsLocations = "" Then MsgBox "Missing NJUNS Locations. Put N/A if there aren't any": Exit Sub
    If SAP.num391Poles = "" Then MsgBox "Missing Number of 391 Forms": Exit Sub
    If SAP.missdigLocations = "" Then MsgBox "Missing Miss Digs Locations. Put N/A if there aren't any": Exit Sub
    If SAP.raildroadPermitRequired = "" Then MsgBox "Missing Railroad Permit Required. Put N/A if there isn't any": Exit Sub
    If SAP.railroadPermitCost = "" Then MsgBox "Missing Railroad Permit Cost. Put N/A if there isn't any": Exit Sub
    If SAP.permits = "" Then MsgBox "Missing Permits. Put N/A if there aren't any": Exit Sub
    If SAP.permitType = "" Then MsgBox "Missing Permit Types. Put N/A if there aren't any": Exit Sub
    If SAP.permitCost = "" Then MsgBox "Missing Permit Cost. Put N/A if there isn't any": Exit Sub
    If SAP.permitLocations = "" Then MsgBox "Missing Permit Locations. Put N/A if there aren't any": Exit Sub
    If SAP.mdot5Day = "" Then MsgBox "Missing MDOT 5-day Permit Locations. Put N/A if there aren't any": Exit Sub
    If SAP.mdotTree = "" Then MsgBox "Missing MDOT Tree Permit Locations. Put N/A if there aren't any": Exit Sub
    
    Dim data_ws As Worksheet
    Set data_ws = ThisWorkbook.sheets("DATA")
    Dim emailCell As Range
    Set emailCell = data_ws.cells.Find("MKR Design Complete Email")
    If Not emailCell Is Nothing Then
        If emailCell.offSet(1, 0).value <> "" Then
            recipient = emailCell.offSet(1, 0).value & " " & jucEmail & ";"
        End If
        If emailCell.offSet(2, 0).value <> "" Then
            ccRecipient = emailCell.offSet(2, 0).value & " " & additionalCC & ";"
        End If
    End If

    For Each leadEmail In SAP.leadEmails
        If InStr(1, recipient, leadEmail, vbTextCompare) = 0 And InStr(1, ccRecipient, leadEmail, vbTextCompare) = 0 Then
            ccRecipient = ccRecipient & leadEmail
        End If
    Next leadEmail

    additionalCC = Utilities.JoinCollection(SAP.coordinatorEmails, ";")
    
    If InStr(LCase(SAP.treeWork), "n/a") > 0 Or InStr(LCase(SAP.treeWork), "na") > 0 Or InStr(LCase(SAP.treeWork), "no") > 0 Then SAP.treeWork = "No"
    If InStr(LCase(SAP.treeWork), "yes") = 0 And InStr(SAP.treeWork, "no") = 0 Then SAP.treeWork = "Yes" & "<br>" & SAP.treeWork
    
    Dim ceWork As String: ceWork = "is not "
    Dim pole As pole
    Dim streetName As String
    Dim streetNames As Scripting.Dictionary: Set streetNames = New Scripting.Dictionary
    For Each pole In project.poles
        If pole.location <> "" Then ceWork = "is "
        If InStr(pole.address, " ") > 0 And InStr(pole.address, ",") > 0 Then
            streetName = Mid(pole.address, InStr(pole.address, " "), InStr(pole.address, ",") - InStr(pole.address, " "))
            If Not streetNames.exists(streetName) Then streetNames(streetName) = 0
            streetNames(streetName) = streetNames(streetName) + 1
        End If
    Next pole
    Dim max As Integer: max = -1
    streetName = ""
    For Each street In streetNames
        If streetNames(street) > max Then
            streetName = street
            max = streetNames(street)
        End If
    Next street
    
    Call initEmail
    
    Subject = SAP.notification & " - " & SAP.constructionMeasure & " - " & SAP.Permit & " - Design Completed " & Format(Date, "mm/dd/yyyy")
    body = "The design for the " & SAP.numPoles & " pole " & SAP.applicant & " job along " & streetName & " has been completed. There is " & ceWork & "CE work required.<br><br>" _
        & "<b>Job Information:</b><br>" _
        & "Designed by: " & SAP.designerName & Chr(32) & designerNumber & "<br>" _
        & "Notification: " & SAP.notification & "<br>" _
        & "Construction Measure #: " & SAP.constructionMeasure & "<br>" _
        & "Permit: " & SAP.Permit & "<br><br>" _
        & "<b>Alt1</b><br>" _
        & "Design: " & SAP.alt1DesDoc & "<br>" _
        & "Order: " & SAP.alt1Order & "<br>" _
        & "<b>Alt2</b><br>" _
        & "Design: " & SAP.alt2DesDoc & "<br>" _
        & "Order: " & SAP.alt2Order & "<br>" _
        & "<b>Alt3</b><br>" _
        & "Design: " & SAP.alt3DesDoc & "<br>" _
        & "Order: " & SAP.alt3Order & "<br><br>" _
        & "<b>Tree work: </b>" & Replace(SAP.treeWork, vbLf, "<br>") & "<br><br>" _
        & "<b>391 Forms:</b><br>" _
        & "# of Poles: " & SAP.num391Poles & "<br><br>" _
        & "<b>CA/Violations NJUNS Project Number:</b> " & SAP.njunsProject & "<br>" _
        & "# of Locations: " & SAP.njunsLocations & "<br><br>" _
        & "<b>CE Maintenance:</b><br>" _
        & "# of Poles: " & SAP.numMaintPoles & "<br><br>" _
        & "# of Denied Poles: " & SAP.numDeniedPoles & "<br><br>"
        
    body = body & "<b>647 Forms: </b>" & form647 & "<br><br>" _
        & "<b>Comm Make Ready Form included:</b> Yes<br><br>" _
        & "<b>Locations where outages will be required: </b>" & SAP.outageRequired & "<br><br>" _
        & "<b>Permits: " & SAP.permits & "</b><br>" _
        & "<b>Type of Permit: " & SAP.permitType & "</b><br>" _
        & "<b>Permit Cost: " & SAP.permitCost & "</b><br>" _
        & "Design Locations requiring permit: " & SAP.permitLocations & "<br>" _
        & "Design Locations requiring MISS DIG forms: " & SAP.missdigLocations & "<br>" _
        & "Design Locations requiring MDOT 5-day notify forms: " & SAP.mdot5Day & "<br>" _
        & "Design Locations requiring MDOT tree-permit forms: " & SAP.mdotTree & "<br><br>" _
        & "Design Location Numbers Staked: " & SAP.stakedLocations & "<br>" _
        & "<b>Staked by:</b> " & SAP.stakedBy & "<br>" _
        & "Date Pole Staked: " & SAP.stakedDate & "<br><br>" _
        & "<b>Railroad Permit Required:</b> " & SAP.raildroadPermitRequired & "<br>" _
        & "<b>Railroad Permit Cost:</b> " & SAP.railroadPermitCost & "<br><br>" _
        & "Any questions please feel free to call or e-mail the designer.<br>" _
        & signiture

    outlookMail.To = recipient
    outlookMail.cc = ccRecipient
    outlookMail.Subject = Subject
    outlookMail.htmlbody = body
    
    MsgBox "Update smart sheets to submitted."
End Sub

Public Sub EmailNewNotification()
    Dim SAP As SAP: Set SAP = New SAP
    Call SAP.extractFromSheet
    
    Dim ohs As Boolean
    Dim demo As Boolean

    ohs = SAP.serviceNotifications.count > 0
    demo = SAP.demoNotification <> ""
    
    If Not ohs And Not demo Then Exit Sub
    
    Call initEmail
    
    recipient = "jason.haas@cmsenergy.com;"
    ccRecipient = "TRC@cmsenergy.com;van.huynh@cmsenergy.com;"
    
    For Each leadEmail In SAP.leadEmails
        If InStr(1, recipient, leadEmail, vbTextCompare) = 0 And InStr(1, ccRecipient, leadEmail, vbTextCompare) = 0 Then
            ccRecipient = ccRecipient & leadEmail
        End If
    Next leadEmail
    
    Subject = IIf(demo, "Demo", "") & IIf(demo And ohs, "/", "") & IIf(ohs, "OHS", "") & " Notification for " & SAP.notification
    
    yellowHighlight = "<span style='background-color: #FFFF00; padding: 2px 4px; display: inline-block;'>"
    
    body = "Hello,<br><br>"
    body = body & "I created a" & IIf(ohs, "n OHS ", "") & IIf(ohs And demo, "and a", "") & IIf(demo, " DEMO ", "") & " notification<br><br>"
    body = body & "The new notification is highlighted below, with the original above it.<br><br>"
    body = body & SAP.notification & " " & SAP.Permit & "<br>"
    If ohs Then
        For i = 1 To SAP.serviceNotifications.count
            Debug.Print SAP.serviceHouseNumbers.count
            body = body & yellowHighlight & SAP.serviceNotifications(i) & " " & SAP.Permit & " - OHS "
            If SAP.serviceHouseNumbers.count >= i Then
                body = body & SAP.serviceHouseNumbers(i)
            Else
                body = body & "[HOUSE NUMBER]"
            End If
            body = body & "</span><br>"
        Next i
    End If
    If demo Then body = body & yellowHighlight & SAP.demoNotification & " " & SAP.Permit & " - DEMO</span><br>"
    body = body & "<br>This is ready for your review.<br>"
    body = body & signiture

    outlookMail.To = recipient
    outlookMail.cc = ccRecipient
    outlookMail.Subject = Subject
    outlookMail.htmlbody = body
End Sub

Private Sub initEmail()
    Set outlookApp = CreateObject("outlook.Application")
    Set outlookMail = outlookApp.createItem(0)
    
    outlookMail.Display
    signiture = outlookMail.htmlbody
End Sub
