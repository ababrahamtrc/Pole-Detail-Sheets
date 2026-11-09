VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public mode As String
Public jobPurpose As String

Public designerName As String
Public designerNumber As String
Public leadEmails As Collection
Public applicant As String
Public numPoles As String
Public address As String
Public dueDate As String
Public coordinatorName As String
Public coordinatorNumber As String
Public coordinatorEmails As Collection
Public notification As String
Public constructionMeasure As String
Public Permit As String
Public num391Poles As String
Public numMaintPoles As String
Public numDeniedPoles As String
Public forms647 As String
Public permits As String
Public permitType As String
Public permitCost As String
Public permitLocations As String
Public raildroadPermitRequired As String
Public railroadPermitCost As String
Public missdigLocations As String
Public mdot5Day As String
Public mdotTree As String
Public treeWork As String
Public njunsProject As String
Public njunsLocations As String
Public stakedLocations As String
Public stakedBy As String
Public stakedDate As String
Public erc As String
Public plia As String
Public outageRequired As String
Public ttc As String
Public trs As String
Public systemOwner As String
Public substation As String
Public circuit As String
Public voltage As String
Public taxCode As String
Public hdq As String
Public county As String
Public twp As String
Public functionalLoc As String
Public tlm As String
Public upstreamDevices As String
Public longLeadNotification As String
Public longLeadOrder As String
Public hotLineTagging As String

'MKR
Public alt1DesDoc As String
Public alt1Order As String
Public alt2DesDoc As String
Public alt2Order As String
Public alt3DesDoc As String
Public alt3Order As String

'SI
Public totalCost As String
Public totalHours As String
Public mainDesDoc As String
Public mainOrder As String
Public mainCost As String
Public mainHours As String
Public demoNotification As String
Public demoLongLeadNotification As String
Public demoLongLeadOrder As String
Public demoConstructionMeasure As String
Public demoDesDoc As String
Public demoOrder As String
Public demoCost As String
Public demoHours As String
Public serviceNotifications As Collection
Public serviceHouseNumbers As Collection
Public serviceDesDocs As Collection
Public serviceOrders As Collection
Public serviceLongLeadNotifications As Collection
Public serviceLongLeadOrders As Collection
Public serviceCosts As Collection
Public serviceHours As Collection
Public docWarehouseLink As String

Private Sub Class_Initialize()
    Me.mode = ThisWorkbook.sheets("Control").Range("MODE").text
    Me.designerName = ThisWorkbook.sheets("Control").Range("NAME")
    Me.designerNumber = ThisWorkbook.sheets("Control").Range("PHONE")
    
    Set Me.leadEmails = New Collection
    parts = Split(ThisWorkbook.sheets("Control").Range("CCEMAILS"), ",")
    For Each part In parts
        Me.leadEmails.Add Replace(Trim$(part), ";", "") & ";"
    Next part
    
    Set Me.coordinatorEmails = New Collection
    
    Set serviceNotifications = New Collection
    Set serviceHouseNumbers = New Collection
    Set serviceDesDocs = New Collection
    Set serviceLongLeadNotifications = New Collection
    Set serviceLongLeadOrders = New Collection
    Set serviceOrders = New Collection
    Set serviceCosts = New Collection
    Set serviceHours = New Collection
End Sub

Public Sub extractFromSheet()
    Dim sapWs As Worksheet
    If Me.mode = "MAKE READY" Then
        Set sapWs = ThisWorkbook.sheets("MKR SAP")
    ElseIf Me.mode = "SYSTEM IMPROVEMENT" Then
        Set sapWs = ThisWorkbook.sheets("SI SAP")
    Else
        Exit Sub
    End If
    
    Me.numPoles = sapWs.cells.Find("# of Poles:", LookAt:=xlWhole).offSet(0, 1).value
    Me.address = sapWs.cells.Find("Address:", LookAt:=xlWhole).offSet(0, 1).value
    Me.dueDate = sapWs.cells.Find("Due Date:", LookAt:=xlWhole).offSet(0, 1).value
     
    Set cell = sapWs.cells.Find("JUC/PC:", LookAt:=xlWhole)
    If Not cell Is Nothing Then
        Me.coordinatorName = cell.offSet(0, 1).value
        Me.coordinatorNumber = cell.offSet(0, 1).offSet(0, 1).value
        
        For i = 0 To 2
            If cell.offSet(i, 0).offSet(0, 1).offSet(0, 1).offSet(0, 1).value = "" Then Exit For
            Me.coordinatorEmails.Add cell.offSet(i, 0).offSet(i, 3).value
        Next i
    End If
    
    Me.notification = sapWs.cells.Find("Notification:", LookAt:=xlWhole).offSet(0, 1).value
    Me.constructionMeasure = sapWs.cells.Find("Construction Measure:", LookAt:=xlWhole).offSet(0, 1).value
    Me.Permit = sapWs.cells.Find("Permit:", LookAt:=xlWhole).offSet(0, 1).value
    Me.num391Poles = sapWs.cells.Find("# of Form 391 Poles:", LookAt:=xlWhole).offSet(0, 1).value
    Me.numMaintPoles = sapWs.cells.Find("# of Ce Maintenance Poles:", LookAt:=xlWhole).offSet(0, 1).value
    Me.numDeniedPoles = sapWs.cells.Find("# of Denied Poles:", LookAt:=xlWhole).offSet(0, 1).value
    Me.forms647 = sapWs.cells.Find("647 Forms:", LookAt:=xlWhole).offSet(0, 1).value
    
    Me.permits = sapWs.cells.Find("Permits:", LookAt:=xlWhole).offSet(0, 1).value
    Me.permitType = sapWs.cells.Find("Type of Permit:", LookAt:=xlWhole).offSet(0, 1).value
    Me.permitCost = sapWs.cells.Find("Permit Cost:", LookAt:=xlWhole).offSet(0, 1).value
    Me.permitLocations = sapWs.cells.Find("Design Locations requiring permit:", LookAt:=xlWhole).offSet(0, 1).value
    Me.missdigLocations = sapWs.cells.Find("Design Locations requiring MISS DIG forms:", LookAt:=xlWhole).offSet(0, 1).value
    Me.mdot5Day = sapWs.cells.Find("Design Locations requiring MDOT 5-day notify forms:", LookAt:=xlWhole).offSet(0, 1).value
    Me.mdotTree = sapWs.cells.Find("Design Locations requiring MDOT tree-permit forms:", LookAt:=xlWhole).offSet(0, 1).value
    
    Me.raildroadPermitRequired = sapWs.cells.Find("Railroad Permit Required:", LookAt:=xlWhole).offSet(0, 1).value
    Me.railroadPermitCost = sapWs.cells.Find("Railroad Permit Cost:", LookAt:=xlWhole).offSet(0, 1).value
    Me.treeWork = sapWs.cells.Find("Tree Work:", LookAt:=xlWhole).offSet(0, 1).value
    Me.njunsProject = sapWs.cells.Find("NJUNS Project Number:", LookAt:=xlWhole).offSet(0, 1).value
    Me.njunsLocations = sapWs.cells.Find("NJUNS Locations:", LookAt:=xlWhole).offSet(0, 1).value
    Me.stakedLocations = sapWs.cells.Find("Design Location Numbers Staked:", LookAt:=xlWhole).offSet(0, 1).value
    Me.stakedBy = sapWs.cells.Find("Staked by:", LookAt:=xlWhole).offSet(0, 1).value
    Me.stakedDate = sapWs.cells.Find("Date Pole Staked:", LookAt:=xlWhole).offSet(0, 1).value
    Me.erc = sapWs.cells.Find("ERC Checklist#:", LookAt:=xlWhole).offSet(0, 1).value
    Me.plia = sapWs.cells.Find("PLIA:", LookAt:=xlWhole).offSet(0, 1).value
    Me.outageRequired = sapWs.cells.Find("Outage Required:", LookAt:=xlWhole).offSet(0, 1).value
    Me.ttc = sapWs.cells.Find("TTC:", LookAt:=xlWhole).offSet(0, 1).value
    Me.trs = sapWs.cells.Find("TRS:", LookAt:=xlWhole).offSet(0, 1).value
    Me.systemOwner = sapWs.cells.Find("System Owner:", LookAt:=xlWhole).offSet(0, 1).value
    Me.substation = sapWs.cells.Find("Substation:", LookAt:=xlWhole).offSet(0, 1).value
    Me.circuit = sapWs.cells.Find("Circuit:", LookAt:=xlWhole).offSet(0, 1).value
    Me.voltage = sapWs.cells.Find("Voltage:", LookAt:=xlWhole).offSet(0, 1).value
    Me.taxCode = sapWs.cells.Find("Taxcode:", LookAt:=xlWhole).offSet(0, 1).value
    Me.hdq = sapWs.cells.Find("HDQ:", LookAt:=xlWhole).offSet(0, 1).value
    Me.county = sapWs.cells.Find("County:", LookAt:=xlWhole).offSet(0, 1).value
    Me.twp = sapWs.cells.Find("TWP:", LookAt:=xlWhole).offSet(0, 1).value
    Me.functionalLoc = sapWs.cells.Find("Functional Loc:", LookAt:=xlWhole).offSet(0, 1).value
    Me.tlm = sapWs.cells.Find("TLM:", LookAt:=xlWhole).offSet(0, 1).value
    Me.upstreamDevices = sapWs.cells.Find("Upstream Device(s):", LookAt:=xlWhole).offSet(0, 1).value
    Me.hotLineTagging = sapWs.cells.Find("Hot Line Tagging Locations:", LookAt:=xlWhole).offSet(0, 1).value
    
    If Me.mode = "MAKE READY" Then
        Me.applicant = sapWs.cells.Find("Applicant:", LookAt:=xlWhole).offSet(0, 1).value
        Me.alt1DesDoc = sapWs.cells.Find("ALT1 Design Document:", LookAt:=xlWhole).offSet(0, 1).value
        Me.alt1Order = sapWs.cells.Find("ALT1 Order:", LookAt:=xlWhole).offSet(0, 1).value
        Me.alt2DesDoc = sapWs.cells.Find("ALT2 Design Document:", LookAt:=xlWhole).offSet(0, 1).value
        Me.alt2Order = sapWs.cells.Find("ALT2 Order:", LookAt:=xlWhole).offSet(0, 1).value
        Me.alt3DesDoc = sapWs.cells.Find("ALT3 Design Document:", LookAt:=xlWhole).offSet(0, 1).value
        Me.alt3Order = sapWs.cells.Find("ALT3 Order:", LookAt:=xlWhole).offSet(0, 1).value
        Me.longLeadNotification = sapWs.cells.Find("Long Lead Notification:", LookAt:=xlWhole).offSet(0, 1).value
        Me.longLeadOrder = sapWs.cells.Find("Long Lead Order:", LookAt:=xlWhole).offSet(0, 1).value
        Me.jobPurpose = "MAKE READY FOR " & UCase(Me.applicant) & vbLf & Me.notification & vbLf & Me.Permit
    ElseIf Me.mode = "SYSTEM IMPROVEMENT" Then
        Me.applicant = "LVDE"
        Me.totalCost = Format(sapWs.cells.Find("Total Cost:", LookAt:=xlWhole).offSet(0, 1).value, "$#,##0.00")
        Me.totalHours = sapWs.cells.Find("Total Planning Hours:", LookAt:=xlWhole).offSet(0, 1).value
        Me.mainDesDoc = sapWs.cells.Find("MAIN Design Document:", LookAt:=xlWhole).offSet(0, 1).value
        Me.mainOrder = sapWs.cells.Find("MAIN Order:", LookAt:=xlWhole).offSet(0, 1).value
        Me.longLeadNotification = sapWs.cells.Find("MAIN LLTM Notification:", LookAt:=xlWhole).offSet(0, 1).value
        Me.longLeadOrder = sapWs.cells.Find("MAIN LLTM Order:", LookAt:=xlWhole).offSet(0, 1).value
        Me.mainCost = Format(sapWs.cells.Find("MAIN Cost:", LookAt:=xlWhole).offSet(0, 1).value, "$#,##0.00")
        Me.mainHours = sapWs.cells.Find("MAIN Planning Hours:", LookAt:=xlWhole).offSet(0, 1).value
        Me.demoNotification = sapWs.cells.Find("DEMO Notification:", LookAt:=xlWhole).offSet(0, 1).value
        Me.demoLongLeadNotification = sapWs.cells.Find("DEMO LLTM Notification:", LookAt:=xlWhole).offSet(0, 1).value
        Me.demoLongLeadOrder = sapWs.cells.Find("DEMO LLTM Order:", LookAt:=xlWhole).offSet(0, 1).value
        Me.demoConstructionMeasure = sapWs.cells.Find("DEMO Construction Measure:", LookAt:=xlWhole).offSet(0, 1).value
        Me.demoDesDoc = sapWs.cells.Find("DEMO Design Document:", LookAt:=xlWhole).offSet(0, 1).value
        Me.demoOrder = sapWs.cells.Find("DEMO Order:", LookAt:=xlWhole).offSet(0, 1).value
        Me.demoCost = Format(sapWs.cells.Find("DEMO Cost:", LookAt:=xlWhole).offSet(0, 1).value, "$#,##0.00")
        Me.demoHours = sapWs.cells.Find("DEMO Planning Hours:", LookAt:=xlWhole).offSet(0, 1).value
        Me.docWarehouseLink = sapWs.cells.Find("Doc Warehouse Link:", LookAt:=xlWhole).offSet(0, 1).value
        Me.jobPurpose = sapWs.cells.Find("Job Purpose:", LookAt:=xlWhole).offSet(0, 1).value
        
        parts = Split(sapWs.cells.Find("SERVICE Notification(s):", LookAt:=xlWhole).offSet(0, 1).value, ",")
        For Each item In parts
            If item <> "N/A" Then Me.serviceNotifications.Add Trim(item)
        Next item
    
        parts = Split(sapWs.cells.Find("SERVICE House Number(s):", LookAt:=xlWhole).offSet(0, 1).value, ",")
        For Each item In parts
            If item <> "N/A" Then Me.serviceHouseNumbers.Add Trim(item)
        Next item
    
        parts = Split(sapWs.cells.Find("SERVICE Design Document(s):", LookAt:=xlWhole).offSet(0, 1).value, ",")
        For Each item In parts
            If item <> "N/A" Then Me.serviceDesDocs.Add Trim(item)
        Next item

        parts = Split(sapWs.cells.Find("SERVICE Order(s):", LookAt:=xlWhole).offSet(0, 1).value, ",")
        For Each item In parts
            If item <> "N/A" Then Me.serviceOrders.Add Trim(item)
        Next item

        parts = Split(sapWs.cells.Find("SERVICE LLTM Notification(s):", LookAt:=xlWhole).offSet(0, 1).value, ",")
        For Each item In parts
            If item <> "N/A" Then Me.serviceLongLeadNotifications.Add Trim(item)
        Next item
        
        parts = Split(sapWs.cells.Find("SERVICE LLTM Order(s):", LookAt:=xlWhole).offSet(0, 1).value, ",")
        For Each item In parts
            If item <> "N/A" Then Me.serviceLongLeadNotifications.Add Trim(item)
        Next item

        parts = Split(sapWs.cells.Find("SERVICE Cost(s):", LookAt:=xlWhole).offSet(0, 1).value, ",")
        For Each item In parts
            If item <> "N/A" Then Me.serviceCosts.Add Trim(item)
        Next item

        parts = Split(sapWs.cells.Find("SERVICE Planning Hours:", LookAt:=xlWhole).offSet(0, 1).value, ",")
        For Each item In parts
            If item <> "N/A" Then Me.serviceHours.Add Trim(item)
        Next item
    End If
End Sub

Public Sub fillSheet()
    Dim sapWs As Worksheet
    If Me.mode = "MAKE READY" Then
        Set sapWs = ThisWorkbook.sheets("MKR SAP")
    ElseIf mode = "SYSTEM IMPROVEMENT" Then
        Set sapWs = ThisWorkbook.sheets("SI SAP")
    Else
        Exit Sub
    End If
    
    If Me.numPoles <> "" Then sapWs.cells.Find("# of Poles:", LookAt:=xlWhole).offSet(0, 1).value = Me.numPoles
    If Me.address <> "" Then sapWs.cells.Find("Address:", LookAt:=xlWhole).offSet(0, 1).value = Me.address
    If Me.dueDate <> "" Then sapWs.cells.Find("Due Date:", LookAt:=xlWhole).offSet(0, 1).value = Me.dueDate
    
    If Me.notification <> "" Then sapWs.cells.Find("Notification:", LookAt:=xlWhole).offSet(0, 1).value = Me.notification
    If Me.constructionMeasure <> "" Then sapWs.cells.Find("Construction Measure:", LookAt:=xlWhole).offSet(0, 1).value = Me.constructionMeasure
    If Me.Permit <> "" Then sapWs.cells.Find("Permit:", LookAt:=xlWhole).offSet(0, 1).value = Me.Permit
    If Me.num391Poles <> "" Then sapWs.cells.Find("# of Form 391 Poles:", LookAt:=xlWhole).offSet(0, 1).value = Me.num391Poles
    If Me.numMaintPoles <> "" Then sapWs.cells.Find("# of Ce Maintenance Poles:", LookAt:=xlWhole).offSet(0, 1).value = Me.numMaintPoles
    If Me.numDeniedPoles <> "" Then sapWs.cells.Find("# of Denied Poles:", LookAt:=xlWhole).offSet(0, 1).value = Me.numDeniedPoles
    If Me.forms647 <> "" Then sapWs.cells.Find("647 Forms:", LookAt:=xlWhole).offSet(0, 1).value = Me.forms647
    
    If Me.permits <> "" Then sapWs.cells.Find("Permits:", LookAt:=xlWhole).offSet(0, 1).value = Me.permits
    If Me.permitType <> "" Then sapWs.cells.Find("Type of Permit:", LookAt:=xlWhole).offSet(0, 1).value = Me.permitType
    If Me.permitCost <> "" Then sapWs.cells.Find("Permit Cost:", LookAt:=xlWhole).offSet(0, 1).value = Me.permitCost
    If Me.permitLocations <> "" Then sapWs.cells.Find("Design Locations requiring permit:", LookAt:=xlWhole).offSet(0, 1).value = Me.permitLocations
    sapWs.cells.Find("Design Locations requiring MISS DIG forms:", LookAt:=xlWhole).offSet(0, 1).NumberFormat = "@"
    If Me.missdigLocations <> "" Then sapWs.cells.Find("Design Locations requiring MISS DIG forms:", LookAt:=xlWhole).offSet(0, 1).value = Me.missdigLocations
    If Me.mdot5Day <> "" Then sapWs.cells.Find("Design Locations requiring MDOT 5-day notify forms:", LookAt:=xlWhole).offSet(0, 1).value = Me.mdot5Day
    If Me.mdotTree <> "" Then sapWs.cells.Find("Design Locations requiring MDOT tree-permit forms:", LookAt:=xlWhole).offSet(0, 1).value = Me.mdotTree
    
    If Me.raildroadPermitRequired <> "" Then sapWs.cells.Find("Railroad Permit Required:", LookAt:=xlWhole).offSet(0, 1).value = Me.raildroadPermitRequired
    If Me.railroadPermitCost <> "" Then sapWs.cells.Find("Railroad Permit Cost:", LookAt:=xlWhole).offSet(0, 1).value = Me.railroadPermitCost
    If Me.treeWork <> "" Then sapWs.cells.Find("Tree Work:", LookAt:=xlWhole).offSet(0, 1).value = Me.treeWork
    If Me.njunsProject <> "" Then sapWs.cells.Find("NJUNS Project Number:", LookAt:=xlWhole).offSet(0, 1).value = Me.njunsProject
    If Me.njunsLocations <> "" Then sapWs.cells.Find("NJUNS Locations:", LookAt:=xlWhole).offSet(0, 1).value = Me.njunsLocations
    sapWs.cells.Find("Design Location Numbers Staked:", LookAt:=xlWhole).offSet(0, 1).NumberFormat = "@"
    If Me.stakedLocations <> "" Then sapWs.cells.Find("Design Location Numbers Staked:", LookAt:=xlWhole).offSet(0, 1).value = Me.stakedLocations
    If Me.stakedBy <> "" Then sapWs.cells.Find("Staked by:", LookAt:=xlWhole).offSet(0, 1).value = Me.stakedBy
    If Me.stakedDate <> "" Then sapWs.cells.Find("Date Pole Staked:", LookAt:=xlWhole).offSet(0, 1).value = Me.stakedDate
    If Me.erc <> "" Then sapWs.cells.Find("ERC Checklist#:", LookAt:=xlWhole).offSet(0, 1).value = Me.erc
    If Me.plia <> "" Then sapWs.cells.Find("PLIA:", LookAt:=xlWhole).offSet(0, 1).value = Me.plia
    If Me.outageRequired <> "" Then sapWs.cells.Find("Outage Required:", LookAt:=xlWhole).offSet(0, 1).value = Me.outageRequired
    If Me.ttc <> "" Then sapWs.cells.Find("TTC:", LookAt:=xlWhole).offSet(0, 1).value = Me.ttc
    If Me.trs <> "" Then sapWs.cells.Find("TRS:", LookAt:=xlWhole).offSet(0, 1).value = Me.trs
    If Me.systemOwner <> "" Then sapWs.cells.Find("System Owner:").offSet(0, 1).value = Me.systemOwner
    If Me.substation <> "" Then sapWs.cells.Find("Substation:", LookAt:=xlWhole).offSet(0, 1).value = Me.substation
    If Me.circuit <> "" Then sapWs.cells.Find("Circuit:", LookAt:=xlWhole).offSet(0, 1).value = Me.circuit
    If Me.voltage <> "" Then sapWs.cells.Find("Voltage:", LookAt:=xlWhole).offSet(0, 1).value = Me.voltage
    If Me.taxCode <> "" Then sapWs.cells.Find("Taxcode:", LookAt:=xlWhole).offSet(0, 1).value = Me.taxCode
    If Me.hdq <> "" Then sapWs.cells.Find("HDQ:", LookAt:=xlWhole).offSet(0, 1).value = Me.hdq
    If Me.county <> "" Then sapWs.cells.Find("County:", LookAt:=xlWhole).offSet(0, 1).value = Me.county
    If Me.twp <> "" Then sapWs.cells.Find("TWP:", LookAt:=xlWhole).offSet(0, 1).value = Me.twp
    If Me.functionalLoc <> "" Then sapWs.cells.Find("Functional Loc:", LookAt:=xlWhole).offSet(0, 1).value = Me.functionalLoc
    If Me.tlm <> "" Then sapWs.cells.Find("TLM:", LookAt:=xlWhole).offSet(0, 1).value = Me.tlm
    If Me.upstreamDevices <> "" Then sapWs.cells.Find("Upstream Device(s):", LookAt:=xlWhole).offSet(0, 1).value = Me.upstreamDevices
    If Me.hotLineTagging <> "" Then sapWs.cells.Find("Hot Line Tagging Locations:", LookAt:=xlWhole).offSet(0, 1).value = Me.hotLineTagging
    
    If Me.mode = "MAKE READY" Then
        Set cell = sapWs.cells.Find("JUC/PC:", LookAt:=xlWhole)
        If Not cell Is Nothing Then
            If Me.coordinatorName <> "" Then cell.offSet(0, 1).value = Me.coordinatorName
            If Me.coordinatorNumber <> "" Then cell.offSet(0, 1).offSet(0, 1).value = Me.coordinatorNumber
            For i = 0 To Me.coordinatorEmails.count - 1
                cell.offSet(i, 0).offSet(0, 1).offSet(0, 1).offSet(0, 1).value = coordinatorEmails(i + 1)
            Next i
        End If
    
        If Me.applicant <> "" Then sapWs.cells.Find("Applicant:", LookAt:=xlWhole).offSet(0, 1).value = Me.applicant
        If Me.alt1DesDoc <> "" Then sapWs.cells.Find("ALT1 Design Document:", LookAt:=xlWhole).offSet(0, 1).value = Me.alt1DesDoc
        If Me.alt1Order <> "" Then sapWs.cells.Find("ALT1 Order:", LookAt:=xlWhole).offSet(0, 1).value = Me.alt1Order
        If Me.alt2DesDoc <> "" Then sapWs.cells.Find("ALT2 Design Document:", LookAt:=xlWhole).offSet(0, 1).value = Me.alt2DesDoc
        If Me.alt2Order <> "" Then sapWs.cells.Find("ALT2 Order:", LookAt:=xlWhole).offSet(0, 1).value = Me.alt2Order
        If Me.alt3DesDoc <> "" Then sapWs.cells.Find("ALT3 Design Document:", LookAt:=xlWhole).offSet(0, 1).value = Me.alt3DesDoc
        If Me.alt3Order <> "" Then sapWs.cells.Find("ALT3 Order:", LookAt:=xlWhole).offSet(0, 1).value = Me.alt3Order
        If Me.longLeadNotification <> "" Then sapWs.cells.Find("Long Lead Notification:", LookAt:=xlWhole).offSet(0, 1).value = Me.longLeadNotification
        If Me.longLeadOrder <> "" Then sapWs.cells.Find("Long Lead Order:", LookAt:=xlWhole).offSet(0, 1).value = Me.longLeadOrder
    ElseIf Me.mode = "SYSTEM IMPROVEMENT" Then
        If Me.mainDesDoc <> "" Then sapWs.cells.Find("MAIN Design Document:", LookAt:=xlWhole).offSet(0, 1).value = Me.mainDesDoc
        If Me.mainOrder <> "" Then sapWs.cells.Find("MAIN Order:", LookAt:=xlWhole).offSet(0, 1).value = Me.mainOrder
        If Me.longLeadNotification <> "" Then sapWs.cells.Find("MAIN LLTM Notification:", LookAt:=xlWhole).offSet(0, 1).value = Me.longLeadNotification
        If Me.longLeadOrder <> "" Then sapWs.cells.Find("MAIN LLTM Order:", LookAt:=xlWhole).offSet(0, 1).value = Me.longLeadOrder
        If Me.mainCost <> "" Then sapWs.cells.Find("MAIN Cost:", LookAt:=xlWhole).offSet(0, 1).value = Me.mainCost
        If Me.mainHours <> "" Then sapWs.cells.Find("MAIN Planning Hours:", LookAt:=xlWhole).offSet(0, 1).value = Me.mainHours
        If Me.demoNotification <> "" Then sapWs.cells.Find("DEMO Notification:", LookAt:=xlWhole).offSet(0, 1).value = Me.demoNotification
        If Me.demoLongLeadNotification <> "" Then sapWs.cells.Find("DEMO LLTM Notification:", LookAt:=xlWhole).offSet(0, 1).value = Me.demoLongLeadNotification
        If Me.demoLongLeadOrder <> "" Then sapWs.cells.Find("DEMO LLTM Order:", LookAt:=xlWhole).offSet(0, 1).value = Me.demoLongLeadOrder
        If Me.demoConstructionMeasure <> "" Then sapWs.cells.Find("DEMO Construction Measure:", LookAt:=xlWhole).offSet(0, 1).value = Me.demoConstructionMeasure
        If Me.demoDesDoc <> "" Then sapWs.cells.Find("DEMO Design Document:", LookAt:=xlWhole).offSet(0, 1).value = Me.demoDesDoc
        If Me.demoOrder <> "" Then sapWs.cells.Find("DEMO Order:", LookAt:=xlWhole).offSet(0, 1).value = Me.demoOrder
        If Me.demoCost <> "" Then sapWs.cells.Find("DEMO Cost:", LookAt:=xlWhole).offSet(0, 1).value = Me.demoCost
        If Me.demoHours <> "" Then sapWs.cells.Find("DEMO Planning Hours:", LookAt:=xlWhole).offSet(0, 1).value = Me.demoHours
        If Me.docWarehouseLink <> "" Then
            If Me.docWarehouseLink <> "Link Not Found" Then
                Set cell = sapWs.cells.Find("Doc Warehouse Link:", LookAt:=xlWhole).offSet(0, 1)
                cell.value = Me.docWarehouseLink
                'cell.Parent.Hyperlinks.Add
                    'anchor:=cell, _
                    'address:=Me.docWarehouseLink
            Else
                sapWs.cells.Find("Doc Warehouse Link:", LookAt:=xlWhole).offSet(0, 1).value = Me.docWarehouseLink
            End If
        End If
        If Me.jobPurpose <> "" Then sapWs.cells.Find("Job Purpose:", LookAt:=xlWhole).offSet(0, 1).value = Me.jobPurpose
        
        If serviceNotifications.count > 0 Then sapWs.cells.Find("SERVICE Notification(s):", LookAt:=xlWhole).offSet(0, 1).value = Utilities.JoinCollection(serviceNotifications, ", ")
        If serviceHouseNumbers.count > 0 Then sapWs.cells.Find("SERVICE House Number(s):", LookAt:=xlWhole).offSet(0, 1).value = Utilities.JoinCollection(serviceHouseNumbers, ", ")
        If serviceDesDocs.count > 0 Then sapWs.cells.Find("SERVICE Design Document(s):", LookAt:=xlWhole).offSet(0, 1).value = Utilities.JoinCollection(serviceDesDocs, ", ")
        If serviceOrders.count > 0 Then sapWs.cells.Find("SERVICE Order(s):", LookAt:=xlWhole).offSet(0, 1).value = Utilities.JoinCollection(serviceOrders, ", ")
        If serviceLongLeadNotifications.count > 0 Then sapWs.cells.Find("SERVICE LLTM Notification(s):", LookAt:=xlWhole).offSet(0, 1).value = Utilities.JoinCollection(serviceLongLeadNotifications, ", ")
        If serviceLongLeadNotifications.count > 0 Then sapWs.cells.Find("SERVICE LLTM Order(s):", LookAt:=xlWhole).offSet(0, 1).value = Utilities.JoinCollection(serviceLongLeadNotifications, ", ")
        If serviceCosts.count > 0 Then sapWs.cells.Find("SERVICE Cost(s):", LookAt:=xlWhole).offSet(0, 1).value = Utilities.JoinCollection(serviceCosts, ", ")
        If serviceHours.count > 0 Then sapWs.cells.Find("SERVICE Planning Hours:", LookAt:=xlWhole).offSet(0, 1).value = Utilities.JoinCollection(serviceHours, ", ")
    
    End If
End Sub

Public Sub clearSheets(Optional sheet As Worksheet)
    Dim sapWorksheets As Collection: Set sapWorksheets = New Collection
    If sheet Is Nothing Then
        sapWorksheets.Add ThisWorkbook.sheets("MKR SAP")
        sapWorksheets.Add ThisWorkbook.sheets("SI SAP")
    Else
        sapWorksheets.Add sheet
    End If
    
    Dim sapWs As Worksheet
    For Each sapWs In sapWorksheets
        sapWs.cells.Find("# of Poles:").offSet(0, 1).value = ""
        sapWs.cells.Find("Address:").offSet(0, 1).value = ""
        sapWs.cells.Find("Due Date:").offSet(0, 1).value = ""
        
        For i = 0 To 3
            Set cell = sapWs.cells.Find("JUC/PC:").offSet(i, 0)
            For j = 0 To 2
                Set cell = cell.offSet(0, 1)
                cell.offSet.value = ""
            Next j
        Next i
        
        sapWs.cells.Find("Notification:").offSet(0, 1).value = ""
        sapWs.cells.Find("Construction Measure:").offSet(0, 1).value = ""
        sapWs.cells.Find("Permit:").offSet(0, 1).value = ""
        sapWs.cells.Find("# of Form 391 Poles:").offSet(0, 1).value = ""
        sapWs.cells.Find("# of Ce Maintenance Poles:").offSet(0, 1).value = ""
        sapWs.cells.Find("# of Denied Poles:").offSet(0, 1).value = ""
        sapWs.cells.Find("647 Forms:").offSet(0, 1).value = ""
        
        sapWs.cells.Find("Permits:").offSet(0, 1).value = ""
        sapWs.cells.Find("Type of Permit:").offSet(0, 1).value = ""
        sapWs.cells.Find("Permit Cost:").offSet(0, 1).value = ""
        sapWs.cells.Find("Design Locations requiring permit:").offSet(0, 1).value = ""
        sapWs.cells.Find("Design Locations requiring MISS DIG forms:").offSet(0, 1).value = ""
        sapWs.cells.Find("Design Locations requiring MDOT 5-day notify forms:").offSet(0, 1).value = ""
        sapWs.cells.Find("Design Locations requiring MDOT tree-permit forms:").offSet(0, 1).value = ""
        
        sapWs.cells.Find("Railroad Permit Required:").offSet(0, 1).value = ""
        sapWs.cells.Find("Railroad Permit Cost:").offSet(0, 1).value = ""
        sapWs.cells.Find("Tree Work:").offSet(0, 1).value = ""
        sapWs.cells.Find("NJUNS Project Number:").offSet(0, 1).value = ""
        sapWs.cells.Find("NJUNS Locations:").offSet(0, 1).value = ""
        sapWs.cells.Find("Design Location Numbers Staked:").offSet(0, 1).value = ""
        sapWs.cells.Find("Staked by:").offSet(0, 1).value = ""
        sapWs.cells.Find("Date Pole Staked:").offSet(0, 1).value = ""
        sapWs.cells.Find("ERC Checklist#:").offSet(0, 1).value = ""
        sapWs.cells.Find("PLIA:").offSet(0, 1).value = ""
        sapWs.cells.Find("Outage Required:").offSet(0, 1).value = ""
        sapWs.cells.Find("TTC:").offSet(0, 1).value = ""
        sapWs.cells.Find("TRS:").offSet(0, 1).value = ""
        sapWs.cells.Find("System Owner:").offSet(0, 1).value = ""
        sapWs.cells.Find("Substation:").offSet(0, 1).value = ""
        sapWs.cells.Find("Circuit:").offSet(0, 1).value = ""
        sapWs.cells.Find("Voltage:").offSet(0, 1).value = ""
        sapWs.cells.Find("Taxcode:").offSet(0, 1).value = ""
        sapWs.cells.Find("HDQ:").offSet(0, 1).value = ""
        sapWs.cells.Find("County:").offSet(0, 1).value = ""
        sapWs.cells.Find("TWP:").offSet(0, 1).value = ""
        sapWs.cells.Find("Functional Loc:").offSet(0, 1).value = ""
        sapWs.cells.Find("TLM:").offSet(0, 1).value = ""
        sapWs.cells.Find("Upstream Device(s):").offSet(0, 1).value = ""
        sapWs.cells.Find("Hot Line Tagging Locations:").offSet(0, 1).value = ""
        
        If sapWs.name = "MKR SAP" Then
            sapWs.cells.Find("Applicant:").offSet(0, 1).value = ""
            sapWs.cells.Find("ALT1 Design Document:").offSet(0, 1).value = ""
            sapWs.cells.Find("ALT1 Order:").offSet(0, 1).value = ""
            sapWs.cells.Find("ALT2 Design Document:").offSet(0, 1).value = ""
            sapWs.cells.Find("ALT2 Order:").offSet(0, 1).value = ""
            sapWs.cells.Find("ALT3 Design Document:").offSet(0, 1).value = ""
            sapWs.cells.Find("ALT3 Order:").offSet(0, 1).value = ""
            sapWs.cells.Find("Long Lead Notification:").offSet(0, 1).value = ""
            sapWs.cells.Find("Long Lead Order:").offSet(0, 1).value = ""
        ElseIf sapWs.name = "SI SAP" Then
            sapWs.cells.Find("MAIN Design Document:").offSet(0, 1).value = ""
            sapWs.cells.Find("MAIN Order:").offSet(0, 1).value = ""
            sapWs.cells.Find("MAIN LLTM Notification:").offSet(0, 1).value = ""
            sapWs.cells.Find("MAIN LLTM Order:").offSet(0, 1).value = ""
            sapWs.cells.Find("MAIN Cost:").offSet(0, 1).value = ""
            sapWs.cells.Find("MAIN Planning Hours:").offSet(0, 1).value = ""
            sapWs.cells.Find("SERVICE Notification(s):").offSet(0, 1).value = ""
            sapWs.cells.Find("SERVICE House Number(s):").offSet(0, 1).value = ""
            sapWs.cells.Find("SERVICE Design Document(s):").offSet(0, 1).value = ""
            sapWs.cells.Find("SERVICE Order(s):").offSet(0, 1).value = ""
            sapWs.cells.Find("SERVICE LLTM Notification(s):").offSet(0, 1).value = ""
            sapWs.cells.Find("SERVICE LLTM Order(s):").offSet(0, 1).value = ""
            sapWs.cells.Find("SERVICE Cost(s):").offSet(0, 1).value = ""
            sapWs.cells.Find("SERVICE Planning Hours:").offSet(0, 1).value = ""
            sapWs.cells.Find("DEMO Notification:").offSet(0, 1).value = ""
            sapWs.cells.Find("DEMO LLTM Notification:").offSet(0, 1).value = ""
            sapWs.cells.Find("DEMO LLTM Order:").offSet(0, 1).value = ""
            sapWs.cells.Find("DEMO Construction Measure:").offSet(0, 1).value = ""
            sapWs.cells.Find("DEMO Design Document:").offSet(0, 1).value = ""
            sapWs.cells.Find("DEMO Order:").offSet(0, 1).value = ""
            sapWs.cells.Find("DEMO Cost:").offSet(0, 1).value = ""
            sapWs.cells.Find("DEMO Planning Hours:").offSet(0, 1).value = ""
            sapWs.cells.Find("Doc Warehouse Link:").offSet(0, 1).value = ""
            sapWs.cells.Find("Job Purpose:").offSet(0, 1).value = ""
        End If
    Next sapWs
End Sub

