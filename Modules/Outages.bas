Attribute VB_Name = "Outages"
Sub DownloadOutageLists()
    Dim project As project: Set project = New project
    Call project.extractFromSheets
    
    Dim token As String: token = GetToken
    If Not testToken(token) Then
        MsgBox "Invalid token, get an up to date one from GIS."
        Exit Sub
    End If
    
    Dim serviceTLM As String
    
    count = 0
    
    Dim locationTLMs As Scripting.Dictionary: Set locationTLMs = New Scripting.Dictionary
    Set locationTLMs("1") = New Collection
    locationTLMs("1").Add "0612111102"
    
    Dim poleCollections As Collection: Set poleCollections = findPoleGroups(project.poles)
    Dim poleCollection As Collection
    For Each poleCollection In poleCollections
        For Each pole In poleCollection
            If locationTLMs.exists(pole.location) Then
                Dim outageList As Collection: Set outageList = New Collection
                Set serviceJson = getElectricJson(poleCollection, 26, token)
                If Not serviceJson Is Nothing Then
                    For Each jsonFeature In serviceJson("features")
                        serviceTLM = jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServicePoint.TLM")
                        For Each tlm In locationTLMs(pole.location)
                            If serviceTLM = tlm Then
                                count = count + 1
                                Dim row As Collection: Set row = New Collection
                                 
                                accountNumber = jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.ACCOUNTNUMBER")
                                accountType = jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.ACCOUNTTYPE")
                                lastName = jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.LASTNAME")
                                firstName = jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.FIRSTNAME")
                                street = jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.STREET")
                                city = jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.CITY")
                                postalCode = jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.POSTALCODE")
                                telephone1 = jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.TELEPHONE1")
                                telephone2 = jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.TELEPHONE2")
                                telephone3 = jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.TELEPHONE3")
                                telephone4 = jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.TELEPHONE4")
                                connectStatus = jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServicePoint.CONNECTIONTYPE")
                                meter = jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServicePoint.METERNUMBER")
                                If InStr(meter, ".") > 0 Then
                                    parts = Split(meter, ".")
                                    meter = parts(UBound(parts))
                                End If
                                phase = jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServicePoint.PHASEDESIGNATION")
                                Group = "Other"
                                 
                                If accountNumber <> "" Then
                                 
                                   row.Add accountNumber
                                   row.Add accountType
                                   row.Add lastName
                                   row.Add firstName
                                   row.Add street
                                   row.Add city
                                   row.Add "MI"
                                   row.Add postalCode
                                   row.Add telephone1
                                   row.Add telephone2
                                   row.Add telephone3
                                   row.Add telephone4
                                   row.Add ""
                                   row.Add ""
                                   row.Add connectStatus
                                   row.Add meter
                                   row.Add ""
                                   row.Add serviceTLM
                                   row.Add phase
                                   row.Add Group
                                    
                                   outageList.Add row
                                End If
                            End If
                        Next tlm
                    Next jsonFeature
                End If
                
                Application.EnableEvents = False
                Application.ScreenUpdating = False
                Application.DisplayAlerts = False ' Suppresses prompts if file already exists
                
                  ' 3. Create a brand new Excel WorkBook in the background
                Dim NewBook As Workbook
                Set NewBook = Workbooks.Add(xlWBATWorksheet) ' xlWBATWorksheet builds it with exactly 1 sheet
                Dim sheet As Worksheet: Set sheet = NewBook.sheets(1)
                
                sheet.Cells(1, 1).Value = "Name"
                sheet.Cells(1, 2).Value = "ReportDate"
                sheet.Cells(1, 3).Value = "DeviceType"
                sheet.Cells(1, 4).Value = "DeviceFID"
                sheet.Cells(1, 5).Value = "DeviceOID"
                sheet.Cells(1, 6).Value = "EID"
                sheet.Cells(1, 7).Value = "FeederID"
                sheet.Cells(1, 8).Value = "FeederID2"
                sheet.Cells(1, 9).Value = "Substation"
                sheet.Cells(1, 10).Value = "Message"
                sheet.Cells(1, 11).Value = "Start Date"
                sheet.Cells(1, 12).Value = "Finish Date"
                sheet.Cells(1, 13).Value = "Alternate Start Date"
                sheet.Cells(1, 14).Value = "Alternate Finish Date"
                sheet.Cells(1, 15).Value = "ConnectedCustomers"
                sheet.Cells(1, 16).Value = "Priority"
                sheet.Cells(1, 17).Value = "Multiphase"
                sheet.Cells(1, 18).Value = "Other"
                sheet.Cells(1, 19).Value = "Critical"
                sheet.Cells(1, 20).Value = "Disconnected"
                sheet.Cells(1, 21).Value = "Connected"
                sheet.Cells(1, 22).Value = "DisconnectedCustomers"
                
                sheet.Cells(2, 1).Value = "Customer List"
                sheet.Cells(2, 2).Value = ""
                sheet.Cells(2, 3).Value = "Secondary Transformers"
                
                sheet.Cells(5, 1).Value = "Account"
                sheet.Cells(5, 2).Value = "AccountType"
                sheet.Cells(5, 3).Value = "LastName"
                sheet.Cells(5, 4).Value = "FirstName"
                sheet.Cells(5, 5).Value = "Street"
                sheet.Cells(5, 6).Value = "City"
                sheet.Cells(5, 7).Value = "State"
                sheet.Cells(5, 8).Value = "PostalCode"
                sheet.Cells(5, 9).Value = "Telephone"
                sheet.Cells(5, 10).Value = "Telephone2"
                sheet.Cells(5, 11).Value = "Telephone3"
                sheet.Cells(5, 12).Value = "Telephone4"
                sheet.Cells(5, 13).Value = "CallbackNumber"
                sheet.Cells(5, 14).Value = "CriticalCustomer"
                sheet.Cells(5, 15).Value = "ConnectStatus"
                sheet.Cells(5, 16).Value = "Meter"
                sheet.Cells(5, 17).Value = "DeviceOID"
                sheet.Cells(5, 18).Value = "TLM"
                sheet.Cells(5, 19).Value = "Phase"
                sheet.Cells(5, 20).Value = "Group"
                
                For i = 1 To outageList.count
                    For j = 1 To outageList(i).count
                        sheet.Cells(i + 5, j).Value = outageList(i)(j)
                    Next j
                Next i
                
                downloadsPath = Environ("USERPROFILE") & "\Downloads\"
                fileName = "CollectionExport_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsx"
                
                
                NewBook.SaveAs fileName:=downloadsPath & fileName, FileFormat:=xlOpenXMLWorkbook
                Application.DisplayAlerts = True
                Application.ScreenUpdating = True
                Application.EnableEvents = True
            End If
        Next pole
    Next poleCollection
End Sub
