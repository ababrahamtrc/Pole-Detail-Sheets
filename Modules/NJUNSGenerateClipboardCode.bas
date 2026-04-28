Attribute VB_Name = "NJUNSGenerateClipboardCode"
Const save As Boolean = True
Private previousSteps As Collection

Public Sub ExportAllNJUNS()
    Dim copiedCode As String
    
    Call LogMessage.SendLogMessage("ExportAllNJUNS")
    If Not checkCodes Then Exit Sub
    
    Dim Project As Project: Set Project = New Project
    Call Project.extractFromSheets
    
    Dim njunsSheets As scripting.Dictionary: Set njunsSheets = New scripting.Dictionary
    njunsSheets.Add "NOTIFY", New Collection
    njunsSheets.Add "CA", New Collection
    njunsSheets.Add "PT", New Collection
    
    Dim pole As pole
    Dim sheet As Worksheet
    For Each pole In Project.poles
        If Utilities.OnlyNumbers(pole.njunsTicket) = -1 And pole.NJUNS <> "" Then
            If InStr(pole.njunsTicket, "NOTIFY") > 0 Then njunsSheets("NOTIFY").Add pole
            If InStr(pole.njunsTicket, "CA") > 0 Then njunsSheets("CA").Add pole
            If InStr(pole.njunsTicket, "PT") > 0 Then njunsSheets("PT").Add pole
        End If
    Next pole
    
    Dim i As Integer
    Dim totalNJUNS As Integer
    totalNJUNS = njunsSheets("NOTIFY").count + njunsSheets("CA").count + njunsSheets("PT").count
    For Each ticketType In njunsSheets
        Dim previousNJUNS As String: previousNJUNS = ""
        Dim previousCompanies As String: previousCompanies = ""
        Call SortNJUNSPoles(njunsSheets(ticketType))
        For Each pole In njunsSheets(ticketType)
            Dim currentNJUNS As String: currentNJUNS = ""
            Dim currentCompanies As String: currentCompanies = ""
            i = i + 1
            
            currentNJUNS = Replace(Replace(pole.NJUNS, vbLf, ""), " ", "")
            
            For Each step In pole.njunsSteps
                company = Trim(UCase(Utilities.GetFirstWord(CStr(step))))
                currentCompanies = currentCompanies & company
            Next step
            
            If save And previousNJUNS = currentNJUNS Then
                copiedCode = copiedCode & getDuplicateSheetNJUNSCode(Project, pole)
            ElseIf save And previousCompanies = currentCompanies Then
                copiedCode = copiedCode & getAlmostDuplicateSheetNJUNSCode(Project, pole)
            Else
                copiedCode = copiedCode & getSelectProjectTabCode(Project) & vbLf & "await clickButton('.v-button.v-widget', 'Create');" & getSheetNJUNSCode(Project, pole, vbYes)
            End If
            
            If copiedCode <> "" Then copiedCode = copiedCode & vbLf & "console.log('" & i & "/" & totalNJUNS & "')"

            previousNJUNS = currentNJUNS
            previousCompanies = currentCompanies
        Next pole
    Next ticketType
    
    If copiedCode <> "" Then
        copiedCode = wrapCode(Project, copiedCode)
        
        Dim DataObj As DataObject: Set DataObj = New DataObject
        DataObj.SetText copiedCode
        DataObj.PutInClipboard
        
        MsgBox ("Copied code to clipboard for all tickets that don't already have ticket numbers on their sheet, add ticket numbers for existing tickets if you don't wish you create a ticket for each pole. go to NJUNS website and press f12 to paste code into console on the project tab for the project you wish to add these tickets to.")
    Else
        MsgBox ("Tickets already created or no NJUNS on sheets.")
    End If
    
    Set previousSteps = Nothing
End Sub

Public Sub ExportSingleNJUNS()
    Dim copiedCode As String
    
    Call LogMessage.SendLogMessage("ExportSingleNJUNS")
    If Not checkCodes Then Exit Sub
    
    Dim Project As Project: Set Project = New Project
    Call Project.extractFromSheets
    
    Dim sheet As Worksheet: Set sheet = ThisWorkbook.ActiveSheet()
    If Not Utilities.IsPDS(sheet) Then
        MsgBox "You need to have a pole detail sheet active to run this script."
        Exit Sub
    End If
    
    Dim pole As pole: Set pole = New pole
    Call pole.extractFromSheet(sheet)
    
    If Utilities.OnlyNumbers(pole.njunsTicket) = -1 And pole.NJUNS <> "" Then
        Dim result As VbMsgBoxResult
        result = MsgBox("Is this ticket accosiated with a project number? (this matters)", vbYesNoCancel + vbQuestion, "Confirm")
        If result <> vbYes And result <> vbNo Then Exit Sub
        If result = vbNo Then copiedCode = copiedCode & vbLf & "consumersCode = """ & InputBox("What is the EXACT CE NJUNS Code for this ticket? (CEJAC,CEMUS, etc...)", "User Input") & """"
        copiedCode = copiedCode & vbLf & IIf(result = vbYes, getSelectProjectTabCode(Project) & vbLf & "await clickButton('.v-button.v-widget', 'Create');", getCreateSingleTicketCode()) & getSheetNJUNSCode(Project, pole, result)
    End If
    
    If copiedCode <> "" Then
        copiedCode = wrapCode(Project, copiedCode)
    
        Dim DataObj As DataObject: Set DataObj = New DataObject
        DataObj.SetText copiedCode
        DataObj.PutInClipboard
        
        If result = vbYes Then
            MsgBox ("Copied code to clipboard, go to NJUNS website and press f12 to paste code into console on the project tab for the project you wish to add the ticket to.")
        Else
            MsgBox ("Copied code to clipboard, go to NJUNS website and press f12 to paste code into console.")
        End If
    Else
        MsgBox ("Ticket already created or no NJUNS on sheet.")
    End If
    
    Set previousSteps = Nothing
End Sub

Private Sub SortNJUNSPoles(poles As Collection)
 
    If poles.count = 0 Then Exit Sub
 
    Dim arr() As Variant
    Dim i As Long

    ReDim arr(1 To poles.count)
    For i = 1 To poles.count
        Set arr(i) = poles(i)
    Next i
 
    QuickSortPoles arr, LBound(arr), UBound(arr)
 
    Do While poles.count > 0
        poles.Remove 1
    Loop
 
    For i = LBound(arr) To UBound(arr)
        poles.Add arr(i)
    Next i
 
End Sub
 
Private Sub QuickSortPoles(arr As Variant, first As Long, last As Long)
    Dim i As Long, j As Long
    Dim pivot As String
    Dim temp As Object
 
    i = first
    j = last
    pivot = GetPoleSortKey(arr((first + last) \ 2))
 
    Do While i <= j
        Do While GetPoleSortKey(arr(i)) < pivot
            i = i + 1
        Loop
 
        Do While GetPoleSortKey(arr(j)) > pivot
            j = j - 1
        Loop
 
        If i <= j Then
            Set temp = arr(i)
            Set arr(i) = arr(j)
            Set arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop
 
    If first < j Then QuickSortPoles arr, first, j
    If i < last Then QuickSortPoles arr, i, last
End Sub
 
Private Function GetPoleSortKey(ByVal pole As pole) As String
    Dim key As String
    Dim step As Variant
    key = Replace(Replace(pole.NJUNS, vbLf, ""), " ", "") & "|"

    For Each step In pole.njunsSteps
        company = UCase(Utilities.GetFirstWord(CStr(step)))
        key = key & company & ">"
    Next step
    GetPoleSortKey = key
End Function

Private Function getSelectProjectTabCode(Project As Project) As String
    Dim copiedCode As String
    
    copiedCode = copiedCode & vbLf & "navigated = false;"
    copiedCode = copiedCode & vbLf & "while(true) {"
    copiedCode = copiedCode & vbLf & "if (!findElementByText('.v-tabsheet-tabitemcell','" & Project.Notification & "')) throw new Error(""Project tab Doesn't Exist. Please create the project and have the project tab open."");"
    copiedCode = copiedCode & vbLf & "if (findElementByText('.v-tabsheet-tabitemcell','" & Project.Notification & "').ariaSelected === 'true') break;"

    copiedCode = copiedCode & vbLf & "realClick(findElementByText('.v-tabsheet-tabitemcell', '" & Project.Notification & "'));"
    copiedCode = copiedCode & vbLf & "navigated = true;"
    copiedCode = copiedCode & vbLf & "if (findElementByText('.v-tabsheet-tabitemcell','" & Project.Notification & "').ariaSelected === 'true') break;"
    
    copiedCode = copiedCode & vbLf & "if (document.querySelector('.v-tabsheet-scrollerPrev-disabled')) {"
    copiedCode = copiedCode & vbLf & "throw new Error(""Can't find job"");"
    copiedCode = copiedCode & vbLf & "} else {"
    copiedCode = copiedCode & vbLf & "realClick(document.querySelector('.v-tabsheet-scrollerPrev'));"
    copiedCode = copiedCode & vbLf & "}"
    copiedCode = copiedCode & vbLf & "}"
    
    
    copiedCode = copiedCode & vbLf & "if (navigated) await waitLoadingTime();"
    copiedCode = copiedCode & vbLf & "if(!consumersCode) consumersCode = document.querySelectorAll('.v-filterselect-input')[0].value;"
    copiedCode = copiedCode & vbLf & "if (navigated) await waitLoadingTime();"

    getSelectProjectTabCode = copiedCode
End Function

Private Function getCreateSingleTicketCode() As String
    Dim copiedCode As String
    
    copiedCode = copiedCode & vbLf & "if (!await clickButton("".v-button.v-widget.link.v-button-link"", ""New Ticket"")) {"
    copiedCode = copiedCode & vbLf & "await realClick(document.querySelectorAll("".v-slot.v-slot-c-app-icon.v-align-middle"")[0]);"
    copiedCode = copiedCode & vbLf & "await clickButton("".v-button.v-widget.link.v-button-link"", ""New Ticket"");"
    copiedCode = copiedCode & vbLf & "}"
    
    getCreateSingleTicketCode = copiedCode
End Function

Private Function wrapCode(Project As Project, copiedCode As String) As String
    
    On Error Resume Next
    copiedCode = "let notification = " & Project.Notification & ";" & vbLf & ThisWorkbook.Worksheets("NJUNSCode").Cells(1, 1) & vbLf & "try {" & vbLf & "while(true) {" & vbLf & copiedCode & vbLf & "break;" & vbLf & "}" & vbLf & "}"
    On Error GoTo 0
    
    alertMsg = "Something went wrong while generating tickets or execution was cancelled. Upload the generated CSV in your downloads folder into the pole detail sheets control file and rexport NJUNS to avoid recreating tickets already generated."
    
    copiedCode = copiedCode & "catch(err) {"
    copiedCode = copiedCode & vbLf & "console.error(err)"
    If save Then copiedCode = copiedCode & vbLf & "if (pdfList.length > 0) alert('" & alertMsg & "')"
    copiedCode = copiedCode & vbLf & "}"
    
    copiedCode = copiedCode & "finally {"
    If save Then copiedCode = copiedCode & vbLf & "if (pdfList.length > 0 && cancelled) alert('" & alertMsg & "')"
    copiedCode = copiedCode & vbLf & "btn.remove()"
    If save Then copiedCode = copiedCode & vbLf & "await downloadAllPdfs();"
    copiedCode = copiedCode & vbLf & "}"
    
    wrapCode = copiedCode
End Function

Private Function checkCodes() As Boolean
    Dim controlWs As Worksheet: Set controlWs = ThisWorkbook.sheets("Control")
    If controlWs.Range("NJUNSCODES").Value = "" Then Call NJUNSCodes.generateNJUNSCodes
    Dim codeRange As Range: Set codeRange = controlWs.Range("NJUNSCODES").EntireColumn
    lastRow = controlWs.Cells(controlWs.Rows.count, controlWs.Range("NJUNSCODES").Column).End(xlUp).row
    Set codeRange = controlWs.Range(controlWs.Cells(3, controlWs.Range("NJUNSCODES").Column), controlWs.Cells(lastRow, controlWs.Range("NJUNSCODES").Column))
    
    Dim cell As Range
    For Each cell In codeRange
        If cell.Value = "" Then Exit For
        If cell.offset(0, 1).Value = "" Then
            MsgBox ("Missing values in the control page for NJUNS codes, these need to be all filled in first")
            checkCodes = False
            Exit Function
        End If
    Next cell
    checkCodes = True
End Function
    
Private Function getSheetNJUNSCode(Project As Project, pole As pole, Optional projectResult As VbMsgBoxResult) As String
    Dim copiedCode As String
    
    Dim controlWs As Worksheet: Set controlWs = ThisWorkbook.sheets("Control")
    If controlWs.Range("NJUNSCODES").Value = "" Then Call NJUNSCodes.generateNJUNSCodes
    Dim codeRange As Range: Set codeRange = controlWs.Range("NJUNSCODES").EntireColumn
    lastRow = controlWs.Cells(controlWs.Rows.count, controlWs.Range("NJUNSCODES").Column).End(xlUp).row
    Set codeRange = controlWs.Range(controlWs.Cells(3, controlWs.Range("NJUNSCODES").Column), controlWs.Cells(lastRow, controlWs.Range("NJUNSCODES").Column))
    
    If pole.njunsSteps.count = 0 Then
        Exit Function
    End If
    
    Set previousSteps = Nothing
    
    Dim ticketType As String
    If InStr(pole.njunsTicket, "CA") = 1 Then
        ticketType = "Violation (VIO)"
        NJUNSType = "CA"
    ElseIf InStr(pole.njunsTicket, "NOTIFY") = 1 Then
        ticketType = "Violation (VIO)"
        NJUNSType = "NOTIFY"
    ElseIf InStr(pole.njunsTicket, "PT") = 1 Then
        ticketType = "Pole Transfer (PT)"
        NJUNSType = "PT"
    Else
        MsgBox ("NJUNS ticket type needs to be CA/PT/NOTIFY")
        Exit Function
    End If
    
    Dim njunsRemarks As String
    njunsRemarks = "Notification: " & Project.Notification & " Permit: " & Project.permit & "\n"
    If ticketType = "Violation (VIO)" Then
        njunsRemarks = njunsRemarks & "Corrective Violation "
    Else
        njunsRemarks = njunsRemarks & "Pole Transfer "
    End If
    njunsRemarks = njunsRemarks & pole.address
    
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    
    If projectResult = vbYes Then
        copiedCode = copiedCode & vbLf & "await selectDropDownOption(document.querySelectorAll("".v-filterselect-input"")[3], """ & ticketType & """, true);"
        copiedCode = copiedCode & vbLf & "if(cancelled) break;"
        copiedCode = copiedCode & vbLf & "await selectDropDownOption(document.querySelectorAll("".v-filterselect-input"")[4], ""Michigan"", true);"
        copiedCode = copiedCode & vbLf & "if(cancelled) break;"
        copiedCode = copiedCode & vbLf & "await selectDropDownOption(document.querySelectorAll("".v-filterselect-input"")[5], """ & Application.WorksheetFunction.Proper(Project.county) & """, true);"
        copiedCode = copiedCode & vbLf & "if(cancelled) break;"
        copiedCode = copiedCode & vbLf & "await selectDropDownOption(document.querySelectorAll("".v-filterselect-input"")[6], """ & Application.WorksheetFunction.Proper(findNJUNSCode("Place")) & """, true);"
        copiedCode = copiedCode & vbLf & "if(cancelled) break;"
        copiedCode = copiedCode & vbLf & "await selectDropDownOption(document.querySelectorAll("".v-filterselect-input"")[7], consumersCode, true)"
        copiedCode = copiedCode & vbLf & "if(cancelled) break;"
        If NJUNSType = "NOTIFY" Then
            copiedCode = copiedCode & vbLf & "await commitLookupValue(document.querySelectorAll("".v-filterselect-input"")[8], """ & findNJUNSCode(Utilities.GetFirstWord(CStr(pole.njunsSteps(1)))) & """);"
        Else
            copiedCode = copiedCode & vbLf & "await commitLookupValue(document.querySelectorAll("".v-filterselect-input"")[8], consumersCode);"
        End If
        copiedCode = copiedCode & vbLf & "if(cancelled) break;"
        If ticketType = "Violation (VIO)" Then
            copiedCode = copiedCode & vbLf & "await selectDropDownOption(document.querySelectorAll("".v-filterselect-input"")[9], ""VIO:PT-Default"", true);"
        Else
            copiedCode = copiedCode & vbLf & "await selectDropDownOption(document.querySelectorAll("".v-filterselect-input"")[9], ""PT:PT-Default"", true);"
        End If
        copiedCode = copiedCode & vbLf & "if(cancelled) break;"
        copiedCode = copiedCode & vbLf & "await commitLookupValue(document.querySelectorAll("".v-filterselect-input"")[10], """ & findNJUNSCode("Applicant") & """);"
    ElseIf projectResult = vbNo Then
        copiedCode = copiedCode & vbLf & "await selectDropDownOption(document.querySelectorAll("".v-filterselect-input"")[0], """ & ticketType & """,true);"
        copiedCode = copiedCode & vbLf & "if(cancelled) break;"
        copiedCode = copiedCode & vbLf & "await selectDropDownOption(document.querySelectorAll("".v-filterselect-input"")[1], ""Michigan"",true);"
        copiedCode = copiedCode & vbLf & "if(cancelled) break;"
        copiedCode = copiedCode & vbLf & "await selectDropDownOption(document.querySelectorAll("".v-filterselect-input"")[2], """ & Application.WorksheetFunction.Proper(Project.county) & """,true);"
        copiedCode = copiedCode & vbLf & "if(cancelled) break;"
        copiedCode = copiedCode & vbLf & "await selectDropDownOption(document.querySelectorAll("".v-filterselect-input"")[3], """ & Application.WorksheetFunction.Proper(findNJUNSCode("Place")) & """,true);"
        copiedCode = copiedCode & vbLf & "if(cancelled) break;"
        copiedCode = copiedCode & vbLf & "await selectDropDownOption(document.querySelectorAll("".v-filterselect-input"")[4], consumersCode,true)"
        copiedCode = copiedCode & vbLf & "if(cancelled) break;"
        copiedCode = copiedCode & vbLf & "await commitLookupValue(document.querySelectorAll("".v-filterselect-input"")[5], consumersCode);"
        copiedCode = copiedCode & vbLf & "if(cancelled) break;"
        If ticketType = "Violation (VIO)" Then
            copiedCode = copiedCode & vbLf & "await selectDropDownOption(document.querySelectorAll("".v-filterselect-input"")[6], ""VIO:PT-Default"",true);"
        Else
            copiedCode = copiedCode & vbLf & "await selectDropDownOption(document.querySelectorAll("".v-filterselect-input"")[6], ""PT:PT-Default"",true);"
        End If
        copiedCode = copiedCode & vbLf & "if(cancelled) break;"
        copiedCode = copiedCode & vbLf & "await commitLookupValue(document.querySelectorAll("".v-filterselect-input"")[7], """ & findNJUNSCode("Applicant") & """);"
    End If
    
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    
    copiedCode = copiedCode & vbLf & "await clickButton("".v-button.v-widget"", ""Create New Ticket"");"
    
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelectorAll("".v-textfield.v-widget.v-has-width"")[8], ""Holly Webb"");"
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelectorAll("".v-textfield.v-widget.v-has-width"")[9], ""517-788-1690"");"
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelectorAll("".v-textfield.v-widget.v-has-width"")[10],  ""holly.webb@cmsenergy.com"");"
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelectorAll("".v-textfield.v-widget.v-has-width"")[11],  """ & Project.Notification & """);"
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelector("".v-textarea.v-widget.v-has-width""), """ & njunsRemarks & """);"
    copiedCode = copiedCode & vbLf & "await selectDropDownOption(document.querySelector("".v-filterselect-input""), ""3"");"
    
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    
    Dim company As String, njunsCode As String, remarks As String, stepType As String
    Dim stepCounter As Integer
    Set previousSteps = New Collection
    For Each step In pole.njunsSteps
        stepCounter = stepCounter + 1
        company = UCase(Utilities.GetFirstWord(CStr(step)))
        If company = "CE" Then company = "CONSUMERS"
        If company <> "CONSUMERS" Then
            njunsCode = findNJUNSCode(company)
        End If
        remarks = Replace(step, vbLf, "\n")
        remarks = Replace(remarks, """", "\""")
        If InStr(pole.njunsTicket, "CA") = 1 Then
            stepType = "VIOLATION"
        ElseIf InStr(pole.njunsTicket, "NOTIFY") = 1 Then
            stepType = "NOTIFY"
        ElseIf InStr(pole.njunsTicket, "PT") = 1 Then
            stepType = "TRANSFER"
            If stepCounter = 1 Or stepCounter = pole.njunsSteps.count Then
                If company <> "CONSUMERS" Then
                    If stepCounter = 1 Then
                        copiedCode = copiedCode & vbLf & "await generateStep(consumersCode, ""Consumers to complete required work."", ""SET POLE"", """ & Project.Notification & """);"
                        previousSteps.Add "Consumers to complete required work."
                    End If
                Else
                    If stepCounter = 1 Then
                        stepType = "SET POLE"
                    Else
                        stepType = "PULL POLE"
                    End If
                End If
            End If
        End If
        
        copiedCode = copiedCode & vbLf & "if(cancelled) break;"
        If company = "CONSUMERS" Then
            copiedCode = copiedCode & vbLf & "await generateStep(consumersCode, """ & remarks & """, """ & stepType & """, """ & Project.Notification & """);"
        Else
            copiedCode = copiedCode & vbLf & "await generateStep(""" & njunsCode & """, """ & remarks & """, """ & stepType & """, """ & Project.Notification & """);"
        End If
        previousSteps.Add remarks
    
        If company <> "CONSUMERS" And stepCounter = pole.njunsSteps.count And InStr(pole.njunsTicket, "PT") = 1 Then
            copiedCode = copiedCode & vbLf & "if(cancelled) break;"
            copiedCode = copiedCode & vbLf & "await generateStep(consumersCode, ""Consumers after comms transfer to new pole, pull topped pole."", ""PULL POLE"", """ & Project.Notification & """);"
            copiedCode = copiedCode & vbLf & "if(cancelled) break;"
            previousSteps.Add "Consumers after comms transfer to new pole, pull topped pole."
        End If
    Next step
    
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    copiedCode = copiedCode & vbLf & "await clickButton("".v-captiontext"", ""Poles/Assets"");"
    copiedCode = copiedCode & vbLf & "await clickButton("".v-button.v-widget.c-primary-action.v-button-c-primary-action.icon"", ""Create"");"
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelector("".v-textfield.v-textfield-large.bold""), """ & pole.existingCEID & """);"
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelectorAll("".v-textfield.v-widget.v-has-width"")[8], """ & pole.latitude & """);"
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelectorAll("".v-textfield.v-widget.v-has-width"")[9], """ & pole.longitude & """);"
    
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    
    copiedCode = copiedCode & vbLf & "await clickButton("".v-button.v-widget.icon"", ""Enable Geocoding"");"
    copiedCode = copiedCode & vbLf & "await clickButton("".v-button.v-widget.primary"", ""Create"");"
    copiedCode = copiedCode & vbLf & "await clickButton("".v-button.v-widget.primary"", ""Apply Changes"");"
    
    If save Then
        copiedCode = copiedCode & vbLf & "if(cancelled) break;"
        copiedCode = copiedCode & vbLf & "await clickButton("".v-button.v-widget.blueIcon.v-button-blueIcon.icon"", ""Save"");"
        copiedCode = copiedCode & vbLf & "ticketNumber = document.querySelectorAll("".v-textfield.v-textfield"")[1].value;"
        
        copiedCode = copiedCode & vbLf & "if(cancelled) break;"
        
        copiedCode = copiedCode & vbLf & "pdfList.push({"
            copiedCode = copiedCode & vbLf & "url: download ? await getPDFURL() : null,"
            copiedCode = copiedCode & vbLf & "filename: `" & Project.Notification & " - " & "${ticketNumber} - " & NJUNSType & " TICKET`,"
            copiedCode = copiedCode & vbLf & "poleNumber: """ & pole.poleNumber & ""","
            copiedCode = copiedCode & vbLf & "ticketNumber: ticketNumber"
        copiedCode = copiedCode & vbLf & "});"
    End If
    
    getSheetNJUNSCode = copiedCode
End Function

Private Function getDuplicateSheetNJUNSCode(Project As Project, pole As pole) As String
    Dim copiedCode As String
    
    Dim controlWs As Worksheet: Set controlWs = ThisWorkbook.sheets("Control")
    If controlWs.Range("NJUNSCODES").Value = "" Then Call NJUNSCodes.generateNJUNSCodes
    Dim codeRange As Range: Set codeRange = controlWs.Range("NJUNSCODES").EntireColumn
    lastRow = controlWs.Cells(controlWs.Rows.count, controlWs.Range("NJUNSCODES").Column).End(xlUp).row
    Set codeRange = controlWs.Range(controlWs.Cells(3, controlWs.Range("NJUNSCODES").Column), controlWs.Cells(lastRow, controlWs.Range("NJUNSCODES").Column))
    
    If pole.njunsSteps.count = 0 Then
        Exit Function
    End If
    
    Dim ticketType As String
    If InStr(pole.njunsTicket, "CA") = 1 Then
        ticketType = "Violation (VIO)"
        NJUNSType = "CA"
    ElseIf InStr(pole.njunsTicket, "NOTIFY") = 1 Then
        ticketType = "Violation (VIO)"
        NJUNSType = "NOTIFY"
    ElseIf InStr(pole.njunsTicket, "PT") = 1 Then
        ticketType = "Pole Transfer (PT)"
        NJUNSType = "PT"
    Else
        MsgBox ("NJUNS ticket type needs to be CA/PT/NOTIFY")
        Exit Function
    End If
    
    Dim njunsRemarks As String
    njunsRemarks = "Notification: " & Project.Notification & " Permit: " & Project.permit & "\n"
    If ticketType = "Violation (VIO)" Then
        njunsRemarks = njunsRemarks & "Corrective Violation "
    Else
        njunsRemarks = njunsRemarks & "Pole Transfer "
    End If
    njunsRemarks = njunsRemarks & pole.address
    
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    copiedCode = copiedCode & vbLf & "await clickButton('.v-button.v-widget.v-popupbutton.borderless.v-button-borderless.noIndicator.v-button-noIndicator.icon.v-button-icon','Actions');"
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    copiedCode = copiedCode & vbLf & "await clickButton('.c-cm-button.v-widget.v-has-width','Clone Ticket');"
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    copiedCode = copiedCode & vbLf & "realClick(document.querySelectorAll('input')[document.querySelectorAll('input').length - 4]);"
    copiedCode = copiedCode & vbLf & "realClick(document.querySelectorAll('input')[document.querySelectorAll('input').length - 3]);"
    copiedCode = copiedCode & vbLf & "realClick(document.querySelectorAll('input')[document.querySelectorAll('input').length - 2]);"
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    copiedCode = copiedCode & vbLf & "await realClick(document.querySelectorAll('input')[document.querySelectorAll('input').length - 1]);"
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    copiedCode = copiedCode & vbLf & "await clickButton('.v-button.v-widget.icon', 'Copy');"
    
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelector("".v-textarea.v-widget.v-has-width""), """ & njunsRemarks & """);"
    
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    copiedCode = copiedCode & vbLf & "await clickButton("".v-captiontext"", ""Poles/Assets"");"
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    copiedCode = copiedCode & vbLf & "await realClick(document.querySelector('.v-grid-row.v-grid-row-has-data'));"
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    copiedCode = copiedCode & vbLf & "await clickButton('.v-button.v-widget.icon', 'Edit');"
    copiedCode = copiedCode & vbLf & "await clickButton('.v-button.v-widget.icon', 'Disable Geocoding');"
    
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelector("".v-textfield.v-textfield-large.bold""), """ & pole.existingCEID & """);"
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelectorAll("".v-textfield.v-widget.v-has-width"")[8], """ & pole.latitude & """);"
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelectorAll("".v-textfield.v-widget.v-has-width"")[9], """ & pole.longitude & """);"
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelectorAll("".v-textfield.v-widget.v-has-width"")[10], """");"
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelectorAll("".v-textfield.v-widget.v-has-width"")[11], """");"
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelectorAll("".v-textfield.v-widget.v-has-width"")[13], """");"
    
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    copiedCode = copiedCode & vbLf & "await clickButton("".v-button.v-widget.icon"", ""Enable Geocoding"");"
    copiedCode = copiedCode & vbLf & "await waitLoadingTime();"
    copiedCode = copiedCode & vbLf & "await clickButton("".v-button.v-widget.icon"", ""Enable Geocoding"");"
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    copiedCode = copiedCode & vbLf & "await clickButton("".v-button.v-widget.primary"", ""Create"");"
    copiedCode = copiedCode & vbLf & "await clickButton("".v-button.v-widget.primary"", ""Apply Changes"");"
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    
    If save Then
        copiedCode = copiedCode & vbLf & "await clickButton("".v-button.v-widget.blueIcon.v-button-blueIcon.icon"", ""Save"");"
        copiedCode = copiedCode & vbLf & "ticketNumber = document.querySelectorAll("".v-textfield.v-textfield"")[1].value;"
        
        copiedCode = copiedCode & vbLf & "if(cancelled) break;"
        
        copiedCode = copiedCode & vbLf & "pdfList.push({"
            copiedCode = copiedCode & vbLf & "url: download ? await getPDFURL() : null,"
            copiedCode = copiedCode & vbLf & "filename: `" & Project.Notification & " - " & "${ticketNumber} - " & NJUNSType & " TICKET`,"
            copiedCode = copiedCode & vbLf & "poleNumber: """ & pole.poleNumber & ""","
            copiedCode = copiedCode & vbLf & "ticketNumber: ticketNumber"
        copiedCode = copiedCode & vbLf & "});"
        
        copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    End If
    
    getDuplicateSheetNJUNSCode = copiedCode
End Function

Private Function getAlmostDuplicateSheetNJUNSCode(Project As Project, pole As pole) As String
    Dim copiedCode As String
    
    Dim controlWs As Worksheet: Set controlWs = ThisWorkbook.sheets("Control")
    If controlWs.Range("NJUNSCODES").Value = "" Then Call NJUNSCodes.generateNJUNSCodes
    Dim codeRange As Range: Set codeRange = controlWs.Range("NJUNSCODES").EntireColumn
    lastRow = controlWs.Cells(controlWs.Rows.count, controlWs.Range("NJUNSCODES").Column).End(xlUp).row
    Set codeRange = controlWs.Range(controlWs.Cells(3, controlWs.Range("NJUNSCODES").Column), controlWs.Cells(lastRow, controlWs.Range("NJUNSCODES").Column))
    
    If pole.njunsSteps.count = 0 Then
        Exit Function
    End If
    
    Dim ticketType As String
    If InStr(pole.njunsTicket, "CA") = 1 Then
        ticketType = "Violation (VIO)"
        NJUNSType = "CA"
    ElseIf InStr(pole.njunsTicket, "NOTIFY") = 1 Then
        ticketType = "Violation (VIO)"
        NJUNSType = "NOTIFY"
    ElseIf InStr(pole.njunsTicket, "PT") = 1 Then
        ticketType = "Pole Transfer (PT)"
        NJUNSType = "PT"
    Else
        MsgBox ("NJUNS ticket type needs to be CA/PT/NOTIFY")
        Exit Function
    End If
    
    Dim njunsRemarks As String
    njunsRemarks = "Notification: " & Project.Notification & " Permit: " & Project.permit & "\n"
    If ticketType = "Violation (VIO)" Then
        njunsRemarks = njunsRemarks & "Corrective Violation "
    Else
        njunsRemarks = njunsRemarks & "Pole Transfer "
    End If
    njunsRemarks = njunsRemarks & pole.address
    
    copiedCode = copiedCode & vbLf & "await clickButton('.v-button.v-widget.v-popupbutton.borderless.v-button-borderless.noIndicator.v-button-noIndicator.icon.v-button-icon','Actions');"
    copiedCode = copiedCode & vbLf & "await clickButton('.c-cm-button.v-widget.v-has-width','Clone Ticket');"
    copiedCode = copiedCode & vbLf & "realClick(document.querySelectorAll('input')[document.querySelectorAll('input').length - 4]);"
    copiedCode = copiedCode & vbLf & "realClick(document.querySelectorAll('input')[document.querySelectorAll('input').length - 3]);"
    copiedCode = copiedCode & vbLf & "realClick(document.querySelectorAll('input')[document.querySelectorAll('input').length - 2]);"
    copiedCode = copiedCode & vbLf & "realClick(document.querySelectorAll('input')[document.querySelectorAll('input').length - 1]);"
    copiedCode = copiedCode & vbLf & "await clickButton('.v-button.v-widget.icon', 'Copy');"
    
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelector("".v-textarea.v-widget.v-has-width""), """ & njunsRemarks & """);"
    Dim offset As Integer
    Dim steps As Collection: Set steps = New Collection
    For i = 0 To pole.njunsSteps.count - 1
        step = pole.njunsSteps(i + 1)
        company = UCase(Utilities.GetFirstWord(CStr(step)))
        If NJUNSType = "PT" Then
            If i = 0 Then
                If company <> "CE" And company <> "CONSUMERS" Then
                    If previousSteps(i + 1) <> "Consumers to complete required work." Then
                        copiedCode = copiedCode & vbLf & "await realClick(document.querySelectorAll('.v-grid-row.v-grid-row-has-data')[" & i & "]);"
                        copiedCode = copiedCode & vbLf & "await clickButton('.v-button.v-widget.icon', 'Edit');"
                        copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelectorAll('.v-textarea.v-widget.v-has-width')[1], 'Consumers to complete required work.');"
                        copiedCode = copiedCode & vbLf & "await clickButton('.v-button.v-widget.primary', 'Create');"
                    End If
                    offset = 1
                    steps.Add "Consumers to complete required work."
                End If
            End If
        End If
        step = Replace(step, vbLf, "\n")
        step = Replace(step, """", "\""")
        If previousSteps(i + 1 + offset) <> step Then
            copiedCode = copiedCode & vbLf & "await realClick(document.querySelectorAll('.v-grid-row.v-grid-row-has-data')[" & i + offset & "]);"
            copiedCode = copiedCode & vbLf & "await clickButton('.v-button.v-widget.icon', 'Edit');"
            copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelectorAll('.v-textarea.v-widget.v-has-width')[1], """ & step & """);"
            copiedCode = copiedCode & vbLf & "await clickButton('.v-button.v-widget.primary', 'Create');"
        End If
        steps.Add step
        If NJUNSType = "PT" Then
            If i = pole.njunsSteps.count - 1 Then
                If company <> "CE" And company <> "CONSUMERS" Then
                    If previousSteps(i + 1 + offset) <> "Consumers after comms transfer to new pole, pull topped pole." Then
                        copiedCode = copiedCode & vbLf & "await realClick(document.querySelectorAll('.v-grid-row.v-grid-row-has-data')[" & i + offset & "]);"
                        copiedCode = copiedCode & vbLf & "await clickButton('.v-button.v-widget.icon', 'Edit');"
                        copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelectorAll('.v-textarea.v-widget.v-has-width')[1], 'Consumers after comms transfer to new pole, pull topped pole.');"
                        copiedCode = copiedCode & vbLf & "await clickButton('.v-button.v-widget.primary', 'Create');"
                    End If
                    steps.Add "Consumers after comms transfer to new pole, pull topped pole."
                End If
            End If
        End If
    Next i
    Set previousSteps = steps
    
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    copiedCode = copiedCode & vbLf & "await clickButton("".v-captiontext"", ""Poles/Assets"");"
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    copiedCode = copiedCode & vbLf & "await realClick(document.querySelector('.v-grid-row.v-grid-row-focused.v-grid-row-has-data'));"
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    copiedCode = copiedCode & vbLf & "await clickButton('.v-button.v-widget.icon', 'Edit');"
    copiedCode = copiedCode & vbLf & "await clickButton('.v-button.v-widget.icon', 'Disable Geocoding');"
    
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelector("".v-textfield.v-textfield-large.bold""), """ & pole.existingCEID & """);"
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelectorAll("".v-textfield.v-widget.v-has-width"")[8], """ & pole.latitude & """);"
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelectorAll("".v-textfield.v-widget.v-has-width"")[9], """ & pole.longitude & """);"
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelectorAll("".v-textfield.v-widget.v-has-width"")[10], """");"
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelectorAll("".v-textfield.v-widget.v-has-width"")[11], """");"
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelectorAll("".v-textfield.v-widget.v-has-width"")[13], """");"
    
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    copiedCode = copiedCode & vbLf & "await clickButton("".v-button.v-widget.icon"", ""Enable Geocoding"");"
    copiedCode = copiedCode & vbLf & "await waitLoadingTime();"
    copiedCode = copiedCode & vbLf & "await clickButton("".v-button.v-widget.icon"", ""Enable Geocoding"");"
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    copiedCode = copiedCode & vbLf & "await clickButton("".v-button.v-widget.primary"", ""Create"");"
    copiedCode = copiedCode & vbLf & "await clickButton("".v-button.v-widget.primary"", ""Apply Changes"");"
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    
    If save Then
        copiedCode = copiedCode & vbLf & "await clickButton("".v-button.v-widget.blueIcon.v-button-blueIcon.icon"", ""Save"");"
        copiedCode = copiedCode & vbLf & "ticketNumber = document.querySelectorAll("".v-textfield.v-textfield"")[1].value;"
        
        copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    
    
        copiedCode = copiedCode & vbLf & "pdfList.push({"
            copiedCode = copiedCode & vbLf & "url: download ? await getPDFURL() : null,"
            copiedCode = copiedCode & vbLf & "filename: `" & Project.Notification & " - " & "${ticketNumber} - " & NJUNSType & " TICKET`,"
            copiedCode = copiedCode & vbLf & "poleNumber: """ & pole.poleNumber & ""","
            copiedCode = copiedCode & vbLf & "ticketNumber: ticketNumber"
        copiedCode = copiedCode & vbLf & "});"
    End If
    
    getAlmostDuplicateSheetNJUNSCode = copiedCode
End Function

Private Function findNJUNSCode(ByVal company As String) As String
    Dim controlWs As Worksheet: Set controlWs = ThisWorkbook.sheets("Control")
    Dim codeRange As Range: Set codeRange = controlWs.Range("NJUNSCODES").EntireColumn
    lastRow = controlWs.Cells(controlWs.Rows.count, controlWs.Range("NJUNSCODES").Column).End(xlUp).row
    Set codeRange = controlWs.Range(controlWs.Cells(3, controlWs.Range("NJUNSCODES").Column), controlWs.Cells(lastRow, controlWs.Range("NJUNSCODES").Column))

    company = UCase(Replace(company, " ", ""))
    company = Replace(company, ":", "")
    company = Replace(company, vbLf, "")

    Dim cell As Range
    For Each cell In codeRange
        If isEmpty(cell.Value) Then Exit For
        If InStr(1, CStr(cell.Value), company, vbTextCompare) = 1 Then
            findNJUNSCode = UCase(cell.offset(0, 1).Value)
            Exit Function
        End If
    Next cell
    findNJUNSCode = ""
End Function

Public Sub ExportUpdateAllNJUNS()
    Dim copiedCode As String
    
    Call LogMessage.SendLogMessage("ExportUpdateAllNJUNS")
    If Not checkCodes Then Exit Sub
    
    Dim Project As Project: Set Project = New Project
    Call Project.extractFromSheets
    
    Dim njunsSheets As scripting.Dictionary: Set njunsSheets = New scripting.Dictionary
    njunsSheets.Add "NOTIFY", New Collection
    njunsSheets.Add "CA", New Collection
    njunsSheets.Add "PT", New Collection
    
    Dim pole As pole
    Dim sheet As Worksheet
    For Each pole In Project.poles
        If Utilities.OnlyNumbers(pole.njunsTicket) <> -1 Then
            If InStr(pole.njunsTicket, "NOTIFY") > 0 Then njunsSheets("NOTIFY").Add pole
            If InStr(pole.njunsTicket, "CA") > 0 Then njunsSheets("CA").Add pole
            If InStr(pole.njunsTicket, "PT") > 0 Then njunsSheets("PT").Add pole
        End If
    Next pole
    
    Dim i As Integer
    Dim totalNJUNS As Integer
    totalNJUNS = njunsSheets("NOTIFY").count + njunsSheets("CA").count + njunsSheets("PT").count
    For Each ticketType In njunsSheets
        For Each pole In njunsSheets(ticketType)
            i = i + 1
            copiedCode = copiedCode & getSelectProjectTabCode(Project) & getUpdateSheetNJUNSCode(Project, pole)
            If copiedCode <> "" Then copiedCode = copiedCode & vbLf & "console.log('" & i & "/" & totalNJUNS & "')"
        Next pole
    Next ticketType
    
    If copiedCode <> "" Then
        copiedCode = wrapCode(Project, copiedCode)
        
        Dim DataObj As DataObject: Set DataObj = New DataObject
        DataObj.SetText copiedCode
        DataObj.PutInClipboard
        
        MsgBox ("Copied code to clipboard for all tickets that already have ticket numbers on their sheet, remove ticket numbers you don't wish to update. go to NJUNS website and press f12 to paste code into console on the project tab for the project you wish to add these tickets to.")
    Else
        MsgBox ("No existing tickets on sheets.")
    End If
End Sub

Private Function getUpdateSheetNJUNSCode(Project As Project, pole As pole) As String
    Dim copiedCode As String
    
    Dim controlWs As Worksheet: Set controlWs = ThisWorkbook.sheets("Control")
    If controlWs.Range("NJUNSCODES").Value = "" Then Call NJUNSCodes.generateNJUNSCodes
    Dim codeRange As Range: Set codeRange = controlWs.Range("NJUNSCODES").EntireColumn
    lastRow = controlWs.Cells(controlWs.Rows.count, controlWs.Range("NJUNSCODES").Column).End(xlUp).row
    Set codeRange = controlWs.Range(controlWs.Cells(3, controlWs.Range("NJUNSCODES").Column), controlWs.Cells(lastRow, controlWs.Range("NJUNSCODES").Column))
    
    If pole.njunsSteps.count = 0 Then
        Exit Function
    End If
    
    Set previousSteps = Nothing
    
    Dim ticketType As String
    If InStr(pole.njunsTicket, "CA") = 1 Then
        ticketType = "Violation (VIO)"
        NJUNSType = "CA"
    ElseIf InStr(pole.njunsTicket, "NOTIFY") = 1 Then
        ticketType = "Violation (VIO)"
        NJUNSType = "NOTIFY"
    ElseIf InStr(pole.njunsTicket, "PT") = 1 Then
        ticketType = "Pole Transfer (PT)"
        NJUNSType = "PT"
    Else
        MsgBox ("NJUNS ticket type needs to be CA/PT/NOTIFY")
        Exit Function
    End If
    
    Dim ticketNumber As String
    ticketNumber = Utilities.OnlyNumbers(pole.njunsTicket)
    
    copiedCode = copiedCode & vbLf & "exists = false;"
    copiedCode = copiedCode & vbLf & "for (const el of document.querySelectorAll(""tbody"")[3].childNodes) {"
    copiedCode = copiedCode & vbLf & "if (/^" & ticketNumber & "/.test(el.textContent.trim())) {"
    copiedCode = copiedCode & vbLf & "exists = true;"
    copiedCode = copiedCode & vbLf & "console.log(""Skipping ticket for pole " & pole.poleNumber & " / " & ticketNumber & " because ticket already exists in project."");"
    copiedCode = copiedCode & vbLf & "break;"
    copiedCode = copiedCode & vbLf & "}};"
    
    copiedCode = copiedCode & vbLf & "if (!exists) {"
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    copiedCode = copiedCode & vbLf & "await clickButton("".v-button.v-widget.icon"", ""Add"");"
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    copiedCode = copiedCode & vbLf & "await realClick(document.querySelector("".c-groupbox-expander""));"
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    copiedCode = copiedCode & vbLf & "await realClick (document.querySelector("".v-button.v-widget.icon-only.v-button-icon-only.v-popupbutton""));"
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    copiedCode = copiedCode & vbLf & "await clickButton("".c-cm-button.v-widget.v-has-width"", ""Ticket/Pole Number Search"");"
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelectorAll("".v-textfield.v-widget.param-field.v-textfield-param-field.v-has-width"")[0],""" & ticketNumber & """);"
    copiedCode = copiedCode & vbLf & "await clickButton("".v-button.v-widget.filter-search-button.v-button-filter-search-button.icon"", ""Search"");"
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    
    copiedCode = copiedCode & vbLf & "if(document.querySelectorAll("".v-table-row"")[0]) {"
    copiedCode = copiedCode & vbLf & "if(!['Open','Draft'].includes(document.querySelectorAll('.v-table-row')[0].childNodes[4].textContent)) {"
    copiedCode = copiedCode & vbLf & "console.log(""Skipping ticket for pole " & pole.poleNumber & " / " & ticketNumber & " because ticket isn't in open or draft status"");"
    copiedCode = copiedCode & vbLf & "break;"
    copiedCode = copiedCode & vbLf & "}"
    copiedCode = copiedCode & vbLf & "await realClick(document.querySelectorAll("".v-table-row"")[0]);"
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    copiedCode = copiedCode & vbLf & "await clickButton("".v-button.v-widget.icon.c-window-action-button.v-button-c-window-action-button.c-primary-action.v-button-c-primary-action"", ""Select"");"
    
    copiedCode = copiedCode & vbLf & "await waitLoadingTime();"
    copiedCode = copiedCode & vbLf & "for (const el of document.querySelectorAll(""tbody"")[3].childNodes) {"
    copiedCode = copiedCode & vbLf & "if (/^" & ticketNumber & "/.test(el.textContent.trim())) {"
    copiedCode = copiedCode & vbLf & "for (i = 0; i < 10; i ++) {"
    copiedCode = copiedCode & vbLf & "el.scrollIntoView();"
    copiedCode = copiedCode & vbLf & "await realClick(el);"
    copiedCode = copiedCode & vbLf & "el.dispatchEvent(new KeyboardEvent(""keydown"", {key: ""Enter"",code: ""Enter"",keyCode: 13,bubbles: true}));"
    copiedCode = copiedCode & vbLf & "await waitLoadingTime();"
    copiedCode = copiedCode & vbLf & "if (document.querySelectorAll("".v-textfield.v-textfield-readonly.c-disabled-or-readonly.v-widget.v-readonly.v-has-width"")[1]) break;"
    copiedCode = copiedCode & vbLf & "}"
    copiedCode = copiedCode & vbLf & "break;"
    copiedCode = copiedCode & vbLf & "}};"
    
    njunsRemarks = "Updated " & Format(Date, "mm/dd/yyyy") & "\nNotification: " & Project.Notification & " Permit: " & Project.permit & "\n"
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelectorAll("".v-textfield.v-widget.v-has-width"")[8], ""Holly Webb"");"
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelectorAll("".v-textfield.v-widget.v-has-width"")[9], ""517-788-1690"");"
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelectorAll("".v-textfield.v-widget.v-has-width"")[10],  ""holly.webb@cmsenergy.com"");"
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelectorAll("".v-textfield.v-widget.v-has-width"")[12],  """ & Project.Notification & """);"
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelector("".v-textarea.v-widget.v-has-width""), """ & njunsRemarks & """ + document.querySelector("".v-textarea.v-widget.v-has-width"").value);"
    copiedCode = copiedCode & vbLf & "await selectDropDownOption(document.querySelector("".v-filterselect-input""), ""3"");"
    
    
    copiedCode = copiedCode & vbLf & "for (i = 0; i < document.querySelectorAll('.v-grid-row.v-grid-row-has-data').length;) {"
    copiedCode = copiedCode & vbLf & "realClick(document.querySelectorAll('.v-grid-row.v-grid-row-has-data')[0],true);"
    copiedCode = copiedCode & vbLf & "if (document.querySelectorAll('.v-grid-row.v-grid-row-has-data').length > 1)  realClick(document.querySelectorAll('.v-grid-row.v-grid-row-has-data')[1],true);"
    copiedCode = copiedCode & vbLf & "if (document.querySelectorAll('.v-grid-row.v-grid-row-has-data').length > 2)  realClick(document.querySelectorAll('.v-grid-row.v-grid-row-has-data')[2],true);"
    copiedCode = copiedCode & vbLf & "if (document.querySelectorAll('.v-grid-row.v-grid-row-has-data').length > 3)  realClick(document.querySelectorAll('.v-grid-row.v-grid-row-has-data')[3],true);"
    copiedCode = copiedCode & vbLf & "if (document.querySelectorAll('.v-grid-row.v-grid-row-has-data').length > 4)  realClick(document.querySelectorAll('.v-grid-row.v-grid-row-has-data')[4],true);"
    copiedCode = copiedCode & vbLf & "if (document.querySelectorAll('.v-grid-row.v-grid-row-has-data').length > 5)  realClick(document.querySelectorAll('.v-grid-row.v-grid-row-has-data')[5],true);"
    copiedCode = copiedCode & vbLf & "await waitLoadingTime();"
    copiedCode = copiedCode & vbLf & "await clickButton("".v-button-caption"", ""Delete"");"
    copiedCode = copiedCode & vbLf & "await clickButton("".v-button.v-widget.c-primary-action.v-button-c-primary-action.icon"", ""OK"");"
    copiedCode = copiedCode & vbLf & "}"
    
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    
    Dim company As String, njunsCode As String, remarks As String, stepType As String
    Dim stepCounter As Integer
    For Each step In pole.njunsSteps
        stepCounter = stepCounter + 1
        company = UCase(Utilities.GetFirstWord(CStr(step)))
        If company = "CE" Then company = "CONSUMERS"
        If company <> "CONSUMERS" Then
            njunsCode = findNJUNSCode(company)
        End If
        remarks = Replace(step, vbLf, "\n")
        remarks = Replace(remarks, """", "\""")
        If InStr(pole.njunsTicket, "CA") = 1 Then
            stepType = "VIOLATION"
        ElseIf InStr(pole.njunsTicket, "NOTIFY") = 1 Then
            stepType = "NOTIFY"
        ElseIf InStr(pole.njunsTicket, "PT") = 1 Then
            stepType = "TRANSFER"
            If stepCounter = 1 Or stepCounter = pole.njunsSteps.count Then
                If company <> "CONSUMERS" Then
                    If stepCounter = 1 Then
                        copiedCode = copiedCode & vbLf & "await generateStep(consumersCode, ""Consumers to complete required work."", ""SET POLE"", """ & Project.Notification & """);"
                   End If
                Else
                    If stepCounter = 1 Then
                        stepType = "SET POLE"
                    Else
                        stepType = "PULL POLE"
                    End If
                End If
            End If
        End If
        
        copiedCode = copiedCode & vbLf & "if(cancelled) break;"
        If company = "CONSUMERS" Then
            copiedCode = copiedCode & vbLf & "await generateStep(consumersCode, """ & remarks & """, """ & stepType & """, """ & Project.Notification & """);"
        Else
            copiedCode = copiedCode & vbLf & "await generateStep(""" & njunsCode & """, """ & remarks & """, """ & stepType & """, """ & Project.Notification & """);"
        End If
    
        If company <> "CONSUMERS" And stepCounter = pole.njunsSteps.count And InStr(pole.njunsTicket, "PT") = 1 Then
            copiedCode = copiedCode & vbLf & "if(cancelled) break;"
            copiedCode = copiedCode & vbLf & "await generateStep(consumersCode, ""Consumers after comms transfer to new pole, pull topped pole."", ""PULL POLE"", """ & Project.Notification & """);"
            copiedCode = copiedCode & vbLf & "if(cancelled) break;"
        End If
    Next step
    
    copiedCode = copiedCode & vbLf & "await clickButton("".v-captiontext"", ""Poles/Assets"");"
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    copiedCode = copiedCode & vbLf & "await realClick(document.querySelector('.v-grid-row.v-grid-row-has-data'));"
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    copiedCode = copiedCode & vbLf & "await clickButton('.v-button.v-widget.icon', 'Edit');"
    copiedCode = copiedCode & vbLf & "await clickButton('.v-button.v-widget.icon', 'Disable Geocoding');"
    
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelector("".v-textfield.v-textfield-large.bold""), """ & pole.existingCEID & """);"
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelectorAll("".v-textfield.v-widget.v-has-width"")[8], """ & pole.latitude & """);"
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelectorAll("".v-textfield.v-widget.v-has-width"")[9], """ & pole.longitude & """);"
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelectorAll("".v-textfield.v-widget.v-has-width"")[10], """");"
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelectorAll("".v-textfield.v-widget.v-has-width"")[11], """");"
    copiedCode = copiedCode & vbLf & "setTextFieldValue(document.querySelectorAll("".v-textfield.v-widget.v-has-width"")[13], """");"
    
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    copiedCode = copiedCode & vbLf & "await clickButton("".v-button.v-widget.icon"", ""Enable Geocoding"");"
    copiedCode = copiedCode & vbLf & "await waitLoadingTime();"
    copiedCode = copiedCode & vbLf & "await clickButton("".v-button.v-widget.icon"", ""Enable Geocoding"");"
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    copiedCode = copiedCode & vbLf & "await clickButton("".v-button.v-widget.primary"", ""Create"");"
    copiedCode = copiedCode & vbLf & "await clickButton("".v-button.v-widget.primary"", ""Apply Changes"");"
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    
    If save Then
        copiedCode = copiedCode & vbLf & "if(cancelled) break;"
        copiedCode = copiedCode & vbLf & "await clickButton("".v-button.v-widget.blueIcon.v-button-blueIcon.icon"", ""Save"");"
        copiedCode = copiedCode & vbLf & "ticketNumber = document.querySelectorAll("".v-textfield.v-textfield"")[1].value;"
        
        copiedCode = copiedCode & vbLf & "if(cancelled) break;"
        
        If ticketType = "Violation (VIO)" Then
            copiedCode = copiedCode & vbLf & "if(document.querySelector("".c-groupbox-caption-text"").textContent.includes(""PT" & ticketNumber & """)) {"
            copiedCode = copiedCode & vbLf & "await clickButton('.v-button.v-widget.v-popupbutton.borderless.v-button-borderless.noIndicator.v-button-noIndicator.icon.v-button-icon','Actions');"
            copiedCode = copiedCode & vbLf & "if(cancelled) break;"
            copiedCode = copiedCode & vbLf & "await clickButton('.c-cm-button.v-widget.v-has-width','Change Ticket Type');"
            copiedCode = copiedCode & vbLf & "if(cancelled) break;"
            copiedCode = copiedCode & vbLf & "await realClick(document.querySelectorAll("".v-grid-row.v-grid-row-has-data"")[document.querySelectorAll("".v-grid-row.v-grid-row-has-data"").length - 1])"
            copiedCode = copiedCode & vbLf & "if(cancelled) break;"
            copiedCode = copiedCode & vbLf & "await clickButton("".v-button.v-widget.primary.v-button-primary"", ""OK"");"
            copiedCode = copiedCode & vbLf & "if(cancelled) break;"
            copiedCode = copiedCode & vbLf & "await clickButton("".v-button.v-widget.icon"", ""Yes"");"
            copiedCode = copiedCode & vbLf & "}"
        ElseIf ticketType = "Pole Transfer (PT)" Then
            copiedCode = copiedCode & vbLf & "if(document.querySelector("".c-groupbox-caption-text"").textContent.includes(""VIO" & ticketNumber & """)) {"
            copiedCode = copiedCode & vbLf & "await clickButton('.v-button.v-widget.v-popupbutton.borderless.v-button-borderless.noIndicator.v-button-noIndicator.icon.v-button-icon','Actions');"
            copiedCode = copiedCode & vbLf & "if(cancelled) break;"
            copiedCode = copiedCode & vbLf & "await clickButton('.c-cm-button.v-widget.v-has-width','Change Ticket Type');"
            copiedCode = copiedCode & vbLf & "if(cancelled) break;"
            copiedCode = copiedCode & vbLf & "await realClick(document.querySelectorAll("".v-grid-row.v-grid-row-has-data"")[document.querySelectorAll("".v-grid-row.v-grid-row-has-data"").length - 1])"
            copiedCode = copiedCode & vbLf & "if(cancelled) break;"
            copiedCode = copiedCode & vbLf & "await clickButton("".v-button.v-widget.primary.v-button-primary"", ""OK"");"
            copiedCode = copiedCode & vbLf & "if(cancelled) break;"
            copiedCode = copiedCode & vbLf & "await clickButton("".v-button.v-widget.icon"", ""Yes"");"
            copiedCode = copiedCode & vbLf & "}"
        End If
        
        copiedCode = copiedCode & vbLf & "pdfList.push({"
            copiedCode = copiedCode & vbLf & "url: download ? await getPDFURL() : null,"
            copiedCode = copiedCode & vbLf & "filename: `" & Project.Notification & " - " & "${ticketNumber} - " & NJUNSType & " TICKET`,"
            copiedCode = copiedCode & vbLf & "poleNumber: """ & pole.poleNumber & ""","
            copiedCode = copiedCode & vbLf & "ticketNumber: ticketNumber"
        copiedCode = copiedCode & vbLf & "});"
    End If
    
    copiedCode = copiedCode & vbLf & "} else {"
    copiedCode = copiedCode & vbLf & "await realClick(document.querySelectorAll("".v-button.v-widget.link.v-button-link"")[0]);"
    copiedCode = copiedCode & vbLf & "}"
    copiedCode = copiedCode & vbLf & "}"
    
    getUpdateSheetNJUNSCode = copiedCode
End Function

Public Sub DownloadAllTickets()
    Dim copiedCode As String
    
    Call LogMessage.SendLogMessage("DownloadAllTickets")
    
    Dim Project As Project: Set Project = New Project
    Call Project.extractFromSheets
    
    idDictString = "const idDict = {"
    latLongDictString = "const latLongDict = {"
    Dim pole As pole
    For Each pole In Project.poles
        idDictString = idDictString & vbLf & "['" & pole.existingCEID & "']: '" & pole.poleNumber & "',"
        latLongDictString = latLongDictString & vbLf & "['" & Application.WorksheetFunction.RoundUp(pole.latitude, 6) & Application.WorksheetFunction.RoundUp(pole.longitude, 6) & "']: '" & pole.poleNumber & "',"
    Next pole
    If InStr(idDictString, ",") > 0 Then idDictString = Left(idDictString, Len(idDictString) - 1)
    If InStr(latLongDictString, ",") > 0 Then latLongDictString = Left(latLongDictString, Len(latLongDictString) - 1)
    idDictString = idDictString & vbLf & "};"
    latLongDictString = latLongDictString & vbLf & "};"
    
    copiedCode = copiedCode & "const ticketsDone = {}"
    copiedCode = copiedCode & getSelectProjectTabCode(Project)
    copiedCode = copiedCode & vbLf & "await waitLoadingTime();"
    copiedCode = copiedCode & vbLf & "await waitLoadingTime();"
    copiedCode = copiedCode & vbLf & "totalTickets = " & "Number(document.querySelector('.v-label.v-widget.c-paging-status.v-label-c-paging-status.v-label-undef-w').textContent.replace(/\D/g,''));"
    copiedCode = copiedCode & vbLf & "for (i = 0; i < totalTickets;) {"
    copiedCode = copiedCode & vbLf & "await waitLoadingTime();"
    copiedCode = copiedCode & vbLf & "element = null"
    copiedCode = copiedCode & vbLf & "for (const el of document.querySelectorAll(""tbody"")[3].childNodes) {"
    copiedCode = copiedCode & vbLf & "if(!(el.children[0].textContent in ticketsDone)) {element = el; break;}"
    copiedCode = copiedCode & vbLf & "}"
    copiedCode = copiedCode & vbLf & "for (j = 0; j < 10; j++) {"
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    copiedCode = copiedCode & vbLf & "element.scrollIntoView();"
    copiedCode = copiedCode & vbLf & "await realClick(element);"
    copiedCode = copiedCode & vbLf & "element.dispatchEvent(new KeyboardEvent('keydown', {key: 'Enter',code: 'Enter',keyCode: 13,bubbles: true}));"
    copiedCode = copiedCode & vbLf & "await waitLoadingTime();"
    copiedCode = copiedCode & vbLf & "if (document.querySelectorAll('.v-textfield.v-textfield-readonly.c-disabled-or-readonly.v-widget.v-readonly.v-has-width')[1]) break;"
    copiedCode = copiedCode & vbLf & "}"

    copiedCode = copiedCode & vbLf & "if(cancelled) break;"
    copiedCode = copiedCode & vbLf & "if(!document.querySelectorAll('.v-textfield.v-textfield-readonly.c-disabled-or-readonly.v-widget.v-readonly.v-has-width')[11]) {"
    copiedCode = copiedCode & vbLf & "await clickButton('.v-captiontext', 'Details');"
    copiedCode = copiedCode & vbLf & "}"
    copiedCode = copiedCode & vbLf & "if(cancelled) break;"

    copiedCode = copiedCode & vbLf & "ticketNumber = document.querySelectorAll('.v-textfield.v-textfield')[1].value;"
    copiedCode = copiedCode & vbLf & "ticketsDone[ticketNumber] = true"
    copiedCode = copiedCode & vbLf & "i++"
    copiedCode = copiedCode & vbLf & "if(document.querySelector('.c-groupbox-caption-text').textContent.includes('PT' + ticketNumber)) ticketType = 'PT';"
    copiedCode = copiedCode & vbLf & "else if(document.querySelectorAll('tbody')[3].childNodes[0].childNodes[3].textContent === 'NOTIFY') ticketType = 'NOTIFY';"
    copiedCode = copiedCode & vbLf & "else ticketType = 'CA';"
        
    copiedCode = copiedCode & vbLf & "assetNumber = document.querySelectorAll('.v-textfield.v-textfield-readonly.c-disabled-or-readonly.v-widget.v-readonly.v-has-width')[4].value;"
    copiedCode = copiedCode & vbLf & "latLong = document.querySelectorAll('.v-textfield.v-textfield-readonly.c-disabled-or-readonly.v-widget.v-readonly.v-has-width')[11].value + document.querySelectorAll('.v-textfield.v-textfield-readonly.c-disabled-or-readonly.v-widget.v-readonly.v-has-width')[12].value;"
    copiedCode = copiedCode & vbLf & "poleNumber = idDict[assetNumber] ?? null"
    copiedCode = copiedCode & vbLf & "if(!poleNumber) poleNumber = latLongDictString[latLong] ?? null"
    copiedCode = copiedCode & vbLf & "if(!poleNumber) poleNumber = '??'"

    copiedCode = copiedCode & vbLf & "pdfList.push({"
    copiedCode = copiedCode & vbLf & "url: download ? await getPDFURL() : null,"
    copiedCode = copiedCode & vbLf & "filename: `" & Project.Notification & " - ${ticketNumber} - ${ticketType} TICKET`,"
    copiedCode = copiedCode & vbLf & "poleNumber: poleNumber,"
    copiedCode = copiedCode & vbLf & "ticketNumber: ticketNumber"
    copiedCode = copiedCode & vbLf & "});"
    copiedCode = copiedCode & vbLf & "console.log(`${i}/${totalTickets}`)"
    copiedCode = copiedCode & getSelectProjectTabCode(Project)
    copiedCode = copiedCode & vbLf & "}"
    
    copiedCode = wrapCode(Project, copiedCode)
    copiedCode = idDictString & vbLf & latLongDictString & vbLf & copiedCode
    
    Dim DataObj As DataObject: Set DataObj = New DataObject
    DataObj.SetText copiedCode
    DataObj.PutInClipboard
    
    MsgBox ("Copied code to clipboard to downloadf all tickets. Go to NJUNS website and press f12 to paste code into console on the project tab for the project you wish to download the tickets of.")
End Sub
