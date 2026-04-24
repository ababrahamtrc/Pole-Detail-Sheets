Attribute VB_Name = "QACheck"
Public Sub QACheckPole()
    On Error Resume Next
    Dim sheet As Worksheet

    Call LogMessage.SendLogMessage("QACheckPole")

    Set sheet = ThisWorkbook.ActiveSheet()
    If sheet.name = "4 Spans" Or sheet.name = "8 Spans" Or sheet.name = "12 Spans" Or sheet.Cells(2, 2).Value <> "Notification:" Then
        MsgBox "You need to have a pole detail sheet active to run this script."
        Exit Sub
    End If
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox QACheck(sheet)
End Sub

Public Sub QACheckAllPoles()
    On Error Resume Next
    
    Call LogMessage.SendLogMessage("QACheckAllPoles")
    
    Dim sheet As Worksheet
    Dim filePath As String
    Dim fileContent As String
    Dim fNum As Integer
    
    issues = ""
    For Each sheet In ThisWorkbook.sheets
        If sheet.name <> "4 Spans" And sheet.name <> "8 Spans" And sheet.name <> "12 Spans" And sheet.Cells(2, 2).Value = "Notification:" Then
            poleIssues = QACheck(sheet)
            If InStr(poleIssues, "No issues") = 0 Then issues = issues & QACheck(sheet) & vbLf & vbLf
        End If
    Next sheet
    
    If issues = "" Then issues = "No issues found on any pole :)"
    
    filePath = ThisWorkbook.path & "\QACheck.txt"
    If InStr(filePath, "sharepoint") > 0 Then filePath = Environ("USERPROFILE") & "\Downloads\QACheck.txt"
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    fNum = FreeFile
    Open filePath For Output As #fNum
    Print #fNum, issues
    Close #fNum
    
    Shell "notepad.exe """ & filePath & """", vbNormalFocus
    
End Sub

Private Function QACheck(pds As Worksheet) As String
    Dim issues As String
    Dim warnings As String
    
    On Error Resume Next
    
    If InStr(pds.Range("PERMIT"), "NAR") Then
        jobType = "MKR"
    ElseIf InStr(pds.Range("PERMIT"), "SC") Then
        jobType = "SC"
    Else
        jobType = "SI"
    End If
    
    Dim pole As pole: Set pole = New pole
    Call pole.extractFromSheet(pds)
    
    ' Checking basic empty cells
    Call checkEmptyCell(pds, issues, "NOTIFICATION", "Notification")
    Call checkEmptyCell(pds, issues, "PERMIT", "Permit")
    Call checkEmptyCell(pds, issues, "APPLICANT", "Applicant")
    Call checkEmptyCell(pds, issues, "INSPECTEDBY", "Inspected By")
    Call checkEmptyCell(pds, issues, "TWP", "Town")
    Call checkEmptyCell(pds, issues, "DATE", "Date")
    Call checkEmptyCell(pds, issues, "COUNTY", "County")
    Call checkEmptyCell(pds, issues, "POLENUM", "Pole Number")
    Call checkEmptyCell(pds, issues, "CLASS", "Pole Class")
    Call checkEmptyCell(pds, issues, "HEIGHT", "Pole Height")
    Call checkEmptyCell(pds, issues, "TYPE", "Pole Type")
    Call checkEmptyCell(pds, issues, "CLASSESTIMATE", "Pole Class Estimate")
    Call checkEmptyCell(pds, issues, "GLC", "GLC")
    Call checkEmptyCell(pds, issues, "HEIGHTESTIMATE", "Pole Height Estimate")
    Call checkEmptyCell(pds, issues, "CEID", "CEID")
    If isEmpty(pds.Range("CEPOLE")) And isEmpty(pds.Range("OTHERPOLE")) Then issues = issues & "• CE Pole or Other pole owner needs to be checked off at the top" & vbLf
    If Not isEmpty(pds.Range("OTHERPOLE")) Then Call checkEmptyCell(pds, issues, "OTHERPOLEOWNER", "Other Pole Owner")
    If isEmpty(pds.Range("TANGENT")) And isEmpty(pds.Range("ANGLE")) And isEmpty(pds.Range("CORNER")) And isEmpty(pds.Range("DEADEND")) And isEmpty(pds.Range("BUCK")) And isEmpty(pds.Range("SECONLY")) Then issues = issues & "• Need to select a framing type" & vbLf
    If (Not isEmpty(pds.Range("TANGENT")) Or Not isEmpty(pds.Range("ANGLE")) Or Not isEmpty(pds.Range("CORNER")) Or Not isEmpty(pds.Range("DEADEND")) Or Not isEmpty(pds.Range("BUCK"))) And Not isEmpty(pds.Range("SECONLY")) Then issues = issues & "• Sec only can't be selected if there are other framing types selected" & vbLf
    If Not isEmpty(ThisWorkbook.sheets("Control").Range("QAFU")) Then
        If isEmpty(pds.Range("SECONLY")) And (isEmpty(pds.Range("TANGENT").offset(0, 1)) And isEmpty(pds.Range("ANGLE").offset(0, 1)) And isEmpty(pds.Range("CORNER").offset(0, 1)) And isEmpty(pds.Range("DEADEND").offset(0, 1)) And isEmpty(pds.Range("BUCK").offset(0, 1))) Then issues = issues & "• Framing unit not specified" & vbLf
    End If
    'Applicant checks
    Call checkEmptyCell(pds, issues, "PROPOSEDHEIGHT", "ProposedHeight")
    If Not isEmpty(pds.Range("PROPOSEDHEIGHT")) And InStr(UCase(pds.Range("PROPOSEDHEIGHT")), "OL") = 0 And InStr(pds.Range("PROPOSEDHEIGHT"), "'") = 0 Then issues = issues & "• Missing foot symbol in proposed attach height" & vbLf
    If Not isEmpty(pds.Range("EXISTINGDIAMETER")) And InStr(UCase(pds.Range("PROPOSEDHEIGHT")), "OL") = 0 Then issues = issues & "• if there's an existing diameter, OL should appear in the proposed attach height" & vbLf
    
    ' Back Sheet Checks
    If Not isEmpty(pds.Range("DL")) Then
        Call checkEmptyCell(pds, issues, "ALTONE", "Alt 1")
    Else
        If Not isEmpty(pds.Range("SUMSHEET3")) Or Not isEmpty(pds.Range("SUMSHEET6")) Then Call checkEmptyCell(pds, issues, "ALTONE", "Alt 1")
    End If
    If Not isEmpty(pds.Range("SUMSHEET3")) Then Call checkEmptyCell(pds, issues, "ALTTWO", "Alt 2")
    If Not isEmpty(pds.Range("SUMSHEET6")) Then Call checkEmptyCell(pds, issues, "ALTTHREE", "Alt 3")
    If Not isEmpty(pds.Range("SUMSHEET4")) Then Call checkEmptyCell(pds, issues, "NJUNS", "NJUNS")
    If Not isEmpty(pds.Range("NJUNSTICKET")) Then Call checkEmptyCell(pds, issues, "NJUNS", "NJUNS")
    If Not isEmpty(pds.Range("SUMSHEET5")) Then Call checkEmptyCell(pds, issues, "MAINT", "CE Maintenance")
    If pds.Range("SUMSHEET11").text = "1‚2" And Replace(pds.Range("ALTTWO").text, " ", "") <> "SAMEASALT1" Then issues = issues & "• Alt 2 should be labeled ""SAME AS ALT 1"" if there's only alt 1 and 2 work" & vbLf
    If pds.Range("SUMSHEET11").text = "1‚3" And Replace(pds.Range("ALTTHREE").text, " ", "") <> "SAMEASALT1" Then issues = issues & "• Alt 3 should be labeled ""SAME AS ALT 1"" if there's only alt 1 and 3 work" & vbLf
    If Not isEmpty(ThisWorkbook.sheets("Control").Range("QACMRW")) Then
        If Not isEmpty(pds.Range("SUMSHEET7")) And Replace(pds.Range("NJUNS").text, " ", "") <> "COMMMAKEREADYWORK" Then issues = issues & "• NJUNS should be labeled ""COMM MAKE READY WORK"" if there's comm make ready work" & vbLf
    End If
    
    ' Pole Denied
    If Not isEmpty(pds.Range("SUMSHEET8")) Then
        Call checkEmptyCell(pds, issues, "SUMSHEET14", "Comments - Applicant")
        If pds.Range("CMRF1").Value <> "DENIED" Or pds.Range("CMRF2").Value <> "DENIED" Or pds.Range("CMRF3").Value <> "DENIED" Then issues = issues & "• If pole is denied, put DENIED in applicant CMRF sections." & vbLf
        If Not isEmpty(pds.Range("SUMSHEET1")) Then issues = issues & "• If pole is denied, ok to attach before work shouldn't be checked off" & vbLf
        If Not isEmpty(pds.Range("SUMSHEET2")) Then issues = issues & "• If pole is denied, work required before attach shouldn't be checked off" & vbLf
        If isEmpty(pds.Range("NEWAPP")) Then issues = issues & "• Even if pole is denied, there should be a new app percentage. Put DENIED if no percentage" & vbLf
    ' Pole Not Denied
    Else
        If isEmpty(pds.Range("SUMSHEET9")) Then Call checkEmptyCell(pds, issues, "NEWAPP", "New App Loading %")
        Call checkEmptyCell(pds, issues, "CMRF1", "New Attacher Height")
    End If
    
    ' Foreign Pole
    If Not isEmpty(pds.Range("SUMSHEET9")) Then
        If pds.Range("CEID") <> "FOREIGN" Then issues = issues & "• CEID should be set to FOREIGN if it's a foreign pole" & vbLf
        If isEmpty(pds.Range("SUMSHEET8")) And (InStr(1, pds.Range("CMRF1"), "APPLY TO", vbTextCompare) = 0 Or InStr(1, pds.Range("CMRF2"), "APPLY TO", vbTextCompare) = 0 Or InStr(1, pds.Range("CMRF3"), "APPLY TO", vbTextCompare) = 0) Then issues = issues & "• If foreign pole, need to say APPLY TO [OWNER] on cmrf new attacher sections" & vbLf
        If isEmpty(pds.Range("SUMSHEET8")) And InStr(1, pds.Range("SUMSHEET14"), "APPLY TO", vbTextCompare) = 0 Then issues = issues & "• If foreign pole, need to say APPLY TO [OWNER] on comments section of summary sheet" & vbLf
        If InStr(pds.Range("ASIS"), "FOREIGN") = 0 Then issues = issues & "• As-is percentage should be FOREIGN on foreign poles" & vbLf
        If InStr(pds.Range("NEWAPP"), "FOREIGN") = 0 Then issues = issues & "• New App percentage should be FOREIGN on foreign poles" & vbLf
    Else
        Call checkEmptyCell(pds, issues, "ASIS", "As-is Percentage")
        Call checkEmptyCell(pds, issues, "ASISPF", "As-Is Pole Loading Pass/Fail")
        If isEmpty(pds.Range("SUMSHEET8")) Then Call checkEmptyCell(pds, issues, "NEWAPPPF", "With-App Pole Loading Pass/Fail")
        If Not isEmpty(pds.Range("NEWAPPLEAD")) Then
            Call checkEmptyCell(pds, issues, "ROOMTOGUY", "Room to guy")
            Call checkEmptyCell(pds, issues, "PGUY", "Top proposed guying ok")
        End If
        If Not isEmpty(pds.Range("NEWAPPLEAD").offset(1, 0)) Then Call checkEmptyCell(pds, issues, "PGUY2", "Bottom proposed guying ok")
        If pds.Range("ROOMTOGUY") = "NO" Then Call checkEmptyCell(pds, issues, "GUYCOMMENT", "Guying Comment")
        If Not isEmpty(pds.Range("SUMSHEET8")) And Not isEmpty(pds.Range("NEWAPPLEAD")) Then
            If pds.Range("PGUY") = "YES" Then issues = issues & "• Top proposed guying shouldn't be YES if pole is denied." & vbLf
            If Not isEmpty(pds.Range("NEWAPPLEAD").offset(1, 0)) Then
                If pds.Range("PGUY2") = "YES" Then issues = issues & "• Bottom proposed guying shouldn't be YES if pole is denied." & vbLf
            End If
        End If
        If Not isEmpty(pds.Range("NJUNS")) And isEmpty(pds.Range("SUMSHEET4")) And InStr(pds.Range("NJUNS"), "COMM MAKE READY WORK") = 0 Then Call checkEmptyCell(pds, issues, "SUMSHEET4", "Comm to correct violations")
        If isEmpty(pds.Range("SUMSHEET8")) And (Not isEmpty(pds.Range("ALTONE")) Or Not isEmpty(pds.Range("ALTTWO")) Or Not isEmpty(pds.Range("ALTTHREE")) Or Not isEmpty(pds.Range("NJUNS")) Or Not isEmpty(pds.Range("MAINT"))) Then
            Call checkEmptyCell(pds, issues, "SUMSHEET2", "Work Required Before Attach") '
        ElseIf isEmpty(pds.Range("SUMSHEET8")) And (isEmpty(pds.Range("SUMSHEET2")) And (Not isEmpty(pds.Range("SUMSHEET3")) Or Not isEmpty(pds.Range("SUMSHEET4")) Or Not isEmpty(pds.Range("SUMSHEET5")) Or Not isEmpty(pds.Range("SUMSHEET6")) Or Not isEmpty(pds.Range("SUMSHEET7")))) Then
            Call checkEmptyCell(pds, issues, "SUMSHEET2", "Work Required Before Attach")
        End If
        If Not isEmpty(pds.Range("STLTBRKT")) Then
            Call checkEmptyCell(pds, issues, "BONDED", "Streetlight Bonded")
            If isEmpty(pds.Range("MBSM").offset(0, 1)) Then issues = issues & "• Miss/Brkn Streetlight molding needs to be selected if there's a streetlight." & vbLf
        End If
    End If
    
    ' CMRF Checks
    If Not isEmpty(pds.Range("NJUNS")) And InStr(pds.Range("NJUNS"), "COMM MAKE READY WORK") = 0 Then Call checkEmptyCell(pds, issues, "NJUNSTICKET", "NJUNS Ticket")
    If Not isEmpty(pds.Range("CMRF1")) Then Call checkEmptyCell(pds, issues, "CMRF2", "New Attacher Midspans")
    If Not isEmpty(pds.Range("NEWAPPLEAD")) And pds.Range("ROOMTOGUY").Value = "YES" Then Call checkEmptyCell(pds, issues, "CMRF3", "New Attacher Down Guy")
    If InStr(pds.Range("NJUNS"), "COMM MAKE READY WORK") > 0 And Not isEmpty(pds.Range("NJUNSTICKET")) Then issues = issues & "• If it's comm make ready work, there shouldn't be an NJUNS ticket" & vbLf
    
    ' Summary Sheet checks
    If Not isEmpty(pds.Range("SUMSHEET1")) And Not isEmpty(pds.Range("SUMSHEET2")) Then issues = issues & "• Can't have Okay to attach and Work required both checked off" & vbLf
    If Not isEmpty(pds.Range("SUMSHEET2")) And (isEmpty(pds.Range("SUMSHEET3")) And isEmpty(pds.Range("SUMSHEET4")) And isEmpty(pds.Range("SUMSHEET5")) And isEmpty(pds.Range("SUMSHEET6")) And isEmpty(pds.Range("SUMSHEET7"))) Then issues = issues & "• Work required is checked off, requires other checkbox to describe type of work" & vbLf
    If isEmpty(pds.Range("SUMSHEET1")) And isEmpty(pds.Range("SUMSHEET2")) And isEmpty(pds.Range("SUMSHEET8")) And isEmpty(pds.Range("SUMSHEET9")) Then issues = issues & "• Missing ok to attach, work required, denied, or foreign from summary sheet" & vbLf

    ' Summary sheet should be checked
    If Not isEmpty(pds.Range("ALTTWO")) Then Call checkEmptyCell(pds, issues, "SUMSHEET3", "CE TO CORRECT VIOLATIONS")
    
    If Not isEmpty(pds.Range("ALTTHREE")) Then Call checkEmptyCell(pds, issues, "SUMSHEET6", "CE MAKE READY WORK REQUIRED")
    If Not isEmpty(pds.Range("NJUNS")) And isEmpty(pds.Range("SUMSHEET7")) And InStr(pds.Range("NJUNS"), "COMM MAKE READY WORK") > 0 Then Call checkEmptyCell(pds, issues, "SUMSHEET8", "Comm make ready work")
    If Not isEmpty(pds.Range("SUMSHEET4")) And Not isEmpty(pds.Range("SUMSHEET7")) Then issues = issues & "• Comm to correct violations and Comm make ready work shouldn't both be checked off" & vbLf
    
    If pds.Range("CMRF1") = "DENIED" Then Call checkEmptyCell(pds, issues, "SUMSHEET8", "ATTACHMENT DENIED")
    If pds.Range("CEID") = "FOREIGN" Or Not isEmpty(pds.Range("OTHERPOLE")) Then Call checkEmptyCell(pds, issues, "SUMSHEET9", "FOREIGN POLE")
    If Not isEmpty(pds.Range("NEWAPPLEAD").offset(1, 0)) Then Call checkEmptyCell(pds, issues, "CMRF3", "New Attacher Down Guy")
    
    If Not isEmpty(pds.Range("ALTONE")) And Not isEmpty(pds.Range("ALTTWO")) And Not isEmpty(pds.Range("ALTTHREE")) Then
        If Replace(Trim(pds.Range("SUMSHEET11").Value), Chr(130), Chr(44)) <> "1" & Chr(44) & "2" & Chr(44) & "3" Then issues = issues & "• Alt work doesn't match alts done on summary sheet, should be 1,2,3" & vbLf
    ElseIf Not isEmpty(pds.Range("ALTONE")) And Not isEmpty(pds.Range("ALTTWO")) Then
        If Replace(Trim(pds.Range("SUMSHEET11").Value), Chr(130), Chr(44)) <> "1" & Chr(44) & "2" Then issues = issues & "• Alt work doesn't match alts done on summary sheet, should be 1,2" & vbLf
    ElseIf Not isEmpty(pds.Range("ALTONE")) And Not isEmpty(pds.Range("ALTTHREE")) Then
        If Replace(Trim(pds.Range("SUMSHEET11").Value), Chr(130), Chr(44)) <> "1" & Chr(44) & "3" Then issues = issues & "• Alt work doesn't match alts done on summary sheet, should be 1,3" & vbLf
    ElseIf Not isEmpty(pds.Range("ALTONE")) Then
        issues = issues & "• Alt work doesn't match alts done on summary sheet" & vbLf
    Else
        If Not isEmpty(pds.Range("SUMSHEET11")) Then issues = issues & "• Alt work doesn't match alts done on summary sheet, should be N/A" & vbLf
    End If
    If Not isEmpty(pds.Range("MAINT")) And isEmpty(pds.Range("SUMSHEET5")) Then issues = issues & "• CE Maintenance Work Required should be checked off if there's maintenance work." & vbLf
    
    
    If pds.Range("SUMSHEET12") = "-" Then issues = issues & "• Pole loading done not selected" & vbLf
    If Not isEmpty(pds.Range("INSTALLANCHOR")) Or Not isEmpty(pds.Range("INSTALLPOLE")) Then Call checkEmptyCell(pds, issues, "SUMSHEET13", "Pole Staking Done")
    If isEmpty(pds.Range("SUMSHEET13")) And Not isEmpty(pds.Range("REPLACEANCHOR")) And isEmpty(pds.Range("INSTALLANCHOR")) And isEmpty(pds.Range("INSTALLPOLE")) Then warnings = warnings & "Warning: If replaced anchor is a new lead then pole staking complete needs to be checked off, ignore this warning if the anchor isn't moving." & vbLf
    If InStr(pds.Range("NJUNSTICKET"), "PT") > 0 And InStr(pds.Range("ALTONE"), "TOP POLE") = 0 And Not isEmpty(pds.Range("ALTONE")) Then issues = issues & "• There needs to be a top pole note if it's a PT ticket" & vbLf
    If InStr(pds.Range("NJUNSTICKET"), "PT") = 0 And InStr(pds.Range("ALTONE"), "TOP POLE") > 0 Then issues = issues & "• If there's a top pole note, it should be a PT ticket" & vbLf
    
    ' Summary sheet should not be checked
    If Not isEmpty(pds.Range("SUMSHEET1")) And (Not isEmpty(pds.Range("ALTONE")) Or Not isEmpty(pds.Range("ALTTWO")) Or Not isEmpty(pds.Range("ALTTHREE")) Or Not isEmpty(pds.Range("NJUNS"))) Then issues = issues & "• Ok to attach shouldn't be checked off if there's work being done" & vbLf
    If Not isEmpty(pds.Range("SUMSHEET3")) And isEmpty(pds.Range("ALTTWO")) Then issues = issues & "• CE to correct violations shouldn't be checked off if there's no alt 2 work" & vbLf
    If Not isEmpty(pds.Range("SUMSHEET4")) And isEmpty(pds.Range("NJUNS")) Then issues = issues & "• Comm to correct violations shouldn't be checked off if there's no NJUNS" & vbLf
    If Not isEmpty(pds.Range("SUMSHEET5")) And isEmpty(pds.Range("MAINT")) Then issues = issues & "• CE maintanence shouldn't be checked off if there's no maintanance work" & vbLf
    If Not isEmpty(pds.Range("SUMSHEET6")) And isEmpty(pds.Range("ALTTHREE")) Then issues = issues & "• CE make ready work shouldn't be checked off if there's no alt 3 work" & vbLf

    If isEmpty(pds.Range("REPLACEANCHOR")) And isEmpty(pds.Range("INSTALLANCHOR")) And isEmpty(pds.Range("INSTALLPOLE")) And Not isEmpty(pds.Range("SUMSHEET13")) Then issues = issues & "• Pole staking should not be checked off if there's no anchor work or new pole install." & vbLf
    
    paapms = False
    Dim foundOwner As Boolean: foundOwner = False
    If pds.Range("PGUY") = "NO" Or pds.Range("PGUY2") = "NO" Then paapms = True
    If InStr(pds.Range("PROPOSEDHEIGHT"), "OL") = 0 Then
        If pds.Range("CMRF1") <> "" And pds.Range("PROPOSEDHEIGHT") <> "" And InStr(pds.Range("CMRF1"), pds.Range("PROPOSEDHEIGHT")) = 0 Then paapms = True
    Else
        If InStr(pds.Range("CMRF1"), "OL") = 0 And isEmpty(pds.Range("SUMSHEET8")) And isEmpty(pds.Range("SUMSHEET9")) Then issues = issues & "• New attacher height should say ""OL @ XX'XX"" if overlash job" & vbLf
        For i = 0 To 100
            If pds.Range("CMOWNER").offset(i, 0).Interior.color <> 16312794 Then Exit For
            If UCase(pds.Range("CMOWNER").offset(i, 0)) = UCase(pds.Range("APPLICANT")) Then
                foundOwner = True
                If InStr(pds.Range("CMOWNER").offset(i, 0), "DG") = 0 Then
                    If pds.Range("CMRF1") <> "" And InStr(pds.Range("CMRF1"), pds.Range("CMHEIGHT").offset(i, 0)) = 0 Then
                        paapms = True
                        Exit For
                    End If
                End If
            End If
        Next i
        If foundOwner = False Then issues = issues & "• Overlash job but applicant doesn't have their name on the owner section of any of the comms" & vbLf
    End If
    
    If pole.applicant.modification > 0 And pole.applicant.height <> pole.applicant.modification Then paapms = True
    
    If Not paapms Then
        For Each midspan In pole.applicant.midspans
            displayMidspan = Utilities.inchesToFeetInches(pole.applicant.midspans(midspan))
            otherPoleNumber = ThisWorkbook.RemoveParentheses(pds.Range("TOPOLE" & midspan))
            lines = Split(pds.Range("CMRF2"), vbLf)
            For Each line In lines
                If InStr(line, "@P" & otherPoleNumber) > 0 And InStr(Replace(line, " ", ""), displayMidspan) = 0 Then
                    paapms = True
                    Exit For
                End If
                If InStr(line, "(" & otherPoleNumber & ")") > 0 And InStr(Replace(line, " ", ""), displayMidspan) = 0 Then
                    paapms = True
                    Exit For
                End If
            Next line
        Next midspan
    End If
    
    commaMidspan = False
    guessMidspan = False
    missingInchOrFoot = False
    connectingWarning = False
    namidSpan = False
    
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Pattern = "^(.*?)(?=\()"
        .IgnoreCase = True
        .Global = False
    End With
    For i = 1 To 12
        For Each name In pds.names
            If name.name = "'" & pds.name & "'" & "!TOPOLE" & i Then
                If Trim(Replace(pds.Range("TOPOLE" & i), "-", "")) <> "" Then
                    If Trim(Replace(pds.Range("TOPOLE" & i).offset(1, 0), "-", "")) <> "" Then
                        If InStr(pds.Range("TOPOLE" & i).offset(1, 0), ",") Then commaMidspan = True
                        If InStr(pds.Range("TOPOLE" & i).offset(1, 0), "GUESS") Then guessMidspan = True
                        If InStr(pds.Range("TOPOLE" & i).offset(1, 0), "'") = 0 Or InStr(pds.Range("TOPOLE" & i).offset(1, 0), """") = 0 And pds.Range("TOPOLE" & i).offset(1, 0) <> "N/A" Then missingInchOrFoot = True
                        If UCase(pds.Range("TOPOLE" & i).offset(1, 0)) = "N/A" Or UCase(pds.Range("TOPOLE" & i).offset(1, 0)) = "NA" Then namidSpan = True
                    End If
                    If regex.test(pds.Range("TOPOLE" & i)) Then
                        Set matches = regex.Execute(pds.Range("TOPOLE" & i))
                        otherSheetName = matches(0)
                        If SheetExists(otherSheetName) Then
                            For Each oSheet In ThisWorkbook.sheets
                                If ThisWorkbook.RemoveParentheses(oSheet.name) = otherSheetName Then
                                    Set otherSheet = oSheet
                                    Exit For
                                End If
                            Next oSheet
                            
                            For j = 1 To 12
                                For Each otherName In otherSheet.names
                                    If otherName.name = "'" & otherSheet.name & "'" & "!TOPOLE" & j Then
                                        If regex.test(otherSheet.Range("TOPOLE" & j)) Then
                                            Set matches = regex.Execute(otherSheet.Range("TOPOLE" & j))
                                            If matches(0) = pds.Range("POLENUM") Then
                                                angle = Trim(Replace(pds.Range("TOPOLE" & i), "-", ""))
                                                If InStr(angle, "(") > 0 And InStr(angle, ")") > 0 Then angle = Mid(angle, InStr(angle, "(") + 1, InStr(angle, ")") - (InStr(angle, "(") + 1))
                                                
                                                If Trim(pds.Range("TOPOLE" & i).offset(1, 0)) <> Trim(otherSheet.Range("TOPOLE" & j).offset(1, 0)) Then issues = issues & "• Proposed midspan towards pole " & otherSheetName & " doesn't match proposed midspan on sheet " & otherSheetName & vbLf
                                                If Trim(pds.Range("TOPOLE" & i).offset(2, 0)) <> Trim(otherSheet.Range("TOPOLE" & j).offset(2, 0)) Then issues = issues & "• Proposed tension towards pole " & otherSheetName & " doesn't match proposed tension on sheet " & otherSheetName & vbLf
                                                
                                                For k = 0 To 100
                                                    If pds.Range("UTTYPE").offset(k, 0).Interior.color <> 16312794 Then Exit For
                                                    comparingMidspan = Trim(Replace(pds.Range("UTMIDSPAN" & i).offset(k, 0), "-", ""))
                                                    If comparingMidspan <> "" Then
                                                        found = False
                                                        For L = 0 To 100
                                                            If otherSheet.Range("UTTYPE").offset(L, 0).Interior.color <> 16312794 Then Exit For
                                                            comparedMidspan = Trim(Replace(otherSheet.Range("UTMIDSPAN" & j).offset(L, 0), "-", ""))
                                                            If comparingMidspan = comparedMidspan Then
                                                                found = True
                                                                Exit For
                                                            End If
                                                        Next L
                                                        If Not found Then issues = issues & "• Couldn't find matching midspan (" & comparingMidspan & ") towards pole " & otherSheetName & " on sheet " & otherSheetName & vbLf
                                                    End If
                                                Next k
                                                For k = 0 To 100
                                                    If pds.Range("CMOWNER").offset(k, 0).Interior.color <> 16312794 Then Exit For
                                                    comparingMidspan = Trim(Replace(pds.Range("CMMIDSPAN" & i).offset(k, 0), "-", ""))
                                                    If comparingMidspan <> "" Then
                                                        found = False
                                                        For L = 0 To 100
                                                            If otherSheet.Range("CMOWNER").offset(L, 0).Interior.color <> 16312794 Then Exit For
                                                            comparedMidspan = Trim(Replace(otherSheet.Range("CMMIDSPAN" & j).offset(L, 0), "-", ""))
                                                            If comparingMidspan = comparedMidspan Then
                                                                found = True
                                                                Exit For
                                                            End If
                                                        Next L
                                                        If Not found Then issues = issues & "• Couldn't find matching midspan (" & comparingMidspan & ") towards pole " & otherSheetName & " on sheet " & otherSheetName & vbLf
                                                    End If
                                                Next k
                                                
                        
                                                
                                                otherAngle = Trim(Replace(otherSheet.Range("TOPOLE" & j), "-", ""))
                                                If InStr(otherAngle, "(") > 0 And InStr(otherAngle, ")") > 0 Then otherAngle = Mid(otherAngle, InStr(otherAngle, "(") + 1, InStr(otherAngle, ")") - (InStr(otherAngle, "(") + 1))
                                                
                                                If Abs(CInt(angle) - CInt(otherAngle)) <> 180 Then issues = issues & "• Angle towards pole " & otherSheetName & " doesn't match angle on sheet " & otherSheetName & vbLf
                                            End If
                                        End If
                                    End If
                                Next otherName
                            Next j
                            
                            If Trim(Replace(pds.Range("TOPOLE" & i).offset(1, 0), "-", "")) <> "" Then
                                If InStr(otherSheet.Range("PROPOSEDHEIGHT"), "OL") = 0 Then
                                    If otherSheet.Range("CMRF1") <> "" And otherSheet.Range("PROPOSEDHEIGHT") <> "" And InStr(otherSheet.Range("CMRF1"), otherSheet.Range("PROPOSEDHEIGHT")) = 0 And InStr(otherSheet.Range("CMRF1"), "APPLY") = 0 Then paapms = True
                                Else
                                    For j = 0 To 100
                                        If otherSheet.Range("CMOWNER").offset(j, 0).Interior.color <> 16312794 Then Exit For
                                        If UCase(otherSheet.Range("CMOWNER").offset(j, 0)) = UCase(pds.Range("APPLICANT")) Then
                                            If InStr(otherSheet.Range("CMOWNER").offset(j, 0), "DG") = 0 Then
                                                If otherSheet.Range("CMRF1") <> "" And InStr(otherSheet.Range("CMRF1"), otherSheet.Range("CMHEIGHT").offset(j, 0)) = 0 And InStr(otherSheet.Range("CMRF1"), "APPLY") = 0 Then
                                                    paapms = True
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    Next j
                                End If
                            End If
                        Else
                            If Trim(Replace(pds.Range("TOPOLE" & i).offset(1, 0), "-", "")) <> "" Then
                                If otherSheetName = "N" Then
                                    connectingWarning = True
                                ElseIf otherSheetName = "E" Then connectingWarning = True
                                ElseIf otherSheetName = "S" Then connectingWarning = True
                                ElseIf otherSheetName = "W" Then connectingWarning = True
                                ElseIf otherSheetName = "NE" Then connectingWarning = True
                                ElseIf otherSheetName = "SE" Then connectingWarning = True
                                ElseIf otherSheetName = "SW" Then connectingWarning = True
                                ElseIf otherSheetName = "NW" Then connectingWarning = True
                                End If
                            End If
                        End If
                    End If
                End If
                Exit For
            End If
        Next name
    Next i
    
    If commaMidspan Then issues = issues & "• There shouldn't be a comma in the proposed midspans section" & vbLf
    If guessMidspan Then issues = issues & "• There shouldn't be a (GUESS) comment in the proposed midspans section, make sure it's in the right column" & vbLf
    If missingInchOrFoot Then issues = issues & "• Missing foot or inch symbol in proposed midspan section" & vbLf
    If connectingWarning Then warnings = warnings & "Warning: proposed midspan going towards pole not on job, check for overlapping jobs and make sure there's a comment in the notes section" & vbLf
    If namidSpan Then issues = issues & "• N/A midspan or tension in applicant section, remove them"
    
    If paapms And isEmpty(pds.Range("SUMSHEET8")) And isEmpty(pds.Range("SUMSHEET9")) Then
        If isEmpty(pds.Range("SUMSHEET10")) Then issues = issues & "• PAAPMS should be checked off if any application information changed" & vbLf
    Else
        If Not isEmpty(pds.Range("SUMSHEET10")) Then issues = issues & "• PAAPMS shouldn't be checked off if none of the application information changed or when the pole is denied or foreign" & vbLf
    End If
    
    ' Streetlight checks
    If InStr(pds.Range("SEP").offset(1, 0), "(Assuming Not Bonded)") > 0 Then issues = issues & "• Remove the (Assuming Not Bonded) or the entire violation if bonded" & vbLf
    
    ' Misc checks
    If InStr(pds.Range("ALTONE").text, "OUTAGE") > 0 Then Call checkEmptyCell(pds, issues, "OUTAGE", "Outage")
    If checkTreeNote(pds.Range("ALTONE")) And (isEmpty(pds.Range("TREE")) And (pds.Range("TREE2") = "NO" Or pds.Range("TREE2") = "N/A")) Then Call checkEmptyCell(pds, issues, "TREE", "Tree Trimming")
    If Not isEmpty(pds.Range("OUTAGE")) And InStr(pds.Range("ALTONE"), "OUTAGE") = 0 Then issues = issues & "• No outage note in alt 1" & vbLf
    If (Not isEmpty(pds.Range("TREE")) Or pds.Range("TREE2") = "YES") And checkTreeNote(pds.Range("ALTONE").text) = 0 Then issues = issues & "• No tree trimming note detected in alt 1" & vbLf
    If isEmpty(pds.Range("SEP").offset(-5, 0)) And Not isEmpty(pds.Range("ALTTHREE")) Then issues = issues & "• Front page CE Make Ready cell empty" & vbLf
    If isEmpty(pds.Range("OTHERPOLE")) And pds.Range("CEID") = "FOREIGN" Then issues = issues & "• Other pole owner should be checked off at the top if foreign" & vbLf
    If Not isEmpty(pds.Range("CEPOLE")) And Not isEmpty(pds.Range("OTHERPOLE")) Then issues = issues & "• Can't have both CE and Other pole owner checked off" & vbLf
    If isEmpty(pds.Range("DL")) And Not isEmpty(pds.Range("ALTONE")) Then issues = issues & "• Location number should be filled in if there's alt 1 work" & vbLf
    If Not isEmpty(ThisWorkbook.sheets("Control").Range("QATTC")) Then
        If Not isEmpty(pds.Range("DL")) And isEmpty(pds.Range("TTC")) Then issues = issues & "• There should be a TTC number if it's a location" & vbLf
    End If
    If InStr(pds.Range("ALTONE").text, "DETERIOR") > 0 And isEmpty(pds.Range("DET").offset(0, 1)) Then issues = issues & "• CE - Work violations Deterioration should be filled out if the pole is deteriorated" & vbLf
    If Not isEmpty(ThisWorkbook.sheets("Control").Range("QACO")) Then
        If InStr(pds.Range("INVENTORY").text, "CO") > 0 And (InStr(pds.Range("INVENTORY").text, "LCOM") = 0 And InStr(pds.Range("INVENTORY").text, "SA") = 0) Then issues = issues & "• If there's a CO, it should specify CO ON LCOM or CO ON SA" & vbLf
    End If
    If InStr(pds.Range("NJUNS"), "[REASON FOR WORK]") > 0 Then issues = issues & "• Remove the [REASON FOR WORK] from the njuns " & vbLf
    If InStr(pds.Range("ALTONE"), "[") > 0 Or InStr(pds.Range("ALTONE"), "]") > 0 Then issues = issues & "• Remove any [] brackets from Alt 1 work and anything between those brackets. These are used for place holder information in scripts" & vbLf
    If InStr(pds.Range("ALTTWO"), "[") > 0 Or InStr(pds.Range("ALTTWO"), "]") > 0 Then issues = issues & "• Remove any [] brackets from Alt 2 work and anything between those brackets. These are used for place holder information in scripts" & vbLf
    If InStr(pds.Range("ALTTHREE"), "[") > 0 Or InStr(pds.Range("ALTTHREE"), "]") > 0 Then issues = issues & "• Remove any [] brackets from Alt 3 work and anything between those brackets. These are used for place holder information in scripts" & vbLf
    
    If pole.replacePole And Not isEmpty(pds.Range("ALTONE")) Then Call checkEmptyCell(pds, issues, "REPLACEPOLE", "Replace pole")
    
    If isEmpty(pds.Range("PPRR")) And (Not isEmpty(pds.Range("TREE")) Or Not isEmpty(pds.Range("TREE2")) Or Not isEmpty(pds.Range("SUMSHEET13"))) Then warnings = warnings & "Warning: if there is tree work or pole staking then make sure there's no PPRR. Ignore this warning if it's all in ROW" & vbLf
    If Not isEmpty(pds.Range("PPRR")) And (isEmpty(pds.Range("TREE")) And isEmpty(pds.Range("TREE2")) And isEmpty(pds.Range("SUMSHEET13"))) Then issues = issues & "• PPRR shouldn't be checked off if there's no tree work or staking, uncheck PPRR or check off the reason for PPRR" & vbLf
    If pds.Range("SUMSHEET12") <> "-" Then
        If pds.Range("NEWAPPPF") <> pds.Range("SUMSHEET12") And Not isEmpty(pds.Range("NEWAPPPF")) Then warnings = warnings & "Warning: Pole Load Done does not match new App Pole Loading Done" & vbLf
        If pds.Range("ASISPF") <> pds.Range("SUMSHEET12") And isEmpty(pds.Range("NEWAPPPF")) And Not isEmpty(pds.Range("ASISPF")) Then warning = warning & "Warning: Pole Load Done does not match as-is Pole Loading Done" & vbLf
    End If
    If Not isEmpty(ThisWorkbook.sheets("Control").Range("QAFCS")) Then
        If (Not isEmpty(pds.Range("REPLACEANCHOR")) Or Not isEmpty(pds.Range("INSTALLANCHOR")) Or Not isEmpty(pds.Range("REPLACEPOLE")) Or Not isEmpty(pds.Range("INSTALLPOLE")) Or Not isEmpty(pds.Range("REPLACERISER")) Or Not isEmpty(pds.Range("REMOVEANCHOR")) Or Not isEmpty(pds.Range("REMOVEPOLE"))) Then
            nameExists = pds.Evaluate("ISREF(" & "FCS" & ")")
            If nameExists Then
                If Trim(pds.Range("FCS").offset(0, 1).Value) = "" Then
                    issues = issues & "• If there's work that requires a Missdig ticket, then there should be a first cross street filled in" & vbLf
                End If
            End If
        End If
    End If
    If ThisWorkbook.RemoveParentheses(pds.name) <> pds.Range("POLENUM") Then issues = issues & "• Pole number doesn't match sheet name, this can cause a lot of scripts to bug out, fix whichever one is wrong." & vbLf
    If Trim(pds.Range("CLASS")) = "7" And isEmpty(pds.Range("SUMSHEET9")) Then warnings = warnings & "Warning: Class 7 poles can't be modeled in pole foreman, make sure to account for this." & vbLf
    If UCase(Trim(pds.Range("NEWAPP"))) = "DENIED" And pds.Range("NEWAPPPF").Value <> "N/A" Then issues = issues & "• New APP Pole Loading should be NA if percentage is Denied." & vbLf
    
    
    ' midspan check
    Dim utilityMidspanIssue As Boolean: utilityMidspanIssue = False
    Dim comMidspanIssue As Boolean: comMidspanIssue = False
    Dim utilityBoundryReached As Boolean: utilityBoundryReached = False
    Dim comBoundryReached As Boolean: comBoundryReached = False
    For i = 0 To 100
        If pds.Range("UTMIDSPAN1").offset(i, 0).Interior.color <> 16312794 Then utilityBoundryReached = True
        If pds.Range("CMMIDSPAN1").offset(i + 1, 0).Interior.color <> 16312794 Then comBoundryReached = True
        For j = 1 To 12
            For Each name In pds.names
                If name.name = "'" & pds.name & "'" & "!UTMIDSPAN" & j Then
                    If Not utilityBoundryReached And Trim(pds.Range("UTMIDSPAN" & j).offset(i, 0)) = "0'0""" Then utilityMidspanIssue = True
                    If Not comBoundryReached And Trim(pds.Range("CMMIDSPAN" & j).offset(i, 0)) = "0'0""" Then comMidspanIssue = True
                    Exit For
                End If
            Next name
            If utilityMidspanIssue And comMidspanIssue Then Exit For
        Next j
        If utilityBoundryReached And comBoundryReached Then Exit For
    Next i
    If utilityMidspanIssue Then issues = issues & "• 0'0"" midspan detected in utility midspans" & vbLf
    If comMidspanIssue Then issues = issues & "• 0'0"" midspan detected in comm midspans" & vbLf
    
    ' Safe attach section
    If pole.location = "" Then
        If (pole.lowestPower > 0 And pole.applicant.modification > pole.lowestPower - 40) Or (pole.streeLightBB > 0 And pole.Bonded <> "YES" And Abs(pole.applicant.modification - pole.streeLightBB) < 40) Then
            issues = issues & "• Attaching in 40"" safety zone violation or within 40"" of unbonded streetlight with no CE work." & vbLf
        End If
    End If
    
    
    ' Troll section
    Randomize
    UserName = LCase(Environ("USERNAME"))
    If UserName = "hnguyen1" Or UserName = "mwatts5" Or UserName = "aabraham" Then
        Const DesiredPercentage As Double = 0.05
        Dim randomNumber As Double: randomNumber = Rnd
        
        If randomNumber <= DesiredPercentage And issues <> "" Then
            randomNumber = Int((4 * Rnd) + 1)
                Select Case randomNumber
                    Case 1
                        warnings = warnings & vbLf & "Wrong Highlight Color" & vbLf
                    Case 2
                        If UserName = "hnguyen1" Then
                            warnings = warnings & vbLf & "I believe in you <3" & vbLf
                        ElseIf UserName = "mwatts5" Then
                            warnings = warnings & vbLf & "I don't believe in you :/" & vbLf
                        End If
                    Case 3
                        warnings = warnings & vbLf & "Did you even look at the pictures!?!?" & vbLf
                    Case 4
                        warnings = warnings & vbLf & "Probably should have someone else design this one tbh" & vbLf
                    Case 5
                        warnings = warnings & vbLf & "Spelling errors, please make sure to doble check your work" & vbLf
                End Select
        End If
    End If
    
    issues = issues & warnings
    If Len(issues) > 1 Then issues = Left(issues, Len(issues) - 1)
    
    If isEmpty(issues) Then issues = "No issues found" & vbLf
    header = "Pole: " & pds.Range("POLENUM")
    If Not isEmpty(pds.Range("CEID")) Then header = header & " CEID: " & pds.Range("CEID")
    If Not isEmpty(pds.Range("DL")) Then header = header & " LOC: " & pds.Range("DL")
    issues = header & vbLf & issues

    QACheck = issues
End Function

Private Function checkTreeNote(str As String) As Boolean
    If InStr(UCase(Replace(str, " ", "")), "TREEWORK") > 0 Then checkTreeNote = True: Exit Function
    If InStr(UCase(Replace(str, " ", "")), "BUSHWORK") > 0 Then checkTreeNote = True: Exit Function
    If InStr(UCase(Replace(str, " ", "")), "BRUSHWORK") > 0 Then checkTreeNote = True: Exit Function
    If InStr(UCase(Replace(str, " ", "")), "TREETRIM") > 0 Then checkTreeNote = True: Exit Function
    If InStr(UCase(Replace(str, " ", "")), "BUSHTRIM") > 0 Then checkTreeNote = True: Exit Function
    If InStr(UCase(Replace(str, " ", "")), "BRUSHTRIM") > 0 Then checkTreeNote = True: Exit Function
    checkTreeNote = False
End Function

Private Function isEmpty(str As String) As Boolean
    If Trim(Replace(str, "-", "")) = "" Or UCase(Trim(Replace(str, "/", ""))) = "NA" Or UCase(Trim(str)) = "FALSE" Then
        isEmpty = True
    Else
        isEmpty = False
    End If
End Function

Private Sub checkEmptyCell(pds As Worksheet, ByRef issues As String, rangeName As String, cellName As String)
    If UCase(pds.Range(rangeName)) = "FALSE" Then
        issues = issues & "• " & cellName & " checkbox unchecked" & vbLf
    ElseIf isEmpty(pds.Range(rangeName)) Then
        issues = issues & "• " & cellName & " cell empty" & vbLf
    End If
End Sub



