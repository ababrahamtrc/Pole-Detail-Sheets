Attribute VB_Name = "ExportMKR"
Sub exportAllMKR()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    Call exportCMRF(True)
    Call exportSS(True)
    Call exportCS(True)
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    
    MsgBox "Finished Exporting"
End Sub

Sub exportCMRF(Optional ignoreMessage As Boolean)
    Dim project As project: Set project = New project
    Call project.extractFromSheets
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    On Error Resume Next
    
    Dim blankUsed As String: blankUsed = "N/A"
    
    jobFolderPath = Utilities.GetJobDirectory(project.notification)
    If jobFolderPath <> "" And Dir(jobFolderPath & "\Design Related Materials\") <> "" Then
        fileSaveName = jobFolderPath & "\Design Related Materials\" & project.notification & " - Comm Make Ready Form.xlsx"
    Else
        fileSaveName = GetLocalPath(ThisWorkbook.Path) & "\" & project.notification & " - Comm Make Ready Form.xlsx"
        If InStr(fileSaveName, "sharepoint") > 0 Or fileSaveName = "" Then fileSaveName = Environ("USERPROFILE") & "\Downloads\" & project.notification & " - Comm Make Ready Form.xlsx"
    End If
    
    Dim cmrfWs As Worksheet: Set cmrfWs = ThisWorkbook.sheets("CMRF")
    Call cmrfWs.Unprotect
    Dim newWorkBook As Workbook: Set newWorkBook = Workbooks.Add
    cmrfWs.visible = xlSheetVisible
    cmrfWs.Copy After:=newWorkBook.sheets(1)
    cmrfWs.visible = xlSheetHidden
    newWorkBook.sheets(1).Delete
    newWorkBook.sheets(1).name = "Sheet 1"
    
    Dim cmrfOwners As Scripting.Dictionary: Set cmrfOwners = New Scripting.Dictionary
    Dim pole As pole
    Dim ROW As Integer: ROW = 4
    For Each pole In project.poles
        
        For col = 1 To 22
            newWorkBook.sheets(1).cells(ROW, col).value = blankUsed
        Next col

        If pole.njunsTicket <> "" Then newWorkBook.sheets(1).cells(ROW, 1).value = pole.njunsTicket
        If pole.address <> "" Then newWorkBook.sheets(1).cells(ROW, 2).value = pole.address
        If pole.existingCEID <> "" Then newWorkBook.sheets(1).cells(ROW, 3).value = pole.existingCEID
        If pole.location <> "" Then newWorkBook.sheets(1).cells(ROW, 4).value = pole.location
        If pole.poleNumber <> "" Then newWorkBook.sheets(1).cells(ROW, 5).value = "M1P" & pole.poleNumber
        If pole.sheet.Range("CMRF1") <> "" Then newWorkBook.sheets(1).cells(ROW, 6).value = pole.sheet.Range("CMRF1")
        If pole.sheet.Range("CMRF2") <> "" Then newWorkBook.sheets(1).cells(ROW, 7).value = pole.sheet.Range("CMRF2")
        If pole.sheet.Range("CMRF3") <> "" Then newWorkBook.sheets(1).cells(ROW, 8).value = pole.sheet.Range("CMRF3")
                
        Dim cell As Range
        For i = 1 To 8
            Set cell = pole.sheet.Range("COMM" & i)
            If Not InStr(cell, "COMM #") > 0 Then
                Dim heights As Scripting.Dictionary: Set heights = New Scripting.Dictionary
                Dim modifications As Scripting.Dictionary: Set modifications = New Scripting.Dictionary
                
                If Not cmrfOwners.exists(cell.value) Then
                    If cmrfOwners.count < 6 Then newWorkBook.sheets(1).cells(2, 9 + (cmrfOwners.count * 2)) = cell.value
                    cmrfOwners(cell.value) = cmrfOwners.count
                End If
                cmrfSlot = cmrfOwners(cell.value)
                
                If cmrfSlot < 6 Then
                    For j = 1 To 100 Step 2
                        If cell.offSet(1 + j, 0).value = "" Then Exit For
                        If Not heights.exists(cell.offSet(1 + j, 0).value) Then heights.Add cell.offSet(1 + j, 0).value, Nothing
                    Next j
                    For j = 1 To 100 Step 2
                        If cell.offSet(1 + j, 0).offSet(0, 1).value = "" Then Exit For
                        If Not modifications.exists(cell.offSet(1 + j, 0).offSet(0, 1).value) Then
                            If ThisWorkbook.RemoveParentheses(cell.offSet(1 + j, 0).value) <> ThisWorkbook.RemoveParentheses(cell.offSet(1 + j, 0).offSet(0, 1).value) Then
                                modifications.Add cell.offSet(1 + j, 0).offSet(0, 1).value, Nothing
                            End If
                        End If
                    Next j
                    
                    If heights.count > 0 Then newWorkBook.sheets(1).cells(ROW, 9 + (cmrfSlot * 2)) = Utilities.JoinCollection(heights, vbLf)
                    If modifications.count > 0 And Utilities.JoinCollection(heights, vbLf) <> Utilities.JoinCollection(modifications, vbLf) Then newWorkBook.sheets(1).cells(ROW, 10 + (cmrfSlot * 2)) = Utilities.JoinCollection(modifications, vbLf)
                End If
            End If
        Next i
        
        ROW = ROW + 1
    Next pole
    
    With newWorkBook.sheets(1).UsedRange
        .NumberFormat = "@"
        .value = .value
    End With
    
    newWorkBook.sheets(1).PageSetup.LeftHeader = "Notification: " & project.notification & vbLf & "Permit: " & project.Permit & vbLf & "Date: " & Date & vbLf & "New Attacher: " & project.applicant
    
    lastRow = newWorkBook.sheets(1).cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).ROW
    lastCol = newWorkBook.sheets(1).cells.Find("*", searchorder:=xlByColumns, searchdirection:=xlPrevious).column
    
    For Each cell In newWorkBook.sheets(1).Rows(2).cells
        If InStr(cell.value, "COMM #") > 0 Then
            lastCol = cell.column - 1
            Exit For
        End If
    Next cell
    
    newWorkBook.sheets(1).PageSetup.PrintArea = newWorkBook.sheets(1).Range(newWorkBook.sheets(1).cells(1, 1), newWorkBook.sheets(1).cells(lastRow, lastCol)).address
    
    For i = newWorkBook.sheets(1).UsedRange.Rows.count To 1 Step -1
        If Intersect(newWorkBook.sheets(1).Rows(i), Range(newWorkBook.sheets(1).PageSetup.PrintArea)) Is Nothing Then
             newWorkBook.sheets(1).Rows(i).Delete
        End If
    Next i
    
    For Each col In newWorkBook.sheets(1).UsedRange.Columns
        If Intersect(col, Range(newWorkBook.sheets(1).PageSetup.PrintArea)) Is Nothing Then
            col.Delete shift:=xlUp
        End If
    Next col
    
    newWorkBook.sheets(1).Range(newWorkBook.sheets(1).cells(1, 9), newWorkBook.sheets(1).cells(1, lastCol)).Merge
    
    Dim dataRange As Range
    Set dataRange = newWorkBook.sheets(1).Range("A4:A" & newWorkBook.sheets(1).cells(newWorkBook.sheets(1).Rows.count, "A").End(xlUp).ROW)
    dataRange.Columns.ColumnWidth = 200
    dataRange.Columns.AutoFit
    dataRange.Columns.AutoFit
    dataRange.Columns.AutoFit
    
    dataRange.Rows.RowHeight = 200
    dataRange.Rows.AutoFit
    dataRange.Rows.AutoFit
    dataRange.Rows.AutoFit
    
    If dataRange.Columns(1).ColumnWidth < 6 Then dataRange.Columns(1).ColumnWidth = 6
    For i = 6 To 22
        newWorkBook.sheets(1).Columns(i).ColumnWidth = 200
        newWorkBook.sheets(1).Columns(i).AutoFit
        newWorkBook.sheets(1).Columns(i).AutoFit
        newWorkBook.sheets(1).Columns(i).AutoFit
        If newWorkBook.sheets(1).Columns(i).ColumnWidth < 6 Then newWorkBook.sheets(1).Columns(i).ColumnWidth = 6
        If newWorkBook.sheets(1).Columns(i).ColumnWidth > 10 Then
            newWorkBook.sheets(1).Columns(i).ColumnWidth = 10
            newWorkBook.sheets(1).Columns(i).WrapText = True
            With newWorkBook.sheets(1)
                If i > 8 Then
                    .cells(3, i).WrapText = False
                Else
                    .Range(.cells(2, i), .cells(3, i)).WrapText = False
                End If
            End With
        End If
    Next i
    
    totalWidth = 0
    For Each col In Range(newWorkBook.sheets(1).PageSetup.PrintArea).Columns
        totalWidth = totalWidth + col.width
    Next col
    
    If totalWidth >= 1227 Then
        If newWorkBook.sheets(1).Columns(2).ColumnWidth - (totalWidth - 1227) > 12 Then
            newWorkBook.sheets(1).Columns(2).ColumnWidth = newWorkBook.sheets(1).Columns(2).ColumnWidth - (totalWidth - 1227)
        Else
            newWorkBook.sheets(1).Columns(2).ColumnWidth = 12
        End If
    End If
    
    newWorkBook.sheets(1).Rows.AutoFit
    newWorkBook.sheets(1).Activate
    ActiveWindow.View = xlPageLayoutView
    
    newWorkBook.SaveAs fileSaveName
    Call ExportPDF(newWorkBook.sheets(1), Replace(fileSaveName, ".xlsx", ".pdf"))
    newWorkBook.Close savechanges:=True
    
    ThisWorkbook.sheets("Control").Activate
    
    If Not ignoreMessage Then
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        
        MsgBox "Finished Exporting Comm Make Ready Form"
    End If
End Sub

Sub exportSS(Optional ignoreMessage As Boolean)
    Dim project As project: Set project = New project
    Call project.extractFromSheets
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    jobFolderPath = Utilities.GetJobDirectory(project.notification)
    If jobFolderPath <> "" And Dir(jobFolderPath & "\Pole Detail Sheets\") <> "" Then
        fileSaveName = jobFolderPath & "\Pole Detail Sheets\" & project.notification & " - Summary Sheet.xlsx"
    Else
        fileSaveName = GetLocalPath(ThisWorkbook.Path) & "\" & project.notification & " - Summary Sheet.xlsx"
        If InStr(fileSaveName, "sharepoint") > 0 Or fileSaveName = "" Then fileSaveName = Environ("USERPROFILE") & "\Downloads\" & project.notification & " - Summary Sheet.xlsx"
    End If
    
    Dim ssWs As Worksheet: Set ssWs = ThisWorkbook.sheets("SS")
    Call ssWs.Unprotect
    Dim newWorkBook As Workbook: Set newWorkBook = Workbooks.Add
    ssWs.visible = xlSheetVisible
    ssWs.Copy After:=newWorkBook.sheets(1)
    ssWs.visible = xlSheetHidden
    newWorkBook.sheets(1).Delete
    newWorkBook.sheets(1).name = "Sheet 1"
    
    newWorkBook.sheets(1).cells(4, 11) = project.Permit
    newWorkBook.sheets(1).cells(5, 11) = project.notification
    newWorkBook.sheets(1).cells(4, 30) = project.fielder
    newWorkBook.sheets(1).cells(5, 30) = Date
    
    Dim pole As pole
    Dim ROW As Integer: ROW = 11
    For Each pole In project.poles
        newWorkBook.sheets(1).cells(ROW, 3).value = pole.poleNumber
        newWorkBook.sheets(1).cells(ROW, 5).value = pole.existingCEID
        
        newWorkBook.sheets(1).cells(ROW, 9).value = pole.sheet.Range("SUMSHEET1")
        newWorkBook.sheets(1).cells(ROW, 10).value = pole.sheet.Range("SUMSHEET2")
        newWorkBook.sheets(1).cells(ROW, 11).value = pole.sheet.Range("SUMSHEET3")
        newWorkBook.sheets(1).cells(ROW, 12).value = pole.sheet.Range("SUMSHEET4")
        newWorkBook.sheets(1).cells(ROW, 13).value = pole.sheet.Range("SUMSHEET5")
        newWorkBook.sheets(1).cells(ROW, 14).value = pole.sheet.Range("SUMSHEET6")
        newWorkBook.sheets(1).cells(ROW, 15).value = pole.sheet.Range("SUMSHEET7")
        newWorkBook.sheets(1).cells(ROW, 16).value = pole.sheet.Range("SUMSHEET8")
        newWorkBook.sheets(1).cells(ROW, 17).value = pole.sheet.Range("SUMSHEET9")
        newWorkBook.sheets(1).cells(ROW, 18).value = pole.sheet.Range("SUMSHEET10")
        newWorkBook.sheets(1).cells(ROW, 19).value = pole.sheet.Range("SUMSHEET14")
        
        newWorkBook.sheets(1).cells(ROW, 38).value = IIf(pole.sheet.Range("SUMSHEET11") <> "-", pole.sheet.Range("SUMSHEET11"), "N/A")
        newWorkBook.sheets(1).cells(ROW, 39).value = IIf(pole.sheet.Range("SUMSHEET12") <> "-", pole.sheet.Range("SUMSHEET12"), "N/A")
        newWorkBook.sheets(1).cells(ROW, 40).value = IIf(pole.location <> "", pole.location, "N/A")
        newWorkBook.sheets(1).cells(ROW, 41).value = pole.sheet.Range("SUMSHEET13")
        
        ROW = ROW + 1
    Next pole
    
    newWorkBook.SaveAs fileSaveName
    newWorkBook.Close savechanges:=True
    
    ThisWorkbook.sheets("Control").Activate
    
    If Not ignoreMessage Then
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        
        MsgBox "Finished Exporting Summary Sheet"
    End If
End Sub

Sub exportCS(Optional ignoreMessage As Boolean)
    Dim project As project: Set project = New project
    Call project.extractFromSheets
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    jobFolderPath = Utilities.GetJobDirectory(project.notification)
    If jobFolderPath <> "" And Dir(jobFolderPath & "\Design Related Materials\") <> "" Then
        fileSaveName = jobFolderPath & "\Design Related Materials\" & project.notification & " - Make Ready Engineering Cost.xlsx"
    Else
        fileSaveName = GetLocalPath(ThisWorkbook.Path) & "\" & project.notification & " - Make Ready Engineering Cost.xlsx"
        If InStr(fileSaveName, "sharepoint") > 0 Or fileSaveName = "" Then fileSaveName = Environ("USERPROFILE") & "\Downloads\" & project.notification & " - Make Ready Engineering Cost.xlsx"
    End If
    
    Dim csWs As Worksheet: Set csWs = ThisWorkbook.sheets("CS")
    Call csWs.Unprotect
    Dim newWorkBook As Workbook: Set newWorkBook = Workbooks.Add
    csWs.visible = xlSheetVisible
    csWs.Copy After:=newWorkBook.sheets(1)
    csWs.visible = xlSheetHidden
    newWorkBook.sheets(1).Delete
    newWorkBook.sheets(1).name = "Sheet 1"
    
    newWorkBook.sheets(1).cells(2, 1).value = newWorkBook.sheets(1).cells(2, 1).value & project.notification
    newWorkBook.sheets(1).cells(3, 1).value = newWorkBook.sheets(1).cells(3, 1).value & project.Permit
    newWorkBook.sheets(1).cells(4, 1).value = newWorkBook.sheets(1).cells(4, 1).value & project.applicant
    
    Const hourlyCost = 68.63
    
    Dim mkrCosts As Scripting.Dictionary: Set mkrCosts = New Scripting.Dictionary
    mkrCosts("Basic Make Ready Work Order") = Round(2.46 * hourlyCost, 2)
    mkrCosts("Secondary Pole") = Round(2.25 * hourlyCost, 2)
    mkrCosts("Single Phase Pole") = Round(3.45 * hourlyCost, 2)
    mkrCosts("2 Phase Pole") = Round(3.12 * hourlyCost, 2)
    mkrCosts("3 Phase Pole") = Round(3.6 * hourlyCost, 2)
    mkrCosts("Primary Riser") = Round(3.5 * hourlyCost, 2)
    mkrCosts("Aerial Spacer Cable") = Round(3 * hourlyCost, 2)
    mkrCosts("Tree Wire") = Round(3 * hourlyCost, 2)
    mkrCosts("New/Midspan Pole") = Round(3 * hourlyCost, 2)
    
    Dim pole As pole
    Dim ROW As Integer: ROW = 6
    For Each pole In project.poles
        If pole.location <> "" Then
            newWorkBook.sheets(1).cells(ROW, 1).value = pole.location
            Dim alt2 As Boolean: alt2 = False
            Dim alt3 As Boolean: alt3 = False
            
            alt2 = pole.sheet.Range("SUMSHEET3") Or (InStr(pole.sheet.Range("SUMSHEET11"), "2") > 0)
            alt3 = pole.sheet.Range("SUMSHEET6") Or (InStr(pole.sheet.Range("SUMSHEET11"), "3") > 0)
            newWorkBook.sheets(1).cells(ROW, 2).value = IIf(alt2 And alt3, "ALT 2 AND 3", IIf(alt2, "ALT 2", IIf(alt3, "ALT 3", "")))
            
            Set results = determineEngineeringCost(pole)
            newWorkBook.sheets(1).cells(ROW, 3).value = results(1)
            newWorkBook.sheets(1).cells(ROW, 4).value = "$" & results(2)
            
            Set cell = newWorkBook.sheets(1).cells(ROW, 1)
            Set targetRange = cell.Resize(1, 4)
            With targetRange.Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            
            ROW = ROW + 1
        End If
    Next pole
    
    newWorkBook.SaveAs fileSaveName
    Call ExportPDF(newWorkBook.sheets(1), Replace(fileSaveName, ".xlsx", ".pdf"))
    newWorkBook.Close savechanges:=True
    
    ThisWorkbook.sheets("Control").Activate
    
    If Not ignoreMessage Then
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.DisplayAlerts = True
    
        MsgBox "Finished Exporting Make Ready Engineering Cost Sheet"
    End If
End Sub

Public Function determineEngineeringCost(pole As pole) As Collection
    Const hourlyCost = 68.63
    
    Dim mkrCosts As Scripting.Dictionary: Set mkrCosts = New Scripting.Dictionary
    mkrCosts("Basic Make Ready Work Order") = Round(2.46 * hourlyCost, 2)
    mkrCosts("Secondary Pole") = Round(2.25 * hourlyCost, 2)
    mkrCosts("Single Phase Pole") = Round(3.45 * hourlyCost, 2)
    mkrCosts("2 Phase Pole") = Round(3.12 * hourlyCost, 2)
    mkrCosts("3 Phase Pole") = Round(3.6 * hourlyCost, 2)
    mkrCosts("Primary Riser") = Round(3.5 * hourlyCost, 2)
    mkrCosts("Aerial Spacer Cable") = Round(3 * hourlyCost, 2)
    mkrCosts("Tree Wire") = Round(3 * hourlyCost, 2)
    mkrCosts("New/Midspan Pole") = Round(3 * hourlyCost, 2)
    
    Dim results As Collection: Set results = New Collection
    
    Dim wire As wire
    If pole.ReplacePole Then
        If pole.primaries.count > 0 Then
            Dim largestPhase As Integer: largestPhase = 1
            For Each wire In pole.primaries
                If wire.phase > largestPhase Then largestPhase = wire.phase
            Next wire
            
            If largestPhase = 1 Then
                results.Add "Single Phase Pole"
            ElseIf largestPhase = 2 Then
                results.Add "2 Phase Pole"
            ElseIf largestPhase >= 3 Then
                results.Add "3 Phase Pole"
            End If
        Else
            results.Add "Secondary Pole"
        End If
    Else
        results.Add "Basic Make Ready Work Order"
    End If
    
    results.Add mkrCosts(results(1))
    Set determineEngineeringCost = results
End Function

Sub ExportPDF(ws As Worksheet, pdfPath As String)
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
End Sub
