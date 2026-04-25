Attribute VB_Name = "ExportMikeLExcel"
Dim colors As Collection:

Sub MikeLExcel()
    On Error Resume Next
    
    Call LogMessage.SendLogMessage("mikeLExcel")
    
    Dim MikeLExcel As Workbook: Set MikeLExcel = Workbooks.Add
    
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    MikeLExcel.sheets(1).Cells(1, 1).Value = "LOC #"
    MikeLExcel.sheets(1).Cells(1, 2).Value = "CREW NOTES"
    With MikeLExcel.sheets(1).Range("A1:B1")
        .Font.Bold = True
        .Font.name = "Aptos Narrow"
        .Font.size = 11
        .Interior.color = RGB(232, 232, 232)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.ColorIndex = 1
    End With
    
    Set colors = New Collection
    colors.Add RGB(218, 242, 208)
    colors.Add RGB(202, 237, 251)
    
    Dim Notification As String: Notification = ""
    Dim row As Integer: row = 2
    
    Dim locations As Collection: Set locations = New Collection
    Dim treeWorkLocations As Collection: Set treeWorkLocations = New Collection
    Dim topPoleWorkLocations As Collection: Set topPoleWorkLocations = New Collection
    Dim outageWorkLocations As Collection: Set outageWorkLocations = New Collection
    
    For Each sheet In ThisWorkbook.sheets
        If sheet.name <> "4 Spans" And sheet.name <> "8 Spans" And sheet.name <> "12 Spans" And sheet.Cells(2, 2).Value = "Notification:" Then
            If Notification = "" And sheet.Range("NOTIFICATION").Value <> "" Then Notification = sheet.Range("NOTIFICATION").Value
            If sheet.Range("DL").Value <> "" Then
                If locations.count = 0 Then
                    locations.Add sheet.Range("DL")
                Else
                    added = False
                    For i = 1 To locations.count
                        If sheet.Range("DL") < locations(i) Then
                            locations.Add item:=sheet.Range("DL"), Before:=i
                            added = True
                            Exit For
                        End If
                    Next i
                    If Not added Then locations.Add sheet.Range("DL")
                End If
                
            End If
        End If
    Next sheet
    
    For Each location In locations
        Set sheet = Nothing
        For Each ws In ThisWorkbook.sheets
            If ws.name <> "4 Spans" And ws.name <> "8 Spans" And ws.name <> "12 Spans" And ws.Cells(2, 2).Value = "Notification:" Then
                If location = ws.Range("DL") Then
                    Set sheet = ws
                    Exit For
                End If
            End If
        Next ws
    
        If InStr(sheet.Range("ALTONE"), "TREE WORK") > 0 Then
            treeWorkLocations.Add location
        ElseIf InStr(sheet.Range("ALTONE"), "TREE TRIM") > 0 Then
            treeWorkLocations.Add location
        ElseIf InStr(sheet.Range("ALTONE"), "BUSH WORK") > 0 Then
            treeWorkLocations.Add location
        ElseIf InStr(sheet.Range("ALTONE"), "BUSH TRIM") > 0 Then
            treeWorkLocations.Add location
        ElseIf InStr(sheet.Range("ALTONE"), "BRUSH WORK") > 0 Then
            treeWorkLocations.Add location
        ElseIf InStr(sheet.Range("ALTONE"), "BRUSH TRIM") > 0 Then
            treeWorkLocations.Add location
        End If
        
        If InStr(sheet.Range("ALTONE"), "TOP POLE") > 0 Then topPoleWorkLocations.Add location
        If InStr(sheet.Range("ALTONE"), "OUTAGE") > 0 Then outageWorkLocations.Add location
    
        If ws Is Nothing Then
            MsgBox "Error generating Mike L Excel. Double check location numbers or contact Alex."
            Exit Sub
        End If
    
        Call generateMikeLExcelRow(MikeLExcel, row, "P" & sheet.Range("POLENUM").Value & "-L" & sheet.Range("DL").Value, sheet.Range("ALTONE").Value)
        row = row + 1
    Next location
    
    Dim Project As Project: Set Project = New Project
    Call Project.extractFromSheets
    Dim pole As pole
    Dim poleReplacements As Collection: Set poleReplacements = New Collection
    For Each pole In Project.poles
        If pole.replacePole Then poleReplacements.Add pole.location
    Next pole
    
    Dim transferWorkLocations As Collection: Set transferWorkLocations = New Collection
    For Each pole In Project.poles
        If InStr(Replace(UCase(pole.Alt1), " ", ""), "TRANSFERAGREEMENT") > 0 Then transferWorkLocations.Add pole.location
    Next pole
    
    Call generateMikeLExcelRow(MikeLExcel, row, "TREE WORK LOCATIONS", combineList(treeWorkLocations))
    Call generateMikeLExcelRow(MikeLExcel, row + 1, "TOP POLE LOCATIONS", combineList(topPoleWorkLocations))
    Call generateMikeLExcelRow(MikeLExcel, row + 2, "CE COMM TRANSFER LOCATIONS", combineList(transferWorkLocations))
    Call generateMikeLExcelRow(MikeLExcel, row + 3, "POLE REPLACEMENT LOCATIONS", combineList(poleReplacements))
    Call generateMikeLExcelRow(MikeLExcel, row + 4, "OUTAGE LOCATIONS", combineList(outageWorkLocations))
    
    
    MikeLExcel.sheets(1).Columns(2).ColumnWidth = 200
    MikeLExcel.sheets(1).Cells.EntireColumn.AutoFit
    MikeLExcel.sheets(1).Cells.EntireRow.AutoFit
    
    filePath = ThisWorkbook.path & "\" & Notification & " - Mike L Excel.xlsx"
    If InStr(filePath, "sharepoint") > 0 Then filePath = Environ("USERPROFILE") & "\" & Notification & " - Mike L Excel.xlsx"
    
    MikeLExcel.SaveAs fileName:=filePath
    
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    
End Sub

Private Sub generateMikeLExcelRow(MikeLExcel As Workbook, row As Integer, value1 As String, value2 As String)
    With MikeLExcel.sheets(1).Cells(row, 1)
            .Value = value1
            .Font.name = "Calibri"
            .Font.size = 11
            .Interior.color = colors((row Mod 2) + 1)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.ColorIndex = 1
    End With
    With MikeLExcel.sheets(1).Cells(row, 2)
            .Value = value2
            .Font.name = "Arial"
            .Font.size = 9
            .VerticalAlignment = xlVAlignTop
            .HorizontalAlignment = xlHAlignLeft
            .Interior.color = colors((row Mod 2) + 1)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.ColorIndex = 1
    End With
End Sub

Private Function combineList(list As Collection) As String
    combinedList = ""
    For Each item In list
        combinedList = combinedList & item & ", "
    Next item
    If Len(combinedList) > 2 Then combineList = Left(combinedList, Len(combinedList) - 2)
End Function

