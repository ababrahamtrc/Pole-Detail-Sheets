Attribute VB_Name = "ChangeColors"
Sub changeHighlightColor()
    Call LogMessage.SendLogMessage("changeHighlightColor")
    
    Dim controlWs As Worksheet: Set controlWs = ThisWorkbook.sheets("Control")
    Dim fc As FormatCondition
    Dim sheet As Worksheet
    Dim sheets As Collection: Set sheets = New Collection
    
    On Error Resume Next
    
    For Each sheet In ThisWorkbook.sheets
        If sheet.cells(2, 2).value = "Notification:" Then
            Call ThisWorkbook.decideTabColor(sheet)
            For i = 1 To sheet.cells.FormatConditions.count
                Set fc = sheet.cells.FormatConditions(i)
                If Application.Intersect(fc.AppliesTo.cells, sheet.Range("CLASSESTIMATE").MergeArea) Is Nothing _
                    And Application.Intersect(fc.AppliesTo.cells, sheet.Range("ASISPF").MergeArea) Is Nothing _
                    And Application.Intersect(fc.AppliesTo.cells, sheet.Range("NEWAPP").MergeArea) Is Nothing _
                    And Application.Intersect(fc.AppliesTo.cells, sheet.Range("PGUY").MergeArea) Is Nothing _
                    And Application.Intersect(fc.AppliesTo.cells, sheet.Range("PGUY2").MergeArea) Is Nothing _
                    And Application.Intersect(fc.AppliesTo.cells, sheet.Range("ROOMTOGUY").MergeArea) Is Nothing _
                    And Application.Intersect(fc.AppliesTo.cells, sheet.Range("BONDED").MergeArea) Is Nothing _
                    And Application.Intersect(fc.AppliesTo.cells, sheet.Range("BONDED").offSet(-1, 1).MergeArea) Is Nothing _
                    And Application.Intersect(fc.AppliesTo.cells, sheet.Range("TREE2").MergeArea) Is Nothing _
                    And Application.Intersect(fc.AppliesTo.cells, sheet.Range("TREE2").offSet(1, 0).MergeArea) Is Nothing _
                    And Application.Intersect(fc.AppliesTo.cells, sheet.Range("TREE2").offSet(2, -2).MergeArea) Is Nothing Then
                    fc.Interior.color = controlWs.Range("Default_Color").offSet(0, 1).Interior.color
                    fc.Font.color = controlWs.Range("Default_Color").offSet(0, 1).Font.color
                End If
            Next i
        End If
    Next sheet
    
    MsgBox "Highlight/Tab colors changed on all sheets!"
End Sub

Sub changeEmptyColor()
    Call LogMessage.SendLogMessage("changeEmptyColor")
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
    
    Dim fc As FormatCondition
    Dim sheet As Worksheet
    Dim sheets As Collection: Set sheets = New Collection
    
    On Error Resume Next
    sheets.Add ThisWorkbook.sheets("MKR SAP")
    sheets.Add ThisWorkbook.sheets("SI SAP")
    
    
    For Each sheet In sheets
        sheet.Range("EMPTYCOLOR").Interior.color = ws.Range("EMPTYCOLOR").Interior.color
        For i = 1 To sheet.cells.FormatConditions.count
            Set fc = sheet.cells.FormatConditions(i)
            fc.Interior.color = ws.Range("EMPTYCOLOR").Interior.color
        Next i
    Next sheet
End Sub

