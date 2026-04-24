Attribute VB_Name = "ChangeColors"
Sub changeHighlightColor()
    Call LogMessage.SendLogMessage("changeHighlightColor")
    
    Dim controlWs As Worksheet: Set controlWs = ThisWorkbook.sheets("Control")
    Dim fc As FormatCondition
    Dim sheet As Worksheet
    Dim sheets As Collection: Set sheets = New Collection
    
    On Error Resume Next
    
    For Each sheet In ThisWorkbook.sheets
        If sheet.Cells(2, 2).Value = "Notification:" Then
            Call ThisWorkbook.decideTabColor(sheet)
            For i = 1 To sheet.Cells.FormatConditions.count
                Set fc = sheet.Cells.FormatConditions(i)
                If Application.Intersect(fc.AppliesTo.Cells, sheet.Range("CLASSESTIMATE").MergeArea) Is Nothing _
                    And Application.Intersect(fc.AppliesTo.Cells, sheet.Range("ASISPF").MergeArea) Is Nothing _
                    And Application.Intersect(fc.AppliesTo.Cells, sheet.Range("NEWAPP").MergeArea) Is Nothing _
                    And Application.Intersect(fc.AppliesTo.Cells, sheet.Range("PGUY").MergeArea) Is Nothing _
                    And Application.Intersect(fc.AppliesTo.Cells, sheet.Range("PGUY2").MergeArea) Is Nothing _
                    And Application.Intersect(fc.AppliesTo.Cells, sheet.Range("ROOMTOGUY").MergeArea) Is Nothing _
                    And Application.Intersect(fc.AppliesTo.Cells, sheet.Range("BONDED").MergeArea) Is Nothing _
                    And Application.Intersect(fc.AppliesTo.Cells, sheet.Range("BONDED").offset(-1, 1).MergeArea) Is Nothing _
                    And Application.Intersect(fc.AppliesTo.Cells, sheet.Range("TREE2").MergeArea) Is Nothing _
                    And Application.Intersect(fc.AppliesTo.Cells, sheet.Range("TREE2").offset(1, 0).MergeArea) Is Nothing _
                    And Application.Intersect(fc.AppliesTo.Cells, sheet.Range("TREE2").offset(2, -2).MergeArea) Is Nothing Then
                    fc.Interior.color = controlWs.Range("Default_Color").offset(0, 1).Interior.color
                    fc.Font.color = controlWs.Range("Default_Color").offset(0, 1).Font.color
                End If
            Next i
        End If
    Next sheet
    
    MsgBox "Highlight/Tab colors changed on all sheets!"
End Sub
