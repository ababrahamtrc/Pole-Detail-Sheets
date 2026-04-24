Attribute VB_Name = "DeletePDS"
Sub ClearWorkbook()
    Dim sheet As Worksheet
    
    On Error Resume Next
    
    Call LogMessage.SendLogMessage("ClearWorkbook")
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    answer = MsgBox("Are you sure you want to delete every pole detail sheet? (This can't be undone)", vbYesNoCancel + vbQuestion, "Confirmation")
    
    If answer = vbYes Then
        For Each sheet In ThisWorkbook.sheets
            If Utilities.IsPDS(sheet) Then
                sheet.Delete
            End If
        Next sheet
    End If
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub
