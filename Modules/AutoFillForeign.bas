Attribute VB_Name = "AutoFillForeign"
Public Sub FillForeignPole()

    On Error Resume Next
    
    Call LogMessage.SendLogMessage("FillForeignPole")
    
    Dim sheet As Worksheet

    Set sheet = ThisWorkbook.ActiveSheet()
    If sheet.name = "4 Spans" Or sheet.name = "8 Spans" Or sheet.name = "12 Spans" Or sheet.cells(2, 2).value <> "Notification:" Then
        MsgBox "You need to have a pole detail sheet active to run this script."
        Exit Sub
    End If
    
    OTHERPOLEOWNER = Trim(sheet.Range("OTHERPOLEOWNER").text)
    If OTHERPOLEOWNER = "" Then
        MsgBox "Please put a foreign pole owner near the top of the sheet (next to the Other checkbox near the CE Pole checkbox)"
        Exit Sub
    End If
    
    answer = MsgBox("Do you want to overwrite the values on this sheet with value for a foreign pole? (This cannot be undone)", vbYesNoCancel + vbQuestion, "Confirmation")
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    If answer = vbYes Then
        sheet.Range("ASIS").value = "FOREIGN"
        sheet.Range("NEWAPP").value = "FOREIGN"
        sheet.Range("SUMSHEET9").value = "TRUE"
        sheet.Range("SUMSHEET12").value = "N/A"
        sheet.Range("SUMSHEET14").value = "APPLY TO " & OTHERPOLEOWNER
        sheet.Range("CMRF1").value = "APPLY TO " & OTHERPOLEOWNER
        sheet.Range("CMRF2").value = "APPLY TO " & OTHERPOLEOWNER
        sheet.Range("CMRF3").value = "APPLY TO " & OTHERPOLEOWNER
        If Utilities.isCEID(sheet.Range("CEID")) Then
            If InStr(sheet.Range("NOTES"), "Old GIS CEID: " & sheet.Range("CEID").value) = 0 Then
                If sheet.Range("NOTES") <> "" Then sheet.Range("NOTES") = vbLf & sheet.Range("NOTES")
                sheet.Range("NOTES") = "Old GIS CEID: " & sheet.Range("CEID").value
            End If
        End If
        sheet.Range("CEID") = "FOREIGN"
        If Trim(sheet.Range("CEID").value) <> "FOREIGN" Then MsgBox "Warning, CEID should be FOREIGN on foreign poles"
    End If
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub



