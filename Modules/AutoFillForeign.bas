Attribute VB_Name = "AutoFillForeign"
Public Sub FillForeignPole()

    On Error Resume Next
    
    Call LogMessage.SendLogMessage("FillForeignPole")
    
    Dim sheet As Worksheet

    Set sheet = ThisWorkbook.ActiveSheet()
    If sheet.name = "4 Spans" Or sheet.name = "8 Spans" Or sheet.name = "12 Spans" Or sheet.Cells(2, 2).Value <> "Notification:" Then
        MsgBox "You need to have a pole detail sheet active to run this script."
        Exit Sub
    End If
    
    otherPoleOwner = Trim(sheet.Range("OTHERPOLEOWNER").text)
    If otherPoleOwner = "" Then
        MsgBox "Please put a foreign pole owner near the top of the sheet (next to the Other checkbox near the CE Pole checkbox)"
        Exit Sub
    End If
    
    answer = MsgBox("Do you want to overwrite the values on this sheet with value for a foreign pole? (This cannot be undone)", vbYesNoCancel + vbQuestion, "Confirmation")
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    If answer = vbYes Then
        sheet.Range("ASIS").Value = "FOREIGN"
        sheet.Range("NEWAPP").Value = "FOREIGN"
        sheet.Range("SUMSHEET9").Value = "TRUE"
        sheet.Range("SUMSHEET12").Value = "N/A"
        sheet.Range("SUMSHEET14").Value = "APPLY TO " & otherPoleOwner
        sheet.Range("CMRF1").Value = "APPLY TO " & otherPoleOwner
        sheet.Range("CMRF2").Value = "APPLY TO " & otherPoleOwner
        sheet.Range("CMRF3").Value = "APPLY TO " & otherPoleOwner
        If Utilities.isCEID(sheet.Range("CEID")) Then
            If InStr(sheet.Range("NOTES"), "Old GIS CEID: " & sheet.Range("CEID").Value) = 0 Then
                If sheet.Range("NOTES") <> "" Then sheet.Range("NOTES") = vbLf & sheet.Range("NOTES")
                sheet.Range("NOTES") = "Old GIS CEID: " & sheet.Range("CEID").Value
            End If
        End If
        sheet.Range("CEID") = "FOREIGN"
        If Trim(sheet.Range("CEID").Value) <> "FOREIGN" Then MsgBox "Warning, CEID should be FOREIGN on foreign poles"
    End If
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub



