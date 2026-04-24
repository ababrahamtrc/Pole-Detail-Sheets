Attribute VB_Name = "CrewNotes"
Public Sub CrewNotesGenerator()
    On Error Resume Next

    Call LogMessage.SendLogMessage("crewNotesGenerator")

    Dim sheet As Worksheet: Set sheet = ThisWorkbook.ActiveSheet()
    If sheet.name = "4 Spans" Or sheet.name = "8 Spans" Or sheet.name = "12 Spans" Or sheet.Cells(2, 2).Value <> "Notification:" Then
        MsgBox "You need to have a pole detail sheet active to run this script."
        Exit Sub
    End If
    
    Unload CrewNotesGenerator_Form
    Call CrewNotesGenerator_Form.Initialize(sheet)
    CrewNotesGenerator_Form.Show vbModeless
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
