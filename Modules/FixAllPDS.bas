Attribute VB_Name = "FixAllPDS"
Sub AbbreviateCrewNotes()
    Dim project As project: Set project = New project
    Call project.extractFromSheets
    
    Dim pole As pole
    Dim sheet As Worksheet
    For Each pole In project.poles
        Dim alt1 As String: alt1 = pole.alt1
        Dim alt2 As String: alt2 = pole.alt2
        Dim alt3 As String: alt3 = pole.alt3
        Call Utilities.applyStandardAbbreviations(alt1)
        Call Utilities.applyStandardAbbreviations(pole.alt2)
        Call Utilities.applyStandardAbbreviations(pole.alt3)
        
        Set sheet = Utilities.GetPDS(pole.poleNumber)
        sheet.Range("ALTONE").Value = alt1
        sheet.Range("ALTTWO").Value = alt2
        sheet.Range("ALTTHREE").Value = alt3
    Next pole
    
    MsgBox "Abbreviated all crewnotes"
End Sub

Sub FixHeaders()
    Dim project As project: Set project = New project
    Call project.extractFromSheets
    
    If project.Notification = "" Then project.Notification = InputBox("Enter Notification:")
    If project.permit = "" Then project.permit = InputBox("Enter Permit:")
    If project.applicant = "" Then project.applicant = InputBox("Enter Applicant:")
    If project.county = "" Then project.county = InputBox("Enter County:")
    If project.township = "" Then project.township = InputBox("Enter Town:")
    
    Dim pole As pole
    Dim sheet As Worksheet
    For Each pole In project.poles
        Set sheet = Utilities.GetPDS(pole.poleNumber)
        sheet.Range("NOTIFICATION").Value = project.Notification
        sheet.Range("PERMIT").Value = project.permit
        sheet.Range("APPLICANT").Value = project.applicant
        sheet.Range("COUNTY").Value = project.county
        sheet.Range("TWP").Value = project.township
    Next pole
    
End Sub
