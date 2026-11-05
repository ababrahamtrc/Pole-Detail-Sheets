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
        sheet.Range("ALTONE").value = alt1
        sheet.Range("ALTTWO").value = alt2
        sheet.Range("ALTTHREE").value = alt3
    Next pole
    
    MsgBox "Abbreviated all crewnotes"
End Sub

Sub FixHeaders()
    Dim project As project: Set project = New project
    Call project.extractFromSheets
    
    If project.notification = "" Then project.notification = InputBox("Enter Notification:")
    If project.Permit = "" Then project.Permit = InputBox("Enter Permit:")
    If project.applicant = "" Then project.applicant = InputBox("Enter Applicant:")
    If project.county = "" Then project.county = InputBox("Enter County:")
    If project.township = "" Then project.township = InputBox("Enter Town:")
    
    Dim pole As pole
    Dim sheet As Worksheet
    For Each pole In project.poles
        Set sheet = Utilities.GetPDS(pole.poleNumber)
        sheet.Range("NOTIFICATION").value = project.notification
        sheet.Range("PERMIT").value = project.Permit
        sheet.Range("APPLICANT").value = project.applicant
        sheet.Range("COUNTY").value = project.county
        sheet.Range("TWP").value = project.township
    Next pole
    
End Sub
