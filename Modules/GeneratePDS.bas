Attribute VB_Name = "GeneratePDS"
Public Sub GenerateAllPDS()
    Call LogMessage.SendLogMessage("GenerateAllPDS")
    
    ThisWorkbook.sheets("4 Spans").Activate
    ThisWorkbook.sheets("4 Spans").Range("A1").Select
    
    ThisWorkbook.sheets("8 Spans").Activate
    ThisWorkbook.sheets("8 Spans").Range("A1").Select
    
    ThisWorkbook.sheets("12 Spans").Activate
    ThisWorkbook.sheets("12 Spans").Range("A1").Select
    
    Dim project As project: Set project = New project
    Call project.extractImportDataFormat
    Call project.createAndFillSheets
    Call NJUNSCodes.clearNJUNSCodes
End Sub
    
Public Sub generateSelectedPDS()
    Call LogMessage.SendLogMessage("GenerateSelectedPDS")
    
    ThisWorkbook.sheets("4 Spans").Activate
    ThisWorkbook.sheets("4 Spans").Range("A1").Select
    
    ThisWorkbook.sheets("8 Spans").Activate
    ThisWorkbook.sheets("8 Spans").Range("A1").Select
    
    ThisWorkbook.sheets("12 Spans").Activate
    ThisWorkbook.sheets("12 Spans").Range("A1").Select
    
    Call GenerateSinglePDS_Form.Initialize
    GenerateSinglePDS_Form.Show vbModeless
End Sub

Public Sub generateBrandNewPDS()
    Call LogMessage.SendLogMessage("GenerateBrandNewPDS")
    
    ThisWorkbook.sheets("4 Spans").Activate
    ThisWorkbook.sheets("4 Spans").Range("A1").Select
    
    ThisWorkbook.sheets("8 Spans").Activate
    ThisWorkbook.sheets("8 Spans").Range("A1").Select
    
    ThisWorkbook.sheets("12 Spans").Activate
    ThisWorkbook.sheets("12 Spans").Range("A1").Select
    
    Call NewPole_Form.Initialize
    NewPole_Form.Show vbModeless
End Sub

