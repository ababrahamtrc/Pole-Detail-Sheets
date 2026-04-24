Attribute VB_Name = "GeneratePDS"
Public Sub GenerateAllPDS()
    Call LogMessage.SendLogMessage("GenerateAllPDS")
    
    Dim project As project: Set project = New project
    Call project.extractImportDataFormat
    Call project.createAndFillSheets
    Call NJUNSCodes.clearNJUNSCodes
End Sub
    
Public Sub generateSelectedPDS()
    Call LogMessage.SendLogMessage("GenerateSelectedPDS")
    
    Call GenerateSinglePDS_Form.Initialize
    GenerateSinglePDS_Form.Show vbModeless
End Sub

Public Sub generateBrandNewPDS()
    Call LogMessage.SendLogMessage("GenerateBrandNewPDS")
    
    Call NewPole_Form.Initialize
    NewPole_Form.Show vbModeless
End Sub

