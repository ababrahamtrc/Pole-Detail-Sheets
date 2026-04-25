Attribute VB_Name = "GeneratePDS"
Public Sub GenerateAllPDS()
    Call LogMessage.SendLogMessage("GenerateAllPDS")
    
    Dim Project As Project: Set Project = New Project
    Call Project.extractImportDataFormat
    Call Project.createAndFillSheets
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

