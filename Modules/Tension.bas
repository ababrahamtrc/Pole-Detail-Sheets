Attribute VB_Name = "Tension"
Public Sub getTension()
    Call LogMessage.SendLogMessage("getTension")

    UserName = LCase(Environ("USERNAME"))
    If UserName <> "zschultz" And UserName <> "aabraham" Then MsgBox "Can't run script": Exit Sub
    
    Call TensionCalculator_Form.Initialize
    TensionCalculator_Form.Show vbModeless
End Sub
