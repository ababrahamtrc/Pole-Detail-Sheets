Attribute VB_Name = "Tension"
Public Sub getTension()
    Call LogMessage.SendLogMessage("getTension")

    userName = LCase(Environ("USERNAME"))
    If userName <> "zschultz" And userName <> "aabraham" Then MsgBox "Can't run script": Exit Sub
    
    Call TensionCalculator_Form.Initialize
    TensionCalculator_Form.Show vbModeless
End Sub
