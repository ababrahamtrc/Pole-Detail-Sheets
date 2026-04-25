Attribute VB_Name = "ImportData"
Public Sub ImportJSONData()
    Call LogMessage.SendLogMessage("ImportData")

    Dim path, msg As String: path = "": msg = ""
    Dim fileDiag As FileDialog: Set fileDiag = Application.FileDialog(msoFileDialogFilePicker)
    Dim json As Object
    
    
    With fileDiag
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add Description:="Import Files", Extensions:="*.json"
        .Title = "Select the File ... "
        .InitialFileName = ThisWorkbook.path & Application.PathSeparator
        If .Show = -1 Then path = Application.FileDialog(msoFileDialogFilePicker).SelectedItems.item(1) Else Exit Sub
    End With
    
    On Error Resume Next
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    ProgressBar_Form.Show vbModeless
    
    ProgressBar_Form.Label1.caption = "Importing Project data... Please wait..."
    ProgressBar_Form.Repaint
    
    If InStr(path, ".json") Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set tso = fso.OpenTextFile(path)
        If tso Is Nothing Then
            MsgBox "Sometimes theres an issue importing the json from the onedrive when the Control file is also on the onedrive, try again after moving the project json to a local directory (like downloads)"
            ProgressBar_Form.Hide
            Application.ScreenUpdating = True
            Application.EnableEvents = True
            Exit Sub
        End If
        Set json = JsonConverter.ParseJson(tso.ReadAll)
        tso.Close
        Set tso = Nothing
        Set fso = Nothing
        
        If TypeName(json) = "Collection" Then
            MsgBox "Please select a spidacalc project json, you can find this by going to ""Project>Export>Project Json..."" on your spidacalc file."
            ProgressBar_Form.Hide
            Application.ScreenUpdating = True
            Application.EnableEvents = True
            Exit Sub
        End If
        
        Dim jsonType As String: jsonType = ""
        
        If json.Exists("date") Then
            jsonType = "Spida"
        ElseIf json.Exists("connections") Then
            jsonType = "Katapult"
        Else
            MsgBox "Please select a spidacalc project json, you can find this by going to ""Project>Export>Project Json..."" on your spidacalc file."
            ProgressBar_Form.Hide
            Application.ScreenUpdating = True
            Application.EnableEvents = True
            Exit Sub
        End If
    
        Call ClearData
        
        Dim Wire As Wire
        Dim Project As Project:
        
        If (jsonType = "Spida") Then
            Set Project = UtilitiesSpidaCalc.InitProjectFromSpidaJson(json)
            Call Project.fillImportDataFormat
        ElseIf jsonType = "Katapult" Then
            Set Project = UtilitiesKatapult.InitProjectFromKatapultJson(json)
            Call Project.fillImportDataFormat
        End If
    End If
    
    ProgressBar_Form.Hide
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub

Public Sub ClearData()

    ThisWorkbook.sheets("Collection").Cells.Clear
    ThisWorkbook.sheets("Job Info").Cells.Clear
    ThisWorkbook.sheets("Span").Cells.Clear
    ThisWorkbook.sheets("Span.Power Circuit").Cells.Clear
    ThisWorkbook.sheets("Span.Communication").Cells.Clear
    ThisWorkbook.sheets("Anchor").Cells.Clear
    ThisWorkbook.sheets("Anchor.Guys").Cells.Clear
    ThisWorkbook.sheets("Equipment").Cells.Clear
    ThisWorkbook.sheets("Control").Range("PHOTODIR").Value = ""
    
End Sub
