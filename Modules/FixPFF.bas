Attribute VB_Name = "FixPFF"
Public Sub FixPoleForemanJSON()
    Call LogMessage.SendLogMessage("FixPFFJSON")

    Dim fileDiag As FileDialog
    Set fileDiag = Application.FileDialog(msoFileDialogFilePicker)
    With fileDiag
        .AllowMultiSelect = False
        .Title = "Select a pole foreman json"
        .Filters.Add "Pole Foreman File", "*.json", 1
        .InitialFileName = ThisWorkbook.path & Application.PathSeparator
        If .Show = -1 Then path = Application.FileDialog(msoFileDialogFilePicker).SelectedItems.item(1) Else Exit Sub
    End With
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set tso = fso.OpenTextFile(path)
    Set json = JsonConverter.ParseJson(tso.ReadAll)
    tso.Close
    Set tso = Nothing
    Set fso = Nothing
    
    Dim pole As pole
    Dim project As project
    
    counter = 0
    
    Dim PFNameMapping As Scripting.Dictionary: Set PFNameMapping = New Scripting.Dictionary
    Call InitializePFNameMapping(PFNameMapping)
    
    For Each jsonPole In json
        Dim poleID As String
        poleID = jsonPole("Structure")("Pole")("PoleNumber")
        If poleID <> "" Then
            found = False
            Set project = New project
            project.extractFromSheets
            For Each pole In project.poles
                If pole.existingCEID = poleID Then
                    jsonPole("Structure")("Pole")("PoleNumber") = "M1P" & pole.poleNumber & "_" & pole.existingCEID & "_" & correctFileName(project.permit) & "_"
                    found = True
                    Exit For
                End If
            Next pole
            
            If Not found Then
                Set project = New project
                project.extractImportDataFormat
                For Each pole In project.poles
                If pole.existingCEID = poleID Then
                    jsonPole("Structure")("Pole")("PoleNumber") = "M1P" & pole.poleNumber & "_" & pole.existingCEID & "_" & correctFileName(project.permit) & "_"
                    Exit For
                End If
                If pole.gisCEID = poleID Then
                    Dim Pole2 As pole: Set Pole2 = New pole
                    Dim project2 As project: Set project2 = New project
                    Call project2.extractFromSheets
                    Set Pole2 = project2.findPole(pole.poleNumber)
                    If Pole2 Is Nothing Then
                        If Utilities.isCEID(pole.existingCEID) Or pole.existingCEID = "FOREIGN" Then jsonPole("Structure")("Pole")("PoleNumber") = "M1P" & pole.poleNumber & "_" & pole.existingCEID & "_" & correctFileName(project.permit) & "_"
                    Else
                        If Utilities.isCEID(Pole2.existingCEID) Or Pole2.existingCEID = "FOREIGN" Then jsonPole("Structure")("Pole")("PoleNumber") = "M1P" & pole.poleNumber & "_" & Pole2.existingCEID & "_" & correctFileName(project.permit) & "_"
                    End If
                    Exit For
                End If
            Next pole
            End If
        End If
        
        Dim estimatedAGL As Double
        If jsonPole("Structure")("Pole")("Length") < 40 Then
            estimatedAGL = jsonPole("Structure")("Pole")("Length") - 6
        Else
            estimatedAGL = (jsonPole("Structure")("Pole")("Length") * 0.9) - 2
        End If
        
        If jsonPole("Structure")("Pole")("AGL") > estimatedAGL And jsonPole("Structure")("Pole")("AGL") - estimatedAGL < 1.5 Then
            jsonPole("Structure")("Pole")("AGL") = estimatedAGL
        End If
        
        If jsonPole("Structure").Exists("Services") And Not IsNull(jsonPole("Structure")("Services")) Then
            For Each jsonService In jsonPole("Structure")("Services")
                If Not IsNull(jsonService) Then
                    If jsonService("Length") < 10 Then
                        jsonService("Length") = 10
                    ElseIf jsonService("Length") > 130 Then
                        jsonService("Length") = 130
                    End If
                End If
            Next jsonService
        End If
        
        For Each jsonSpan In jsonPole("Structure")("Spans")
            For Each jsonCircuit In jsonSpan("Power")("Circuit")
                If jsonCircuit.Exists("Primary") Then
                    If PFNameMapping.Exists((jsonCircuit("Primary")("ConductorDescription"))) Then
                        jsonCircuit("Primary")("RulingSpan") = PFNameMapping(jsonCircuit("Primary")("ConductorDescription"))
                    End If
                End If
                If jsonCircuit.Exists("Neutral") Then
                    If PFNameMapping.Exists(jsonCircuit("Neutral")("ConductorDescription")) Then
                        jsonCircuit("Neutral")("RulingSpan") = PFNameMapping(jsonCircuit("Neutral")("ConductorDescription"))
                    End If
                End If
                If jsonCircuit.Exists("Secondary") Then
                    If PFNameMapping.Exists(jsonCircuit("Secondary")("ConductorDescription")) Then
                        jsonCircuit("Secondary")("RulingSpan") = PFNameMapping(jsonCircuit("Secondary")("ConductorDescription"))
                    End If
                End If
            Next jsonCircuit
            
            Dim oppositeBearings As Scripting.Dictionary: Set oppositeBearings = New Scripting.Dictionary
            Dim oppositeSpan As Object: Set oppositeSpan = Nothing
            Dim oppositeBearing As Double
            Dim bearingDifference As Double
            Dim closestBearingDifference As Double: closestBearingDifference = 0
            Dim bearingKey As String: bearingKey = ""
            Dim otherbearingKey As String: otherbearingKey = ""
            If jsonSpan.Exists("Communication") And Not IsNull(jsonSpan("Communication")) Then
                For Each jsonCommunication In jsonSpan("Communication")
                    For Each jsonOtherSpan In jsonPole("Structure")("Spans")
                        If jsonOtherSpan.Exists("Communication") And Not IsNull(jsonOtherSpan("Communication")) Then
                            If jsonSpan("Length") <> jsonOtherSpan("Length") And jsonSpan("Bearing") <> jsonOtherSpan("Bearing") Then
                                If jsonSpan("Bearing") >= 3.141592 Then
                                    oppositeBearing = jsonSpan("Bearing") - 3.141592
                                Else
                                    oppositeBearing = jsonSpan("Bearing") + 3.141592
                                End If
                                
                                bearing = jsonOtherSpan("Bearing")
                                bearingDifference = Abs(bearing - oppositeBearing)
                                If bearingDifference > (3.141592 / 3) Then bearingDifference = Abs(bearing - (oppositeBearing + (2 * 3.141592)))
                                If bearingDifference <= (3.141592 / 3) Then
                                    For Each jsonOtherCommunication In jsonOtherSpan("Communication")
                                        If jsonCommunication("Owner") = jsonOtherCommunication("Owner") And Abs(jsonCommunication("Height") - jsonOtherCommunication("Height")) < 2 Then
                                            If closestBearingDifference >= bearingDifference Or closestBearingDifference = 0 Then
                                                Set oppositeSpan = jsonOtherSpan
                                                closestBearingDifference = bearingDifference
                                                bearingKey = JsonConverter.ConvertToJson(jsonSpan) & JsonConverter.ConvertToJson(jsonCommunication)
                                                otherbearingKey = JsonConverter.ConvertToJson(jsonOtherSpan) & JsonConverter.ConvertToJson(jsonOtherCommunication)
                                                oppositeBearings(bearingKey) = bearingDifference
                                            End If
                                        End If
                                    Next jsonOtherCommunication
                                End If
                            End If
                        End If
                    Next jsonOtherSpan
                    
                    For Each jsonOtherSpan In jsonPole("Structure")("Spans")
                        If jsonSpan("Length") <> jsonOtherSpan("Length") And jsonSpan("Bearing") <> jsonOtherSpan("Bearing") Then
                            If jsonSpan("Bearing") > 3.141592 Then
                                oppositeBearing = jsonSpan("Bearing") - 3.141592
                            Else
                                oppositeBearing = jsonSpan("Bearing") + 3.141592
                            End If
                            
                            bearing = jsonOtherSpan("Bearing")
                            bearingDifference = Abs(bearing - oppositeBearing)
                            If bearingDifference > (3.141592 / 3) Then bearingDifference = Abs(bearing - (oppositeBearing + (2 * 3.141592)))
                            If bearingDifference <= (3.141592 / 3) Then
                                For Each jsonOtherCommunication In jsonSpan("Communication")
                                    If jsonCommunication("Owner") = jsonOtherCommunication("Owner") And Abs(jsonCommunication("Height") - jsonOtherCommunication("Height")) < 2 Then
                                        If closestBearingDifference <> bearingDifference Then
                                            bearingKey = JsonConverter.ConvertToJson(jsonSpan) & JsonConverter.ConvertToJson(jsonCommunication)
                                            If oppositeBearings.Exists(bearingKey) And oppositeBearings(bearingKey) <> bearingDifference Then
                                                jsonOtherCommunication("RulingSpan") = Application.WorksheetFunction.Ceiling(CDbl(jsonOtherSpan("Length")), 50)
                                            End If
                                        End If
                                    End If
                                Next jsonOtherCommunication
                            End If
                        End If
                    Next jsonOtherSpan
                    
                    If Not oppositeSpan Is Nothing And (Not oppositeBearings.Exists(otherbearingKey) Or (oppositeBearings.Exists(otherbearingKey) And oppositeBearings(otherbearingKey) = closestBearingDifference)) Then
                        jsonCommunication("RulingSpan") = Application.WorksheetFunction.Ceiling(CDbl((jsonSpan("Length")) + CDbl(oppositeSpan("Length"))) / 2, 50)
                    Else
                        jsonCommunication("RulingSpan") = Application.WorksheetFunction.Ceiling(CDbl(jsonSpan("Length")), 50)
                    End If
                    
                Next jsonCommunication
            End If
        Next jsonSpan
    Next jsonPole
    
    Dim jsonText As String
    jsonText = JsonConverter.ConvertToJson(json, Whitespace:=2)
    
    Dim file As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set tso = fso.CreateTextFile(path, True, False)
    
    tso.Write jsonText
    tso.Close
    
    MsgBox "Done"
End Sub

Private Sub InitializePFNameMapping(PFNameMapping As Scripting.Dictionary)
    'Primary/Neutral sizes
    PFNameMapping("4 ACSR (7/1)") = 380
    PFNameMapping("2 ACSR (7/1)") = "440"
    PFNameMapping("2 ACSR (7/1) HDPE (45)") = "380"
    PFNameMapping("1/0 ACSR (6/1)") = 440
    PFNameMapping("1/0 ACSR (6/1) HDPE (60)") = "380"
    PFNameMapping("3/0 ACSR (6/1)") = "440"
    PFNameMapping("3/0 ACSR (6/1) HDPE (60)") = "380"
    PFNameMapping("336 ACSR (26/7)") = "300"
    PFNameMapping("795 ACSR (26/7)") = "350"
    PFNameMapping("795 AL (37)") = "230"
    PFNameMapping("80 KCMIL ACSR (8/1)") = "840"
    PFNameMapping("1/0 AERIAL SPACER CABLE") = "35"
    PFNameMapping("336 AERIAL SPACER CABLE") = "35"
    PFNameMapping("795 AERIAL SPACER CABLE") = "35"
    PFNameMapping("4 WP ACSR") = "250"
    PFNameMapping("350 AAC (19)") = "150"
    PFNameMapping("6 CU (1)") = "150"
    PFNameMapping("4 CU (1)") = "200"
    PFNameMapping("3 CU (1)") = "250"
    PFNameMapping("2 CU (1)") = "280"
    'PFNameMapping("4 ACSR (7/1) - SLACK 3PH") = "100"
    'PFNameMapping("2 ACSR (7/1) - SLACK 3PH") = "100"
    'PFNameMapping("1/0 ACSR (6/1) - SLACK 3PH") = "100"
    'PFNameMapping("3/0 ACSR (6/1) - SLACK 3PH") = "100"
    'PFNameMapping("336 ACSR (26/7) - SLACK 3PH") = "75"
    'PFNameMapping("795 ACSR (26/7) - SLACK 3PH") = "75"
    'PFNameMapping("795 AL (37) - SLACK 3PH") = "75"
    PFNameMapping("1/0 ACSR (6/1) TREE WIRE") = "300"
    PFNameMapping("336 AAC (19) TREE WIRE") = "200"
    'PFNameMapping("4 ACSR (7/1) - SLACK 1PH") = "125"
    'PFNameMapping("2 ACSR (7/1) - SLACK 1PH") = "125"
    'PFNameMapping("1/0 ACSR (6/1) - SLACK 1PH") = "125"
    'PFNameMapping("3/0 ACSR (6/1) - SLACK 1PH") = "100"
    'PFNameMapping("336 ACSR (26/7) - SLACK 1PH") = "100"
    'PFNameMapping("795 ACSR (26/7) - SLACK 1PH") = "75"
    'PFNameMapping("795 AL (37) - SLACK 1PH") = "75"
    
     'Neutral sizes
    PFNameMapping("052 (1/0-1Ø)") = "350"
    PFNameMapping("052 (336-1Ø)") = "350"
    PFNameMapping("052 (1/0-3Ø)") = "250"
    PFNameMapping("052 (336-3Ø)") = "200"
    PFNameMapping("7#6 (795-3Ø)") = "150"
    PFNameMapping("052 AWA Shield") = "300"
    PFNameMapping("7#6") = "150"
    PFNameMapping("350 AAC") = "150"
    
    'Secondary sizes
    PFNameMapping("4-4-4-4 ACSR QX") = "180"
    PFNameMapping("2-2-2-2 ACSR QX") = "180"
    PFNameMapping("1/0-1/0-1/0-1/0 ACSR QX") = "180"
    PFNameMapping("3/0-3/0-3/0-3/0 ACSR QX") = "180"
    'PFNameMapping("4 WP ACSR") = "250" Duplicate
    PFNameMapping("2 WP ACSR") = "350"
    PFNameMapping("1/0 HDPE ACSR") = "350"
    PFNameMapping("3/0 HDPE ACSR") = "350"
    'PFNameMapping("6-6 ACSR DX SERV") = "130"
    'PFNameMapping("2-2 ACSR DX SERV") = "130"
    'PFNameMapping("4-4-4 ACSR TX SERV") = "130"
    'PFNameMapping("1/0-2-1/0 ACSR TX SERV") = "130"
    'PFNameMapping("3/0-1/0-3/0 ACSR TX SERV") = "130"
    'PFNameMapping("4-4-4-4 ACSR QX SERV") = "130"
    'PFNameMapping("2-2-2-2 ACSR QX SERV") = "130"
    'PFNameMapping("1/0-1/0-1/0-1/0 ACSR QX SERV") = "130"
    'PFNameMapping("3/0-3/0-3/0-3/0 ACSR QX SERV") = "130"
    PFNameMapping("6-6 ACSR DX") = "200"
    PFNameMapping("2-2 ACSR DX") = "175"
    PFNameMapping("4-4-4 ACSR TX") = "200"
    PFNameMapping("1/0-2-1/0 ACSR TX") = "200"
    PFNameMapping("3/0-1/0-3/0 ACSR TX") = "200"
    PFNameMapping("2-4-2 AWAC TX") = "300"
    PFNameMapping("1/0-4-1/0 AWAC TX") = "300"
    PFNameMapping("3/0-2-3/0 AWAC TX") = "300"
    'PFNameMapping("6 CU (1)") = "150" DUPLICATE
End Sub
