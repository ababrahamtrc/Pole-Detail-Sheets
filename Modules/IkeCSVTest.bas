Attribute VB_Name = "IkeCSVTest"
Public globalPoleCsv As Scripting.Dictionary
Public globalAnchorCsv As Collection
Public globalSpanCsv As Collection
Public polesUsed As Scripting.Dictionary

Sub exportIkeCSV()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim project As project: Set project = New project
    Call project.extractFromSheets
    
    Set polesUsed = New Scripting.Dictionary
    
    Set poleCollections = findPoleGroups(project.poles)
    Dim poleCollection As Collection
    Dim i As Integer: i = 1
    For Each poleCollection In poleCollections
        Set globalPoleCsv = New Scripting.Dictionary
        Set globalAnchorCsv = New Collection
        Set globalSpanCsv = New Collection
        
        Call recurTest(project, poleCollection(1), New Scripting.Dictionary, New Collection, New Collection, poleCollection, True)
        
        If poleCollection.count > 1 Then
            Call generateCsv("polegroup" & i, globalPoleCsv, globalAnchorCsv, globalSpanCsv, project.notification)
            i = i + 1
        End If
    Next poleCollection
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Sub recurTest(project As project, pole As pole, poleCsv As Scripting.Dictionary, anchorCSV As Collection, spanCsv As Collection, poles As Collection, createCSV As Boolean)
    Dim poleLabel As String: poleLabel = pole.poleNumber
    
    If polesUsed.exists(poleLabel) Then Exit Sub
    polesUsed.Add poleLabel, Nothing
    
    Dim poleValues As Collection: Set poleValues = New Collection
    Call poleValues.Add(poleLabel)
    Call poleValues.Add(pole.latitude)
    Call poleValues.Add(pole.longitude)
    If (CInt(pole.height / 12)) > 0 Then
        Call poleValues.Add(CInt(pole.height / 12))
    Else
        Call poleValues.Add(40)
    End If
    If Utilities.OnlyNumbers(pole.Class) > 0 Then
        Call poleValues.Add(pole.Class)
    Else
        Call poleValues.Add(4)
    End If
    poleCsv.Add poleLabel, poleValues
    globalPoleCsv.Add poleLabel, poleValues
    
    Dim anchorValues As Collection
    Dim anchor As anchor
    For Each anchor In pole.anchors
        Set anchorValues = New Collection
        Call anchorValues.Add(poleLabel)
        Call anchorValues.Add(anchor.distance)
        Call anchorValues.Add(anchor.angle)
        anchorCSV.Add anchorValues
        globalAnchorCsv.Add anchorValues
    Next anchor
    
    Dim spanCounter As Integer: spanCounter = 0
    Dim Span As Span
    For Each Span In pole.spans
        If Span.utilWires.count <> Span.services.count Or Span.commWires.count <> Span.drops.count Then
            Dim otherPole As pole: Set otherPole = project.findPole(Span.otherPole)
            Dim otherLabel As String
            Set poleValues = New Collection
            If Span.otherPole <> "" Then
                otherLabel = Span.otherPole
                Call poleValues.Add(otherLabel)
            Else
                spanCounter = spanCounter + 1
                otherLabel = poleLabel & "-" & spanCounter
                Call poleValues.Add(otherLabel)
            End If
            
            Dim spanValues As Collection: Set spanValues = New Collection
            Call spanValues.Add(poleLabel)
            Call spanValues.Add(otherLabel)
            If Not poleCsv.exists(otherLabel) Then spanCsv.Add spanValues
            If Not polesUsed.exists(otherLabel) Then globalSpanCsv.Add spanValues
            
            If Not otherPole Is Nothing Then
                If Span.spanGuys.count > 0 Then
                    If Not polesUsed.exists(otherPole.poleNumber) Then
                        Call recurTest(project, otherPole, poleCsv, anchorCSV, spanCsv, poles, False)
                    End If
                Else
                    Call poleValues.Add(otherPole.latitude)
                    Call poleValues.Add(otherPole.longitude)
                    If (CInt(otherPole.height / 12)) > 0 Then
                        Call poleValues.Add(CInt(otherPole.height / 12))
                    Else
                        Call poleValues.Add(40)
                    End If
                    If Utilities.OnlyNumbers(pole.Class) > 0 Then
                        Call poleValues.Add(otherPole.Class)
                    Else
                        Call poleValues.Add(4)
                    End If
                    If Not poleCsv.exists(otherLabel) Then poleCsv.Add otherLabel, poleValues
                    
                    Call recurTest(project, otherPole, New Scripting.Dictionary, New Collection, New Collection, poles, True)
                End If
            Else
                results = latLonOffsetPoint(pole.latitude, pole.longitude, Span.angle, Span.distance)
                Call poleValues.Add(results(0))
                Call poleValues.Add(results(1))
                Call poleValues.Add(40)
                Call poleValues.Add(4)
                poleCsv.Add otherLabel, poleValues
                globalPoleCsv.Add otherLabel, poleValues
            End If
        End If
    Next Span
    
    If createCSV Then Call generateCsv("", poleCsv, anchorCSV, spanCsv, project.notification, True)
End Sub

Sub generateCsv(fileName As String, poleCsv As Scripting.Dictionary, anchorCSV As Collection, spanCsv As Collection, notification As String, Optional group As Boolean)
    If group Then
        fileName = ""
        Dim poleNumberCol As Collection: Set poleNumberCol = New Collection
        For Each poleNumber In poleCsv
            If InStr(poleNumber, "-") = 0 Then
                 poleNumberCol.Add poleNumber
            End If
           
        Next poleNumber
        fileName = Utilities.compactListString(poleNumberCol)
    End If
    
    If Dir(GetLocalPath(ThisWorkbook.Path) & "\" & notification & " - Ike Poleforeman CSVs", vbDirectory) = "" Then
        MkDir GetLocalPath(ThisWorkbook.Path) & "\" & notification & " - Ike Poleforeman CSVs"
    End If
    
    filePath = GetLocalPath(ThisWorkbook.Path) & "\" & notification & " - Ike Poleforeman CSVs\" & fileName & ".csv"
    
    FileNumber = FreeFile
    Open filePath For Output As #FileNumber

    Print #FileNumber, "[POLES]"
    Print #FileNumber, "Pole_Label,Pole_Latitude,Pole_Longitude,Pole_Elevation,Length,Class"
    For Each poleNumber In poleCsv
        Set poleValues = poleCsv(poleNumber)
        Print #FileNumber, poleValues(1) & "," & poleValues(2) & "," & poleValues(3) & ",1000," & poleValues(4) & "," & poleValues(5) & ","
    Next poleNumber
    
    Print #FileNumber, "[ANCHORS]"
    Print #FileNumber, "Pole_Label,Anchor_LeadLength,Anchor_Direction"
    For Each anchorValues In anchorCSV
        Print #FileNumber, anchorValues(1) & "," & anchorValues(2) & "," & anchorValues(3)
    Next anchorValues
    
    Print #FileNumber, "[SPANS]"
    Print #FileNumber, "Pole1_Label,Pole2_Label"
    For Each spanValues In spanCsv
        Print #FileNumber, spanValues(1) & "," & spanValues(2) & ","
    Next spanValues
    
    Print #FileNumber, "[GROUNDLINE_TOPOGRAPHY]"
    Print #FileNumber, "Pole1_Label,Pole2_Label,Distance_From_Pole1,Elevation"
    
    Close #FileNumber
    
    Dim csvWb As Workbook: Set csvWb = Workbooks.Open(filePath)
    csvWb.save
    csvWb.Close
End Sub

Function latLonOffsetPoint(lat As Double, lon As Double, bearing As Double, distance As Double) As Variant()
    Dim PI As Double
    Dim lat1 As Double, lon1 As Double, brng As Double
    Dim lat2 As Double, lon2 As Double
    Dim dR As Double
    Dim radius As Double:  radius = 20925721.784777
    PI = 4 * Atn(1)
    
    lat1 = lat * PI / 180
    lon1 = lon * PI / 180
    brng = bearing * PI / 180
    
    dR = distance / radius

    lat2 = Asin(Sin(lat1) * Cos(dR) + Cos(lat1) * Sin(dR) * Cos(brng))
    
    lon2 = lon1 + Atn2(Cos(dR) - Sin(lat1) * Sin(lat2), Sin(brng) * Sin(dR) * Cos(lat1))
    
    lat2 = lat2 * 180 / PI
    lon2 = lon2 * 180 / PI
    
    latLonOffsetPoint = Array(lat2, lon2)
End Function

Function Asin(x As Double) As Double
    If Abs(x) = 1 Then
        Asin = Sgn(x) * PI / 2
    Else
        Asin = Atn(x / Sqr(1 - x * x))
    End If
End Function

