Attribute VB_Name = "MicrostationUtilities"
Function getElectricJson(poles As Collection, layer As Integer, token As String) As Object
    Dim lowestLatitude As Double
    Dim lowestLongitude As Double
    Dim highestLatitude As Double
    Dim highestLongitude As Double
    
    Dim pole As pole
    For Each pole In poles
        If lowestLatitude = 0 Or pole.latitude < lowestLatitude Then lowestLatitude = pole.latitude
        If lowestLongitude = 0 Or pole.longitude < lowestLongitude Then lowestLongitude = pole.longitude
        If highestLatitude = 0 Or pole.latitude > highestLatitude Then highestLatitude = pole.latitude
        If highestLongitude = 0 Or pole.longitude > highestLongitude Then highestLongitude = pole.longitude
    Next pole
    
    Dim radius As Integer: radius = 1000
    
    results = LatLonToMI2253(lowestLatitude, lowestLongitude)
    x1 = results(0) - radius
    y1 = results(1) - radius
    
    results = LatLonToMI2253(highestLatitude, highestLongitude)
    x2 = results(0) + radius
    y2 = results(1) + radius
    
    Dim bbox1 As String: bbox1 = x1 & "," & y1
    Dim bbox2 As String: bbox2 = x2 & "," & y2
    Dim bbox As String: bbox = bbox1 & "," & bbox2
    
    Dim url As String: url = "https://gis.consumersenergy.com/mapping/rest/services/Electric/Electric_PUB/MapServer/" & layer & "/query?where=1=1&outFields=*&returnGeometry=true&geometryType=esriGeometryEnvelope&geometry=" & bbox & "&spatialRel=esriSpatialRelIntersects&inSR=2253&outSR=2253+&f=json&token=" & token

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", url, False
    http.setRequestHeader "Authorization", "Bearer " & token
    http.setRequestHeader "Content-Type", "application/json"
    http.send
    
    If http.Status = 200 Then
        Set getElectricJson = JsonConverter.ParseJson(http.responseText)
    Else
        Set getElectricJson = Nothing
        Debug.Print "Error: " & http.Status & " - " & http.statusText
    End If
End Function

Function getOtherPole(x As Double, y As Double, radius As Integer, token) As Object
    x1 = x - radius
    y1 = y - radius
    
    x2 = x + radius
    y2 = y + radius
    
    Dim bbox1 As String: bbox1 = x1 & "," & y1
    Dim bbox2 As String: bbox2 = x2 & "," & y2
    Dim bbox As String: bbox = bbox1 & "," & bbox2
    
    Dim url As String: url = "https://gis.consumersenergy.com/mapping/rest/services/Electric/Electric_PUB/MapServer/3/query?where=1=1&outFields=*&returnGeometry=true&geometryType=esriGeometryEnvelope&geometry=" & bbox & "&spatialRel=esriSpatialRelIntersects&inSR=2253&outSR=2253+&f=json&token=" & token

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", url, False
    http.setRequestHeader "Authorization", "Bearer " & token
    http.setRequestHeader "Content-Type", "application/json"
    http.send
    
    If http.Status = 200 Then
        Set getOtherPole = JsonConverter.ParseJson(http.responseText)
    Else
        Set getOtherPole = Nothing
        Debug.Print "Error: " & http.Status & " - " & http.statusText
    End If
    
    Set http = Nothing
End Function

Function getROWJSON(poles As Collection, layer As Integer, token As String) As Object
    Dim lowestLatitude As Double
    Dim lowestLongitude As Double
    Dim highestLatitude As Double
    Dim highestLongitude As Double
    
    Dim pole As pole
    For Each pole In poles
        If lowestLatitude = 0 Or pole.latitude < lowestLatitude Then lowestLatitude = pole.latitude
        If lowestLongitude = 0 Or pole.longitude < lowestLongitude Then lowestLongitude = pole.longitude
        If highestLatitude = 0 Or pole.latitude > highestLatitude Then highestLatitude = pole.latitude
        If highestLongitude = 0 Or pole.longitude > highestLongitude Then highestLongitude = pole.longitude
    Next pole
    
    Dim radius As Integer
    If poles.count > 1 Then
        radius = 200
    Else
        radius = 500
    End If
    
    results = LatLonToMI2253(lowestLatitude, lowestLongitude)
    x1 = results(0) - radius
    y1 = results(1) - radius
    
    results = LatLonToMI2253(highestLatitude, highestLongitude)
    x2 = results(0) + radius
    y2 = results(1) + radius
    
    Dim bbox1 As String: bbox1 = x1 & "," & y1
    Dim bbox2 As String: bbox2 = x2 & "," & y2
    Dim bbox As String: bbox = bbox1 & "," & bbox2
    
    Dim url As String: url = "https://gis.consumersenergy.com/mapping/rest/services/Landbase/Landbase_Grids_Boundaries_PUB/MapServer/" & layer & "/query?where=1=1&outFields=*&returnGeometry=true&geometryType=esriGeometryEnvelope&geometry=" & bbox & "&spatialRel=esriSpatialRelIntersects&inSR=2253&outSR=2253+&f=json&token=" & token

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", url, False
    http.setRequestHeader "Authorization", "Bearer " & token
    http.setRequestHeader "Content-Type", "application/json"
    http.send
    
    If http.Status = 200 Then
        Set getROWJSON = JsonConverter.ParseJson(http.responseText)
    Else
        Set getROWJSON = Nothing
        Debug.Print "Error: " & http.Status & " - " & http.statusText
    End If
    
    Set http = Nothing
End Function

Public Function getItems(pole As pole, ByVal i As String, uttype As String) As Collection
    Dim items As Collection: Set items = New Collection
    Dim sheet As Worksheet: Set sheet = Utilities.GetPDS(pole.poleNumber)
    
    For j = 0 To 100
        If sheet.Range("UTTYPE").offset(j, 0).Interior.color <> 16312794 Then Exit For
        If InStr(sheet.Range("UTTYPE").offset(j, 0), uttype) > 0 Then
            If Replace(sheet.Range("UTMIDSPAN" & i).offset(j, 0), "-", "") <> "" Then
                Dim item As Scripting.Dictionary: Set item = New Scripting.Dictionary
                item("size") = OnlyNumbers(sheet.Range("UTSIZE").offset(j, 0), True)
                item("type") = uttype
                If uttype = "PRI" Then
                    If InStr(sheet.Range("UTSIZE").offset(j, 0), "Ø") > 1 Then
                        item("phase") = Mid(sheet.Range("UTSIZE").offset(j, 0), InStr(sheet.Range("UTSIZE").offset(j, 0), "Ø") - 1, 1)
                        If CInt(item("phase")) > 3 Then item("phase") = "3"
                        item("size") = Left(item("size"), Len(item("size")) - 1)
                        neutralSpanCount = 0
                        Dim neutrals As Collection: Set neutrals = New Collection
                        neutralSize = ""
                        secondarySpanCount = 0
                        Dim secondaries As Collection: Set secondaries = New Collection
                        neutralShareHeight = False
                        For k = 0 To 100
                            If sheet.Range("UTTYPE").offset(k, 0).Interior.color <> 16312794 Then Exit For
                            If InStr(sheet.Range("UTTYPE").offset(k, 0), "NEUT") > 0 Then
                                neutrals.Add k
                                If Replace(sheet.Range("UTMIDSPAN" & i).offset(k, 0), "-", "") <> "" Then
                                    neutralSize = Utilities.OnlyNumbers(sheet.Range("UTSIZE").offset(k, 0), True)
                                    neutralSpanCount = neutralSpanCount + 1
                                    netrualHeight = Utilities.convertToInches(sheet.Range("UTHEIGHT").offset(k, 0))
                                    primaryHeight = Utilities.convertToInches(sheet.Range("UTHEIGHT").offset(j, 0))
                                    If (Abs(primaryHeight - netrualHeight) < 18) Then
                                        neutralShareHeight = True
                                    End If
                                End If
                            ElseIf InStr(sheet.Range("UTTYPE").offset(k, 0), "SEC") > 0 Or InStr(sheet.Range("UTTYPE").offset(k, 0), "OW") > 0 Then
                                secondaries.Add k
                                If Replace(sheet.Range("UTMIDSPAN" & i).offset(k, 0), "-", "") <> "" Then
                                    secondarySpanCount = secondarySpanCount + 1
                                End If
                            End If
                        Next k
                        If (neutralShareHeight) Then
                            item("config") = "N"
                            item("neutralSize") = neutralSize
                        ElseIf (secondarySpanCount > 0) Then
                            item("config") = "SN"
                        ElseIf (neutralSpanCount > 0) Then
                            item("config") = "NB"
                            item("neutralSize") = neutralSize
                        ElseIf (neutrals.count > 0 And secondaries.count = 0) Then
                            item("config") = "NB"
                            item("neutralSize") = neutralSize
                        ElseIf (secondaries.count > 0 And neutrals.count = 0) Then
                            item("config") = "SN"
                        ElseIf (secondaries.count = 0 And neutrals.count = 0) Then
                            item("config") = "N"
                            item("neutralSize") = item("size")
                        Else
                            closestSecAngleDif = 360
                            closestNeutAngleDif = 360
                            Set distanceAngle = getDistanceAngle(sheet, i)
                            angle = distanceAngle(2)
                            For Each secondary In secondaries
                                Set results = getClosestAngle(sheet, secondary, angle)
                                If closestSecAngleDif > results(1) Then closestSecAngleDif = results(1)
                            Next secondary
                            
                            For Each neutral In neutrals
                                Set results = getClosestAngle(sheet, neutral, angle)
                                If closestNeutAngleDif > results(1) Then
                                    closestNeutAngleDif = results(1)
                                    neutralSize = results(2)
                                End If
                            Next neutral

                            If closestSecAngleDif <= closestNeutAngleDif And closestSecAngleDif < 30 Then
                                item("config") = "SN"
                            ElseIf closestNeutAngleDif <= closestSecAngleDif And closestNeutAngleDif < 30 Then
                                item("config") = "NB"
                                item("neutralSize") = neutralSize
                            Else
                                item("config") = "N"
                                item("neutralSize") = item("size")
                            End If
                        End If
                    End If
                ElseIf uttype = "SEC" Then
                    size = sheet.Range("UTSIZE").offset(j, 0)
                    If InStr(size, "TX") > 0 Then
                        item("size") = item("size") & "TX"
                    ElseIf InStr(size, "DX") > 0 Then
                        item("size") = item("size") & "DX"
                    ElseIf InStr(size, "QX") > 0 Then
                        item("size") = item("size") & "QX"
                    ElseIf InStr(size, "AWAC") > 0 Then
                        item("size") = item("size") & "AWAC"
                    End If
                End If
                items.Add item
            End If
        End If
    Next j
    Set getItems = items
End Function

Public Function getDistanceAngle(sheet As Worksheet, ByVal i As String) As Collection
    distance = OnlyNumbers(sheet.Range("SPAN" & i).Value)
    angle = OnlyNumbers(Mid(sheet.Range("TOPOLE" & i).Value, InStr(sheet.Range("TOPOLE" & i).Value, "(")))
    
    Dim distanceAngle As Collection: Set distanceAngle = New Collection
    distanceAngle.Add distance
    distanceAngle.Add angle
    
    Set getDistanceAngle = distanceAngle
End Function

Public Function getClosestAngle(sheet As Worksheet, ByVal k As Integer, ByVal angle As Integer) As Collection
    Set closestAngle = New Collection
    
    smallestAngleDif = 360
    size = ""
    
    For i = 1 To 12
        For Each name In sheet.names
            If name.name = "'" & sheet.name & "'" & "!TOPOLE" & i Then
                If Trim(Replace(sheet.Range("TOPOLE" & i), "-", "")) <> "" Then
                    If Replace(sheet.Range("UTMIDSPAN" & i).offset(k, 0), "-", "") <> "" Then
                        Set results = getDistanceAngle(sheet, i)
                        angleDif = Abs(angle - results(2))
                        If smallestAngleDif > angleDif Then
                            smallestAngleDif = angleDif
                            size = OnlyNumbers(sheet.Range("UTSIZE").offset(k, 0), True)
                        End If
                    End If
                End If
            End If
        Next name
    Next i
    
    closestAngle.Add angleDif
    closestAngle.Add size
    Set getClosestAngle = closestAngle
End Function

Function LatLonToMI2253(latDeg As Double, lonDeg As Double) As Variant
    Const PI As Double = 3.14159265358979
 
    Const a As Double = 6378137#
    Const F As Double = 1# / 298.257222101
 
    Dim e As Double
    e = Sqr(2 * F - F * F)
 
    Dim lat0 As Double
    Dim lon0 As Double
    Dim sp1 As Double
    Dim sp2 As Double
 
    lat0 = 41.5 * PI / 180#
    lon0 = -84.3666666666667 * PI / 180#
    sp1 = 42.1 * PI / 180#
    sp2 = 43.6666666666667 * PI / 180#
 
    Dim lat As Double
    Dim lon As Double
 
    lat = latDeg * PI / 180#
    lon = lonDeg * PI / 180#
 
    Dim m1 As Double, m2 As Double
    Dim t1 As Double, t2 As Double
    Dim t0 As Double, t As Double
 
    m1 = Cos(sp1) / Sqr(1 - e ^ 2 * Sin(sp1) ^ 2)
    m2 = Cos(sp2) / Sqr(1 - e ^ 2 * Sin(sp2) ^ 2)
 
    t1 = Tan(PI / 4 - sp1 / 2) / _
         (((1 - e * Sin(sp1)) / (1 + e * Sin(sp1))) ^ (e / 2))
 
    t2 = Tan(PI / 4 - sp2 / 2) / _
         (((1 - e * Sin(sp2)) / (1 + e * Sin(sp2))) ^ (e / 2))
 
    t0 = Tan(PI / 4 - lat0 / 2) / _
         (((1 - e * Sin(lat0)) / (1 + e * Sin(lat0))) ^ (e / 2))
 
    t = Tan(PI / 4 - lat / 2) / _
        (((1 - e * Sin(lat)) / (1 + e * Sin(lat))) ^ (e / 2))
 
    Dim n As Double
    Dim lccF As Double
 
    n = (Log(m1) - Log(m2)) / (Log(t1) - Log(t2))
    lccF = m1 / (n * t1 ^ n)
 
    Dim rho As Double
    Dim rho0 As Double
 
    rho = a * lccF * t ^ n
    rho0 = a * lccF * t0 ^ n
 
    Dim xMeters As Double
    Dim yMeters As Double
 
    xMeters = rho * Sin(n * (lon - lon0))
    yMeters = rho0 - rho * Cos(n * (lon - lon0))
 
    Const FE As Double = 13123359.58
    Const FT_PER_M As Double = 3.28083333333333
 
    Dim xft As Double
    Dim yFt As Double
 
    xft = xMeters * FT_PER_M + FE
    yFt = yMeters * FT_PER_M
 
    LatLonToMI2253 = Array(xft, yFt)
End Function
