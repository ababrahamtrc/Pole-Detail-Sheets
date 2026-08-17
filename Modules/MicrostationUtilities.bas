Attribute VB_Name = "MicrostationUtilities"
Function testToken(token As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim url As String: url = "https://gis.consumersenergy.com/mapping/rest/services/Landbase/Landbase_Grids_Boundaries_PUB/MapServer/1/query?where=1=1&outFields=*&returnGeometry=true&geometryType=esriGeometryEnvelope&geometry=&spatialRel=esriSpatialRelIntersects&inSR=2253&outSR=2253+&f=json&token=" & token
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", url, False
    http.setRequestHeader "Authorization", "Bearer " & token
    http.setRequestHeader "Content-Type", "application/json"
    http.SetTimeouts 15000, 15000, 15000, 15000
    http.send
    If InStr(http.responseText, "Invalid Token") > 0 Then
        testToken = False
    Else
        testToken = True
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "HTTP Request Failed or Timed Out!" & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, vbCritical
End Function

Function getElectricJson(poles As Collection, layer As Integer, token As String, Optional query As String) As Object
    On Error GoTo ErrorHandler
    
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
    
    Dim radius As Double: radius = 750
    
    results = LatLonToMI2253(lowestLatitude, lowestLongitude)
    x1 = results(0) - radius
    y1 = results(1) - radius
    
    results = LatLonToMI2253(highestLatitude, highestLongitude)
    x2 = results(0) + radius
    y2 = results(1) + radius
    
    Dim bbox1 As String: bbox1 = x1 & "," & y1
    Dim bbox2 As String: bbox2 = x2 & "," & y2
    Dim bbox As String: bbox = bbox1 & "," & bbox2
    
    Dim url As String: url = "https://gis.consumersenergy.com/mapping/rest/services/Electric/Electric_PUB/MapServer/" & layer & "/query?where=" & IIf(query <> "", query, "1=1") & "&outFields=*&returnGeometry=true&geometryType=esriGeometryEnvelope&geometry=" & bbox & "&spatialRel=esriSpatialRelIntersects&inSR=2253&outSR=2253+&f=json&token=" & token
    Debug.Print url

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", url, False
    http.setRequestHeader "Authorization", "Bearer " & token
    http.setRequestHeader "Content-Type", "application/json"
    http.SetTimeouts 15000, 15000, 15000, 15000
    http.send
    
    If http.status = 200 Then
        Set getElectricJson = JsonConverter.ParseJson(http.responseText)
    Else
        Set getElectricJson = Nothing
        Debug.Print "Error: " & http.status & " - " & http.statusText
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "HTTP Request Failed or Timed Out!" & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, vbCritical
    
End Function

Function getOtherPole(x As Double, y As Double, radius As Integer, token) As Object
    On Error GoTo ErrorHandler
    
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
    http.SetTimeouts 15000, 15000, 15000, 15000
    http.send
    
    If http.status = 200 Then
        Set getOtherPole = JsonConverter.ParseJson(http.responseText)
    Else
        Set getOtherPole = Nothing
        Debug.Print "Error: " & http.status & " - " & http.statusText
    End If
    
    Set http = Nothing
    Exit Function
    
ErrorHandler:
    MsgBox "HTTP Request Failed or Timed Out!" & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, vbCritical
End Function

Function getROWJSON(poles As Collection, layer As Integer, token As String) As Object
    On Error GoTo ErrorHandler
    
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
    Debug.Print url
    
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", url, False
    http.setRequestHeader "Authorization", "Bearer " & token
    http.setRequestHeader "Content-Type", "application/json"
    http.SetTimeouts 15000, 15000, 15000, 15000
    http.send
    
    If http.status = 200 Then
        Set getROWJSON = JsonConverter.ParseJson(http.responseText)
    Else
        Set getROWJSON = Nothing
        Debug.Print "Error: " & http.status & " - " & http.statusText
    End If
    
    Set http = Nothing
    Exit Function
    
ErrorHandler:
    MsgBox "HTTP Request Failed or Timed Out!" & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, vbCritical
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
                    If Replace(sheet.Range("UTMIDSPAN" & i).OFFSET(k, 0), "-", "") <> "" Then
                        Set results = getDistanceAngle(sheet, i)
                        angleDif = Abs(angle - results(2))
                        If smallestAngleDif > angleDif Then
                            smallestAngleDif = angleDif
                            size = OnlyNumbers(sheet.Range("UTSIZE").OFFSET(k, 0), True)
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

Public Function Atn2(dy As Double, dx As Double) As Double
    Dim PI As Double
    PI = 4 * Atn(1)
    
    If dx > 0 Then
        Atn2 = Atn(dy / dx)
    ElseIf dx < 0 Then
        If dy >= 0 Then
            Atn2 = Atn(dy / dx) + PI
        Else
            Atn2 = Atn(dy / dx) - PI
        End If
    Else
        If dy > 0 Then
            Atn2 = PI / 2
        ElseIf dy < 0 Then
            Atn2 = -PI / 2
        Else
            Atn2 = 0
        End If
    End If
End Function

Function LatLonToMI2253(latDeg As Double, lonDeg As Double) As Variant
    Const PI As Double = 3.14159265358979
 
    Const A As Double = 6378137#
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
 
    rho = A * lccF * t ^ n
    rho0 = A * lccF * t0 ^ n
 
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
