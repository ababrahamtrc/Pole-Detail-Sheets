Attribute VB_Name = "Outages"
Public ignoreIds As Scripting.Dictionary
Const secondaryOffsetMaxDistance As Integer = 35

Function AreTwoPointsEqual(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Boolean
    AreTwoPointsEqual = Abs(x1 - x2) < 0.1 And Abs(y1 - y2) < 0.1
End Function

Function FindTLMFromSec(jsonSecondary As Variant, secondaryJson As Object, transformerJson As Object, Optional first As Boolean) As String
    Dim x As Double, y As Double
    If first Then Set ignoreIds = New Scripting.Dictionary
        
    For Each Path In jsonSecondary("geometry")("paths")
        For Each Point In Path
            For Each jsonTransformerFeature In transformerJson("features")
                x = jsonTransformerFeature("geometry")("x")
                y = jsonTransformerFeature("geometry")("y")
                If AreTwoPointsEqual(CDbl(Point(1)), CDbl(Point(2)), x, y) Then
                    FindTLMFromSec = jsonTransformerFeature("attributes")("TLM")
                    Exit Function
                End If
            Next jsonTransformerFeature
        Next Point
    Next Path

    For Each jsonFeature In secondaryJson("features")
        If Not ignoreIds.exists(jsonSecondary("attributes")("OBJECTID")) Then ignoreIds.Add jsonSecondary("attributes")("OBJECTID"), Nothing
        If Not ignoreIds.exists(jsonFeature("attributes")("OBJECTID")) Then
            For Each Path In jsonFeature("geometry")("paths")
                For Each Point In Path
                    For Each path2 In jsonSecondary("geometry")("paths")
                        For Each point2 In path2
                            If AreTwoPointsEqual(CDbl(Point(1)), CDbl(Point(2)), CDbl(point2(1)), CDbl(point2(2))) Then
                                tlm = FindTLMFromSec(jsonFeature, secondaryJson, transformerJson)
                                If tlm <> "" Then
                                    FindTLMFromSec = tlm
                                    Exit Function
                                End If
                            End If
                        Next point2
                    Next path2
                Next Point
            Next Path
        End If
    Next jsonFeature
End Function

Function GetClosestPointOnLine(Ax As Double, Ay As Double, Bx As Double, By As Double, Px As Double, Py As Double) As Double()
    Dim ABx As Double, ABy As Double
    Dim APx As Double, APy As Double
    Dim dot_AP_AB As Double
    Dim dot_AB_AB As Double
    Dim t As Double
    Dim result(0 To 1) As Double

    ABx = Bx - Ax
    ABy = By - Ay
    APx = Px - Ax
    APy = Py - Ay
    
    dot_AP_AB = (APx * ABx) + (APy * ABy)
    dot_AB_AB = (ABx * ABx) + (ABy * ABy)
    
    If dot_AB_AB = 0 Then
        result(0) = Ax
        result(1) = Ay
        GetClosestPointOnLine = result
        Exit Function
    End If
    
    t = dot_AP_AB / dot_AB_AB
    
    If t < 0 Then
        t = 0
    ElseIf t > 1 Then
        t = 1
    End If
    
    result(0) = Ax + (t * ABx)
    result(1) = Ay + (t * ABy)
    
    GetClosestPointOnLine = result
End Function

Sub DownloadOutageLists()
    MsgBox "Extracting Information From GIS, this will take a moment. Leave Excel open and don't click anywhere. Wait for a form to appear."
    
    Call LogMessage.SendLogMessage("DownloadOutageLists")
    
    Dim pole As pole
    Dim project As project: Set project = New project
    Call project.extractFromSheets
    Dim poleCollections As Collection: Set poleCollections = findPoleGroups(project.poles)
    Dim poleCollection As Collection
    
    Dim outageCardExists As Boolean: outageCardExists = Dir(Environ("TEMP") & "\Outage Card.docx") <> ""
    If Not outageCardExists Then
        Dim url As String: url = "https://api.github.com/repos/ababrahamtrc/Pole-Detail-Sheets/contents/"
        Dim file As Object
        Dim http As Object: Set http = CreateObject("MSXML2.XMLHTTP")
        http.Open "GET", url & "Blanks/Outage Card.docx", False
        http.setRequestHeader "User-Agent", "ExcelVBA"
        http.send
        If http.status = 200 Then
            Set file = JsonConverter.ParseJson(http.responseText)
            If Not UpdatePoleDetailSheets.DownloadFile(file) Then
                MsgBox "Failed to download outage card template"
                Exit Sub
            End If
        Else
            MsgBox "Failed to download outage card template"
            Exit Sub
        End If
    End If
    
    Dim token As String: token = GetToken
    If Not testToken(token) Then
        MsgBox "Invalid token, get an up to date one from GIS."
        Exit Sub
    End If
    
    Dim locationTLMs As Scripting.Dictionary: Set locationTLMs = New Scripting.Dictionary
    Dim serviceJsons As Collection: Set serviceJsons = New Collection
    For Each poleCollection In poleCollections
        outageInPoleCollection = False
        For Each pole In poleCollection
            If pole.outage Then outageInPoleCollection = True: Exit For
        Next pole
        If outageInPoleCollection Then
            Dim secondaryJson As Object: Set secondaryJson = getElectricJson(poleCollection, 32, token)
            Dim secondaryUGJson As Object: Set secondaryUGJson = getElectricJson(poleCollection, 33, token)
            For Each jsonFeature In secondaryUGJson("features")
                secondaryJson("features").Add jsonFeature
            Next jsonFeature
            Dim transformerJson As Object: Set transformerJson = getElectricJson(poleCollection, 27, token)
            Dim transformerConnectorJson As Object: Set transformerConnectorJson = getElectricJson(poleCollection, 30, token)
            Dim tapPointsJson As Object: Set tapPointsJson = getElectricJson(poleCollection, 13, token)
            Dim polesJson As Object: Set polesJson = getElectricJson(poleCollection, 3, token)
            Dim serviceJson As Object: Set serviceJson = getElectricJson(poleCollection, 26, token)
            Call serviceJsons.Add(serviceJson)
            Dim circuitJson As Object: Set circuitJson = getElectricJson(poleCollection, 123, token)
            For Each pole In poleCollection
                If pole.outage Then
                    Set locationTLMs(pole.location) = New Collection
                    
                    results = getClosestPole(polesJson, pole.x, pole.y)
                    Dim closestPole As Object: Set closestPole = results(0)
                    Dim sourceX As Double: sourceX = 0
                    Dim sourceY As Double: sourceY = 0
                    Dim x As Double, y As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double
                    x = 0: y = 0: x1 = 0: y1 = 0: x2 = 0: y2 = 0
                    If Not closestPole Is Nothing Then
                        Dim closestPoleX As Double: closestPoleX = closestPole("geometry")("x")
                        Dim closestPoleY As Double: closestPoleY = closestPole("geometry")("y")
                        If pole.transformerSizes <> 0 Then
                            results = getClosestPole(transformerJson, closestPoleX, closestPoleY)
                            Set closestTransformer = results(0)
                            duplicate = False
                            For Each tlmValue In locationTLMs(pole.location)
                                If tlmValue = closestTransformer("attributes")("TLM") Then duplicate = True
                            Next tlmValue
                            If Not duplicate Then locationTLMs(pole.location).Add closestTransformer("attributes")("TLM")
                            sourceX = closestTransformer("geometry")("x")
                            sourceY = closestTransformer("geometry")("y")
                            For Each jsonFeature In transformerConnectorJson("features")
                                x1 = jsonFeature("geometry")("paths")(1)(1)(1)
                                y1 = jsonFeature("geometry")("paths")(1)(1)(2)
                                x2 = jsonFeature("geometry")("paths")(1)(2)(1)
                                y2 = jsonFeature("geometry")("paths")(1)(2)(2)
                                If AreTwoPointsEqual(x1, y1, CDbl(closestTransformer("geometry")("x")), CDbl(closestTransformer("geometry")("y"))) Then
                                    If Not AreTwoPointsEqual(x2, y2, closestPoleX, closestPoleY) Then
                                        sourceX = x2
                                        sourceY = y2
                                        Exit For
                                    End If
                                ElseIf AreTwoPointsEqual(x2, y2, CDbl(closestTransformer("geometry")("x")), CDbl(closestTransformer("geometry")("y"))) Then
                                    If Not AreTwoPointsEqual(x1, y1, closestPoleX, closestPoleY) Then
                                        sourceX = x1
                                        sourceY = y1
                                        Exit For
                                    End If
                                End If
                            Next jsonFeature
                        End If
                    
                        For Each jsonFeature In secondaryJson("features")
                            For Each Path In jsonFeature("geometry")("paths")
                                For Each Point In Path
                                    If AreTwoPointsEqual(CDbl(Point(1)), CDbl(Point(2)), closestPoleX, closestPoleY) Then
                                        sourceX = closestPole("geometry")("x")
                                        sourceY = closestPole("geometry")("y")
                                        tlm = FindTLMFromSec(jsonFeature, secondaryJson, transformerJson, True)
                                        If tlm <> "" Then
                                            duplicate = False
                                            For Each tlmValue In locationTLMs(pole.location)
                                                If tlmValue = tlm Then duplicate = True
                                            Next tlmValue
                                            If Not duplicate Then locationTLMs(pole.location).Add tlm
                                            Exit For
                                        End If
                                    End If
                                Next Point
                            Next Path
                        Next jsonFeature
                        
                        Dim attributes As Scripting.Dictionary: Set attributes = New Scripting.Dictionary
                        Set attributes("SUBTYPECD") = New Collection
                        attributes("SUBTYPECD").Add 8
                        attributes("SUBTYPECD").Add 7
                        results = getClosestPole(tapPointsJson, closestPoleX, closestPoleY, attributes, True)
                        If results(1) > 0 And results(1) < secondaryOffsetMaxDistance Then
                            Set closestSource = results(0)
                            sourceX = closestSource("geometry")("x")
                            sourceY = closestSource("geometry")("y")
                            For Each jsonFeature In secondaryJson("features")
                                For Each Path In jsonFeature("geometry")("paths")
                                    For Each Point In Path
                                        If AreTwoPointsEqual(CDbl(Point(1)), CDbl(Point(2)), sourceX, sourceY) Then
                                            tlm = FindTLMFromSec(jsonFeature, secondaryJson, transformerJson, True)
                                            If tlm <> "" Then
                                                duplicate = False
                                                For Each tlmValue In locationTLMs(pole.location)
                                                    If tlmValue = tlm Then duplicate = True
                                                Next tlmValue
                                                If Not duplicate Then locationTLMs(pole.location).Add tlm
                                                Exit For
                                            End If
                                        End If
                                    Next Point
                                Next Path
                            Next jsonFeature
                        End If
                        
                        If sourceX = 0 And sourceY = 0 Then
                            For Each jsonFeature In secondaryJson("features")
                                For Each Path In jsonFeature("geometry")("paths")
                                    For Each Point In Path
                                        x = Point(1)
                                        y = Point(2)
                                        distance = Sqr((closestPoleX - x) ^ 2 + (closestPoleY - y) ^ 2)
                                        If distance < secondaryOffsetMaxDistance Then
                                            tlm = FindTLMFromSec(jsonFeature, secondaryJson, transformerJson, True)
                                            If tlm <> "" Then
                                                duplicate = False
                                                For Each tlmValue In locationTLMs(pole.location)
                                                    If tlmValue = tlm Then duplicate = True
                                                Next tlmValue
                                                If Not duplicate Then locationTLMs(pole.location).Add tlm
                                                sourceX = x
                                                sourceY = y
                                            End If
                                        End If
                                    Next Point
                                Next Path
                            Next jsonFeature
                        End If
                        
                        If sourceX = 0 And sourceY = 0 Then
                            For Each jsonFeature In secondaryJson("features")
                                For Each Path In jsonFeature("geometry")("paths")
                                    For i = 1 To Path.count - 1
                                        Set point1 = Path(i)
                                        Set point2 = Path(i + 1)
                                        result = GetClosestPointOnLine(CDbl(point1(1)), CDbl(point1(2)), CDbl(point2(1)), CDbl(point2(2)), closestPoleX, closestPoleY)
                                        x = result(0)
                                        y = result(1)
                                        distance = Sqr((closestPoleX - x) ^ 2 + (closestPoleY - y) ^ 2)
                                        If distance < secondaryOffsetMaxDistance Then
                                            tlm = FindTLMFromSec(jsonFeature, secondaryJson, transformerJson, True)
                                            If tlm <> "" Then
                                                duplicate = False
                                                For Each tlmValue In locationTLMs(pole.location)
                                                    If tlmValue = tlm Then duplicate = True
                                                Next tlmValue
                                                If Not duplicate Then locationTLMs(pole.location).Add tlm
                                                sourceX = x
                                                sourceY = y
                                            End If
                                        End If
                                    Next i
                                Next Path
                            Next jsonFeature
                        End If
                        
                        If sourceX <> 0 And sourceY <> 0 Then
                            For Each jsonFeature In tapPointsJson("features")
                                If jsonFeature("attributes")("SUBTYPECD") = 9 Then
                                    x = jsonFeature("geometry")("x")
                                    y = jsonFeature("geometry")("y")
                                    distance = Sqr((sourceX - x) ^ 2 + (sourceY - y) ^ 2)
                                    If distance < secondaryOffsetMaxDistance Then
                                        For Each jsonSecondaryFeature In secondaryJson("features")
                                            For Each Path In jsonSecondaryFeature("geometry")("paths")
                                                For Each Point In Path
                                                    If Point(1) = x And Point(2) = y Then
                                                        tlm = FindTLMFromSec(jsonSecondaryFeature, secondaryJson, transformerJson, True)
                                                        If tlm <> "" Then
                                                            duplicate = False
                                                            For Each tlmValue In locationTLMs(pole.location)
                                                                If tlmValue = tlm Then duplicate = True
                                                            Next tlmValue
                                                            If Not duplicate Then locationTLMs(pole.location).Add tlm
                                                        End If
                                                    End If
                                                Next Point
                                            Next Path
                                        Next jsonSecondaryFeature
                                    End If
                                End If
                            Next jsonFeature
                        End If
                        
                        For Each jsonFeature In tapPointsJson("features")
                            If jsonFeature("attributes")("SUBTYPECD") = 9 Then
                                x = jsonFeature("geometry")("x")
                                y = jsonFeature("geometry")("y")
                                distance = Sqr((closestPoleX - x) ^ 2 + (closestPoleY - y) ^ 2)
                                If distance < secondaryOffsetMaxDistance Then
                                    For Each jsonSecondaryFeature In secondaryJson("features")
                                        For Each Path In jsonSecondaryFeature("geometry")("paths")
                                            For Each Point In Path
                                                If Point(1) = x And Point(2) = y Then
                                                    tlm = FindTLMFromSec(jsonSecondaryFeature, secondaryJson, transformerJson, True)
                                                    If tlm <> "" Then
                                                        duplicate = False
                                                        For Each tlmValue In locationTLMs(pole.location)
                                                            If tlmValue = tlm Then duplicate = True
                                                        Next tlmValue
                                                        If Not duplicate Then locationTLMs(pole.location).Add tlm
                                                    End If
                                                End If
                                            Next Point
                                        Next Path
                                    Next jsonSecondaryFeature
                                End If
                            End If
                        Next jsonFeature
                    End If
                End If
            Next pole
        Else
            Call serviceJsons.Add(Nothing)
        End If
    Next poleCollection
    
    Dim downloadPath As String: downloadPath = GetLocalPath(ThisWorkbook.Path) & "\"
    If InStr(downloadPath, "sharepoint") > 0 Then downloadPath = Environ("USERPROFILE") & "\Downloads\"
    
    Set objNetwork = CreateObject("WScript.Network")
    Dim DesignerName As String: DesignerName = GetObject("WinNT://" & objNetwork.UserDomain & "/" & objNetwork.UserName & ",user").FullName
    If ThisWorkbook.sheets("Control").Range("NAME") <> "" Then DesignerName = ThisWorkbook.sheets("Control").Range("NAME")
    
    Dim phases As Scripting.Dictionary: Set phases = New Scripting.Dictionary
    phases(1) = "Z"
    phases(2) = "Y"
    phases(3) = "YZ"
    phases(4) = "X"
    phases(5) = "XZ"
    phases(6) = "XY"
    phases(7) = "3P"
    
    Dim serviceTLM As String

    Call Outages_Form.Initialize(locationTLMs)
    Outages_Form.Show
    If Not Outages_Form.finished Then Exit Sub
    
    Call LoadingBar_Form2.InitProgress(locationTLMs.count, True, 1)
    
    Dim usedPoles As New Scripting.Dictionary
    Dim tlms1 As Collection, tlms2 As Collection
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim poleCollectionCount As Integer: poleCollectionCount = 0
    For Each poleCollection In poleCollections
        poleCollectionCount = poleCollectionCount + 1
        Set serviceJson = serviceJsons(poleCollectionCount)
        For Each pole In poleCollection
            If LoadingBar_Form2.gTotal = 0 Then Exit Sub
            If locationTLMs.exists(pole.location) And Not usedPoles.exists(pole.location) Then
                Dim locationsUsed As Collection: Set locationsUsed = New Collection
                usedPoles.Add pole.location, Nothing
                locationsUsed.Add pole.location
                Set tlms1 = locationTLMs(pole.location)
                For Each location In locationTLMs
                    If location <> pole.location Then
                        Set tlms2 = locationTLMs(location)
                        If AreCollectionsEqualUnordered(tlms1, tlms2) Then
                            usedPoles.Add location, Nothing
                            locationsUsed.Add location
                        End If
                    End If
                Next location
                
                Dim outageList As Collection: Set outageList = New Collection
                Dim resCount As Integer: resCount = 0
                Dim comCount As Integer: comCount = 0
                Dim feederId As String: feederId = ""
                Dim feederId2 As String: feederId2 = ""
                Dim county As String: counter = ""
                Dim city As String: city = ""
                Dim substation As String: substation = ""
                Dim circuit As String: circuit = ""
                Dim streets As New Scripting.Dictionary
                Dim priorityCount As Integer: priorityCount = 0
                Dim multiPhaseCount As Integer: multiPhaseCount = 0
                Dim otherCount As Integer: otherCount = 0
                Dim criticalCount As Integer: criticalCount = 0
                Dim connectedCount As Integer: connectedCount = 0
                Dim disconnectedCount As Integer: disconnectedCount = 0
    
                If Not serviceJson Is Nothing Then
                    Dim accountNumbers As Scripting.Dictionary: Set accountNumbers = New Scripting.Dictionary
                    For Each tlm In locationTLMs(pole.location)
                        If LoadingBar_Form2.gTotal = 0 Then Exit Sub
                        For Each jsonFeature In serviceJson("features")
                            If LoadingBar_Form2.gTotal = 0 Then Exit Sub
                            serviceTLM = jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServicePoint.TLM")
                            If serviceTLM = tlm Then
                                Dim row As Collection: Set row = New Collection
                                
                                If Not IsNull(jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServicePoint.FEEDERID")) Then feederId = jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServicePoint.FEEDERID")
                                If Not IsNull(jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServicePoint.FEEDERID2")) Then feederId2 = jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServicePoint.FEEDERID2")
                                If Not IsNull(jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.COUNTY")) Then county = jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.COUNTY")
                                If Not IsNull(jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.CITY")) Then city = jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.CITY")
                                
                                If substation = "" And feederId <> "" Then
                                     For Each jsonFeature2 In circuitJson("features")
                                        If jsonFeature2("attributes")("FEEDERID") = feederId Then
                                            If Not IsNull(jsonFeature2("attributes")("SUBSTATION")) Then substation = jsonFeature2("attributes")("SUBSTATION")
                                            If Not IsNull(jsonFeature2("attributes")("CIRCUIT")) Then circuit = jsonFeature2("attributes")("CIRCUIT")
                                        End If
                                     Next jsonFeature2
                                End If
                                
                                Dim accountNumber As String: If Not IsNull(jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.ACCOUNTNUMBER")) Then accountNumber = jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.ACCOUNTNUMBER") Else accountNumber = ""
                                If Not accountNumbers.exists(accountNumber) Then
                                    accountNumbers.Add accountNumber, Nothing
                                    
                                    Dim accountType As String: If Not IsNull(jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.ACCOUNTTYPE")) Then accountType = jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.ACCOUNTTYPE") Else accountType = ""
                                    
                                    If accountType = "RES" Then
                                        resCount = resCount + 1
                                    ElseIf accountType = "COM" Then
                                        comCount = comCount + 1
                                    End If
                                    
                                    Dim lastName As String: If Not IsNull(jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.LASTNAME")) Then lastName = jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.LASTNAME") Else lastName = ""
                                    Dim firstName As String: If Not IsNull(jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.FIRSTNAME")) Then firstName = jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.FIRSTNAME") Else firstName = ""
                                    Dim street As String: If Not IsNull(jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.STREET")) Then street = jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.STREET") Else street = ""
                                    Dim streetName As String
                                    If InStr(street, " ") > 0 Then
                                        streetName = Right(street, Len(street) - InStr(street, " "))
                                        If Not streets.exists(streetName) Then streets.Add streetName, Nothing
                                    End If
                                    Dim postalCode As String: If Not IsNull(jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.POSTALCODE")) Then postalCode = jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.POSTALCODE") Else postalCode = ""
                                    
                                    Dim telephones As Scripting.Dictionary: Set telephones = New Scripting.Dictionary
                                    For i = 1 To 4
                                        telephone = jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.TELEPHONE" & i)
                                        If Not IsNull(telephone) Then
                                            If telephone <> "" Then
                                                If Not telephones.exists(telephone) Then telephones.Add telephone, Nothing
                                            End If
                                        End If
                                    Next i
                                    
                                    Dim telephone1 As String, telephone2 As String, telephone3 As String, telephone4 As String
                                    If telephones.count > 0 Then telephone1 = telephones.keys()(0) Else telephone1 = ""
                                    If telephones.count > 1 Then telephone2 = telephones.keys()(1) Else telephone2 = ""
                                    If telephones.count > 2 Then telephone3 = telephones.keys()(2) Else telephone3 = ""
                                    If telephones.count > 3 Then telephone4 = telephones.keys()(3) Else telephone4 = ""
                                    
                                    Dim meter As String: If Not IsNull(jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServicePoint.METERNUMBER")) Then meter = jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServicePoint.METERNUMBER")
                                    If InStr(meter, ".") > 0 Then
                                        parts = Split(meter, ".")
                                        meter = parts(UBound(parts))
                                    End If
                                    Dim phase As String: If Not IsNull(jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServicePoint.PHASEDESIGNATION")) Then phase = phases(jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServicePoint.PHASEDESIGNATION")) Else phase = ""
                                    Dim criticalCustomer As String: If Not IsNull(jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.CRITICALRESTORATION")) Then criticalCustomer = jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServiceAddress.CRITICALRESTORATION") Else criticalCustomer = ""
                                    Dim priorityRestoration As String: If Not IsNull(jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServicePoint.PRIORITYRESTORATION")) Then priorityRestoration = jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServicePoint.PRIORITYRESTORATION") Else priorityRestoration = ""
                                    Dim disconnectReason As String: If Not IsNull(jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServicePoint.DISCONNECTIONREASON")) Then disconnectReason = jsonFeature("attributes")("geoAIM_ElecDist.ELECDIST.ServicePoint.DISCONNECTIONREASON") Else disconnectReason = ""
                                    Dim connectStatus As String
                                    
                                    If disconnectReason <> "" Then
                                        connectStatus = "Disconnected"
                                    Else
                                        connectStatus = "Connected"
                                    End If
                                    
                                    Dim group As String: group = ""
                                    If criticalCustomer <> "" Then group = "Critical"
                                    If priorityRestoration <> "" Then
                                        If group <> "" Then group = group & "|"
                                        group = group & "Priority"
                                    End If
                                    If phase = "YZ" Or phase = "XZ" Or phase = "XY" Or phase = "3P" Then
                                        If group <> "" Then group = group & "|"
                                        group = group & "Multiphase"
                                    End If
                                    If group = "" Then group = "Other"
                                    
                                    If accountNumber <> "" Then
                                        If InStr(group, "Priority") > 0 Then priorityCount = priorityCount + 1
                                        If InStr(group, "Multiphase") > 0 Then multiPhaseCount = multiPhaseCount + 1
                                        If InStr(group, "Other") > 0 Then otherCount = otherCount + 1
                                        If InStr(group, "Critical") > 0 Then criticalCount = criticalCount + 1
                                        
                                        If connectStatus = "Connected" Then connectedCount = connectedCount + 1
                                        If connectStatus = "Disconnected" Then disconnectedCount = disconnectedCount + 1
                                        
                                        row.Add accountNumber
                                        row.Add accountType
                                        row.Add lastName
                                        row.Add firstName
                                        row.Add street
                                        row.Add city
                                        row.Add "MI"
                                        row.Add postalCode
                                        row.Add telephone1
                                        row.Add telephone2
                                        row.Add telephone3
                                        row.Add telephone4
                                        row.Add ""
                                        row.Add criticalCustomer
                                        row.Add connectStatus
                                        row.Add meter
                                        row.Add ""
                                        row.Add serviceTLM
                                        row.Add phase
                                        row.Add group
                                         
                                        outageList.Add row
                                    End If
                                End If
                            End If
                        Next jsonFeature
                    Next tlm
                End If
                
                If LoadingBar_Form2.gTotal = 0 Then Exit Sub
                
                Dim NewBook As Workbook
                Set NewBook = Workbooks.Add(xlWBATWorksheet)
                Dim sheet As Worksheet: Set sheet = NewBook.sheets(1)
                
                sheet.Cells(1, 1).Value = "Name"
                sheet.Cells(1, 2).Value = "ReportDate"
                sheet.Cells(1, 3).Value = "DeviceType"
                sheet.Cells(1, 4).Value = "DeviceFID"
                sheet.Cells(1, 5).Value = "DeviceOID"
                sheet.Cells(1, 6).Value = "EID"
                sheet.Cells(1, 7).Value = "FeederID"
                sheet.Cells(2, 7).Value = feederId
                sheet.Cells(1, 8).Value = "FeederID2"
                sheet.Cells(2, 8).Value = feederId2
                sheet.Cells(1, 9).Value = "Substation"
                sheet.Cells(2, 9).Value = substation
                sheet.Cells(1, 10).Value = "Message"
                sheet.Cells(1, 11).Value = "Start Date"
                sheet.Cells(1, 12).Value = "Finish Date"
                sheet.Cells(1, 13).Value = "Alternate Start Date"
                sheet.Cells(1, 14).Value = "Alternate Finish Date"
                sheet.Cells(1, 15).Value = "ConnectedCustomers"
                sheet.Cells(1, 16).Value = "Priority"
                sheet.Cells(2, 16).Value = priorityCount
                sheet.Cells(1, 17).Value = "Multiphase"
                sheet.Cells(2, 17).Value = multiPhaseCount
                sheet.Cells(1, 18).Value = "Other"
                sheet.Cells(2, 18).Value = otherCount
                sheet.Cells(1, 19).Value = "Critical"
                sheet.Cells(2, 19).Value = criticalCount
                sheet.Cells(1, 20).Value = "Disconnected"
                sheet.Cells(2, 20).Value = disconnectedCount
                sheet.Cells(1, 21).Value = "Connected"
                sheet.Cells(2, 21).Value = connectedCount
                sheet.Cells(1, 22).Value = "DisconnectedCustomers"
                
                sheet.Cells(2, 1).Value = "Customer List"
                sheet.Cells(2, 2).Value = Format(Now, "m/d/yyyy hh:nn:ss")
                sheet.Cells(2, 3).Value = "Secondary Transformers"
                
                sheet.Cells(5, 1).Value = "Account"
                sheet.Cells(5, 2).Value = "AccountType"
                sheet.Cells(5, 3).Value = "LastName"
                sheet.Cells(5, 4).Value = "FirstName"
                sheet.Cells(5, 5).Value = "Street"
                sheet.Cells(5, 6).Value = "City"
                sheet.Cells(5, 7).Value = "State"
                sheet.Cells(5, 8).Value = "PostalCode"
                sheet.Cells(5, 9).Value = "Telephone"
                sheet.Cells(5, 10).Value = "Telephone2"
                sheet.Cells(5, 11).Value = "Telephone3"
                sheet.Cells(5, 12).Value = "Telephone4"
                sheet.Cells(5, 13).Value = "CallbackNumber"
                sheet.Cells(5, 14).Value = "CriticalCustomer"
                sheet.Cells(5, 15).Value = "ConnectStatus"
                sheet.Cells(5, 16).Value = "Meter"
                sheet.Cells(5, 17).Value = "DeviceOID"
                sheet.Cells(5, 18).Value = "TLM"
                sheet.Cells(5, 19).Value = "Phase"
                sheet.Cells(5, 20).Value = "Groups"
                
                For i = 1 To outageList.count
                    For j = 1 To outageList(i).count
                        sheet.Cells(i + 5, j).Value = outageList(i)(j)
                    Next j
                Next i
                
                For Each cell In sheet.UsedRange
                    cell.WrapText = False
                    If (cell.row < 3 And cell.Column < 23) Or (cell.Column < 21 And cell.row > 4) Then
                        With cell.Borders
                            .LineStyle = xlContinuous
                            .Weight = xlThin
                        End With
                    End If
                    If cell.row = 1 Or cell.row = 5 Then
                        cell.Font.Bold = True
                    End If
                Next cell
                sheet.Cells.EntireColumn.AutoFit
                sheet.Cells.EntireRow.AutoFit
                
                locationNumbers = compactListString(locationsUsed)
                
                FileName = project.Notification & " - Outage List Loc " & locationNumbers & ".xlsx"
                NewBook.SaveAs FileName:=downloadPath & FileName, FileFormat:=xlOpenXMLWorkbook
                NewBook.Close savechanges:=False
                
                If LoadingBar_Form2.gTotal = 0 Then Exit Sub
                
                If comCount > 3 Or resCount > 9 Then
                    Dim cardStreets As New Collection
                    For Each streetKey In streets
                        Dim addStreet As Boolean: addStreet = True
                        For Each streetKey2 In streets
                            If streetKey <> streetKey2 Then
                                If Len(streetKey2) > 4 And InStr(streetKey, streetKey2) Then addStreet = False: Exit For
                            End If
                        Next streetKey2
                        duplicate = False
                        For Each cardstreet In cardStreets
                            If cardstreet = streetKey Then duplicate = True: Exit For
                        Next cardstreet
                        If addStreet And Not duplicate Then cardStreets.Add streetKey
                    Next streetKey
                    
                    Set wdApp = CreateObject("Word.Application")
                    wdApp.visible = False
                    Set wdDoc = wdApp.Documents.Open(Environ("TEMP") & "\Outage Card.docx")
        
                    Dim center As String: center = ""
                    If InStr(county, "KENT") > 0 Then center = "Grand Rapids Work Management Center"
                    If InStr(county, "JACKSON") > 0 Then center = "Jackson Work Management Center"
                    If InStr(county, "SAGINAW") > 0 Then center = "Saginaw Work Management Center"
                    
                    wdDoc.ContentControls(1).Range.text = DesignerName
                    wdDoc.ContentControls(6).Range.text = center
                    wdDoc.ContentControls(12).Range.text = JoinCollection(locationTLMs(pole.location), ", ")
                    If feederId <> "" Then
                        wdDoc.ContentControls(13).Range.text = substation & "-" & Left(feederId, 4)
                        wdDoc.ContentControls(14).Range.text = circuit & "-" & Right(feederId, 2)
                    End If
                    
                    Dim secRegex As Object
                    Set secRegex = CreateObject("VBScript.RegExp")
                    
                    secRegex.Pattern = "\s*(\d*)'[ OF]*\s*(.*)SEC\s*\/\s*(.*)SEC\s*(.*)"
                    secRegex.Global = True
                    secRegex.IgnoreCase = True
                    
                    Dim crewNotes As String: crewNotes = pole.alt1
                    Call Utilities.applyStandardAbbreviations(crewNotes)
                    lines = Split(crewNotes, vbLf)
                    Dim reconductored As Boolean
                    For Each line In lines
                        If InStr(line, "SEC") > 0 Then
                            If secRegex.test(line) Then
                                reconductored = True
                                Exit For
                            End If
                        End If
                    Next line
                    Dim reason As String
                    If pole.ReplacePole Then
                        reason = "Pole Replacement"
                    ElseIf reconductored Then
                        reason = "Secondary Reconductoring"
                    ElseIf InStr(crewNotes, "TRIM SEC") > 0 Then
                        reason = "Trimming Secondary"
                    ElseIf InStr(crewNotes, "RISER &") > 0 Then
                        reason = "Replacing Riser"
                    ElseIf InStr(crewNotes, "RAISE") > 0 Then
                        reason = "Raising Secondary"
                    End If
                    
                    wdDoc.ContentControls(15).Range.text = reason
                    wdDoc.ContentControls(16).Range.text = comCount + resCount
                    wdDoc.ContentControls(17).Range.text = project.Notification
                    wdDoc.ContentControls(18).Range.text = JoinCollection(cardStreets, ", ")
                    wdDoc.ContentControls(19).Range.text = WorksheetFunction.Proper(county & "/" & city)
                    
                    FileName = project.Notification & " - Outage Card Loc " & locationNumbers & ".docx"
                    wdDoc.saveAs2 downloadPath & FileName
                    wdDoc.Close savechanges:=False
                    wdApp.Quit
                End If
            End If
            If locationTLMs.exists(pole.location) Then Call LoadingBar_Form2.UpdateProgress("Downloading Outage Lists/Cards", "Locations Downloaded", True)
        Next pole
    Next poleCollection
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    Call LoadingBar_Form2.FinishProgress
    MsgBox "Finished generate List/Cards"
End Sub
