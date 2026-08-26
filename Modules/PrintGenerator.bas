Attribute VB_Name = "PrintGenerator"
Const PI As Double = 3.14159265358979
Dim existingLineIDs As Object

Public Sub GeneratePrint()
     Set testText = CreateTextElement1(Nothing, "Test", Point3dFromXYZ(0, 0, 0), Matrix3dIdentity)
    
    On Error Resume Next
    ActiveModelReference.AddElement testText
    If Err.Number <> 0 Then
        MsgBox "File is Read Only"
        Err.Clear
        Exit Sub
    Else
        IsFileReadOnlyByOperation = False
        ActiveModelReference.RemoveElement testText
    End If
    On Error GoTo 0
    
    Unload PrintOptions
    Call PrintOptions.Initialize
    PrintOptions.Show vbModeless
End Sub

Public Sub DrawPrint()
    Dim oDataBlock As New DataBlock
    Dim json As Object
    Set json = ReadJSON("print")
    
    Dim oSettings As Settings: Set oSettings = ActiveSettings
    
    Dim textHeight As Double
    textHeight = ActiveSettings.TextStyle.height
    ActiveSettings.TextStyle = ActiveDesignFile.TextStyles("Engineering")
    
    If json Is Nothing Then
        MsgBox "JSON file doesn't exist, generate one from pole detail sheets."
        Exit Sub
    End If
    
    Dim oCriteria As New ElementScanCriteria
    oCriteria.ExcludeNonGraphical
    oCriteria.ExcludeAllTypes
    oCriteria.IncludeType msdElementTypeLine
    
    Dim oEnumerator As ElementEnumerator
    Set oEnumerator = ActiveModelReference.Scan(oCriteria)
    
    Set existingLineIDs = CreateObject("Scripting.Dictionary")
    Do While oEnumerator.MoveNext
        Dim strID As String
        strID = CStr(oEnumerator.Current.ID64)
        If Not existingLineIDs.exists(strID) Then
            existingLineIDs.Add strID, True
        End If
    Loop
    
    If PrintOptions.DrawCenterLines Then
        Dim centerlineLine As LineElement
        For Each centerline In json("centerlines")
            i = 0
            Dim centerlinePoints() As Point3d
            For Each centerlinePoint In centerline
                ReDim Preserve centerlinePoints(i)
                centerlinePoints(i) = Point3dFromXYZ(centerlinePoint(1), centerlinePoint(2), 0)
                i = i + 1
            Next centerlinePoint
            Set centerlineLine = CreateLineElement1(Nothing, centerlinePoints)
            centerlineLine.color = 6
            centerlineLine.LineStyle = ActiveDesignFile.LineStyles.Find("4")
            oDataBlock.CopyString "CENTERLINE", True
            Call centerlineLine.AddUserAttributeData(11111, oDataBlock)
            ActiveModelReference.AddElement centerlineLine
            If PrintOptions.ShowDrawing Then DoEvents
        Next centerline
    End If
    
    If PrintOptions.DrawROW Then
        Dim roadLine As LineElement
        For Each road In json("roads")
            i = 0
            Dim roadPoints() As Point3d
            For Each roadPoint In road
                ReDim Preserve roadPoints(i)
                roadPoints(i) = Point3dFromXYZ(roadPoint(1), roadPoint(2), 0)
                i = i + 1
            Next roadPoint
            Set roadLine = CreateLineElement1(Nothing, roadPoints)
            roadLine.color = 0
            ActiveModelReference.AddElement roadLine
            If PrintOptions.ShowDrawing Then DoEvents
        Next road
    End If
    
    If PrintOptions.DrawEquipment Then
        If PrintOptions.DrawTransformers Then
            Dim jsonTransformer As Object
            For Each jsonTransformer In json("transformers")
                If Not jsonTransformer("adjacent") Or (PrintOptions.DrawAdjacentPoleEquipment And jsonTransformer("adjacent")) Then
                    Call placeTransformer(jsonTransformer)
                End If
            Next jsonTransformer
        End If
        
        If PrintOptions.DrawCapacitors Then
            Dim jsonCapacitor As Object
            For Each jsonCapacitor In json("capacitors")
                If Not jsonCapacitor("adjacent") Or (PrintOptions.DrawAdjacentPoleEquipment And jsonCapacitor("adjacent")) Then
                    Call placeCapacitor(jsonCapacitor)
                End If
            Next jsonCapacitor
        End If
        
        If PrintOptions.DrawRegulators Then
            Dim jsonRegulator As Object
            For Each jsonRegulator In json("regulators")
                If Not jsonRegulator("adjacent") Or (PrintOptions.DrawAdjacentPoleEquipment And jsonRegulator("adjacent")) Then
                    Call placeRegulator(jsonRegulator)
                End If
            Next jsonRegulator
        End If
        
        If PrintOptions.DrawIsolators Then
            Dim jsonIsolator As Object
            For Each jsonIsolator In json("isolators")
                If Not jsonIsolator("adjacent") Or (PrintOptions.DrawAdjacentPoleEquipment And jsonIsolator("adjacent")) Then
                    Call placeIsolator(jsonIsolator)
                End If
            Next jsonIsolator
        End If
        
        If PrintOptions.DrawStreetlights Then
            Dim jsonStreetlight As Object
            For Each jsonStreetlight In json("streetlights")
                If Not jsonStreetlight("adjacent") Or (PrintOptions.DrawAdjacentPoleEquipment And jsonStreetlight("adjacent")) Then
                    Call placeStreetlight(jsonStreetlight)
                End If
            Next jsonStreetlight
        End If
        
        If PrintOptions.DrawFuses Then
            Dim jsonFuse As Object
            For Each jsonFuse In json("fuses")
                If Not jsonFuse("adjacent") Or (PrintOptions.DrawAdjacentPoleEquipment And jsonFuse("adjacent")) Then
                    Call placeLCPObject(jsonFuse, "Fuse")
                End If
            Next jsonFuse
        End If
        
        If PrintOptions.DrawReclosers Then
            Dim jsonRecloser As Object
            For Each jsonRecloser In json("reclosers")
                If Not jsonRecloser("adjacent") Or (PrintOptions.DrawAdjacentPoleEquipment And jsonRecloser("adjacent")) Then
                    Call placeLCPObject(jsonRecloser, "Recloser")
                End If
            Next jsonRecloser
        End If
        
        If PrintOptions.DrawSectionalizers Then
            Dim jsonSectionalizer As Object
            For Each jsonSectionalizer In json("sectionalizers")
                If Not jsonSectionalizer("adjacent") Or (PrintOptions.DrawAdjacentPoleEquipment And jsonSectionalizer("adjacent")) Then
                    Call placeLCPObject(jsonSectionalizer, "Sectionalizer")
                End If
            Next jsonSectionalizer
        End If
        
        If PrintOptions.DrawSwitches Then
            Dim jsonSwitch As Object
            For Each jsonSwitch In json("switches")
                If Not jsonSwitch("adjacent") Or (PrintOptions.DrawAdjacentPoleEquipment And jsonSwitch("adjacent")) Then
                    Call placeLCPObject(jsonSwitch, "Switch")
                End If
            Next jsonSwitch
        End If
        
        If PrintOptions.DrawSensors Then
            Dim jsonSensor As Object
            For Each jsonSensor In json("sensors")
                If Not jsonSensor("adjacent") Or (PrintOptions.DrawAdjacentPoleEquipment And jsonSensor("adjacent")) Then
                    Call placeLCPObject(jsonSensor, "Sensor")
                End If
            Next jsonSensor
        End If
        
        Dim jsonRiser As Object
        For Each jsonRiser In json("risers")
            If Not jsonRiser("adjacent") Or (PrintOptions.DrawAdjacentPoleEquipment And jsonRiser("adjacent")) Then
                If jsonRiser("type") = "Secondary" And PrintOptions.DrawSecondaryRisers Then
                    Call placeLCPObject(jsonRiser, "Riser")
                ElseIf jsonRiser("type") = "Primary" And PrintOptions.DrawPrimaryRisers Then
                    Call placeLCPObject(jsonRiser, "Riser")
                End If
            End If
        Next jsonRiser
    End If
    
    If PrintOptions.DrawSpanGuys Then
        Dim jsonSpanguy As Object
        For Each jsonSpanguy In json("spanguys")
            Call placeSpanguy(jsonSpanguy)
        Next jsonSpanguy
    End If
    
    If PrintOptions.DrawConductors Then
        If PrintOptions.DrawSecondary Then
            Dim jsonSec As Object
            For Each jsonSec In json("secWires")
                Call placeSecondary(jsonSec)
            Next jsonSec
        End If
        
        If PrintOptions.DrawOpenWire Then
            For Each jsonSec In json("openWires")
                Call placeSecondary(jsonSec)
            Next jsonSec
        End If
        If PrintOptions.DrawDeadends And (PrintOptions.DrawSecondary Or PrintOptions.DrawOpenWire) Then Call ShortenLines("CE-EX-ELEC-OH-SEC-COND")
        
        If PrintOptions.DrawPrimary Then
            Dim jsonPri As Object
            For Each jsonPri In json("priWires")
                Call placePrimary(jsonPri)
            Next jsonPri
            If PrintOptions.DrawDeadends Then Call ShortenLines("CE-EX-ELEC-OH-PRI-COND")
        End If
    End If
    
    If PrintOptions.DrawServices Then
        Dim jsonService As Object
        For Each jsonService In json("services")
            If Not jsonService("adjacent") Or (PrintOptions.DrawAdjacentServices And jsonService("adjacent")) Then
                If Not jsonService("ug") Or (PrintOptions.DrawUGServices And jsonService("ug")) Then
                    Call drawService(jsonService)
                End If
            End If
        Next jsonService
    End If
    
    Dim largestX As Double
    Dim largestY As Double
    Dim jsonPole As Object
    For Each jsonPole In json("poles")
        If jsonPole("x") > largestX Then largestX = jsonPole("x")
        If jsonPole("y") > largestY Then largestY = jsonPole("y")
        Call placePole(jsonPole)
    Next jsonPole
    
    Dim startPoint As Point3d
    
    Dim OFFSET As Point3d
    OFFSET = Point3dFromXYZ(-25, -10, 0)

    Dim groupStartingPoints As Object: Set groupStartingPoints = CreateObject("Scripting.Dictionary")
    For Each groupKey In json("groups")
        groupStartingPoints(groupKey) = Point3dFromXYZ(json("groups")(groupKey)("maxX") + 500, json("groups")(groupKey)("maxY"), 0)
    Next groupKey
    
    If PrintOptions.DrawCrewNotes Then
        Dim group As String
        Dim strLine As String
        For Each jsonPole In json("poles")
            If jsonPole("location") <> "" Then
                If jsonPole("crewNotes") <> "" Then
                    crewNotes = jsonPole("crewNotes")
                    strLines = Split(jsonPole("crewNotes"), vbCrLf)
                    group = CStr(jsonPole("group"))
                    startPoint = groupStartingPoints(group)
                    Dim txt As TextNodeElement: Set txt = CreateTextNodeElement1(Nothing, startPoint, Matrix3dIdentity)
                    txt.color = 0
                    For i = 0 To UBound(strLines)
                        strLine = strLines(i)
                        txt.AddTextLine strLine
                    Next i
                    ActiveModelReference.AddElement txt
                    If PrintOptions.ShowDrawing Then DoEvents
                    
                    Dim oSubElements As ElementEnumerator: Set oSubElements = txt.GetSubElements
                    Dim oSubElem As element
                    Dim oStyle As TextStyle
                    Do While oSubElements.MoveNext
                        Set oSubElem = oSubElements.Current
                        If oSubElem.Type = msdElementTypeText Then
                            Set oStyle = oSubElem.AsTextElement.TextStyle
                            oStyle.BorderAndBackgroundVisible = True
                            oStyle.BackgroundFillColor = 255
                            oStyle.BorderColor = 255
                            
                            If oSubElem.AsTextElement.text = "INSTALL" Or oSubElem.AsTextElement.text = "REMOVE" Or oSubElem.AsTextElement.text = "REPLACE" Or oSubElem.AsTextElement.text = "TRANSFER" Then
                                oStyle.IsUnderlined = True
                            End If
                            oSubElem.AsTextElement.TextStyle = oStyle
                            oSubElem.Rewrite
                            If PrintOptions.ShowDrawing Then DoEvents
                        End If
                    Loop
                    'txt.Redraw
                    'If PrintOptions.ShowDrawing Then DoEvents
                    
                    Call placeLocation(Point3dAdd(startPoint, OFFSET), jsonPole("location"))
                    
                    startPoint.y = startPoint.y - ((UBound(strLines) + 2) * 15) - (45)
                    groupStartingPoints(group) = startPoint
                End If
            End If
        Next jsonPole
    End If
    
    ActiveSettings.TextStyle = oSettings.TextStyle
    
    MsgBox "Done Placing Poles"
End Sub

Sub ShortenLines(targetLevelName As String)
    Dim oScan As ElementScanCriteria
    Dim oEnumerator As ElementEnumerator
    Dim oElement As element
    Dim lines() As LineElement
    Dim pIntersectArray() As Point3d
    Dim count As Long
    Dim i As Long, j As Long
    Dim pIntersect As Point3d
    Dim status As Boolean
    
    On Error Resume Next
    Set oLevel = ActiveDesignFile.Levels(targetLevelName)
    On Error GoTo 0
    
    If oLevel Is Nothing Then
        MsgBox "Level '" & targetLevelName & "' not found in this DGN file.", vbCritical
        Exit Sub
    End If
    
    Set oScan = New ElementScanCriteria
    oScan.ExcludeAllTypes
    oScan.IncludeType msdElementTypeLine
    oScan.ExcludeAllLevels
    oScan.IncludeLevel oLevel
    
    Set oEnumerator = ActiveModelReference.Scan(oScan)
    
    count = 0
    While oEnumerator.MoveNext
        ReDim Preserve lines(count)
        Set lines(count) = oEnumerator.Current.AsLineElement
        count = count + 1
    Wend
    
    For i = 0 To count - 2
        For j = i + 1 To count - 1
            pIntersectArray = lines(i).GetIntersectionPoints(lines(j), Matrix3dIdentity)

            If (Not Not pIntersectArray) <> 0 Then
                If (UBound(pIntersectArray) >= LBound(pIntersectArray)) Then
                    pIntersect = pIntersectArray(LBound(pIntersectArray))
                    AdjustLineLength lines(i), pIntersect
                    AdjustLineLength lines(j), pIntersect
                End If
            End If
        Next j
    Next i

    'Dim ofeature As Object
    Dim ofeature As xft.feature
    Dim oFeatureMgr As Object
    Set oFeatureMgr = CreateObject("xft.FeatureMgr")
    For i = 0 To count - 1
        On Error Resume Next
        
        If Not existingLineIDs.exists(CStr(lines(i).ID64)) Then
            Set ofeature = oFeatureMgr.CreateFeature(lines(i))
            If ofeature.GetProperty("StartDeadend") Then
                If Err.Number = 0 Then Call createDeadend(lines(i), 0)
            End If
            If ofeature.GetProperty("EndDeadend") Then
                If Err.Number = 0 Then Call createDeadend(lines(i), 1)
            End If
        End If
        
        On Error GoTo 0
    Next i
    
End Sub

Private Sub AdjustLineLength(oLine As LineElement, pInt As Point3d)
    Dim pStart As Point3d, pEnd As Point3d
    Dim targetLevel As Level
    
    pStart = oLine.startPoint
    pEnd = oLine.endPoint
    Set targetLevel = oLine.Level

    If DistanceBetweenPoints(pStart, pInt) < DistanceBetweenPoints(pEnd, pInt) Then
        If Not IsVertexConnected(pStart, oLine, targetLevel) Then
            Call oLine.ModifyVertex(0, pInt)
            oLine.Rewrite
            If PrintOptions.ShowDrawing Then DoEvents
        End If
    Else
        If Not IsVertexConnected(pEnd, oLine, targetLevel) Then
            Call oLine.ModifyVertex(1, pInt)
            oLine.Rewrite
            If PrintOptions.ShowDrawing Then DoEvents
        End If
    End If
End Sub

Private Function IsVertexConnected(pVertex As Point3d, oCurrentLine As LineElement, oLevel As Level) As Boolean
    Dim oScan As New ElementScanCriteria
    Dim oEnumerator As ElementEnumerator
    Dim oCheckLine As LineElement
    Dim pCheckStart As Point3d, pCheckEnd As Point3d
    Dim tolerance As Double
    
    tolerance = 0.0001
    IsVertexConnected = False
    
    oScan.ExcludeAllTypes
    oScan.IncludeType msdElementTypeLine
    oScan.ExcludeAllLevels
    oScan.IncludeLevel oLevel
    
    Set oEnumerator = ActiveModelReference.Scan(oScan)
    
    While oEnumerator.MoveNext
        Set oCheckLine = oEnumerator.Current.AsLineElement
        
        If oCheckLine.id.Low <> oCurrentLine.id.Low Or oCheckLine.id.High <> oCurrentLine.id.High Then
            
            pCheckStart = oCheckLine.startPoint
            pCheckEnd = oCheckLine.endPoint

            If DistanceBetweenPoints(pVertex, pCheckStart) < tolerance Or _
               DistanceBetweenPoints(pVertex, pCheckEnd) < tolerance Then
                
                IsVertexConnected = True
                Exit Function
            End If
        End If
    Wend
End Function
Private Function DistanceBetweenPoints(p1 As Point3d, p2 As Point3d) As Double
    DistanceBetweenPoints = Sqr((p2.x - p1.x) ^ 2 + (p2.y - p1.y) ^ 2)
End Function

Public Sub createDeadend(oLine As LineElement, vertexIndex As Integer)
    Dim pStart As Point3d
    Dim pEnd As Point3d
    Dim pNewTarget As Point3d
    Dim vDirection As Point3d
    Dim lineLength As Double
    Dim distanceToShorten As Double
    Dim rMatrix As Matrix3d
    
    distanceToShorten = 10
    
    pStart = oLine.startPoint
    pEnd = oLine.endPoint
    
    vDirection.x = pEnd.x - pStart.x
    vDirection.y = pEnd.y - pStart.y
    vDirection.Z = pEnd.Z - pStart.Z
    
    lineLength = Sqr(vDirection.x ^ 2 + vDirection.y ^ 2 + vDirection.Z ^ 2)
    
    If lineLength <= distanceToShorten Then
        Exit Sub
    End If
    
    vDirection.x = vDirection.x / lineLength
    vDirection.y = vDirection.y / lineLength
    vDirection.Z = vDirection.Z / lineLength
    
    If vertexIndex = 0 Then
        pNewTarget.x = pStart.x + (vDirection.x * distanceToShorten)
        pNewTarget.y = pStart.y + (vDirection.y * distanceToShorten)
        pNewTarget.Z = pStart.Z + (vDirection.Z * distanceToShorten)
        
        oLine.ModifyVertex 0, pNewTarget
    ElseIf vertexIndex = 1 Then
        pNewTarget.x = pEnd.x - (vDirection.x * distanceToShorten)
        pNewTarget.y = pEnd.y - (vDirection.y * distanceToShorten)
        pNewTarget.Z = pEnd.Z - (vDirection.Z * distanceToShorten)
        
        oLine.ModifyVertex 1, pNewTarget
    End If
    
    oLine.Rewrite
    'oLine.Redraw msdDrawingModeNormal
    'If PrintOptions.ShowDrawing Then DoEvents
    
    Dim dx As Double, dy As Double
    dx = pEnd.x - pStart.x
    dy = pEnd.y - pStart.y
    lineAngle = Atn2(dy, dx)
    rMatrix = Matrix3dFromAxisAndRotationAngle(2, lineAngle)
    
    Dim scl As Point3d: scl = Point3dFromXYZ(1, 1, 1)
    Dim deadend As CellElement
    If oLine.Level.name = "CE-EX-ELEC-OH-PRI-COND" Or oLine.Level.name = "CE-RP-ELEC-OH-PRI-COND" Then Set deadend = CreateCellElement2("TERM", pNewTarget, scl, True, rMatrix)
    If oLine.Level.name = "CE-EX-ELEC-OH-SEC-COND" Or oLine.Level.name = "CE-RP-ELEC-OH-SEC-COND" Then Set deadend = CreateCellElement2("TERM_SEC", pNewTarget, scl, True, rMatrix)
    deadend.color = oLine.color
    ActiveModelReference.AddElement deadend
    If PrintOptions.ShowDrawing Then DoEvents
    
    Dim ofeature As Object
    Dim oFeatureMgr As Object
    Set oFeatureMgr = CreateObject("xft.FeatureMgr")
    Set ofeature = oFeatureMgr.CreateFeature(deadend)
    
    If oLine.Level.name = "CE-EX-ELEC-OH-PRI-COND" Or oLine.Level.name = "CE-RP-ELEC-OH-PRI-COND" Then ofeature.name = "CE_PRI_DE"
    If oLine.Level.name = "CE-EX-ELEC-OH-SEC-COND" Or oLine.Level.name = "CE-RP-ELEC-OH-SEC-COND" Then ofeature.name = "CE_SEC_DE"
    
    ofeature.Write (True)
    'deadend.Redraw msdDrawingModeNormal
    'If PrintOptions.ShowDrawing Then DoEvents
End Sub

Function ReadJSON(FileName As String) As Object
    Dim filePath As String
    Dim jsonText As String
    Dim fileNum As Integer
    
    filePath = "C:\ProgramData\Bentley\OpenUtilities Map Connect Edition\Configuration\WorkSpaces\ConsumersEnergy\Standards\vba\" & FileName & ".json"

    fileNum = FreeFile
    
    On Error Resume Next
    Open filePath For Input As #fileNum
        jsonText = Input$(LOF(fileNum), fileNum)
    Close #fileNum
    On Error GoTo 0

    If jsonText = "" Then Exit Function

    Set ReadJSON = JsonConverter.ParseJson(jsonText)
End Function

Public Sub placePole(jsonPole As Object)
    Dim pt As Point3d
    Dim txt As TextElement
    Dim location As CellElement
    Dim pole As CellElement
    Dim Tree As CellElement
    Dim scl As Point3d: scl = Point3dFromXYZ(1, 1, 1)
    
    pt = Point3dFromXYZ(jsonPole("x"), jsonPole("y"), 0)
    
    If jsonPole("ceid") = "FOREIGN" Then
        If jsonPole("replace") Then
            Set pole = CreateCellElement2("TPOLRP", pt, scl, True, Matrix3dIdentity)
            Set txt = CreateTextElement1(Nothing, IIf(jsonPole("skipSpan"), "SKIP SPAN" & vbLf, "") & "RP " & jsonPole("height") & "-" & jsonPole("class") & "/" & jsonPole("newHeight") & "-" & jsonPole("newClass"), pt, Matrix3dIdentity)
            txt.color = 65
        Else
            Set pole = CreateCellElement2("TPOLE", pt, scl, True, Matrix3dIdentity)
            Set txt = CreateTextElement1(Nothing, IIf(jsonPole("skipSpan"), "SKIP SPAN" & vbLf, "") & jsonPole("height") & "-" & jsonPole("class"), pt, Matrix3dIdentity)
            txt.color = 73
        End If
    Else
        If jsonPole("replace") Then
            Set pole = CreateCellElement2("POLERP", pt, scl, True, Matrix3dIdentity)
            Set txt = CreateTextElement1(Nothing, IIf(jsonPole("skipSpan"), "SKIP SPAN" & vbLf, "") & "RP " & jsonPole("height") & "-" & jsonPole("class") & "/" & jsonPole("newHeight") & "-" & jsonPole("newClass") & vbLf & "CE:" & jsonPole("ceid"), pt, Matrix3dIdentity)
            txt.color = 65
        Else
            Set pole = CreateCellElement2("POLE", pt, scl, True, Matrix3dIdentity)
            If jsonPole("hvd") <> "" Then
                Set txt = CreateTextElement1(Nothing, "T-" & jsonPole("height") & vbLf & "#" & jsonPole("hvd"), pt, Matrix3dIdentity)
            Else
                Set txt = CreateTextElement1(Nothing, IIf(jsonPole("skipSpan"), "SKIP SPAN" & vbLf, "") & jsonPole("height") & "-" & jsonPole("class") & vbLf & "CE:" & jsonPole("ceid"), pt, Matrix3dIdentity)
            End If
            txt.color = 73
        End If
    End If
    
    txt.TextStyle.BorderAndBackgroundVisible = True
    txt.TextStyle.BackgroundFillColor = 255
    txt.TextStyle.BorderColor = 255

    On Error Resume Next
    txt.Level = ActiveDesignFile.Levels("CE-EX-OH-PRI-POLE")
    On Error GoTo 0
    
    ActiveModelReference.AddElement pole
    If PrintOptions.ShowDrawing Then DoEvents
    
    Dim ofeature As Object
    Dim oFeatureMgr As Object
    Set oFeatureMgr = CreateObject("xft.FeatureMgr")

    Set ofeature = oFeatureMgr.CreateFeature(pole)
    ofeature.name = "CE_SUPPORTSTRUCTURE"

    If jsonPole("replace") Then
        ofeature.SetProperty "LIFECYCLE", 4
    Else
        ofeature.SetProperty "LIFECYCLE", 1
    End If

    ofeature.SetProperty "SUBTYPECD", 7
    
    If jsonPole("ceid") = "FOREIGN" Then
        ofeature.SetProperty "OWNER", "Foreign"
    Else
        ofeature.SetProperty "OWNER", "Consumers Energy"
    End If
    
    If jsonPole("hvd") <> "" Then ofeature.SetProperty "HVD_TAG", jsonPole("hvd")
    
    If Not IsNull(jsonPole("height")) Then ofeature.SetProperty "HEIGHT", jsonPole("height")
    If Not IsNull(jsonPole("class")) Then ofeature.SetProperty "CLASS", jsonPole("class")
    If Not IsNull(jsonPole("ceid")) Then ofeature.SetProperty "CE_TAG", jsonPole("ceid")
    ofeature.SetProperty "JOINT", "No"

    ofeature.Write (True)
    'pole.Redraw msdDrawingModeNormal
    'If PrintOptions.ShowDrawing Then DoEvents

    Dim scaleMatrix As Point3d
    scaleMatrix = Point3dFromXYZ(4, 4, 1)
    If PrintOptions.DrawTrees Then
        If jsonPole("tree") Then
            Set Tree = CreateCellElement2("TREE1", pt, scaleMatrix, True, Matrix3dIdentity)
            ActiveModelReference.AddElement Tree
            If PrintOptions.ShowDrawing Then DoEvents
            
            Set ofeature = oFeatureMgr.CreateFeature(Tree)
            ofeature.name = "CE_TREE"
            ofeature.Write (True)
            'tree.Redraw msdDrawingModeNormal
            'If PrintOptions.ShowDrawing Then DoEvents
        End If
    End If
    
    If PrintOptions.DrawDownGuys Then
        Dim guyAngleDict As Object: Set guyAngleDict = CreateObject("Scripting.Dictionary")
        If jsonPole.exists("guys") Then
            Dim jsonGuy As Object
            For Each jsonGuy In jsonPole("guys")
                If Not guyAngleDict.exists(jsonGuy("angle")) Then guyAngleDict(jsonGuy("angle")) = -1
                guyAngleDict(jsonGuy("angle")) = guyAngleDict(jsonGuy("angle")) + 1
                Call placeGuy(jsonGuy, pt, guyAngleDict(jsonGuy("angle")))
            Next jsonGuy
        End If
    End If

    Call txt.Move(Point3dFromXYZ(5, -10, 0))
    ActiveModelReference.AddElement txt
    If PrintOptions.ShowDrawing Then DoEvents
    

    Dim closestPt As Point3d
    If jsonPole("location") <> "" Then
        Call placeLocation(Point3dAdd(pt, Point3dFromXYZ(-15, -20, 0)), jsonPole("location"))
        If PrintOptions.DrawCenterLines And PrintOptions.DrawCenterLineDistances Then
            results = FindClosestPointOnCenterlines(pt)
            closestPt = results(0)
            If closestPt.x <> -1 And closestPt.y <> -1 And closestPt.Z <> -1 Then
                Dim rMatrix As Matrix3d
                rMatrix = results(1)
                Call CreateLinearDimensionFullyCorrected(pt, closestPt, rMatrix)
            End If
        End If
    End If
End Sub

Sub placeTransformer(jsonTransformer)
    Dim pt As Point3d
    Dim transformer As CellElement
    Dim scl As Point3d: scl = Point3dFromXYZ(1, 1, 1)
    Dim size2 As String
    Dim ofeature As Object
    Dim oFeatureMgr As Object
    Set oFeatureMgr = CreateObject("xft.FeatureMgr")
    
    Dim phase As String
    phase = jsonTransformer("phase")
    size2 = jsonTransformer("size2")
    
    pt = Point3dFromXYZ(jsonTransformer("x"), jsonTransformer("y"), 0)
    pt = Point3dAdd(pt, Point3dFromXYZ(10, 40, 0))
    
    ' 1. Temporarily create the base cell in memory (not placed yet)
    Dim cellName As String
    If phase = "X" Or phase = "Y" Or phase = "Z" Then
        cellName = "TRF1"
    ElseIf phase = "XY" Or phase = "YZ" Or phase = "XZ" Then
        cellName = "TRF2"
    ElseIf phase = "3P" Then
        cellName = "TRF3"
    End If
    
    Set transformer = CreateCellElement2(cellName & IIf(size2 <> "", "RP", ""), pt, scl, True, Matrix3dIdentity)
    
    ActiveModelReference.AddElement transformer
    If PrintOptions.ShowDrawing Then DoEvents

    Set ofeature = oFeatureMgr.CreateFeature(transformer)
    ofeature.name = "CE_OH_XFRM"
    If size2 <> "" Then
        'oFeature.SetProperty "LIFECYCLE", 4
    Else
        'oFeature.SetProperty "LIFECYCLE", 1
    End If
    'oFeature.SetProperty "PHASE", phase
    'oFeature.SetProperty "TLM", jsonTransformer("TLM")
    'oFeature.SetProperty "SIZE", jsonTransformer("size")

    Call ofeature.Write(False)
    transformer.Redraw msdDrawingModeNormal
    If PrintOptions.ShowDrawing Then DoEvents

    Dim txt As TextElement
    Dim txtString As String
    If size2 <> "" Then
        If jsonTransformer("size") = size2 Then
            txtString = "RP " & jsonTransformer("size") & IIf(phase <> "3P", phase, "") & vbLf & jsonTransformer("TLM") & IIf(Len(phase) > 1, vbLf & jsonTransformer("lowSideVoltage"), "")
        Else
            txtString = "RP " & jsonTransformer("size") & "/" & size2 & IIf(phase <> "3P", phase, "") & vbLf & jsonTransformer("TLM") & IIf(Len(phase) > 1, vbLf & jsonTransformer("lowSideVoltage"), "")
        End If
    Else
        txtString = jsonTransformer("size") & IIf(phase <> "3P", phase, "") & vbLf & jsonTransformer("TLM") & IIf(Len(phase) > 1, vbLf & jsonTransformer("lowSideVoltage"), "")
    End If
    
    Dim txtOffset As Point3d
    If cellName = "TRF1" Then txtOffset = Point3dFromXYZ(-5, -5, 0)
    If cellName = "TRF2" Then txtOffset = Point3dFromXYZ(-15, -5, 0)
    If cellName = "TRF3" Then txtOffset = Point3dFromXYZ(-15, -10, 0)
    Set txt = CreateTextElement1(Nothing, txtString, Point3dAdd(pt, txtOffset), Matrix3dIdentity)
    If size2 <> "" Then
        txt.color = 64
    Else
        txt.color = 72
    End If
    txt.TextStyle.BorderAndBackgroundVisible = True
    txt.TextStyle.BackgroundFillColor = 255
    txt.TextStyle.BorderColor = 255
    ActiveModelReference.AddElement txt
    If PrintOptions.ShowDrawing Then DoEvents
End Sub

Sub placeCapacitor(jsonCapacitor)
    Dim pt As Point3d
    Dim capacitor As CellElement
    Dim scl As Point3d: scl = Point3dFromXYZ(1, 1, 1)
    
    Dim ofeature As Object
    Dim oFeatureMgr As Object
    Set oFeatureMgr = CreateObject("xft.FeatureMgr")
    
    pt = Point3dFromXYZ(jsonCapacitor("x"), jsonCapacitor("y"), 0)
    pt = Point3dAdd(pt, Point3dFromXYZ(10, 40, 0))
    
    If jsonCapacitor("switched") Then
        Set capacitor = CreateCellElement2("CAPSW", pt, scl, True, Matrix3dIdentity)
    Else
        Set capacitor = CreateCellElement2("CAPU", pt, scl, True, Matrix3dIdentity)
    End If
    
    ActiveModelReference.AddElement capacitor
    If PrintOptions.ShowDrawing Then DoEvents
    
    Set ofeature = oFeatureMgr.CreateFeature(capacitor)
    ofeature.name = "CE_CAPACITOR"
    ofeature.SetProperty "LIFECYCLE", 1
    ofeature.SetProperty "LCP", jsonCapacitor("lcp")
    If jsonCapacitor("lcp") <> "" Then ofeature.SetProperty "LCP", jsonCapacitor("lcp")

    ofeature.Write (True)
    'capacitor.Redraw msdDrawingModeNormal
    'If PrintOptions.ShowDrawing Then DoEvents

    Dim txt As TextElement
    Dim txtString As String
    txtString = jsonCapacitor("lcp") & vbLf & jsonCapacitor("size") & "kVAR"
    
    Dim txtOffset As Point3d
    txtOffset = Point3dFromXYZ(-15, -10, 0)
    Set txt = CreateTextElement1(Nothing, txtString, Point3dAdd(pt, txtOffset), Matrix3dIdentity)
    txt.color = 72
    txt.TextStyle.BorderAndBackgroundVisible = True
    txt.TextStyle.BackgroundFillColor = 255
    txt.TextStyle.BorderColor = 255
    ActiveModelReference.AddElement txt
    If PrintOptions.ShowDrawing Then DoEvents
End Sub

Sub placeRegulator(jsonRegulator)
    Dim pt As Point3d
    Dim regulator As CellElement
    Dim scl As Point3d: scl = Point3dFromXYZ(1, 1, 1)
    
    Dim ofeature As Object
    Dim oFeatureMgr As Object
    Set oFeatureMgr = CreateObject("xft.FeatureMgr")
    
    Dim phase As String
    phase = jsonRegulator("phase")
    
    pt = Point3dFromXYZ(jsonRegulator("x"), jsonRegulator("y"), 0)
    pt = Point3dAdd(pt, Point3dFromXYZ(10, 40, 0))
    
    Dim cellName As String
    If jsonRegulator("auto") Then
        If phase = "X" Or phase = "Y" Or phase = "Z" Then
            cellName = "AUTOB1EX"
        ElseIf phase = "XY" Or phase = "YZ" Or phase = "XZ" Then
            cellName = "AUTOB2EX"
        ElseIf phase = "3P" Then
            cellName = "AUTOB3EX"
        End If
    ElseIf jsonRegulator("fixed") Then
        If phase = "X" Or phase = "Y" Or phase = "Z" Then
            cellName = "BOOST1EX"
        ElseIf phase = "XY" Or phase = "YZ" Or phase = "XZ" Then
            cellName = "BOOST2EX"
        ElseIf phase = "3P" Then
            cellName = "BOOST3EX"
        End If
    Else
        If phase = "X" Or phase = "Y" Or phase = "Z" Then
            cellName = "REG1EX"
        ElseIf phase = "XY" Or phase = "YZ" Or phase = "XZ" Then
            cellName = "REG2EX"
        ElseIf phase = "3P" Then
            cellName = "REG3EX"
        End If
    End If
    
    Set regulator = CreateCellElement2(cellName, pt, scl, True, Matrix3dIdentity)
    
    ActiveModelReference.AddElement regulator
    If PrintOptions.ShowDrawing Then DoEvents
    
    Set ofeature = oFeatureMgr.CreateFeature(regulator)
    If jsonRegulator("auto") Then
        ofeature.name = "CE_AUTOBOOSTER"
    ElseIf jsonRegulator("fixed") Then
        ofeature.name = "CE_FIXEDBOOSTER"
    Else
        ofeature.name = "CE_REGULATOR"
    End If
    ofeature.SetProperty "LIFECYCLE", 1
    ofeature.SetProperty "PHASE", phase
    ofeature.SetProperty "SIZE", jsonRegulator("size") & "A"
    If jsonRegulator("lcp") <> "" Then ofeature.SetProperty "LCP", jsonRegulator("lcp")

    ofeature.Write (True)
    'regulator.Redraw msdDrawingModeNormal
    'If PrintOptions.ShowDrawing Then DoEvents

    Dim txt As TextElement
    Dim txtString As String
    If jsonRegulator("auto") Then
        txtString = jsonRegulator("lcp") & vbLf & jsonRegulator("size") & "A"
    Else
        txtString = jsonRegulator("lcp") & vbLf & jsonRegulator("size") & "KVA"
    End If
    
    Dim txtOffset As Point3d
    txtOffset = Point3dFromXYZ(-15, -10, 0)
    Set txt = CreateTextElement1(Nothing, txtString, Point3dAdd(pt, txtOffset), Matrix3dIdentity)
    txt.color = 72
    txt.TextStyle.BorderAndBackgroundVisible = True
    txt.TextStyle.BackgroundFillColor = 255
    txt.TextStyle.BorderColor = 255
    ActiveModelReference.AddElement txt
    If PrintOptions.ShowDrawing Then DoEvents
End Sub

Sub placeIsolator(jsonIsolator)
    Dim pt As Point3d
    Dim isolator As CellElement
    Dim scl As Point3d: scl = Point3dFromXYZ(1, 1, 1)
    
    Dim ofeature As Object
    Dim oFeatureMgr As Object
    Set oFeatureMgr = CreateObject("xft.FeatureMgr")
    
    Dim phase As String
    phase = jsonIsolator("phase")
    
    pt = Point3dFromXYZ(jsonIsolator("x"), jsonIsolator("y"), 0)
    pt = Point3dAdd(pt, Point3dFromXYZ(10, 40, 0))
    
    Dim cellName As String
    If phase = "X" Or phase = "Y" Or phase = "Z" Then
        cellName = "ISO1EX"
    ElseIf phase = "XY" Or phase = "YZ" Or phase = "XZ" Then
        cellName = "ISO2EX"
    ElseIf phase = "3P" Then
        cellName = "ISO3EX"
    End If
    
    Set isolator = CreateCellElement2(cellName, pt, scl, True, Matrix3dIdentity)
    
    ActiveModelReference.AddElement isolator
    If PrintOptions.ShowDrawing Then DoEvents
    
    Set ofeature = oFeatureMgr.CreateFeature(isolator)
    ofeature.name = "CE_ISOLATOR"
    ofeature.SetProperty "LIFECYCLE", 1
    ofeature.SetProperty "LCP", jsonIsolator("lcp")
    If jsonIsolator("lcp") <> "" Then ofeature.SetProperty "LCP", jsonIsolator("lcp")

    ofeature.Write (True)
    'capacitor.Redraw msdDrawingModeNormal
    'If PrintOptions.ShowDrawing Then DoEvents

    Dim txt As TextElement
    Dim txtString As String
    txtString = jsonIsolator("lcp") & vbLf & jsonIsolator("size") & "KVA"
    
    Dim txtOffset As Point3d
    txtOffset = Point3dFromXYZ(-15, -10, 0)
    Set txt = CreateTextElement1(Nothing, txtString, Point3dAdd(pt, txtOffset), Matrix3dIdentity)
    txt.color = 72
    txt.TextStyle.BorderAndBackgroundVisible = True
    txt.TextStyle.BackgroundFillColor = 255
    txt.TextStyle.BorderColor = 255
    ActiveModelReference.AddElement txt
    If PrintOptions.ShowDrawing Then DoEvents
End Sub

Sub placeStreetlight(jsonStreetlight As Object)
    Dim pt As Point3d
    Dim closestPt As Point3d
    Dim streetlight As CellElement
    Dim scl As Point3d: scl = Point3dFromXYZ(1, 1, 1)
    
    pt = Point3dFromXYZ(jsonStreetlight("x"), jsonStreetlight("y"), 0)
    
    results = FindClosestPointOnCenterlines(pt)
    closestPt = results(0)
    If closestPt.x <> -1 And closestPt.y <> -1 And closestPt.Z <> -1 Then
        Dim rMatrix As Matrix3d
        
        Dim directionVector As Point3d
        Dim baseCellVector As Point3d
        directionVector = Point3dSubtract(closestPt, pt)
        Dim calculatedAngle As Double

        directionVector.Z = 0
        baseCellVector = Point3dFromXYZ(1, 0, 0)
        calculatedAngle = Point3dAngleBetweenVectors(baseCellVector, directionVector)
        If directionVector.y < 0 Then
            calculatedAngle = -calculatedAngle
        End If
    
        rMatrix = Matrix3dFromAxisAndRotationAngle(2, calculatedAngle)
        
        Dim offsetAngle As Double: offsetAngle = calculatedAngle + 45
        pt.x = pt.x + Cos(offsetAngle) * 5
        pt.y = pt.y + Sin(offsetAngle) * 5

        Set streetlight = CreateCellElement2("LITE", pt, scl, True, rMatrix)
    Else
        Set streetlight = CreateCellElement2("LITE", pt, scl, True, Matrix3dIdentity)
    End If
    
    ActiveModelReference.AddElement streetlight
    If PrintOptions.ShowDrawing Then DoEvents

    Dim ofeature As Object
    Dim oFeatureMgr As Object
    Set oFeatureMgr = CreateObject("xft.FeatureMgr")
    Set ofeature = oFeatureMgr.CreateFeature(streetlight)
    ofeature.name = "CE_STREETLIGHT"
    ofeature.SetProperty "LIFECYCLE", 1
    ofeature.SetProperty "MOUNT_TYPE", "Bracket"
    ofeature.SetProperty "LIGHT_TYPE", "S"
    ofeature.SetProperty "SIZE", "40"

    ofeature.Write (True)
    'streetlight.Redraw msdDrawingModeNormal
    'If PrintOptions.ShowDrawing Then DoEvents
End Sub

Sub placeLocation(pt As Point3d, location As String)
    Dim scl As Point3d: scl = Point3dFromXYZ(1, 1, 1)
    Set oCell = CreateCellElement2("LOCATION_NUMBER", pt, scl, True, Matrix3dIdentity)
    
    Dim subElements() As element
    Dim subCount As Long
    subCount = 0

    Set oEnum = oCell.GetSubElements
    Do While oEnum.MoveNext
        Set oSubEl = oEnum.Current
        
        If oSubEl.IsTextElement Then
            Set oTextEl = oSubEl.AsTextElement
            If oTextEl.text = "[W]" Then
                oTextEl.text = location
            End If
            Set oSubEl = oTextEl
        End If
        
        ReDim Preserve subElements(subCount)
        Set subElements(subCount) = oSubEl
        subCount = subCount + 1
    Loop
    
    Dim oFinalCell As CellElement
    Set oFinalCell = CreateCellElement1("MyCellName", subElements, pt, False)
    
    ActiveModelReference.AddElement oFinalCell
    If PrintOptions.ShowDrawing Then DoEvents
    
    Dim ofeature As Object
    Dim oFeatureMgr As Object
    Set oFeatureMgr = CreateObject("xft.FeatureMgr")
    
    Set ofeature = oFeatureMgr.CreateFeature(oFinalCell)
    ofeature.name = "CE_LOCATION_NUMBER"
    
    ofeature.SetProperty "LOCATION_NUMBER", location
    
    ofeature.Write (True)
    'oFinalCell.Redraw msdDrawingModeNormal
    'If PrintOptions.ShowDrawing Then DoEvents
End Sub

Public Sub placeGuy(jsonGuy As Object, pt As Point3d, OFFSET As Integer)
    Dim guy As CellElement
    Dim scl As Point3d: scl = Point3dFromXYZ(1, 1, 1)
    Dim guyPt As Point3d
    
    angle = jsonGuy("angle")
    angle = 180 - 90 - angle
    
    guyPt.x = pt.x + OFFSET * 25 * Cos(angle * (PI / 180))
    guyPt.y = pt.y + OFFSET * 25 * Sin(angle * (PI / 180))
    guyPt.Z = 0
    
    If jsonGuy("count") < 2 Then
        guyType = "S"
    ElseIf jsonGuy("count") = 1 Then
        guyType = "D"
    Else
        guyType = "T"
    End If
    
    If jsonGuy("replace") Then
        Set guy = CreateCellElement2(guyType & "GUYRP", guyPt, scl, True, Matrix3dFromAxisAndRotationAngle(2, Radians(angle)))
    Else
        Set guy = CreateCellElement2(guyType & "GUY", guyPt, scl, True, Matrix3dFromAxisAndRotationAngle(2, Radians(angle)))
    End If
    
    ActiveModelReference.AddElement guy
    If PrintOptions.ShowDrawing Then DoEvents
    
    Dim ofeature As Object
    Dim oFeatureMgr As Object
    Set oFeatureMgr = CreateObject("xft.FeatureMgr")

    Set ofeature = oFeatureMgr.CreateFeature(guy)
    ofeature.name = "CE_GUY"

    If jsonGuy("replace") Then
        ofeature.SetProperty "LIFECYCLE", 4
    Else
        ofeature.SetProperty "LIFECYCLE", 1
    End If

    ofeature.SetProperty "TYPE", jsonGuy("count")
    ofeature.SetProperty "FOREIGN", "No"

    ofeature.Write (True)
    'guy.Redraw msdDrawingModeNormal
    'If PrintOptions.ShowDrawing Then DoEvents
End Sub

Public Sub drawService(jsonService As Object)
    Dim startPoint As Point3d
    Dim endPoint As Point3d
    Dim angle As Double
    
    startPoint.x = jsonService("x")
    startPoint.y = jsonService("y")
    startPoint.Z = 0
    
    distance = jsonService("distance")
    angle = jsonService("angle")
    angle = 90 - angle
    
    startPoint.x = startPoint.x + 5 * Cos(angle * (PI / 180))
    startPoint.y = startPoint.y + 5 * Sin(angle * (PI / 180))
    startPoint.Z = 0
    
    endPoint.x = startPoint.x + (distance - 5) * Cos(angle * (PI / 180))
    endPoint.y = startPoint.y + (distance - 5) * Sin(angle * (PI / 180))
    endPoint.Z = 0
    
    Dim oLine As LineElement
    Set oLine = CreateLineElement2(Nothing, startPoint, endPoint)
    oLine.color = 73

    ActiveModelReference.AddElement oLine
    If PrintOptions.ShowDrawing Then DoEvents
    If jsonService("ug") Then Call PlaceTextAboveAndBelow(oLine, "", "UG", True, True)
    Call drawAddress(endPoint, jsonService("address"), angle)
End Sub

Public Sub drawAddress(pt As Point3d, address As String, angle As Double)
    Dim txt As TextElement
    Dim closestPt As Point3d
    Dim rMatrix As Matrix3d
    Dim directionVector As Point3d
    Dim baseCellVector As Point3d
    Dim calculatedAngle As Double
    Dim vectorBetween As Vector3d
    Dim xAxis As Vector3d
    Dim angleRadians As Double
    Dim angleDeg As Double

    If address <> "" Then
        results = FindClosestPointOnCenterlines(pt)
        closestPt = results(0)
        If closestPt.x <> -1 And closestPt.y <> -1 And closestPt.Z <> -1 Then
            vectorBetween = Vector3dSubtractPoint3dPoint3d(pt, closestPt)
            xAxis = Vector3dFromXYZ(1, 0, 0)
            angleRadians = Vector3dAngleBetweenVectors(vectorBetween, xAxis)
            
            Dim rotationAxis As Point3d
            rotationAxis = Point3dFromXYZ(0, 0, 1)
            
            angleDeg = angleRadians * (180 / PI)
            If angleDeg > 70 And angleDeg <= 315 Then
                angleRadians = Abs(angleRadians - PI)
            End If
            
            rMatrix = Matrix3dFromVectorAndRotationAngle(rotationAxis, angleRadians)
                    
            Set txt = CreateTextElement1(Nothing, address, pt, rMatrix)

        Else
            Set txt = CreateTextElement1(Nothing, address, pt, Matrix3dIdentity)
        End If
        
        
        Call AlignTextEdgeToPointB(txt, closestPt, pt)
        Call txt.Move(Point3dFromXYZ(5 * Cos(angle * (PI / 180)), 5 * Sin(angle * (PI / 180)), 0))
        
        txt.color = 51
        txt.LineWeight = 2
        txt.TextStyle.BorderAndBackgroundVisible = True
        txt.TextStyle.BackgroundFillColor = 255
        txt.TextStyle.BorderColor = 255
    
        ActiveModelReference.AddElement txt
        If PrintOptions.ShowDrawing Then DoEvents
    End If
End Sub

Public Sub AlignTextEdgeToPointB(ByRef txtEl As TextElement, pointA As Point3d, pointB As Point3d)
    Dim rng As Range3d
    rng = txtEl.Range
    
    Dim midTop As Point3d, midBottom As Point3d, midLeft As Point3d, midRight As Point3d
    

    midTop.x = (rng.Low.x + rng.High.x) / 2
    midTop.y = rng.High.y
    midTop.Z = (rng.Low.Z + rng.High.Z) / 2
    
    midBottom.x = (rng.Low.x + rng.High.x) / 2
    midBottom.y = rng.Low.y
    midBottom.Z = (rng.Low.Z + rng.High.Z) / 2
    
    midLeft.x = rng.Low.x
    midLeft.y = (rng.Low.y + rng.High.y) / 2
    midLeft.Z = (rng.Low.Z + rng.High.Z) / 2
    
    midRight.x = rng.High.x
    midRight.y = (rng.Low.y + rng.High.y) / 2
    midRight.Z = (rng.Low.Z + rng.High.Z) / 2

    Dim closestMidpoint As Point3d
    Dim minDistance As Double
    Dim currentDist As Double
    
    closestMidpoint = midTop
    minDistance = Point3dDistance(midTop, pointA)
    
    currentDist = Point3dDistance(midBottom, pointA)
    If currentDist < minDistance Then
        minDistance = currentDist
        closestMidpoint = midBottom
    End If
    
    currentDist = Point3dDistance(midLeft, pointA)
    If currentDist < minDistance Then
        minDistance = currentDist
        closestMidpoint = midLeft
    End If
    
    currentDist = Point3dDistance(midRight, pointA)
    If currentDist < minDistance Then
        minDistance = currentDist
        closestMidpoint = midRight
    End If
    
    Dim moveVector As Point3d
    moveVector = Point3dSubtract(pointB, closestMidpoint)

    txtEl.Move moveVector
End Sub

Public Sub placeSpanguy(jsonSpanguy As Object)
    Dim oLineString As LineElement
    Dim points(0 To 1) As Point3d
    Dim i As Integer
    Dim top As Boolean
    Dim bottom As Boolean
    
    angle = jsonSpanguy("angle")
    radAngle = (90 - angle) * (PI / 180)
    
    x1 = jsonSpanguy("x1") + (5 * Cos(radAngle))
    x2 = jsonSpanguy("x2") - (5 * Cos(radAngle))
    
    y1 = jsonSpanguy("y1") + (5 * Sin(radAngle))
    y2 = jsonSpanguy("y2") - (5 * Sin(radAngle))
    
    points(0) = Point3dFromXYZ(x1, y1, 0)
    points(1) = Point3dFromXYZ(x2, y2, 0)
    
    Set oLineString = CreateLineElement1(Nothing, points)
    oLineString.LineStyle = ActiveDesignFile.LineStyles.Find("SPN")
    oLineString.Class = primary
    
    On Error Resume Next
    oLineString.Level = ActiveDesignFile.Levels("CE-EX-ELEC-SPANGUY")
    On Error GoTo 0
    
    oLineString.LineWeight = 0
    oLineString.color = 73
    ActiveModelReference.AddElement oLineString
    If PrintOptions.ShowDrawing Then DoEvents

    Dim ofeature As Object
    Dim oFeatureMgr As Object
    Set oFeatureMgr = CreateObject("xft.FeatureMgr")

    Set ofeature = oFeatureMgr.CreateFeature(oLineString)
    ofeature.name = "CE_SPAN_GUY"

    ofeature.SetProperty "SPAN_TYPE", IIf(jsonSpanguy("count") <= 3, jsonSpanguy("count"), 3)
    ofeature.SetProperty "LIFECYCLE", 1

    ofeature.Write (True)
   'oLineString.Redraw msdDrawingModeNormal
    'If PrintOptions.ShowDrawing Then DoEvents
    
    top = jsonSpanguy("top")
    bottom = True
    
    Dim calculatedAngle As Double
    Dim pt As Point3d: pt = Point3dFromXYZ(jsonSpanguy("x1"), jsonSpanguy("y1"), 0)
    Dim closestPt As Point3d
    results = FindClosestPointOnCenterlines(pt)
    closestPt = results(0)
    If closestPt.x <> -1 And closestPt.y <> -1 And closestPt.Z <> -1 Then
        Dim rMatrix As Matrix3d
        
        Dim directionVector As Point3d
        Dim baseCellVector As Point3d
        directionVector = Point3dSubtract(closestPt, pt)

        directionVector.Z = 0
        baseCellVector = Point3dFromXYZ(1, 0, 0)
        calculatedAngle = Point3dAngleBetweenVectors(baseCellVector, directionVector)
        If directionVector.y < 0 Then
            calculatedAngle = -calculatedAngle
        End If
    End If
    
    calculatedAngle = calculatedAngle * (180 / PI)
    If calculatedAngle < 0 Then calculatedAngle = calculatedAngle + 360
    
    angle = 90 - jsonSpanguy("angle")
    If angle < 0 Then angle = angle + 360
    If angle > 90 And angle <= 180 Then angle = 180 - angle
    If angle > 180 And angle <= 270 Then angle = angle - 180
    If angle > 270 And angle <= 360 Then angle = 360 - angle
    
    Dim dx As Double, dy As Double
    dx = jsonSpanguy("x1") - jsonSpanguy("x2")
    dy = jsonSpanguy("y1") - jsonSpanguy("y2")
    
    Dim angleRad As Double
    angleRad = Atn2(dy, dx)
    
    Dim angleDeg As Double
    angleDeg = angleRad * (180 / PI)

    If angleDeg < 0 Then angleDeg = angleDeg + 360
    
    
    If angle > 45 And calculatedAngle >= 135 And calculatedAngle < 225 Then
        If angleDeg <= 95 Or angleDeg > 265 Then
            If Not top Or Not bottom Then
                If top Then
                    top = False
                    bottom = True
                ElseIf bottom Then
                    bottom = False
                    top = True
                End If
            End If
        End If
    ElseIf angle <= 45 And (calculatedAngle < 225 Or calculatedAngle >= 315) Then
        If angleDeg > 95 Or angleDeg <= 265 Then
            If Not top Or Not bottom Then
                If top Then
                    top = False
                    bottom = True
                ElseIf bottom Then
                    bottom = False
                    top = True
                End If
            End If
        End If
    End If
    
    Call PlaceTextAboveAndBelow(oLineString, "11K", jsonSpanguy("length") & "'", top, bottom)
End Sub

Public Sub placeSecondary(jsonSec As Object)
    Dim oLineString As LineElement
    Dim points(0 To 1) As Point3d
    Dim i As Integer
    Dim layer As Integer
    Dim size1 As String
    Dim size2 As String
    Dim top As Boolean
    Dim bottom As Boolean

    size1 = jsonSec("size")
    size2 = jsonSec("size2")
    layer = jsonSec("layer")
    top = jsonSec("top")
    bottom = jsonSec("bottom")
    

    angle = 90 - jsonSec("angle")
    If angle < 0 Then angle = angle + 360
    If angle > 90 And angle <= 180 Then angle = 180 - angle
    If angle > 180 And angle <= 270 Then angle = angle - 180
    If angle > 270 And angle <= 360 Then angle = 360 - angle
    
    Dim calculatedAngle As Double
    Dim pt As Point3d: pt = Point3dFromXYZ(jsonSec("x1"), jsonSec("y1"), 0)
    Dim closestPt As Point3d
    results = FindClosestPointOnCenterlines(pt)
    closestPt = results(0)
    If closestPt.x <> -1 And closestPt.y <> -1 And closestPt.Z <> -1 Then
        Dim rMatrix As Matrix3d
        
        Dim directionVector As Point3d
        Dim baseCellVector As Point3d
        directionVector = Point3dSubtract(closestPt, pt)

        directionVector.Z = 0
        baseCellVector = Point3dFromXYZ(1, 0, 0)
        calculatedAngle = Point3dAngleBetweenVectors(baseCellVector, directionVector)
        If directionVector.y < 0 Then
            calculatedAngle = -calculatedAngle
        End If
    End If
    
    calculatedAngle = calculatedAngle * (180 / PI)
    If calculatedAngle < 0 Then calculatedAngle = calculatedAngle + 360
    
    x1 = jsonSec("x1")
    x2 = jsonSec("x2")
    
    y1 = jsonSec("y1")
    y2 = jsonSec("y2")
    
    Dim dx As Double, dy As Double
    dx = jsonSec("x1") - jsonSec("x2")
    dy = jsonSec("y1") - jsonSec("y2")
    
    Dim angleRad As Double
    angleRad = Atn2(dy, dx)
    
    Dim angleDeg As Double
    angleDeg = angleRad * (180 / PI)

    If angleDeg < 0 Then angleDeg = angleDeg + 360

    If angle > 45 Then
        If (calculatedAngle >= 135) And (calculatedAngle < 225) Then
            x1 = x1 + (PrintOptions.ConductorInitOffset + ((layer - 1) * PrintOptions.ConductorOffsetAmount))
            x2 = x2 + (PrintOptions.ConductorInitOffset + ((layer - 1) * PrintOptions.ConductorOffsetAmount))
            If angleDeg <= 95 Or angleDeg > 265 Then
                If Not top Or Not bottom Then
                    If top Then
                        top = False
                        bottom = True
                    ElseIf bottom Then
                        bottom = False
                        top = True
                    End If
                End If
            End If
        Else
            x1 = x1 - (PrintOptions.ConductorInitOffset + ((layer - 1) * PrintOptions.ConductorOffsetAmount))
            x2 = x2 - (PrintOptions.ConductorInitOffset + ((layer - 1) * PrintOptions.ConductorOffsetAmount))
        End If
    Else
        If (calculatedAngle >= 225) And (calculatedAngle < 315) Then
            y1 = y1 + (PrintOptions.ConductorInitOffset + ((layer - 1) * PrintOptions.ConductorOffsetAmount))
            y2 = y2 + (PrintOptions.ConductorInitOffset + ((layer - 1) * PrintOptions.ConductorOffsetAmount))
        Else
            y1 = y1 - (PrintOptions.ConductorInitOffset + ((layer - 1) * PrintOptions.ConductorOffsetAmount))
            y2 = y2 - (PrintOptions.ConductorInitOffset + ((layer - 1) * PrintOptions.ConductorOffsetAmount))
            If angleDeg > 95 Or angleDeg <= 265 Then
                If Not top Or Not bottom Then
                    If top Then
                        top = False
                        bottom = True
                    ElseIf bottom Then
                        bottom = False
                        top = True
                    End If
                End If
            End If
        End If
    End If
    
    points(0) = Point3dFromXYZ(x1, y1, 0)
    points(1) = Point3dFromXYZ(x2, y2, 0)
    
    Set oLineString = CreateLineElement1(Nothing, points)

    If size2 <> "" Then
        oLineString.color = 65
        oLineString.LineStyle = ActiveDesignFile.LineStyles.Find("SECRP")
        On Error Resume Next
        oLineString.Level = ActiveDesignFile.Levels("CE-RP-ELEC-OH-SEC-COND")
        On Error GoTo 0
    Else
        oLineString.color = 73
        oLineString.LineStyle = ActiveDesignFile.LineStyles.Find("SEC")
        On Error Resume Next
        oLineString.Level = ActiveDesignFile.Levels("CE-EX-ELEC-OH-SEC-COND")
        On Error GoTo 0
    End If
    oLineString.Class = primary
    oLineString.LineWeight = 1
    
    
    ActiveModelReference.AddElement oLineString
    If PrintOptions.ShowDrawing Then DoEvents

    Dim ofeature As Object
    Dim oFeatureMgr As Object
    Set oFeatureMgr = CreateObject("xft.FeatureMgr")

    Set ofeature = oFeatureMgr.CreateFeature(oLineString)
    ofeature.name = "CE_SEC_OH_COND"
    
    ofeature.SetProperty "TYPE", "Lighting"
    'oFeature.SetProperty "SEC_CONFIG", "TX"
    'oFeature.SetProperty "SEC_MX_SIZE", 1
    If size2 <> "" Then
        ofeature.SetProperty "LIFECYCLE", 4
    Else
        ofeature.SetProperty "LIFECYCLE", 1
    End If
    ofeature.SetProperty "StartDeadend", IIf(jsonSec.exists("startDeadend"), jsonSec("startDeadend"), False)
    ofeature.SetProperty "EndDeadend", IIf(jsonSec.exists("endDeadend"), jsonSec("endDeadend"), False)

    ofeature.Write (True)
    'oLineString.Redraw msdDrawingModeNormal
    'If PrintOptions.ShowDrawing Then DoEvents
    
    Dim sizeLabel As String, lengthLabel As String
    
    sizeLabel = size1
    If size2 <> "" Then sizeLabel = "RP " & sizeLabel & "/" & size2
    lengthLabel = jsonSec("length") & IIf(jsonSec("length") <> "", "'", "")
    
    Call PlaceTextAboveAndBelow(oLineString, sizeLabel, lengthLabel, top, bottom)
End Sub

Public Sub placePrimary(jsonPri As Object)
    Dim oLineString As LineElement
    Dim points(0 To 1) As Point3d
    Dim i As Integer
    Dim layer As Integer
    Dim size As String
    Dim size2 As String
    Dim neutSize As String
    Dim neutSize2 As String
    Dim config As String
    Dim top As Boolean
    Dim bottom As Boolean
    
    layer = jsonPri("layer")
    phase = jsonPri("phase")
    size1 = jsonPri("size")
    size2 = jsonPri("size2")
    neutSize = jsonPri("neutSize")
    neutSize2 = jsonPri("neutSize2")
    config = jsonPri("configuration")
    
    top = jsonPri("top")
    bottom = jsonPri("bottom")
    
    angle = 90 - jsonPri("angle")
    If angle < 0 Then angle = angle + 360
    If angle > 90 And angle <= 180 Then angle = 180 - angle
    If angle > 180 And angle <= 270 Then angle = angle - 180
    If angle > 270 And angle <= 360 Then angle = 360 - angle
    
    Dim calculatedAngle As Double
    Dim pt As Point3d: pt = Point3dFromXYZ(jsonPri("x1"), jsonPri("y1"), 0)
    Dim closestPt As Point3d
    results = FindClosestPointOnCenterlines(pt)
    closestPt = results(0)
    If closestPt.x <> -1 And closestPt.y <> -1 And closestPt.Z <> -1 Then
        Dim rMatrix As Matrix3d
        
        Dim directionVector As Point3d
        Dim baseCellVector As Point3d
        directionVector = Point3dSubtract(closestPt, pt)

        directionVector.Z = 0
        baseCellVector = Point3dFromXYZ(1, 0, 0)
        calculatedAngle = Point3dAngleBetweenVectors(baseCellVector, directionVector)
        If directionVector.y < 0 Then
            calculatedAngle = -calculatedAngle
        End If
    End If
    
    calculatedAngle = calculatedAngle * (180 / PI)
    If calculatedAngle < 0 Then calculatedAngle = calculatedAngle + 360
    
    x1 = jsonPri("x1")
    x2 = jsonPri("x2")
    
    y1 = jsonPri("y1")
    y2 = jsonPri("y2")
    
    Dim dx As Double, dy As Double
    dx = jsonPri("x1") - jsonPri("x2")
    dy = jsonPri("y1") - jsonPri("y2")
    
    Dim angleRad As Double
    angleRad = Atn2(dy, dx)
    
    Dim angleDeg As Double
    angleDeg = angleRad * (180 / PI)

    If angleDeg < 0 Then angleDeg = angleDeg + 360
    
    If angle > 45 Then
        If (calculatedAngle >= 135) And (calculatedAngle < 225) Then
            x1 = x1 + (PrintOptions.ConductorInitOffset + ((layer - 1) * PrintOptions.ConductorOffsetAmount))
            x2 = x2 + (PrintOptions.ConductorInitOffset + ((layer - 1) * PrintOptions.ConductorOffsetAmount))
            If angleDeg <= 95 Or angleDeg > 265 Then
                If Not top Or Not bottom Then
                    If top Then
                        top = False
                        bottom = True
                    ElseIf bottom Then
                        bottom = False
                        top = True
                    End If
                End If
            End If
        Else
            x1 = x1 - (PrintOptions.ConductorInitOffset + ((layer - 1) * PrintOptions.ConductorOffsetAmount))
            x2 = x2 - (PrintOptions.ConductorInitOffset + ((layer - 1) * PrintOptions.ConductorOffsetAmount))
        End If
    Else
        If (calculatedAngle >= 225) And (calculatedAngle < 315) Then
            y1 = y1 + (PrintOptions.ConductorInitOffset + ((layer - 1) * PrintOptions.ConductorOffsetAmount))
            y2 = y2 + (PrintOptions.ConductorInitOffset + ((layer - 1) * PrintOptions.ConductorOffsetAmount))
        Else
            y1 = y1 - (PrintOptions.ConductorInitOffset + ((layer - 1) * PrintOptions.ConductorOffsetAmount))
            y2 = y2 - (PrintOptions.ConductorInitOffset + ((layer - 1) * PrintOptions.ConductorOffsetAmount))
            If angleDeg > 95 Or angleDeg <= 265 Then
                If Not top Or Not bottom Then
                    If top Then
                        top = False
                        bottom = True
                    ElseIf bottom Then
                        bottom = False
                        top = True
                    End If
                End If
            End If
        End If
    End If
    
    points(0) = Point3dFromXYZ(x1, y1, 0)
    points(1) = Point3dFromXYZ(x2, y2, 0)
    
    Set oLineString = CreateLineElement1(Nothing, points)
    oLineString.LineStyle = ActiveDesignFile.LineStyles.Find("PRI" & phase & "P" & IIf(size2 <> "", "RP", ""))
    oLineString.Class = primary
    oLineString.LineWeight = 1
    If size2 <> "" Then
        oLineString.color = 64
        On Error Resume Next
        oLineString.Level = ActiveDesignFile.Levels("CE-RP-ELEC-OH-PRI-COND")
        On Error GoTo 0
    Else
        oLineString.color = 72
        On Error Resume Next
        oLineString.Level = ActiveDesignFile.Levels("CE-EX-ELEC-OH-PRI-COND")
        On Error GoTo 0
    End If
    ActiveModelReference.AddElement oLineString
    If PrintOptions.ShowDrawing Then DoEvents

    Dim ofeature As Object
    Dim oFeatureMgr As Object
    Set oFeatureMgr = CreateObject("xft.FeatureMgr")

    Set ofeature = oFeatureMgr.CreateFeature(oLineString)
    ofeature.name = "CE_PRIMARY_OH_CONDUCTOR"
    
    If phase = "1" Then
        ofeature.SetProperty "PHASE", "Y"
    ElseIf phase = "2" Then
        ofeature.SetProperty "PHASE", "XY"
    Else
        ofeature.SetProperty "PHASE", "3P"
    End If
    ofeature.SetProperty "PRIMARY_SIZE", size
    ofeature.SetProperty "PRIMARY_MATERIAL", 4
    ofeature.SetProperty "NEUTRAL_CONFIGURATION", config
    ofeature.SetProperty "NEUTRAL_SIZE", neutSize
    ofeature.SetProperty "NEUTRAL_MATERIAL", 1
    ofeature.SetProperty "VOLTAGE", 1
    If size2 <> "" Then
        ofeature.SetProperty "LIFECYCLE", 4
    Else
        ofeature.SetProperty "LIFECYCLE", 1
    End If
    ofeature.SetProperty "StartDeadend", IIf(jsonPri.exists("startDeadend"), jsonPri("startDeadend"), False)
    ofeature.SetProperty "EndDeadend", IIf(jsonPri.exists("endDeadend"), jsonPri("endDeadend"), False)

    ofeature.Write (True)
    'oLineString.Redraw msdDrawingModeNormal
    'If PrintOptions.ShowDrawing Then DoEvents
    
    Dim sizeLabel1 As String, sizeLabel2 As String, lengthLabel As String
    
    sizeLabel1 = size1 & IIf(config <> "", "+" & IIf(neutSize <> "" And neutSize <> size1, neutSize, "") & config, "")
    If neutSize2 <> "" Then
        sizeLabel2 = size2 & IIf(config <> "", "+" & IIf(neutSize2 <> size2, neutSize2, "") & config, "")
    Else
        sizeLabel2 = size2 & IIf(config <> "", "+" & IIf(neutSize <> "" And neutSize <> size2, neutSize, "") & config, "")
    End If
    lengthLabel = jsonPri("length") & IIf(jsonPri("length") <> "", "'", "")
    
    If size2 <> "" Then
        sizeLabel1 = "RP " & sizeLabel1 & "/" & sizeLabel2
    End If
    
    Call PlaceTextAboveAndBelow(oLineString, sizeLabel1, lengthLabel, top, bottom)
End Sub

Public Sub placeLCPObject(jsonObject As Object, cellType As String)
    Dim scl As Point3d: scl = Point3dFromXYZ(1, 1, 1)
    Dim pt As Point3d
    Dim rotationAxis As Point3d
    Dim rMatrix As Matrix3d
    Dim angleRadians As Double

    pt = Point3dFromXYZ(jsonObject("x"), jsonObject("y"), 0)
    
    rotationAxis = Point3dFromXYZ(0, 0, 1)
    
    angleRadians = jsonObject("rotation") * (PI / 180)
    
    rMatrix = Matrix3dFromVectorAndRotationAngle(rotationAxis, angleRadians)
    
    Dim fuse As CellElement
    Dim recloser As CellElement
    Dim sectionalizer As CellElement
    Dim switch As CellElement
    Dim sensor As CellElement
    Dim riser As CellElement
    Dim ofeature As Object
    Dim oFeatureMgr As Object
    Set oFeatureMgr = CreateObject("xft.FeatureMgr")
    Dim txt As TextElement
    Dim txtString As String
    
    If cellType = "Fuse" Then
        If jsonObject("open") Then
            Set fuse = CreateCellElement2("FUSEOEX", pt, scl, True, rMatrix)
        Else
            Set fuse = CreateCellElement2("FUSECEX", pt, scl, True, rMatrix)
        End If
        
        ActiveModelReference.AddElement fuse
        If PrintOptions.ShowDrawing Then DoEvents
        
        Set ofeature = oFeatureMgr.CreateFeature(fuse)
        ofeature.name = "CE_OH_FUSE"
        ofeature.SetProperty "LIFECYCLE", 1
        ofeature.SetProperty "STATE", IIf(jsonObject("open"), "Open", "Closed")
        ofeature.SetProperty "TYPE", 1
        ofeature.SetProperty "SIZE", jsonObject("size")
        If jsonObject("lcp") <> "" Then ofeature.SetProperty "LCP", jsonObject("lcp")
    
        ofeature.Write (True)
        'fuse.Redraw msdDrawingModeNormal
        'If PrintOptions.ShowDrawing Then DoEvents

        txtString = jsonObject("lcp") & vbLf & jsonObject("size") & "A"
        If jsonObject("open") Then txtString = txtString & vbLf & "OPEN"
    ElseIf cellType = "Recloser" Then
        Set recloser = CreateCellElement2("RECLEX", pt, scl, True, rMatrix)
        ActiveModelReference.AddElement recloser
        If PrintOptions.ShowDrawing Then DoEvents
        
        Set ofeature = oFeatureMgr.CreateFeature(recloser)
        ofeature.name = "CE_OH_RECLOSER"
        ofeature.SetProperty "LIFECYCLE", 1
        ofeature.SetProperty "SIZE", jsonObject("size")
        ofeature.SetProperty "ATR", IIf(jsonObject("atr"), "Yes", "No")
        If jsonObject("lcp") Then ofeature.SetProperty "LCP", jsonObject("lcp")
        
        ofeature.Write (True)
        'recloser.Redraw msdDrawingModeNormal
        'If PrintOptions.ShowDrawing Then DoEvents
        
        txtString = jsonObject("lcp") & vbLf & jsonObject("size") & "A"
        If jsonObject("atr") Then txtString = txtString & vbLf & "ATR"
    ElseIf cellType = "Sectionalizer" Then
        Set sectionalizer = CreateCellElement2("SECTEX", pt, scl, True, rMatrix)
        ActiveModelReference.AddElement sectionalizer
        If PrintOptions.ShowDrawing Then DoEvents
        
        Set ofeature = oFeatureMgr.CreateFeature(sectionalizer)
        ofeature.name = "CE_OH_SECTIONALIZER"
        ofeature.SetProperty "LIFECYCLE", 1
        ofeature.SetProperty "SIZE", jsonObject("size")
        If jsonObject("lcp") Then ofeature.SetProperty "LCP", jsonObject("lcp")
        
        ofeature.Write (True)
        'recloser.Redraw msdDrawingModeNormal
        'If PrintOptions.ShowDrawing Then DoEvents
        
        txtString = jsonObject("lcp") & vbLf & jsonObject("size") & "A"
    ElseIf cellType = "Switch" Then
        If jsonObject("open") Then
            Set switch = CreateCellElement2("SWITCHLOEX", pt, scl, True, rMatrix)
        Else
            Set switch = CreateCellElement2("SWITCHLCEX", pt, scl, True, rMatrix)
        End If
        
        ActiveModelReference.AddElement switch
        If PrintOptions.ShowDrawing Then DoEvents
        
        Set ofeature = oFeatureMgr.CreateFeature(switch)
        ofeature.name = "CE_OH_SWITCH"
        'oFeature.SetProperty "LIFECYCLE", 1
        'oFeature.SetProperty "STATE", IIf(jsonObject("open"), "Open", "Closed")
        'If jsonObject("lcp") <> "" Then oFeature.SetProperty "LCP", jsonObject("lcp")
        
        ofeature.Write (True)
        'switch.Redraw msdDrawingModeNormal
        'If PrintOptions.ShowDrawing Then DoEvents
        
        txtString = jsonObject("lcp") & vbLf & jsonObject("size") & "A"
        If jsonObject("open") Then txtString = txtString & vbLf & "OPEN"
    ElseIf cellType = "Sensor" Then
        If jsonObject("power") Then
            Set sensor = CreateCellElement2("PSNSREX", pt, scl, True, rMatrix)
        Else
            Set sensor = CreateCellElement2("ISNSREX", pt, scl, True, rMatrix)
        End If
        
        ActiveModelReference.AddElement sensor
        If PrintOptions.ShowDrawing Then DoEvents
        
        Set ofeature = oFeatureMgr.CreateFeature(sensor)
        ofeature.name = "CE_SENSOR"
        ofeature.SetProperty "LIFECYCLE", 1
        ofeature.SetProperty "TYPE", IIf(jsonObject("power"), "PS", "IS")
        If jsonObject("lcp") Then ofeature.SetProperty "LCP", jsonObject("lcp")
        
        ofeature.Write (True)
        'sensor.Redraw msdDrawingModeNormal
        'If PrintOptions.ShowDrawing Then DoEvents
        
        txtString = ""
    ElseIf cellType = "Riser" Then
        If jsonObject("type") = "Secondary" Then
            Set riser = CreateCellElement2("RISSECEX", pt, scl, True, rMatrix)
        ElseIf jsonObject("type") = "Primary" Then
            Set riser = CreateCellElement2("RISPRIEX", pt, scl, True, rMatrix)
        End If
        
        ActiveModelReference.AddElement riser
        If PrintOptions.ShowDrawing Then DoEvents
        
        If jsonObject("type") = "Primary" Then txtString = jsonObject("lcp") & vbLf & jsonObject("size") & "A"
    End If

    If txtString <> "" Then
        Dim txtOffset As Point3d: txtOffset = Point3dFromXYZ(-20, -20, 0)
        Set txt = CreateTextElement1(Nothing, txtString, Point3dAdd(pt, txtOffset), Matrix3dIdentity)
        txt.color = 72
        txt.TextStyle.BorderAndBackgroundVisible = True
        txt.TextStyle.BackgroundFillColor = 255
        txt.TextStyle.BorderColor = 255
        ActiveModelReference.AddElement txt
        If PrintOptions.ShowDrawing Then DoEvents
    End If
End Sub

Function getNormalAngle(ByVal angle As Double) As Double
    If (angle > 0 And angle <= 180) Then
        angle = angle - 90
    Else
        angle = angle + 90
    End If
    If angle < 0 Then aangle = angle + 360
    If angle >= 360 Then angle = angle - 360
    
    getNormalAngle = angle
End Function

Sub PlaceTextAboveAndBelow(oLine As LineElement, text1String As String, text2String As String, top As Boolean, bottom As Boolean)
    Dim startPt As Point3d
    Dim endPt As Point3d
    
    startPt = oLine.startPoint
    endPt = oLine.endPoint
    
    Dim dx As Double, dy As Double
    dx = endPt.x - startPt.x
    dy = endPt.y - startPt.y
    
    Dim angleRad As Double
    angleRad = Atn2(dy, dx)
    
    Dim angleDeg As Double
    angleDeg = angleRad * (180 / PI)

    If angleDeg < 0 Then angleDeg = angleDeg + 360

    If angleDeg > 85 And angleDeg <= 275 Then
        If angleDeg >= 95 Then
            angleRad = angleRad - PI
        End If
    End If

    If Not bottom Then text2String = ""

    Dim rotMatrix As Matrix3d
    Dim rotationAxis As Point3d
    rotationAxis = Point3dFromXYZ(0, 0, 1)
    rotMatrix = Matrix3dFromVectorAndRotationAngle(rotationAxis, angleRad) ' Rotate around Z-axis
    
    Dim centerPt As Point3d
    centerPt.x = (startPt.x + endPt.x) / 2
    centerPt.y = (startPt.y + endPt.y) / 2
    centerPt.Z = (startPt.Z + endPt.Z) / 2

    Dim textOffsetDistance As Double
    textOffsetDistance = 10
    
    Dim offsetX As Double, offsetY As Double
    offsetX = -Sin(angleRad) * textOffsetDistance
    offsetY = Cos(angleRad) * textOffsetDistance
    
    Dim ptText1 As Point3d, ptText2 As Point3d

    ' Text 1 is on top / left
    ptText1.x = centerPt.x + offsetX
    ptText1.y = centerPt.y + offsetY
    ptText1.Z = centerPt.Z
    
    ' Text 2 is on bottom / right
    ptText2.x = centerPt.x - offsetX
    ptText2.y = centerPt.y - offsetY
    ptText2.Z = centerPt.Z

    If text1String <> "" Then
        Dim oText1 As TextElement
        Set oText1 = CreateTextElement1(Nothing, text1String, ptText1, rotMatrix)
        oText1.TextStyle.Justification = msdTextJustificationCenterCenter
        oText1.TextStyle.BorderAndBackgroundVisible = True
        oText1.TextStyle.BackgroundFillColor = 255
        oText1.TextStyle.BorderColor = 255
        ActiveModelReference.AddElement oText1
        If PrintOptions.ShowDrawing Then DoEvents
        oText1.TextStyle.color = oLine.color
        oText1.Rewrite
        If PrintOptions.ShowDrawing Then DoEvents
    End If
        
    If text2String <> "" Then
        Dim oText2 As TextElement
        Set oText2 = CreateTextElement1(Nothing, text2String, ptText2, rotMatrix)
        oText2.TextStyle.Justification = msdTextJustificationCenterCenter
        oText2.TextStyle.BorderAndBackgroundVisible = True
        oText2.TextStyle.BackgroundFillColor = 255
        oText2.TextStyle.BorderColor = 255
        ActiveModelReference.AddElement oText2
        If PrintOptions.ShowDrawing Then DoEvents
        oText2.TextStyle.color = oLine.color
        oText2.Rewrite
        If PrintOptions.ShowDrawing Then DoEvents
    End If
End Sub

Private Function Atn2(dy As Double, dx As Double) As Double
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

Sub CreateLinearDimensionFullyCorrected(ptStart As Point3d, ptEnd As Point3d, rMatrix As Matrix3d)
    Dim oDimStyle As DimensionStyle
    Dim oDim As DimensionElement
    Dim directionVector As Point3d
    Dim distance As Double
    Dim oNewTextStyle As TextStyle
    
    Set oDim = CreateDimensionElement1(Nothing, rMatrix, msdDimTypeCustomLinear)
    
    Set oDimStyle = ActiveSettings.DimensionStyle
    oDimStyle.ExtensionLineExtend = False
    oDimStyle.ExtensionLineOffset = 0
    oDimStyle.textColor = 6
    oDimStyle.TerminatorColor = 6
    oDimStyle.TextFrameType = MsdDimTextFrameTypeNone
    oDimStyle.OverallColor = 6
    
    oDimStyle.ExtensionLineColor = 6
    oDimStyle.OverrideExtensionLineColor = True
    Set oDim.DimensionStyle = oDimStyle
    oDim.color = 6
    
    distance = 0
    directionVector.x = rMatrix.RowY.x
    directionVector.y = rMatrix.RowY.y
    directionVector.Z = rMatrix.RowY.Z

    directionVector = Point3dNormalize(directionVector)
    
    Dim shiftVector As Point3d
    shiftVector.x = directionVector.x * distance
    shiftVector.y = directionVector.y * distance
    shiftVector.Z = directionVector.Z * distance

    ptStart = Point3dAdd(ptStart, shiftVector)
    ptEnd = Point3dAdd(ptEnd, shiftVector)
    
    oDim.InsertReferencePoint ActiveModelReference, 1, ptEnd
    oDim.InsertReferencePoint ActiveModelReference, 2, ptStart
    
    ActiveModelReference.AddElement oDim
    If PrintOptions.ShowDrawing Then DoEvents
End Sub

Function FindClosestPointOnCenterlines(startPoint As Point3d) As Variant()
    Dim oScanCriteria As New ElementScanCriteria
    Dim oEnumerator As ElementEnumerator
    Dim oEl As element
    Dim rMatrix As Matrix3d
    Dim closestPointOnAnyLine As Point3d
    Dim shortestDistance As Double
    Dim currentDistance As Double
    Dim tempPoint As Point3d
    Dim hasFoundAny As Boolean
    
    oScanCriteria.ExcludeAllTypes
    oScanCriteria.IncludeType msdElementTypeLine
    oScanCriteria.IncludeType msdElementTypeLineString
    
    Set oEnumerator = ActiveModelReference.Scan(oScanCriteria)
    
    shortestDistance = 999999999#
    hasFoundAny = False
    
    Do While oEnumerator.MoveNext
        Set oEl = oEnumerator.Current
        
        If oEl.color = 6 And oEl.LineStyle.Number = 4 Then
            Dim tagStr As String
            tagStr = ""
            
            Dim oLine As LineElement
            Set oLine = oEl
    
            arrVertices = oLine.GetVertices
            For i = LBound(arrVertices) To UBound(arrVertices) - 1
                Dim linePoint1 As Point3d: linePoint1 = arrVertices(i)
                Dim linePoint2 As Point3d: linePoint2 = arrVertices(i + 1)
                tempPoint = ClosestPointOnSegment(linePoint1, linePoint2, startPoint)
                currentDistance = Point3dDistance(startPoint, tempPoint)
                If currentDistance < shortestDistance Then
                    shortestDistance = currentDistance
                    closestPointOnAnyLine = tempPoint
                    
                    Dim directionVec As Vector3d
                    Dim lineAngle As Double
                    
                    directionVec = Vector3dFromXY(linePoint2.x - linePoint1.x, linePoint2.y - linePoint1.y)
                    lineAngle = Vector3dPolarAngle(directionVec) + (PI / 2)
                    rMatrix = Matrix3dFromAxisAndRotationAngle(2, lineAngle)
                    
                    hasFoundAny = True
                End If
            Next i
        End If
    Loop
    
    If hasFoundAny Then
        FindClosestPointOnCenterlines = Array(closestPointOnAnyLine, rMatrix)
    Else
        FindClosestPointOnCenterlines = Array(Point3dFromXYZ(-1, -1, -1))
    End If
End Function

Function ClosestPointOnSegment(A As Point3d, b As Point3d, p As Point3d) As Point3d
    Dim Ax As Double: Ax = A.x
    Dim Ay As Double: Ay = A.y
    Dim Bx As Double: Bx = b.x
    Dim By As Double: By = b.y
    Dim Px As Double: Px = p.x
    Dim Py As Double: Py = p.y
    
    Dim ABx As Double, ABy As Double
    Dim APx As Double, APy As Double
    Dim dot_AP_AB As Double
    Dim dot_AB_AB As Double
    Dim t As Double
    Dim result(1 To 2) As Double

    ABx = Bx - Ax
    ABy = By - Ay
    APx = Px - Ax
    APy = Py - Ay
    
    dot_AP_AB = (APx * ABx) + (APy * ABy)
    dot_AB_AB = (ABx * ABx) + (ABy * ABy)
    
    If dot_AB_AB = 0 Then
        result(1) = Ax
        result(2) = Ay
        ClosestPointOnSegment = Point3dFromXYZ(Ax, Ay, 0)
        Exit Function
    End If
    
    t = dot_AP_AB / dot_AB_AB
    
    If t < 0 Then
        t = 0
    ElseIf t > 1 Then
        t = 1
    End If
    
    ClosestPointOnSegment = Point3dFromXYZ(Ax + (t * ABx), Ay + (t * ABy), 0)
End Function
