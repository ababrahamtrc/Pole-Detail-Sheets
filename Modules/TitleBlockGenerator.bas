Attribute VB_Name = "TitleBlockGenerator"
Sub AutofillTitleBlocks()
    Dim json As Object: Set json = ReadJSON("title_block")
    
    If json Is Nothing Then
        MsgBox "JSON file doesn't exist, generate one from pole detail sheets."
        Exit Sub
    End If

    ActiveDesignFile.DefaultModelReference.Activate

    Dim printTypes As Object: Set printTypes = CreateObject("Scripting.Dictionary")
    Dim modelPrintType As Object: Set modelPrintType = CreateObject("Scripting.Dictionary")
    For Each mdl In ActiveDesignFile.Models
        If mdl.Type = msdModelTypeSheet Then
            printType = ChangeSheetText(mdl, json)
            If Not printTypes.exists(printType) Then printTypes.Add printType, 0
            printTypes(printType) = printTypes(printType) + 1
            modelPrintType.Add mdl.name, Array(printType, printTypes(printType))
        End If
    Next mdl
    
    Dim printNumber As Integer
    Dim totalPrintType As Integer
    For Each mdl In ActiveDesignFile.Models
        If modelPrintType.exists(mdl.name) Then
            printType = modelPrintType(mdl.name)(0)
            printNumber = modelPrintType(mdl.name)(1)
            totalPrintType = printTypes(printType)
            Call updateSheetNumbers(mdl, printNumber, totalPrintType)
        End If
    Next mdl
    
    MsgBox "Finished Autofilling Title Blocks"
End Sub


Function ChangeSheetText(ByVal mdl As ModelReference, json As Object) As String
    Dim sc As New ElementScanCriteria
    Dim ee As ElementEnumerator
    Dim EL As Element
    Dim cell As CellElement
    
    sc.ExcludeAllLevels
    sc.IncludeLevel ActiveDesignFile.Levels("Title Block")

    Set titleBlock = GetTitleBlock(mdl)
    
    If Not titleBlock Is Nothing Then
        sheettype = GetSheetType(titleBlock) & " " & GetOrientation(titleBlock)
    Else
        sheettype = ""
    End If
    
    Dim printType As String
    
    If sheettype <> "" Then
        Set ee = mdl.Scan(sc)
        
        Dim treeWork As Boolean
        Dim locations As Collection: Set locations = New Collection
        Set cells = getVisibleCells(mdl)
        For Each cell In cells
            If InStr(1, cell.name, "tree", vbTextCompare) > 0 Then treeWork = True
            If InStr(1, cell.name, "PULLPOLE", vbTextCompare) > 0 Then printType = "Pull"
            If cell.name = "LOCATION_NUMBER" Then
                Set oEnum = cell.GetSubElements
                Do While oEnum.MoveNext
                    Set oSubEl = oEnum.Current
                     If oSubEl.IsTextElement Then
                        Set oTextEl = oSubEl.AsTextElement
                        If oTextEl.text <> "" Then locations.Add oTextEl.text
                    End If
                Loop
            End If
        Next cell

        If printType = "" Then
            If locations.count = 0 Then
                printType = "NJUNS"
            Else
                printType = "Normal"
            End If
        End If

        Dim errorMsg As String

        If json("DesignerName") <> "" Then
            FindAndReplaceText ee, "DESIGNER", json("DesignerName"), 4
        Else
            errorMsg = errorMsg & "Missing 'Designer Name' on JSON, not filled in on title block." & vbLf
        End If
        
        If json("DesignerNumber") <> "" Then
            FindAndReplaceText ee, "DESIGNER", json("DesignerNumber"), 3
        Else
            errorMsg = errorMsg & "Missing 'Designer Number' on JSON, not filled in on title block." & vbLf
        End If
        
        If json("CoordinaterName") <> "" Then
            FindAndReplaceText ee, "DESIGNER", json("CoordinaterName"), 5
        Else
            errorMsg = errorMsg & "Missing 'Coordinator Name' on JSON, not filled in on title block." & vbLf
        End If
        
        If json("CoordinaterNumber") <> "" Then
            FindAndReplaceText ee, "DESIGNER", json("CoordinaterNumber"), 2
        Else
            errorMsg = errorMsg & "Missing 'Coordinator Number' on JSON, not filled in on title block." & vbLf
        End If
        
        If json("County") <> "" Then
            FindAndReplaceText ee, "CO", json("County"), 2
        Else
            errorMsg = errorMsg & "Missing 'County' on JSON, not filled in on title block." & vbLf
        End If
        
        If json("TWP") <> "" Then
            FindAndReplaceText ee, "TWP", json("TWP"), 2
        Else
            errorMsg = errorMsg & "Missing 'TWP' on JSON, not filled in on title block." & vbLf
        End If
        
        
        lines = Split(json("JobPurpose"), vbCrLf)
        ArrLen = UBound(lines) - LBound(lines) + 1
        If ArrLen > 0 Then
            FindAndReplaceText ee, "ENTER JOB PURPOSE", lines(0), 0
        Else
            errorMsg = errorMsg & "Missing 'Job Purpose' on JSON, not filled in on title block." & vbLf
        End If
        If ArrLen > 1 Then
            FindAndReplaceText ee, "INTO THIS", lines(1), 0
        ElseIf ArrLen > 0 Then
            FindAndReplaceText ee, "INTO THIS", "", 0
        End If
        If ArrLen > 2 Then
            FindAndReplaceText ee, "TEXT FIELD", lines(2), 0
        ElseIf ArrLen > 0 Then
            FindAndReplaceText ee, "TEXT FIELD", "", 0
        End If
        
        If treeWork And printType <> "Pull" Then
            FindAndReplaceText ee, "FORESTRY REQ'D", "X", 4
            FindAndReplaceText ee, "FORESTRY REQ'D", " ", 6
        Else
            FindAndReplaceText ee, "FORESTRY REQ'D", "X", 6
            FindAndReplaceText ee, "FORESTRY REQ'D", " ", 4
        End If
        
        Dim staking As Boolean
        For Each location In locations
            If json("Staking").exists(location) Then staking = True: Exit For
        Next location
        
        If staking And printType <> "Pull" Then
            FindAndReplaceText ee, "CE STAKING REQ'D", "X", -1
            FindAndReplaceText ee, "CE STAKING REQ'D", " ", 3
        Else
            FindAndReplaceText ee, "CE STAKING REQ'D", "X", 3
            FindAndReplaceText ee, "CE STAKING REQ'D", " ", -1
        End If


        If printType <> "NJUNS" Then
            Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
            
            Dim substation As String
            Dim substationNumber As String
            For Each location In locations
                If json("Substation").exists(CStr(location)) Then
                    substation = Replace(json("Substation")(location), " -", "-")
                    If Not seen.exists(substation) Then seen.Add substation, True
                    If InStr(substation, "-") > 0 Then
                        substationNumber = Mid(substation, InStr(substation, "-") + 2, Len(substation) - InStr(substation, "-") - 1)
                        substation = Left(substation, InStr(substation, "-") - 1)
                    End If
                End If
            Next location
            
            If seen.count > 1 Then
                substation = "VARIES, SEE PRINT"
                substationNumber = ""
            End If
            
            If substation <> "" Then
                FindAndReplaceText ee, "SUBSTATION", substation, 1
            ElseIf locations.count > 0 Then
                errorMsg = errorMsg & "Missing 'Substation' on JSON, not filled in on title block." & vbLf
            End If
            
            If substationNumber <> "" Then
                FindAndReplaceText ee, "WD NO.", substationNumber, 1
            ElseIf substation <> "VARIES, SEE PRINT" And locations.count > 0 Then
                errorMsg = errorMsg & "Missing 'Substation Number' on JSON, not filled in on title block." & vbLf
            End If
            
            Set seen = CreateObject("Scripting.Dictionary")
            
            Dim circuit As String
            Dim circuitNumber As String
            For Each location In locations
                If json("Circuit").exists(CStr(location)) Then
                    circuit = Replace(json("Circuit")(location), " - ", "-")
                    If Not seen.exists(circuit) Then seen.Add circuit, True
                    If InStr(circuit, "-") > 0 Then
                        circuitNumber = Mid(circuit, InStr(circuit, "-") + 1, Len(circuit) - InStr(circuit, "-"))
                        circuit = Left(circuit, InStr(circuit, "-") - 1)
                    End If
                End If
            Next location
            
            If seen.count > 1 Then
                circuit = "VARIES, SEE PRINT"
                circuitNumber = ""
            End If
            
            If circuit <> "" Then
                FindAndReplaceText ee, "CIRCUIT", circuit, 1
            ElseIf locations.count > 0 Then
                errorMsg = errorMsg & "Missing 'Circuit' on JSON, not filled in on title block." & vbLf
            End If
            
            If circuitNumber <> "" Then
                If sheettype = "A Portrait" Then
                    FindAndReplaceText ee, "CKT NO.", circuitNumber, 1
                Else
                    FindAndReplaceText ee, "WD NO.", circuitNumber, 2
                End If
            ElseIf circuit <> "VARIES, SEE PRINT" And locations.count > 0 Then
                errorMsg = errorMsg & "Missing 'Circuit Number' on JSON, not filled in on title block." & vbLf
            End If
            
            Set seen = CreateObject("Scripting.Dictionary")
            
            Dim upstreamDeviceLocation As String
            For Each location In locations
                If json("UpstreamDevice").exists(CStr(location)) Then
                    upstreamDeviceLocation = Replace(json("UpstreamDevice")(location), "\n", vbCrLf)
                    If Not seen.exists(upstreamDeviceLocation) Then seen.Add upstreamDeviceLocation, True
                End If
            Next location
            If seen.count > 1 Then
                upstreamDeviceLocation = "VARIES, SEE PRINT" & vbCrLf & "VARIES, SEE PRINT"
            End If
    
            lines = Split(upstreamDeviceLocation, vbCrLf)
            ArrLen = UBound(lines) - LBound(lines) + 1
            If ArrLen > 0 Then
                FindAndReplaceText ee, "XXXX-##A", lines(0), 0
            ElseIf locations.count > 0 Then
                errorMsg = errorMsg & "Missing 'Upstream Device' on JSON, not filled in on title block." & vbLf
            End If
            If ArrLen > 1 Then
                FindAndReplaceText ee, "ENTER LOCATION INTO", lines(1), 0
            ElseIf ArrLen > 0 Then
                FindAndReplaceText ee, "ENTER LOCATION INTO", " ", 0
            End If
            If ArrLen > 2 Then
                FindAndReplaceText ee, "THIS TEXT FIELD", lines(2), 0
            ElseIf ArrLen > 0 Then
                FindAndReplaceText ee, "THIS TEXT FIELD", " ", 0
            End If
        End If
    End If
    
    Dim notificationOffsets As Object: Set notificationOffsets = CreateObject("Scripting.Dictionary")
    Set notificationOffsets(1) = CreateObject("Scripting.Dictionary")
    Set notificationOffsets(2) = CreateObject("Scripting.Dictionary")
    Set notificationOffsets(3) = CreateObject("Scripting.Dictionary")
    
    If sheettype = "A Landscape" Then
        notificationOffsets(1)(1) = 7
        notificationOffsets(1)(2) = 8
        notificationOffsets(1)(3) = 9
        notificationOffsets(1)(4) = 10
        
        notificationOffsets(2)(1) = 12
        notificationOffsets(2)(2) = 13
        notificationOffsets(2)(3) = 14
        notificationOffsets(2)(4) = 15
    
        notificationOffsets(3)(1) = 19
        notificationOffsets(3)(2) = 20
        notificationOffsets(3)(3) = 21
        notificationOffsets(3)(4) = 22
    ElseIf sheettype = "A Portrait" Then
        notificationOffsets(1)(1) = 5
        notificationOffsets(1)(2) = 6
        notificationOffsets(1)(3) = 7
        notificationOffsets(1)(4) = 8
        
        notificationOffsets(2)(1) = 10
        notificationOffsets(2)(2) = 11
        notificationOffsets(2)(3) = 12
        notificationOffsets(2)(4) = 13
    ElseIf sheettype = "B Landscape" Then
        notificationOffsets(1)(1) = 6
        notificationOffsets(1)(2) = 7
        notificationOffsets(1)(3) = 8
        notificationOffsets(1)(4) = 23
        
        notificationOffsets(2)(1) = 25
        notificationOffsets(2)(2) = 26
        notificationOffsets(2)(3) = 27
        notificationOffsets(2)(4) = 28
    ElseIf sheettype = "B Portrait" Then
        notificationOffsets(1)(1) = 7
        notificationOffsets(1)(2) = 8
        notificationOffsets(1)(3) = 9
        notificationOffsets(1)(4) = 10
        
        notificationOffsets(2)(1) = 12
        notificationOffsets(2)(2) = 13
        notificationOffsets(2)(3) = 14
        notificationOffsets(2)(4) = 15
    
        notificationOffsets(3)(1) = 20
        notificationOffsets(3)(2) = 21
        notificationOffsets(3)(3) = 22
        notificationOffsets(3)(4) = 23
    Else
        notificationOffsets(1)(1) = 6
        notificationOffsets(1)(2) = 7
        notificationOffsets(1)(3) = 8
        notificationOffsets(1)(4) = 9
        
        notificationOffsets(2)(1) = 11
        notificationOffsets(2)(2) = 12
        notificationOffsets(2)(3) = 13
        notificationOffsets(2)(4) = 14
    
        notificationOffsets(3)(1) = 20
        notificationOffsets(3)(2) = 21
        notificationOffsets(3)(3) = 22
        notificationOffsets(3)(4) = 23
    End If
    
    maxOffset = 3
    If sheettype = "A Portrait" Or sheettype = "B Landscape" Then maxOffset = 2
    
    If printType <> "NJUNS" Then
        If json("Mode") = "SYSTEM IMPROVEMENT" Then
            Dim offSet As Integer: offSet = 1
            If json("DemoNotification") <> "" Then
                If printType = "Pull" Then
                    FindAndReplaceText ee, "MAT.", "ESIC", notificationOffsets(offSet)(1)
                    FindAndReplaceText ee, "MAT.", "OHL", notificationOffsets(offSet)(2)
                    FindAndReplaceText ee, "MAT.", json("MainNotification"), notificationOffsets(offSet)(3)
                    FindAndReplaceText ee, "MAT.", json("MainDesDoc"), notificationOffsets(offSet)(4)
                ElseIf printType = "Normal" Then
                    FindAndReplaceText ee, "MAT.", "ERET", notificationOffsets(offSet)(1)
                    FindAndReplaceText ee, "MAT.", "OHL", notificationOffsets(offSet)(2)
                    FindAndReplaceText ee, "MAT.", json("DemoNotification"), notificationOffsets(offSet)(3)
                    FindAndReplaceText ee, "MAT.", json("DemoDesDoc"), notificationOffsets(offSet)(4)
                End If
                offSet = offSet + 1
            End If
            For i = 1 To json("ServiceNotifications").count
                If offSet > maxOffset Then Exit For
                FindAndReplaceText ee, "MAT.", "OHS", notificationOffsets(offSet)(1)
                FindAndReplaceText ee, "MAT.", "OHS", notificationOffsets(offSet)(2)
                FindAndReplaceText ee, "MAT.", json("ServiceNotifications")(i), notificationOffsets(offSet)(3)
                If json("ServiceDesDocs").count >= i Then FindAndReplaceText ee, "MAT.", json("ServiceDesDocs")(i), notificationOffsets(offSet)(4)
                
                offSet = offSet + 1
            Next i
        End If
    End If
    
    If errorMsg <> "" Then
        MsgBox mdl.name & vbLf & errorMsg
    End If
    
    ChangeSheetText = printType
End Function

Sub updateSheetNumbers(ByVal mdl As ModelReference, num1 As Integer, num2 As Integer)
    Dim sc As New ElementScanCriteria
    Dim ee As ElementEnumerator
    
    sc.ExcludeAllLevels
    sc.IncludeLevel ActiveDesignFile.Levels("Title Block")
    
    Dim printType As String
    
    Set ee = mdl.Scan(sc)

    FindAndReplaceText ee, "SCALE", "SHEET " & num1 & " OF " & num2, 1
End Sub

Function GetTitleBlock(ByVal mdl As ModelReference) As CellElement

    Dim sc As New ElementScanCriteria
    Dim ee As ElementEnumerator
    Dim EL As Element

    sc.ExcludeAllLevels
    sc.IncludeLevel ActiveDesignFile.Levels("Title Block")

    Set ee = mdl.Scan(sc)

    Do While ee.MoveNext
        Set EL = ee.Current

        If EL.Type = msdElementTypeCellHeader Then
            Set GetTitleBlock = EL.AsCellElement
            Exit Function
        End If

    Loop

    Set GetTitleBlock = Nothing

End Function

Function GetSheetType(ByVal cell As CellElement) As String

    Dim ee As ElementEnumerator
    Dim EL As Element
    Dim txt As String
    Dim result As String

    Set ee = cell.GetSubElements

    Do While ee.MoveNext

        Set EL = ee.Current

        If EL.Type = msdElementTypeText Then

            txt = UCase$(Trim$(EL.AsTextElement.text))

            Select Case txt

                Case "SHEET A"
                    GetSheetType = "A"
                    Exit Function

                Case "SHEET B"
                    GetSheetType = "B"
                    Exit Function

                Case "SHEET C"
                    GetSheetType = "C"
                    Exit Function
                    
                Case "SHEET D"
                    GetSheetType = "C"
                    Exit Function
            End Select

        ElseIf EL.Type = msdElementTypeCellHeader Then

            result = GetSheetType(EL.AsCellElement)

            If result <> "" Then
                GetSheetType = result
                Exit Function
            End If

        End If

    Loop

End Function

Function GetOrientation(ByVal cell As CellElement) As String

    Dim rng As Range3d
    Dim width As Double
    Dim height As Double

    rng = cell.Range

    width = rng.High.x - rng.Low.x
    height = rng.High.y - rng.Low.y

    If width > height Then
        GetOrientation = "Landscape"
    ElseIf height > width Then
        GetOrientation = "Portrait"
    Else
        GetOrientation = "Square"
    End If

End Function

Private Sub FindAndReplaceText(ByVal ee As ElementEnumerator, _
                               ByVal oldText As String, _
                               ByVal newText As String, _
                               ByVal offSet As Integer)

    Dim texts As New Collection
    Dim i As Long
    Dim targetIndex As Long
    Dim txt As TextElement

    CollectTextElements ee, texts

    For i = 1 To texts.count
        Set txt = texts(i)

        If StrComp(Trim$(txt.text), Trim$(oldText), vbTextCompare) = 0 Then

            found = True
            targetIndex = i + offSet

            If targetIndex >= 1 And targetIndex <= texts.count Then
                Set txt = texts(targetIndex)

                txt.text = newText
                txt.Rewrite

                Debug.Print "Changed to: " & newText
            End If
            Exit Sub
        End If
    Next i

End Sub

Private Sub CollectTextElements(ByVal ee As ElementEnumerator, _
                                ByRef texts As Collection)

    Dim EL As Element
    Dim childEE As ElementEnumerator

    ee.reset

    Do While ee.MoveNext

        Set EL = ee.Current

        Select Case EL.Type

            Case msdElementTypeText
                texts.Add EL.AsTextElement

            Case msdElementTypeCellHeader
                Set childEE = EL.AsCellElement.GetSubElements
                CollectTextElements childEE, texts

            Case msdElementTypeTextNode
                Set childEE = EL.AsTextNodeElement.GetSubElements
                CollectTextElements childEE, texts

        End Select

    Loop

End Sub

Private Function getVisibleCells(ByVal mdl As ModelReference) As Collection
    Dim sheetRng As Range3d, refRng As Range3d
    Dim xf As Transform3d, inv As Transform3d
    Dim att As Attachment
    Dim sc As ElementScanCriteria
    Dim ee As ElementEnumerator
    Dim cel As CellElement
    Dim p As Point3d

    Dim cells As Collection
    Set cells = New Collection

    sheetRng = mdl.Range(False)

    For Each att In mdl.Attachments
        xf = att.GetReferenceToMasterTransform
        inv = Transform3dInverse(xf)


        refRng = TransformRange(inv, sheetRng)
        refRng = GetBoundaryRange(att, refRng)
        refRng.High.Z = 1E+20
        refRng.Low.Z = 1E-20

        Set sc = New ElementScanCriteria
        sc.ExcludeAllTypes
        sc.IncludeType msdElementTypeCellHeader
        sc.IncludeOnlyWithinRange refRng

        Set ee = att.Scan(sc)
        Do While ee.MoveNext
            Set cel = ee.Current.AsCellElement
            Dim cr As Range3d: cr = cel.Range
            cells.Add cel
        Loop
    Next att
    
    Set getVisibleCells = cells
End Function

Private Function TransformRange(xf As Transform3d, r As Range3d) As Range3d
    Dim c(3) As Point3d, i As Integer, out As Range3d
    c(0) = Point3dFromXYZ(r.Low.x, r.Low.y, r.Low.Z)
    c(1) = Point3dFromXYZ(r.High.x, r.Low.y, r.Low.Z)
    c(2) = Point3dFromXYZ(r.High.x, r.High.y, r.High.Z)
    c(3) = Point3dFromXYZ(r.Low.x, r.High.y, r.High.Z)
    out = Range3dFromPoint3d(Point3dFromTransform3dTimesPoint3d(xf, c(0)))
    For i = 1 To 3
        out = Range3dUnionPoint3d(out, Point3dFromTransform3dTimesPoint3d(xf, c(i)))
    Next i

    out.Low.Z = -1E+20: out.High.Z = 1E+20
    TransformRange = out
End Function

Private Function GetBoundaryRange(att As Attachment, approx As Range3d) As Range3d
    Dim sc As New ElementScanCriteria
    Dim ee As ElementEnumerator
    Dim r As Range3d, bndRng As Range3d
    Dim area As Double, best As Double

    bndRng.Low.x = 1E+20: bndRng.Low.y = 1E+20
    bndRng.High.x = -1E+20: bndRng.High.y = -1E+20

    sc.ExcludeAllTypes
    sc.IncludeType msdElementTypeShape

    Set ee = ActiveModelReference.Scan(sc)
    Do While ee.MoveNext
        r = ee.Current.Range
        If r.Low.x >= approx.Low.x And r.High.x <= approx.High.x And _
           r.Low.y >= approx.Low.y And r.High.y <= approx.High.y Then
            area = (r.High.x - r.Low.x) * (r.High.y - r.Low.y)
            If area > best Then
                best = area
                bndRng = r
            End If
        End If
    Loop

    GetBoundaryRange = bndRng
End Function

Private Function IsEmptyRange(r As Range3d) As Boolean
    IsEmptyRange = (r.Low.x > r.High.x Or r.Low.y > r.High.y)
End Function




