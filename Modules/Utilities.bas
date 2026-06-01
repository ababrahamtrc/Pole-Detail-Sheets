Attribute VB_Name = "Utilities"
Public Function SheetExists(ByVal sheetName As String, Optional wb As Workbook) As Boolean
    Dim sheet As Worksheet
    If wb Is Nothing Then Set wb = ThisWorkbook
 
    For Each sheet In wb.Worksheets
        If ThisWorkbook.RemoveParentheses(sheet.name) = ThisWorkbook.RemoveParentheses(sheetName) Then
            SheetExists = True
            Exit Function
        End If
    Next sheet
 
    SheetExists = False
End Function

Public Sub CheckAndCloseWorkbook(filePath As String)
    If IsWorkbookOpen(filePath) Then
        Workbooks.Open(filePath).Close SaveChanges:=False
    End If
End Sub

Private Function IsWorkbookOpen(ByVal filePath As String) As Boolean
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks.Open(filePath)
    IsWorkbookOpen = Not wb Is Nothing
    On Error GoTo 0
End Function

Public Function isCEID(text As String) As Boolean
    isCEID = False
    If IsNumeric(text) Then
        If Len(text) = 7 Then isCEID = True
    End If
End Function

Public Function OnlyNumbers(text As String, Optional slashes As Boolean = False) As String
    Dim result As String
    Dim i As Long
    For i = 1 To Len(text)
        If slashes Then
            If Mid(text, i, 1) Like "[0-9/]" Then result = result & Mid(text, i, 1)
        Else
            If Mid(text, i, 1) Like "[0-9]" Then result = result & Mid(text, i, 1)
        End If
    Next i
    
    If Not slashes And Not IsNumeric(result) Then result = "-1"
    If slashes And Not IsNumeric(Replace(result, "/", "")) Then result = "-1"
    
    OnlyNumbers = result
End Function

Public Function OnlyLetters(text As String) As String
    Dim result As String
    Dim i As Long
    For i = 1 To Len(text)
        If Mid(text, i, 1) Like "[a-zA-Z]" Then result = result & Mid(text, i, 1)
    Next i
    
    OnlyLetters = result
End Function

Public Function inchesToFeetInches(ByVal inches As Double) As String
    Dim feet As Integer

    If inches = 9999 Then
        inchesToFeetInches = "N/A"
        Exit Function
    End If

    feet = Int(inches / 12)
    inches = Round(inches - (feet * 12), 0)
    
    While inches >= 12
        feet = feet + 1
        inches = inches - 12
    Wend
    
    If feet < 0 Then
        inchesToFeetInches = "N/A"
    ElseIf feet = 0 And inches <> 0 Then
        inchesToFeetInches = inches & """"
    Else
        inchesToFeetInches = feet & "'" & inches & """"
    End If
End Function

Public Function convertToInches(ByVal txt As String) As Double
    Dim i As Long
    Dim Ch As String
    Dim num As String
    Dim feet As Long: feet = 0
    Dim inches As Long: inches = 0
    
    If InStr(txt, "'") = 0 And InStr(txt, """") = 0 Then
        convertToInches = -1
        Exit Function
    End If
    
    If UBound(Split(txt, "'")) > 1 Or UBound(Split(txt, """")) > 1 Then
        convertToInches = -1
        Exit Function
    End If
    
    While (InStr(txt, "  ") > 0)
        txt = Replace(txt, "  ", " ")
    Wend
    txt = Replace(txt, " '", "'")
    txt = Replace(txt, " """, """")
    txt = Replace(txt, vbLf, " ")
    parts = Split(txt, " ")
    txt = ""
    For Each part In parts
        If InStr(part, "'") > 0 Or InStr(part, """") > 0 Then txt = txt & part
    Next part
    
    For i = 1 To Len(txt)
        Ch = Mid(txt, i, 1)
        If Ch Like "[0-9.]" Then
            num = num & Ch
        Else
            If num <> "" Then
                Select Case Ch
                    Case "'"
                        If feet = 0 Then feet = feet + CLng(num)
                        num = ""
                    Case """"
                        If inches = 0 Then inches = inches + CLng(num)
                        Exit For
                End Select
            End If
        End If
    Next i

    If inches = 0 And num <> "" Then inches = CLng(num)
    
    If inches = 0 And feet = 0 Then
        convertToInches = -1
    Else
        convertToInches = feet * 12 + inches
    End If
End Function

Public Function sortComponents(components As Variant) As Collection
    Dim arr() As Object
    Dim i As Long, j As Long
    Dim temp As Object
    Set sortedcomponents = New Collection
    
    If TypeOf components Is Collection Then
        If components.count < 2 Then
            Set sortComponents = components
            Exit Function
        End If
        
        ReDim arr(0 To components.count - 1)
        For i = 0 To components.count - 1
            Set arr(i) = components(i + 1)
        Next i
    ElseIf TypeOf components Is Scripting.Dictionary Then
        If components.count < 2 Then
            For Each v In components.items: sortedcomponents.Add v: Next
            Set sortComponents = sortedcomponents
            Exit Function
        End If
    
        items = components.items
        
        ReDim arr(0 To components.count - 1)
        For i = 0 To components.count - 1
            Set arr(i) = items(i)
        Next i
    Else
        Set sortComponents = Nothing
        Exit Function
    End If
    
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i).height < arr(j).height Then
                Set temp = arr(i)
                Set arr(i) = arr(j)
                Set arr(j) = temp
            End If
        Next j
    Next i
    
    For i = LBound(arr) To UBound(arr)
        sortedcomponents.Add arr(i)
    Next i
    
    Set sortComponents = sortedcomponents
End Function

Public Function degreesToDirection(degrees As Integer) As String
    Dim NormalizeAngle As Integer
    
    NormalizeAngle = ((degrees Mod 360) + 360) Mod 360
    
    Select Case NormalizeAngle
        Case 0 To 22
            degreesToDirection = "N"
        Case 23 To 67
            degreesToDirection = "NE"
        Case 68 To 112
            degreesToDirection = "E"
        Case 113 To 157
            degreesToDirection = "SE"
        Case 158 To 202
            degreesToDirection = "S"
        Case 203 To 247
            degreesToDirection = "SW"
        Case 248 To 292
            degreesToDirection = "W"
        Case 293 To 337
            degreesToDirection = "NW"
        Case 338 To 359
            degreesToDirection = "N"
    End Select
End Function

Public Function JoinCollection(col As Variant, delimiter As String) As String
    Dim item As Variant
    Dim result As String
    
    For Each item In col
        If item <> "" Then result = result & item & delimiter
    Next item
    
    If Len(result) > 0 Then
        result = Left(result, Len(result) - Len(delimiter))
    End If
    
    JoinCollection = result
End Function

Public Function IsPDS(sheet As Worksheet) As Boolean
    IsPDS = False
    
    If sheet.name <> "4 Spans" And sheet.name <> "8 Spans" And sheet.name <> "12 Spans" Then
        If sheet.Cells(2, 2).Value = "Notification:" Then
            IsPDS = True
            Exit Function
        End If
    End If
End Function

Function ShareAKey(dict1 As Scripting.Dictionary, dict2 As Scripting.Dictionary) As Boolean
    Dim key As Variant
    
    ShareAKey = False
    
    For Each key In dict1.keys
        If dict2.Exists(key) Then
            ShareAKey = True
            Exit Function
        End If
    Next key
End Function

Public Function RangeExists(sheet As Worksheet, rngAddress As String) As Boolean
    Dim v As Variant
    
    v = sheet.Evaluate("ISREF(" & rngAddress & ")")
    
    RangeExists = (VarType(v) = vbBoolean And v = True)
End Function

Public Function GetPDS(name As String) As Worksheet
    If Utilities.SheetExists(name) Then
        For Each sheet In ThisWorkbook.sheets
            If ThisWorkbook.RemoveParentheses(sheet.name) = name Then
                Set GetPDS = sheet
                Exit Function
            End If
        Next sheet
    End If
    
    Set GetPDS = Nothing
End Function

Public Function SplitByBlankLine(ByVal text As String) As Collection
    Dim result As New Collection
    Dim parts() As String
    Dim i As Long
 
    text = Replace(text, vbCrLf, vbLf)
    text = Replace(text, vbCr, vbLf)
 
    While (InStr(text, vbLf & vbLf & vbLf) > 0)
        text = Replace(text, vbLf & vbLf & vbLf, vbLf & vbLf)
    Wend
 
    parts = Split(text, vbLf & vbLf)
 
    For i = LBound(parts) To UBound(parts)
        If Trim(parts(i)) <> "" Then result.Add parts(i)
    Next i
 
    Set SplitByBlankLine = result
End Function

Public Function GetFirstWord(txt As String) As String
    Dim posSpace As Long, posNewline As Long, pos As Long
    txt = Trim(txt)
    If txt = "" Then Exit Function
    posSpace = InStr(txt, " ")
    posNewline = InStr(txt, vbLf)
    If posNewline = 0 Then posNewline = InStr(txt, vbCr)

    If posSpace = 0 Then
        pos = posNewline
    ElseIf posNewline = 0 Then
        pos = posSpace
    Else
        pos = IIf(posSpace < posNewline, posSpace, posNewline)
    End If
    If pos = 0 Then
        GetFirstWord = txt
    Else
        GetFirstWord = Left(txt, pos - 1)
    End If
End Function

Public Function correctFileName(str As String) As String
    Dim outputStr As String
    outputStr = Replace(str, "*", "")
    outputStr = Replace(outputStr, "<", "")
    outputStr = Replace(outputStr, ">", "")
    outputStr = Replace(outputStr, ":", "")
    outputStr = Replace(outputStr, """", "")
    outputStr = Replace(outputStr, "/", "")
    outputStr = Replace(outputStr, "\", "")
    outputStr = Replace(outputStr, "|", "")
    outputStr = Replace(outputStr, "?", "")
    
    correctFileName = outputStr
End Function

Public Function SwapWords(text As String, word1 As String, word2 As String) As String
    Dim temp As String

    temp = ""
    text = Replace(text, word1, temp)
    text = Replace(text, word2, word1)
    text = Replace(text, temp, word2)
    SwapWords = text
End Function

Public Function autoGLC(poleHeight As Integer, poleSpecies As String, poleClass As String) As String
    If poleSpecies <> "WC" Then poleSpecies = "SP"
    
    Select Case poleHeight
        Case 35
            Select Case poleSpecies
                Case "SP"
                    Select Case poleClass
                        Case "4"
                            autoGLC = "32"" (Auto)"
                        Case "5"
                            autoGLC = "30 1/4"" (Auto)"
                        Case "6"
                            autoGLC = "28"" (Auto)"
                        Case "7"
                            autoGLC = "26"" (Auto)"
                        Case Else
                            autoGLC = ""
                    End Select
                Case "WC", "WRC"
                    Select Case poleClass
                        Case "4"
                            autoGLC = "35 2/4"" (Auto)"
                        Case "5"
                            autoGLC = "33 1/4"" (Auto)"
                        Case "6"
                            autoGLC = "31"" (Auto)"
                        Case Else
                            autoGLC = ""
                    End Select
                Case Else
                    autoGLC = ""
            End Select
        Case 40
            Select Case poleSpecies
                Case "SP"
                    Select Case poleClass
                        Case "2"
                            autoGLC = "39 2/4"" (Auto)"
                        Case "3"
                            autoGLC = "37 1/4"" (Auto)"
                        Case "4"
                            autoGLC = "34 3/4"" (Auto)"
                        Case "5"
                            autoGLC = "32 1/4"" (Auto)"
                        Case "6"
                            autoGLC = "29 3/4"" (Auto)"
                        Case "7"
                            autoGLC = "27 2/4"" (Auto)"
                        Case Else
                            autoGLC = ""
                    End Select
                Case "WC", "WRC"
                    Select Case poleClass
                        Case "2"
                            autoGLC = "43 2/4"" (Auto)"
                        Case "3"
                            autoGLC = "41"" (Auto)"
                        Case "4"
                            autoGLC = "38"" (Auto)"
                        Case "5"
                            autoGLC = "35 1/4"" (Auto)"
                        Case "6"
                            autoGLC = "32 3/4"" (Auto)"
                        Case Else
                            autoGLC = ""
                    End Select
                Case Else
                    autoGLC = ""
            End Select
        Case 45
            Select Case poleSpecies
                Case "SP"
                    Select Case poleClass
                        Case "2"
                            autoGLC = "41 1/4"" (Auto)"
                        Case "3"
                            autoGLC = "38 3/4"" (Auto)"
                        Case "4"
                            autoGLC = "36"" (Auto)"
                        Case "5"
                            autoGLC = "33 2/4"" (Auto)"
                        Case Else
                            autoGLC = ""
                    End Select
                Case "WC", "WRC"
                    Select Case poleClass
                        Case "2"
                            autoGLC = "45 1/4"" (Auto)"
                        Case "3"
                            autoGLC = "42 3/4"" (Auto)"
                        Case "4"
                            autoGLC = "39 3/4"" (Auto)"
                        Case "5"
                            autoGLC = "37 1/4"" (Auto)"
                        Case Else
                            autoGLC = ""
                    End Select
                Case Else
                    autoGLC = ""
            End Select
        Case 50
            Select Case poleSpecies
                Case "SP"
                    Select Case poleClass
                        Case "2"
                            autoGLC = "42 2/4"" (Auto)"
                        Case "3"
                            autoGLC = "40"" (Auto)"
                        Case "4"
                            autoGLC = "37 1/4"" (Auto)"
                        Case Else
                            autoGLC = ""
                    End Select
                Case "WC", "WRC"
                    Select Case poleClass
                        Case "2"
                            autoGLC = "47"" (Auto)"
                        Case "3"
                            autoGLC = "44 2/4"" (Auto)"
                        Case "4"
                            autoGLC = "41 3/4"" (Auto)"
                        Case Else
                            autoGLC = ""
                    End Select
                Case Else
                    autoGLC = ""
            End Select
        Case 55
            Select Case poleSpecies
                Case "SP"
                    Select Case poleClass
                        Case "2"
                            autoGLC = "44"" (Auto)"
                        Case "3"
                            autoGLC = "41 2/4"" (Auto)"
                        Case "4"
                            autoGLC = "38 3/4"" (Auto)"
                        Case Else
                            autoGLC = ""
                    End Select
                Case "WC", "WRC"
                    Select Case poleClass
                        Case "2"
                            autoGLC = "48 3/4"" (Auto)"
                        Case "3"
                            autoGLC = "46"" (Auto)"
                        Case "4"
                            autoGLC = "42 3/4"" (Auto)"
                        Case Else
                            autoGLC = ""
                    End Select
                Case Else
                    autoGLC = ""
            End Select
        Case 60
            Select Case poleSpecies
                Case "SP"
                    Select Case poleClass
                        Case "2"
                            autoGLC = "45 1/4"" (Auto)"
                        Case "3"
                            autoGLC = "42 3/4"" (Auto)"
                        Case "4"
                            autoGLC = "39 3/4"" (Auto)"
                        Case Else
                            autoGLC = ""
                    End Select
                Case "WC", "WRC"
                    Select Case poleClass
                        Case "2"
                            autoGLC = "50"" (Auto)"
                        Case "3"
                            autoGLC = "47 1/4"" (Auto)"
                        Case "4"
                            autoGLC = "44"" (Auto)"
                        Case Else
                            autoGLC = ""
                    End Select
                Case Else
                    autoGLC = ""
            End Select
        Case 65
            Select Case poleSpecies
                Case "SP"
                    Select Case poleClass
                        Case "2"
                            autoGLC = "46 2/4"" (Auto)"
                        Case "3"
                            autoGLC = "44"" (Auto)"
                        Case "4"
                            autoGLC = "41 1/4"" (Auto)"
                        Case Else
                            autoGLC = ""
                    End Select
                Case "WC", "WRC"
                    Select Case poleClass
                        Case "2"
                            autoGLC = "51 2/4"" (Auto)"
                        Case "3"
                            autoGLC = "48 3/4"" (Auto)"
                        Case "4"
                            autoGLC = "44 3/4"" (Auto)"
                        Case Else
                            autoGLC = ""
                    End Select
                Case Else
                    autoGLC = ""
            End Select
        Case 70
            Select Case poleSpecies
                Case "SP"
                    Select Case poleClass
                        Case "2"
                            autoGLC = "48"" (Auto)"
                        Case "3"
                            autoGLC = "45 2/4"" (Auto)"
                        Case "4"
                            autoGLC = "42 1/4"" (Auto)"
                        Case Else
                            autoGLC = ""
                    End Select
                Case "WC", "WRC"
                    Select Case poleClass
                        Case "2"
                            autoGLC = "52 3/4"" (Auto)"
                        Case "3"
                            autoGLC = "50"" (Auto)"
                        Case "4"
                            autoGLC = "46 2/4"" (Auto)"
                        Case Else
                            autoGLC = ""
                    End Select
                Case Else
                    autoGLC = ""
            End Select
        Case 75
            Select Case poleSpecies
                Case "SP"
                    Select Case poleClass
                        Case "2"
                            autoGLC = "48 3/4"" (Auto)"
                        Case "3"
                            autoGLC = "46 1/4"" (Auto)"
                        Case "4"
                            autoGLC = "43 2/4"" (Auto)"
                        Case Else
                            autoGLC = ""
                    End Select
                Case "WC", "WRC"
                    Select Case poleClass
                        Case "2"
                            autoGLC = "54"" (Auto)"
                        Case "3"
                            autoGLC = "51 1/4"" (Auto)"
                        Case "4"
                            autoGLC = "48"" (Auto)"
                        Case Else
                            autoGLC = ""
                    End Select
                Case Else
                    autoGLC = ""
            End Select
        Case 80
            Select Case poleSpecies
                Case "SP"
                    Select Case poleClass
                        Case "2"
                            autoGLC = "50"" (Auto)"
                        Case "3"
                            autoGLC = "47 2/4"" (Auto)"
                        Case Else
                            autoGLC = ""
                    End Select
                Case "WC", "WRC"
                    Select Case poleClass
                        Case "2"
                            autoGLC = "55 1/4"" (Auto)"
                        Case "3"
                            autoGLC = "52 2/4"" (Auto)"
                        Case Else
                            autoGLC = ""
                    End Select
                Case Else
                    autoGLC = ""
            End Select
        Case 85
            Select Case poleSpecies
                Case "SP"
                    Select Case poleClass
                        Case "2"
                            autoGLC = "51"" (Auto)"
                        Case "3"
                            autoGLC = "48 1/4"" (Auto)"
                        Case Else
                            autoGLC = ""
                    End Select
                Case "WC", "WRC"
                    Select Case poleClass
                        Case "2"
                            autoGLC = "56 1/4"" (Auto)"
                        Case "3"
                            autoGLC = "51"" (Auto)"
                        Case Else
                            autoGLC = ""
                    End Select
                Case Else
                    autoGLC = ""
            End Select
        Case Else
            autoGLC = ""
    End Select
End Function
