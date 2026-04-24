Attribute VB_Name = "PoleClassHeightChecker"
Public Sub CheckPole()
    On Error Resume Next

    Call LogMessage.SendLogMessage("CheckPole")

    Set sheet = Application.ActiveSheet()
    If sheet.name = "4 Spans" Or sheet.name = "8 Spans" Or sheet.name = "12 Spans" Or sheet.Cells(2, 2).Value <> "Notification:" Then
        MsgBox "You need to have a pole detail sheet active to run this script."
        Exit Sub
    End If
    
    Dim expectedClass As String
    Dim potentialClass As String
    Dim Class As String
    
    height = sheet.Range("HEIGHT")
    Class = sheet.Range("CLASS")
    species = sheet.Range("TYPE")
    glc = sheet.Range("GLC")
    If (InStr(glc, "(Auto)") > 1) Then
        expectedClass = Class
        potentialClass = Class
    ElseIf (InStr(Class, "H") > 0) Then
        expectedClass = 0
        potentialClass = 0
    Else
        glc = CDbl(GLCConvert(glc))
        expectedClass = getExpectedClass(val(height), CStr(species), glc)
        potentialClass = getExpectedClass(val(height), CStr(species), glc + 0.5)
    End If

    If glc = 0 Or height = "" Or species = "" Then
        MsgBox "Missing GLC, Height, or Species from sheet."
        Exit Sub
    End If
    
    If (potentialClass <> expectedClass And expectedClass = Class) Then
        MsgBox "Measured GLC 0.5 inches or less away from a thicker class. Check if pole is branded."
    ElseIf (potentialClass <> expectedClass And potentialClass = Class) Then
        MsgBox "Measured GlC is within range of given class and species."
    ElseIf (expectedClass <> 0 And expectedClass <> Class) Then
        MsgBox "Measured GLC not within range of given class and species. Expected class is " & expectedClass & ". Check if pole is branded."
    ElseIf (expectedClass = 0) Then
        MsgBox "Skipped by Pole Checker, measured GLC and species doesn't fit any range in CE Standard Work."
    ElseIf expectedClass = Class Then
        MsgBox "Measured GlC is within range of given class and species."
    End If

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Private Function GLCConvert(glc As Variant) As String
    glc = CStr(glc)
    
    glc = Replace(glc, """", "")
    glc = Replace(glc, " ", "")
    glc = Replace(glc, "(Auto)", "")
    glc = Replace(glc, "0/4", ".00")
    glc = Replace(glc, "1/4", ".25")
    glc = Replace(glc, "1/2", ".50")
    glc = Replace(glc, "2/4", ".50")
    glc = Replace(glc, "3/4", ".75")
    If glc = "" Then glc = "0"
    
    GLCConvert = glc
End Function

Private Function getExpectedClass(poleHeight As Integer, poleSpecies As String, poleGLC As Variant) As String
    Select Case poleHeight
        Case 35
            Select Case poleSpecies
                Case "SP"
                    Select Case poleGLC
                        Case 31.5 To 100
                            getExpectedClass = 4
                        Case 29 To 31.5
                            getExpectedClass = 5
                        Case 27 To 29
                            getExpectedClass = 6
                        Case 25 To 27
                            getExpectedClass = 7
                        Case Else
                            getExpectedClass = 0
                    End Select
                Case "WC", "WRC"
                    Select Case poleGLC
                        Case 34.5 To 100
                            getExpectedClass = 4
                        Case 32 To 34.5
                            getExpectedClass = 5
                        Case 30 To 32
                            getExpectedClass = 6
                        Case Else
                            getExpectedClass = 0
                    End Select
                Case Else
                    getExpectedClass = 0
            End Select
        Case 40
            Select Case poleSpecies
                Case "SP"
                    Select Case poleGLC
                        Case 38.5 To 100
                            getExpectedClass = 2
                        Case 36 To 38.5
                            getExpectedClass = 3
                        Case 33.5 To 36
                            getExpectedClass = 4
                        Case 31 To 33.5
                            getExpectedClass = 5
                        Case 28.5 To 31
                            getExpectedClass = 6
                        Case 26.5 To 28.5
                            getExpectedClass = 7
                        Case Else
                            getExpectedClass = 0
                    End Select
                Case "WC", "WRC"
                    Select Case poleGLC
                        Case 42.5 To 100
                            getExpectedClass = 2
                        Case 39.5 To 42.5
                            getExpectedClass = 3
                        Case 36.5 To 39.5
                            getExpectedClass = 4
                        Case 34 To 36.5
                            getExpectedClass = 5
                        Case 31.5 To 34
                            getExpectedClass = 6
                        Case Else
                            getExpectedClass = 0
                    End Select
                Case Else
                    getExpectedClass = 0
            End Select
        Case 45
            Select Case poleSpecies
                Case "SP"
                    Select Case poleGLC
                        Case 40.25 To 100
                            getExpectedClass = 2
                        Case 37.25 To 40.25
                            getExpectedClass = 3
                        Case 34.75 To 37.25
                            getExpectedClass = 4
                        Case 32.5 To 34.75
                            getExpectedClass = 5
                        Case Else
                            getExpectedClass = 0
                    End Select
                Case "WC", "WRC"
                    Select Case poleGLC
                        Case 44.25 To 100
                            getExpectedClass = 2
                        Case 41.25 To 44.25
                            getExpectedClass = 3
                        Case 38.25 To 41.25
                            getExpectedClass = 4
                        Case 36.2 To 38.25
                            getExpectedClass = 5
                        Case Else
                            getExpectedClass = 0
                    End Select
                Case Else
                    getExpectedClass = 0
            End Select
        Case 50
            Select Case poleSpecies
                Case "SP"
                    Select Case poleGLC
                        Case 41.5 To 100
                            getExpectedClass = 2
                        Case 38.5 To 41.5
                            getExpectedClass = 3
                        Case 36 To 38.5
                            getExpectedClass = 4
                        Case Else
                            getExpectedClass = 0
                    End Select
                Case "WC", "WRC"
                    Select Case poleGLC
                        Case 46 To 100
                            getExpectedClass = 2
                        Case 43 To 46
                            getExpectedClass = 3
                        Case 40.4 To 43
                            getExpectedClass = 4
                        Case Else
                            getExpectedClass = 0
                    End Select
                Case Else
                    getExpectedClass = 0
            End Select
        Case 55
            Select Case poleSpecies
                Case "SP"
                    Select Case poleGLC
                        Case 42.9 To 100
                            getExpectedClass = 2
                        Case 40 To 42.9
                            getExpectedClass = 3
                        Case 37.5 To 40
                            getExpectedClass = 4
                        Case Else
                            getExpectedClass = 0
                    End Select
                Case "WC", "WRC"
                    Select Case poleGLC
                        Case 47.75 To 100
                            getExpectedClass = 2
                        Case 44.25 To 47.75
                            getExpectedClass = 3
                        Case 41.4 To 44.25
                            getExpectedClass = 4
                        Case Else
                            getExpectedClass = 0
                    End Select
                Case Else
                    getExpectedClass = 0
            End Select
        Case 60
            Select Case poleSpecies
                Case "SP"
                    Select Case poleGLC
                        Case 44.25 To 100
                            getExpectedClass = 2
                        Case 41.25 To 44.25
                            getExpectedClass = 3
                        Case 38.25 To 41.25
                            getExpectedClass = 4
                        Case Else
                            getExpectedClass = 0
                    End Select
                Case "WC", "WRC"
                    Select Case poleGLC
                        Case 49 To 100
                            getExpectedClass = 2
                        Case 45.5 To 49
                            getExpectedClass = 3
                        Case 42.7 To 45.5
                            getExpectedClass = 4
                        Case Else
                            getExpectedClass = 0
                    End Select
                Case Else
                    getExpectedClass = 0
            End Select
        Case 65
            Select Case poleSpecies
                Case "SP"
                    Select Case poleGLC
                        Case 45.5 To 100
                            getExpectedClass = 2
                        Case 42.5 To 45.5
                            getExpectedClass = 3
                        Case 39.7 To 42.5
                            getExpectedClass = 4
                        Case Else
                            getExpectedClass = 0
                    End Select
                Case "WC", "WRC"
                    Select Case poleGLC
                        Case 50.4 To 100
                            getExpectedClass = 2
                        Case 46.9 To 50.4
                            getExpectedClass = 3
                        Case 42.5 To 46.9
                            getExpectedClass = 4
                        Case Else
                            getExpectedClass = 0
                    End Select
                Case Else
                    getExpectedClass = 0
            End Select
        Case 70
            Select Case poleSpecies
                Case "SP"
                    Select Case poleGLC
                        Case 46.9 To 100
                            getExpectedClass = 2
                        Case 44 To 46.9
                            getExpectedClass = 3
                        Case 40.5 To 44
                            getExpectedClass = 4
                        Case Else
                            getExpectedClass = 0
                    End Select
                Case "WC", "WRC"
                    Select Case poleGLC
                        Case 51.7 To 100
                            getExpectedClass = 2
                        Case 48.25 To 51.7
                            getExpectedClass = 3
                        Case 44.75 To 48.25
                            getExpectedClass = 4
                        Case Else
                            getExpectedClass = 0
                    End Select
                Case Else
                    getExpectedClass = 0
            End Select
        Case 75
            Select Case poleSpecies
                Case "SP"
                    Select Case poleGLC
                        Case 47.75 To 100
                            getExpectedClass = 2
                        Case 44.75 To 47.75
                            getExpectedClass = 3
                        Case 41.9 To 44.75
                            getExpectedClass = 4
                        Case Else
                            getExpectedClass = 0
                    End Select
                Case "WC", "WRC"
                    Select Case poleGLC
                        Case 53 To 100
                            getExpectedClass = 2
                        Case 49.5 To 53
                            getExpectedClass = 3
                        Case 46.2 To 49.5
                            getExpectedClass = 3
                        Case Else
                            getExpectedClass = 0
                    End Select
                Case Else
                    getExpectedClass = 0
            End Select
        Case 80
            Select Case poleSpecies
                Case "SP"
                    Select Case poleGLC
                        Case 49 To 100
                            getExpectedClass = 2
                        Case 45.7 To 49
                            getExpectedClass = 3
                        Case Else
                            getExpectedClass = 0
                    End Select
                Case "WC", "WRC"
                    Select Case poleGLC
                        Case 54.25 To 100
                            getExpectedClass = 2
                        Case 50.4 To 54.25
                            getExpectedClass = 3
                        Case Else
                            getExpectedClass = 0
                    End Select
                Case Else
                    getExpectedClass = 0
            End Select
        Case 85
            Select Case poleSpecies
                Case "SP"
                    Select Case poleGLC
                        Case 50 To 100
                            getExpectedClass = 2
                        Case 46.5 To 50
                            getExpectedClass = 3
                        Case Else
                            getExpectedClass = 0
                    End Select
                Case "WC", "WRC"
                    Select Case poleGLC
                        Case 55.2 To 100
                            getExpectedClass = 2
                        Case 46.5 To 55.2
                            getExpectedClass = 3
                        Case Else
                            getExpectedClass = 0
                    End Select
                Case Else
                    getExpectedClass = 0
            End Select
        Case Else
            getExpectedClass = 0
    End Select
End Function
