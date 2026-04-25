VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal Target As Range)
    On Error Resume Next
    
    Application.EnableEvents = True
    
    If sh.name <> "4 Spans" And sh.name <> "8 Spans" And sh.name <> "12 Spans" Then
        If sh.Cells(2, 2).Value = "Notification:" Then
            If Not Intersect(Target, sh.Range("DL")) Is Nothing Or Not Intersect(Target, sh.Range("SUMSHEET9")) Is Nothing Or Not Intersect(Target, sh.Range("SUMSHEET8")) Is Nothing Or Not Intersect(Target, sh.Range("SUMSHEET4")) Is Nothing Or Not Intersect(Target, sh.Range("SUMSHEET1")) Is Nothing Or Not Intersect(Target, sh.Range("SUMSHEET7")) Is Nothing Then
                Call decideTabColor(sh)
            ElseIf Not Intersect(Target, sh.Range("SUMSHEET3")) Is Nothing Or Not Intersect(Target, sh.Range("SUMSHEET6")) Is Nothing Then
                If sh.Range("SUMSHEET3").Value Then
                    If sh.Range("SUMSHEET6") Then
                        sh.Range("SUMSHEET11").Value = "1" & Chr(44) & "2" & Chr(44) & "3"
                    Else
                        sh.Range("SUMSHEET11").Value = "1" & Chr(44) & "2"
                    End If
                ElseIf sh.Range("SUMSHEET6") Then
                    sh.Range("SUMSHEET11").Value = "1" & Chr(44) & "3"
                Else
                    sh.Range("SUMSHEET11").Value = ""
                End If
            End If
        End If
    End If
End Sub

Public Sub decideTabColor(sh As Worksheet)
    On Error Resume Next

    If Trim(sh.Range("DL")) <> "" And ThisWorkbook.sheets("Control").Range("COLOR1").offset(0, 1).Interior.ColorIndex <> xlNone Then
        If sh.Tab.color <> ThisWorkbook.sheets("Control").Range("COLOR1").offset(0, 1).Interior.color Then sh.Tab.color = ThisWorkbook.sheets("Control").Range("COLOR1").offset(0, 1).Interior.color
        If sh.name <> Trim(sh.Range("POLENUM")) & "(" & Trim(fixSpecialCharacters(sh.Range("DL"))) & ")" Then sh.name = Trim(sh.Range("POLENUM")) & "(" & Trim(fixSpecialCharacters(sh.Range("DL"))) & ")"
    ElseIf sh.Range("SUMSHEET9") = True And ThisWorkbook.sheets("Control").Range("COLOR2").offset(0, 1).Interior.ColorIndex <> xlNone Then
        If sh.Tab.color <> ThisWorkbook.sheets("Control").Range("COLOR1").offset(0, 1).Interior.color Then sh.Tab.color = ThisWorkbook.sheets("Control").Range("COLOR2").offset(0, 1).Interior.color
        If Trim(sh.Range("DL")) <> "" Then
            If sh.name <> Trim(sh.Range("POLENUM")) & "(" & Trim(fixSpecialCharacters(sh.Range("DL"))) & ")" Then sh.name = Trim(sh.Range("POLENUM")) & "(" & Trim(fixSpecialCharacters(sh.Range("DL"))) & ")"
        Else
            If sh.name <> Trim(sh.Range("POLENUM")) Then sh.name = Trim(sh.Range("POLENUM"))
        End If
    ElseIf sh.Range("SUMSHEET8") = True And ThisWorkbook.sheets("Control").Range("COLOR3").offset(0, 1).Interior.ColorIndex <> xlNone Then
        If sh.Tab.color <> ThisWorkbook.sheets("Control").Range("COLOR3").offset(0, 1).Interior.color Then sh.Tab.color = ThisWorkbook.sheets("Control").Range("COLOR3").offset(0, 1).Interior.color
        If Trim(sh.Range("DL")) <> "" Then
            If sh.name <> Trim(sh.Range("POLENUM")) & "(" & Trim(fixSpecialCharacters(sh.Range("DL"))) & ")" Then sh.name = Trim(sh.Range("POLENUM")) & "(" & Trim(fixSpecialCharacters(sh.Range("DL"))) & ")"
        Else
            If sh.name <> Trim(sh.Range("POLENUM")) Then sh.name = Trim(sh.Range("POLENUM"))
        End If
    ElseIf sh.Range("SUMSHEET4") = True And ThisWorkbook.sheets("Control").Range("COLOR4").offset(0, 1).Interior.ColorIndex <> xlNone Then
        If sh.Tab.color <> ThisWorkbook.sheets("Control").Range("COLOR1").offset(0, 1).Interior.color Then sh.Tab.color = ThisWorkbook.sheets("Control").Range("COLOR4").offset(0, 1).Interior.color
        If Trim(sh.Range("DL")) <> "" Then
            If sh.name <> Trim(sh.Range("POLENUM")) & "(" & Trim(fixSpecialCharacters(sh.Range("DL"))) & ")" Then sh.name = Trim(sh.Range("POLENUM")) & "(" & Trim(fixSpecialCharacters(sh.Range("DL"))) & ")"
        Else
            If sh.name <> Trim(sh.Range("POLENUM")) Then sh.name = Trim(sh.Range("POLENUM"))
        End If
    ElseIf (sh.Range("SUMSHEET1") = True Or sh.Range("SUMSHEET7") = True) And ThisWorkbook.sheets("Control").Range("COLOR5").offset(0, 1).Interior.ColorIndex <> xlNone Then
        If sh.Tab.color <> ThisWorkbook.sheets("Control").Range("COLOR1").offset(0, 1).Interior.color Then sh.Tab.color = ThisWorkbook.sheets("Control").Range("COLOR5").offset(0, 1).Interior.color
        If Trim(sh.Range("DL")) <> "" Then
            If sh.name <> Trim(sh.Range("POLENUM")) & "(" & Trim(fixSpecialCharacters(sh.Range("DL"))) & ")" Then sh.name = Trim(sh.Range("POLENUM")) & "(" & Trim(fixSpecialCharacters(sh.Range("DL"))) & ")"
        Else
            If sh.name <> Trim(sh.Range("POLENUM")) Then sh.name = Trim(sh.Range("POLENUM"))
        End If
    Else
        If sh.Tab.color <> False Then sh.Tab.color = False
        If Trim(sh.Range("DL")) <> "" Then
            If sh.name <> Trim(sh.Range("POLENUM")) & "(" & Trim(fixSpecialCharacters(sh.Range("DL"))) & ")" Then sh.name = Trim(sh.Range("POLENUM")) & "(" & Trim(fixSpecialCharacters(sh.Range("DL"))) & ")"
        ElseIf sh.name <> "4 Spans" And sh.name <> "8 Spans" And sh.name <> "12 Spans" Then
            If sh.name <> Trim(sh.Range("POLENUM")) Then sh.name = sh.Range("POLENUM")
        End If
    End If
    
End Sub


Public Function RemoveParentheses(ByVal inputStr As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    regex.Pattern = "\([^)]*\)"
    regex.Global = True
    regex.IgnoreCase = True
 
    RemoveParentheses = regex.Replace(inputStr, "")
End Function

Function InsideParentheses(ByVal inputStr As String) As String
    Dim openParenPos As Integer
    Dim closeParenPos As Integer
    Dim valueLength As Integer
    
    openParenPos = InStr(inputStr, "(")
    closeParenPos = InStr(inputStr, ")")
    
    If openParenPos > 0 And closeParenPos > 0 And closeParenPos > openParenPos Then
        valueLength = closeParenPos - openParenPos - 1
        InsideParentheses = Mid(inputStr, openParenPos + 1, valueLength)
    Else
        InsideParentheses = ""
    End If
End Function

Public Function fixSpecialCharacters(ByVal inputStr As String) As String
    inputStr = Replace(inputStr, "*", "#")
    inputStr = Replace(inputStr, ":", "#")
    inputStr = Replace(inputStr, "/", "#")
    inputStr = Replace(inputStr, "\", "#")
    inputStr = Replace(inputStr, "?", "#")
    inputStr = Replace(inputStr, "[", "#")
    inputStr = Replace(inputStr, "]", "#")

    fixSpecialCharacters = inputStr
End Function

Private Sub Workbook_Open()
    
    ' Checks if user is using an outdated PDS
    Const VERSION_URL As String = _
        "https://raw.githubusercontent.com/ElijahRademaker/automation-tools/refs/heads/main/version.txt"

    Const APP_NAME As String = "PoleDetailSheets"
    Const SECTION_NAME As String = "VersionCheck"
    Const SETTING_KEY As String = "LastWarningDate"

    Dim http As Object
    Dim response As String
    Dim lines() As String
    Dim i As Long

    Dim remoteVersion As String
    Dim localVersion As String
    Dim fileName As String
    Dim found As Boolean
    Dim vPos As Long

    Dim lastWarnDate As String

    lastWarnDate = GetSetting(APP_NAME, SECTION_NAME, SETTING_KEY, "")

    If lastWarnDate = Format(Date, "yyyy-mm-dd") Then
         Exit Sub
    End If

    fileName = ThisWorkbook.name

    If InStrRev(fileName, ".") > 0 Then
        fileName = Left(fileName, InStrRev(fileName, ".") - 1)
    End If

    vPos = InStrRev(UCase(fileName), "V")

    If vPos = 0 Or vPos = Len(fileName) Then
        MsgBox _
            "Unable to determine Pole Detail Sheet version from filename." & vbCrLf & _
            "Expected format:  Pole Detail Sheets V<version>", _
            vbExclamation, "Version Warning"
        Exit Sub
    End If

    localVersion = Trim(ThisWorkbook.RemoveParentheses(Mid(fileName, vPos + 1)))
    If InStr(localVersion, " ") > 0 Then localVersion = Split(localVersion, " ")(0)
    Debug.Print localVersion

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", VERSION_URL, False
    http.Send

    If http.Status <> 200 Then
        Exit Sub
    End If

    response = http.responseText
    lines = Split(response, vbLf)

    For i = LBound(lines) To UBound(lines)
        If Trim(lines(i)) Like "POLEDETAILSHEETVERSION=*" Then
            remoteVersion = Trim(Split(lines(i), "=")(1))
            found = True
            Exit For
        End If
    Next i

    If Not found Then Exit Sub

    If remoteVersion <> localVersion Then

        MsgBox _
            "=== OUTDATED POLE DETAIL SHEETS ===" & vbCrLf & vbCrLf & _
            "Your version:       v" & localVersion & vbCrLf & _
            "Current version: v" & remoteVersion & vbCrLf & vbCrLf & _
            "Please obtain the latest version when possible.", _
            vbExclamation, "Outdated Version Warning"

        SaveSetting APP_NAME, SECTION_NAME, SETTING_KEY, Format(Date, "yyyy-mm-dd")

    Else
        SaveSetting APP_NAME, SECTION_NAME, SETTING_KEY, ""
    End If

End Sub





