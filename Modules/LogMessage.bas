Attribute VB_Name = "LogMessage"
Dim httpAsync As Object

Sub SendLogMessage(script As String)
    On Error Resume Next

    Dim url As String
    Dim UserName As String
    
    UserName = Environ$("USERDOMAIN") & " " & Environ$("USERNAME")
    
    fileName = ThisWorkbook.name

    If InStrRev(fileName, ".") > 0 Then
        fileName = Left(fileName, InStrRev(fileName, ".") - 1)
    End If

    vPos = InStrRev(UCase(fileName), "V")

    If vPos = 0 Or vPos = Len(fileName) Then
        MsgBox _
            "Unable to determine local version from filename." & vbCrLf & _
            "Expected format:  Pole Detail Sheets V<version>", _
            vbCritical, "Version Check Failed"
        Exit Sub
    End If

    localVersion = Trim(ThisWorkbook.RemoveParentheses(Mid(fileName, vPos + 1)))
    If InStr(localVersion, " ") > 0 Then localVersion = Split(localVersion, " ")(0)
 
    Debug.Print "Sending log message for " & UserName & ", V" & localVersion & ": " & script
    
    If Environ$("USERDOMAIN") = "CE" Then
       url = "https://script.google.com/macros/s/AKfycbyhq0VhByIT6hSXzxsyM6q8WqtddyH3ugBbmSzPFrvpZgBI428i1dxGoheIHnV5V1IZEA/exec" & _
             "?user=" & UserName & _
             "&message=" & "V" & localVersion & " PDS " & script
    
       Set httpAsync = CreateObject("MSXML2.XMLHTTP.6.0")
       httpAsync.Open "Get", url, True
       httpAsync.Send
    End If
End Sub



