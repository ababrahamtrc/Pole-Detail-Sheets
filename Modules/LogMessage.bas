Attribute VB_Name = "LogMessage"
Dim httpAsync As Object

Sub SendLogMessage(script As String)
    On Error Resume Next

    Dim url As String
    Dim userName As String
    Dim project As project: Set project = New project
    
    mode = IIf(project.mode = "MAKE READY", "MKR", "SI")
    
    userName = Environ$("USERNAME")
    
    If Environ$("USERDOMAIN") = "CE" Then
        Debug.Print "Sending log message for " & userName & ": " & script
        url = "https://script.google.com/macros/s/AKfycbwsg9vKpbad32eug1B0lSddjABtCb0MmZBjygFLXt3UlNiPdDhupg3gwYbSQpsviE8H4w/exec" & _
             "?user=" & userName & _
             "&message=" & "PDS " & mode & " " & script
    
       Set httpAsync = CreateObject("MSXML2.XMLHTTP.6.0")
       httpAsync.Open "Get", url, True
       httpAsync.Send
    End If
End Sub



