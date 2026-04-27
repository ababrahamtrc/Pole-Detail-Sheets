Attribute VB_Name = "LogMessage"
Dim httpAsync As Object

Sub SendLogMessage(script As String)
    On Error Resume Next

    Dim url As String
    Dim UserName As String
    
    UserName = Environ$("USERNAME")
    
    If Environ$("USERDOMAIN") = "CE" Then
        Debug.Print "Sending log message for " & UserName & ": " & script
        url = "https://script.google.com/macros/s/AKfycbwsg9vKpbad32eug1B0lSddjABtCb0MmZBjygFLXt3UlNiPdDhupg3gwYbSQpsviE8H4w/exec" & _
             "?user=" & UserName & _
             "&message=" & "PDS " & script
    
       Set httpAsync = CreateObject("MSXML2.XMLHTTP.6.0")
       httpAsync.Open "Get", url, True
       httpAsync.Send
    End If
End Sub



