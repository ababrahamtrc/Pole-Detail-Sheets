Attribute VB_Name = "mod_getDocWarehouseLink"
Public Function getDocWarehouseLink(ByVal notification As Variant) As String

    Const EAM_REDIRECT_BASE As String = _
        "http://eamdashboard/Base/RedirectToDocWareHouse?Notification="

    Const WINHTTP_OPTION_ENABLE_REDIRECTS As Long = 6
    Const AUTO_LOGON_ALWAYS As Long = 0

    Dim request As Object
    Dim requestUrl As String
    Dim redirectUrl As String
    Dim statusCode As Long
    Dim redirectCount As Long
    Dim notificationNumber As String

    notificationNumber = Trim$(CStr(notification))

    If Len(notificationNumber) = 0 Then Exit Function

    requestUrl = EAM_REDIRECT_BASE & notificationNumber

    On Error GoTo Failed

    For redirectCount = 1 To 10

        Set request = CreateObject("WinHttp.WinHttpRequest.5.1")

        request.Open "GET", requestUrl, False

        'Use current Windows credentials
        request.SetAutoLogonPolicy AUTO_LOGON_ALWAYS

        'Do not automatically follow redirects
        request.Option(WINHTTP_OPTION_ENABLE_REDIRECTS) = False

        request.Send

        statusCode = request.status

        Select Case statusCode

            Case 301, 302, 303, 307, 308

                redirectUrl = ""
                redirectUrl = request.GetResponseHeader("Location")

                If Len(redirectUrl) = 0 Then Exit Function

                'If we have reached SharePoint, return it
                If InStr(1, redirectUrl, _
                    "consumersenergy.sharepoint.com", _
                    vbTextCompare) > 0 Then

                    getDocWarehouseLink = redirectUrl
                    Exit Function

                End If

                requestUrl = redirectUrl

            Case 200

                'If the resulting URL itself is SharePoint
                If InStr(1, requestUrl, _
                    "consumersenergy.sharepoint.com", _
                    vbTextCompare) > 0 Then

                    getDocWarehouseLink = requestUrl
                    Exit Function

                End If

                Exit Function

            Case Else
                Exit Function

        End Select

    Next redirectCount

    Exit Function

Failed:
    getDocWarehouseLink = ""

End Function

