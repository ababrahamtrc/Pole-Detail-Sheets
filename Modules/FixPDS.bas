Attribute VB_Name = "FixPDS"
Public Sub fixAttachmentHeights()
    On Error Resume Next
    
    Call LogMessage.SendLogMessage("fixAttachmentHeights")
    
    Dim sheet As Worksheet: Set sheet = ThisWorkbook.ActiveSheet()
    If Not Utilities.IsPDS(sheet) Then
        MsgBox "You need to have a pole detail sheet active to run this script."
        Exit Sub
    End If
    
    Dim spans As Integer: spans = 0
    
    answer = MsgBox("Do you want to automatically sort the attachments in the Utility and Communications sections?" & pole, vbYesNoCancel + vbQuestion, "Confirmation")
    If answer <> vbYes Then Exit Sub
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    For i = 1 To 12
        For Each name In sheet.names
            If name.name = "'" & sheet.name & "'" & "!" & "TOPOLE" & i Then spans = spans + 1
        Next name
    Next i
    
    Dim comp As Component
    Dim utilityComponentsDict As scripting.Dictionary: Set utilityComponentsDict = New scripting.Dictionary
    Dim utilityComponents As Collection: Set utilityComponents = New Collection
    Dim commComponentsDict As scripting.Dictionary: Set commComponentsDict = New scripting.Dictionary
    Dim commComponents As Collection: Set commComponents = New Collection
    Dim midspans As scripting.Dictionary
    
    For i = 1 To 100
        If sheet.Range("UTTYPE").offset(i - 1, 0).Interior.color <> 16312794 Then Exit For
        If Trim(sheet.Range("UTHEIGHT").offset(i - 1, 0).Value) <> "" Then
            Set comp = New Component
            comp.height = Utilities.convertToInches(ThisWorkbook.RemoveParentheses(sheet.Range("UTHEIGHT").offset(i - 1, 0).text))
            comp.bottomHeight = Utilities.convertToInches(ThisWorkbook.InsideParentheses(sheet.Range("UTHEIGHT").offset(i - 1, 0).text))
            comp.componentType = sheet.Range("UTTYPE").offset(i - 1, 0).text
            comp.size = sheet.Range("UTSIZE").offset(i - 1, 0).text
            key = comp.height & comp.componentType & comp.size
            For j = 1 To spans
                midspan = sheet.Range("UTMIDSPAN" & j).offset(i - 1, 0).text
                If midspan = "" Then midspan = "-"
                comp.midspans.Add j, midspan
            Next j
            Do
                sharingSpan = False
                If utilityComponentsDict.Exists(key) Then
                    For Each midspanSlot In comp.midspans
                        If Replace(comp.midspans(midspanSlot), "-", "") <> "" And Replace(utilityComponentsDict(key).midspans(midspanSlot), "-", "") <> "" Then
                            sharingSpan = True
                            Exit For
                        End If
                    Next midspanSlot
                    If sharingSpan Then
                        key = key & 1
                    Else
                        For Each midspanSlot In comp.midspans
                            If Replace(comp.midspans(midspanSlot), "-", "") <> "" And Replace(utilityComponentsDict(key).midspans(midspanSlot), "-", "") = "" Then
                                utilityComponentsDict(key).midspans(midspanSlot) = comp.midspans(midspanSlot)
                            End If
                        Next midspanSlot
                        Exit Do
                    End If
                End If
            Loop While utilityComponentsDict.Exists(key)
            If Not utilityComponentsDict.Exists(key) Then utilityComponentsDict.Add key, comp
        End If
    Next i
    
    For i = 1 To 100
        If sheet.Range("CMOWNER").offset(i - 1, 0).Interior.color <> 16312794 Then Exit For
        If Trim(sheet.Range("CMHEIGHT").offset(i - 1, 0).Value) <> "" Then
            Set comp = New Component
            comp.height = Utilities.convertToInches(ThisWorkbook.RemoveParentheses(sheet.Range("CMHEIGHT").offset(i - 1, 0).text))
            comp.bottomHeight = Utilities.convertToInches(ThisWorkbook.InsideParentheses(sheet.Range("CMHEIGHT").offset(i - 1, 0).text))
            comp.owner = sheet.Range("CMOWNER").offset(i - 1, 0).text
            comp.size = sheet.Range("CMSIZE").offset(i - 1, 0).text
            key = comp.height & comp.owner & comp.size
            For j = 1 To spans
                midspan = sheet.Range("CMMIDSPAN" & j).offset(i - 1, 0).text
                If midspan = "" Then midspan = "-"
                comp.midspans.Add j, midspan
            Next j
            Do
                sharingSpan = False
                If commComponentsDict.Exists(key) Then
                    For Each midspanSlot In comp.midspans
                        If Replace(comp.midspans(midspanSlot), "-", "") <> "" And Replace(commComponentsDict(key).midspans(midspanSlot), "-", "") <> "" Then
                            sharingSpan = True
                            Exit For
                        End If
                    Next midspanSlot
                    If sharingSpan Then
                        key = key & 1
                    Else
                        For Each midspanSlot In comp.midspans
                            If Replace(comp.midspans(midspanSlot), "-", "") <> "" And Replace(commComponentsDict(key).midspans(midspanSlot), "-", "") = "" Then
                                commComponentsDict(key).midspans(midspanSlot) = comp.midspans(midspanSlot)
                            End If
                        Next midspanSlot
                        Exit Do
                    End If
                End If
            Loop While commComponentsDict.Exists(key)
            If Not commComponentsDict.Exists(key) Then commComponentsDict.Add key, comp
        End If
    Next i
   
    Set utilityComponents = Utilities.sortComponents(utilityComponentsDict)
    Set commComponents = Utilities.sortComponents(commComponentsDict)
   
    For i = 1 To 100
        If sheet.Range("UTTYPE").offset(i - 1, 0).Interior.color <> 16312794 Then Exit For
        If i <= utilityComponents.count Then
            Set comp = utilityComponents(i)
            displayHeight = Utilities.inchesToFeetInches(comp.height)
            If comp.bottomHeight > 0 Then displayHeight = displayHeight & "(" & Utilities.inchesToFeetInches(comp.bottomHeight) & ")"
            
            If sheet.Range("UTHEIGHT").offset(i - 1, 0).text <> displayHeight Then sheet.Range("UTHEIGHT").offset(i - 1, 0) = displayHeight
            If sheet.Range("UTTYPE").offset(i - 1, 0).text <> comp.componentType Then sheet.Range("UTTYPE").offset(i - 1, 0) = comp.componentType
            If sheet.Range("UTSIZE").offset(i - 1, 0).text <> comp.size Then sheet.Range("UTSIZE").offset(i - 1, 0) = comp.size
            For j = 1 To comp.midspans.count
                If sheet.Range("UTMIDSPAN" & j).offset(i - 1, 0).text <> comp.midspans(j) Then sheet.Range("UTMIDSPAN" & j).offset(i - 1, 0) = comp.midspans(j)
                If comp.midspans(j) <> "-" Then
                    If sheet.Range("UTMIDSPAN" & j).offset(i - 1, 0).HorizontalAlignment <> xlLeft Then sheet.Range("UTMIDSPAN" & j).offset(i - 1, 0).HorizontalAlignment = xlLeft
                Else
                    If sheet.Range("UTMIDSPAN" & j).offset(i - 1, 0).HorizontalAlignment <> xlCenter Then sheet.Range("UTMIDSPAN" & j).offset(i - 1, 0).HorizontalAlignment = xlCenter
                End If
            Next j
        Else
            If sheet.Range("UTHEIGHT").offset(i - 1, 0).text <> "" Then sheet.Range("UTHEIGHT").offset(i - 1, 0) = ""
            If sheet.Range("UTTYPE").offset(i - 1, 0).text <> "" Then sheet.Range("UTTYPE").offset(i - 1, 0) = ""
            If sheet.Range("UTSIZE").offset(i - 1, 0).text <> "" Then sheet.Range("UTSIZE").offset(i - 1, 0) = ""
            For j = 1 To comp.midspans.count
                If sheet.Range("UTMIDSPAN" & j).offset(i - 1, 0).text <> "" Then sheet.Range("UTMIDSPAN" & j).offset(i - 1, 0) = ""
                If sheet.Range("UTMIDSPAN" & j).offset(i - 1, 0).HorizontalAlignment <> xlCenter Then sheet.Range("UTMIDSPAN" & j).offset(i - 1, 0).HorizontalAlignment = xlCenter
            Next j
        End If
    Next i
    
    For i = 1 To 100
        If sheet.Range("CMOWNER").offset(i - 1, 0).Interior.color <> 16312794 Then Exit For
        If sheet.Range("CMOWNER").offset(i - 1, -1).Value <> CStr(i) Then sheet.Range("CMOWNER").offset(i - 1, -1).Value = i
        If i <= commComponents.count Then
            Set comp = commComponents(i)
            displayHeight = Utilities.inchesToFeetInches(comp.height)
            If comp.bottomHeight > 0 Then displayHeight = displayHeight & "(" & Utilities.inchesToFeetInches(comp.bottomHeight) & ")"
            
            If sheet.Range("CMHEIGHT").offset(i - 1, 0).Value <> displayHeight Then sheet.Range("CMHEIGHT").offset(i - 1, 0).Value = displayHeight
            If sheet.Range("CMOWNER").offset(i - 1, 0).Value <> comp.owner Then sheet.Range("CMOWNER").offset(i - 1, 0).Value = comp.owner
            If sheet.Range("CMSIZE").offset(i - 1, 0).Value <> comp.size Then sheet.Range("CMSIZE").offset(i - 1, 0) = comp.size
            For j = 1 To comp.midspans.count
                If sheet.Range("CMMIDSPAN" & j).offset(i - 1, 0).text <> comp.midspans(j) Then sheet.Range("CMMIDSPAN" & j).offset(i - 1, 0) = comp.midspans(j)
                If comp.midspans(j) <> "-" Then
                    If sheet.Range("CMMIDSPAN" & j).offset(i - 1, 0).HorizontalAlignment <> xlLeft Then sheet.Range("CMMIDSPAN" & j).offset(i - 1, 0).HorizontalAlignment = xlLeft
                Else
                    If sheet.Range("CMMIDSPAN" & j).offset(i - 1, 0).HorizontalAlignment <> xlCenter Then sheet.Range("CMMIDSPAN" & j).offset(i - 1, 0).HorizontalAlignment = xlCenter
                End If
            Next j
        Else
            If sheet.Range("CMHEIGHT").offset(i - 1, 0).text <> "" Then sheet.Range("CMHEIGHT").offset(i - 1, 0) = ""
            If sheet.Range("CMOWNER").offset(i - 1, 0).text <> "" Then sheet.Range("CMOWNER").offset(i - 1, 0) = ""
            If sheet.Range("CMSIZE").offset(i - 1, 0).text <> "" Then sheet.Range("CMSIZE").offset(i - 1, 0) = ""
            For j = 1 To comp.midspans.count
                If sheet.Range("CMMIDSPAN" & j).offset(i - 1, 0).text <> "" Then sheet.Range("CMMIDSPAN" & j).offset(i - 1, 0) = ""
                If sheet.Range("CMMIDSPAN" & j).offset(i - 1, 0).HorizontalAlignment <> xlCenter Then sheet.Range("CMMIDSPAN" & j).offset(i - 1, 0).HorizontalAlignment = xlCenter
            Next j
        End If
    Next i
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "Finished sorting attachments"
    
End Sub

Public Sub fixCommMakeReadyForm()
    On Error Resume Next

    Call LogMessage.SendLogMessage("fixCommMakeReadyForm")
    
    Dim sheet As Worksheet: Set sheet = ThisWorkbook.ActiveSheet()
    If Not Utilities.IsPDS(sheet) Then
        MsgBox "You need to have a pole detail sheet active to run this script."
        Exit Sub
    End If
    
    Dim protected As Boolean: protected = sheet.ProtectContents
    If protected Then sheet.Unprotect
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim pole As pole: Set pole = New pole
    Call pole.extractFromSheet(sheet)
    Call pole.clearCMRF(sheet)
    Call pole.fillCMRF(sheet)
    
    If protected Then sheet.Protect _
        Password:="", _
        DrawingObjects:=False, _
        contents:=True, _
        Scenarios:=False, _
        AllowFormattingCells:=True, _
        AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, _
        AllowInsertingColumns:=False, _
        AllowInsertingRows:=False, _
        AllowDeletingColumns:=False, _
        AllowDeletingRows:=False
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "Finished regenerating CMRF"
    
End Sub
