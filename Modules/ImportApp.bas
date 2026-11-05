Attribute VB_Name = "ImportApp"
Sub ImportApplication()
    Call LogMessage.SendLogMessage("ImportApplication")

    Set fileDiag = Application.FileDialog(msoFileDialogFilePicker)
    With fileDiag
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add Description:="Excel Files", Extensions:="*.xls,*.xlsx,*.csv"
        .Title = "Select the File ... "
        .InitialFileName = GetLocalPath(ThisWorkbook.Path) & Application.PathSeparator
        If .Show = -1 Then Path = Application.FileDialog(msoFileDialogFilePicker).SelectedItems.item(1) Else Exit Sub
    End With
    
    On Error Resume Next
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    ProgressBar_Form.Show vbModeless
    
    ProgressBar_Form.Label1.caption = "Importing Application... Please wait..."
    ProgressBar_Form.Repaint
    
    Dim applicationWb As Workbook
    Set applicationWb = Workbooks.Open(Path, False)
    If applicationWb Is Nothing Then
        ProgressBar_Form.Hide
        MsgBox "Either close or enable editing on the application"
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        Exit Sub
    End If
    Dim applicationWs As Worksheet: Set applicationWs = applicationWb.sheets(1)
    
    Dim i As Integer
    Dim applicationWsLastRow As Integer: applicationWsLastRow = applicationWs.cells(applicationWs.Rows.count, "A").End(xlUp).ROW
    Dim applicationWsLastCol As Integer: applicationWsLastCol = applicationWs.cells(1, applicationWs.Columns.count).End(xlToLeft).column
    Dim applicationWsHeaders As Scripting.Dictionary: Set applicationWsHeaders = New Scripting.Dictionary
    For i = 1 To applicationWsLastCol
        header = Trim(applicationWs.cells(1, i).value)
        If header <> "" Then
            applicationWsHeaders(header) = i
        End If
    Next i
    
    Dim sheet As Worksheet
    Dim sheetName As String
    Dim cell As Variant
    
    Call clearApp
    
    For i = 2 To applicationWsLastRow
        sheetName = CStr(applicationWs.cells(i, applicationWsHeaders("POLE NUMBER")))
        toPoleSheetName = CStr(applicationWs.cells(i, applicationWsHeaders("TO POLE")))
        Dim added As Boolean
        If SheetExists(sheetName) Then
            Set sheet = Utilities.GetPDS(sheetName)
        
            If applicationWsHeaders.exists("PROPOSED ATT. HEIGHT") Then
                sheet.Range("PROPOSEDHEIGHT").value = CleanFeetInches(applicationWs.cells(i, applicationWsHeaders("PROPOSED ATT. HEIGHT")))
            End If
            If (applicationWsHeaders.exists("OL ATT. HEIGHT")) Then
                If UCase(applicationWs.cells(i, applicationWsHeaders("OL ATT. HEIGHT")).value) <> "NA" And UCase(applicationWs.cells(i, applicationWsHeaders("OL ATT. HEIGHT")).value) <> "N/A" And Trim(applicationWs.cells(i, applicationWsHeaders("OL ATT. HEIGHT")).value) <> "" Then
                    sheet.Range("PROPOSEDHEIGHT").value = sheet.Range("PROPOSEDHEIGHT").value & " OL"
                End If
            End If
            If (applicationWsHeaders.exists("GUY SIZE")) Then
                If (UCase(applicationWs.cells(i, applicationWsHeaders("GUY SIZE"))) <> "NA" And UCase(applicationWs.cells(i, applicationWsHeaders("GUY SIZE"))) <> "N/A") Then
                    sheet.Range("NEWAPPSIZE").value = applicationWs.cells(i, applicationWsHeaders("GUY SIZE"))
                End If
            End If
            If (applicationWsHeaders.exists("GUY LEAD")) Then
                If (UCase(applicationWs.cells(i, applicationWsHeaders("GUY LEAD"))) <> "NA" And UCase(applicationWs.cells(i, applicationWsHeaders("GUY LEAD"))) <> "N/A") Then
                    sheet.Range("NEWAPPLEAD").value = applicationWs.cells(i, applicationWsHeaders("GUY LEAD"))
                End If
            End If
            If (applicationWsHeaders.exists("GUY DIRECTION")) Then
                If (UCase(applicationWs.cells(i, applicationWsHeaders("GUY DIRECTION"))) <> "NA" And UCase(applicationWs.cells(i, applicationWsHeaders("GUY DIRECTION"))) <> "N/A") Then
                    sheet.Range("NEWAPPDIR").value = applicationWs.cells(i, applicationWsHeaders("GUY DIRECTION"))
                End If
            End If
            
            If applicationWsHeaders.exists("EXISTING DIAMETER") Then
                sheet.Range("EXISTINGDIAMETER").value = applicationWs.cells(i, applicationWsHeaders("EXISTING DIAMETER"))
            End If
            If applicationWsHeaders.exists("DIAMETER") Then
                sheet.Range("PROPOSEDDIAMETER").value = applicationWs.cells(i, applicationWsHeaders("DIAMETER"))
            End If
            
            If applicationWsHeaders.exists("ADDITIONAL SPANS") Then
                If applicationWs.cells(i, applicationWsHeaders("ADDITIONAL SPANS")).value <> "" Then
                    If InStr(sheet.Range("NOTES").value, "ADDITIONAL SPANS: " & applicationWs.cells(i, applicationWsHeaders("ADDITIONAL SPANS"))) = 0 Then
                        If sheet.Range("NOTES").value <> "" Then
                            sheet.Range("NOTES").value = sheet.Range("NOTES").value & vbLf
                        End If
                        sheet.Range("NOTES").value = sheet.Range("NOTES").value & "ADDITIONAL SPANS: " & applicationWs.cells(i, applicationWsHeaders("ADDITIONAL SPANS"))
                    End If
                End If
            End If
            
            If applicationWsHeaders.exists("COMMENTS") Then
                If applicationWs.cells(i, applicationWsHeaders("COMMENTS")).value <> "" Then
                    If InStr(sheet.Range("NOTES").value, applicationWs.cells(i, applicationWsHeaders("COMMENTS"))) = 0 Then
                        If sheet.Range("NOTES").value <> "" Then
                            sheet.Range("NOTES").value = sheet.Range("NOTES").value & vbLf
                        End If
                        sheet.Range("NOTES").value = sheet.Range("NOTES").value & "APPLICANT COMMENTS: " & applicationWs.cells(i, applicationWsHeaders("COMMENTS"))
                    End If
                End If
            End If
            
            For j = 2 To applicationWsLastRow
                If sheetName = CStr(applicationWs.cells(j, applicationWsHeaders("TO POLE"))) Then
                    fromPoleSheetName = CStr(applicationWs.cells(j, applicationWsHeaders("POLE NUMBER")))
                    For k = 1 To 12
                        If Utilities.RangeExists(sheet, "TOPOLE" & k) Then
                            fromPoleNumber = ThisWorkbook.RemoveParentheses(sheet.Range("TOPOLE" & k).value)
                            If CStr(fromPoleNumber) = CStr(fromPoleSheetName) Then
                                sheet.Range("TOPOLE" & k).offSet(1, 0).value = CleanFeetInches(applicationWs.cells(j, applicationWsHeaders("MIDSPAN")))
                                sheet.Range("TOPOLE" & k).offSet(2, 0).value = applicationWs.cells(j, applicationWsHeaders("TENSION"))
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next k
                End If
            Next j
            
            added = False
            For j = 1 To 12
                If Utilities.RangeExists(sheet, "TOPOLE" & j) Then
                    toPoleNumber = ThisWorkbook.RemoveParentheses(sheet.Range("TOPOLE" & j).value)
                    If CStr(toPoleNumber) = CStr(toPoleSheetName) Then
                        sheet.Range("TOPOLE" & j).offSet(1, 0).value = CleanFeetInches(applicationWs.cells(i, applicationWsHeaders("MIDSPAN")))
                        sheet.Range("TOPOLE" & j).offSet(2, 0).value = applicationWs.cells(i, applicationWsHeaders("TENSION"))
                        added = True
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next j
            
            If Not added Then
                For j = 1 To 12
                    If Utilities.RangeExists(sheet, "TOPOLE" & j) Then
                        toPoleNumber = ThisWorkbook.RemoveParentheses(sheet.Range("TOPOLE" & j).value)
                        If sheet.Range("TOPOLE" & j).offSet(1, 0).value = "" Then
                            If Not Utilities.SheetExists(CStr(toPoleNumber)) Or Not Utilities.RangeExists(sheet, "TOPOLE" & j + 1) Then
                                sheet.Range("TOPOLE" & j).offSet(1, 0).value = CleanFeetInches(applicationWs.cells(i, applicationWsHeaders("MIDSPAN"))) & " (GUESS)"
                                sheet.Range("TOPOLE" & j).offSet(2, 0).value = applicationWs.cells(i, applicationWsHeaders("TENSION"))
                                Exit For
                            End If
                        End If
                    Else
                        Exit For
                    End If
                Next j
            End If
        End If
    Next i
    
    applicationWb.Close savechanges:=False
    
    ProgressBar_Form.Hide
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub

Private Sub clearApp()
    Dim sheet As Worksheet
    For Each sheet In ThisWorkbook.sheets
        If Utilities.IsPDS(sheet) Then
            For i = 1 To 12
                If Utilities.RangeExists(sheet, "TOPOLE" & i) Then
                    sheet.Range("TOPOLE" & i).offSet(1, 0).value = ""
                    sheet.Range("TOPOLE" & i).offSet(2, 0).value = ""
                Else
                    Exit For
                End If
            Next i
            sheet.Range("NEWAPPSIZE").value = ""
            sheet.Range("NEWAPPSIZE").offSet(1, 0).value = ""
            sheet.Range("NEWAPPLEAD").value = ""
            sheet.Range("NEWAPPLEAD").offSet(1, 0).value = ""
            sheet.Range("NEWAPPDIR").value = ""
            sheet.Range("NEWAPPDIR").offSet(1, 0).value = ""
            sheet.Range("PROPOSEDHEIGHT").value = ""
            sheet.Range("PROPOSEDDIAMETER").value = ""
            sheet.Range("EXISTINGDIAMETER").value = ""
        End If
    Next sheet
End Sub

Private Function CleanFeetInches(inputStr As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .IgnoreCase = True
        ' Match optional space + foot symbol + optional space + optional dash + optional space + inch digits + optional inch symbol
        .Pattern = "(\d+)\s*['’`]\s*[-]?\s*(\d+)\s*(?:[""”]|''){0,1}"
    End With
    If regex.Test(inputStr) Then
        CleanFeetInches = regex.Replace(inputStr, "$1'$2""")
    Else
        CleanFeetInches = inputStr
    End If
End Function
