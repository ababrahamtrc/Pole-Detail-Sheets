Attribute VB_Name = "ImportApp"
Sub ImportApplication()
    Call LogMessage.SendLogMessage("ImportApplication")

    Set fileDiag = Application.FileDialog(msoFileDialogFilePicker)
    With fileDiag
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add Description:="Excel Files", Extensions:="*.xls,*.xlsx,*.csv"
        .Title = "Select the File ... "
        .InitialFileName = ThisWorkbook.path & Application.PathSeparator
        If .Show = -1 Then path = Application.FileDialog(msoFileDialogFilePicker).SelectedItems.item(1) Else Exit Sub
    End With
    
    On Error Resume Next
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    ProgressBar_Form.Show vbModeless
    
    ProgressBar_Form.Label1.caption = "Importing Application... Please wait..."
    ProgressBar_Form.Repaint
    
    Dim applicationWb As Workbook
    Set applicationWb = Workbooks.Open(path, False)
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
    Dim applicationWsLastRow As Integer: applicationWsLastRow = applicationWs.Cells(applicationWs.Rows.count, "A").End(xlUp).row
    Dim applicationWsLastCol As Integer: applicationWsLastCol = applicationWs.Cells(1, applicationWs.Columns.count).End(xlToLeft).Column
    Dim applicationWsHeaders As Scripting.Dictionary: Set applicationWsHeaders = New Scripting.Dictionary
    For i = 1 To applicationWsLastCol
        header = Trim(applicationWs.Cells(1, i).Value)
        If header <> "" Then
            applicationWsHeaders(header) = i
        End If
    Next i
    
    Dim sheet As Worksheet
    Dim sheetName As String
    Dim cell As Variant
    
    Call clearApp
    
    For i = 2 To applicationWsLastRow
        sheetName = CStr(applicationWs.Cells(i, applicationWsHeaders("POLE NUMBER")))
        toPoleSheetName = CStr(applicationWs.Cells(i, applicationWsHeaders("TO POLE")))
        Dim added As Boolean
        If SheetExists(sheetName) Then
            Set sheet = Utilities.GetPDS(sheetName)
        
            If applicationWsHeaders.Exists("PROPOSED ATT. HEIGHT") Then
                sheet.Range("PROPOSEDHEIGHT").Value = CleanFeetInches(applicationWs.Cells(i, applicationWsHeaders("PROPOSED ATT. HEIGHT")))
            End If
            If (applicationWsHeaders.Exists("OL ATT. HEIGHT")) Then
                If UCase(applicationWs.Cells(i, applicationWsHeaders("OL ATT. HEIGHT")).Value) <> "NA" And UCase(applicationWs.Cells(i, applicationWsHeaders("OL ATT. HEIGHT")).Value) <> "N/A" And Trim(applicationWs.Cells(i, applicationWsHeaders("OL ATT. HEIGHT")).Value) <> "" Then
                    sheet.Range("PROPOSEDHEIGHT").Value = sheet.Range("PROPOSEDHEIGHT").Value & " OL"
                End If
            End If
            If (applicationWsHeaders.Exists("GUY SIZE")) Then
                If (UCase(applicationWs.Cells(i, applicationWsHeaders("GUY SIZE"))) <> "NA" And UCase(applicationWs.Cells(i, applicationWsHeaders("GUY SIZE"))) <> "N/A") Then
                    sheet.Range("NEWAPPSIZE").Value = applicationWs.Cells(i, applicationWsHeaders("GUY SIZE"))
                End If
            End If
            If (applicationWsHeaders.Exists("GUY LEAD")) Then
                If (UCase(applicationWs.Cells(i, applicationWsHeaders("GUY LEAD"))) <> "NA" And UCase(applicationWs.Cells(i, applicationWsHeaders("GUY LEAD"))) <> "N/A") Then
                    sheet.Range("NEWAPPLEAD").Value = applicationWs.Cells(i, applicationWsHeaders("GUY LEAD"))
                End If
            End If
            If (applicationWsHeaders.Exists("GUY DIRECTION")) Then
                If (UCase(applicationWs.Cells(i, applicationWsHeaders("GUY DIRECTION"))) <> "NA" And UCase(applicationWs.Cells(i, applicationWsHeaders("GUY DIRECTION"))) <> "N/A") Then
                    sheet.Range("NEWAPPDIR").Value = applicationWs.Cells(i, applicationWsHeaders("GUY DIRECTION"))
                End If
            End If
            
            If applicationWsHeaders.Exists("EXISTING DIAMETER") Then
                sheet.Range("EXISTINGDIAMETER").Value = applicationWs.Cells(i, applicationWsHeaders("EXISTING DIAMETER"))
            End If
            If applicationWsHeaders.Exists("DIAMETER") Then
                sheet.Range("PROPOSEDDIAMETER").Value = applicationWs.Cells(i, applicationWsHeaders("DIAMETER"))
            End If
            
            If applicationWsHeaders.Exists("ADDITIONAL SPANS") Then
                If applicationWs.Cells(i, applicationWsHeaders("ADDITIONAL SPANS")).Value <> "" Then
                    If InStr(sheet.Range("NOTES").Value, "ADDITIONAL SPANS: " & applicationWs.Cells(i, applicationWsHeaders("ADDITIONAL SPANS"))) = 0 Then
                        If sheet.Range("NOTES").Value <> "" Then
                            sheet.Range("NOTES").Value = sheet.Range("NOTES").Value & vbLf
                        End If
                        sheet.Range("NOTES").Value = sheet.Range("NOTES").Value & "ADDITIONAL SPANS: " & applicationWs.Cells(i, applicationWsHeaders("ADDITIONAL SPANS"))
                    End If
                End If
            End If
            
            If applicationWsHeaders.Exists("COMMENTS") Then
                If applicationWs.Cells(i, applicationWsHeaders("COMMENTS")).Value <> "" Then
                    If InStr(sheet.Range("NOTES").Value, applicationWs.Cells(i, applicationWsHeaders("COMMENTS"))) = 0 Then
                        If sheet.Range("NOTES").Value <> "" Then
                            sheet.Range("NOTES").Value = sheet.Range("NOTES").Value & vbLf
                        End If
                        sheet.Range("NOTES").Value = sheet.Range("NOTES").Value & "APPLICANT COMMENTS: " & applicationWs.Cells(i, applicationWsHeaders("COMMENTS"))
                    End If
                End If
            End If
            
            For j = 2 To applicationWsLastRow
                If sheetName = CStr(applicationWs.Cells(j, applicationWsHeaders("TO POLE"))) Then
                    fromPoleSheetName = CStr(applicationWs.Cells(j, applicationWsHeaders("POLE NUMBER")))
                    For k = 1 To 12
                        If Utilities.RangeExists(sheet, "TOPOLE" & k) Then
                            fromPoleNumber = ThisWorkbook.RemoveParentheses(sheet.Range("TOPOLE" & k).Value)
                            If CStr(fromPoleNumber) = CStr(fromPoleSheetName) Then
                                sheet.Range("TOPOLE" & k).offset(1, 0).Value = CleanFeetInches(applicationWs.Cells(j, applicationWsHeaders("MIDSPAN")))
                                sheet.Range("TOPOLE" & k).offset(2, 0).Value = applicationWs.Cells(j, applicationWsHeaders("TENSION"))
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
                    toPoleNumber = ThisWorkbook.RemoveParentheses(sheet.Range("TOPOLE" & j).Value)
                    If CStr(toPoleNumber) = CStr(toPoleSheetName) Then
                        sheet.Range("TOPOLE" & j).offset(1, 0).Value = CleanFeetInches(applicationWs.Cells(i, applicationWsHeaders("MIDSPAN")))
                        sheet.Range("TOPOLE" & j).offset(2, 0).Value = applicationWs.Cells(i, applicationWsHeaders("TENSION"))
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
                        toPoleNumber = ThisWorkbook.RemoveParentheses(sheet.Range("TOPOLE" & j).Value)
                        If sheet.Range("TOPOLE" & j).offset(1, 0).Value = "" Then
                            If Not Utilities.SheetExists(CStr(toPoleNumber)) Or Not Utilities.RangeExists(sheet, "TOPOLE" & j + 1) Then
                                sheet.Range("TOPOLE" & j).offset(1, 0).Value = CleanFeetInches(applicationWs.Cells(i, applicationWsHeaders("MIDSPAN"))) & " (GUESS)"
                                sheet.Range("TOPOLE" & j).offset(2, 0).Value = applicationWs.Cells(i, applicationWsHeaders("TENSION"))
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
    
    applicationWb.Close SaveChanges:=False
    
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
                    sheet.Range("TOPOLE" & i).offset(1, 0).Value = ""
                    sheet.Range("TOPOLE" & i).offset(2, 0).Value = ""
                Else
                    Exit For
                End If
            Next i
            sheet.Range("NEWAPPSIZE").Value = ""
            sheet.Range("NEWAPPSIZE").offset(1, 0).Value = ""
            sheet.Range("NEWAPPLEAD").Value = ""
            sheet.Range("NEWAPPLEAD").offset(1, 0).Value = ""
            sheet.Range("NEWAPPDIR").Value = ""
            sheet.Range("NEWAPPDIR").offset(1, 0).Value = ""
            sheet.Range("PROPOSEDHEIGHT").Value = ""
            sheet.Range("PROPOSEDDIAMETER").Value = ""
            sheet.Range("EXISTINGDIAMETER").Value = ""
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
    If regex.test(inputStr) Then
        CleanFeetInches = regex.Replace(inputStr, "$1'$2""")
    Else
        CleanFeetInches = inputStr
    End If
End Function
