Attribute VB_Name = "ImportPDS"
Sub ImportAllPDS()
    Call LogMessage.SendLogMessage("ImportPDS")

    Dim fileDiag As FileDialog
    Set fileDiag = Application.FileDialog(msoFileDialogFolderPicker)
    With fileDiag
        .AllowMultiSelect = False
        .Title = "Select a folder that contains the pole detail sheets"
        .InitialFileName = ThisWorkbook.path & Application.PathSeparator
        If .Show = -1 Then path = Application.FileDialog(msoFileDialogFilePicker).SelectedItems.item(1) & Application.PathSeparator Else Exit Sub
    End With
    
    On Error Resume Next
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    Call NJUNSCodes.clearNJUNSCodes
    
    ProgressBar_Form.Show vbModeless
    
    ProgressBar_Form.Label1.caption = "Importing Pole Detail Sheets... Please wait..."
    ProgressBar_Form.Repaint


    Dim poleNumbers As scripting.Dictionary: Set poleNumbers = New scripting.Dictionary
    Dim duplicateNumbers As scripting.Dictionary: Set duplicateNumbers = New scripting.Dictionary
    Dim duplicateNumbersString As String: duplicateNumbersString = ""
    Dim sheet As Worksheet
    Dim sheetCount As Integer: sheetCount = 0
    Dim fileName As String: fileName = Dir(path & "*.xlsx")
    Dim sourceWb As Workbook
    Dim insertSpot As Worksheet
    Do While fileName <> ""
        ProgressBar_Form.Label1.caption = "Importing Pole Detail Sheets... " & sheetCount & " sheets imported."
        ProgressBar_Form.Repaint
        
        Set sourceWb = Workbooks.Open(path & fileName)
        If sourceWb.sheets(1).Cells(2, 2).Value = "Notification:" Then
            rpSourceName = ThisWorkbook.RemoveParentheses(sourceWb.sheets(1).name)
            If Not poleNumbers.Exists(rpSourceName) Then
                poleNumbers.Add rpSourceName, Nothing
            Else
                If Not duplicateNumbers.Exists(rpSourceName) Then
                    duplicateNumbers.Add rpSourceName, Nothing
                    If duplicateNumbersString <> "" Then duplicateNumbersString = duplicateNumbersString & vbLf
                    duplicateNumbersString = duplicateNumbersString & "Pole: " & rpSourceName
                End If
            End If
            If Not SheetExists(rpSourceName) Then
                Set insertSpot = Nothing
                For Each sheet In ThisWorkbook.sheets
                    rpName = ThisWorkbook.RemoveParentheses(sheet.name)
                    If rpName <> "4 Spans" And rpName <> "8 Spans" And rpName <> "12 Spans" Then
                        If sheet.Cells(2, 2).Value = "Notification:" Then
                            If val(rpName) > val(rpSourceName) Or (val(rpName) = val(rpSourceName) And rpName > rpSourceName) Then
                                Set insertSpot = sheet
                                Exit For
                            End If
                        End If
                    End If
                Next sheet
                
                Call ThisWorkbook.decideTabColor(sourceWb.sheets(1))
                
                If insertSpot Is Nothing Then
                    sourceWb.sheets(1).Copy after:=ThisWorkbook.sheets(ThisWorkbook.sheets.count)
                Else
                    sourceWb.sheets(1).Copy Before:=insertSpot
                End If
                sheetCount = sheetCount + 1
            End If
        End If
        sourceWb.Close SaveChanges:=False
        fileName = Dir
    Loop
    ThisWorkbook.sheets("Control").Activate
    
    Application.ScreenUpdating = True
    
    For Each sheet In ThisWorkbook.sheets
        If sheet.name <> "4 Spans" And sheet.name <> "8 Spans" And sheet.name <> "12 Spans" Then
            If sheet.Cells(2, 2).Value = "Notification:" Then
                sheet.Activate
                Call Figures.getSheetFigures(sheet)
                sheet.Range("A1").Select
            End If
        End If
    Next sheet
    
    ThisWorkbook.sheets("Control").Activate
    
    ProgressBar_Form.Hide
    
    ThisWorkbook.sheets("Control").Range("PHOTODIR").Value = ""
    
    If duplicateNumbersString <> "" Then MsgBox "The following pole numbers have duplicate files in the selected folder, please make sure the correct one was imported and that they are given differen't pole numbers, or delete the one that's not needed." & vbLf & duplicateNumbersString
    
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub

Sub ImportSinglePDS()
    Call LogMessage.SendLogMessage("ImportSinglePDS")

    Dim fileDiag As FileDialog
    Set fileDiag = Application.FileDialog(msoFileDialogFilePicker)
    With fileDiag
        .AllowMultiSelect = False
        .Title = "Select a pole detail sheet"
        .Filters.Add "Pole Detail Sheet Files", "*.xls; *.xlsx, *.xlsm; *.xlsb", 1
        .InitialFileName = ThisWorkbook.path & Application.PathSeparator
        If .Show = -1 Then path = Application.FileDialog(msoFileDialogFilePicker).SelectedItems.item(1) Else Exit Sub
    End With
    
    On Error Resume Next
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    Call NJUNSCodes.clearNJUNSCodes
    
    ProgressBar_Form.Show vbModeless
    
    ProgressBar_Form.Label1.caption = "Importing Pole Detail Sheet... Please wait..."
    ProgressBar_Form.Repaint

    Dim sourceWb As Workbook: Set sourceWb = Workbooks.Open(path)
    Dim newSheet As Worksheet: Set newSheet = Nothing
    Dim insertSpot As Worksheet
    If sourceWb.sheets(1).Cells(2, 2).Value = "Notification:" Then
        rpSourceName = ThisWorkbook.RemoveParentheses(sourceWb.sheets(1).name)
        If Not SheetExists(rpSourceName) Then
            Set insertSpot = Nothing
            For Each sheet In ThisWorkbook.sheets
                rpName = ThisWorkbook.RemoveParentheses(sheet.name)
                If rpName <> "4 Spans" And rpName <> "8 Spans" And rpName <> "12 Spans" Then
                    If sheet.Cells(2, 2).Value = "Notification:" Then
                        If val(rpName) > val(rpSourceName) Or (val(rpName) = val(rpSourceName) And rpName > rpSourceName) Then
                            Set insertSpot = sheet
                            Exit For
                        End If
                    End If
                End If
            Next sheet
            
            Call ThisWorkbook.decideTabColor(sourceWb.sheets(1))
            Call Figures.getSheetFigures(sourceWb.sheets(1))
            
            If insertSpot Is Nothing Then
                sourceWb.sheets(1).Copy after:=ThisWorkbook.sheets(ThisWorkbook.sheets.count)
            Else
                sourceWb.sheets(1).Copy Before:=insertSpot
            End If
        End If
    End If
    
    sourceWb.Close SaveChanges:=False
    'ThisWorkBook.sheets("Control").Activate
    
    ProgressBar_Form.Hide
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub

Sub ImportLocationPDS()
    Dim fileDiag As FileDialog
    Set fileDiag = Application.FileDialog(msoFileDialogFolderPicker)
    With fileDiag
        .AllowMultiSelect = False
        .Title = "Select a folder that contains the pole detail sheets"
        .InitialFileName = ThisWorkbook.path & Application.PathSeparator
        If .Show = -1 Then path = Application.FileDialog(msoFileDialogFilePicker).SelectedItems.item(1) & Application.PathSeparator Else Exit Sub
    End With
    
    On Error Resume Next
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    ProgressBar_Form.Show vbModeless
    
    ProgressBar_Form.Label1.caption = "Importing Pole Detail Sheets... Please wait..."
    ProgressBar_Form.Repaint
    
    Dim poleNumbers As scripting.Dictionary: Set poleNumbers = New scripting.Dictionary
    Dim duplicateNumbers As scripting.Dictionary: Set duplicateNumbers = New scripting.Dictionary
    Dim duplicateNumbersString As String: duplicateNumbersString = ""
    
    Dim sheetCount As Integer: sheetCount = 0
    Dim fileName As String: fileName = Dir(path & "*.xlsx")
    Dim sourceWb As Workbook
    Dim insertSpot As Worksheet
    Do While fileName <> ""
        ProgressBar_Form.Label1.caption = "Importing Pole Detail Sheets... " & sheetCount & " sheets imported."
        ProgressBar_Form.Repaint
        Set sourceWb = Workbooks.Open(path & fileName)
        If sourceWb.sheets(1).Cells(2, 2).Value = "Notification:" Then
            If Trim(sourceWb.sheets(1).Range("DL").Value) <> "" Then
                rpSourceName = ThisWorkbook.RemoveParentheses(sourceWb.sheets(1).name)
                If Not poleNumbers.Exists(rpSourceName) Then
                    poleNumbers.Add rpSourceName, Nothing
                Else
                    If Not duplicateNumbers.Exists(rpSourceName) Then
                        duplicateNumbers.Add rpSourceName, Nothing
                        If duplicateNumbersString <> "" Then duplicateNumbersString = duplicateNumbersString & vbLf
                        duplicateNumbersString = duplicateNumbersString & "Pole: " & rpSourceName
                    End If
                End If
                If Not SheetExists(rpSourceName) Then
                    Set insertSpot = Nothing
                    For Each sheet In ThisWorkbook.sheets
                        rpName = ThisWorkbook.RemoveParentheses(sheet.name)
                        If rpName <> "4 Spans" And rpName <> "8 Spans" And rpName <> "12 Spans" Then
                            If sheet.Cells(2, 2).Value = "Notification:" Then
                                If val(rpName) > val(rpSourceName) Or (val(rpName) = val(rpSourceName) And rpName > rpSourceName) Then
                                    Set insertSpot = sheet
                                    Exit For
                                End If
                            End If
                        End If
                    Next sheet
                    
                    Call ThisWorkbook.decideTabColor(sourceWb.sheets(1))
                    Call Figures.getSheetFigures(sourceWb.sheets(1))
                    
                    If insertSpot Is Nothing Then
                        sourceWb.sheets(1).Copy after:=ThisWorkbook.sheets(ThisWorkbook.sheets.count)
                    Else
                        sourceWb.sheets(1).Copy Before:=insertSpot
                    End If
                    sheetCount = sheetCount + 1
                End If
            End If
        End If
        sourceWb.Close SaveChanges:=False
        fileName = Dir
    Loop
    ThisWorkbook.sheets("Control").Activate
    
    Application.ScreenUpdating = True
    
    For Each sheet In ThisWorkbook.sheets
        If sheet.name <> "4 Spans" And sheet.name <> "8 Spans" And sheet.name <> "12 Spans" Then
            If sheet.Cells(2, 2).Value = "Notification:" Then
                sheet.Activate
                sheet.Range("A1").Select
            End If
        End If
    Next sheet
    
    ThisWorkbook.sheets("Control").Activate
    
    ProgressBar_Form.Hide
    
    If duplicateNumbersString <> "" Then MsgBox "The following pole numbers have duplicate files in the selected folder, please make sure the correct one was imported and that they are given differen't pole numbers, or delete the one that's not needed." & vbLf & duplicateNumbersString
    
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub

Sub ImportNjunsPDS()
    Dim fileDiag As FileDialog
    Set fileDiag = Application.FileDialog(msoFileDialogFolderPicker)
    With fileDiag
        .AllowMultiSelect = False
        .Title = "Select a folder that contains the pole detail sheets"
        .InitialFileName = ThisWorkbook.path & Application.PathSeparator
        If .Show = -1 Then path = Application.FileDialog(msoFileDialogFilePicker).SelectedItems.item(1) & Application.PathSeparator Else Exit Sub
    End With
    
    On Error Resume Next
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    ProgressBar_Form.Show vbModeless
    
    ProgressBar_Form.Label1.caption = "Importing Pole Detail Sheets... Please wait..."
    ProgressBar_Form.Repaint
    
    Dim poleNumbers As scripting.Dictionary: Set poleNumbers = New scripting.Dictionary
    Dim duplicateNumbers As scripting.Dictionary: Set duplicateNumbers = New scripting.Dictionary
    Dim duplicateNumbersString As String: duplicateNumbersString = ""
    
    Dim sheetCount As Integer: sheetCount = 0
    Dim fileName As String: fileName = Dir(path & "*.xlsx")
    Dim sourceWb As Workbook
    Dim insertSpot As Worksheet
    Do While fileName <> ""
        ProgressBar_Form.Label1.caption = "Importing Pole Detail Sheets... " & sheetCount & " sheets imported."
        ProgressBar_Form.Repaint
        Set sourceWb = Workbooks.Open(path & fileName)
        If sourceWb.sheets(1).Cells(2, 2).Value = "Notification:" Then
            If (Trim(sourceWb.sheets(1).Range("NJUNS").Value) <> "" And LCase(Trim(sourceWb.sheets(1).Range("NJUNS").Value)) <> "n/a" And LCase(Trim(sourceWb.sheets(1).Range("NJUNS").Value)) <> "comm make ready work") Or _
                (Trim(sourceWb.sheets(1).Range("NJUNSTICKET").Value) <> "" And LCase(Trim(sourceWb.sheets(1).Range("NJUNSTICKET").Value)) <> "n/a") Then
                rpSourceName = ThisWorkbook.RemoveParentheses(sourceWb.sheets(1).name)
                If Not poleNumbers.Exists(rpSourceName) Then
                    poleNumbers.Add rpSourceName, Nothing
                Else
                    If Not duplicateNumbers.Exists(rpSourceName) Then
                        duplicateNumbers.Add rpSourceName, Nothing
                        If duplicateNumbersString <> "" Then duplicateNumbersString = duplicateNumbersString & vbLf
                        duplicateNumbersString = duplicateNumbersString & "Pole: " & rpSourceName
                    End If
                End If
                If Not SheetExists(rpSourceName) Then
                    Set insertSpot = Nothing
                    For Each sheet In ThisWorkbook.sheets
                        rpName = ThisWorkbook.RemoveParentheses(sheet.name)
                        If rpName <> "4 Spans" And rpName <> "8 Spans" And rpName <> "12 Spans" Then
                            If sheet.Cells(2, 2).Value = "Notification:" Then
                                If val(rpName) > val(rpSourceName) Or (val(rpName) = val(rpSourceName) And rpName > rpSourceName) Then
                                    Set insertSpot = sheet
                                    Exit For
                                End If
                            End If
                        End If
                    Next sheet
                    
                    Call ThisWorkbook.decideTabColor(sourceWb.sheets(1))
                    
                    If insertSpot Is Nothing Then
                        sourceWb.sheets(1).Copy after:=ThisWorkbook.sheets(ThisWorkbook.sheets.count)
                    Else
                        sourceWb.sheets(1).Copy Before:=insertSpot
                    End If
                    sheetCount = sheetCount + 1
                End If
            End If
        End If
        sourceWb.Close SaveChanges:=False
        fileName = Dir
    Loop
    ThisWorkbook.sheets("Control").Activate
    
    Application.ScreenUpdating = True
    
    For Each sheet In ThisWorkbook.sheets
        If sheet.name <> "4 Spans" And sheet.name <> "8 Spans" And sheet.name <> "12 Spans" Then
            If sheet.Cells(2, 2).Value = "Notification:" Then
                sheet.Activate
                sheet.Range("A1").Select
            End If
        End If
    Next sheet
    
    ThisWorkbook.sheets("Control").Activate
    
    ProgressBar_Form.Hide
    
    If duplicateNumbersString <> "" Then MsgBox "The following pole numbers have duplicate files in the selected folder, please make sure the correct one was imported and that they are given differen't pole numbers, or delete the one that's not needed." & vbLf & duplicateNumbersString
    
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub

