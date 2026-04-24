Attribute VB_Name = "ExportPDS"
Sub ExportAllPDS()
    Call LogMessage.SendLogMessage("ExportPDS")

    Dim fileDiag As FileDialog
    Set fileDiag = Application.FileDialog(msoFileDialogFolderPicker)
    With fileDiag
        .AllowMultiSelect = False
        .Title = "Select a folder "
        .InitialFileName = ThisWorkbook.path & Application.PathSeparator
        If .Show = -1 Then outputPath = Application.FileDialog(msoFileDialogFilePicker).SelectedItems.item(1) & Application.PathSeparator Else Exit Sub
    End With
    
    Dim sheet As Worksheet
    
    On Error Resume Next
    
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    ProgressBar_Form.Show vbModeless
    
    ProgressBar_Form.Label1.caption = "Exporting Pole Detail Sheets... Please wait..."
    ProgressBar_Form.Repaint

    Application.ScreenUpdating = True
    
    Dim sheetTotal As Integer: sheetTotal = 0
    For Each sheet In ThisWorkbook.sheets
        If sheet.name <> "4 Spans" And sheet.name <> "8 Spans" And sheet.name <> "12 Spans" Then
            If sheet.Cells(2, 2).Value = "Notification:" Then
                sheetTotal = sheetTotal + 1
                sheet.Activate
                sheet.Range("A1").Select
            End If
        End If
    Next sheet
    
    ThisWorkbook.sheets("Control").Activate
    
    Application.ScreenUpdating = False

    Dim sheetCount As Integer: sheetCount = 0
    Dim tempWb As Workbook
    Dim fileName As String
    Dim fullPath As String
    For Each sheet In ThisWorkbook.sheets
        If sheet.name <> "4 Spans" And sheet.name <> "8 Spans" And sheet.name <> "12 Spans" And sheet.Cells(2, 2).Value = "Notification:" Then
            fullPath = "": fileName = ""
            fileName = "M1P" & fixIllegalCharacters(sheet.Range("POLENUM").text) & "_" & fixIllegalCharacters(sheet.Range("CEID").text) & "_" & fixIllegalCharacters(sheet.Range("PERMIT").text)
            fullPath = outputPath & fileName & ".xlsx"
            Call Figures.clearSheetFigures(sheet)
            sheet.Copy
            
            For Each wb In Application.Workbooks
                If wb.FullName = fullPath Then wb.Close SaveChanges:=True
            Next wb
            
            Set tempWb = ActiveWorkbook
            
            tempWb.SaveAs fileName:=fullPath, FileFormat:=xlOpenXMLWorkbook
            tempWb.Close SaveChanges:=True
            
            sheetCount = sheetCount + 1
            ProgressBar_Form.Label1.caption = "Exporting Pole Detail Sheets... " & sheetCount & " sheets exported.."
            ProgressBar_Form.Repaint
            
        End If
    Next sheet
    
    ProgressBar_Form.Hide
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub

Private Function fixIllegalCharacters(str As String) As String
    str = Replace(str, "<", "")
    str = Replace(str, ">", "")
    str = Replace(str, ":", "")
    str = Replace(str, """", "")
    str = Replace(str, "/", "")
    str = Replace(str, "\", "")
    str = Replace(str, "|", "")
    str = Replace(str, "?", "")
    str = Replace(str, "*", "")
    
    fixIllegalCharacters = str
End Function

Sub ExportSinglePDS()

    Dim sheet As Worksheet

    Call LogMessage.SendLogMessage("ExportSinglePDS")
    
    Set sheet = Application.ActiveSheet()
    If sheet.name = "4 Spans" Or sheet.name = "8 Spans" Or sheet.name = "12 Spans" Or sheet.Cells(2, 2).Value <> "Notification:" Then
        MsgBox "You need to have a pole detail sheet active to run this script."
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Exit Sub
    End If

    Dim fileDiag As FileDialog
    Set fileDiag = Application.FileDialog(msoFileDialogFolderPicker)
    With fileDiag
        .AllowMultiSelect = False
        .Title = "Select a folder "
        .InitialFileName = ThisWorkbook.path & Application.PathSeparator
        If .Show = -1 Then outputPath = Application.FileDialog(msoFileDialogFilePicker).SelectedItems.item(1) & Application.PathSeparator Else Exit Sub
    End With
    
    On Error Resume Next
    
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    ProgressBar_Form.Show vbModeless
    
    ProgressBar_Form.Label1.caption = "Exporting Pole Detail Sheet... Please wait..."
    ProgressBar_Form.Repaint

    Application.ScreenUpdating = True

    sheet.Range("A1").Select
    
    Application.ScreenUpdating = False

    Dim tempWb As Workbook
    Dim fileName As String
    Dim fullPath As String
    If sheet.name <> "4 Spans" And sheet.name <> "8 Spans" And sheet.name <> "12 Spans" And sheet.Cells(2, 2).Value = "Notification:" Then
        fileName = "M1P" & fixIllegalCharacters(sheet.Range("POLENUM").text) & "_" & fixIllegalCharacters(sheet.Range("CEID").text) & "_" & fixIllegalCharacters(sheet.Range("PERMIT").text)
        fullPath = outputPath & fileName & ".xlsx"
        
        For Each wb In Application.Workbooks
            If wb.FullName = fullPath Then wb.Close SaveChanges:=True
        Next wb
        
        Call Figures.clearSheetFigures(sheet)
        sheet.Copy
        Set tempWb = ActiveWorkbook
        
        tempWb.SaveAs fileName:=fullPath, FileFormat:=xlOpenXMLWorkbook
        tempWb.Close SaveChanges:=True
        
    End If
    
    'ThisWorkbook.sheets("Control").Activate
    
    ProgressBar_Form.Hide
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub


