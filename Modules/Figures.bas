Attribute VB_Name = "Figures"
Public Sub getFigures()
    Dim sheet As Worksheet: Set sheet = ThisWorkbook.ActiveSheet()
    If sheet.name = "4 Spans" Or sheet.name = "8 Spans" Or sheet.name = "12 Spans" Or sheet.Cells(2, 2).Value <> "Notification:" Then
        MsgBox "You need to have a pole detail sheet active to run this script."
        Exit Sub
    End If
    
    Call getSheetFigures(sheet)
End Sub

Public Sub clearFigures()
    Dim sheet As Worksheet: Set sheet = ThisWorkbook.ActiveSheet()
    If sheet.name = "4 Spans" Or sheet.name = "8 Spans" Or sheet.name = "12 Spans" Or sheet.Cells(2, 2).Value <> "Notification:" Then
        MsgBox "You need to have a pole detail sheet active to run this script."
        Exit Sub
    End If
    
    Call clearSheetFigures(sheet)
End Sub

Public Sub getSheetFigures(sheet As Worksheet)
    Application.ScreenUpdating = False
    
    Dim cell As Range
    Dim edemPath As String
    Dim line As String
    Dim linkCount As Integer: linkCount = 0
    Dim maxLinkCount As Integer: maxLinkCount = 10
    Dim match As Object
    Dim lines() As String
    Dim i As Integer
    Dim col As Integer
    col = sheet.Range("DL").Column + 43
    
    edemPath = "C:\Distrib\CE Standard Work\doc\edem\EDEM_"
    edmPath = "C:\Distrib\CE Standard Work\doc\edm\EDM_"
    
    On Error Resume Next
    sheet.Unprotect
    
    altOneString = sheet.Range("ALTONE").text
    lines = Split(altOneString, Chr(10))
    For i = LBound(lines) To UBound(lines)
        line = Replace(lines(i), " ", "")
        If InStr(line, "FIGURE") > 0 Then
            Dim regex As Object: Set regex = CreateObject("VBScript.RegExp")
                
            regex.Pattern = "FIGURE(\d+-\d+)"
            regex.IgnoreCase = True
            regex.Global = False
            
            If regex.test(line) Then
                Set match = regex.Execute(line)(0)
                
                Set cell = sheet.Cells(linkCount + 2, col)
                If Not cell.Locked Then cell.Locked = False
                
                cell.Value = "FIGURE " & match.SubMatches(0)
                cell.EntireColumn.AutoFit
                cell.Hyperlinks.Add _
                    Anchor:=cell, _
                    address:=IIf(match.SubMatches(0) = "22-405", edmPath, edemPath) & match.SubMatches(0) & ".html", _
                    TextToDisplay:=cell.Value
                
                linkCount = linkCount + 1
            End If
        End If
    Next i
    For i = linkCount To maxLinkCount
        Set cell = sheet.Cells(i + 2, col)
        cell.Locked = False
        If cell.Value <> "" Then cell.Value = ""
    Next i
    
    sheet.Protect _
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
    
    Application.ScreenUpdating = True

End Sub

Public Sub clearSheetFigures(sheet As Worksheet)
    Dim i As Integer
    sheet.Unprotect
    
    For i = 0 To 10
        Set cell = sheet.Cells(i + 2, 83)
        If cell.Value <> "" Then cell.Value = ""
    Next i
End Sub


