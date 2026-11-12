VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    
    Application.EnableEvents = True
    
    If Sh.name <> "4 Spans" And Sh.name <> "8 Spans" And Sh.name <> "12 Spans" Then
        If Sh.cells(2, 2).value = "Notification:" Then
            If Not Intersect(Target, Sh.Range("DL")) Is Nothing Or Not Intersect(Target, Sh.Range("SUMSHEET9")) Is Nothing Or Not Intersect(Target, Sh.Range("SUMSHEET8")) Is Nothing Or Not Intersect(Target, Sh.Range("SUMSHEET4")) Is Nothing Or Not Intersect(Target, Sh.Range("SUMSHEET1")) Is Nothing Or Not Intersect(Target, Sh.Range("SUMSHEET7")) Is Nothing Or Not Intersect(Target, Sh.Range("NJUNS")) Is Nothing Then
                Call decideTabColor(Sh)
            ElseIf Not Intersect(Target, Sh.Range("SUMSHEET3")) Is Nothing Or Not Intersect(Target, Sh.Range("SUMSHEET6")) Is Nothing Then
                If Sh.Range("SUMSHEET3").value Then
                    If Sh.Range("SUMSHEET6") Then
                        Sh.Range("SUMSHEET11").value = "1" & Chr(44) & "2" & Chr(44) & "3"
                    Else
                        Sh.Range("SUMSHEET11").value = "1" & Chr(44) & "2"
                    End If
                ElseIf Sh.Range("SUMSHEET6") Then
                    Sh.Range("SUMSHEET11").value = "1" & Chr(44) & "3"
                Else
                    Sh.Range("SUMSHEET11").value = ""
                End If
            End If
        End If
    End If
    
    If Sh.name = "Control" Then
        If Not Intersect(Target, Sh.Range("MODE")) Is Nothing Then
            If Sh.Range("MODE") = "MAKE READY" Then
                ThisWorkbook.sheets("MKR SAP").visible = True
                ThisWorkbook.sheets("SI SAP").visible = False
            ElseIf Sh.Range("MODE") = "SYSTEM IMPROVEMENT" Then
                ThisWorkbook.sheets("MKR SAP").visible = False
                ThisWorkbook.sheets("SI SAP").visible = True
            End If
        End If
    End If
    
    If Sh.name <> "4 Spans" And Sh.name <> "8 Spans" And Sh.name <> "12 Spans" Then
        If Sh.cells(2, 2).value = "Notification:" Then
            If Utilities.RangeExists(Sh, "TROWSTART") And Utilities.RangeExists(Sh, "TROWEND") Then
                Set endrange = Sh.Range("TROWEND").MergeArea
                Dim insideRange As Range
                startrow = Sh.Range("TROWSTART").ROW + Sh.Range("TROWSTART").Rows.count
                endRow = Sh.Range("TROWEND").ROW - 2
                If startrow <= endRow Then
                    Set insideRange = Sh.Rows(startrow & ":" & endRow)
                Else
                    Set insideRange = Nothing
                End If
                If Not endrange Is Nothing Then
                    If Not Intersect(Target.EntireRow, endrange.EntireRow) Is Nothing Then
                        If Target.value = False Or Target.value = "" Then Exit Sub
                        Sh.Unprotect
                        
                        Application.EnableEvents = False
                        Application.ScreenUpdating = False
                        
                        pasteRow = endrange.ROW + endrange.Rows.count
                        Set targetRows = Sh.Rows(pasteRow).Resize(endrange.Rows.count)
                        
                        endrange.EntireRow.Copy Destination:=targetRows
                        
                        Set activeCells = Intersect(targetRows, ActiveSheet.UsedRange)
                        
                        For Each cell In activeCells.cells
                            If VarType(cell.value) <> vbBoolean Then
                                If cell.value <> "" Then cell.value = ""
                            Else
                                If cell.value <> False Then cell.value = False
                                Exit For
                            End If
                        Next cell
                        
                        Set newEndRange = endrange.offSet(endrange.Rows.count, 0)
                        Sh.Names("TROWEND").RefersTo = newEndRange
                        
                        Sh.Protect
                        
                        Application.CutCopyMode = False
                        Target.Select
                        Application.EnableEvents = True
                        Application.ScreenUpdating = True
                    ElseIf Not Intersect(Target.EntireRow, insideRange.EntireRow) Is Nothing Then
                        Sh.Unprotect
                        
                        Application.EnableEvents = False
                        Application.ScreenUpdating = False
                        
                        For i = endRow To startrow Step -2
                            Set targetRows = Sh.Rows(i - 1).Resize(2)
                            Set activeCells = Intersect(targetRows, ActiveSheet.UsedRange)
                            
                            Dim emptyRow As Boolean: emptyRow = True
                            For Each cell In activeCells.cells
                                If VarType(cell.value) <> vbBoolean Then
                                    If cell.value <> "" Then
                                        emptyRow = False
                                        Exit For
                                    End If
                                Else
                                    If cell.value <> False Then emptyRow = False
                                    Exit For
                                End If
                            Next cell
                            
                            If emptyRow Then
                                Call Sh.Rows(i).Delete
                                Call Sh.Rows(i - 1).Delete
                            End If
                        Next i
                        Sh.Protect
                        
                        Application.EnableEvents = True
                        Application.ScreenUpdating = True
                    End If
                End If
            End If
        End If
    End If
End Sub

Public Sub decideTabColor(Sh As Worksheet)
    On Error Resume Next
    
    If Trim(Sh.Range("DL")) <> "" Then
        If Sh.name <> Trim(Sh.Range("POLENUM")) & "(" & Trim(fixSpecialCharacters(Sh.Range("DL"))) & ")" Then Sh.name = Trim(Sh.Range("POLENUM")) & "(" & Trim(fixSpecialCharacters(Sh.Range("DL"))) & ")"
    Else
        If Sh.name <> Trim(Sh.Range("POLENUM")) Then Sh.name = Trim(Sh.Range("POLENUM"))
    End If

    If Trim(Sh.Range("DL")) <> "" And (Sh.Range("SUMSHEET4") = True Or Sh.Range("NJUNS") <> "") And Sh.Range("NJUNS") <> "COMM MAKE READY WORK" And ThisWorkbook.sheets("Control").Range("COLOR1").offSet(0, 1).Interior.ColorIndex <> xlNone Then
        If Sh.Tab.color <> ThisWorkbook.sheets("Control").Range("COLOR1").offSet(0, 1).Interior.color Then Sh.Tab.color = ThisWorkbook.sheets("Control").Range("COLOR1").offSet(0, 1).Interior.color
    ElseIf Trim(Sh.Range("DL")) <> "" And ThisWorkbook.sheets("Control").Range("COLOR2").offSet(0, 1).Interior.ColorIndex <> xlNone Then
        If Sh.Tab.color <> ThisWorkbook.sheets("Control").Range("COLOR2").offSet(0, 1).Interior.color Then Sh.Tab.color = ThisWorkbook.sheets("Control").Range("COLOR2").offSet(0, 1).Interior.color
    ElseIf Sh.Range("SUMSHEET9") = True And ThisWorkbook.sheets("Control").Range("COLOR3").offSet(0, 1).Interior.ColorIndex <> xlNone Then
        If Sh.Tab.color <> ThisWorkbook.sheets("Control").Range("COLOR3").offSet(0, 1).Interior.color Then Sh.Tab.color = ThisWorkbook.sheets("Control").Range("COLOR3").offSet(0, 1).Interior.color
    ElseIf Sh.Range("SUMSHEET8") = True And Sh.Range("SUMSHEET4") = True And ThisWorkbook.sheets("Control").Range("COLOR4").offSet(0, 1).Interior.ColorIndex <> xlNone Then
        If Sh.Tab.color <> ThisWorkbook.sheets("Control").Range("COLOR4").offSet(0, 1).Interior.color Then Sh.Tab.color = ThisWorkbook.sheets("Control").Range("COLOR4").offSet(0, 1).Interior.color
    ElseIf Sh.Range("SUMSHEET8") = True And ThisWorkbook.sheets("Control").Range("COLOR5").offSet(0, 1).Interior.ColorIndex <> xlNone Then
        If Sh.Tab.color <> ThisWorkbook.sheets("Control").Range("COLOR5").offSet(0, 1).Interior.color Then Sh.Tab.color = ThisWorkbook.sheets("Control").Range("COLOR5").offSet(0, 1).Interior.color
    ElseIf Sh.Range("SUMSHEET4") = True And ThisWorkbook.sheets("Control").Range("COLOR6").offSet(0, 1).Interior.ColorIndex <> xlNone Then
        If Sh.Tab.color <> ThisWorkbook.sheets("Control").Range("COLOR6").offSet(0, 1).Interior.color Then Sh.Tab.color = ThisWorkbook.sheets("Control").Range("COLOR6").offSet(0, 1).Interior.color
    ElseIf (Sh.Range("SUMSHEET1") = True Or Sh.Range("SUMSHEET7") = True) And ThisWorkbook.sheets("Control").Range("COLOR7").offSet(0, 1).Interior.ColorIndex <> xlNone Then
        If Sh.Tab.color <> ThisWorkbook.sheets("Control").Range("COLOR7").offSet(0, 1).Interior.color Then Sh.Tab.color = ThisWorkbook.sheets("Control").Range("COLOR7").offSet(0, 1).Interior.color
    ElseIf InStr(Sh.Range("DET").offSet(0, 1), "Hammer Test") > 0 Then
        If Sh.Tab.color <> ThisWorkbook.sheets("Control").Range("COLOR8").offSet(0, 1).Interior.color Then Sh.Tab.color = ThisWorkbook.sheets("Control").Range("COLOR8").offSet(0, 1).Interior.color
    Else
        If Sh.Tab.color <> False Then Sh.Tab.color = False
    End If
    
End Sub


Public Function RemoveParentheses(ByVal inputStr As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    regex.Pattern = "\([^)]*\)"
    regex.Global = True
    regex.IgnoreCase = True
 
    RemoveParentheses = regex.Replace(inputStr, "")
End Function

Function InsideParentheses(ByVal inputStr As String) As String
    Dim openParenPos As Integer
    Dim closeParenPos As Integer
    Dim valueLength As Integer
    
    openParenPos = InStr(inputStr, "(")
    closeParenPos = InStr(inputStr, ")")
    
    If openParenPos > 0 And closeParenPos > 0 And closeParenPos > openParenPos Then
        valueLength = closeParenPos - openParenPos - 1
        InsideParentheses = Mid(inputStr, openParenPos + 1, valueLength)
    Else
        InsideParentheses = ""
    End If
End Function

Public Function fixSpecialCharacters(ByVal inputStr As String) As String
    inputStr = Replace(inputStr, "*", "#")
    inputStr = Replace(inputStr, ":", "#")
    inputStr = Replace(inputStr, "/", "#")
    inputStr = Replace(inputStr, "\", "#")
    inputStr = Replace(inputStr, "?", "#")
    inputStr = Replace(inputStr, "[", "#")
    inputStr = Replace(inputStr, "]", "#")

    fixSpecialCharacters = inputStr
End Function

Private Sub Workbook_Open()
    Call UpdatePoleDetailSheets.CheckVersion(True)
End Sub








