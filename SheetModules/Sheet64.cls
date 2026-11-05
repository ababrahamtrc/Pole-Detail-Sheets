VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo SafeExit
    
    Application.EnableEvents = True

    If Target.CountLarge > 1 Then GoTo SafeExit
    
    Dim data_ws As Worksheet
    Set data_ws = ThisWorkbook.sheets("DATA")
    
    Dim catagoryCell As Range
    Dim hourlycostCell As Range
    
    If Not Intersect(Target, Me.Range("C6:C56")) Is Nothing Then

        Set catagoryCell = data_ws.cells.Find(Target.value)
        Set hourlycostCell = data_ws.cells.Find("Engineering Hourly Cost")
        If Target.value <> "" Then
            Target.offSet(0, 1).value = "$ " & Round((data_ws.cells(catagoryCell.ROW, catagoryCell.column + 1).value * data_ws.cells(hourlycostCell.ROW + 1, hourlycostCell.column).value), 2)
        Else
            Target.offSet(0, 1).value = ""
        End If
    End If
    
SafeExit:
    Application.EnableEvents = True
End Sub


