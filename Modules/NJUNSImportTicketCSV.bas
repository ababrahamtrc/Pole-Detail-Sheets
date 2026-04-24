Attribute VB_Name = "NJUNSImportTicketCSV"
Sub ImportNJUNSTicketCSV()
 
    Dim tickets As Object: Set tickets = CreateObject("Scripting.Dictionary")
    Dim filePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select a CSV File"
        .Filters.clear
        .Filters.Add "CSV Files", "*.csv"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            MsgBox "No file selected."
            Exit Sub
        End If
        filePath = .SelectedItems(1)
    End With
    
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 1)
    Dim line As String
    Dim parts As Variant

    If Not ts.AtEndOfStream Then ts.ReadLine
    Do While Not ts.AtEndOfStream
        line = ts.ReadLine
        If Trim(line) <> "" Then
            parts = Split(line, ",")

            If UBound(parts) >= 1 Then
                tickets(Replace(parts(0), """", "")) = Replace(parts(1), """", "")
            End If
        End If
    Loop
    ts.Close

    Dim project As project: Set project = New project
    Call project.extractFromSheets
    
    Dim ticketsImported As Integer
    Dim pole As pole
    For Each pole In project.poles
        If pole.NJUNS <> "" Then
            If tickets.Exists(pole.poleNumber) Then
                If Utilities.OnlyNumbers(pole.njunsTicket) = -1 Then
                    If InStr(pole.njunsTicket, "NOTIFY") > 0 Then
                        pole.njunsTicket = "NOTIFY-" & tickets(pole.poleNumber)
                    ElseIf InStr(pole.njunsTicket, "CA") > 0 Then
                        pole.njunsTicket = "CA-" & tickets(pole.poleNumber)
                    ElseIf InStr(pole.njunsTicket, "PT") > 0 Then
                        pole.njunsTicket = "PT-" & tickets(pole.poleNumber)
                    End If
                    Set sheet = Utilities.GetPDS(pole.poleNumber)
                    If Not sheet Is Nothing Then
                        sheet.Range("NJUNSTICKET") = pole.njunsTicket
                        ticketsImported = ticketsImported + 1
                    End If
                End If
            End If
        End If
    Next pole
 
    MsgBox ("Done, " & ticketsImported & " tickets imported.")
End Sub
