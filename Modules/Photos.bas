Attribute VB_Name = "Photos"
Public Sub photoNameValidate()
    On Error Resume Next
    
    Call LogMessage.SendLogMessage("photoNameValidate")
    
    Dim pole As pole
    Dim path As String
    Dim fileName As String
    
    Set fileDiag = Application.FileDialog(msoFileDialogFolderPicker)
    With fileDiag
        .AllowMultiSelect = False
        .Title = "Select the Photos folder"
        .InitialFileName = ThisWorkbook.path & Application.PathSeparator
        If .Show = -1 Then path = Application.FileDialog(msoFileDialogFilePicker).SelectedItems.item(1) & Application.PathSeparator Else Exit Sub
    End With
    
    ThisWorkbook.sheets("Control").Range("PHOTODIR").Value = path
    Shell "powershell -command ""Get-ChildItem '" & path & "' | Unblock-File""", vbHide
    
    Dim matches As Object
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    Dim photoCounter As String
    Dim newName As String
    Dim deleteCounter As Integer
    Dim renameCount As Integer: renameCount = 0
    Dim photoCounts As scripting.Dictionary: Set photoCounts = New scripting.Dictionary
    
    regex.Pattern = "M1P([^-]+)-([^_]+)_([^_]+)_([^.]+)\.(jpg|png)"
    regex.Global = True
    regex.IgnoreCase = True
    
    Dim cleanedName As String
    
    Dim Project As New Project
    Call Project.extractFromSheets
    
    Dim regex2 As Object
    Set regex2 = CreateObject("VBScript.RegExp")
    
    Dim regEx3 As Object
    Set regEx3 = CreateObject("VBScript.RegExp")
    
    Dim regEx4 As Object
    Set regEx4 = CreateObject("VBScript.RegExp")
    
    Dim regEx5 As Object
    Set regEx5 = CreateObject("VBScript.RegExp")
    
    Dim regEx6 As Object
    Set regEx6 = CreateObject("VBScript.RegExp")
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    For Each pole In Project.poles
        fileName = Dir(path & "*")
        photoCounter = 1
        Do While fileName <> ""
              
            If InStr(fileName, "jpg") > 0 Then
                fileExtension = "jpg"
            ElseIf InStr(fileName, "png") > 0 Then
                fileExtension = "png"
            End If
            
            regex2.Pattern = "_" & "0*" & pole.poleNumber & "_\d+"
            regex2.IgnoreCase = True
            regEx3.Pattern = "[_-]\(0*" & pole.poleNumber & "\)[_-]"
            regEx3.IgnoreCase = True
            regEx4.Pattern = "[_-]\(.*\)\(0*" & pole.poleNumber & "\)[_-]"
            regEx4.IgnoreCase = True
            regEx5.Pattern = "Location\s*" & pole.location & "-"
            regEx5.IgnoreCase = True
            
            If InStr(fileName, "_Misc_") > 0 Or InStr(fileName, "_(No Tag)") > 0 Then
                newName = "DELETE" & "." & fileExtension
            
                Do While fso.FileExists(path & newName)
                    deleteCounter = deleteCounter + 1
                    newName = "DELETE" & deleteCounter & "." & fileExtension
                Loop
                Name path & fileName As path & newName
                renameCount = renameCount + 1
            ElseIf regex2.test(fileName) Then
                newName = "M1P" & pole.poleNumber & "-" & photoCounter & "_" & pole.existingCEID & "_" & Project.permit & "." & fileExtension
                
                Do While fso.FileExists(path & correctFileName(newName))
                    photoCounter = photoCounter + 1
                    newName = "M1P" & pole.poleNumber & "-" & photoCounter & "_" & pole.existingCEID & "_" & Project.permit & "." & fileExtension
                Loop
                
                Name path & fileName As path & correctFileName(newName)
                photoCounter = photoCounter + 1
                renameCount = renameCount + 1
            ElseIf regEx3.test(fileName) Then
                newName = "M1P" & pole.poleNumber & "-" & photoCounter & "_" & pole.existingCEID & "_" & Project.permit & "." & fileExtension
                            
                Do While fso.FileExists(path & correctFileName(newName))
                    photoCounter = photoCounter + 1
                    newName = "M1P" & pole.poleNumber & "-" & photoCounter & "_" & pole.existingCEID & "_" & Project.permit & "." & fileExtension
                Loop
                
                Name path & fileName As path & correctFileName(newName)
                photoCounter = photoCounter + 1
                renameCount = renameCount + 1
            ElseIf regEx4.test(fileName) Then
                newName = "M1P" & pole.poleNumber & "-" & photoCounter & "_" & pole.existingCEID & "_" & Project.permit & "." & fileExtension
                
                Do While fso.FileExists(path & correctFileName(newName))
                    photoCounter = photoCounter + 1
                    newName = "M1P" & pole.poleNumber & "-" & photoCounter & "_" & pole.existingCEID & "_" & Project.permit & "." & fileExtension
                Loop
                
                Name path & fileName As path & correctFileName(newName)
                photoCounter = photoCounter + 1
                renameCount = renameCount + 1
            ElseIf regEx5.test(fileName) Then
                newName = "M1P" & pole.poleNumber & "-" & photoCounter & "_" & pole.existingCEID & "_" & Project.permit & "." & fileExtension
                
                Do While fso.FileExists(path & correctFileName(newName))
                    photoCounter = photoCounter + 1
                    newName = "M1P" & pole.poleNumber & "-" & photoCounter & "_" & pole.existingCEID & "_" & Project.permit & "." & fileExtension
                Loop
                
                Name path & fileName As path & correctFileName(newName)
                photoCounter = photoCounter + 1
                renameCount = renameCount + 1
            End If
        
            fileName = Dir()
        Loop
    Next pole
    
    For Each pole In Project.poles
        fileName = Dir(path & "*")
        photoCounter = 1
        
        regEx6.Pattern = "M1P" & pole.poleNumber & "-(\d*)_.*_"
        regEx6.IgnoreCase = True
        
        Do While fileName <> ""
            If regEx6.test(fileName) Then
                photoCounter = 1
                newName = "M1P" & pole.poleNumber & "-" & photoCounter & "_" & pole.existingCEID & "_" & Project.permit & "." & fileExtension
                Set matches = regEx6.Execute(fileName)
                existingPhotoCounter = matches(0).SubMatches(0)
                
                
                Do While fso.FileExists(path & correctFileName(newName))
                    photoCounter = photoCounter + 1
                    newName = "M1P" & pole.poleNumber & "-" & photoCounter & "_" & pole.existingCEID & "_" & Project.permit & "." & fileExtension
                Loop
                
                If CInt(existingPhotoCounter) > CInt(photoCounter) Then
                    Name path & fileName As path & correctFileName(newName)
                    renameCount = renameCount + 1
                End If
            End If
            fileName = Dir()
        Loop
    Next pole
                
    fileName = Dir(path & "*")
    Do While fileName <> ""
        cleanedName = fileName
        If Len(cleanedName) - Len(Replace(cleanedName, "_", "")) > 2 Then
            firstPos = InStr(1, cleanedName, "_")
            lastPos = InStrRev(cleanedName, "_")
            middle = Replace(Mid$(cleanedName, firstPos + 1, lastPos - firstPos - 1), "_", "")
            cleanedName = Left$(cleanedName, firstPos) & middle & Mid$(cleanedName, lastPos)
        End If
        If regex.test(cleanedName) Then
            Set matches = regex.Execute(cleanedName)
            Dim poleNumber As String
            poleNumber = matches(0).SubMatches(0)
            photoCounter = matches(0).SubMatches(1)
            ceid = matches(0).SubMatches(2)
            permit = matches(0).SubMatches(3)
            fileExtension = matches(0).SubMatches(4)
            
            If Utilities.SheetExists(poleNumber) Then
                If Not photoCounts.Exists(poleNumber) Then photoCounts.Add poleNumber, 0
                photoCounts(poleNumber) = photoCounts(poleNumber) + 1
                
                Dim pds As Worksheet
                For Each sheet In ThisWorkbook.sheets
                    If ThisWorkbook.RemoveParentheses(sheet.name) = poleNumber Then
                        Set pds = sheet
                        Exit For
                    End If
                Next sheet
                If pds Is Nothing Then Exit Sub
                If permit = Utilities.correctFileName(pds.Range("PERMIT")) Then
                    If ceid <> pds.Range("CEID") Then
                        newName = "M1P" & poleNumber & "-" & photoCounter & "_" & pds.Range("CEID") & "_" & permit & "." & fileExtension
                        
                        Do While fso.FileExists(path & correctFileName(newName))
                            photoCounter = photoCounter + 1
                            newName = "M1P" & poleNumber & "-" & photoCounter & "_" & pds.Range("CEID") & "_" & permit & "." & fileExtension
                        Loop
                        
                        Name path & fileName As path & correctFileName(newName)
                        renameCount = renameCount + 1
                    End If
                Else
                    MsgBox "Missmatching permit on photo and pole " & poleNumber & vbLf & "Can't verify this is the correct photos folder"
                    Exit Sub
                End If
            ElseIf IsNumeric(ceid) Then
                For Each sheet In ThisWorkbook.sheets
                    If sheet.name <> "4 Spans" And sheet.name <> "8 Spans" And sheet.name <> "12 Spans" Then
                        If sheet.Cells(2, 2).Value = "Notification:" Then
                            If permit = Utilities.correctFileName(sheet.Range("PERMIT")) Then
                                If ceid = sheet.Range("CEID") Then
                                    
                                    poleNumber = sheet.Range("POLENUM")
                                    If Not photoCounts.Exists(poleNumber) Then photoCounts.Add poleNumber, 0
                                    photoCounts(poleNumber) = photoCounts(poleNumber) + 1
                                        
                                    newName = "M1P" & sheet.Range("POLENUM") & "-" & photoCounter & "_" & sheet.Range("CEID") & "_" & permit & "." & fileExtension
                                    
                                    Do While fso.FileExists(path & correctFileName(newName))
                                        photoCounter = photoCounter + 1
                                        newName = "M1P" & sheet.Range("POLENUM") & "-" & photoCounter & "_" & sheet.Range("CEID") & "_" & permit & "." & fileExtension
                                    Loop
                                    
                                    Name path & fileName As path & correctFileName(newName)
                                    renameCount = renameCount + 1
                                End If
                            Else
                                MsgBox "Missmatching permit on photo and pole " & poleNumber & vbLf & "Can't verify this is the correct photos folder"
                                Exit Sub
                            End If
                        End If
                    End If
                Next sheet
            End If
        End If
        fileName = Dir()
    Loop
    
    For Each poleNum In photoCounts
        For Each sheet In ThisWorkbook.sheets
            If sheet.name <> "4 Spans" And sheet.name <> "8 Spans" And sheet.name <> "12 Spans" Then
                If sheet.Cells(2, 2).Value = "Notification:" Then
                    If poleNum = sheet.Range("POLENUM") Then
                        sheet.Range("PICTURES").Value = "1-" & photoCounts(poleNum)
                    End If
                End If
            End If
        Next sheet
    Next poleNum
    
    MsgBox renameCount & " Photos renamed"
    
End Sub
    
Public Sub OpenPolePhoto(Optional getDir As Boolean = True)
    Dim sheet As Worksheet: Set sheet = ThisWorkbook.ActiveSheet()
    If Not Utilities.IsPDS(sheet) Then
        Exit Sub
    End If

    Dim pole As pole: Set pole = New pole
    Call pole.extractFromSheet(sheet)

    path = ThisWorkbook.sheets("Control").Range("PHOTODIR")
    
    If getDir Then
        If path = "" Then
            Set fileDiag = Application.FileDialog(msoFileDialogFolderPicker)
            With fileDiag
                .AllowMultiSelect = False
                .Title = "Select the Photos folder"
                .InitialFileName = ThisWorkbook.path & Application.PathSeparator
                If .Show = -1 Then path = Application.FileDialog(msoFileDialogFilePicker).SelectedItems.item(1) & Application.PathSeparator Else Exit Sub
            End With
            ThisWorkbook.sheets("Control").Range("PHOTODIR").Value = path
        ElseIf Dir(path) = "" Then
            Set fileDiag = Application.FileDialog(msoFileDialogFolderPicker)
            With fileDiag
                .AllowMultiSelect = False
                .Title = "Select the Photos folder"
                .InitialFileName = ThisWorkbook.path & Application.PathSeparator
                If .Show = -1 Then path = Application.FileDialog(msoFileDialogFilePicker).SelectedItems.item(1) & Application.PathSeparator Else Exit Sub
            End With
            ThisWorkbook.sheets("Control").Range("PHOTODIR").Value = path
        End If
    End If
    
    photoName = "M1P" & pole.poleNumber & "-1_" & pole.existingCEID & "_" & Utilities.correctFileName(pole.permit) & ".jpg"
    filePath = path & photoName
    
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    
    If Dir(filePath) <> "" Then
        wsh.Run "cmd /c start """" """ & filePath & """", vbHide
    End If
End Sub

Sub PhotoRenamer()
    Call LogMessage.SendLogMessage("PhotoRenamer")
    Call PhotoRenamer_Form.Initialize
    PhotoRenamer_Form.Show vbModeless
End Sub

