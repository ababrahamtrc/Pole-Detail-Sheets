VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PhotoRenamer_Form 
   Caption         =   "Photo Renamer"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7215
   OleObjectBlob   =   "PhotoRenamer_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PhotoRenamer_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Dim colWs As Worksheet
Dim colWsLastRow As Integer
Dim colWsHeaders As Scripting.Dictionary
Dim downloaded As Boolean

Public Sub Initialize()
    On Error Resume Next
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.height) - (0.5 * Me.height)
    
    Dim header As String
    Set colWs = ThisWorkbook.sheets("Collection")
    colWsLastRow = colWs.Cells(colWs.Rows.count, "A").End(xlUp).row
    Dim colWsLastCol As Integer: colWsLastCol = colWs.Cells(1, colWs.Columns.count).End(xlToLeft).Column
    Set colWsHeaders = New Scripting.Dictionary
    For i = 1 To colWsLastCol
        header = Trim(colWs.Cells(1, i).Value)
        If header <> "" Then
            colWsHeaders(header) = i
        End If
    Next i
    
    Dim jobInfoWs As Worksheet: Set jobInfoWs = ThisWorkbook.sheets("Job Info")
    Dim jobInfoWsLastRow As Integer: jobInfoWsLastRow = jobInfoWs.Cells(jobInfoWs.Rows.count, "A").End(xlUp).row
    Dim jobInfoWsLastCol As Integer: jobInfoWsLastCol = jobInfoWs.Cells(1, jobInfoWs.Columns.count).End(xlToLeft).Column
    Dim jobInfoWsHeaders As Scripting.Dictionary: Set jobInfoWsHeaders = New Scripting.Dictionary
    For i = 1 To jobInfoWsLastCol
        header = Trim(jobInfoWs.Cells(1, i).Value)
        If header <> "" Then
            jobInfoWsHeaders(header) = i
        End If
    Next i
    
    If jobInfoWsHeaders.Exists("Permit") Then TextBox2.Value = jobInfoWs.Cells(2, jobInfoWsHeaders("Permit"))
End Sub

Private Sub CommandButton1_Click()
    On Error Resume Next
    Dim folderPath As String
    folderPath = TextBox1.Value
    
    If folderPath = "" Or Dir(folderPath, vbDirectory) = "" Then
        MsgBox "Please select a directory"
        Exit Sub
    End If
    
    Dim counter As Integer
    For i = 2 To colWsLastRow
        folderName = folderPath & Application.PathSeparator & colWs.Cells(i, colWsHeaders("ID"))
        If Dir(folderName, vbDirectory) = "" Then
            MkDir folderName
            counter = counter + 1
        End If
    Next i
    
    MsgBox counter & " folders created."
End Sub

Private Sub CommandButton2_Click()
    On Error Resume Next
    
    Dim folderPath, folderName, photosFolder, oldPath, newFileName, newPath, permit, poleNumber, ceid, poleNum, fileName As String
    Dim photoCounter, counter As Integer
    
    folderPath = TextBox1.Value & Application.PathSeparator
    
    If TextBox2.Value = "" Then
        MsgBox "Please enter a permit number."
        Exit Sub
    ElseIf folderPath = "" Or Dir(folderPath, vbDirectory) = "" Then
        MsgBox "Please select a directory"
        Exit Sub
    End If
    
    For i = 2 To colWsLastRow
        folderName = folderPath & colWs.Cells(i, colWsHeaders("ID")).Value
        If Dir(folderName, vbDirectory) = "" Then
            MsgBox "Not all folders created in directory, please generate folders."
            Exit Sub
        End If
    Next i
    
    photosFolder = folderPath & "Photos"
    If Dir(photosFolder, vbDirectory) = "" Then
        MkDir photosFolder
    End If
    
    permit = TextBox2.Value
    For i = 2 To colWsLastRow
        poleNum = colWs.Cells(i, colWsHeaders("ID")).Value
        ceid = ""
        If colWsHeaders.Exists("New CE ID Tag") Then ceid = colWs.Cells(i, colWsHeaders("New CE ID Tag"))
        If ceid = "" Then ceid = colWs.Cells(i, colWsHeaders("CE ID Tag"))
        If ceid = "" Then ceid = "FOREIGN"
        fileName = Dir(folderPath & poleNum & Application.PathSeparator & "*")
        Dim oldFiles As Collection: Set oldFiles = New Collection
        Do While fileName <> ""
            oldPath = folderPath & poleNum & Application.PathSeparator & fileName
            If GetAttr(oldPath) <> vbDirectory Then
                oldFiles.Add oldPath
            End If
            fileName = Dir()
        Loop
        photoCounter = 1
        For Each oldFile In oldFiles
            Do
                newFileName = "M1P" & poleNum & "-" & photoCounter & "_" & ceid & "_" & permit & ".jpg"
                newPath = photosFolder & Application.PathSeparator & newFileName
                photoCounter = photoCounter + 1
            Loop While Dir(newPath) <> ""
            Name oldFile As newPath
            counter = counter + 1
        Next oldFile
        RmDir folderPath & poleNum
    Next i
    MsgBox counter & " photos renamed"
End Sub

Private Sub CommandButton3_click()
    On Error Resume Next
    
    Dim fileDiag As FileDialog
    Set fileDiag = Application.FileDialog(msoFileDialogFolderPicker)
    With fileDiag
        .AllowMultiSelect = False
        .Title = "Select a folder "
        .InitialFileName = "C:\" & Application.PathSeparator
        If .Show = -1 Then outputPath = Application.FileDialog(msoFileDialogFilePicker).SelectedItems.item(1) Else Exit Sub
    End With
    
    TextBox1.Value = outputPath
End Sub

Private Sub CommandButton4_click()
    Dim header, imageUrl, photosFolder, folderPath, fileName, savepath, poleNum, ceid, permit As String
    Dim photoCounter As Integer
    Dim http, stream As Object
    
    On Error Resume Next
    
    folderPath = TextBox1.Value & Application.PathSeparator
    If TextBox2.Value = "" Then
        MsgBox "Please enter a permit number."
        Exit Sub
    ElseIf folderPath = "" Or Dir(folderPath, vbDirectory) = "" Then
        MsgBox "Please select a directory"
        Exit Sub
    ElseIf downloaded Then
        MsgBox "Photos already downloaded"
        Exit Sub
    End If
    
    Dim colWs As Worksheet: Set colWs = ThisWorkbook.sheets("Collection")
    Dim colWsLastRow As Integer: colWsLastRow = colWs.Cells(colWs.Rows.count, "A").End(xlUp).row
    Dim colWsLastCol As Integer: colWsLastCol = colWs.Cells(1, colWs.Columns.count).End(xlToLeft).Column
    Dim colWsHeaders As Scripting.Dictionary: Set colWsHeaders = New Scripting.Dictionary
    For i = 1 To colWsLastCol
        header = Trim(colWs.Cells(1, i).Value)
        If header <> "" Then
            colWsHeaders(header) = i
        End If
    Next i
    
    Dim imageWs As Worksheet: Set imageWs = ThisWorkbook.sheets("Images")
    Dim imageWsLastRow As Integer: imageWsLastRow = imageWs.Cells(imageWs.Rows.count, "A").End(xlUp).row
    Dim imageWsLastCol As Integer: imageWsLastCol = imageWs.Cells(1, imageWs.Columns.count).End(xlToLeft).Column
    Dim imageWsHeaders As Scripting.Dictionary: Set imageWsHeaders = New Scripting.Dictionary
    For i = 1 To imageWsLastCol
        header = Trim(imageWs.Cells(1, i).Value)
        If header <> "" Then
            imageWsHeaders(header) = i
        End If
    Next i
    
    photosFolder = folderPath & "Photos"
    If Dir(photosFolder, vbDirectory) = "" Then
        MkDir photosFolder
    End If
    
    ProgressBar_Form.Show vbModeless
    
    Dim poleNumRange, cell As Range
    Set poleNumRange = colWs.Range(colWs.Cells(2, colWsHeaders("ID")), colWs.Cells(colWsLastRow, colWsHeaders("ID")))
    permit = TextBox2.Value
    Set http = CreateObject("MSXML2.XMLHTTP")
    For i = 2 To imageWsLastRow
        ProgressBar_Form.Label1.caption = "Downloading image " & (i - 1) & " of " & (imageWsLastRow - 1) & " ... please wait..."
        ProgressBar_Form.Repaint
        imageUrl = imageWs.Cells(i, imageWsHeaders("value"))
        If poleNum <> imageWs.Cells(i, imageWsHeaders("ID")) Then
            ceid = ""
            poleNum = imageWs.Cells(i, imageWsHeaders("ID"))
            Set cell = poleNumRange.find(what:=poleNum, LookIn:=xlValues, lookat:=xlWhole)
            If colWsHeaders.Exists("New CE ID Tag") Then ceid = colWs.Cells(cell.row, colWsHeaders("New CE ID Tag"))
            If ceid = "" Then ceid = colWs.Cells(cell.row, colWsHeaders("CE ID Tag"))
            If ceid = "" Then ceid = "FOREIGN"
            photoCounter = 1
        End If
        Do
            fileName = "M1P" & poleNum & "-" & photoCounter & "_" & ceid & "_" & permit & ".jpg"
            savepath = photosFolder & Application.PathSeparator & fileName
            photoCounter = photoCounter + 1
        Loop While Dir(savepath) <> ""
        http.Open "Get", imageUrl, False
        http.Send
        If http.Status = 200 Then
            Set stream = CreateObject("ADODB.Stream")
            stream.Type = 1
            stream.Open
            stream.Write http.responseBody
            stream.saveToFile savepath, 2
            stream.Close
        End If
    Next i
    Unload ProgressBar_Form
    downloaded = True
End Sub
