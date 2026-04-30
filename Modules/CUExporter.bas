Attribute VB_Name = "CUExporter"
Public addedCU As Boolean
Public guySection As Boolean
Public hotsite As Boolean
Public timeAdder As Integer
Public vpoPole As Boolean
Public serviceAmount As Integer
Public reconductored As Boolean
Public streetlightMolding As String

Public Sub CopyCUImportCode()
    Call LogMessage.SendLogMessage("CopyCUImportCode")

    Dim url As String: url = "https://api.github.com/repos/ElijahRademaker/Automation-tools/contents/cuimport.js"
    Dim file As Object
    Dim strText As String
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.Send
 
    If http.Status <> 200 Then
        MsgBox "Failed to get cuimport.js from github: " & http.Status & vbLf & JsonConverter.ParseJson(http.responseText)("message")
        Exit Sub
    End If
 
    Set file = JsonConverter.ParseJson(http.responseText)
    Call UpdatePoleDetailSheets.DownloadFile(file)

    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile Environ$("TEMP") & "\cuimport.js"
    strText = stm.ReadText
    stm.Close

    Dim DataObj As DataObject: Set DataObj = New DataObject
    If Len(strText) > 0 Then
        DataObj.SetText strText
        DataObj.PutInClipboard
        MsgBox "The code has been copied. Go to your Design Doc in EAM, Press F12, then paste and hit enter to load the Importer.", vbInformation
    Else
        MsgBox "The Text Box is empty.", vbExclamation
    End If
End Sub

Public Sub ExportAllSheetCUs()
    Call LogMessage.SendLogMessage("ExportAllCUs")
    
    Dim Project As Project: Set Project = New Project
    Call Project.extractFromSheets

    Dim CU As Variant
    Dim cus As Collection: Set cus = New Collection
    Dim missedLines As Collection: Set missedLines = New Collection
    Dim cusTemp As Collection
    Dim missedLinesTemp As Collection
    Dim inputCol As Collection
    Dim sheet As Worksheet
    
    For Each sheet In ThisWorkbook.sheets
        If Utilities.IsPDS(sheet) Then
            Set inputCol = ExportSheetCUs(Project, sheet)
            If Not inputCol Is Nothing Then
                Set cusTemp = inputCol(1)
                Set missedLinesTemp = inputCol(2)
                For Each CU In cusTemp
                    cus.Add CU
                Next CU
                For Each line In missedLinesTemp
                    missedLines.Add "Location " & sheet.Range("DL") & ": " & line
                Next line
            End If
        End If
    Next sheet
    
    ThisWorkbook.sheets("Control").Activate
    
    If cus.count > 0 Then
        Call generateCSV(Project, cus)
        If missedLines.count > 0 Then
            Call generateMissedLinesTXT(missedLines)
        Else
            MsgBox "All lines successfully turned into CUs."
        End If
    Else
        MsgBox "No CUs generated."
    End If
End Sub

Public Sub ExportSingleSheetCUs()
    Call LogMessage.SendLogMessage("ExportSingleCUs")

    Dim Project As Project: Set Project = New Project
    Call Project.extractFromSheets
    
    Dim sheet As Worksheet: Set sheet = ThisWorkbook.ActiveSheet()
    If Not Utilities.IsPDS(sheet) Then
        MsgBox "You need to have a pole detail sheet active to run this script."
        Exit Sub
    End If
    
    Dim cus As Collection
    Dim missedLines As Collection
    
    Dim inputCol As Collection: Set inputCol = New Collection
    Set inputCol = ExportSheetCUs(Project, sheet)
    If Not inputCol Is Nothing Then
        Set cus = inputCol(1)
        Set missedLines = inputCol(2)
    End If
    
    If Not cus Is Nothing Then
        If cus.count > 0 Then
            Call generateCSV(Project, cus)
            If missedLines.count > 0 Then
                MsgBox "Lines unable to turn into CUS." & vbLf & Utilities.JoinCollection(missedLines, vbLf)
            Else
                MsgBox "All lines successfully turned into CUs."
            End If
        Else
            MsgBox "No CUs generated"
        End If
    Else
        MsgBox "No CUs generated"
    End If
End Sub

Private Function ExportSheetCUs(Project As Project, sheet As Worksheet) As Collection
    Dim installSection As Boolean, replaceSection As Boolean, removeSection As Boolean, transferSection As Boolean
    Dim line As Variant
    Dim lines() As String
    Dim installNotes As String, replaceNotes As String, removeNotes As String, transferNotes As String, notes As String
    Dim cus As Collection: Set cus = New Collection
    Dim missedLines As Collection: Set missedLines = New Collection
    Dim needAdditionalCUs As Collection: Set needAdditionalCUs = New Collection
    Dim pole As pole: Set pole = New pole
    Call pole.extractFromSheet(sheet)
    
    guySection = False
    hotsite = False
    reconductored = False
    vpoPole = False
    serviceAmount = 0
    streetlightMolding = ""
    
    lines = Split(pole.Alt1, vbLf)
    
    If pole.location = "" Or UBound(lines) < 1 Then Exit Function
    timeAdder = 1
    
    If Replace(Replace(pole.Alt1, "/", ""), "NA", "") = "" Then
        Set ExportSheetCUs = Nothing
        Exit Function
    End If
    
    If pole.commComponents.count > 0 Then
        cus.Add Array(properLocation(pole.location), 1.45)
    Else
        cus.Add Array(properLocation(pole.location), 1.3)
    End If
    
    For i = 0 To UBound(lines)
        line = lines(i)
        line = Replace(line, "/0", "|0")
        line = Replace(line, "/LA", "|LA")
        line = Replace(line, "/C", "|C")
        line = Replace(line, "1/2", "1|2")
        line = Replace(line, "AT&T", "ATT")
        line = Replace(line, "W/INS", "W|INS")
        
        
        If InStr(line, "NOTE") > 0 Then
            installSection = False
            replaceSection = False
            removeSection = False
            transferSection = False
        End If
    
        If Trim(line) = "INSTALL" Then
            installSection = True
            replaceSection = False
            removeSection = False
            transferSection = False
        ElseIf Trim(line) = "REPLACE" Then
            installSection = False
            replaceSection = True
            removeSection = False
            transferSection = False
        ElseIf Trim(line) = "REMOVE" Then
            installSection = False
            replaceSection = False
            removeSection = True
            transferSection = False
        ElseIf Trim(line) = "TRANSFER" Then
            installSection = False
            replaceSection = False
            removeSection = False
            transferSection = True
        Else
            If installSection Then
                If line <> "" Then Call parseLineToCUs(needAdditionalCUs, missedLines, cus, pole, line, "Install")
            ElseIf replaceSection Then
                If line <> "" Then Call parseLineToCUs(needAdditionalCUs, missedLines, cus, pole, line, "Replace")
            ElseIf removeSection Then
                If line <> "" Then Call parseLineToCUs(needAdditionalCUs, missedLines, cus, pole, line, "Remove")
            ElseIf transferSection Then
                If line <> "" Then Call parseLineToCUs(needAdditionalCUs, missedLines, cus, pole, line, "Transfer")
            Else
                If line <> "" Then Call parseLineToCUs(needAdditionalCUs, missedLines, cus, pole, line, "Note")
            End If
        End If
    Next i
    
    If IsNumeric(pole.ttc) Then
        Call generateTTCCU(cus, pole.location, CInt(Utilities.OnlyNumbers(pole.ttc)))
    Else
        missedLines.Add "Missing TTC in pole detail sheet, can't generate TTC CU"
    End If
 
    Call generateCU(cus, pole.location, "100417", timeAdder, "INSTALL")
    If hotsite Then Call generateCU(cus, pole.location, "106268", 1, "INSTALL")
    
    Call fixCUErrors(cus, needAdditionalCUs, missedLines)
    
    If needAdditionalCUs.count > 0 Then
        sheet.Activate
        sheet.Range("A1").Select
        Call findAdditonalCUs(cus, pole, needAdditionalCUs, missedLines)
    End If
    
    If Not reconductored Then Call checkForAdjacentPoleRecondcutoring(cus, Project, pole, missedLines)
    
    Dim outputCol As Collection: Set outputCol = New Collection
    outputCol.Add cus
    outputCol.Add missedLines
    
    Set ExportSheetCUs = outputCol
End Function

Private Sub fixCUErrors(cus As Collection, needAdditionalCUs As Collection, missedLines As Collection)
    Dim transferSpg As Integer
    Dim removeSpg As Boolean
    Dim CU As Variant
    Dim transferServices As Boolean
    Dim missedLine As String
    Dim hardware As String
    Dim regex As Object: Set regex = CreateObject("VBScript.RegExp")
    
    For i = 1 To cus.count
        If TypeOf cus(i) Is CU Then
            Set CU = cus(i)
            If TypeOf CU Is CU Then
                If CU.code = "106121" Then transferSpg = i
                If CU.code = "505040" And CU.action = "RET REM" Then removeSpg = True
                If CU.code = "106115" Then transferServices = True
                If vpoPole And CU.code = "100052" Then CU.qty = CU.qty - 1
            End If
        End If
    Next i
    
    For i = needAdditionalCUs.count To 1 Step -1
        neededCU = needAdditionalCUs(i)
        If vpoPole And neededCU(0) = "100052" Then
            neededCU(1) = neededCU(1) - 1
            If neededCU(1) = 0 Then Call needAdditionalCUs.Remove(i)
        End If
    Next i
    
    For i = missedLines.count To 1 Step -1
        missedLine = missedLines(i)
        regex.Pattern = "\((\d+)\)(.+)"
        regex.Global = True
        regex.IgnoreCase = True
        missedLine = Replace(missedLine, " ", "")
        missedLine = Replace(missedLine, "Install", "")
        missedLine = Replace(missedLine, "Remove", "")
        hardware = missedLine
        If regex.test(missedLine) Then
            Set matches = regex.Execute(missedLine)
            hardware = Trim(matches(0).SubMatches(1))
        End If
        
        If transferServices Then
            If InStr(Replace(hardware, "DEADEND", "DE"), "SERVICEDE") = 1 Then Call missedLines.Remove(i)
            If InStr(Replace(Replace(Replace(hardware, " ", ""), "DEADEND", "DE"), "OPENWIRE", "OW"), "OWSERVICEDE") = 1 Then Call missedLines.Remove(i)
        End If
    Next i
    
    If transferSpg > 0 And Not removeSpg Then Call cus.Remove(transferSpg)
End Sub

Private Sub generateTTCCU(cus As Collection, location As String, ttc As Integer)
    If (ttc >= 3 And ttc <= 7) Or ttc = 16 Or ttc = 17 Then Call generateCU(cus, location, "999013", 1, "INSTALL")
    If ttc = 8 Or ttc = 10 Or ttc = 11 Or ttc = 18 Or ttc = 19 Or ttc = 22 Then Call generateCU(cus, location, "999014", 1, "INSTALL")
    If (ttc >= 12 And ttc <= 14) Then Call generateCU(cus, location, "999015", 1, "INSTALL")
    If ttc = 15 Or ttc = 21 Then Call generateCU(cus, location, "999017", 1, "INSTALL")
    If ttc = 23 Then Call generateCU(cus, location, "999018", 1, "INSTALL")
End Sub

Private Sub findAdditonalCUs(cus As Collection, pole As pole, needAdditionalCUs As Collection, missedLines As Collection)
    Dim Wire As Wire
    Dim cuCode As String
    Dim priCount As Integer
    Dim neutCount As Integer
    Dim secCount As Integer
    Dim owCount As Integer
    Dim priSizes As scripting.Dictionary: Set priSizes = New scripting.Dictionary
    Dim neutSizes As scripting.Dictionary: Set neutSizes = New scripting.Dictionary
    Dim secSizes As scripting.Dictionary: Set secSizes = New scripting.Dictionary
    Dim owSizes As scripting.Dictionary: Set owSizes = New scripting.Dictionary
    
    Call SortCollectionByAction(needAdditionalCUs)
    
    'Running total of sizes, these go down as they're matched to a CU
    For Each Wire In pole.primaries
        If Not priSizes.Exists(Wire.size) Then priSizes(Wire.size) = 0
        For Each midspan In Wire.midspans
            priCount = priCount + Wire.phase
            priSizes(Wire.size) = priSizes(Wire.size) + Wire.phase
        Next midspan
    Next Wire
    For Each Wire In pole.neutrals
        If Not neutSizes.Exists(Wire.size) Then neutSizes(Wire.size) = 0
        For Each midspan In Wire.midspans
            neutCount = neutCount + 1
            neutSizes(Wire.size) = neutSizes(Wire.size) + 1
        Next midspan
    Next Wire
    For Each Wire In pole.secondaries
        If Not secSizes.Exists(Wire.size) Then secSizes(Wire.size) = 0
        For Each midspan In Wire.midspans
            secCount = secCount + 1
            secSizes(Wire.size) = secSizes(Wire.size) + 1
        Next midspan
    Next Wire
    For Each Wire In pole.openWires
        If Not owSizes.Exists(Wire.size) Then owSizes(Wire.size) = 0
        For Each midspan In Wire.midspans
            owCount = owCount + 1
            owSizes(Wire.size) = owSizes(Wire.size) + 1
        Next midspan
    Next Wire
    
    Dim CU As Variant
    For Each CU In cus
        If TypeOf CU Is CU Then
            If CU.location = properLocation(pole.location) And CU.code = "101036" And CU.action = "RET REM" Then
                owCount = owCount - CU.qty
                Exit For
            End If
        End If
    Next CU
    
    Dim neededCU() As Variant
    Dim hardware As String
    Dim amount As Integer
    Dim action As String
    
    'Find and calculate size CUs for deadends
    Dim i As Integer
    For i = needAdditionalCUs.count To 1 Step -1
        neededCU = needAdditionalCUs(i)
        hardware = Replace(neededCU(0), " ", "")
        amount = neededCU(1)
        action = neededCU(2)
        
        If InStr(hardware, "PRI") > 0 And InStr(hardware, "DE") > 0 Then
            Call getExtraDECU(cus, pole, needAdditionalCUs, i, priSizes, priCount, neededCU, "PRI")
        ElseIf InStr(hardware, "NEUT") > 0 And InStr(hardware, "DE") > 0 Then
            Call getExtraDECU(cus, pole, needAdditionalCUs, i, neutSizes, neutCount, neededCU)
        ElseIf InStr(hardware, "SEC") > 0 And InStr(hardware, "DE") > 0 Then
            Call getExtraDECU(cus, pole, needAdditionalCUs, i, secSizes, secCount, neededCU)
        End If
    Next i

    'Find all the lines that would require a spool tie
    Dim neededSpoolCUs As scripting.Dictionary: Set neededSpoolCUs = New scripting.Dictionary
    neededSpoolCUs("INSTALL") = 0
    neededSpoolCUs("RET REM") = 0
    For i = needAdditionalCUs.count To 1 Step -1
        neededCU = needAdditionalCUs(i)
        hardware = Replace(neededCU(0), " ", "")
        amount = neededCU(1)
        action = neededCU(2)
        
        If InStr(hardware, "WR") > 0 Or InStr(hardware, "1VPO") > 0 Or InStr(hardware, "2VPO") > 0 Or InStr(hardware, "3VPO") > 0 Then
            neededSpoolCUs(action) = neededSpoolCUs(action) + amount
        End If
    Next i
    
    'Calculate the size of the spool tie and remove from additionalCUs if size found unambiguously
    If neededSpoolCUs("RET REM") > 0 Then
        If neededSpoolCUs("RET REM") = neededSpoolCUs("INSTALL") Then
            Call getSpoolTies(cus, pole, needAdditionalCUs, neutSizes, neutCount, secSizes, secCount, owSizes, owCount, neededSpoolCUs("INSTALL"), "INSTALL")
        End If
        Call getSpoolTies(cus, pole, needAdditionalCUs, neutSizes, neutCount, secSizes, secCount, owSizes, owCount, neededSpoolCUs("RET REM"), "RET REM")
    End If
    
    'Prompt user for top/side ties if size can be found
    Dim topSideTie As String
    Dim uniqueTopSideTieSizes As scripting.Dictionary: Set uniqueTopSideTieSizes = New scripting.Dictionary
    
    For Each priSize In priSizes
        If Not uniqueTopSideTieSizes.Exists(Utilities.OnlyNumbers(CStr(priSize))) Then uniqueTopSideTieSizes.Add Utilities.OnlyNumbers(CStr(priSize)), Nothing
    Next priSize
    For Each neutSize In neutSizes
        If Not uniqueTopSideTieSizes.Exists(Utilities.OnlyNumbers(CStr(neutSize))) Then uniqueTopSideTieSizes.Add Utilities.OnlyNumbers(CStr(neutSize)), Nothing
    Next neutSize
    If owCount > 0 Then
        For Each owSize In owSizes
            If Not uniqueTopSideTieSizes.Exists(Utilities.OnlyNumbers(CStr(owSize))) Then uniqueTopSideTieSizes.Add Utilities.OnlyNumbers(CStr(owSize)), Nothing
        Next owSize
    End If
    
    If uniqueTopSideTieSizes.count = 0 Then
        For Each Wire In pole.primaries
            If Not uniqueTopSideTieSizes.Exists(Utilities.OnlyNumbers(CStr(Wire.size))) Then uniqueTopSideTieSizes.Add Utilities.OnlyNumbers(CStr(Wire.size)), Nothing
        Next Wire
    End If
    
    Dim size As String
    If priSizes.count = 1 Or uniqueTopSideTieSizes.count = 1 Then
        If priSizes.count = 1 Then
            size = priSizes.keys()(0)
        ElseIf uniqueTopSideTieSizes.count = 1 Then
            size = uniqueTopSideTieSizes.keys()(0)
        End If
        
        For i = needAdditionalCUs.count To 1 Step -1
            neededCU = needAdditionalCUs(i)
            hardware = Replace(neededCU(0), " ", "")
            amount = neededCU(1)
            action = neededCU(2)
            cuCode = ""
            
            If InStr(hardware, "PTP") > 0 Or InStr(hardware, "SPINS") > 0 Then
                If topSideTie = "" Then
                    Call OpenPolePhoto(False)
                    Unload CU_Form
                    Call CU_Form.Initialize(size)
                    CU_Form.Show vbModal
                    If CU_Form.IsCancelled Then
                        ThisWorkbook.sheets("Control").Activate
                        End
                    End If
                    If CU_Form.OptionButton1 Then
                        topSideTie = "TOP"
                    ElseIf CU_Form.OptionButton2 Then
                        topSideTie = "SIDE"
                    End If
                End If
                cuCode = CUNameMapping.getCUNameMapping(Utilities.OnlyNumbers(size) & topSideTie & "TIE")
            ElseIf InStr(hardware, "SCORS") > 0 Then
                cuCode = CUNameMapping.getCUNameMapping(Utilities.OnlyNumbers(size) & "SIDETIE")
            End If
            
            If cuCode <> "" Then
                Call generateCU(cus, pole.location, cuCode, amount, action)
                Call needAdditionalCUs.Remove(i)
            End If
        Next i
    End If
    
    For i = 1 To needAdditionalCUs.count
        neededCU = needAdditionalCUs(i)
        hardware = Replace(neededCU(0), " ", "")
        amount = neededCU(1)
        action = neededCU(2)
        
        If InStr(hardware, "PRI") > 0 And InStr(hardware, "DE") > 0 Then
            missedLines.Add action & IIf(amount <> 1, " (" & amount & ")", " ") & hardware & " PRIMARY GRIP CU MISSING"
        ElseIf InStr(hardware, "NEUT") > 0 And InStr(hardware, "DE") > 0 Then
            missedLines.Add action & IIf(amount <> 1, " (" & amount & ")", " ") & hardware & " NEUTRAL GRIP CU MISSING"
        ElseIf InStr(hardware, "SEC") > 0 And InStr(hardware, "DE") > 0 Then
            missedLines.Add action & IIf(amount <> 1, " (" & amount & ")", " ") & hardware & " SECONDARY DE CU MISSING"
        ElseIf InStr(hardware, "WR") > 0 Or InStr(hardware, "1VPO") > 0 Or InStr(hardware, "2VPO") > 0 Or InStr(hardware, "3VPO") > 0 Then
            missedLines.Add action & IIf(amount <> 1, " (" & amount & ")", " ") & hardware & " SPOOL TIE CU MISSING"
        ElseIf InStr(hardware, "PTP") > 0 Or InStr(hardware, "SPINS") > 0 Then
            missedLines.Add action & IIf(amount <> 1, " (" & amount & ")", " ") & hardware & " TOP/SIDE TIE CU MISSING"
        ElseIf InStr(hardware, "SCORS") > 0 Then
            missedLines.Add action & IIf(amount <> 1, " (" & amount & ")", " ") & hardware & " SIDE TIE CU MISSING"
        End If
    Next i
End Sub

Public Sub SortCollectionByAction(col As Collection)
    Dim arr() As Variant
    Dim i As Long, j As Long
    Dim temp As Variant

    ReDim arr(1 To col.count)
    For i = 1 To col.count
        arr(i) = col(i)
    Next i

    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If StrComp(arr(i)(2), arr(j)(2), vbTextCompare) < 0 Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i

    Do While col.count > 0
        col.Remove 1
    Loop

    For i = 1 To UBound(arr)
        col.Add arr(i)
    Next i
End Sub

Private Sub getSpoolTies(cus As Collection, pole As pole, needAdditionalCUs As Collection, neutSizes As scripting.Dictionary, neutCount As Integer, secSizes As scripting.Dictionary, secCount As Integer, owSizes As scripting.Dictionary, owCount As Integer, amount As Integer, action As String)
    Dim uniqueSizes As scripting.Dictionary: Set uniqueSizes = New scripting.Dictionary
    Dim totalWires As Integer
    Dim failed As Boolean
    
    totalWires = neutCount + secCount + owCount
    
    For Each size In neutSizes
        sizeNumber = Utilities.OnlyNumbers(CStr(size))
        If Not uniqueSizes.Exists(sizeNumber) Then
            uniqueSizes(sizeNumber) = 0
        End If
        uniqueSizes(sizeNumber) = uniqueSizes(sizeNumber) + 1
    Next size
    
    For Each size In secSizes
        sizeNumber = Utilities.OnlyNumbers(CStr(size))
        If Not uniqueSizes.Exists(sizeNumber) Then
            uniqueSizes(sizeNumber) = 0
        End If
        uniqueSizes(sizeNumber) = uniqueSizes(sizeNumber) + 1
    Next size
    
    For Each size In owSizes
        sizeNumber = Utilities.OnlyNumbers(CStr(size))
        If Not uniqueSizes.Exists(sizeNumber) Then
            uniqueSizes(sizeNumber) = 0
        End If
        uniqueSizes(sizeNumber) = uniqueSizes(sizeNumber) + 1
    Next size
    
    Dim cuCode As String
    
    If uniqueSizes.count = 1 Then
        cuCode = CUNameMapping.getCUNameMapping(sizeNumber & " SPOOL TIE")
        If cuCode <> "" Then
            Call generateCU(cus, pole.location, cuCode, amount, action)
            For i = needAdditionalCUs.count To 1 Step -1
                neededCU = needAdditionalCUs(i)
                hardware = Replace(neededCU(0), " ", "")
                If action = neededCU(2) And (InStr(hardware, "WR") > 0 Or InStr(hardware, "1VPO") > 0 Or InStr(hardware, "2VPO") > 0 Or InStr(hardware, "3VPO") > 0) Then Call needAdditionalCUs.Remove(i)
            Next i
            If action = "RET REM" Then
                If neutCount = amount * 2 And secCount = 0 And owCount = 0 Then
                    Call neutSizes.RemoveAll
                    neutCount = 0
                ElseIf neutCount = 0 And secCount = amount * 2 And owCount = 0 Then
                    Call secSizes.RemoveAll
                    secCount = 0
                ElseIf neutCount = 0 And secCount = 0 And owCount = amount Then
                    Call owSizes.RemoveAll
                    owCount = 0
                ElseIf neutCount + secCount + owCount = amount * 2 Then
                    Call neutSizes.RemoveAll
                    Call secSizes.RemoveAll
                    Call owSizes.RemoveAll
                    neutCount = 0
                    secCount = 0
                    owCount = 0
                End If
            End If
        End If
    ElseIf totalWires = amount Then
        For Each size In neutSizes
            sizeNumber = Utilities.OnlyNumbers(CStr(size))
            cuCode = CUNameMapping.getCUNameMapping(sizeNumber & " SPOOL TIE")
            If cuCode <> "" Then
                Call generateCU(cus, pole.location, cuCode, neutSizes(size), action)
            Else
                failed = True
            End If
        Next size
        For Each size In secSizes
            sizeNumber = Utilities.OnlyNumbers(CStr(size))
            cuCode = CUNameMapping.getCUNameMapping(sizeNumber & " SPOOL TIE")
            If cuCode <> "" Then
                Call generateCU(cus, pole.location, cuCode, secSizes(size), action)
            Else
                failed = True
            End If
        Next size
        For Each size In owSizes
            sizeNumber = Utilities.OnlyNumbers(CStr(size))
            cuCode = CUNameMapping.getCUNameMapping(sizeNumber & " SPOOL TIE")
            If cuCode <> "" Then
                Call generateCU(cus, pole.location, cuCode, owSizes(size), action)
            Else
                failed = True
            End If
        Next size
        If action = "RET REM" And Not failed Then
            Call neutSizes.RemoveAll
            Call secSizes.RemoveAll
            Call owSizes.RemoveAll
            neutCount = 0
            secCount = 0
            owCount = 0
            For i = needAdditionalCUs.count To 1 Step -1
                neededCU = needAdditionalCUs(i)
                hardware = Replace(neededCU(0), " ", "")
                If action = neededCU(2) And (InStr(hardware, "WR") > 0 Or InStr(hardware, "1VPO") > 0 Or InStr(hardware, "2VPO") > 0 Or InStr(hardware, "3VPO") > 0) Then Call needAdditionalCUs.Remove(i)
            Next i
        End If
    End If
End Sub

Private Sub getExtraDECU(cus As Collection, pole As pole, needAdditionalCUs As Collection, index As Integer, sizes As scripting.Dictionary, sizeCount As Integer, neededCU() As Variant, Optional componentType As String)
    Dim hardware As String: hardware = neededCU(0)
    Dim amount As Integer: amount = neededCU(1)
    Dim action As String: action = neededCU(2)
    Dim cuCode As String
    Dim amountUsed As Integer
    Dim singleSize As Boolean
    
    If sizes.count = 1 Or sizeCount = amount Then
        If sizes.count = 1 Then singleSize = True
        For Each size In sizes
            cuCode = CUNameMapping.getCUNameMapping(size & "DE")
            If cuCode = "" Then cuCode = CUNameMapping.getCUNameMapping(size & "DEGRIP")
            If cuCode <> "" Then
                Call generateCU(cus, pole.location, cuCode, IIf(Not singleSize, sizes(size), amount), action)
                amountUsed = amountUsed + IIf(Not singleSize, sizes(size), amount)
                If action = "RET REM" Then
                    sizeCount = sizeCount - IIf(Not singleSize, sizes(size), amount)
                    sizes(size) = sizes(size) - IIf(Not singleSize, sizes(size), amount)
                    If sizes(size) = 0 Then Call sizes.Remove(size)
                End If
            End If
        Next size
        If amountUsed = amount Then Call needAdditionalCUs.Remove(index)
    Else
        If componentType = "PRI" Then Exit Sub
        'find the ONLY size with an odd number of spans (must be the deadend)
        oddCount = 0
        Dim oddSize As String
        For Each size In sizes
            If sizes(size) Mod 2 > 0 Then
                oddCount = oddCount + 1
                oddSize = size
            End If
        Next size
        If oddCount = 1 And amount = 1 Then
            cuCode = CUNameMapping.getCUNameMapping(oddSize & "DEGRIP")
            If cuCode <> "" Then
                Call generateCU(cus, pole.location, cuCode, amount, action)
                Call needAdditionalCUs.Remove(index)
                If action = "RET REM" Then
                    sizeCount = sizeCount - 1
                    sizes(oddSize) = sizes(oddSize) - 1
                    If sizes(oddSize) = 0 Then Call sizes.Remove(oddSize)
                End If
            End If
        End If
    End If
End Sub

Private Sub generateCSV(Project As Project, cus As Collection)
    Dim CU As Variant
    Dim filePath As String
    
    filePath = ThisWorkbook.path & "\" & Project.Notification & " - " & "cus.csv"
    If InStr(filePath, "sharepoint") > 0 Then filePath = Environ("USERPROFILE") & "\Downloads\" & Project.Notification & " - " & "cus.csv"
    
    Call CheckAndCloseWorkbook(filePath)
    
    FileNumber = FreeFile
    Open filePath For Output As #FileNumber

    Print #FileNumber, "Location, CU, QTY, ACTION, CMPLX"
    For Each CU In cus
        If TypeOf CU Is CU Then
            Print #FileNumber, CU.location & "," & CU.code & "," & CU.qty & "," & CU.action & ", "
        Else
            Print #FileNumber, CU(0) & ", , , ," & CU(1)
        End If
    Next CU

    Close #FileNumber
    
    Application.ScreenUpdating = False
    
    Dim csvWb As Workbook
    Set csvWb = Workbooks.Open(filePath)
    Dim csvWs As Worksheet
    Set csvWs = csvWb.sheets(1)
    
    Dim foundCell As Range
    Set cuSortWs = ThisWorkbook.sheets("CUSortOrder")
    
    For Each cell In csvWs.UsedRange.Columns(2).Cells
        If Trim(cell.Value) <> "CU" Then
            If Trim(cell.offset(0, 3).Value) <> "" Then
                cell.offset(0, 4).Value = 0.1
            Else
                Set foundCell = cuSortWs.UsedRange.find(what:=cell.Value, LookIn:=xlValues, lookat:=xlWhole)
                If Not foundCell Is Nothing Then cell.offset(0, 4).Value = foundCell.offset(0, 1)
            End If
        End If
    Next cell
    
    With csvWs.Sort
        .SortFields.Clear
        .SortFields.Add key:=csvWs.Range("A1"), Order:=xlAscending
        .SortFields.Add key:=csvWs.Range("F1"), Order:=xlAscending
        .SortFields.Add key:=csvWs.Range("D1"), Order:=xlDescending
        .SetRange csvWs.UsedRange
        .header = xlYes
        .Apply
    End With
    
    'csvWs.Columns(6).Delete
    
    csvWb.save
    csvWb.Close

    Application.ScreenUpdating = True
End Sub

Private Sub generateMissedLinesTXT(missedLines As Collection)
    Dim issues As String
    Dim Project As Project: Set Project = New Project
    Call Project.extractFromSheets
    
    issues = "Lines unable to turn into CUS." & vbLf & Utilities.JoinCollection(missedLines, vbLf)
    filePath = ThisWorkbook.path & "\" & Project.Notification & " - " & "MissedLineCUs.txt"
    If InStr(filePath, "sharepoint") > 0 Then filePath = Environ("USERPROFILE") & "\" & Project.Notification & " - " & "MissedLineCUs.txt"
    
    fNum = FreeFile
    Open filePath For Output As #fNum
    Print #fNum, issues
    Close #fNum
    
    Shell "notepad.exe """ & filePath & """", vbNormalFocus
End Sub

Private Sub parseLineToCUs(needAdditionalCUs As Collection, missedLines As Collection, cus As Collection, pole As pole, ByVal line As String, mode As String)
    Dim regex As Object: Set regex = CreateObject("VBScript.RegExp")
    Dim regex2 As Object: Set regex2 = CreateObject("VBScript.RegExp")
    Dim amount As Integer: amount = 1
    Dim hardware As String: hardware = Trim(line)
    Dim CU As CU, otherCu As CU
    Dim cuCode As String
    
    addedCU = False
    vpoPole = False
    
    regex.Pattern = "\((\d+)\)(.+)"
    regex.Global = True
    regex.IgnoreCase = True
    If regex.test(line) Then
        Set matches = regex.Execute(line)
        amount = matches(0).SubMatches(0)
        hardware = Trim(ThisWorkbook.RemoveParentheses(matches(0).SubMatches(1)))
    End If
    
    'Replace section handler
    If mode = "Replace" Then
        Dim line1 As String, line2 As String
        line1 = line
        line2 = line
        
        regex.Pattern = "(\d[0:5]\s*-\s*\d)\s*\/\s*(\d[0:5]\s*-\s*\d)"
        regex.Global = True
        regex.IgnoreCase = True
        
        regex2.Pattern = "(\d[0:5]\s*-\s*\d)\s*"
        regex2.Global = True
        regex2.IgnoreCase = True
        
        If regex.test(line) Then
            Set matches = regex.Execute(line)
            line1 = Trim(matches(0).SubMatches(0))
            line2 = Trim(matches(0).SubMatches(1))
            If pole.primaries.count > 0 Then hotsite = True
            If timeAdder = 1 Then timeAdder = 2
            If pole.buck Then timeAdder = 3
            For Each equipment In pole.equipments
                If equipment.componentType = "XFMR" Then timeAdder = 3
            Next equipment
        ElseIf InStr(line, "SERVICE RISER") > 0 And InStr(line, "|C") > 0 Then
            line1 = Left(line, InStr(line, "SERVICE RISER") - 1) & "RISER"
            line2 = Left(line, InStr(line, "SERVICE RISER") - 1) & "RISER"
        ElseIf InStr(line, "SECONDARY RISER") > 0 And InStr(line, "|C") > 0 Then
            line1 = Left(line, InStr(line, "SECONDARY RISER") - 1) & "RISER"
            line2 = Left(line, InStr(line, "SECONDARY RISER") - 1) & "RISER"
        ElseIf InStr(line, "/") > 0 Then
            parts = Split(line, "/")
            line1 = Trim(parts(0))
            line2 = Trim(parts(1))
        ElseIf regex2.test(line) And InStr(line, "FIGURE") = 0 Then
            Set matches = regex2.Execute(line)
            If matches.count = 1 Then
                line1 = Trim(matches(0).SubMatches(0))
                line2 = line1
            End If
        End If
        
        If InStr(Replace(Replace(hardware, " ", ""), "DEADEND", "DE"), "SERVICEDE") = 1 Then serviceAmount = serviceAmount + amount
        If InStr(Replace(Replace(Replace(hardware, " ", ""), "DEADEND", "DE"), "OPENWIRE", "OW"), "OWSERVICEDE") = 1 Then serviceAmount = serviceAmount + amount
        
        If InStr(line, "11K") = 0 And InStr(line, "20K") = 0 Then guySection = False
        If (InStr(line, "11K") > 0 Or InStr(line, "20K") > 0) And InStr(line, "/") > 0 And InStr(line, "XP") = 0 And InStr(line, "XFG") = 0 Then guySection = True
    
        If guySection And InStr(line, "/") = 0 Then
            Do While InStr(line, "  ") > 0 ' Find if there are any double spaces
                line = Replace(line, "  ", " ") ' Replace all double spaces with a single space
            Loop
            If Left(line, 1) = " " Then
                line1 = ""
                line2 = line
            ElseIf InStr(line, " ") = 0 Then
                line1 = line
                line2 = line
            Else
                parts = Split(line, " ")
                If UBound(parts) = 1 Then
                    line1 = parts(0)
                    line2 = parts(1)
                Else
                    line1 = line
                    line2 = line
                End If
            End If
        End If
    
        If InStr(line, "DEEP") > 0 And InStr(line, "SET") > 0 And InStr(line, "'") > 0 Then
            temp = Utilities.OnlyNumbers(Mid(line, InStr(line, "'") - 1, 1))
            If IsNumeric(temp) Then Call generateCU(cus, pole.location, "100041", CInt(temp), "INSTALL")
            line1 = ""
            line2 = ""
        End If
    
        If InStr(line1, "'") > 0 And (InStr(line1, "OPEN WIRE") > 0 Or InStr(line1, "SECONDARY") > 0) And InStr(line2, "SECONDARY") > 0 Then
            Call generateReconductorCUs(cus, pole, line1, line2)
            line1 = ""
            line2 = ""
        ElseIf InStr(line, "CO") > 0 And (InStr(line, " ON SA") > 0 Or InStr(line, " ON LCOM") > 0 Or InStr(line, " TO SA") > 0 Or InStr(line, " TO LCOM") > 0) Then
            Call generateTransferCOCU(cus, pole.location, amount, hardware)
            line1 = ""
            line2 = ""
        ElseIf InStr(line, "|LA ") > 0 Or InStr(hardware, "LA ") = 1 And (InStr(line, "TRANSFORMER") = 0 Or InStr(line, "TO TRANSFORMER") > 0) Then
            Call generateCU(cus, pole.location, "200155", amount, "INSTALL")
            line1 = ""
            line2 = ""
        Else
            If line1 <> "" Then Call parseLineToCUs(needAdditionalCUs, missedLines, cus, pole, line1, "Remove")
            If line2 <> "" Then Call parseLineToCUs(needAdditionalCUs, missedLines, cus, pole, line2, "Install")
        End If
        
    Else
        line = Replace(line, ",", "+")
        line = Replace(line, "&", "+")
        If InStr(line, "+") > 0 Then
            parts = Split(line, "+")
            For i = 0 To UBound(parts)
                Call parseLineToCUs(needAdditionalCUs, missedLines, cus, pole, parts(i), mode)
            Next i
            Exit Sub
        End If
    
        If InStr(hardware, " TO ALLOW") > 0 Then hardware = Left(hardware, InStr(hardware, " TO ALLOW") - 1)
        If InStr(hardware, " TO CORRECT") > 0 Then hardware = Left(hardware, InStr(hardware, " TO CORRECT") - 1)
        If InStr(hardware, " TO UPGRADE") > 0 Then hardware = Left(hardware, InStr(hardware, " TO UPGRADE") - 1)
        If InStr(hardware, " DUE TO") > 0 Then hardware = Left(hardware, InStr(hardware, " DUE TO") - 1)
        If InStr(hardware, "@") > 0 And InStr(hardware, "LIGHT") = 0 Then hardware = Left(hardware, InStr(hardware, "@") - 1)
        If InStr(hardware, " @") > 0 And InStr(hardware, "LIGHT") = 0 Then hardware = Left(hardware, InStr(hardware, " @") - 1)
        
        'Riser Install/Remove
        If mode <> "Transfer" Then
            If InStr(line, "PRIMARY RISER") > 0 And InStr(line, "|C") > 0 Then
                Call generatePrimaryRiserCU(cus, pole, hardware, properAction(mode))
            ElseIf InStr(line, "RISER") > 0 And InStr(line, "|C") > 0 And InStr(line, "PRIMARY") = 0 Then
                Call generateSecondaryRiserCU(cus, pole, hardware, properAction(mode))
            End If
            regex2.Pattern = "(\d[0:5]\s*-\s*\d)\s*"
            regex2.Global = True
            regex2.IgnoreCase = True
            If regex2.test(hardware) And InStr(hardware, "FIGURE") = 0 Then
                Set matches = regex2.Execute(line)
                If matches.count = 1 Then
                    hardware = Trim(matches(0).SubMatches(0))
                End If
            End If
             
            If (InStr(hardware, "STREET") > 0 Or InStr(hardware, "FLOOD") > 0) And InStr(hardware, "LIGHT") > 0 And InStr(hardware, "MOLDING") > 0 Then
                If pole.replacePole Then
                    streetlightMolding = streetlightMolding & mode
                    addedCU = True
                Else
                    Call generateReplaceStreetlightMoldingCU(cus, pole, hardware, properAction(mode), missedLines)
                End If
            End If
            
            If InStr(hardware, "MIDSPAN") > 0 And InStr(hardware, "TAP") > 0 Then
                Call generateCU(cus, pole.location, "100196", amount, properAction(mode))
            End If
        End If
        
        'Guy handler
        If (InStr(line, "11K") > 0 Or InStr(line, "20K") > 0) And InStr(line, "XP") = 0 And InStr(line, "XFG") = 0 Then
            If InStr(line, "SPAN GUY") > 0 Or InStr(line, "SPANGUY") > 0 Then
                If mode = "Transfer" Then
                    Call generateCU(cus, pole.location, "106121", amount, "INSTALL")
                    Call generateCU(cus, pole.location, "505040", amount, "RET REM")
                    Call generateCU(cus, pole.location, "505040", amount, "INSTALL")
                Else
                    If mode = "Install" Then Call generateCU(cus, pole.location, "106121", amount, "INSTALL")
                    Call generateCU(cus, pole.location, "505040", amount, properAction(mode))
                End If
            Else
                If mode = "Install" Or mode = "Remove" Then
                    Call generateGuyCU(cus, pole.location, hardware, amount, properAction(mode))
                End If
            End If
        ElseIf mode = "Transfer" And (InStr(line, "SPAN GUY") > 0 Or InStr(line, "SPANGUY") > 0) Then
            Call generateCU(cus, pole.location, "106121", amount, "INSTALL")
            Call generateCU(cus, pole.location, "505040", amount, "RET REM")
            Call generateCU(cus, pole.location, "505040", amount, "INSTALL")
        End If
        
        'Install section handler
        If mode = "Install" Then
            If InStr(line, "BOND STREETLIGHT") > 0 Then Call generateCU(cus, pole.location, "100144", amount, "INSTALL")
        End If
        
        'Remove section handler
        If mode = "Remove" Then
            If InStr(line, "FIRE") > 0 And InStr(line, "WIRE") > 0 Then Call generateCU(cus, pole.location, "201389", amount, "INSTALL")
        End If
        
        'Transfer section handler
        If mode = "Transfer" Then
            If InStr(line, "TRANSFORMER") > 0 And InStr(UCase(line), "KVA") > 0 Then Call generateTransferTransformerCU(cus, pole.location, amount)
            If (InStr(line, "STREET") > 0 Or InStr(hardware, "FLOOD") > 0) And InStr(line, "LIGHT") > 0 And InStr(line, "@") > 0 Then Call generateTransferStreetlightCU(line, cus, pole, missedLines)
            If InStr(line, "TRIM") > 0 And InStr(line, "DRIP") Then Call generateCU(cus, pole.location, "101023", 1, "INSTALL")
            If InStr(line, "CO") > 0 And (InStr(line, " ON SA") > 0 Or InStr(line, " ON LCOM") > 0 Or InStr(line, " TO SA") > 0 Or InStr(line, " TO LCOM") > 0) Then Call generateTransferCOCU(cus, pole.location, amount, hardware)
            If InStr(line, "|LA ") > 0 Or InStr(hardware, "LA ") = 1 And InStr(line, "TRANSFORMER") = 0 Then Call generateCU(cus, pole.location, "200155", amount, "INSTALL")
            If Replace(hardware, " ", "") = "SERVICE" Or hardware = "SERVICES" Or Replace(hardware, " ", "") = "OHSERVICE" Or Replace(hardware, " ", "") = "OHSERVICE" Then Call generateTransferServiceCU(cus, pole, amount)
        End If
        
        'Note section handler
        If mode = "Note" Then
            If InStr(line, "TOP") > 0 And InStr(line, "POLE") > 0 And InStr(line, "ABOVE") > 0 Then Call generateCU(cus, pole.location, "100910", 1, "INSTALL")
            If InStr(line, "DEEPSET") > 0 And InStr(line, "'") > 0 Then
                temp = Utilities.OnlyNumbers(Mid(InStr(line, "'") - 1, 1))
                If IsNumeric(temp) Then Call generateCU(cus, pole.location, "100041", CInt(temp), "INSTALL")
            End If
        End If
        
        'Get CU and check if it needs additional CUs
        If Not addedCU Then
            If InStr(hardware, "VPO") > 0 Then
                hardware = amount & hardware
                vpoPole = True
            End If
            cuCode = CUNameMapping.getCUNameMapping(hardware)
            If cuCode = "" Then cuCode = CUNameMapping.getCUNameMapping(Utilities.OnlyLetters(hardware))
            If cuCode = "" And InStr(hardware, "SWAMP FIXTURE") > 0 Then cuCode = "100085"
            If cuCode <> "" Then Call generateCU(cus, pole.location, cuCode, amount, properAction(mode))
            
            If CUNameMapping.CheckForAdditionalCUs(hardware) Then needAdditionalCUs.Add Array(hardware, amount, properAction(mode))
        End If
    End If
    
    'Add missed lines
    If Not addedCU And mode <> "Note" Then
        If (mode = "Replace" And (line1 = "" Or line2 = "")) Or mode <> "Replace" Then
            If Not MissedLineIgnorable(pole, hardware) Then missedLines.Add mode & " " & hardware
        End If
    End If
End Sub

Private Sub generateReplaceStreetlightMoldingCU(cus As Collection, pole As pole, hardware As String, mode As String, missedLines As Collection)
    Dim distance As Integer
    Dim cuCode As String
    Dim closestDistance As Integer

    If pole.slBottomBracketHeight > 1 Then streetlightBottomBracketHeight = pole.slBottomBracketHeight

    If InStr(hardware, "'") > 0 Then distance = Utilities.OnlyNumbers(Left(hardware, InStr(hardware, "'") + 1))
    
    If distance > 0 Then
        Call generateCU(cus, pole.location, "100598", Application.WorksheetFunction.RoundUp(distance / 8, 0), mode)
    Else
        closestDistance = 0
        For Each Wire In pole.utilWires
            If Wire.componentType = "OW" Or Wire.componentType = "SEC" Or Wire.componentType = "TRAFFIC" Then
                If Abs(Wire.height - streetlightBottomBracketHeight) < closestDistance Or closestDistance = 0 Then closestDistance = Abs(Wire.height - streetlightBottomBracketHeight)
            End If
        Next Wire
        Call generateCU(cus, pole.location, "100598", WorksheetFunction.RoundUp((closestDistance / 12) / 8, 0), mode)
    End If
End Sub

Private Sub generateTransferServiceCU(cus As Collection, pole As pole, amount As Integer)
    Dim cuCode As String
    Dim totalServices As Integer
    Dim serviceDict As scripting.Dictionary: Set serviceDict = New scripting.Dictionary
    
    For Each service In pole.services
        For Each midspan In service.midspans
            If Not serviceDict.Exists(midspan) Then serviceDict.Add midspan, Nothing
            totalServices = totalServices + 1
        Next midspan
    Next service
    
    Call generateCU(cus, pole.location, 106115, serviceDict.count, "INSTALL")
    
    If serviceAmount = totalServices Then
        For Each service In pole.services
            cuCode = CUNameMapping.getCUNameMapping(service.size & "DE")
            If cuCode <> "" Then
                For Each midspan In service.midspans
                    Call generateCU(cus, pole.location, cuCode, 1, "RET REM")
                    Call generateCU(cus, pole.location, cuCode, 1, "INSTALL")
                Next midspan
            End If
        Next service
    End If
End Sub

Private Sub generateReconductorCUs(cus As Collection, pole As pole, line1 As String, line2 As String)
    Dim regex As Object: Set regex = CreateObject("VBScript.RegExp")
    
    Dim distance As Integer
    Dim cuCode As String
    Dim secSize As String
    
    reconductored = False
    
    If InStr(line1, "'") > 0 Then distance = Utilities.OnlyNumbers(Left(line1, InStr(line1, "'") + 1))
    If distance = -1 Then distance = 0
    
    If InStr(line1, "OPEN WIRE") > 0 Then
        owSizes = Mid(line1, InStr(line1, "'") + 1, InStr(line1, "OPEN WIRE") - InStr(line1, "'") - 1)
        
        parts = Split(owSizes, "-")
        For i = 0 To UBound(parts)
            owSize = Left(Trim(parts(i)), 1)
            cuCode = CUNameMapping.getOWNameMapping(owSize)
            If cuCode <> "" Then
                Call generateCU(cus, pole.location, "290048", 1, "INSTALL")
                Call generateCU(cus, pole.location, cuCode, distance, "RET REM")
                reconductored = True
            End If
        Next i
    ElseIf InStr(line1, "SECONDARY") > 0 Then
        secSize = Left(line1, InStr(line1, "SECONDARY") - 1)
        If InStr(secSize, "'") > 0 Then secSize = Mid(secSize, InStr(secSize, "'") + 1, Len(secSize) - InStr(secSize, "'") - 1)
        cuCode = CUNameMapping.getSecNameMapping(secSize)
        If cuCode <> "" Then
            Call generateCU(cus, pole.location, "290048", 1, "INSTALL")
            Call generateCU(cus, pole.location, cuCode, distance, "RET REM")
            reconductored = True
        End If
    End If
    
    secSize = Left(line2, InStr(line2, "SECONDARY") - 1)
    If InStr(secSize, "'") > 0 Then secSize = Mid(secSize, InStr(secSize, "'") + 1, Len(secSize) - InStr(secSize, "'") - 1)
    cuCode = CUNameMapping.getSecNameMapping(secSize)
    If cuCode <> "" Then
        Call generateCU(cus, pole.location, "290048", 1, "INSTALL")
        Call generateCU(cus, pole.location, cuCode, distance, "INSTALL")
        reconductored = True
    Else
        addedCU = False
    End If
    
    If reconductored = True And (pole.secondaries.count > 0 Or pole.services.count > 0) Then
        Call generateCU(cus, pole.location, "290061", 1, "INSTALL")
    End If
End Sub

Private Sub checkForAdjacentPoleRecondcutoring(cus As Collection, Project As Project, pole As pole, missedLines As Collection)
    Dim Span As Span
    Dim otherPole As pole
    Dim otherSpan As Span
    Dim count As Integer
    Dim lines() As String
    For Each Span In pole.spans
        If Span.otherPole <> "" Then
            Set otherPole = Project.findPole(Span.otherPole)
            If InStr(otherPole.Alt1, "'") > 0 And InStr(line, "OPEN WIRE") > 0 And InStr(line, "SECONDARY") > 0 Then
                If InStr(otherPole.Alt1, vbLf) > 0 Then
                    lines = Split(otherPole.Alt1, vbLf)
                    For Each line In lines
                        If InStr(line, "'") > 0 And InStr(line, "OPEN WIRE") > 0 And InStr(line, "SECONDARY") > 0 Then
                            distance = Utilities.OnlyNumbers(Left(line, InStr(line, "'")))
                            If IsNumeric(distance) Then
                                If CInt(distance) = Span.distance Then
                                    For Each otherSpan In otherPole.spans
                                        If Span.distance = otherSpan.distance And otherSpan.otherPole <> "" Then count = count + 1
                                    Next otherSpan
                                    If count = 1 Then
                                        Call generateCU(cus, pole.location, "290048", 1, "INSTALL")
                                    Else
                                        missedLines.Add "AMBIGUOUS SPAN LENGTHS ON OTHER POLE FOR RECONDUCTORING"
                                    End If
                                End If
                            End If
                            Exit For
                        End If
                    Next line
                End If
            End If
        End If
    Next Span
End Sub

Private Function MissedLineIgnorable(pole As pole, ByVal line As String) As Boolean
    MissedLineIgnorable = False
    line = Replace(ThisWorkbook.RemoveParentheses(line), " ", "")
    If InStr(line, "(") > 0 Then line = Left(line, InStr(line, "(") - 1)
    If InStr(line, ")") > 0 Then line = Left(line, InStr(line, ")") - 1)
    If InStr(line, "FIGURE") = 1 Then MissedLineIgnorable = True: Exit Function
    If InStr(line, "D=") = 1 Then MissedLineIgnorable = True: Exit Function
    If InStr(line, "P=") = 1 Then MissedLineIgnorable = True: Exit Function
    If InStr(line, "TOCORRECT") = 1 Then MissedLineIgnorable = True: Exit Function
    If InStr(line, "TOMAKE") = 1 Then MissedLineIgnorable = True: Exit Function
    If InStr(line, "TOALLOW") = 1 Then MissedLineIgnorable = True: Exit Function
    If InStr(line, "@11""FROM") = 1 Then MissedLineIgnorable = True: Exit Function
    If InStr(Replace(line, "DEADEND", "DE"), "SECONDARYDE") = 1 Then MissedLineIgnorable = True: Exit Function
    If InStr(line, "AS-IS") > 0 Then MissedLineIgnorable = True: Exit Function
    If InStr(line, "ASIS") > 0 Then MissedLineIgnorable = True: Exit Function
    If Replace(line, vbLf, "") = "" Then MissedLineIgnorable = True: Exit Function
    If line = "PRIMARY" Then MissedLineIgnorable = True: Exit Function
    If line = "NEUTRAL" Then MissedLineIgnorable = True: Exit Function
    If line = "SECONDARY" Then MissedLineIgnorable = True: Exit Function
    If line = "EXTENDTONEWHEIGHT" Then MissedLineIgnorable = True: Exit Function
    If line = "OPENWIRE" Then MissedLineIgnorable = True: Exit Function
    If line = "OW" Then MissedLineIgnorable = True: Exit Function
    If line = "LAONTRANSFORMER" Then MissedLineIgnorable = True: Exit Function
    Dim comp As Variant
    For Each comp In pole.commComponents
        If InStr(line, Replace(Replace(comp.owner, " ", ""), "&", "")) = 1 Then MissedLineIgnorable = True: Exit Function
    Next comp
End Function

Private Sub generateCU(cus As Collection, location As String, code As String, qty As Integer, action As String)
    Dim CU As CU: Set CU = New CU
    CU.location = properLocation(location)
    CU.code = code
    CU.qty = qty
    CU.action = action
    Call AddCu(cus, CU)
End Sub

Private Sub generateGuyCU(cus As Collection, location As String, hardware As String, qty As Integer, action As String)
    If InStr(hardware, "-RS") > 0 Then Call generateCU(cus, location, "100131", qty, action)
    If InStr(hardware, "-RT") > 0 Then Call generateCU(cus, location, "100133", qty, action)
    If InStr(hardware, "-STE") > 0 Then Call generateCU(cus, location, "100136", qty, action)
    
    If InStr(hardware, "11K") > 0 Then Call generateCU(cus, location, "100421", qty, action)
    If InStr(hardware, "20K") > 0 Then Call generateCU(cus, location, "100422", qty, action)
    
    Dim pQty As Integer
    If InStr(hardware, "P") > 0 Then
        pQty = qty
        If InStr(hardware, "P") > 1 Then
            If IsNumeric(Mid(hardware, InStr(hardware, "P") - 1, 1)) Then pQty = Mid(hardware, InStr(hardware, "P") - 1, 1)
        End If
        
        If InStr(hardware, "11K") > 0 Then
            Call generateCU(cus, location, "100194", pQty, action)
        ElseIf InStr(hardware, "20K") > 0 Then
            Call generateCU(cus, location, "100195", pQty, action)
        End If
    End If
    
    Dim fgQty As Integer
    If InStr(hardware, "FG") > 0 Then
        hotsite = True
        fgQty = qty
        If InStr(hardware, "FG") > 1 Then
            If IsNumeric(Mid(hardware, InStr(hardware, "FG") - 1, 1)) Then fgQty = Mid(hardware, InStr(hardware, "FG") - 1, 1)
        End If
        
        Call generateCU(cus, location, "100192", fgQty, action)
    End If
End Sub

Private Sub generateTransferCOCU(cus As Collection, location As String, qty As Integer, hardware As String)
    Call generateCU(cus, location, "106122", qty, "INSTALL")
    If InStr(hardware, "ON SA") Then
        Call generateCU(cus, location, "100063", qty, "RET REM")
        If InStr(hardware, "TO LCOM") > 0 Then
            Call generateCU(cus, location, "100160", qty, "INSTALL")
        Else
            Call generateCU(cus, location, "100063", qty, "INSTALL")
        End If
    ElseIf InStr(hardware, "ON LCOM") > 0 Then
        Call generateCU(cus, location, "100160", qty, "RET REM")
        If InStr(hardware, "TO SA") > 0 Then
            Call generateCU(cus, location, "100163", qty, "INSTALL")
        Else
            Call generateCU(cus, location, "100160", qty, "INSTALL")
        End If
    End If
End Sub

Private Sub generateTransferStreetlightCU(line As String, cus As Collection, pole As pole, missedLines As Collection)
    Call generateCU(cus, pole.location, "106132", 1, "INSTALL")
    Dim streetlightBottomBracketHeight As Integer
    Dim amount As Integer
    
    If InStr(line, "@") > 0 Then streetlightBottomBracketHeight = Utilities.convertToInches(Mid(line, InStr(line, "@")))
    If streetlightBottomBracketHeight < 1 Then
        For Each equipment In pole.equipments
            If equipment.componentType = "SL" Then
                streetlightBottomBracketHeight = equipment.bottomHeight
                Exit For
            End If
        Next equipment
    End If
    
    
    If streetlightBottomBracketHeight < 1 Then
        Call generateCU(cus, pole.location, "718146", 0, "INSTALL")
        Call generateCU(cus, pole.location, "718146", 0, " RET REM")
        If streetlightMolding <> "" Then
            If InStr(streetlightMolding, "Remove") > 0 Then Call generateCU(cus, pole.location, "718146", 0, " RET REM")
            If InStr(streetlightMolding, "Install") > 0 Then Call generateCU(cus, pole.location, "100598", 0, "INSTALL")
        End If

        missedLines.Add "Replace 2/C-10 CU STLT quantity not set"
        Exit Sub
    End If
    
    If pole.primaries.count > 0 Then
        amount = WorksheetFunction.RoundUp(((pole.newHeight * 0.9) - 36 - pole.dSpace - pole.pSpace - streetlightBottomBracketHeight) / 12, 0)
    Else
        amount = WorksheetFunction.RoundUp((pole.newHeight - 83 - streetlightBottomBracketHeight) / 12, 0)
    End If
    
    Call generateCU(cus, pole.location, "718146", amount, "INSTALL")
    If InStr(streetlightMolding, "Install") > 0 Then Call generateCU(cus, pole.location, "100598", Application.WorksheetFunction.RoundUp(amount / 8, 0), "INSTALL")

    closestDistance = 0
    For Each Wire In pole.utilWires
        If Wire.componentType = "OW" Or Wire.componentType = "SEC" Or Wire.componentType = "TRAFFIC" Then
            If Abs(Wire.height - streetlightBottomBracketHeight) < closestDistance Or closestDistance = 0 Then closestDistance = Abs(Wire.height - streetlightBottomBracketHeight)
        End If
    Next Wire
    amount = WorksheetFunction.RoundUp(closestDistance / 12, 0)

    Call generateCU(cus, pole.location, "718146", amount, "RET REM")
    If InStr(streetlightMolding, "Remove") > 0 Then Call generateCU(cus, pole.location, "100598", Application.WorksheetFunction.RoundUp(amount / 8, 0), "RET REM")
End Sub

Private Sub generateTransferTransformerCU(cus As Collection, location As String, qty As Integer)
    timeAdder = 3
    Call generateCU(cus, location, "106124", 1, "INSTALL")
    If qty > 1 Then Call generateCU(cus, location, "200548", 1, "INSTALL")
    Call generateCU(cus, location, "200352", 1, "INSTALL")
    Call generateCU(cus, location, "106129", 3 * qty, "INSTALL")
    Call generateCU(cus, location, "100101", 1, "INSTALL")
End Sub

Private Sub generatePrimaryRiserCU(cus As Collection, pole As pole, hardware As String, action As String)
    hotsite = True
    
    Dim cuCode As String
    Dim amount As Integer
    
    If InStr(hardware, "'-") > 0 Then
        If IsNumeric(Utilities.OnlyNumbers(Left(hardware, InStr(hardware, "'-")))) Then
            amount = Utilities.OnlyNumbers(Left(hardware, InStr(hardware, "'-")))
        End If
        hardware = Mid(hardware, InStr(hardware, "'-") + 2, Len(hardware) - InStr(hardware, "'-"))
    End If
    
    If InStr(hardware, "'") > 0 Then
        If IsNumeric(Utilities.OnlyNumbers(Left(hardware, InStr(hardware, "'")))) Then
            amount = Utilities.OnlyNumbers(Left(hardware, InStr(hardware, "'")))
        End If
        hardware = Mid(hardware, InStr(hardware, "'") + 1, Len(hardware) - InStr(hardware, "'"))
    End If
    
    
End Sub

Private Sub generateSecondaryRiserCU(cus As Collection, pole As pole, hardware As String, action As String)
    Dim cuCode As String
    Dim amount As Integer
    
    If InStr(hardware, "'-") > 0 Then
        If IsNumeric(Utilities.OnlyNumbers(Left(hardware, InStr(hardware, "'-")))) Then
            amount = Utilities.OnlyNumbers(Left(hardware, InStr(hardware, "'-")))
        End If
        hardware = Mid(hardware, InStr(hardware, "'-") + 2, Len(hardware) - InStr(hardware, "'-"))
    End If
    
    If InStr(hardware, "'") > 0 Then
        If IsNumeric(Utilities.OnlyNumbers(Left(hardware, InStr(hardware, "'")))) Then
            amount = Utilities.OnlyNumbers(Left(hardware, InStr(hardware, "'")))
        End If
        hardware = Mid(hardware, InStr(hardware, "'") + 1, Len(hardware) - InStr(hardware, "'"))
    End If
    
    cuCode = CUNameMapping.getCUNameMapping(hardware)
    If cuCode <> "" Then
        If action = "INSTALL" Then
            If pole.primaries.count > 0 Then
                If amount = 0 Then amount = WorksheetFunction.RoundUp((pole.newHeight - (pole.newHeight * 0.1) - 36 - pole.dSpace - pole.pSpace) / 12, 0)
            Else
                If amount = 0 Then amount = WorksheetFunction.RoundUp((pole.newHeight - 83) / 12, 0)
            End If
        ElseIf action = "RET REM" Then
            For Each equipment In pole.equipments
                If equipment.componentType = "RISER" Then
                    If amount = 0 Then amount = WorksheetFunction.RoundUp(equipment.height / 12, 0)
                End If
            Next equipment
        End If
        
        Call generateCU(cus, pole.location, "101523", amount, action)
        Call generateCU(cus, pole.location, cuCode, amount + 7, action)
        
        If action = "INSTALL" Then
            cuCode = CUNameMapping.getCUNameMapping(hardware & "SPLICE")
            If cuCode <> "" Then
                Call generateCU(cus, pole.location, cuCode, 1, action)
            End If
            Call generateCU(cus, pole.location, "201365", 1, "INSTALL")
        End If
    End If
End Sub

Private Function properLocation(location As String) As String
    properLocation = "L" & Format(location, "000") & " ALT1"
End Function

Private Function properAction(mode As String) As String
     Select Case mode
        Case "Install", "Transfer", "Note"
            properAction = "INSTALL"
        Case "Remove"
            properAction = "RET REM"
     End Select
End Function

Private Sub AddCu(cus As Collection, CU As CU)
    Dim alreadyExists As Boolean
    Dim otherCu As CU
    For i = 1 To cus.count
        If TypeOf cus(i) Is CU Then
            Set otherCu = cus(i)
            If CU.Equals(otherCu) Then
                If CU.code <> "290048" And CU.code <> "200548" Then otherCu.qty = otherCu.qty + CU.qty
                alreadyExists = True
                Exit For
            End If
        End If
    Next i
    
    addedCU = True
    If Not alreadyExists Then cus.Add CU
End Sub
