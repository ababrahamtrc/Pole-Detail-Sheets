Attribute VB_Name = "CUExporter"
Public addedCU As Boolean
Public guySection As Boolean
Public hotsite As Boolean
Public timeAdder As Integer
Public vpoPole As Boolean
Public serviceAmount As Integer
Public reconductored As Boolean
Public primaryReconductored As Boolean
Public streetlightMolding As String

Public Sub CopyCUImportCode()
    Call LogMessage.SendLogMessage("CopyCUImportCode")

    Dim url As String: url = "https://api.github.com/repos/ElijahRademaker/Automation-tools/contents/cuimport.js"
    Dim file As Object
    Dim strText As String
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.send
 
    If http.status <> 200 Then
        MsgBox "Failed to get cuimport.js from github: " & http.status & vbLf & JsonConverter.ParseJson(http.responseText)("message")
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
    
    Dim project As project: Set project = New project
    Call project.extractFromSheets

    Dim cu As Variant
    Dim cus As Collection: Set cus = New Collection
    Dim demoCus As Collection: Set demoCus = New Collection
    Dim missedlines As Collection: Set missedlines = New Collection
    Dim cusTemp As Collection
    Dim demoCusTemp As Collection
    Dim missedLinesTemp As Collection
    Dim inputCol As Collection
    Dim sheet As Worksheet
    
    For Each sheet In ThisWorkbook.sheets
        If Utilities.IsPDS(sheet) Then
            Set inputCol = ExportSheetCUs(project, sheet)
            If Not inputCol Is Nothing Then
                Set cusTemp = inputCol(1)
                Set demoCusTemp = inputCol(2)
                Set missedLinesTemp = inputCol(3)
                For Each cu In cusTemp
                    cus.Add cu
                Next cu
                For Each cu In demoCusTemp
                    demoCus.Add cu
                Next cu
                For Each line In missedLinesTemp
                    missedlines.Add "Location " & sheet.Range("DL") & ": " & line
                Next line
            End If
        End If
    Next sheet
    
    ThisWorkbook.sheets("Control").Activate
    
    If cus.count > 0 Then
        Call generateCSV(project, cus)
        Call generateCSV(project, demoCus, True)
        If missedlines.count > 0 Then
            Call generateMissedLinesTXT(missedlines)
        Else
            MsgBox "All lines successfully turned into CUs."
        End If
    Else
        MsgBox "No CUs generated."
    End If
End Sub

Public Sub ExportSingleSheetCUs()
    Call LogMessage.SendLogMessage("ExportSingleCUs")

    Dim project As project: Set project = New project
    Call project.extractFromSheets
    
    Dim sheet As Worksheet: Set sheet = ThisWorkbook.ActiveSheet()
    If Not Utilities.IsPDS(sheet) Then
        MsgBox "You need to have a pole detail sheet active to run this script."
        Exit Sub
    End If
    
    Dim cus As Collection
    Dim demoCus As Collection
    Dim missedlines As Collection
    
    Dim inputCol As Collection: Set inputCol = New Collection
    Set inputCol = ExportSheetCUs(project, sheet)
    If Not inputCol Is Nothing Then
        Set cus = inputCol(1)
        Set demoCus = inputCol(2)
        Set missedlines = inputCol(3)
    End If
    
    If Not cus Is Nothing Then
        If cus.count > 0 Then
            Call generateCSV(project, cus)
            If missedlines.count > 0 Then
                MsgBox "Lines unable to turn into CUS." & vbLf & Utilities.JoinCollection(missedlines, vbLf)
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

Private Function ExportSheetCUs(project As project, sheet As Worksheet) As Collection
    Dim installSection As Boolean, replaceSection As Boolean, removeSection As Boolean, transferSection As Boolean
    Dim line As Variant
    Dim lines() As String
    Dim installNotes As String, replaceNotes As String, removeNotes As String, transferNotes As String, notes As String
    Dim cus As Collection: Set cus = New Collection
    Dim demoCus As Collection: Set demoCus = New Collection
    Dim missedlines As Collection: Set missedlines = New Collection
    Dim needAdditionalCUs As Collection: Set needAdditionalCUs = New Collection
    Dim pole As pole: Set pole = New pole
    Call pole.extractFromSheet(sheet)
    
    guySection = False
    hotsite = False
    reconductored = False
    primaryReconductored = False
    vpoPole = False
    serviceAmount = 0
    streetlightMolding = ""
    
    lines = Split(pole.alt1, vbLf)
    
    If pole.location = "" Or UBound(lines) < 1 Then Exit Function
    timeAdder = 1
    
    If Replace(Replace(pole.alt1, "/", ""), "NA", "") = "" Then
        Set ExportSheetCUs = Nothing
        Exit Function
    End If
    
    If pole.commComponents.count > 0 Then
        If project.mode = "SYSTEM IMPROVEMENT" Then
            cus.Add Array(properLocation(pole.location), 1.15)
            demoCus.Add Array(properLocation(pole.location), 1.15)
        Else
            cus.Add Array(properLocation(pole.location), 1.45)
        End If
    Else
        If project.mode = "SYSTEM IMPROVEMENT" Then
            cus.Add Array(properLocation(pole.location), 1)
            demoCus.Add Array(properLocation(pole.location), 1)
        Else
            cus.Add Array(properLocation(pole.location), 1.3)
        End If
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
                If line <> "" Then Call parseLineToCUs(project, needAdditionalCUs, missedlines, cus, demoCus, pole, line, "Install")
            ElseIf replaceSection Then
                If line <> "" Then Call parseLineToCUs(project, needAdditionalCUs, missedlines, cus, demoCus, pole, line, "Replace")
            ElseIf removeSection Then
                If line <> "" Then Call parseLineToCUs(project, needAdditionalCUs, missedlines, cus, demoCus, pole, line, "Remove")
            ElseIf transferSection Then
                If line <> "" Then Call parseLineToCUs(project, needAdditionalCUs, missedlines, cus, demoCus, pole, line, "Transfer")
            Else
                If line <> "" Then Call parseLineToCUs(project, needAdditionalCUs, missedlines, cus, demoCus, pole, line, "Note")
            End If
        End If
    Next i
    
    If project.mode <> "SYSTEM IMPROVEMENT" Then
        If IsNumeric(pole.ttc) Then
            Call generateTTCCU(cus, pole.location, CInt(Utilities.OnlyNumbers(pole.ttc)))
        Else
            missedlines.Add "Missing TTC in pole detail sheet, can't generate TTC CU"
        End If
    End If
 
    Call generateCU(cus, pole.location, "100417", timeAdder, "INSTALL")
    If pole.ReplacePole And pole.primaries.count > 0 Then hotsite = True
    If hotsite Then Call generateCU(cus, pole.location, "106268", 1, "INSTALL")
    
    Call fixCUErrors(cus, needAdditionalCUs, missedlines)
    
    If needAdditionalCUs.count > 0 Then
        sheet.Activate
        sheet.Range("A1").Select
        Call findAdditonalCUs(cus, pole, needAdditionalCUs, missedlines)
    End If
    
    If Not reconductored Then Call checkForAdjacentPoleRecondcutoring(cus, project, pole, missedlines)
    
    Dim outputCol As Collection: Set outputCol = New Collection
    outputCol.Add cus
    outputCol.Add demoCus
    outputCol.Add missedlines
    
    Set ExportSheetCUs = outputCol
End Function

Private Sub fixCUErrors(cus As Collection, needAdditionalCUs As Collection, missedlines As Collection)
    Dim transferSpg As Integer
    Dim removeSpg As Boolean
    Dim cu As Variant
    Dim transferServices As Boolean
    Dim missedLine As String
    Dim hardware As String
    Dim regex As Object: Set regex = CreateObject("VBScript.RegExp")
    
    For i = cus.count To 1 Step -1
        If TypeOf cus(i) Is cu Then
            Set cu = cus(i)
            If cu.code = "106121" Then transferSpg = i
            If cu.code = "505040" And cu.action = "RET REM" Then removeSpg = True
            If cu.code = "106115" Then transferServices = True
            If vpoPole And cu.code = "100052" Then cu.qty = cu.qty - 1
            If cu.code = "100417" And primaryReconductored And cu.qty = 1 Then Call cus.Remove(i)
            If cu.code = "100417" And primaryReconductored And cu.qty > 1 Then cu.qty = 1
            If cu.code = "106268" And primaryReconductored Then Call cus.Remove(i)
        End If
    Next i
    
    For i = needAdditionalCUs.count To 1 Step -1
        neededCU = needAdditionalCUs(i)
        If vpoPole And neededCU(0) = "100052" Then
            neededCU(1) = neededCU(1) - 1
            If neededCU(1) = 0 Then Call needAdditionalCUs.Remove(i)
        End If
    Next i
    
    For i = missedlines.count To 1 Step -1
        missedLine = missedlines(i)
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
            If InStr(hardware, "SVCDE") = 1 Then Call missedlines.Remove(i)
            If InStr(hardware, "OWSVCDE") = 1 Then Call missedlines.Remove(i)
        End If
    Next i
    
    If transferSpg > 0 And Not removeSpg Then Call cus.Remove(transferSpg)
    
    Call condenseAssemblies(cus)
End Sub

Sub condenseAssemblies(cus As Collection)
    Dim CUAssemblyMapping As Scripting.Dictionary
    Dim cuCode As String
    Dim cuQty As Integer
    Dim location As String
    Dim locationAssemblyMap As Scripting.Dictionary
    Dim isAssembly As Boolean
    Dim cu As Variant
    Dim childCu As cu
    Dim minqty As Integer
    Dim action As String
    
    Set CUAssemblyMapping = CUNameMapping.getCUAssemblyMapping

    For Each CUAssemblyCode In CUAssemblyMapping
        assembly = CUAssemblyMapping(CUAssemblyCode)
        Set locationAssemblyMap = New Scripting.Dictionary
        For i = 0 To UBound(assembly)
            cuCode = assembly(i)
            Dim locationActionAssemblyCus() As Long
            ReDim locationActionAssemblyCus(0 To UBound(assembly)) As Long
            For Each cu In cus
                If TypeOf cu Is cu Then
                    If Not locationAssemblyMap.exists(cu.location & "|" & cu.action) Then locationAssemblyMap(cu.location & "|" & cu.action) = locationActionAssemblyCus
                    If cu.code = cuCode Then
                        Dim temp() As Long
                        temp = locationAssemblyMap(cu.location & "|" & cu.action)
                        temp(i) = temp(i) + cu.qty
                        locationAssemblyMap(cu.location & "|" & cu.action) = temp
                    End If
                End If
            Next cu
        Next i
        
        For Each locationAction In locationAssemblyMap
            results = Split(locationAction, "|")
            location = results(0)
            action = results(1)
            
            locationActionAssemblyCus = locationAssemblyMap(locationAction)
            
            isAssembly = True
            minqty = -1
            For i = 0 To UBound(locationActionAssemblyCus)
                cuQty = locationActionAssemblyCus(i)
                If cuQty < 1 Then isAssembly = False: Exit For
                If minqty = -1 Or cuQty < minqty Then minqty = cuQty
            Next i
            
            If isAssembly Then
                For i = cus.count To 1 Step -1
                    If TypeOf cus(i) Is cu Then
                        Set cu = cus(i)
                        For j = 0 To UBound(assembly)
                            cuCode = assembly(j)
                            If cu.code = cuCode Then
                                If minqty > 1 Then cu.qty = cu.qty - minqty
                                If minqty = 1 Or cu.qty < 1 Then Call cus.Remove(i)
                            End If
                        Next j
                    End If
                Next i
                
                If minqty = 1 Then
                    For i = 0 To UBound(assembly)
                        Set childCu = New cu
                        childCu.location = location
                        childCu.parentCU = CUAssemblyCode
                        childCu.childCode = assembly(i)
                        childCu.childQty = locationActionAssemblyCus(i)
                        childCu.parentInstance = 1

                        For Each cu In cus
                            If TypeOf cu Is cu Then
                                If cu.action = action And cu.location = location And cu.code = CUAssemblyCode Then childCu.parentInstance = childCu.parentInstance + 1
                            End If
                        Next cu
                        
                        cus.Add childCu
                    Next i
                End If
                Call generateCU(cus, location, CStr(CUAssemblyCode), minqty, action)
            End If
        Next locationAction
    Next CUAssemblyCode
End Sub

Private Sub generateTTCCU(cus As Collection, location As String, ttc As Integer)
    If (ttc >= 3 And ttc <= 7) Or ttc = 16 Or ttc = 17 Then Call generateCU(cus, location, "999013", 1, "INSTALL")
    If ttc = 8 Or ttc = 10 Or ttc = 11 Or ttc = 18 Or ttc = 19 Or ttc = 22 Then Call generateCU(cus, location, "999014", 1, "INSTALL")
    If (ttc >= 12 And ttc <= 14) Then Call generateCU(cus, location, "999015", 1, "INSTALL")
    If ttc = 15 Or ttc = 21 Then Call generateCU(cus, location, "999017", 1, "INSTALL")
    If ttc = 23 Then Call generateCU(cus, location, "999018", 1, "INSTALL")
End Sub

Private Sub findAdditonalCUs(cus As Collection, pole As pole, needAdditionalCUs As Collection, missedlines As Collection)
    Dim wire As wire
    Dim cuCode As String
    Dim priCount As Integer
    Dim neutCount As Integer
    Dim secCount As Integer
    Dim owCount As Integer
    Dim priSizes As Scripting.Dictionary: Set priSizes = New Scripting.Dictionary
    Dim neutSizes As Scripting.Dictionary: Set neutSizes = New Scripting.Dictionary
    Dim secSizes As Scripting.Dictionary: Set secSizes = New Scripting.Dictionary
    Dim owSizes As Scripting.Dictionary: Set owSizes = New Scripting.Dictionary
    
    Call SortCollectionByAction(needAdditionalCUs)
    
    'Running total of sizes, these go down as they're matched to a CU
    For Each wire In pole.primaries
        If Not priSizes.exists(wire.size) Then priSizes(wire.size) = 0
        For Each midspan In wire.midspans
            priCount = priCount + wire.phase
            priSizes(wire.size) = priSizes(wire.size) + wire.phase
        Next midspan
    Next wire
    For Each wire In pole.neutrals
        If Not neutSizes.exists(wire.size) Then neutSizes(wire.size) = 0
        For Each midspan In wire.midspans
            neutCount = neutCount + 1
            neutSizes(wire.size) = neutSizes(wire.size) + 1
        Next midspan
    Next wire
    For Each wire In pole.secondaries
        If Not secSizes.exists(wire.size) Then secSizes(wire.size) = 0
        For Each midspan In wire.midspans
            secCount = secCount + 1
            secSizes(wire.size) = secSizes(wire.size) + 1
        Next midspan
    Next wire
    For Each wire In pole.openWires
        If Not owSizes.exists(wire.size) Then owSizes(wire.size) = 0
        For Each midspan In wire.midspans
            owCount = owCount + 1
            owSizes(wire.size) = owSizes(wire.size) + 1
        Next midspan
    Next wire
    
    Dim cu As Variant
    For Each cu In cus
        If TypeOf cu Is cu Then
            If cu.location = properLocation(pole.location) And cu.code = "101036" And cu.action = "RET REM" Then
                owCount = owCount - cu.qty
                Exit For
            End If
        End If
    Next cu
    
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
    Dim neededSpoolCUs As Scripting.Dictionary: Set neededSpoolCUs = New Scripting.Dictionary
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
    Dim uniqueTopSideTieSizes As Scripting.Dictionary: Set uniqueTopSideTieSizes = New Scripting.Dictionary
    
    For Each priSize In priSizes
        If Not uniqueTopSideTieSizes.exists(Utilities.OnlyNumbers(CStr(priSize))) Then uniqueTopSideTieSizes.Add Utilities.OnlyNumbers(CStr(priSize)), Nothing
    Next priSize
    For Each neutSize In neutSizes
        If Not uniqueTopSideTieSizes.exists(Utilities.OnlyNumbers(CStr(neutSize))) Then uniqueTopSideTieSizes.Add Utilities.OnlyNumbers(CStr(neutSize)), Nothing
    Next neutSize
    If owCount > 0 Then
        For Each owSize In owSizes
            If Not uniqueTopSideTieSizes.exists(Utilities.OnlyNumbers(CStr(owSize))) Then uniqueTopSideTieSizes.Add Utilities.OnlyNumbers(CStr(owSize)), Nothing
        Next owSize
    End If
    
    If uniqueTopSideTieSizes.count = 0 Then
        For Each wire In pole.primaries
            If Not uniqueTopSideTieSizes.exists(Utilities.OnlyNumbers(CStr(wire.size))) Then uniqueTopSideTieSizes.Add Utilities.OnlyNumbers(CStr(wire.size)), Nothing
        Next wire
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
            
            If InStr(hardware, "PTP") > 0 Or InStr(hardware, "SPIN") > 0 Then
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
            missedlines.Add action & IIf(amount <> 1, " (" & amount & ")", " ") & hardware & " PRIMARY GRIP CU MISSING"
        ElseIf InStr(hardware, "NEUT") > 0 And InStr(hardware, "DE") > 0 Then
            missedlines.Add action & IIf(amount <> 1, " (" & amount & ")", " ") & hardware & " NEUTRAL GRIP CU MISSING"
        ElseIf InStr(hardware, "SEC") > 0 And InStr(hardware, "DE") > 0 Then
            missedlines.Add action & IIf(amount <> 1, " (" & amount & ")", " ") & hardware & " SECONDARY DE CU MISSING"
        ElseIf InStr(hardware, "WR") > 0 Or InStr(hardware, "1VPO") > 0 Or InStr(hardware, "2VPO") > 0 Or InStr(hardware, "3VPO") > 0 Then
            missedlines.Add action & IIf(amount <> 1, " (" & amount & ")", " ") & hardware & " SPOOL TIE CU MISSING"
        ElseIf InStr(hardware, "PTP") > 0 Or InStr(hardware, "SPIN") > 0 Then
            missedlines.Add action & IIf(amount <> 1, " (" & amount & ")", " ") & hardware & " TOP/SIDE TIE CU MISSING"
        ElseIf InStr(hardware, "SCORS") > 0 Then
            missedlines.Add action & IIf(amount <> 1, " (" & amount & ")", " ") & hardware & " SIDE TIE CU MISSING"
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

Private Sub getSpoolTies(cus As Collection, pole As pole, needAdditionalCUs As Collection, neutSizes As Scripting.Dictionary, neutCount As Integer, secSizes As Scripting.Dictionary, secCount As Integer, owSizes As Scripting.Dictionary, owCount As Integer, amount As Integer, action As String)
    Dim uniqueSizes As Scripting.Dictionary: Set uniqueSizes = New Scripting.Dictionary
    Dim totalWires As Integer
    Dim failed As Boolean
    
    totalWires = neutCount + secCount + owCount
    
    For Each size In neutSizes
        sizeNumber = Utilities.OnlyNumbers(CStr(size))
        If Not uniqueSizes.exists(sizeNumber) Then
            uniqueSizes(sizeNumber) = 0
        End If
        uniqueSizes(sizeNumber) = uniqueSizes(sizeNumber) + 1
    Next size
    
    For Each size In secSizes
        sizeNumber = Utilities.OnlyNumbers(CStr(size))
        If Not uniqueSizes.exists(sizeNumber) Then
            uniqueSizes(sizeNumber) = 0
        End If
        uniqueSizes(sizeNumber) = uniqueSizes(sizeNumber) + 1
    Next size
    
    For Each size In owSizes
        sizeNumber = Utilities.OnlyNumbers(CStr(size))
        If Not uniqueSizes.exists(sizeNumber) Then
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
                ElseIf neutCount = 0 And secCount = 0 And owCount = amount * 2 Then
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
    ElseIf totalWires = amount * 2 Then
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
        If Not failed Then
            For i = needAdditionalCUs.count To 1 Step -1
                neededCU = needAdditionalCUs(i)
                hardware = Replace(neededCU(0), " ", "")
                If action = neededCU(2) And (InStr(hardware, "WR") > 0 Or InStr(hardware, "1VPO") > 0 Or InStr(hardware, "2VPO") > 0 Or InStr(hardware, "3VPO") > 0) Then Call needAdditionalCUs.Remove(i)
            Next i
        End If
        If action = "RET REM" And Not failed Then
            Call neutSizes.RemoveAll
            Call secSizes.RemoveAll
            Call owSizes.RemoveAll
            neutCount = 0
            secCount = 0
            owCount = 0
        End If
    End If
End Sub

Private Sub getExtraDECU(cus As Collection, pole As pole, needAdditionalCUs As Collection, index As Integer, sizes As Scripting.Dictionary, sizeCount As Integer, neededCU() As Variant, Optional componentType As String)
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

Private Sub generateCSV(project As project, cus As Collection, Optional demo As Boolean)
    Dim cu As Variant
    Dim filePath As String
    
    If demo Then
        filePath = ThisWorkbook.path & "\" & project.Notification & " - " & "CU - Demo.csv"
    Else
        filePath = ThisWorkbook.path & "\" & project.Notification & " - " & "CU.csv"
    End If
    If InStr(filePath, "sharepoint") > 0 Then filePath = Environ("USERPROFILE") & "\Downloads\" & project.Notification & " - " & "cus.csv"
    
    Call CheckAndCloseWorkbook(filePath)
    
    FileNumber = FreeFile
    Open filePath For Output As #FileNumber

    Print #FileNumber, "Location, CU, QTY, ACTION, CMPLX, PARENT CU, PARENT INSTANCE, CHILD CU, CHILD QTY"
    
    For Each cu In cus
        If TypeOf cu Is cu Then
            Print #FileNumber, cu.location & "," & cu.code & "," & IIf(cu.childQty > 0, " ", cu.qty) & "," & cu.action & ", ," & cu.parentCU & "," & IIf(cu.parentInstance > 0, cu.parentInstance, " ") & "," & cu.childCode & "," & IIf(cu.childQty > 0, cu.childQty, " ")
        Else
            Print #FileNumber, cu(0) & ", , , ," & cu(1) & ", , , , "
        End If
    Next cu

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
            If Trim(cell.OFFSET(0, 3).Value) <> "" Then
                cell.OFFSET(0, 8).Value = 0.1
            Else
                Set foundCell = cuSortWs.UsedRange.Find(what:=cell.Value, LookIn:=xlValues, lookat:=xlWhole)
                If Not foundCell Is Nothing Then cell.OFFSET(0, 8).Value = foundCell.OFFSET(0, 1)
            End If
        End If
        
        cell.OFFSET(0, 9).Value = Utilities.OnlyNumbers(cell.OFFSET(0, -1))
    Next cell
    
    With csvWs.Sort
        .SortFields.Clear
        .SortFields.Add key:=csvWs.Range("K1"), Order:=xlAscending
        .SortFields.Add key:=csvWs.Range("J1"), Order:=xlAscending
        .SortFields.Add key:=csvWs.Range("D1"), Order:=xlDescending
        '.SortFields.Add key:=csvWs.Range("C1"), Order:=xlDescending
        .SetRange csvWs.UsedRange
        .header = xlYes
        .Apply
    End With
    
    csvWs.Columns(11).Delete
    
    csvWb.save
    csvWb.Close

    Application.ScreenUpdating = True
End Sub

Private Sub generateMissedLinesTXT(missedlines As Collection)
    Dim issues As String
    Dim project As project: Set project = New project
    Call project.extractFromSheets
    
    issues = "Lines unable to turn into CUS." & vbLf & Utilities.JoinCollection(missedlines, vbLf)
    filePath = ThisWorkbook.path & "\" & project.Notification & " - " & "MissedLineCUs.txt"
    If InStr(filePath, "sharepoint") > 0 Then filePath = Environ("USERPROFILE") & "\" & project.Notification & " - " & "MissedLineCUs.txt"
    
    fNum = FreeFile
    Open filePath For Output As #fNum
    Print #fNum, issues
    Close #fNum
    
    Shell "notepad.exe """ & filePath & """", vbNormalFocus
End Sub

Private Sub parseLineToCUs(project As project, needAdditionalCUs As Collection, missedlines As Collection, cus As Collection, demoCus As Collection, pole As pole, ByVal line As String, mode As String)
    Dim regex As Object: Set regex = CreateObject("VBScript.RegExp")
    Dim regex2 As Object: Set regex2 = CreateObject("VBScript.RegExp")
    Dim amount As Integer: amount = 1
    Dim cu As cu, otherCu As cu
    Dim cuCode As String
    
    addedCU = False
    vpoPole = False
    
    regex.Pattern = "\((\d+)\)(.+)"
    regex.Global = True
    regex.IgnoreCase = True
    
    Call applyStandardAbbreviations(line)
    line = Replace(line, "OPEN WIRE", "OW")
    line = Replace(line, "OPENWIRE", "OW")
    
    Dim hardware As String: hardware = Trim(line)
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
            If timeAdder = 1 Then timeAdder = 2
            If pole.buck Then timeAdder = 3
            For Each equipment In pole.equipments
                If equipment.componentType = "XFMR" Then timeAdder = 3
            Next equipment
        ElseIf InStr(line, "SVC RISER") > 0 And InStr(line, "|C") > 0 Then
            line1 = Left(line, InStr(line, "SVC RISER") - 1) & "RISER"
            line2 = Left(line, InStr(line, "SVC RISER") - 1) & "RISER"
        ElseIf InStr(line, "SEC RISER") > 0 And InStr(line, "|C") > 0 Then
            line1 = Left(line, InStr(line, "SEC RISER") - 1) & "RISER"
            line2 = Left(line, InStr(line, "SEC RISER") - 1) & "RISER"
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
        
        If InStr(hardware, "OWSERVICEDE") = 1 Then serviceAmount = serviceAmount + amount
        If InStr(hardware, "SVCDE") = 1 Then serviceAmount = serviceAmount + amount
        
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
    
        Dim priRegex As Object
        Set priRegex = CreateObject("VBScript.RegExp")
        
        priRegex.Pattern = "\s*(\d*)'[ OF]*\s*(\d)PH\s*(.*)PRI\s*\/\s*[\d' ]*[ OF]*(\d)PH(.*)PRI\s*(.*)"
        priRegex.Global = True
        priRegex.IgnoreCase = True
    
        Dim neutRegex As Object
        Set neutRegex = CreateObject("VBScript.RegExp")
        
        neutRegex.Pattern = "\s*(\d*)'[ OF]*\s*(.*)NEUT\s*\/\s*[\d' ]*[OF]*(.*)NEUT\s*(.*)"
        neutRegex.Global = True
        neutRegex.IgnoreCase = True
    
        Dim distance As Integer
        Dim phase1 As Integer
        Dim size1 As String
        Dim phase2 As Integer
        Dim size2 As String
        If InStr(line1, "'") > 0 And (InStr(line1, "OW") > 0 Or InStr(line1, "SEC") > 0) And InStr(line2, "SEC") > 0 Then
            Call generateReconductorCUs(cus, pole, line1, line2)
            line1 = ""
            line2 = ""
        ElseIf priRegex.test(line) Then
            Set matches = priRegex.Execute(line)
            distance = CInt(matches(0).SubMatches(0))
            phase1 = CInt(matches(0).SubMatches(1))
            size1 = matches(0).SubMatches(2)
            phase2 = CInt(matches(0).SubMatches(3))
            size2 = matches(0).SubMatches(4)
            
            Call generatePrimaryReconductorCUs(cus, missedlines, pole, distance, phase1, size1, phase2, size2)
        ElseIf neutRegex.test(line) Then
            Set matches = neutRegex.Execute(line)
            distance = CInt(matches(0).SubMatches(0))
            size1 = matches(0).SubMatches(1)
            size2 = matches(0).SubMatches(2)
            
            Call generateNeutralReconductorCUs(cus, pole, distance, size1, size2)
        ElseIf InStr(line, "CO") > 0 And (InStr(line, " ON SA") > 0 Or InStr(line, " ON LCOM") > 0 Or InStr(line, " TO SA") > 0 Or InStr(line, " TO LCOM") > 0) Then
            Call generateTransferCOCU(cus, pole.location, amount, hardware)
            line1 = ""
            line2 = ""
        ElseIf InStr(line, "|LA ") > 0 Or InStr(hardware, "LA ") = 1 And (InStr(line, "XFMR") = 0 Or InStr(line, "TO XFMR") > 0) Then
            Call generateCU(cus, pole.location, "200155", amount, "INSTALL")
            line1 = ""
            line2 = ""
        Else
            If line1 <> "" Then Call parseLineToCUs(project, needAdditionalCUs, missedlines, cus, demoCus, pole, line1, "Remove")
            If line2 <> "" Then Call parseLineToCUs(project, needAdditionalCUs, missedlines, cus, demoCus, pole, line2, "Install")
        End If
        
    Else
        line = Replace(line, ",", "+")
        line = Replace(line, "&", "+")
        If InStr(line, "+") > 0 Then
            parts = Split(line, "+")
            For i = 0 To UBound(parts)
                Call parseLineToCUs(project, needAdditionalCUs, missedlines, cus, demoCus, pole, parts(i), mode)
            Next i
            Exit Sub
        End If
    
        If InStr(hardware, " TO ALLOW") > 0 Then hardware = Left(hardware, InStr(hardware, " TO ALLOW") - 1)
        If InStr(hardware, " TO CORRECT") > 0 Then hardware = Left(hardware, InStr(hardware, " TO CORRECT") - 1)
        If InStr(hardware, " TO PREVENT") > 0 Then hardware = Left(hardware, InStr(hardware, " TO PREVENT") - 1)
        If InStr(hardware, " TO UPGRADE") > 0 Then hardware = Left(hardware, InStr(hardware, " TO UPGRADE") - 1)
        If InStr(hardware, " DUE TO") > 0 Then hardware = Left(hardware, InStr(hardware, " DUE TO") - 1)
        If InStr(hardware, "@") > 0 And InStr(hardware, "LIGHT") = 0 Then hardware = Left(hardware, InStr(hardware, "@") - 1)
        If InStr(hardware, " @") > 0 And InStr(hardware, "LIGHT") = 0 Then hardware = Left(hardware, InStr(hardware, " @") - 1)
        
        'Riser Install/Remove
        If mode <> "Transfer" Then
            If InStr(line, "PRIRISER") > 0 And InStr(line, "|C") > 0 Then
                Call generatePrimaryRiserCU(cus, pole, hardware, properAction(mode))
            ElseIf InStr(line, "RISER") > 0 And InStr(line, "|C") > 0 And InStr(line, "PRI") = 0 Then
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
             
            If (InStr(hardware, "STLT") > 0 Or (InStr(hardware, "FLOOD") > 0 And InStr(hardware, "LIGHT") > 0)) And InStr(hardware, "MOLDING") > 0 Then
                If pole.ReplacePole Then
                    streetlightMolding = streetlightMolding & mode
                    addedCU = True
                Else
                    Call generateReplaceStreetlightMoldingCU(cus, pole, hardware, properAction(mode), missedlines)
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
            If InStr(line, "BOND STLT") > 0 Then Call generateCU(cus, pole.location, "100144", amount, "INSTALL")
        End If
        
        'Remove section handler
        If mode = "Remove" Then
            If InStr(line, "FIRE") > 0 And InStr(line, "WIRE") > 0 Then Call generateCU(cus, pole.location, "201389", amount, "INSTALL")
        End If
        
        'Transfer section handler
        If mode = "Transfer" Then
            If InStr(line, "XFMR") > 0 And InStr(UCase(line), "KVA") > 0 Then Call generateTransferTransformerCU(cus, pole.location, amount)
            If InStr(line, "STLT") > 0 Or (InStr(hardware, "FLOOD") > 0 And InStr(line, "LIGHT") > 0) And InStr(line, "@") > 0 Then Call generateTransferStreetlightCU(line, cus, pole, missedlines)
            If InStr(line, "TRIM") > 0 And InStr(line, "DRIP") Then Call generateCU(cus, pole.location, "101023", 1, "INSTALL")
            If InStr(line, "CO") > 0 And (InStr(line, " ON SA") > 0 Or InStr(line, " ON LCOM") > 0 Or InStr(line, " TO SA") > 0 Or InStr(line, " TO LCOM") > 0) Then Call generateTransferCOCU(cus, pole.location, amount, hardware)
            If InStr(line, "|LA ") > 0 Or InStr(hardware, "LA ") = 1 And InStr(line, "XFMR") = 0 Then Call generateCU(cus, pole.location, "200155", amount, "INSTALL")
            If Replace(hardware, " ", "") = "SVC" Or Replace(hardware, " ", "") = "OHSVC" Then Call generateTransferServiceCU(cus, pole, amount)
        End If
        
        'Note section handler
        If mode = "Note" Then
            If InStr(line, "TOP") > 0 And InStr(line, "POLE") > 0 And InStr(line, "ABOVE") > 0 Then
                Call generateCU(cus, pole.location, "100910", 1, "INSTALL")
                If project.mode = "SYSTEM IMPROVEMENT" Then
                
                    Set polesDict = New Scripting.Dictionary
                    Set cuSortWs = ThisWorkbook.sheets("CUSortOrder")
                    
                    lastRow = cuSortWs.Cells(cuSortWs.Rows.count, "A").End(xlUp).row
                    
                    For i = 1 To lastRow
                        If cuSortWs.Cells(i, "B").Value = "5" Then
                            polesDict(CStr(cuSortWs.Cells(i, "A").Value)) = True
                        End If
                    Next i
                
                    For i = cus.count To 1 Step -1
                        If TypeOf cus(i) Is cu Then
                            Set cu = cus(i)
                            
                            If polesDict.exists(cu.code) And cu.action = "RET REM" And cu.location = properLocation(pole.location) Then
                                Call generateCU(demoCus, pole.location, "100417", 1, "INSTALL")
                                demoCus.Add cus(i)
                                Call generateCU(demoCus, pole.location, "100066", 1, "INSTALL")
                                cus.Remove (i)
                                Exit For
                            End If
                        End If
                    Next i
                End If
            End If
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
            If InStr(hardware, "PRI DE") > 0 And Utilities.OnlyNumbers(hardware) <> "-1" Then
                Call generateCU(cus, pole.location, 290014, amount, properAction(mode))
                cuCode = CUNameMapping.getCUNameMapping(hardware)
                If cuCode = "" Then
                    missedlines.Add properAction(mode) & hardware & " GRIP CU MISSING"
                Else
                    Call generateCU(cus, pole.location, cuCode, amount, properAction(mode))
                End If
            ElseIf InStr(hardware, "NEUT DE") > 0 And Utilities.OnlyNumbers(hardware) <> "-1" Then
                Call generateCU(cus, pole.location, 290034, amount, properAction(mode))
                cuCode = CUNameMapping.getCUNameMapping(hardware)
                If cuCode = "" Then
                    missedlines.Add properAction(mode) & hardware & " GRIP CU MISSING"
                Else
                    Call generateCU(cus, pole.location, cuCode, amount, properAction(mode))
                End If
            End If
            If Not addedCU Then
                cuCode = CUNameMapping.getCUNameMapping(hardware)
                If cuCode = "" Then cuCode = CUNameMapping.getCUNameMapping(Utilities.OnlyLetters(hardware))
                If cuCode = "" And InStr(hardware, "SWAMP FIXTURE") > 0 Then cuCode = "100085"
                If cuCode = "200495" Then Call generateCU(cus, pole.location, "200961", amount, properAction(mode))
                If cuCode <> "" Then
                    If InStr(hardware, "JUMPER") > 0 And (InStr(hardware, "SPIN") > 0 Or InStr(hardware, "PTP") > 0) Then Call generateCU(cus, pole.location, "100724", amount, properAction(mode))
                    Call generateCU(cus, pole.location, cuCode, amount, properAction(mode))
                End If
                
                If CUNameMapping.CheckForAdditionalCUs(hardware) Then needAdditionalCUs.Add Array(hardware, amount, properAction(mode))
            End If
        End If
    End If
    
    'Add missed lines
    If Not addedCU And mode <> "Note" Then
        If (mode = "Replace" And (line1 = "" Or line2 = "")) Or mode <> "Replace" Then
            If Not MissedLineIgnorable(pole, hardware) Then missedlines.Add mode & " " & hardware
        End If
    End If
End Sub

Private Sub generateReplaceStreetlightMoldingCU(cus As Collection, pole As pole, hardware As String, mode As String, missedlines As Collection)
    Dim distance As Integer
    Dim cuCode As String
    Dim closestDistance As Integer

    If pole.slBottomBracketHeight > 1 Then streetlightBottomBracketHeight = pole.slBottomBracketHeight

    If InStr(hardware, "'") > 0 Then distance = Utilities.OnlyNumbers(Left(hardware, InStr(hardware, "'") + 1))
    
    If distance > 0 Then
        Call generateCU(cus, pole.location, "100598", Application.WorksheetFunction.RoundUp(distance / 8, 0), mode)
    Else
        closestDistance = 0
        For Each wire In pole.utilWires
            If wire.componentType = "OW" Or wire.componentType = "SEC" Or wire.componentType = "TRAFFIC" Then
                If Abs(wire.height - streetlightBottomBracketHeight) < closestDistance Or closestDistance = 0 Then closestDistance = Abs(wire.height - streetlightBottomBracketHeight)
            End If
        Next wire
        Call generateCU(cus, pole.location, "100598", WorksheetFunction.RoundUp((closestDistance / 12) / 8, 0), mode)
    End If
End Sub

Private Sub generateTransferServiceCU(cus As Collection, pole As pole, amount As Integer)
    Dim cuCode As String
    Dim totalServices As Integer
    Dim serviceDict As Scripting.Dictionary: Set serviceDict = New Scripting.Dictionary
    
    For Each Service In pole.services
        For Each midspan In Service.midspans
            If Not serviceDict.exists(midspan) Then serviceDict.Add midspan, Nothing
            totalServices = totalServices + 1
        Next midspan
    Next Service
    
    Call generateCU(cus, pole.location, 106115, serviceDict.count, "INSTALL")
    
    If serviceAmount = totalServices Then
        For Each Service In pole.services
            cuCode = CUNameMapping.getCUNameMapping(Service.size & "DE")
            If cuCode <> "" Then
                For Each midspan In Service.midspans
                    Call generateCU(cus, pole.location, cuCode, 1, "RET REM")
                    Call generateCU(cus, pole.location, cuCode, 1, "INSTALL")
                Next midspan
            End If
        Next Service
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
    
    If InStr(line1, "OW") > 0 Then
        owSizes = Mid(line1, InStr(line1, "'") + 1, InStr(line1, "OW") - InStr(line1, "'") - 1)
        
        parts = Split(owSizes, "-")
        
        Set cuCodeDistanceMap = New Scripting.Dictionary
        For i = 0 To UBound(parts)
            owSize = Left(Trim(parts(i)), 1)
            cuCode = CUNameMapping.getOWNameMapping(owSize)
            If cuCode <> "" Then
                If Not cuCodeDistanceMap.exists(cuCode) Then cuCodeDistanceMap(cuCode) = 0
                Call generateCU(cus, pole.location, "290048", 1, "INSTALL")
                cuCodeDistanceMap(cuCode) = cuCodeDistanceMap(cuCode) + distance
                reconductored = True
            End If
        Next i
        For Each cuCodeKey In cuCodeDistanceMap
            Call generateCU(cus, pole.location, CStr(cuCodeKey), CInt(cuCodeDistanceMap(cuCodeKey)), "RET REM", True)
        Next cuCodeKey
        
    ElseIf InStr(line1, "SEC") > 0 Then
        secSize = Left(line1, InStr(line1, "SEC") - 1)
        If InStr(secSize, "'") > 0 Then secSize = Mid(secSize, InStr(secSize, "'") + 1, Len(secSize) - InStr(secSize, "'") - 1)
        cuCode = CUNameMapping.getSecNameMapping(secSize)
        If cuCode <> "" Then
            Call generateCU(cus, pole.location, "290048", 1, "INSTALL")
            Call generateCU(cus, pole.location, cuCode, distance, "RET REM", True)
            reconductored = True
        End If
    End If
    
    secSize = Left(line2, InStr(line2, "SEC") - 1)
    If InStr(secSize, "'") > 0 Then secSize = Mid(secSize, InStr(secSize, "'") + 1, Len(secSize) - InStr(secSize, "'") - 1)
    cuCode = CUNameMapping.getSecNameMapping(secSize)
    If cuCode <> "" Then
        Call generateCU(cus, pole.location, "290048", 1, "INSTALL")
        Call generateCU(cus, pole.location, cuCode, distance, "INSTALL", True)
        reconductored = True
    Else
        addedCU = False
    End If
    
    If reconductored = True And (pole.secondaries.count > 0 Or pole.services.count > 0) Then
        Call generateCU(cus, pole.location, "290061", 1, "INSTALL")
    End If
End Sub

Private Sub generatePrimaryReconductorCUs(cus As Collection, missedlines As Collection, pole As pole, distance As Integer, phase1 As Integer, size1 As String, phase2 As Integer, size2 As String)
    Dim finished1 As Boolean
    Dim cuCode As String
    cuCode = CUNameMapping.getPriNameMapping(phase1 & size1)
    If cuCode = "" Then cuCode = CUNameMapping.getPriNameMapping(phase1 & Utilities.OnlyNumbers(size1))
    If cuCode <> "" Then
       Call generateCU(cus, pole.location, cuCode, distance, "RET REM", True)
       finished1 = True
    End If
    
    Dim finished2 As Boolean
    cuCode = CUNameMapping.getPriNameMapping(phase2 & size2)
    If cuCode = "" Then cuCode = CUNameMapping.getPriNameMapping(phase2 & Utilities.OnlyNumbers(size2))
    If cuCode <> "" Then
       Call generateCU(cus, pole.location, cuCode, distance, "INSTALL", True)
       finished2 = True
       If Not finished1 Then missedlines.Add "MISSING REMOVE PRIMARY SIZE"
    ElseIf finished1 Then
        missedlines.Add "MISSING INSTALL PRIMARY SIZE"
    End If
    
    If finished1 And finished2 Then
        Dim equipment As equipment
        Dim withEquipment As Boolean
        For Each equipment In pole.equipments
            If equipment.componentType = "XFMR" Or equipment.componentType = "CAPACITOR" Or equipment.componentType = "RECLOSER" Or equipment.componentType = "REGULATOR" Then withEquipment = True
        Next equipment
        
        If phase1 = phase2 And phase1 = 3 Then
            If withEquipment Then
                Call generateCU(cus, pole.location, 290043, 1, "INSTALL", True)
            Else
                Call generateCU(cus, pole.location, 290046, 1, "INSTALL", True)
            End If
            primaryReconductored = True
        ElseIf phase1 = phase2 And phase1 = 2 Then
            If withEquipment Then
                Call generateCU(cus, pole.location, 290042, 1, "INSTALL", True)
            Else
                Call generateCU(cus, pole.location, 290045, 1, "INSTALL", True)
            End If
            primaryReconductored = True
        ElseIf phase1 = phase2 And phase1 = 1 Then
            If withEquipment Then
                Call generateCU(cus, pole.location, 290041, 1, "INSTALL", True)
            Else
                Call generateCU(cus, pole.location, 290044, 1, "INSTALL", True)
            End If
            primaryReconductored = True
        ElseIf phase1 = 1 And phase2 = 2 And size1 = size2 Then
            Call generateCU(cus, pole.location, 290053, 1, "INSTALL", True)
            If withEquipment Then Call generateCU(cus, pole.location, 290054, 1, "INSTALL", True)
            primaryReconductored = True
            missedlines.Add "POTENITALLY IRRELEVANT PRIMARY RECONDUCTOR CUS ADDED, REMOVE IF NOT APPLICABLE"
        ElseIf phase1 = 2 And phase2 = 3 And size1 = size2 Then
            If withEquipment Then
                Call generateCU(cus, pole.location, 290056, 1, "INSTALL", True)
            Else
                Call generateCU(cus, pole.location, 290057, 1, "INSTALL", True)
            End If
            primaryReconductored = True
        ElseIf phase1 = 1 And phase2 = 3 And size1 = size2 Then
            If withEquipment Then
                Call generateCU(cus, pole.location, 290049, 1, "INSTALL", True)
            Else
                Call generateCU(cus, pole.location, 290050, 1, "INSTALL", True)
            End If
            primaryReconductored = True
        End If
    ElseIf finished1 Or finished2 Then
        missedlines.Add "MISSING PRIMARY RECONDUCTOR CU"
    End If
End Sub

Private Sub generateNeutralReconductorCUs(cus As Collection, pole As pole, distance As Integer, size1 As String, size2 As String)
    Dim cuCode As String
    cuCode = CUNameMapping.getPriNameMapping(1 & size1)
    If cuCode = "" Then cuCode = CUNameMapping.getPriNameMapping(1 & Utilities.OnlyNumbers(size1))
    If cuCode <> "" Then
       Call generateCU(cus, pole.location, cuCode, distance, "RET REM", True)
    End If
    
    cuCode = CUNameMapping.getPriNameMapping(1 & size2)
    If cuCode = "" Then cuCode = CUNameMapping.getPriNameMapping(1 & Utilities.OnlyNumbers(size2))
    If cuCode <> "" Then
       Call generateCU(cus, pole.location, cuCode, distance, "INSTALL", True)
    End If
End Sub

Private Sub checkForAdjacentPoleRecondcutoring(cus As Collection, project As project, pole As pole, missedlines As Collection)
    Dim Span As Span
    Dim otherPole As pole
    Dim otherSpan As Span
    Dim count As Integer
    Dim lines() As String
    For Each Span In pole.spans
        If Span.otherPole <> "" Then
            Set otherPole = project.findPole(Span.otherPole)
            If InStr(otherPole.alt1, "'") > 0 Then
                If InStr(otherPole.alt1, vbLf) > 0 Then
                    lines = Split(otherPole.alt1, vbLf)
                    For Each line In lines
                        If InStr(line, "'") > 0 And (InStr(line, "OPEN WIRE") > 0 Or InStr(line, "SECONDARY") > 0 Or InStr(line, "SEC") > 0) And (InStr(line, "SECONDARY") > 0 Or InStr(line, "SEC") > 0) Then
                            distance = Utilities.OnlyNumbers(Left(line, InStr(line, "'")))
                            If IsNumeric(distance) Then
                                If CInt(distance) = Span.distance Then
                                    For Each otherSpan In otherPole.spans
                                        If Span.distance = otherSpan.distance And otherSpan.otherPole <> "" Then count = count + 1
                                    Next otherSpan
                                    If count = 1 Then
                                        Call generateCU(cus, pole.location, "290048", 1, "INSTALL")
                                    Else
                                        missedlines.Add "AMBIGUOUS SPAN LENGTHS ON OTHER POLE FOR RECONDUCTORING"
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
    If InStr(line, "TOREPLACEOW") = 1 Then MissedLineIgnorable = True: Exit Function
    If InStr(line, "D=") = 1 Then MissedLineIgnorable = True: Exit Function
    If InStr(line, "P=") = 1 Then MissedLineIgnorable = True: Exit Function
    If InStr(line, "TOCORRECT") = 1 Then MissedLineIgnorable = True: Exit Function
    If InStr(line, "TOMAKE") = 1 Then MissedLineIgnorable = True: Exit Function
    If InStr(line, "TOALLOW") = 1 Then MissedLineIgnorable = True: Exit Function
    If InStr(line, "@11""FROM") = 1 Then MissedLineIgnorable = True: Exit Function
    If InStr(line, "SECDE") = 1 Then MissedLineIgnorable = True: Exit Function
    If InStr(line, "AS-IS") > 0 Then MissedLineIgnorable = True: Exit Function
    If InStr(line, "ASIS") > 0 Then MissedLineIgnorable = True: Exit Function
    If Replace(line, vbLf, "") = "" Then MissedLineIgnorable = True: Exit Function
    If InStr(line, "TO") = 1 Then MissedLineIgnorable = True: Exit Function
    If line = "PRIMARY" Then MissedLineIgnorable = True: Exit Function
    If line = "PRI" Then MissedLineIgnorable = True: Exit Function
    If line = "NEUTRAL" Then MissedLineIgnorable = True: Exit Function
    If line = "NEUT" Then MissedLineIgnorable = True: Exit Function
    If line = "SECONDARY" Then MissedLineIgnorable = True: Exit Function
    If line = "SEC" Then MissedLineIgnorable = True: Exit Function
    If line = "EXTENDTONEWHEIGHT" Then MissedLineIgnorable = True: Exit Function
    If line = "OPENWIRE" Then MissedLineIgnorable = True: Exit Function
    If line = "OW" Then MissedLineIgnorable = True: Exit Function
    If line = "LAONXFMR" Then MissedLineIgnorable = True: Exit Function
    Dim comp As Variant
    For Each comp In pole.commComponents
        If InStr(line, Replace(Replace(comp.owner, " ", ""), "&", "")) = 1 Then MissedLineIgnorable = True: Exit Function
    Next comp
End Function

Private Sub generateCU(cus As Collection, location As String, code As String, qty As Integer, action As String, Optional duplicate As Boolean)
    Dim cu As cu: Set cu = New cu
    If InStr(location, "ALT") = 0 And InStr(location, "LOC") = 0 Then
        cu.location = properLocation(location)
    Else
        cu.location = location
    End If
    cu.code = code
    cu.qty = qty
    cu.action = action
    
    While cu.qty > 999
        Dim cu2 As cu: Set cu2 = New cu
        cu2.location = cu.location
        cu2.code = cu.code
        cu2.qty = cu.qty
        cu2.action = cu.action
        cu2.qty = cu.qty - 999
        If cu2.qty > 999 Then cu2.qty = 999
        cu.qty = cu.qty - cu2.qty
        Call AddCu(cus, cu2, duplicate)
    Wend
    
    Call AddCu(cus, cu, duplicate)
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

Private Sub generateTransferStreetlightCU(line As String, cus As Collection, pole As pole, missedlines As Collection)
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

        missedlines.Add "Replace 2/C-10 CU STLT quantity not set"
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
    For Each wire In pole.utilWires
        If wire.componentType = "OW" Or wire.componentType = "SEC" Or wire.componentType = "TRAFFIC" Then
            If Abs(wire.height - streetlightBottomBracketHeight) < closestDistance Or closestDistance = 0 Then closestDistance = Abs(wire.height - streetlightBottomBracketHeight)
        End If
    Next wire
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
    Dim project As project: Set project = New project
    
    If project.mode = "SYSTEM IMPROVEMENT" Then
        properLocation = "LOC " & location
    Else
        properLocation = "L" & Format(location, "000") & " ALT1"
    End If
End Function

Private Function properAction(mode As String) As String
     Select Case mode
        Case "Install", "Transfer", "Note"
            properAction = "INSTALL"
        Case "Remove"
            properAction = "RET REM"
     End Select
End Function

Private Sub AddCu(cus As Collection, cu As cu, Optional duplicate As Boolean)
    Dim alreadyExists As Boolean
    Dim otherCu As cu
    
    If Not duplicate Then
        For i = 1 To cus.count
            If TypeOf cus(i) Is cu Then
                Set otherCu = cus(i)
                If cu.Equals(otherCu) Then
                    If cu.code <> "290048" And cu.code <> "200548" And otherCu.qty + cu.qty > 999 Then
                        alreadyExists = False
                    Else
                        If cu.code <> "290048" And cu.code <> "200548" Then otherCu.qty = otherCu.qty + cu.qty
                        alreadyExists = True
                        Exit For
                    End If
                End If
            End If
        Next i
    End If
    
    addedCU = True
    If Not alreadyExists Then cus.Add cu
End Sub
