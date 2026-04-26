Attribute VB_Name = "NJUNSCodes"
Public NJUNSCodes As Scripting.Dictionary

Public Sub generateNJUNSCodes()
    Call clearNJUNSCodes
    
    Dim companies As Scripting.Dictionary: Set companies = New Scripting.Dictionary
    Dim controlWs As Worksheet: Set controlWs = ThisWorkbook.sheets("Control")
    Dim Project As Project: Set Project = New Project
    Call Project.extractFromSheets
    
    controlWs.Range("NJUNSCODES").offset(0, 0).Value = "Place(Check NJUNS)"
    controlWs.Range("NJUNSCODES").offset(0, 1).Value = Project.township
    
    controlWs.Range("NJUNSCODES").offset(1, 0).Value = "Applicant(" & Project.applicant & ")"
    controlWs.Range("NJUNSCODES").offset(1, 1).Value = getNJUNSNameMapping(Project, Project.applicant)
    
    Dim pole As pole
    For Each pole In Project.poles
        If pole.njunsTicket <> "" Then
            For Each step In pole.njunsSteps
                company = Application.WorksheetFunction.Proper(Utilities.GetFirstWord(CStr(step)))
                company = Replace(company, ":", "")
                If company = "Ce" Then company = "Consumers"
                If company <> "Consumers" Then
                    If Not companies.Exists(company) Then companies.Add company, 0
                    companies(company) = companies(company) + 1
                End If
            Next step
        End If
    Next pole
    
    Dim i As Integer: i = 2
    For Each company In companies
        controlWs.Range("NJUNSCODES").offset(i, 0).Value = company & "(" & companies(company) & ")"
        controlWs.Range("NJUNSCODES").offset(i, 1).Value = getNJUNSNameMapping(Project, company)
        i = i + 1
    Next company
    
    Set NJUNSCodes = Nothing
End Sub

Public Sub clearNJUNSCodes()
    Dim controlWs As Worksheet: Set controlWs = ThisWorkbook.sheets("Control")
    
    controlWs.Unprotect
    
    Dim startCell As Range: Set startCell = controlWs.Range("NJUNSCODES")
    
    controlWs.Range(startCell, startCell.offset(26, 1)).ClearContents
    
    controlWs.Protect _
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
End Sub

Private Function getNJUNSNameMapping(Project, ByVal key As String) As String
    key = UCase(Replace(key, " ", ""))
    key = Replace(key, ":", "")
    key = Replace(key, vbLf, "")
    
    If NJUNSCodes Is Nothing Then
        Set NJUNSCodes = New Scripting.Dictionary
        Call InitializeNJUNSNameCorrecting
    End If
    
    If NJUNSCodes.Exists(key) Then
        getNJUNSNameMapping = NJUNSCodes(key)
    Else
        If NJUNSCodes.Exists(key & UCase(Project.county)) Then
            getNJUNSNameMapping = NJUNSCodes(key & UCase(Project.county))
        Else
            getNJUNSNameMapping = ""
        End If
    End If
End Function

Private Sub InitializeNJUNSNameCorrecting()
    NJUNSCodes("COMCAST") = "COMCMI"
    NJUNSCodes("KEPS") = "KEPTCE"
    NJUNSCodes("METRONET") = "MTRFMI"
    NJUNSCodes("CLIMAX") = "MTRFMI"
    NJUNSCodes("AT&TALCONA") = "ATT101"
    NJUNSCodes("AT&TALLEGAN") = "ATT103"
    NJUNSCodes("AT&TANTRIM") = "ATT105"
    NJUNSCodes("AT&TARENAC") = "ATT106"
    NJUNSCodes("AT&TBARRY") = "ATT108"
    NJUNSCodes("AT&TBAY") = "ATT109"
    NJUNSCodes("AT&TBENZIE") = "ATT110"
    NJUNSCodes("AT&TBERRIEN") = "ATT111"
    NJUNSCodes("AT&TCHEBOYGAN") = "ATT116"
    NJUNSCodes("AT&TCHIPPEWA") = "ATT117"
    NJUNSCodes("AT&TCLARE") = "ATT118"
    NJUNSCodes("AT&TCLINTON") = "ATT119"
    NJUNSCodes("AT&TDELTA") = "ATT121"
    NJUNSCodes("AT&TEATON") = "ATT123"
    NJUNSCodes("AT&TEMMET") = "ATT124"
    NJUNSCodes("AT&TGENESEE") = "ATT125"
    NJUNSCodes("AT&TGLADWIN") = "ATT126"
    NJUNSCodes("AT&THILLSDALE") = "ATT130"
    NJUNSCodes("AT&TIONIA") = "ATT134"
    NJUNSCodes("AT&TJACKSON") = "ATT138"
    NJUNSCodes("AT&TKALAMAZOO") = "ATT139"
    NJUNSCodes("AT&TKENT") = "ATT141"
    NJUNSCodes("AT&TMUSKEGON") = "ATT161"
    NJUNSCodes("AT&TNEWAYGO") = "ATT162"
    NJUNSCodes("AT&TOTTAWA") = "ATT170"
    NJUNSCodes("AT&TSAGINAW") = "ATT173"
    
    NJUNSCodes("CHARTERNEWAYGO") = "CRTROC"
    NJUNSCodes("CHARTERALCONA") = "CRTSAG"
    NJUNSCodes("CHARTERSAGINAW") = "CRTSAG"
    
    NJUNSCodes("WOW") = "MILLDG"
    NJUNSCodes("SFN") = "MASDCE"
    NJUNSCodes("MCI") = "BFBRCE"
    NJUNSCodes("WPS") = "WPSDCE"
    NJUNSCodes("PFN") = "PFNCE"
    NJUNSCodes("LYNX") = "GLCCE"
    NJUNSCodes("MHS") = "MHPSCE"
    NJUNSCodes("SRESD") = "SHSDCE"
    NJUNSCodes("EVERSTREAM") = "GLCCE"
    NJUNSCodes("USSIGNAL") = "RVPFCE"
    NJUNSCodes("US") = "RVPFCE"
    NJUNSCodes("NC") = "NCSDCE"
    NJUNSCodes("CENTURYLINK") = "CNTLCE"
    NJUNSCodes("BRIGHTSPEED") = "CNTLCE"
End Sub

