VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CrewNotesGenerator_Form 
   Caption         =   "Crew Notes Generator"
   ClientHeight    =   6750
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19800
   OleObjectBlob   =   "CrewNotesGenerator_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CrewNotesGenerator_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim pds As Worksheet
Dim pole As pole
Dim crewNotes As String, installNotes As String, removeNotes As String, replaceNotes As String, transferNotes, notes As String
Dim ohsCount As Integer, dg11KCount As Integer, dg20KCount As Integer, spg11KCount As Integer, spg20KCount As Integer, replace11kCount As Integer, replace20kCount As Integer
Dim priBool As Boolean, neutBool As Boolean, secBool As Boolean, owBool As Boolean, riserBool As Boolean
Dim fuRP1 As Boolean, fuRP2 As Boolean, fuRP3 As Boolean, fuIN1 As Boolean, fuIN2 As Boolean, fuIN3 As Boolean, fuRM1 As Boolean, fuRM2 As Boolean, fuRM3 As Boolean
Dim priRP1 As Boolean, priRP2 As Boolean, priRP3 As Boolean, priRP4 As Boolean, priRP5 As Boolean, priRP6 As Boolean, priRP7 As Boolean
Dim priIN1 As Boolean, priIN2 As Boolean, priIN3 As Boolean, priIN4 As Boolean, priIN5 As Boolean, priIN6 As Boolean, priIN7 As Boolean
Dim priRM1 As Boolean, priRM2 As Boolean, priRM3 As Boolean, priRM4 As Boolean, priRM5 As Boolean, priRM6 As Boolean, priRM7 As Boolean
Dim snRP1 As Boolean, snRP2 As Boolean, snRP3 As Boolean, snRP4 As Boolean, snRP5 As Boolean, snRP6 As Boolean, snRP7 As Boolean
Dim snIN1 As Boolean, snIN2 As Boolean, snIN3 As Boolean, snIN4 As Boolean, snIN5 As Boolean, snIN6 As Boolean, snIN7 As Boolean
Dim snRM1 As Boolean, snRM2 As Boolean, snRM3 As Boolean, snRM4 As Boolean, snRM5 As Boolean, snRM6 As Boolean, snRM7 As Boolean
Dim PreviousValueRPFUC1 As Integer, PreviousValueRPFUC2 As Integer, PreviousValueRPFUC3 As Integer, PreviousValueINFUC1 As Integer, PreviousValueINFUC2 As Integer, PreviousValueINFUC3 As Integer, PreviousValueRMFUC1 As Integer, PreviousValueRMFUC2 As Integer, PreviousValueRMFUC3 As Integer
Dim PreviousValueRPPRIC1 As Integer, PreviousValueRPPRIC2 As Integer, PreviousValueRPPRIC3 As Integer, PreviousValueRPPRIC4 As Integer, PreviousValueRPPRIC5 As Integer, PreviousValueRPPRIC6 As Integer, PreviousValueRPPRIC7 As Integer
Dim PreviousValueINPRIC1 As Integer, PreviousValueINPRIC2 As Integer, PreviousValueINPRIC3 As Integer, PreviousValueINPRIC4 As Integer, PreviousValueINPRIC5 As Integer, PreviousValueINPRIC6 As Integer, PreviousValueINPRIC7 As Integer
Dim PreviousValueRMPRIC1 As Integer, PreviousValueRMPRIC2 As Integer, PreviousValueRMPRIC3 As Integer, PreviousValueRMPRIC4 As Integer, PreviousValueRMPRIC5 As Integer, PreviousValueRMPRIC6 As Integer, PreviousValueRMPRIC7 As Integer
Dim PreviousValueRPSNC1 As Integer, PreviousValueRPSNC2 As Integer, PreviousValueRPSNC3 As Integer, PreviousValueRPSNC4 As Integer, PreviousValueRPSNC5 As Integer, PreviousValueRPSNC6 As Integer, PreviousValueRPSNC7 As Integer
Dim PreviousValueINSNC1 As Integer, PreviousValueINSNC2 As Integer, PreviousValueINSNC3 As Integer, PreviousValueINSNC4 As Integer, PreviousValueINSNC5 As Integer, PreviousValueINSNC6 As Integer, PreviousValueINSNC7 As Integer
Dim PreviousValueRMSNC1 As Integer, PreviousValueRMSNC2 As Integer, PreviousValueRMSNC3 As Integer, PreviousValueRMSNC4 As Integer, PreviousValueRMSNC5 As Integer, PreviousValueRMSNC6 As Integer, PreviousValueRMSNC7 As Integer
Dim previousValueRPMHC1 As Integer, previousValueINMHC1 As Integer, previousValueRMMHC1 As Integer, previousValueRPMHC2 As Integer, previousValueINMHC2 As Integer, previousValueRMMHC2 As Integer
Dim PreviousValueRECA1 As Integer, PreviousValueRECA2 As Integer, PreviousValueRECA3 As Integer, PreviousValueRECA4 As Integer
Dim mhRP1 As Boolean, mhIN1 As Boolean, mhRM1 As Boolean, mhRP2 As Boolean, mhIN2 As Boolean, mhRM2 As Boolean
Dim figuresUsed As Scripting.Dictionary
Dim comms As Scripting.Dictionary
Dim commTransfers As Collection

Public Sub Initialize(sheet As Worksheet)
    Dim priC As Integer, snC As Integer
    
    Set pds = sheet
    Set pole = New pole
    Call pole.extractFromSheet(sheet)
    Me.MultiPage1.value = 0
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.width) - (0.5 * Me.width)
    Me.top = Application.top + (0.5 * Application.height) - (0.5 * Me.height)
    
    Set figuresUsed = New Scripting.Dictionary
    
    priBool = Not pds.Range("SECONLY")
    neutBool = False
    secBool = False
    owBool = False
    ohsCount = 0
    dg11KCount = 0
    dg20KCount = 0
    spg11KCount = 0
    spg20KCount = 0
    
    previousValueRPMHC1 = 1
    previousValueINMHC1 = 1
    previousValueRMMHC1 = 1
    mhRP1 = False
    mhIN1 = False
    mhRM1 = False
    
    previousValueRPMHC2 = 1
    previousValueINMHC2 = 1
    previousValueRMMHC2 = 1
    mhRP2 = False
    mhIN2 = False
    mhRM2 = False
    
    fuRP1 = False
    fuRP2 = False
    fuRP3 = False
    
    PreviousValueRPFUC1 = 1
    PreviousValueRPFUC2 = 1
    PreviousValueRPFUC3 = 1
    
    fuIN1 = False
    fuIN2 = False
    fuIN3 = False
    
    PreviousValueINFUC1 = 1
    PreviousValueINFUC2 = 1
    PreviousValueINFUC3 = 1
    
    fuRM1 = False
    fuRM2 = False
    fuRM3 = False
    
    PreviousValueRMFUC1 = 1
    PreviousValueRMFUC2 = 1
    PreviousValueRMFUC3 = 1
    
    PreviousValueRECA1 = 1
    PreviousValueRECA2 = 1
    PreviousValueRECA3 = 1
    PreviousValueRECA4 = 1
    
    priRP1 = False
    priRP2 = False
    priRP3 = False
    priRP4 = False
    priRP5 = False
    priRP6 = False
    priRP7 = False
    
    priIN1 = False
    priIN2 = False
    priIN3 = False
    priIN4 = False
    priIN5 = False
    priIN6 = False
    priIN7 = False
    
    priRM1 = False
    priRM2 = False
    priRM3 = False
    priRM4 = False
    priRM5 = False
    priRM6 = False
    priRM7 = False
    
    snRP1 = False
    snRP2 = False
    snRP3 = False
    snRP4 = False
    snRP5 = False
    snRP6 = False
    snRP7 = False
    
    snIN1 = False
    snIN2 = False
    snIN3 = False
    snIN4 = False
    snIN5 = False
    snIN6 = False
    snIN7 = False
    
    snRM1 = False
    snRM2 = False
    snRM3 = False
    snRM4 = False
    snRM5 = False
    snRM6 = False
    snRM7 = False
    
    PreviousValueRPPRIC1 = 1
    PreviousValueRPPRIC2 = 1
    PreviousValueRPPRIC3 = 1
    PreviousValueRPPRIC4 = 1
    PreviousValueRPPRIC5 = 1
    PreviousValueRPPRIC6 = 1
    PreviousValueRPPRIC7 = 1
    
    PreviousValueINPRIC1 = 1
    PreviousValueINPRIC2 = 1
    PreviousValueINPRIC3 = 1
    PreviousValueINPRIC4 = 1
    PreviousValueINPRIC5 = 1
    PreviousValueINPRIC6 = 1
    PreviousValueINPRIC7 = 1
    
    PreviousValueRMPRIC1 = 1
    PreviousValueRMPRIC2 = 1
    PreviousValueRMPRIC3 = 1
    PreviousValueRMPRIC4 = 1
    PreviousValueRMPRIC5 = 1
    PreviousValueRMPRIC6 = 1
    PreviousValueRMPRIC7 = 1
    
    PreviousValueRPSNC1 = 1
    PreviousValueRPSNC2 = 1
    PreviousValueRPSNC3 = 1
    PreviousValueRPSNC4 = 1
    PreviousValueRPSNC5 = 1
    PreviousValueRPSNC6 = 1
    PreviousValueRPSNC7 = 1
    
    PreviousValueINSNC1 = 1
    PreviousValueINSNC2 = 1
    PreviousValueINSNC3 = 1
    PreviousValueINSNC4 = 1
    PreviousValueINSNC5 = 1
    PreviousValueINSNC6 = 1
    PreviousValueINSNC7 = 1
    
    PreviousValueRMSNC1 = 1
    PreviousValueRMSNC2 = 1
    PreviousValueRMSNC3 = 1
    PreviousValueRMSNC4 = 1
    PreviousValueRMSNC5 = 1
    PreviousValueRMSNC6 = 1
    PreviousValueRMSNC7 = 1
    
    Set comms = New Scripting.Dictionary
    
    Dim name As Variant
    
    For i = 0 To 50
        If pds.Range("CMOWNER").offSet(i, 0).Interior.color = 16777215 Then Exit For
        If pds.Range("CMOWNER").offSet(i, 0) = "" Then Exit For
        If pds.Range("CMOWNER").offSet(i, 0).text = "Clearance Requirement" Then Exit For
        If InStr(pds.Range("CMSIZE").offSet(i, 0).text, """") > 0 Then
            owner = Replace(pds.Range("CMOWNER").offSet(i, 0).text, " SVC", "")
            If Not comms.exists(owner) Then comms.Add owner, Nothing
        End If
    Next i
    
    For i = 0 To 50
        If pds.Range("UTTYPE").offSet(i, 0).Interior.color = 16777215 Then Exit For
        If pds.Range("UTTYPE").offSet(i, 0) = "" Then Exit For
        If pds.Range("UTTYPE").offSet(i, 0).text = "STLT. BOTTOM BRKT." Then Exit For
        If InStr(pds.Range("UTTYPE").offSet(i, 0).text, "NEUT") > 0 Then neutBool = True
        If InStr(pds.Range("UTTYPE").offSet(i, 0).text, "SEC") > 0 Then secBool = True
        If InStr(pds.Range("UTTYPE").offSet(i, 0).text, "OW") > 0 Then owBool = True
        If InStr(pds.Range("UTTYPE").offSet(i, 0).text, "RISER") > 0 Then riserBool = True
        If InStr(pds.Range("UTTYPE").offSet(i, 0).text, "DG") > 0 And InStr(pds.Range("UTSIZE").offSet(i, 0).text, "11K") Then dg11KCount = dg11KCount + 1
        If InStr(pds.Range("UTTYPE").offSet(i, 0).text, "DG") > 0 And InStr(pds.Range("UTSIZE").offSet(i, 0).text, "20K") Then dg20KCount = dg20KCount + 1
        If InStr(pds.Range("UTTYPE").offSet(i, 0).text, "SVC") > 0 Then
            For j = 1 To 12
                For Each name In pds.Names
                    If name.name = "'" & pds.name & "'" & "!" & "TOPOLE" & j Then
                        If Trim(Replace(pds.Range("UTMIDSPAN" & j).offSet(i, 0), "-", "")) <> "" Then
                            ohsCount = ohsCount + 1
                        End If
                    End If
                Next name
            Next j
        End If
        If InStr(pds.Range("UTTYPE").offSet(i, 0).text, "SPG") > 0 Then
            For j = 1 To 12
                For Each name In pds.Names
                    If name.name = "'" & pds.name & "'" & "!" & "TOPOLE" & j Then
                        If Trim(Replace(pds.Range("UTMIDSPAN" & j).offSet(i, 0), "-", "")) <> "" And InStr(pds.Range("UTSIZE").offSet(i, 0).text, "11K") Then spg11KCount = spg11KCount + 1
                        If Trim(Replace(pds.Range("UTMIDSPAN" & j).offSet(i, 0), "-", "")) <> "" And InStr(pds.Range("UTSIZE").offSet(i, 0).text, "20K") Then spg20KCount = spg20KCount + 1
                    End If
                Next name
            Next j
        End If
        If InStr(pds.Range("UTTYPE").offSet(i, 0).text, "OW") > 0 Or InStr(pds.Range("UTTYPE").offSet(i, 0).text, "SEC") > 0 Or InStr(pds.Range("UTTYPE").offSet(i, 0).text, "NEUT") > 0 Then
            For j = 1 To 12
                For Each name In pds.Names
                    If name.name = "'" & pds.name & "'" & "!" & "TOPOLE" & j Then
                        If Trim(Replace(pds.Range("UTMIDSPAN" & j).offSet(i, 0), "-", "")) <> "" Then snC = snC + 1
                    End If
                Next name
            Next j
        End If
    Next i
    
    For i = 0 To 3
        If InStr(pds.Range("DESC").offSet(i, 0).value, "3") > 0 Then priC = priC + (3 * (CountCharInString(pds.Range("DIRECTION").offSet(i, 0), "/") + 1))
        If InStr(pds.Range("DESC").offSet(i, 0).value, "2") > 0 Then priC = priC + (2 * (CountCharInString(pds.Range("DIRECTION").offSet(i, 0), "/") + 1))
        If InStr(pds.Range("DESC").offSet(i, 0).value, "1") > 0 Then priC = priC + (CountCharInString(pds.Range("DIRECTION").offSet(i, 0), "/") + 1)
    Next i
    
    UAP.value = priC
    UASN.value = snC
    
    If RPPSIZE = "" And priC > 0 Then RPPSIZE = 40
    
    Call Initialize_Basic
    Call Initialize_ReplacePole
    Call Initialize_Reconductor
    Call Initialize_Downguys
End Sub

Private Function CountCharInString(ByVal targetString As String, ByVal charToFind As String) As Long
    If Len(charToFind) = 0 Then
        CountCharInString = 0
    Else
        CountCharInString = (Len(targetString) - Len(Replace(targetString, charToFind, ""))) / Len(charToFind)
    End If
End Function

Private Sub BAFIG_Click()
    If PSV.value = "" And BAFIG.value Then PSV.value = "4'0"""
End Sub

Private Sub BSL_Click()
    BSLRFW.visible = BSL.value
    RPBSL.value = BSL.value
    RPBSLRFW.visible = RPP.value And BSL.value
    SLRFWL.visible = RPP.value And BSL.value
    RFWL2.visible = BSL.value
End Sub

Private Sub BSLRFW_Change()
    RPBSLRFW.value = BSLRFW.value
End Sub

Private Sub MHT1_Change()
    If MHT1.value = "RP" Then
        MH1.value = mhRP1
        MHC1.value = previousValueRPMHC1
    ElseIf MHT1.value = "IN" Then
        MH1.value = mhIN1
        MHC1.value = previousValueINMHC1
    ElseIf MHT1.value = "RM" Then
        MH1.value = mhRM1
        MHC1.value = previousValueRMMHC1
    End If
End Sub

Private Sub MH1_Click()
    MHC1.visible = MH1.value
    If MHT1.value = "RP" Then
        If mhRP1 <> MH1.value Then mhRP1 = MH1.value
    ElseIf MHT1.value = "IN" Then
        If mhIN1 <> MH1.value Then mhIN1 = MH1.value
    ElseIf MHT1.value = "RM" Then
        If mhRM1 <> MH1.value Then mhRM1 = MH1.value
    End If
End Sub

Private Sub MHC1_Change()
    If MHT1.value = "RP" Then
        previousValueRPMHC1 = MHC1.value
    ElseIf MHT1.value = "IN" Then
        previousValueINMHC1 = MHC1.value
    ElseIf MHT1.value = "RM" Then
        previousValueRMMHC1 = MHC1.value
    End If
End Sub

Private Sub MHT2_Change()
    If MHT2.value = "RP" Then
        MH2.value = mhRP2
        MHC2.value = previousValueRPMHC2
    ElseIf MHT2.value = "IN" Then
        MH2.value = mhIN2
        MHC2.value = previousValueINMHC2
    ElseIf MHT2.value = "RM" Then
        MH2.value = mhRM2
        MHC2.value = previousValueRMMHC2
    End If
End Sub

Private Sub MH2_Click()
    MHC2.visible = MH2.value
    If MHT2.value = "RP" Then
        If mhRP2 <> MH2.value Then mhRP2 = MH2.value
    ElseIf MHT2.value = "IN" Then
        If mhIN2 <> MH2.value Then mhIN2 = MH2.value
    ElseIf MHT2.value = "RM" Then
        If mhRM2 <> MH2.value Then mhRM2 = MH2.value
    End If
End Sub

Private Sub MHC2_Change()
    If MHT2.value = "RP" Then
        previousValueRPMHC2 = MHC2.value
    ElseIf MHT2.value = "IN" Then
        previousValueINMHC2 = MHC2.value
    ElseIf MHT2.value = "RM" Then
        previousValueRMMHC2 = MHC2.value
    End If
End Sub

Private Sub FUT1_Change()
    If FUT1.value = "RP" Then
        FU1.value = fuRP1
        FUC1.value = PreviousValueRPFUC1
    ElseIf FUT1.value = "IN" Then
        FU1.value = fuIN1
        FUC1.value = PreviousValueINFUC1
    ElseIf FUT1.value = "RM" Then
        FU1.value = fuRM1
        FUC1.value = PreviousValueRMFUC1
    End If
End Sub

Private Sub FU1_Click()
    FUC1.visible = FU1.value
    If FUT1.value = "RP" Then
        If fuRP1 <> FU1.value Then fuRP1 = FU1.value
    ElseIf FUT1.value = "IN" Then
        If fuIN1 <> FU1.value Then fuIN1 = FU1.value
    ElseIf FUT1.value = "RM" Then
        If fuRM1 <> FU1.value Then fuRM1 = FU1.value
    End If
End Sub

Private Sub FUC1_Change()
    If FUT1.value = "RP" Then
        PreviousValueRPFUC1 = FUC1.value
    ElseIf FUT1.value = "IN" Then
        PreviousValueINFUC1 = FUC1.value
    ElseIf FUT1.value = "RM" Then
        PreviousValueRMFUC1 = FUC1.value
    End If
End Sub

Private Sub FUT2_Change()
    If FUT2.value = "RP" Then
        FU2.value = fuRP2
        FUC2.value = PreviousValueRPFUC2
    ElseIf FUT2.value = "IN" Then
        FU2.value = fuIN2
        FUC2.value = PreviousValueINFUC2
    ElseIf FUT2.value = "RM" Then
        FU2.value = fuRM2
        FUC2.value = PreviousValueRMFUC2
    End If
End Sub

Private Sub FU2_Click()
    FUC2.visible = FU2.value
    If FUT2.value = "RP" Then
        If fuRP2 <> FU2.value Then fuRP2 = FU2.value
    ElseIf FUT2.value = "IN" Then
        If fuIN2 <> FU2.value Then fuIN2 = FU2.value
    ElseIf FUT2.value = "RM" Then
        If fuRM2 <> FU2.value Then fuRM2 = FU2.value
    End If
End Sub

Private Sub FUC2_Change()
    If FUT2.value = "RP" Then
        PreviousValueRPFUC2 = FUC2.value
    ElseIf FUT2.value = "IN" Then
        PreviousValueINFUC2 = FUC2.value
    ElseIf FUT2.value = "RM" Then
        PreviousValueRMFUC2 = FUC2.value
    End If
End Sub

Private Sub FUT3_Change()
    If FUT3.value = "RP" Then
        FU3.value = fuRP3
        FUC3.value = PreviousValueRPFUC3
    ElseIf FUT3.value = "IN" Then
        FU3.value = fuIN3
        FUC3.value = PreviousValueINFUC3
    ElseIf FUT3.value = "RM" Then
        FU3.value = fuRM3
        FUC3.value = PreviousValueRMFUC3
    End If
End Sub

Private Sub FU3_Click()
    FUC3.visible = FU3.value
    If FUT3.value = "RP" Then
        If fuRP3 <> FU3.value Then fuRP3 = FU3.value
    ElseIf FUT3.value = "IN" Then
        If fuIN3 <> FU3.value Then fuIN3 = FU3.value
    ElseIf FUT3.value = "RM" Then
        If fuRM3 <> FU3.value Then fuRM3 = FU3.value
    End If
End Sub

Private Sub FUC3_Change()
    If FUT3.value = "RP" Then
        PreviousValueRPFUC3 = FUC3.value
    ElseIf FUT3.value = "IN" Then
        PreviousValueINFUC3 = FUC3.value
    ElseIf FUT3.value = "RM" Then
        PreviousValueRMFUC3 = FUC3.value
    End If
End Sub

Private Sub IA11_Change()
    If IA11.value = "RS" Then
        IA12.list = Array("11K")
        IA12.ListIndex = 0
    ElseIf IA11.value = "RT" Then
        IA12.list = Array("(2)11K", "20K")
        IA12.ListIndex = 0
    ElseIf IA11.value = "STE" Then
        IA12.list = Array("(3)11K", "20K + 11K", "(2)20K")
        IA12.ListIndex = 0
    End If
End Sub

Private Sub IA21_Change()
    If IA21.value = "RS" Then
        IA22.list = Array("11K")
        IA22.ListIndex = 0
    ElseIf IA21.value = "RT" Then
        IA22.list = Array("(2)11K", "20K")
        IA22.ListIndex = 0
    ElseIf IA21.value = "STE" Then
        IA22.list = Array("(3)11K", "20K + 11K", "(2)20K")
        IA22.ListIndex = 0
    End If
End Sub

Private Sub IA31_Change()
    If IA31.value = "RS" Then
        IA32.list = Array("11K")
        IA32.ListIndex = 0
    ElseIf IA31.value = "RT" Then
        IA32.list = Array("(2)11K", "20K")
        IA32.ListIndex = 0
    ElseIf IA31.value = "STE" Then
        IA32.list = Array("(3)11K", "20K + 11K", "(2)20K")
        IA32.ListIndex = 0
    End If
End Sub

Private Sub EA11_Change()
    RA11.value = EA11.value
    If EA11.value = "RS" Then
        EA12.list = Array("11K")
        EA12.ListIndex = 0
    ElseIf EA11.value = "RT" Then
        EA12.list = Array("(2)11K", "20K")
        EA12.ListIndex = 0
    ElseIf EA11.value = "STE" Then
        EA12.list = Array("(3)11K", "20K + 11K", "(2)20K")
        EA12.ListIndex = 0
    End If
End Sub

Private Sub EA21_Change()
    If EA21.value = "RS" Then
        EA22.list = Array("11K")
        EA22.ListIndex = 0
    ElseIf EA21.value = "RT" Then
        EA22.list = Array("(2)11K", "20K")
        EA22.ListIndex = 0
    ElseIf EA21.value = "STE" Then
        EA22.list = Array("(3)11K", "20K + 11K", "(2)20K")
        EA22.ListIndex = 0
    End If
    RA21.value = EA21.value
End Sub

Private Sub EA31_Change()
    If EA31.value = "RS" Then
        EA32.list = Array("11K")
        EA32.ListIndex = 0
    ElseIf EA31.value = "RT" Then
        EA32.list = Array("(2)11K", "20K")
        EA32.ListIndex = 0
    ElseIf EA31.value = "STE" Then
        EA32.list = Array("(3)11K", "20K + 11K", "(2)20K")
        EA32.ListIndex = 0
    End If
    RA31.value = EA31.value
End Sub

Private Sub EA41_Change()
    If EA41.value = "RS" Then
        EA42.list = Array("11K")
        EA42.ListIndex = 0
    ElseIf EA41.value = "RT" Then
        EA42.list = Array("(2)11K", "20K")
        EA42.ListIndex = 0
    ElseIf EA41.value = "STE" Then
        EA42.list = Array("(3)11K", "20K + 11K", "(2)20K")
        EA42.ListIndex = 0
    End If
    RA41.value = EA41.value
End Sub

Private Sub PSV_Change()
    Call determineRecDSpace
End Sub

Private Sub RA11_Change()
    If RA11.value = "RS" Then
        RA12.list = Array("11K")
        RA12.ListIndex = 0
    ElseIf RA11.value = "RT" Then
        RA12.list = Array("(2)11K", "20K")
        RA12.ListIndex = 0
    ElseIf RA11.value = "STE" Then
        RA12.list = Array("(3)11K", "20K + 11K", "(2)20K")
        RA12.ListIndex = 0
    End If
End Sub

Private Sub RA21_Change()
    If RA21.value = "RS" Then
        RA22.list = Array("11K")
        RA22.ListIndex = 0
    ElseIf RA21.value = "RT" Then
        RA22.list = Array("(2)11K", "20K")
        RA22.ListIndex = 0
    ElseIf RA21.value = "STE" Then
        RA22.list = Array("(3)11K", "20K + 11K", "(2)20K")
        RA22.ListIndex = 0
    End If
End Sub

Private Sub RA31_Change()
    If RA31.value = "RS" Then
        RA32.list = Array("11K")
        RA32.ListIndex = 0
    ElseIf RA31.value = "RT" Then
        RA32.list = Array("(2)11K", "20K")
        RA32.ListIndex = 0
    ElseIf RA31.value = "STE" Then
        RA32.list = Array("(3)11K", "20K + 11K", "(2)20K")
        RA32.ListIndex = 0
    End If
End Sub

Private Sub RA41_Change()
    If RA41.value = "RS" Then
        RA42.list = Array("11K")
        RA42.ListIndex = 0
    ElseIf RA41.value = "RT" Then
        RA42.list = Array("(2)11K", "20K")
        RA42.ListIndex = 0
    ElseIf RA41.value = "STE" Then
        RA42.list = Array("(3)11K", "20K + 11K", "(2)20K")
        RA42.ListIndex = 0
    End If
End Sub

Private Sub EA12_Change()
    RA12.value = EA12.value
End Sub

Private Sub EA22_Change()
    RA22.value = EA22.value
End Sub

Private Sub EA32_Change()
    RA32.value = EA32.value
End Sub

Private Sub EA42_Change()
    RA42.value = EA42.value
End Sub

Private Sub IAC_Change()
    For i = 1 To 3
        If CInt(IAC.value) >= i Then
            Me.Controls("IA" & i & "1").visible = True
            Me.Controls("IA" & i & "2").visible = True
            Me.Controls("IA" & i & "3").visible = True
            Me.Controls("IA" & i & "4").visible = True
            Me.Controls("IA" & i & "5").visible = True
            Me.Controls("IGRFW" & i).visible = True
            Me.Controls("IGFIG" & i).visible = True
            IL1.visible = True
            IL2.visible = True
            IL3.visible = True
            IL4.visible = True
        Else
            Me.Controls("IA" & i & "1").visible = False
            Me.Controls("IA" & i & "2").visible = False
            Me.Controls("IA" & i & "3").visible = False
            Me.Controls("IA" & i & "4").visible = False
            Me.Controls("IA" & i & "5").visible = False
            Me.Controls("IGRFW" & i).visible = False
            Me.Controls("IGFIG" & i).visible = False
            If i = 1 Then
                IL1.visible = False
                IL2.visible = False
                IL3.visible = False
                IL4.visible = False
            End If
        End If
    Next i
End Sub

Private Sub OWSPAN1_Change()
    OWSIZE1.visible = OWSPAN1.value <> ""
    OWNEWSIZE1.visible = OWSPAN1.value <> ""
    OWLENGTH1.visible = OWSPAN1.value <> ""
    OW1L2.visible = OWSPAN1.value <> ""
    OW1L3.visible = OWSPAN1.value <> ""
    OW1L4.visible = OWSPAN1.value <> ""
    If OWSPAN1.value <> "" Then
        OWSIZE1.value = getOpenWireSizeFromSpan(OWSPAN1.value)
        Set cell = pds.UsedRange.Find(What:=OWSPAN1.value, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
        If Not cell Is Nothing Then
            OWLENGTH1.value = Round(Utilities.OnlyNumbers(cell.offSet(-1, 0).value), 0) & "'"
        End If
    End If
End Sub

Private Sub OWSPAN2_Change()
    OWSIZE2.visible = OWSPAN2.value <> ""
    OWNEWSIZE2.visible = OWSPAN2.value <> ""
    OWLENGTH2.visible = OWSPAN2.value <> ""
    OW2L2.visible = OWSPAN2.value <> ""
    OW2L3.visible = OWSPAN2.value <> ""
    OW2L4.visible = OWSPAN2.value <> ""
    If OWSPAN2.value <> "" Then
        OWSIZE2.value = getOpenWireSizeFromSpan(OWSPAN2.value)
        Set cell = pds.UsedRange.Find(What:=OWSPAN2.value, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
        If Not cell Is Nothing Then
            OWLENGTH2.value = Round(Utilities.OnlyNumbers(cell.offSet(-1, 0).value), 0) & "'"
        End If
    End If
End Sub

Private Sub OWSPAN3_Change()
    OWSIZE3.visible = OWSPAN3.value <> ""
    OWNEWSIZE3.visible = OWSPAN3.value <> ""
    OWLENGTH3.visible = OWSPAN3.value <> ""
    OW3L2.visible = OWSPAN3.value <> ""
    OW3L3.visible = OWSPAN3.value <> ""
    OW3L4.visible = OWSPAN3.value <> ""
    If OWSPAN3.value <> "" Then
        OWSIZE3.value = getOpenWireSizeFromSpan(OWSPAN3.value)
        Set cell = pds.UsedRange.Find(What:=OWSPAN3.value, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
        If Not cell Is Nothing Then
            OWLENGTH3.value = Round(Utilities.OnlyNumbers(cell.offSet(-1, 0).value), 0) & "'"
        End If
    End If
End Sub

Private Sub OWSPAN4_Change()
    OWSIZE4.visible = OWSPAN4.value <> ""
    OWNEWSIZE4.visible = OWSPAN4.value <> ""
    OWLENGTH4.visible = OWSPAN4.value <> ""
    OW4L2.visible = OWSPAN4.value <> ""
    OW4L3.visible = OWSPAN4.value <> ""
    OW4L4.visible = OWSPAN4.value <> ""
    If OWSPAN4.value <> "" Then
        OWSIZE4.value = getOpenWireSizeFromSpan(OWSPAN4.value)
        Set cell = pds.UsedRange.Find(What:=OWSPAN4.value, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
        If Not cell Is Nothing Then
            OWLENGTH4.value = Round(Utilities.OnlyNumbers(cell.offSet(-1, 0).value), 0) & "'"
        End If
    End If
End Sub

Private Sub OWSPAN5_Change()
    OWSIZE5.visible = OWSPAN5.value <> ""
    OWNEWSIZE5.visible = OWSPAN5.value <> ""
    OWLENGTH5.visible = OWSPAN5.value <> ""
    OW5L2.visible = OWSPAN5.value <> ""
    OW5L3.visible = OWSPAN5.value <> ""
    OW5L4.visible = OWSPAN5.value <> ""
    If OWSPAN5.value <> "" Then
        OWSIZE5.value = getOpenWireSizeFromSpan(OWSPAN5.value)
        Set cell = pds.UsedRange.Find(What:=OWSPAN5.value, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
        If Not cell Is Nothing Then
            OWLENGTH5.value = Round(Utilities.OnlyNumbers(cell.offSet(-1, 0).value), 0) & "'"
        End If
    End If
End Sub

Private Function getOpenWireSizeFromSpan(Span As String) As String
    Dim found As Integer: found = 0
    Dim name As Variant
    Dim openWireSizes As String: output = ""
    For i = 1 To 12
        For Each name In pds.Names
            If name.name = "'" & pds.name & "'" & "!" & "TOPOLE" & i Then
                If pds.Range("TOPOLE" & i).value = Span Then found = i
                Exit For
            End If
        Next name
        If found > 0 Then Exit For
    Next i
    
    For i = 0 To 50
        If pds.Range("UTTYPE").offSet(i, 0).Interior.color = 16777215 Then Exit For
        If pds.Range("UTTYPE").offSet(i, 0) = "" Then Exit For
        If pds.Range("UTTYPE").offSet(i, 0).text = "STLT. BOTTOM BRKT." Then Exit For
        If InStr(pds.Range("UTTYPE").offSet(i, 0).text, "OW") > 0 Then
            If Trim(Replace(pds.Range("UTMIDSPAN" & found).offSet(i, 0).value, "-", "")) <> "" Then
                openWireSizes = openWireSizes & ExtractNumbers(pds.Range("UTSIZE").offSet(i, 0).value) & "-"
            End If
        End If
    Next i
    
    If Len(openWireSizes) > 0 Then openWireSizes = Left(openWireSizes, Len(openWireSizes) - 1)
    
    getOpenWireSizeFromSpan = openWireSizes
    
End Function

Private Function ExtractNumbers(ByVal inputString As String) As String
    Dim i As Long
    Dim NumericString As String

    For i = 1 To Len(inputString)
        If IsNumeric(Mid(inputString, i, 1)) Then
            NumericString = NumericString & Mid(inputString, i, 1)
        End If
    Next i

    ExtractNumbers = NumericString
End Function

Private Sub PRIT1_Change()
    If PRIT1.value = "RP" Then
        PRI1.value = priRP1
        PRIC1.value = PreviousValueRPPRIC1
    ElseIf PRIT1.value = "IN" Then
        PRI1.value = priIN1
        PRIC1.value = PreviousValueINPRIC1
    ElseIf PRIT1.value = "RM" Then
        PRI1.value = priRM1
        PRIC1.value = PreviousValueRMPRIC1
    End If
End Sub

Private Sub PRI1_Click()
    PRIC1.visible = PRI1.value
    
    If PRIT1.value = "RP" Then
        If priRP1 <> PRI1.value Then
            priRP1 = PRI1.value
            If PRI1.value Then
                UAP.value = CDbl(UAP.value) - (CInt(PRIC1.value) * 2)
            Else
                UAP.value = CDbl(UAP.value) + (CInt(PRIC1.value) * 2)
            End If
        End If
    ElseIf PRIT1.value = "IN" Then
        If priIN1 <> PRI1.value Then
            priIN1 = PRI1.value
            If PRI1.value Then
                UAP.value = CDbl(UAP.value) - (CInt(PRIC1.value))
            Else
                UAP.value = CDbl(UAP.value) + (CInt(PRIC1.value))
            End If
        End If
    ElseIf PRIT1.value = "RM" Then
        If priRM1 <> PRI1.value Then
            priRM1 = PRI1.value
            If PRI1.value Then
                UAP.value = CDbl(UAP.value) - (CInt(PRIC1.value))
            Else
                UAP.value = CDbl(UAP.value) + (CInt(PRIC1.value))
            End If
        End If
    End If
End Sub

Private Sub PRIC1_Change()
    If PRIT1.value = "RP" Then
        If PreviousValueRPPRIC1 <> PRIC1.value Then
            UAP.value = CDbl(UAP.value) - (2 * (PRIC1.value - PreviousValueRPPRIC1))
            PreviousValueRPPRIC1 = PRIC1.value
        End If
    ElseIf PRIT1.value = "IN" Then
        If PreviousValueINPRIC1 <> PRIC1.value Then
            UAP.value = CDbl(UAP.value) - (PRIC1.value - PreviousValueINPRIC1)
            PreviousValueINPRIC1 = PRIC1.value
        End If
    ElseIf PRIT1.value = "RM" Then
        If PreviousValueRMPRIC1 <> PRIC1.value Then
            UAP.value = CDbl(UAP.value) - (PRIC1.value - PreviousValueRMPRIC1)
            PreviousValueRMPRIC1 = PRIC1.value
        End If
    End If
End Sub

Private Sub PRIT2_Change()
    If PRIT2.value = "RP" Then
        PRI2.value = priRP2
        PRIC2.value = PreviousValueRPPRIC2
    ElseIf PRIT2.value = "IN" Then
        PRI2.value = priIN2
        PRIC2.value = PreviousValueINPRIC2
    ElseIf PRIT2.value = "RM" Then
        PRI2.value = priRM2
        PRIC2.value = PreviousValueRMPRIC2
    End If
End Sub

Private Sub PRI2_Click()
    PRIC2.visible = PRI2.value
    
    If PRIT2.value = "RP" Then
        If priRP2 <> PRI2.value Then
            priRP2 = PRI2.value
            If PRI2.value Then
                UAP.value = CDbl(UAP.value) - (CInt(PRIC2.value) * 2)
            Else
                UAP.value = CDbl(UAP.value) + (CInt(PRIC2.value) * 2)
            End If
        End If
    ElseIf PRIT2.value = "IN" Then
        If priIN2 <> PRI2.value Then
            priIN2 = PRI2.value
            If PRI2.value Then
                UAP.value = CDbl(UAP.value) - (CInt(PRIC2.value))
            Else
                UAP.value = CDbl(UAP.value) + (CInt(PRIC2.value))
            End If
        End If
    ElseIf PRIT2.value = "RM" Then
        If priRM2 <> PRI2.value Then
            priRM2 = PRI2.value
            If PRI2.value Then
                UAP.value = CDbl(UAP.value) - (CInt(PRIC2.value))
            Else
                UAP.value = CDbl(UAP.value) + (CInt(PRIC2.value))
            End If
        End If
    End If
End Sub

Private Sub PRIC2_Change()
    If PRIT2.value = "RP" Then
        If PreviousValueRPPRIC2 <> PRIC2.value Then
            UAP.value = CDbl(UAP.value) - (2 * (PRIC2.value - PreviousValueRPPRIC2))
            PreviousValueRPPRIC2 = PRIC2.value
        End If
    ElseIf PRIT2.value = "IN" Then
        If PreviousValueINPRIC2 <> PRIC2.value Then
            UAP.value = CDbl(UAP.value) - (PRIC2.value - PreviousValueINPRIC2)
            PreviousValueINPRIC2 = PRIC2.value
        End If
    ElseIf PRIT2.value = "RM" Then
        If PreviousValueRMPRIC2 <> PRIC2.value Then
            UAP.value = CDbl(UAP.value) - (PRIC2.value - PreviousValueRMPRIC2)
            PreviousValueRMPRIC2 = PRIC2.value
        End If
    End If
End Sub

Private Sub PRIT3_Change()
    If PRIT3.value = "RP" Then
        PRI3.value = priRP3
        PRIC3.value = PreviousValueRPPRIC3
    ElseIf PRIT3.value = "IN" Then
        PRI3.value = priIN3
        PRIC3.value = PreviousValueINPRIC3
    ElseIf PRIT3.value = "RM" Then
        PRI3.value = priRM3
        PRIC3.value = PreviousValueRMPRIC3
    End If
End Sub

Private Sub PRI3_Click()
    PRIC3.visible = PRI3.value
    
    If PRIT3.value = "RP" Then
        If priRP3 <> PRI3.value Then
            priRP3 = PRI3.value
            If PRI3.value Then
                UAP.value = CDbl(UAP.value) - (CInt(PRIC3.value) * 2)
            Else
                UAP.value = CDbl(UAP.value) + (CInt(PRIC3.value) * 2)
            End If
        End If
    ElseIf PRIT3.value = "IN" Then
        If priIN3 <> PRI3.value Then
            priIN3 = PRI3.value
            If PRI3.value Then
                UAP.value = CDbl(UAP.value) - (CInt(PRIC3.value))
            Else
                UAP.value = CDbl(UAP.value) + (CInt(PRIC3.value))
            End If
        End If
    ElseIf PRIT3.value = "RM" Then
        If priRM3 <> PRI3.value Then
            priRM3 = PRI3.value
            If PRI3.value Then
                UAP.value = CDbl(UAP.value) - (CInt(PRIC3.value))
            Else
                UAP.value = CDbl(UAP.value) + (CInt(PRIC3.value))
            End If
        End If
    End If
End Sub

Private Sub PRIC3_Change()
    If PRIT3.value = "RP" Then
        If PreviousValueRPPRIC3 <> PRIC3.value Then
            UAP.value = CDbl(UAP.value) - (2 * (PRIC3.value - PreviousValueRPPRIC3))
            PreviousValueRPPRIC3 = PRIC3.value
        End If
    ElseIf PRIT3.value = "IN" Then
        If PreviousValueINPRIC3 <> PRIC3.value Then
            UAP.value = CDbl(UAP.value) - (PRIC3.value - PreviousValueINPRIC3)
            PreviousValueINPRIC3 = PRIC3.value
        End If
    ElseIf PRIT3.value = "RM" Then
        If PreviousValueRMPRIC3 <> PRIC3.value Then
            UAP.value = CDbl(UAP.value) - (PRIC3.value - PreviousValueRMPRIC3)
            PreviousValueRMPRIC3 = PRIC3.value
        End If
    End If
End Sub

Private Sub PRIT4_Change()
    If PRIT4.value = "RP" Then
        PRI4.value = priRP4
        PRIC4.value = PreviousValueRPPRIC4
    ElseIf PRIT4.value = "IN" Then
        PRI4.value = priIN4
        PRIC4.value = PreviousValueINPRIC4
    ElseIf PRIT4.value = "RM" Then
        PRI4.value = priRM4
        PRIC4.value = PreviousValueRMPRIC4
    End If
End Sub

Private Sub PRI4_Click()
    PRIC4.visible = PRI4.value
    
    If PRIT4.value = "RP" Then
        If priRP4 <> PRI4.value Then
            priRP4 = PRI4.value
            If PRI4.value Then
                UAP.value = CDbl(UAP.value) - (CInt(PRIC4.value) * 2)
            Else
                UAP.value = CDbl(UAP.value) + (CInt(PRIC4.value) * 2)
            End If
        End If
    ElseIf PRIT4.value = "IN" Then
        If priIN4 <> PRI4.value Then
            priIN4 = PRI4.value
            If PRI4.value Then
                UAP.value = CDbl(UAP.value) - (CInt(PRIC4.value))
            Else
                UAP.value = CDbl(UAP.value) + (CInt(PRIC4.value))
            End If
        End If
    ElseIf PRIT4.value = "RM" Then
        If priRM4 <> PRI4.value Then
            priRM4 = PRI4.value
            If PRI4.value Then
                UAP.value = CDbl(UAP.value) - (CInt(PRIC4.value))
            Else
                UAP.value = CDbl(UAP.value) + (CInt(PRIC4.value))
            End If
        End If
    End If
End Sub

Private Sub PRIC4_Change()
    If PRIT4.value = "RP" Then
        If PreviousValueRPPRIC4 <> PRIC4.value Then
            UAP.value = CDbl(UAP.value) - (2 * (PRIC4.value - PreviousValueRPPRIC4))
            PreviousValueRPPRIC4 = PRIC4.value
        End If
    ElseIf PRIT4.value = "IN" Then
        If PreviousValueINPRIC4 <> PRIC4.value Then
            UAP.value = CDbl(UAP.value) - (PRIC4.value - PreviousValueINPRIC4)
            PreviousValueINPRIC4 = PRIC4.value
        End If
    ElseIf PRIT4.value = "RM" Then
        If PreviousValueRMPRIC4 <> PRIC4.value Then
            UAP.value = CDbl(UAP.value) - (PRIC4.value - PreviousValueRMPRIC4)
            PreviousValueRMPRIC4 = PRIC4.value
        End If
    End If
End Sub

Private Sub PRIT5_Change()
    If PRIT5.value = "RP" Then
        PRI5.value = priRP5
        PRIC5.value = PreviousValueRPPRIC5
    ElseIf PRIT5.value = "IN" Then
        PRI5.value = priIN5
        PRIC5.value = PreviousValueINPRIC5
    ElseIf PRIT5.value = "RM" Then
        PRI5.value = priRM5
        PRIC5.value = PreviousValueRMPRIC5
    End If
End Sub

Private Sub PRI5_Click()
    PRIC5.visible = PRI5.value
    
    If PRIT5.value = "RP" Then
        If priRP5 <> PRI5.value Then
            priRP5 = PRI5.value
            If PRI5.value Then
                UAP.value = CDbl(UAP.value) - (CInt(PRIC5.value) * 2)
            Else
                UAP.value = CDbl(UAP.value) + (CInt(PRIC5.value) * 2)
            End If
        End If
    ElseIf PRIT5.value = "IN" Then
        If priIN5 <> PRI5.value Then
            priIN5 = PRI5.value
            If PRI5.value Then
                UAP.value = CDbl(UAP.value) - (CInt(PRIC5.value))
            Else
                UAP.value = CDbl(UAP.value) + (CInt(PRIC5.value))
            End If
        End If
    ElseIf PRIT5.value = "RM" Then
        If priRM5 <> PRI5.value Then
            priRM5 = PRI5.value
            If PRI5.value Then
                UAP.value = CDbl(UAP.value) - (CInt(PRIC5.value))
            Else
                UAP.value = CDbl(UAP.value) + (CInt(PRIC5.value))
            End If
        End If
    End If
End Sub

Private Sub PRIC5_Change()
    If PRIT5.value = "RP" Then
        If PreviousValueRPPRIC5 <> PRIC5.value Then
            UAP.value = CDbl(UAP.value) - (2 * (PRIC5.value - PreviousValueRPPRIC5))
            PreviousValueRPPRIC5 = PRIC5.value
        End If
    ElseIf PRIT5.value = "IN" Then
        If PreviousValueINPRIC5 <> PRIC5.value Then
            UAP.value = CDbl(UAP.value) - (PRIC5.value - PreviousValueINPRIC5)
            PreviousValueINPRIC5 = PRIC5.value
        End If
    ElseIf PRIT5.value = "RM" Then
        If PreviousValueRMPRIC5 <> PRIC5.value Then
            UAP.value = CDbl(UAP.value) - (PRIC5.value - PreviousValueRMPRIC5)
            PreviousValueRMPRIC5 = PRIC5.value
        End If
    End If
End Sub

Private Sub PRIT6_Change()
    If PRIT6.value = "RP" Then
        PRI6.value = priRP6
        PRIC6.value = PreviousValueRPPRIC6
    ElseIf PRIT6.value = "IN" Then
        PRI6.value = priIN6
        PRIC6.value = PreviousValueINPRIC6
    ElseIf PRIT6.value = "RM" Then
        PRI6.value = priRM6
        PRIC6.value = PreviousValueRMPRIC6
    End If
End Sub

Private Sub PRI6_Click()
    PRIC6.visible = PRI6.value
    
    If PRIT6.value = "RP" Then
        If priRP6 <> PRI6.value Then
            priRP6 = PRI6.value
            If PRI6.value Then
                UAP.value = CDbl(UAP.value) - (CInt(PRIC6.value))
            Else
                UAP.value = CDbl(UAP.value) + (CInt(PRIC6.value))
            End If
        End If
    ElseIf PRIT6.value = "IN" Then
        If priIN6 <> PRI6.value Then
            priIN6 = PRI6.value
            If PRI6.value Then
                UAP.value = CDbl(UAP.value) - (CInt(PRIC6.value) * 0.5)
            Else
                UAP.value = CDbl(UAP.value) + (CInt(PRIC6.value) * 0.5)
            End If
        End If
    ElseIf PRIT6.value = "RM" Then
        If priRM6 <> PRI6.value Then
            priRM6 = PRI6.value
            If PRI6.value Then
                UAP.value = CDbl(UAP.value) - (CInt(PRIC6.value) * 0.5)
            Else
                UAP.value = CDbl(UAP.value) + (CInt(PRIC6.value) * 0.5)
            End If
        End If
    End If
End Sub

Private Sub PRIC6_Change()
    If PRIT6.value = "RP" Then
        If PreviousValueRPPRIC6 <> PRIC6.value Then
            UAP.value = CDbl(UAP.value) - (PRIC6.value - PreviousValueRPPRIC6)
            PreviousValueRPPRIC6 = PRIC6.value
        End If
    ElseIf PRIT6.value = "IN" Then
        If PreviousValueINPRIC6 <> PRIC6.value Then
            UAP.value = CDbl(UAP.value) - (0.5 * (PRIC6.value - PreviousValueINPRIC6))
            PreviousValueINPRIC6 = PRIC6.value
        End If
    ElseIf PRIT6.value = "RM" Then
        If PreviousValueRMPRIC6 <> PRIC6.value Then
            UAP.value = CDbl(UAP.value) - (0.5 * (PRIC6.value - PreviousValueRMPRIC6))
            PreviousValueRMPRIC6 = PRIC6.value
        End If
    End If
End Sub

Private Sub PRIT7_Change()
    If PRIT7.value = "RP" Then
        PRI7.value = priRP7
        PRIC7.value = PreviousValueRPPRIC7
    ElseIf PRIT7.value = "IN" Then
        PRI7.value = priIN7
        PRIC7.value = PreviousValueINPRIC7
    ElseIf PRIT7.value = "RM" Then
        PRI7.value = priRM7
        PRIC7.value = PreviousValueRMPRIC7
    End If
End Sub

Private Sub PRI7_Click()
    PRIC7.visible = PRI7.value
    
    If PRIT7.value = "RP" Then
        If priRP7 <> PRI7.value Then
            priRP7 = PRI7.value
            If PRI7.value Then
                UAP.value = CDbl(UAP.value) - (CInt(PRIC7.value) * 2)
            Else
                UAP.value = CDbl(UAP.value) + (CInt(PRIC7.value) * 2)
            End If
        End If
    ElseIf PRIT7.value = "IN" Then
        If priIN7 <> PRI7.value Then
            priIN7 = PRI7.value
            If PRI7.value Then
                UAP.value = CDbl(UAP.value) - (CInt(PRIC7.value))
            Else
                UAP.value = CDbl(UAP.value) + (CInt(PRIC7.value))
            End If
        End If
    ElseIf PRIT7.value = "RM" Then
        If priRM7 <> PRI7.value Then
            priRM7 = PRI7.value
            If PRI7.value Then
                UAP.value = CDbl(UAP.value) - (CInt(PRIC7.value))
            Else
                UAP.value = CDbl(UAP.value) + (CInt(PRIC7.value))
            End If
        End If
    End If
End Sub

Private Sub PRIC7_Change()
    If PRIT7.value = "RP" Then
        If PreviousValueRPPRIC7 <> PRIC7.value Then
            UAP.value = CDbl(UAP.value) - (2 * (PRIC7.value - PreviousValueRPPRIC7))
            PreviousValueRPPRIC7 = PRIC7.value
        End If
    ElseIf PRIT7.value = "IN" Then
        If PreviousValueINPRIC7 <> PRIC7.value Then
            UAP.value = CDbl(UAP.value) - (PRIC7.value - PreviousValueINPRIC7)
            PreviousValueINPRIC7 = PRIC7.value
        End If
    ElseIf PRIT7.value = "RM" Then
        If PreviousValueRMPRIC7 <> PRIC7.value Then
            UAP.value = CDbl(UAP.value) - (PRIC7.value - PreviousValueRMPRIC7)
            PreviousValueRMPRIC7 = PRIC7.value
        End If
    End If
End Sub

Private Sub REC1_Click()
    RECA1.visible = REC1.value
    RECO1.visible = REC1.value
    
    If REC1.value Then
        UASN.value = CInt(UASN.value) - (CInt(RECA1.value) * 6)
    Else
        UASN.value = CInt(UASN.value) + (CInt(RECA1.value) * 6)
    End If
End Sub

Private Sub REC2_Click()
    RECA2.visible = REC2.value
    RECO2.visible = REC2.value
    
    If REC2.value Then
        UASN.value = CInt(UASN.value) - (CInt(RECA2.value) * 3)
    Else
        UASN.value = CInt(UASN.value) + (CInt(RECA2.value) * 3)
    End If
End Sub

Private Sub REC3_Click()
    RECA3.visible = REC3.value
    RECO3.visible = REC3.value
    
    If REC3.value Then
        UASN.value = CInt(UASN.value) - (CInt(RECA3.value) * 6)
    Else
        UASN.value = CInt(UASN.value) + (CInt(RECA3.value) * 6)
    End If
End Sub

Private Sub REC4_Click()
    RECA4.visible = REC4.value
    RECO4.visible = REC4.value
    
    If REC4.value Then
        UASN.value = CInt(UASN.value) - (CInt(RECA4.value) * 6)
    Else
        UASN.value = CInt(UASN.value) + (CInt(RECA4.value) * 6)
    End If
End Sub

Private Sub RECA1_Change()
    If REC1.value Then
        If PreviousValueRECA1 <> RECA1.value Then
            UASN.value = CDbl(UASN.value) - ((RECA1.value - PreviousValueRECA1) * 6)
            PreviousValueRECA1 = RECA1.value
        End If
    End If
End Sub

Private Sub RECA2_Change()
    If REC2.value Then
        If PreviousValueRECA2 <> RECA2.value Then
            UASN.value = CDbl(UASN.value) - ((RECA2.value - PreviousValueRECA2) * 3)
            PreviousValueRECA2 = RECA2.value
        End If
    End If
End Sub

Private Sub RECA3_Change()
    If REC3.value Then
        If PreviousValueRECA3 <> RECA3.value Then
            UASN.value = CDbl(UASN.value) - ((RECA3.value - PreviousValueRECA3) * 6)
            PreviousValueRECA3 = RECA3.value
        End If
    End If
End Sub

Private Sub RECA4_Change()
    If REC4.value Then
        If PreviousValueRECA4 <> RECA4.value Then
            UASN.value = CDbl(UASN.value) - ((RECA4.value - PreviousValueRECA4) * 6)
            PreviousValueRECA4 = RECA4.value
        End If
    End If
End Sub

Private Sub RPBSL_Click()
    BSL.value = RPBSL.value
    RPBSLRFW.visible = BSL.value
    BSLRFW.visible = BSL.value
    SLRFWL.visible = BSL.value
End Sub

Private Sub RPBSLRFW_Change()
    BSLRFW.value = RPBSLRFW.value
End Sub

Private Sub RPP_Click()
    MHL.visible = RPP.value
    MH1.visible = RPP.value
    MHC1.visible = MH1.value And RPP.value
    MHT1.visible = RPP.value
    MH2.visible = RPP.value
    MHC2.visible = MH2.value And RPP.value
    MHT2.visible = RPP.value
    RPPSIZE.visible = RPP.value
    RPPCLASS.visible = RPP.value
    RPPRFW.visible = RPP.value
    RDST.visible = RPP.value
    RDSV.visible = RPP.value
    DGFV.visible = RPP.value And RPDG.value
    DGFL.visible = RPP.value And RPDG.value
    RPSWB.visible = RPP.value And RPDG.value
    DSL.visible = RPP.value
    DSV.visible = RPP.value
    PSL.visible = RPP.value
    PSV.visible = RPP.value
    BAFIG.visible = RPP.value
    FIGL.visible = priBool And RPP.value
    FIG.visible = priBool And RPP.value
    TRT.visible = RPP.value
    RPT.visible = RPP.value
    RPR.visible = RPP.value
    TRSL.visible = RPP.value
    SLM1.visible = RPP.value And TRSL
    SLM2.visible = RPP.value And TRSL
    SLM3.value = SLM1.value And RPP.value
    SLM4.value = SLM2.value And RPP.value
    TRSPG.visible = RPP.value
    RPR.visible = RPP.value
    RPDG.visible = RPP.value
    TRSIZE.visible = RPP.value And (TRT.value Or RPT.value)
    TRRPSIZE.visible = RPP.value And RPT.value
    STLTBTMBRKT.visible = RPP.value And TRSL.value
    COLA.visible = RPP.value And (TRT.value Or RPT.value)
    TRLETTER.visible = RPP.value And (TRT.value Or RPT.value)
    RPBSL.visible = RPP.value
    RPBSLRFW.visible = RPP.value And RPBSL.value
    CETL.visible = RPP.value
    CETV.visible = RPP.value
    UAPL.visible = RPP.value
    UAP.visible = RPP.value
    FUT1.visible = RPP.value
    FUT2.visible = RPP.value
    FUT3.visible = RPP.value
    PRIT1.visible = RPP.value
    PRIT2.visible = RPP.value
    PRIT3.visible = RPP.value
    PRIT4.visible = RPP.value
    PRIT5.visible = RPP.value
    PRIT6.visible = RPP.value
    PRIT7.visible = RPP.value
    PRI1.visible = RPP.value
    PRI2.visible = RPP.value
    PRI3.visible = RPP.value
    PRI4.visible = RPP.value
    PRI5.visible = RPP.value
    PRI6.visible = RPP.value
    PRI7.visible = RPP.value
    PRIC1.visible = PRI1.value And RPP.value
    PRIC2.visible = PRI2.value And RPP.value
    PRIC3.visible = PRI3.value And RPP.value
    PRIC4.visible = PRI4.value And RPP.value
    PRIC5.visible = PRI5.value And RPP.value
    PRIC6.visible = PRI6.value And RPP.value
    PRIC7.visible = PRI7.value And RPP.value
    USNL.visible = RPP.value
    UASN.visible = RPP.value
    USNL2.visible = RPP.value
    UASN2.visible = RPP.value
    SNT1.visible = RPP.value
    SNT2.visible = RPP.value
    SNT3.visible = RPP.value
    SNT4.visible = RPP.value
    SNT5.visible = RPP.value
    SNT6.visible = RPP.value
    SNT7.visible = RPP.value
    SN1.visible = RPP.value
    SN2.visible = RPP.value
    SN3.visible = RPP.value
    SN4.visible = RPP.value
    SN5.visible = RPP.value
    SN6.visible = RPP.value
    SN7.visible = RPP.value
    SNC1.visible = SN1.value And RPP.value
    SNC2.visible = SN2.value And RPP.value
    SNC3.visible = SN3.value And RPP.value
    SNC4.visible = SN4.value And RPP.value
    SNC5.visible = SN5.value And RPP.value
    SNC6.visible = SN6.value And RPP.value
    SNC7.visible = SN7.value And RPP.value
    FUL.visible = RPP.value
    FU1.visible = RPP.value
    FU2.visible = RPP.value
    FU3.visible = RPP.value
    FUC1.visible = FU1.value And RPP.value
    FUC2.visible = FU2.value And RPP.value
    FUC3.visible = FU3.value And RPP.value
    NPL.visible = RPP.value
    RFWL.visible = RPP.value
    XRL.visible = RPP.value And RPT.value
    XSL.visible = RPP.value And (TRT.value Or RPT.value)
    CLL.visible = RPP.value And (TRT.value Or RPT.value)
    XDL.visible = RPP.value And (TRT.value Or RPT.value)
    SLL.visible = RPP.value And TRSL.value
    SLRFWL.visible = RPP.value And RPBSL.value
End Sub

Private Sub RPPSIZE_Change()
    Call determineRecDSpace
End Sub

Private Sub determineRecDSpace()
    If pole.primaries.count > 0 Then
        If RPPSIZE.value >= 50 Then
            If TRT.value Or RPT.value Then
                RDSV.caption = "[" & Utilities.inchesToFeetInches(((((RPPSIZE * 0.9) - 2) - 32) * 12) - 11 - IIf(Utilities.convertToInches(PSV.value) > 0, Utilities.convertToInches(PSV.value), 0)) & "]"
                If Utilities.convertToInches(RDSV.caption) < 60 Then RDSV.caption = "[5'0""]"
            Else
                RDSV.caption = "[" & Utilities.inchesToFeetInches(((((RPPSIZE * 0.9) - 2) - 35) * 12) - 11 - IIf(Utilities.convertToInches(PSV.value) > 0, Utilities.convertToInches(PSV.value), 0)) & "]"
                If Utilities.convertToInches(RDSV.caption) < 60 Then RDSV.caption = "[5'0""]"
            End If
        Else
            If TRT.value Or RPT.value Then
                If InStr(TRLETTER.value, "C") = 1 Then
                    RDSV.caption = "[7'8""]"
                Else
                    RDSV.caption = "[7'6""]"
                End If
            Else
                RDSV.caption = "[5'0""]"
            End If
        End If
    End If
End Sub

Private Sub RSL_Click()
    RSLD.visible = RSL.value
    RSLRFW.visible = RSL.value
    RFWL4.visible = RSL.value
End Sub

Private Sub RSPANS_Change()
    
    If RSPANS.value >= 1 Then
        OWSPAN1.visible = True
        OW1L1.visible = True
    Else
        OWSPAN1.visible = False
        OW1L1.visible = False
    End If
    
    If RSPANS.value >= 2 Then
        OWSPAN2.visible = True
        OW2L1.visible = True
    Else
        OWSPAN2.visible = False
        OW2L1.visible = False
    End If
    
    If RSPANS.value >= 3 Then
        OWSPAN3.visible = True
        OW3L1.visible = True
    Else
        OWSPAN3.visible = False
        OW3L1.visible = False
    End If
    
    If RSPANS.value >= 4 Then
        OWSPAN4.visible = True
        OW4L1.visible = True
    Else
        OWSPAN4.visible = False
        OW4L1.visible = False
    End If
    
    If RSPANS.value >= 5 Then
        OWSPAN5.visible = True
        OW5L1.visible = True
    Else
        OWSPAN5.visible = False
        OW5L1.visible = False
    End If
    
    OWSIZE1.visible = OWSPAN1.value <> "" And OWSPAN1.visible
    OWNEWSIZE1.visible = OWSPAN1.value <> "" And OWSPAN1.visible
    OWLENGTH1.visible = OWSPAN1.value <> "" And OWSPAN1.visible
    OW1L2.visible = OWSPAN1.value <> "" And OWSPAN1.visible
    OW1L3.visible = OWSPAN1.value <> "" And OWSPAN1.visible
    OW1L4.visible = OWSPAN1.value <> "" And OWSPAN1.visible
    
    OWSIZE2.visible = OWSPAN2.value <> "" And OWSPAN2.visible
    OWNEWSIZE2.visible = OWSPAN2.value <> "" And OWSPAN2.visible
    OWLENGTH2.visible = OWSPAN2.value <> "" And OWSPAN2.visible
    OW2L2.visible = OWSPAN2.value <> "" And OWSPAN2.visible
    OW2L3.visible = OWSPAN2.value <> "" And OWSPAN2.visible
    OW2L4.visible = OWSPAN2.value <> "" And OWSPAN2.visible
    
    OWSIZE3.visible = OWSPAN3.value <> "" And OWSPAN3.visible
    OWNEWSIZE3.visible = OWSPAN3.value <> "" And OWSPAN3.visible
    OWLENGTH3.visible = OWSPAN3.value <> "" And OWSPAN3.visible
    OW3L2.visible = OWSPAN3.value <> "" And OWSPAN3.visible
    OW3L3.visible = OWSPAN3.value <> "" And OWSPAN3.visible
    OW3L4.visible = OWSPAN3.value <> "" And OWSPAN3.visible
    
    OWSIZE4.visible = OWSPAN4.value <> "" And OWSPAN4.visible
    OWNEWSIZE4.visible = OWSPAN4.value <> "" And OWSPAN4.visible
    OWLENGTH4.visible = OWSPAN4.value <> "" And OWSPAN4.visible
    OW4L2.visible = OWSPAN4.value <> "" And OWSPAN4.visible
    OW4L3.visible = OWSPAN4.value <> "" And OWSPAN4.visible
    OW4L4.visible = OWSPAN4.value <> "" And OWSPAN4.visible
    
    OWSIZE5.visible = OWSPAN5.value <> "" And OWSPAN5.visible
    OWNEWSIZE5.visible = OWSPAN5.value <> "" And OWSPAN5.visible
    OWLENGTH5.visible = OWSPAN5.value <> "" And OWSPAN5.visible
    OW5L2.visible = OWSPAN5.value <> "" And OWSPAN5.visible
    OW5L3.visible = OWSPAN5.value <> "" And OWSPAN5.visible
    OW5L4.visible = OWSPAN5.value <> "" And OWSPAN5.visible
    
End Sub

Private Sub SLDL_Click()
    SLDLRFW.visible = SLDL.value
    RFWL3.visible = SLDL.value
End Sub

Private Sub SLM1_Click()
    If RPP.value Then SLM3.value = SLM1.value
End Sub

Private Sub SLM2_Click()
    If RPP.value Then SLM4.value = SLM2.value
End Sub

Private Sub SLM3_Click()
    SLM1.value = SLM3.value
End Sub

Private Sub SLM4_Click()
    SLM2.value = SLM4.value
End Sub

Private Sub TRSL_Click()
    SLL.visible = RPP.value And TRSL.value
    STLTBTMBRKT.visible = RPP.value And TRSL.value
End Sub

Private Sub TSDL_Click()
    TSDLRFW.visible = TSDL.value
    RFWL1.visible = TSDL.value
End Sub

Private Sub CommandButton1_Click()
    Set figuresUsed = New Scripting.Dictionary

    crewNotes = ""
    installNotes = ""
    removeNotes = ""
    replaceNotes = ""
    transferNotes = ""
    notes = ""
    replace11kCount = 0
    replace20kCount = 0
    Set commTransfers = New Collection

    Call generateDownguysSection
    If RPP.value Then Call generateReplacePoleSection
    Call generateReconductorSection
    Call generateBasicSection
    Call generateNotesSection
    Call compileCrewNotes
    
    Errors = ""
    warnings = ""
    
    Errors = findErrors
    warnings = findWarnings
    
    If Errors <> "" Then
        MsgBox Errors & vbLf & warnings
        Exit Sub
    End If
    
    If warnings <> "" Then
        answer = MsgBox(warnings & vbLf & "Do you want to generate crew notes even with the warnings? Select no or cancel to fix the issues first.", vbYesNoCancel + vbQuestion, "Confirm")
        If answer = vbNo Or answer = vbCancel Then Exit Sub
    End If
    
    If crewNotes <> "" Then
        crewNotes = Left(crewNotes, Len(crewNotes) - 1)
        Dim project As project: Set project = New project
        Call Utilities.applyStandardAbbreviations(crewNotes)

        Dim DataObj As DataObject: Set DataObj = New DataObject
        DataObj.SetText crewNotes
        DataObj.PutInClipboard
    
        If Trim(pds.Range("ALTONE")) = "" Then
            pds.Range("ALTONE").value = crewNotes
            Call Figures.getSheetFigures(pds)
        End If
    
        MsgBox crewNotes
    Else
        MsgBox "No work selected"
    End If
    
End Sub

Private Function findErrors() As String
    Errors = ""
    
    If RPP.value And PSV.value = "" And BAFIG.value Then
        Errors = Errors & "Error: P space shouldn't be blank if there's a buckarm on the pole." & vbLf
    End If
    
    If RPP.value And pole.commWires.count > CInt(CETV.value) And Not RPPNOTE1.value Then
        Errors = Errors & "Error: Not all comms accounted for on pole replace, either have CE transfer more comms or add top pole note." & vbLf
    End If
    
    If RPP.value And priBool And DSV.value = "" Then
        Errors = Errors & "Error: D space shouldn't be blank if there's primary on the pole." & vbLf
    End If
    
    If Not RPPNOTE2.value And RPP.value And (TRT.value Or RPT) Then
        Errors = Errors & "Error: Outage note not selected, outage needs to be selected when replacing a pole with a transformer." & vbLf
    ElseIf RPP.value And (ohsCount > 0) And Not RPPNOTE2.value Then
        Errors = Errors & "Error: Outage note not selected, outage needs to be selected when replacing services." & vbLf
    ElseIf Not RPPNOTE2.value And (REC1.value Or REC2.value Or REC3.value Or REC4.value Or (OWSPAN1.value <> "" And OWSPAN1.visible) Or (OWSPAN2.value <> "" And OWSPAN2.visible) Or (OWSPAN3.value <> "" And OWSPAN3.visible) Or (OWSPAN4.value <> "" And OWSPAN4.visible) Or (OWSPAN5.value <> "" And OWSPAN5.visible)) Then
        Errors = Errors & "Error: Outage note not selected, outage needs to be selected when reconductoring." & vbLf
    End If
    
    If RPP.value And (TRT.value Or RPT.value) And COLA.value = "" Then
        Errors = Errors & "Error: If transfering or replacing transformer, CO/LA configuration needs to be chosen. Choose Other if nothing fits." & vbLf
    End If
    
    If RPP.value And TRSL.value And STLTBTMBRKT.value = "" Then
        Errors = Errors & "Error: If transfering streetlight, streetlight attach height needs to be filled in." & vbLf
    End If
    
    findErrors = Errors
End Function

Private Function findWarnings() As String
    warnings = ""
    
    If TSDL.value And Not RPPNOTE2.value Then
        warnings = warnings & "Warning: Outage required not selected when trimming driploop." & vbLf
    End If
    
    If RPP.value And PSV.value = "" And Not BAFIG.value Then
        warnings = warnings & "Warning: P space is blank, make sure there's no P space needed on this pole." & vbLf
    End If
    
    If RPP.value And priBool And Not (FU1.value Or FU2.value Or FU3.value) Then
        warnings = warnings & "Warning: Make sure the pole doesn't need any framing hardware replaced." & vbLf
    End If
    
    If RPP.value And UAP <> 0 Then
        warnings = warnings & "Warning: Unaccounted Primary should be 0 to account for all primary spans." & vbLf
    End If
    
    If RPP.value And UASN <> 0 Then
        warnings = warnings & "Warning: Unaccounted OW/Sec/Neut should be 0 to account for all OW/Sec/Neut spans." & vbLf
    End If
    
    If RPP.value And ohsCount = 0 And (TRT.value Or RPT.value Or secBool Or owBool) And Not RPPNOTE2.value Then
        warnings = warnings & "Warning: Outage note not selected, check if the transformer/secondaries power a house or streetlight." & vbLf
    End If
    
    If RPP.value And Not RPPNOTE3.value And Not RPPNOTE4.value Then
        warnings = warnings & "Warning: Make sure there's no tree or brush work required to replace this pole." & vbLf
    End If
    
    findWarnings = warnings
End Function

Private Sub generateReplacePoleSection()
    Dim replacePoleSection As String
    
    poleHeight = "-"
    If InStr(pds.Range("HEIGHT"), "(") > 0 Then
        poleHeight = Left(Trim(pds.Range("HEIGHT")), InStr(Trim(pds.Range("HEIGHT")), "(") - 1) & poleHeight
    Else
        poleHeight = Trim(pds.Range("HEIGHT")) & poleHeight
    End If
    
    replacePoleSection = replacePoleSection & poleHeight & pds.Range("CLASS") & "/" & RPPSIZE.value & "-" & RPPCLASS.value & " " & Replace(RPPRFW.value, "OTHER", "[REASON FOR WORK]") & vbLf
    If Trim(PSV.value) <> "N/A" And Trim(PSV.value) <> "" Then replacePoleSection = replacePoleSection & "P = " & Trim(PSV.value) & vbLf
    If RDSV.caption <> "" Then
        If Trim(DSV.value) <> "N/A" And Trim(DSV.value) <> "" And Trim(DSV.value) <> "[X']" Then
            replacePoleSection = replacePoleSection & "D = " & Trim(DSV.value) & vbLf
        Else
            replacePoleSection = replacePoleSection & "D = " & RDSV.caption & vbLf
        End If
    End If
    If Not pds.Range("SECONLY") Then
        If Trim(FIG.text) = "" Or Trim(FIG.text) = "OTHER" Then
            replacePoleSection = replacePoleSection & "FIGURE [INSERT FIGURE NUMBER]" & vbLf
        Else
            replacePoleSection = replacePoleSection & Split(FIG.text, ",")(0) & vbLf
        End If
    End If
    If BAFIG.value Then replacePoleSection = replacePoleSection & "FIGURE 23-229-1" & vbLf
    
    
    Dim insulators As String: insulators = "[XP-FG]"
    If pole.primaries.count = 0 Then insulators = "[XP]"
    If RPDG.value Then
        If dg11KCount - replace11kCount > 0 Or dg20KCount - replace20kCount > 0 Then
            If dg11KCount - replace11kCount > 0 Then
                replacePoleSection = replacePoleSection & IIf(dg11KCount - replace11kCount > 1, "(" & dg11KCount - replace11kCount & ")", "") & "11K-" & insulators & " GUY STRAND" & IIf(RPPSIZE.value > poleHeight, " AND EXTEND TO NEW HEIGHT", "") & vbLf
            End If
            If dg20KCount - replace20kCount > 0 Then
                replacePoleSection = replacePoleSection & IIf(dg20KCount - replace20kCount > 1, "(" & dg20KCount - replace20kCount & ")", "") & "20K-" & insulators & " GUY STRAND" & IIf(RPPSIZE.value > poleHeight, " AND EXTEND TO NEW HEIGHT", "") & vbLf
            End If
            If Me.Controls("DGFV").value = "" Or Me.Controls("DGFV" & i).value = "OTHER" Then
                replacePoleSection = replacePoleSection & "FIGURE [INSERT DOWNGUY FIGURE]" & vbLf
            Else
                replacePoleSection = replacePoleSection & Split(Me.Controls("DGFV").value, ",")(0) & vbLf
            End If
        End If
        If RPSWB.value Then
            replacePoleSection = replacePoleSection & "SIDEWALK BRACE" & vbLf
            replacePoleSection = replacePoleSection & "FIGURE 22-450-1" & vbLf
        End If
    End If
    
    If RPT.value Then
        If UCase(TRSIZE.value) <> UCase(TRRPSIZE.value) Then
            replacePoleSection = replacePoleSection & UCase(TRSIZE.value) & "/" & UCase(TRRPSIZE.value) & " TRANSFORMER"
        ElseIf UCase(TRSIZE.value) = UCase(TRRPSIZE.value) Then
            replacePoleSection = replacePoleSection & UCase(TRRPSIZE.value) & " TRANSFORMER"
        End If
        If RPPSIZE.value >= 50 Then replacePoleSection = replacePoleSection & " [@35'0""]"
        replacePoleSection = replacePoleSection & vbLf
        If COLA.ListIndex = 0 Or COLA.value = "" Then
            replacePoleSection = replacePoleSection & "[INSERT CO/LA INFO]" & vbLf
        ElseIf COLA.ListIndex = 1 Then
            replacePoleSection = replacePoleSection & "CO/LA ON SA TO LCOM" & vbLf
        ElseIf COLA.ListIndex = 2 Then
            replacePoleSection = replacePoleSection & "CO/LA ON LCOM" & vbLf
        ElseIf COLA.ListIndex = 3 Then
            replacePoleSection = replacePoleSection & "CO/LA ON SA" & vbLf
        ElseIf COLA.ListIndex = 4 Then
            replacePoleSection = replacePoleSection & "CO ON LCOM" & vbLf
            replacePoleSection = replacePoleSection & "LA ON TRANSFORMER" & vbLf
        ElseIf COLA.ListIndex = 5 Then
            replacePoleSection = replacePoleSection & "CO ON SA" & vbLf
            replacePoleSection = replacePoleSection & "LA ON TRANSFORMER" & vbLf
        ElseIf COLA.ListIndex = 6 Then
            replacePoleSection = replacePoleSection & "CO ON SA TO LCOM" & vbLf
            replacePoleSection = replacePoleSection & "LA ON TRANSFORMER" & vbLf
        End If
        If TRLETTER.value = "3 PHASE" Then
            replacePoleSection = replacePoleSection & "FIGURE 26-301-1" & vbLf
        Else
            replacePoleSection = replacePoleSection & "FIGURE 26-101-1 DETAIL " & IIf(Trim(TRLETTER.text) = "", "[A/B/C/D]", Left(TRLETTER.value, 1)) & vbLf
        End If
    End If
    
    If RPR.value Then
        If TRT.value Or RPT.value Then
            replacePoleSection = replacePoleSection & "[X]/C-[XX] AL SVC [OR SECONDARY] RISER & EXTEND MOLDING TO 12"" BELOW BOTTOM OF TRANSFORMER" & vbLf
            replacePoleSection = replacePoleSection & "FIGURE 63-20-1" & vbLf
        Else
            replacePoleSection = replacePoleSection & "[X]/C-[XX]AL SVC [OR SECONDARY] RISER & EXTEND MOLDING TO 4"" BELOW [OR ABOVE?] SECONDARY ATTACH HEIGHT" & vbLf
            replacePoleSection = replacePoleSection & "FIGURE 63-20-1" & vbLf
        End If
    End If
    
    Dim rpDict As Scripting.Dictionary: Set rpDict = New Scripting.Dictionary
    Dim inDict As Scripting.Dictionary: Set inDict = New Scripting.Dictionary
    Dim rmDict As Scripting.Dictionary: Set rmDict = New Scripting.Dictionary
    
    If fuRP1 Then rpDict("S8S") = PreviousValueRPFUC1
    If fuIN1 Then inDict("S8S") = PreviousValueINFUC1
    If fuRM1 Then rmDict("S8S") = PreviousValueRMFUC1
    
    If fuRP2 Then rpDict("D8S") = PreviousValueRPFUC2
    If fuIN2 Then inDict("D8S") = PreviousValueINFUC2
    If fuRM2 Then rmDict("D8S") = PreviousValueRMFUC2
    
    If fuRP3 Then rpDict("S8FGDE") = PreviousValueRPFUC3
    If fuIN3 Then inDict("S8FGDE") = PreviousValueINFUC3
    If fuRM3 Then rmDict("S8FGDE") = PreviousValueRMFUC3
    
    If mhRP1 Then rpDict("JUMPER SPINS") = previousValueRPMHC1
    If mhIN1 Then inDict("JUMPER SPINS") = previousValueINMHC1
    If mhRM1 Then rmDict("JUMPER SPINS") = previousValueRMMHC1
    
    Call extractFromRPINRMDicts(replacePoleSection, rpDict, inDict, rmDict)
    
    If priRP1 Or priRP2 Or snRP1 Or snRP2 Then
        spinCount = 0
        If priRP1 Then spinCount = spinCount + PreviousValueRPPRIC1
        If priRP2 Then spinCount = spinCount + (PreviousValueRPPRIC2 * 2)
        If snRP1 Then spinCount = spinCount + PreviousValueRPSNC1
        If snRP2 Then spinCount = spinCount + (PreviousValueRPSNC2 * 2)
        If mhRP1 Then spinCount = spinCount + previousValueRPMHC1
        
        rpDict("SPINS") = spinCount
    End If
    
    If priIN1 Or priIN2 Or snIN1 Or snIN2 Then
        spinCount = 0
        If priIN1 Then spinCount = spinCount + PreviousValueINPRIC1
        If priIN2 Then spinCount = spinCount + (PreviousValueINPRIC2 * 2)
        If snIN1 Then spinCount = spinCount + PreviousValueINSNC1
        If snIN2 Then spinCount = spinCount + (PreviousValueINSNC2 * 2)
        If mhIN1 Then spinCount = spinCount + previousValueINMHC1
        
        inDict("SPINS") = spinCount
    End If
    
    If priRM1 Or priRM2 Or snRM1 Or snRM2 Then
        spinCount = 0
        If priRM1 Then spinCount = spinCount + PreviousValueRMPRIC1
        If priRM2 Then spinCount = spinCount + (PreviousValueRMPRIC2 * 2)
        If snRM1 Then spinCount = spinCount + PreviousValueRMSNC1
        If snRM2 Then spinCount = spinCount + (PreviousValueRMSNC2 * 2)
        If mhRM1 Then spinCount = spinCount + previousValueRMMHC1
        
        rmDict("SPINS") = spinCount
    End If
    
    If mhRP2 Then rpDict("FG STANDOFF BRACKET") = PreviousValueRPFUC3
    If mhIN2 Then inDict("FG STANDOFF BRACKET") = PreviousValueINFUC3
    If mhRM2 Then rmDict("FG STANDOFF BRACKET") = PreviousValueRMFUC3
    
    If priRP3 Or priRP4 Then
        ptpCount = 0
        If priRP3 Then ptpCount = ptpCount + PreviousValueRPPRIC3
        If priRP4 Then ptpCount = ptpCount + (PreviousValueRPPRIC4 * 2)
        
        rpDict("PTP") = ptpCount
    End If
    
    If priIN3 Or priIN4 Then
        ptpCount = 0
        If priIN3 Then ptpCount = ptpCount + PreviousValueINPRIC3
        If priIN4 Then ptpCount = ptpCount + (PreviousValueINPRIC4 * 2)
        
        inDict("PTP") = ptpCount
    End If
    
    If priRM3 Or priRM4 Then
        ptpCount = 0
        If priRM3 Then ptpCount = ptpCount + PreviousValueRMPRIC3
        If priRM4 Then ptpCount = ptpCount + (PreviousValueRMPRIC4 * 2)
        
        rmDict("PTP") = ptpCount
    End If
    
    If priRP5 Then rpDict("VPO") = PreviousValueRPPRIC5
    If priIN5 Then inDict("VPO") = PreviousValueINPRIC5
    If priRM5 Then rmDict("VPO") = PreviousValueRMPRIC5
    
    If priRP6 Then rpDict("PRIMARY DEADEND") = PreviousValueRPPRIC6
    If priIN6 Then inDict("PRIMARY DEADEND") = PreviousValueINPRIC6
    If priRM6 Then rmDict("PRIMARY DEADEND") = PreviousValueRMPRIC6
    
    If priRP7 Then rpDict("SCOR") = PreviousValueRPPRIC7
    If priIN7 Then inDict("SCOR") = PreviousValueINPRIC7
    If priRM7 Then rmDict("SCOR") = PreviousValueRMPRIC7
    
    Call extractFromRPINRMDicts(replacePoleSection, rpDict, inDict, rmDict)
    
    If snRP3 Then rpDict("WR") = PreviousValueRPSNC3
    If snIN3 Then inDict("WR") = PreviousValueINSNC3
    If snRM3 Then rmDict("WR") = PreviousValueRMSNC3
    
    If snRP4 Then rpDict("NEUTRAL DEADEND") = PreviousValueRPSNC4
    If snIN4 Then inDict("NEUTRAL DEADEND") = PreviousValueINSNC4
    If snRM4 Then rmDict("NEUTRAL DEADEND") = PreviousValueRMSNC4
    
    If snRP5 Then rpDict("SECONDARY DEADEND") = PreviousValueRPSNC5
    If snIN5 Then inDict("SECONDARY DEADEND") = PreviousValueINSNC5
    If snRM5 Then rmDict("SECONDARY DEADEND") = PreviousValueRMSNC5
    
    If snRP6 Then rpDict("OPEN WIRE DEADEND") = PreviousValueRPSNC6
    If snIN6 Then inDict("OPEN WIRE DEADEND") = PreviousValueINSNC6
    If snRM6 Then rmDict("OPEN WIRE DEADEND") = PreviousValueRMSNC6
    
    If snRP7 Then rpDict("TANGENT CLAMP") = PreviousValueRPSNC7
    If snIN7 Then inDict("TANGENT CLAMP") = PreviousValueINSNC7
    If snRM7 Then rmDict("TANGENT CLAMP") = PreviousValueRMSNC7
    
    Call extractFromRPINRMDicts(replacePoleSection, rpDict, inDict, rmDict)
    
    Dim conductors As String
    
    If priBool Then conductors = "PRIMARY"
    If neutBool Then
        If conductors <> "" Then conductors = conductors & ","
        conductors = conductors & "NEUTRAL"
    End If
    If secBool Then
        If conductors <> "" Then conductors = conductors & ","
        conductors = conductors & "SECONDARY"
    End If
    If owBool Then
        If conductors <> "" Then conductors = conductors & ","
        conductors = conductors & "OPEN WIRE"
    End If
    If ohsCount > 0 Then
        If conductors <> "" Then conductors = conductors & ","
        conductors = conductors & "SVC"
        Dim serviceSizes As Scripting.Dictionary: Set serviceSizes = New Scripting.Dictionary
        For Each Service In pole.services
            For Each midspan In Service.midspans
                If Not serviceSizes.exists(Service.size) Then serviceSizes(Service.size) = 0
                serviceSizes(Service.size) = serviceSizes(Service.size) + 1
            Next midspan
        Next Service
        
        For Each size In serviceSizes
            If serviceSizes(size) > 1 Then
                replacePoleSection = replacePoleSection & "(" & serviceSizes(size) & ")" & Replace(size, " ", "") & " SVC DEADEND" & vbLf
            Else
                replacePoleSection = replacePoleSection & Replace(size, " ", "") & " SVC DEADEND" & vbLf
            End If
        Next size
        replacePoleSection = replacePoleSection & secondaryFigures("SECONDARY DEADEND")
    End If
    
    If conductors <> "" Then transferNotes = transferNotes & conductors & vbLf
    If priBool = False Then transferNotes = transferNotes & "@11"" FROM TOP OF POLE" & vbLf
    
    i = 0
    For Each key In comms
        If i >= CInt(CETV.value) Then Exit For
        For j = 1 To 8
            If UCase(Trim(pds.Range("COMM" & j).value)) = UCase(Trim(key)) Then
                If pds.Range("COMM" & j).offSet(2, 0).offSet(0, 1).value <> "" Then
                    transferNotes = transferNotes & key & " @" & pds.Range("COMM" & j).offSet(2, 0).offSet(0, 1).value & vbLf
                Else
                    transferNotes = transferNotes & key & " @" & pds.Range("COMM" & j).offSet(2, 0).value & vbLf
                End If
                commTransfers.Add key
                Exit For
            End If
        Next j
        i = i + 1
    Next key
    
    If TRSL.value Then
        transferNotes = transferNotes & "STREETLIGHT BOTTOM BRACKET @" & STLTBTMBRKT.value & vbLf
        transferNotes = transferNotes & "FIGURE 43-116-1" & vbLf
    End If
    If TRSPG.value Then
        If spg11KCount > 0 Or spg20KCount > 0 Then
            If spg11KCount > 0 Then
                transferNotes = transferNotes & IIf(spg11KCount > 1, "(" & spg11KCount & ")", "") & "11K SPAN GUY" & vbLf
            End If
            If spg20KCount > 0 Then
                transferNotes = transferNotes & IIf(spg20KCount > 1, "(" & spg20KCount & ")", "") & "20K SPAN GUY" & vbLf
            End If
            transferNotes = transferNotes & "FIGURE 22-420-1" & vbLf
        End If
    End If
    If TRT.value Then
        transferNotes = transferNotes & UCase(TRSIZE.value) & " TRANSFORMER"
        If RPPSIZE.value >= 50 Then transferNotes = transferNotes & " [@35'0""]"
        transferNotes = transferNotes & vbLf
        If COLA.ListIndex = 0 Or COLA.value = "" Then
            transferNotes = transferNotes & "[INSERT CO/LA INFO]" & vbLf
        ElseIf COLA.ListIndex = 1 Then
            transferNotes = transferNotes & "CO/LA ON SA TO LCOM" & vbLf
        ElseIf COLA.ListIndex = 2 Then
            transferNotes = transferNotes & "CO/LA ON LCOM" & vbLf
        ElseIf COLA.ListIndex = 3 Then
            transferNotes = transferNotes & "CO/LA ON SA" & vbLf
        ElseIf COLA.ListIndex = 4 Then
            transferNotes = transferNotes & "CO ON LCOM" & vbLf
            transferNotes = transferNotes & "LA ON TRANSFORMER" & vbLf
        ElseIf COLA.ListIndex = 5 Then
            transferNotes = transferNotes & "CO ON SA" & vbLf
            transferNotes = transferNotes & "LA ON TRANSFORMER" & vbLf
        ElseIf COLA.ListIndex = 6 Then
            transferNotes = transferNotes & "CO ON SA TO LCOM" & vbLf
            transferNotes = transferNotes & "LA ON TRANSFORMER" & vbLf
        End If
        If TRLETTER.value = "3 PHASE" Then
            transferNotes = transferNotes & "FIGURE 26-301-1" & vbLf
        Else
            transferNotes = transferNotes & "FIGURE 26-101-1 DETAIL " & IIf(Trim(TRLETTER.text) = "", "[A/B/C/D]", Left(TRLETTER.value, 1)) & vbLf
        End If
    End If
    
    transferNotes = transferNotes & "CE ID TAG" & vbLf
    
    replaceNotes = replacePoleSection & replaceNotes
End Sub

Private Sub extractFromRPINRMDicts(ByRef replacePoleSection As String, rpDict As Scripting.Dictionary, inDict As Scripting.Dictionary, rmDict As Scripting.Dictionary)
    For Each key In inDict
        If rmDict.exists(key) Then
            If inDict(key) >= rmDict(key) Then
                If Not rpDict.exists(key) Then rpDict(key) = 0
                rpDict(key) = rpDict(key) + rmDict(key)
                inDict(key) = inDict(key) - rmDict(key)
                rmDict(key) = rmDict(key) - rmDict(key)
            Else
                If Not rpDict.exists(key) Then rpDict(key) = 0
                rpDict(key) = rpDict(key) + inDict(key)
                rmDict(key) = rmDict(key) - inDict(key)
                inDict(key) = inDict(key) - inDict(key)
            End If
            If inDict(key) = 0 Then Call inDict.Remove(key)
            If rmDict(key) = 0 Then Call rmDict.Remove(key)
        End If
    Next key
    
    For Each key In rpDict
        If rmDict.exists(key) And inDict.count = 0 Then
            replacePoleSection = replacePoleSection & "(" & rpDict(key) + rmDict(key) & ")" & key & "/" & IIf(rpDict(key) <> 1, "(" & rpDict(key) & ")", "") & key & vbLf
            replacePoleSection = replacePoleSection & secondaryFigures(key)
            Call rmDict.Remove(key)
        ElseIf inDict.exists(key) And rmDict.count = 0 Then
            eplacePoleSection = replacePoleSection & IIf(rpDict(key) <> 1, "(" & rpDict(key) & ")", "") & key & "/(" & rpDict(key) + inDict(key) & ")" & key & vbLf
            replacePoleSection = replacePoleSection & secondaryFigures(key)
            Call inDict.Remove(key)
        Else
            replacePoleSection = replacePoleSection & IIf(rpDict(key) <> 1, "(" & rpDict(key) & ")", "") & key & vbLf
            replacePoleSection = replacePoleSection & secondaryFigures(key)
        End If
        
    Next key
    
    If inDict.count = 1 And rmDict.count = 1 Then
        For Each key In inDict
            For Each Key2 In rmDict
                replacePoleSection = replacePoleSection & IIf(rmDict(Key2) <> 1, "(" & rmDict(Key2) & ")", "") & Key2 & "/"
                replacePoleSection = replacePoleSection & IIf(inDict(key) <> 1, "(" & inDict(key) & ")", "") & key & vbLf
                replacePoleSection = replacePoleSection & secondaryFigures(key)
                Call inDict.Remove(key)
                Call rmDict.Remove(Key2)
            Next Key2
        Next key
    End If
    
    For Each key In inDict
        installNotes = installNotes & IIf(inDict(key) <> 1, "(" & inDict(key) & ")", "") & key & vbLf
        installNotes = installNotes & secondaryFigures(key)
    Next key
    
    For Each key In rmDict
        removeNotes = removeNotes & IIf(rmDict(key) <> 1, "(" & rmDict(key) & ")", "") & key & vbLf
        removeNotes = removeNotes & secondaryFigures(key)
    Next key
    
    Call rpDict.RemoveAll: Set rpDict = New Scripting.Dictionary
    Call inDict.RemoveAll: Set inDict = New Scripting.Dictionary
    Call rmDict.RemoveAll: Set rmDict = New Scripting.Dictionary
End Sub

Private Function secondaryFigures(hardware As Variant) As String
    Dim figure As String
    
    hardware = CStr(hardware)
    
    If hardware = "WR" Then
        If Not figuresUsed.exists("[FIGURE 23-301-2 or FIGURE 23-303-1 DETAIL A (FOR OPEN WIRE)]") Then
            figuresUsed.Add "[FIGURE 23-301-2 or FIGURE 23-303-1 DETAIL A (FOR OPEN WIRE)]", Nothing
            figure = "[FIGURE 23-301-2 or FIGURE 23-303-1 DETAIL A (FOR OPEN WIRE)]" & vbLf
        End If
    ElseIf hardware = "SECONDARY DEADEND" Then
        If Not figuresUsed.exists("FIGURE 23-302-1 DETAIL [A/B]") Then
            figuresUsed.Add "FIGURE 23-302-1 DETAIL [A/B]", Nothing
            figure = "FIGURE 23-302-1 DETAIL [A/B]" & vbLf
        End If
    ElseIf hardware = "OPEN WIRE DEADEND" Then
        If Not figuresUsed.exists("FIGURE 23-303-1 DETAIL B") Then
            figuresUsed.Add "FIGURE 23-303-1 DETAIL B", Nothing
            figure = "FIGURE 23-303-1 DETAIL B" & vbLf
        End If
    ElseIf hardware = "TANGENT CLAMP" Then
        If Not figuresUsed.exists("FIGURE 23-301-1") Then
            figuresUsed.Add "FIGURE 23-301-1", Nothing
            figure = "FIGURE 23-301-1" & vbLf
        End If
    ElseIf hardware = "SECONDARY AWAC DEADEND" Then
        If Not figuresUsed.exists("FIGURE 23-302-1 DETAIL B") Then
            figuresUsed.Add "FIGURE 23-302-1 DETAIL B", Nothing
            figure = "FIGURE 23-302-1 DETAIL B" & vbLf
        End If
    End If
    
    secondaryFigures = figure
End Function

Private Sub generateReconductorSection()
    If OWSPAN1.value <> "" And OWSPAN1.visible Then
        direction = getDirection(OWSPAN1)
        replaceNotes = replaceNotes & OWLENGTH1.value & " " & OWSIZE1.value & " OPEN WIRE / " & OWNEWSIZE1.value & " SECONDARY " & direction & vbLf
    End If
    
    If OWSPAN2.value <> "" And OWSPAN2.visible Then
        direction = getDirection(OWSPAN2)
        replaceNotes = replaceNotes & OWLENGTH2.value & " " & OWSIZE2.value & " OPEN WIRE / " & OWNEWSIZE2.value & " SECONDARY " & direction & vbLf
    End If
    
    If OWSPAN3.value <> "" And OWSPAN3.visible Then
        direction = getDirection(OWSPAN3)
        replaceNotes = replaceNotes & OWLENGTH3.value & " " & OWSIZE3.value & " OPEN WIRE / " & OWNEWSIZE3.value & " SECONDARY " & direction & vbLf
    End If
    
    If OWSPAN4.value <> "" And OWSPAN4.visible Then
        direction = getDirection(OWSPAN4)
        replaceNotes = replaceNotes & OWLENGTH4.value & " " & OWSIZE4.value & " OPEN WIRE / " & OWNEWSIZE4.value & " SECONDARY " & direction & vbLf
    End If
    
    If OWSPAN5.value <> "" And OWSPAN5.visible Then
        direction = getDirection(OWSPAN5)
        replaceNotes = replaceNotes & OWLENGTH5.value & " " & OWSIZE5.value & " OPEN WIRE / " & OWNEWSIZE5.value & " SECONDARY " & direction & vbLf
    End If
    
    If REC1.value Then
        replaceNotes = replaceNotes & "(" & 3 * RECA1.value & ")WR/" & IIf(RECA1.value > 1, "(" & RECA1.value & ")", "")
        If RECO1.value Then
            replaceNotes = replaceNotes & "TANGENT CLAMP" & vbLf
            replaceNotes = replaceNotes & secondaryFigures("TANGENT CLAMP")
        Else
            replaceNotes = replaceNotes & "WR" & vbLf
            replaceNotes = replaceNotes & secondaryFigures("WR")
        End If
    End If
    
    If REC2.value Then
        replaceNotes = replaceNotes & "(" & 3 * RECA2.value & ")OPEN WIRE DEADEND/" & IIf(RECA2.value > 1, "(" & RECA2.value & ")", "")
        If RECO2.value Then
            replaceNotes = replaceNotes & "SECONDARY AWAC DEADEND" & vbLf
            replaceNotes = replaceNotes & secondaryFigures("SECONDARY AWAC DEADEND")
        Else
            replaceNotes = replaceNotes & "SECONDARY DEADEND" & vbLf
            replaceNotes = replaceNotes & secondaryFigures("SECONDARY DEADEND")
        End If
    End If
    
    If REC3.value Then
        replaceNotes = replaceNotes & "(" & 3 * RECA3.value & ")WR/" & "(" & 3 * RECA3.value & ")OPEN WIRE DEADEND" & IIf(RECA3.value > 1, "(" & RECA3.value & ")", "")
        If RECO3.value Then
            replaceNotes = replaceNotes & "+ SECONDARY AWAC DEADEND" & vbLf
            replaceNotes = replaceNotes & secondaryFigures("SECONDARY AWAC DEADEND")
        Else
            replaceNotes = replaceNotes & "+ SECONDARY DEADEND" & vbLf
            replaceNotes = replaceNotes & "FIGURE 23-302-1 DETAIL A" & vbLf
        End If
    End If
    
    If REC4.value Then
        replaceNotes = replaceNotes & "(" & 6 * RECA4.value & ")OPEN WIRE DEADEND/" & IIf(RECA4.value > 1, "(" & RECA4.value & ")", "")
        If RECO4.value Then
            replaceNotes = replaceNotes & "TANGENT CLAMP" & vbLf
            replaceNotes = replaceNotes & secondaryFigures("TANGENT CLAMP")
        Else
            replaceNotes = replaceNotes & "WR" & vbLf
            replaceNotes = replaceNotes & secondaryFigures("WR")
        End If
    End If
    
    If REC1.value Or REC2.value Or REC3.value Or REC4.value Or (OWSPAN1.value <> "" And OWSPAN1.visible) Or (OWSPAN2.value <> "" And OWSPAN2.visible) Or (OWSPAN3.value <> "" And OWSPAN3.visible) Or (OWSPAN4.value <> "" And OWSPAN4.visible) Or (OWSPAN5.value <> "" And OWSPAN5.visible) Then
        replaceNotes = replaceNotes & "TO REPLACE OPEN WIRE" & vbLf
    End If
    
End Sub

Private Function getDirection(Span As String) As String

    StartPos = InStr(Span, "(")
    EndPos = InStr(StartPos + 1, Span, ")")
    
    If StartPos > 0 And EndPos > 0 Then
        degrees = CInt(Mid(Span, StartPos + 1, EndPos - StartPos - 1))
    Else
        degrees = 0
    End If
    
    degrees = degrees Mod 360
    If degrees < 0 Then degrees = degrees + 360
    
    
    Select Case degrees
        Case 337.5 To 360, 0 To 22.5
            getDirection = "NORTH"
        Case 22.5 To 67.5
            getDirection = "NORTHEAST"
        Case 67.5 To 112.5
            getDirection = "EAST"
        Case 112.5 To 157.5
            getDirection = "SOUTHEAST"
        Case 157.5 To 202.5
            getDirection = "SOUTH"
        Case 202.5 To 247.5
            getDirection = "SOUTHWEST"
        Case 247.5 To 292.5
            getDirection = "WEST"
        Case 292.5 To 337.5
            getDirection = "NORHTWEST"
        Case Else
            getDirection = "UNKNOWN"
    End Select
End Function

Private Sub generateBasicSection()
    If BSL.value Then
        installNotes = installNotes & "BOND STREETLIGHT BRACKET TO NEUTRAL " & Replace(BSLRFW.value, "OTHER", "[REASON FOR WORK]") & vbLf
        installNotes = installNotes & "FIGURE 42-105-1" & vbLf
    End If
    
    If SLM3.value Then
        replaceNotes = replaceNotes & "STREETLIGHT MOLDING" & vbLf
    End If
    
    If SLM4.value Then
        installNotes = installNotes & "STREETLIGHT MOLDING" & vbLf
    End If
    
    If TSDL.value Then
        transferNotes = transferNotes & "TRIM SECONDARY DRIPLOOP TO 4"" BELOW SECONDARY ATTACHMENT HEIGHT " & Replace(TSDLRFW.value, "OTHER", "[REASON FOR WORK]") & vbLf
    End If
    
    If RSL.value Then
        transferNotes = transferNotes & "RAISE STREETLIGHT BRACKET " & RSLD.value & " " & Replace(RSLRFW.value, "OTHER", "[REASON FOR WORK]") & vbLf
        transferNotes = transferNotes & "FIGURE 43-116-1" & vbLf
    End If
    
    If SLDL.value Then
        transferNotes = transferNotes & "TRIM STREETLIGHT DRIP LOOP TO 1"" BELOW STREETLIGHT BRACKET" & vbLf
        transferNotes = transferNotes & Replace(SLDLRFW.value, "OTHER", "[REASON FOR WORK]") & vbLf
    End If
End Sub

Private Sub generateNotesSection()
    For i = 1 To commTransfers.count
        If i = 1 Then notes = notes & "NOTE: "
        If i <> commTransfers.count Then
            notes = notes & UCase(commTransfers(i)) & ", "
        ElseIf i = commTransfers.count And i <> 1 Then
            notes = notes & "AND " & UCase(commTransfers(i)) & " HAVE A TRANSFER AGREEMENT" & vbLf
        Else
            notes = notes & UCase(commTransfers(i)) & " HAS A TRANSFER AGREEMENT" & vbLf
        End If
    Next i

    If RPPNOTE1.value Then
        notes = notes & "NOTE: TOP POLE ABOVE COMMS, TOPPED POLE TO BE REMOVED AT A LATER DATE" & vbLf
    End If
    If RPPNOTE2.value Then
        notes = notes & "NOTE: OUTAGE REQUIRED, SCHEDULING TO NOTIFY" & vbLf
    End If
    If RPPNOTE3.value Then
        notes = notes & "NOTE: TREE WORK REQUIRED" & vbLf
    End If
    If RPPNOTE4.value Then
        notes = notes & "NOTE: BRUSH WORK REQUIRED" & vbLf
    End If
    If RPPNOTE5.value Then
        notes = notes & "NOTE: POLE LOCATED IN ACTIVE FARMFIELD" & vbLf
    End If
    If RPPNOTE6.value Then
        notes = notes & "NOTE: FENCE REMOVAL REQUIRED" & vbLf
    End If
    If RPPNOTE7.value Then
        notes = notes & "NOTE: BACKYARD MACHINE REQUIRED" & vbLf
    End If
    If RPPNOTE8.value Then
        notes = notes & "NOTE: [INSERT POLE ACCESS NOTES]" & vbLf
    End If
End Sub

Private Sub generateDownguysSection()
    Dim insulators As String: insulators = "[XP-FG]"
    If pole.primaries.count = 0 Then insulators = "[XP]"
    For i = 1 To 3
        If Me.Controls("IA" & i & "1").value <> "" And Me.Controls("IA" & i & "1").visible Then
            If Me.Controls("IA" & i & "1").value = "RS" Then
                installNotes = installNotes & "11K-" & insulators & "-" & IIf(Me.Controls("IA" & i & "3").value = "", [lead], Me.Controls("IA" & i & "3").value) & "-RS " & IIf(Me.Controls("IA" & i & "4").value = "", "[DIRECTION]", Me.Controls("IA" & i & "4").value) & vbLf
            ElseIf Me.Controls("IA" & i & "1").value = "RT" Then
                If Me.Controls("IA" & i & "2").value = "20K" Then
                    installNotes = installNotes & "20K-" & insulators & "-" & IIf(Me.Controls("IA" & i & "3").value = "", [lead], Me.Controls("IA" & i & "3").value) & "-RT " & IIf(Me.Controls("IA" & i & "4").value = "", "[DIRECTION]", Me.Controls("IA" & i & "4").value) & vbLf
                ElseIf Me.Controls("IA" & i & "2").value = "(2)11K" Then
                    installNotes = installNotes & "11K-" & insulators & "-" & IIf(Me.Controls("IA" & i & "3").value = "", [lead], Me.Controls("IA" & i & "3").value) & "-RT " & IIf(Me.Controls("IA" & i & "4").value = "", "[DIRECTION]", Me.Controls("IA" & i & "4").value) & vbLf
                    installNotes = installNotes & "11K-" & insulators & vbLf
                End If
            ElseIf Me.Controls("IA" & i & "1").value = "STE" Then
                If Me.Controls("IA" & i & "2").value = "20K + 11K" Then
                    installNotes = installNotes & "20K-" & insulators & "-" & IIf(Me.Controls("IA" & i & "3").value = "", [lead], Me.Controls("IA" & i & "3").value) & "-STE " & IIf(Me.Controls("IA" & i & "4").value = "", "[DIRECTION]", Me.Controls("IA" & i & "4").value) & vbLf
                    installNotes = installNotes & "11K-" & insulators & vbLf
                ElseIf Me.Controls("IA" & i & "2").value = "(3)11K" Then
                    installNotes = installNotes & "11K-" & insulators & "-" & IIf(Me.Controls("IA" & i & "3").value = "", [lead], Me.Controls("IA" & i & "3").value) & "-STE " & IIf(Me.Controls("IA" & i & "4").value = "", "[DIRECTION]", Me.Controls("IA" & i & "4").value) & vbLf
                    installNotes = installNotes & "11K-" & insulators & vbLf
                    installNotes = installNotes & "11K-" & insulators & vbLf
                ElseIf Me.Controls("IA" & i & "2").value = "(2)20K" Then
                    installNotes = installNotes & "20K-" & insulators & "-" & IIf(Me.Controls("IA" & i & "3").value = "", [lead], Me.Controls("IA" & i & "3").value) & "-STE " & IIf(Me.Controls("IA" & i & "4").value = "", "[DIRECTION]", Me.Controls("IA" & i & "4").value) & vbLf
                    installNotes = installNotes & "20K-" & insulators & vbLf
                End If
            End If
            
            If Me.Controls("IA" & i & "5").value Then
                installNotes = installNotes & "SIDEWALK BRACE" & vbLf
                installNotes = installNotes & "FIGURE 22-450-1" & vbLf
            End If

            If Me.Controls("IGRFW" & i).value = "" Or Me.Controls("IGRFW" & i).value = "OTHER" Then
                installNotes = installNotes & "[REASON FOR WORK]" & vbLf
            Else
                installNotes = installNotes & Me.Controls("IGRFW" & i) & vbLf
            End If
            
            If Me.Controls("IGFIG" & i).value = "" Or Me.Controls("IGFIG" & i).value = "OTHER" Then
                installNotes = installNotes & "FIGURE [INSERT DOWNGUY FIGURE]" & vbLf
            Else
                installNotes = installNotes & Split(Me.Controls("IGFIG" & i).value, ",")(0) & vbLf
            End If
        End If
    Next i
    
    For i = 1 To 4
        Dim eLine1, eLine2, eLine3, rLine1, rLine2, rLine3 As String
        eLine1 = ""
        eLine2 = ""
        eLine3 = ""
        rLine1 = ""
        rLine2 = ""
        rLine3 = ""
        If Me.Controls("EA" & i & "1").value <> Me.Controls("RA" & i & "1").value Or Me.Controls("EA" & i & "2").value <> Me.Controls("RA" & i & "2").value Or Me.Controls("EA" & i & "3").value <> Me.Controls("RA" & i & "3").value Then
            If Me.Controls("EA" & i & "1").value <> "" Then
                If Me.Controls("EA" & i & "1").value = "RS" Then
                    eLine1 = "11K-" & insulators & "-" & IIf(Me.Controls("EA" & i & "3").value = "", [lead], Me.Controls("EA" & i & "3").value) & "-RS"
                    replace11kCount = replace11kCount + 1
                ElseIf Me.Controls("EA" & i & "1").value = "RT" Then
                    If Me.Controls("EA" & i & "2").value = "20K" Then
                        eLine1 = "20K-" & insulators & "-" & IIf(Me.Controls("EA" & i & "3").value = "", [lead], Me.Controls("EA" & i & "3").value) & "-RT"
                        replace20kCount = replace20kCount + 1
                    ElseIf Me.Controls("EA" & i & "2").value = "(2)11K" Then
                        eLine1 = "11K-" & insulators & "-" & IIf(Me.Controls("EA" & i & "3").value = "", [lead], Me.Controls("EA" & i & "3").value) & "-RT"
                        eLine2 = "11K-" & insulators
                        replace11kCount = replace11kCount + 2
                    End If
                ElseIf Me.Controls("EA" & i & "1").value = "STE" Then
                    If Me.Controls("EA" & i & "2").value = "20K + 11K" Then
                        eLine1 = "20K-" & insulators & "-" & IIf(Me.Controls("EA" & i & "3").value = "", [lead], Me.Controls("EA" & i & "3").value) & "-STE"
                        eLine2 = "11K-" & insulators
                        replace20kCount = replace20kCount + 1
                        replace11kCount = replace11kCount + 1
                    ElseIf Me.Controls("EA" & i & "2").value = "(3)11K" Then
                        eLine1 = "11K-" & insulators & "-" & IIf(Me.Controls("EA" & i & "3").value = "", [lead], Me.Controls("EA" & i & "3").value) & "-STE"
                        eLine2 = "11K-" & insulators
                        eLine3 = "11K-" & insulators
                        replace11kCount = replace11kCount + 3
                    ElseIf Me.Controls("EA" & i & "2").value = "(2)20K" Then
                        eLine1 = "20K-" & insulators & "-" & IIf(Me.Controls("EA" & i & "3").value = "", [lead], Me.Controls("EA" & i & "3").value) & "-STE"
                        eLine2 = "20K-" & insulators
                        replace20kCount = replace20kCount + 2
                    End If
                End If
            End If
            If Me.Controls("RA" & i & "1").value <> "" Then
                If Me.Controls("RA" & i & "1").value = "RS" Then
                    rLine1 = "11K-" & insulators & "-" & IIf(Me.Controls("RA" & i & "3").value = "", [lead], Me.Controls("RA" & i & "3").value) & "-RS"
                ElseIf Me.Controls("RA" & i & "1").value = "RT" Then
                    If Me.Controls("RA" & i & "2").value = "20K" Then
                        rLine1 = "20K-" & insulators & "-" & IIf(Me.Controls("RA" & i & "3").value = "", [lead], Me.Controls("RA" & i & "3").value) & "-RT"
                    ElseIf Me.Controls("RA" & i & "2").value = "(2)11K" Then
                        rLine1 = "11K-" & insulators & "-" & IIf(Me.Controls("RA" & i & "3").value = "", [lead], Me.Controls("RA" & i & "3").value) & "-RT"
                        rLine2 = "11K-" & insulators
                    End If
                ElseIf Me.Controls("RA" & i & "1").value = "STE" Then
                    If Me.Controls("RA" & i & "2").value = "20K + 11K" Then
                        rLine1 = "20K-" & insulators & "-" & IIf(Me.Controls("RA" & i & "3").value = "", [lead], Me.Controls("RA" & i & "3").value) & "-STE"
                        rLine2 = "11K-" & insulators
                    ElseIf Me.Controls("RA" & i & "2").value = "(3)11K" Then
                        rLine1 = "11K-" & insulators & "-" & IIf(Me.Controls("RA" & i & "3").value = "", [lead], Me.Controls("RA" & i & "3").value) & "-STE"
                        rLine2 = "11K-" & insulators
                        rLine3 = "11K-" & insulators
                    ElseIf Me.Controls("RA" & i & "2").value = "(2)20K" Then
                        rLine1 = "20K-" & insulators & "-" & IIf(Me.Controls("RA" & i & "3").value = "", [lead], Me.Controls("RA" & i & "3").value) & "-STE"
                        rLine2 = "20K-" & insulators
                    End If
                End If
            End If
            replaceNotes = replaceNotes & eLine1 & " / " & rLine1 & " " & IIf(Me.Controls("EA" & i & "4").value = "", "[DIRECTION]", Me.Controls("EA" & i & "4").value) & vbLf
            If eLine2 <> "" Or rLine2 <> "" Then
                replaceNotes = replaceNotes & eLine2 & Space((Len(eLine1) * 2) - Len(eLine2) - 4) & rLine2 & vbLf
            End If
            If eLine3 <> "" Or rLine3 <> "" Then
                replaceNotes = replaceNotes & eLine3 & Space((Len(eLine1) * 2) - Len(eLine3) - 4) & rLine3 & vbLf
            End If
            
            If Me.Controls("DGRFW" & i).value = "" Or Me.Controls("DGRFW" & i).value = "OTHER" Then
                replaceNotes = replaceNotes & "[REASON FOR WORK]" & vbLf
            Else
                replaceNotes = replaceNotes & Me.Controls("DGRFW" & i) & vbLf
            End If
            
            If Me.Controls("DGFIG" & i).value = "" Or Me.Controls("DGFIG" & i).value = "OTHER" Then
                replaceNotes = replaceNotes & "FIGURE [INSERT DOWNGUY FIGURE]" & vbLf
            Else
                replaceNotes = replaceNotes & Split(Me.Controls("DGFIG" & i).value, ",")(0) & vbLf
            End If
        End If
    Next i
    
End Sub

Private Sub compileCrewNotes()
    If replaceNotes <> "" Then
        crewNotes = crewNotes & "REPLACE" & vbLf & replaceNotes & vbLf
    End If
    If installNotes <> "" Then
        crewNotes = crewNotes & "INSTALL" & vbLf & installNotes & vbLf
    End If
    If transferNotes <> "" Then
        crewNotes = crewNotes & "TRANSFER" & vbLf & transferNotes & vbLf
    End If
    If removeNotes <> "" Then
        crewNotes = crewNotes & "REMOVE" & vbLf & removeNotes & vbLf
    End If
    If notes <> "" Then
        crewNotes = crewNotes & notes
    End If
End Sub

Private Sub Initialize_ReplacePole()
    RPPSIZE.list = Array("35", "40", "45", "50", "55", "60", "65", "70")
    RPPCLASS.list = Array("1", "2", "3", "4")
    If InStr(pds.Range("HEIGHT"), "(") > 0 Then
        If RPPSIZE.value = "" Then
            tempHeight = Left(Trim(pds.Range("HEIGHT")), InStr(Trim(pds.Range("HEIGHT")), "(") - 1)
            If IsNumeric(tempHeight) And tempHeight < "35" Then tempHeight = "35"
            RPPSIZE.value = tempHeight
        ElseIf CInt(Left(Trim(pds.Range("HEIGHT")), InStr(Trim(pds.Range("HEIGHT")), "(") - 1)) > 40 Then
            RPPSIZE.value = Left(Trim(pds.Range("HEIGHT")), InStr(Trim(pds.Range("HEIGHT")), "(") - 1)
        End If
    Else
        If RPPSIZE.value = "" Then
            RPPSIZE.value = Trim(pds.Range("HEIGHT"))
        ElseIf CInt(Trim(pds.Range("HEIGHT"))) > 40 Then
            RPPSIZE.value = Trim(pds.Range("HEIGHT"))
        End If
    End If
    If pds.Range("SECONLY").value = True Then
        RPPCLASS.value = 4
        DSV.value = "N/A"
        PSV.value = "N/A"
    ElseIf InStr(pds.Range("DESC"), "3") > 0 Or InStr(pds.Range("DESC").offSet(1, 0), "3") > 0 Or InStr(pds.Range("DESC").offSet(2, 0), "3") > 0 Or InStr(pds.Range("DESC").offSet(3, 0), "3") > 0 Then
        RPPCLASS.value = 2
    Else
        RPPCLASS.value = 3
    End If
    
    RPPRFW.list = Array("OTHER", "TO CORRECT POLE LOADING FAILURE @[XX]% [AS-IS/WITH APPLICANT]", "TO CORRECT POLE DETERIORATION", "DUE TO FAILED HAMMER TEST", "TO ALLOW COMMS TO CORRECT VIOLATIONS", "TO MAKE ROOM FOR APPLICANT")
    RPPRFW.ListIndex = 0
    
    RPBSLRFW.list = Array("OTHER", "TO CORRECT 40"" STREETLIGHT SEPARATION VIOLATION", "TO MAKE ROOM FOR APPLICANT", "TO ALLOW COMMS TO CORRECT VIOLATIONS")
    RPBSLRFW.ListIndex = 1
    
    COLA.list = Array("OTHER", "CO/LA ON SA TO LCOM", "CO/LA ON LCOM", "CO/LA ON SA", "CO ON LCOM AND LA ON XFMR", "CO ON SA AND LA ON XFMR", "CO ON SA TO LCOM AND LA ON XFMR")
    COLA.ListIndex = 0
    
    CETV.list = Array(0, 1, 2, 3, 4, 5, 6, 7, 8)
    CETV.ListIndex = 0
    
    TRSIZE.list = Array("10kVA", "25kVA", "50kVA", "75kVA", "100kVA")
    TRRPSIZE.list = Array("10kVA", "25kVA", "50kVA", "75kVA", "100kVA")
    
    If pds.Range("EQUIPMENT") = "XFMR" Or pds.Range("EQUIPMENT") = "TRANSFORMER" Then
        TRT.value = True
        TRSIZE.value = Replace(pds.Range("EQUIPMENTSIZE"), " ", "")
        TRRPSIZE.value = Replace(pds.Range("EQUIPMENTSIZE"), " ", "")
    End If
    
    If Trim(pds.Range("STLTBRKT").value) <> "N/A" And Trim(pds.Range("STLTBRKT").value) <> "" Then
        TRSL.value = True
        STLTBTMBRKT.value = Trim(Split(pds.Range("STLTBRKT").value, vbLf)(0))
        If pds.Range("MBSM").offSet(0, 1).value = "YES" Then
            SLM2.value = True
        ElseIf pds.Range("MBSM").offSet(0, 1).value = "NO" Then
            SLM1.value = True
        End If
    End If
    
    If riserBool Then
        RPR.value = True
    End If
    
    If dg11KCount > 0 Or dg20KCount > 0 Then
        RPDG.value = True
    End If
    
    If spg11KCount > 0 Or spg20KCount > 0 Then
        TRSPG.value = True
    End If
    
    MHT1.list = Array("RP", "IN", "RM")
    MHT2.list = Array("RP", "IN", "RM")
    
    PRIT1.list = Array("RP", "IN", "RM")
    PRIT2.list = Array("RP", "IN", "RM")
    PRIT3.list = Array("RP", "IN", "RM")
    PRIT4.list = Array("RP", "IN", "RM")
    PRIT5.list = Array("RP", "IN", "RM")
    PRIT6.list = Array("RP", "IN", "RM")
    PRIT7.list = Array("RP", "IN", "RM")
    
    SNT1.list = Array("RP", "IN", "RM")
    SNT2.list = Array("RP", "IN", "RM")
    SNT3.list = Array("RP", "IN", "RM")
    SNT4.list = Array("RP", "IN", "RM")
    SNT5.list = Array("RP", "IN", "RM")
    SNT6.list = Array("RP", "IN", "RM")
    SNT7.list = Array("RP", "IN", "RM")
    
    FUT1.list = Array("RP", "IN", "RM")
    FUT2.list = Array("RP", "IN", "RM")
    FUT3.list = Array("RP", "IN", "RM")
    
    Dim countArray As Variant: countArray = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12)
    
    MHC1.list = countArray
    MHC2.list = countArray
    
    PRIC1.list = countArray
    PRIC2.list = countArray
    PRIC3.list = countArray
    PRIC4.list = countArray
    PRIC5.list = countArray
    PRIC6.list = countArray
    PRIC7.list = countArray
    
    SNC1.list = countArray
    SNC2.list = countArray
    SNC3.list = countArray
    SNC4.list = countArray
    SNC5.list = countArray
    SNC6.list = countArray
    SNC7.list = countArray
    
    FUC1.list = countArray
    FUC2.list = countArray
    FUC3.list = countArray
    
    Dim figArray(7) As String
    figArray(0) = "OTHER"
    figArray(1) = "FIGURE 23-102-1 DETAIL [A/B], Single-Phase—Vertical Pole Top Pin"
    figArray(2) = "FIGURE 23-103-1 DETAIL [A/B/C/D/E], Single-Phase—Vertical Pull-Off"
    figArray(3) = "FIGURE 23-104-1 DETAIL [A/B], Single-Phase Vertical Dead End"
    figArray(4) = "FIGURE 23-110-1, Crossarm Tangent or Corner Structure Single-Phase Wye or Delta and Open Wye"
    figArray(5) = "FIGURE 23-111-1, Primary Dead Ends on Arms—Single-Phase Delta or Wye"
    figArray(6) = "FIGURE 23-140-1 DETAIL [A/B], Crossarm Tangent or Corner Structure—Three-Phase Wye or Delta"
    figArray(7) = "FIGURE 23-150-1, Dead End Open Wye Updated Crossarms"
    
    FIG.list = figArray
    FIG.ListIndex = 0
    
    TRLETTER.list = Array("A, PTP", "B, VDE", "C, XARM", "D, FGARM", "3 PHASE")
End Sub

Private Sub Initialize_Basic()
    TSDLRFW.list = Array("TO CORRECT 40"" SAFETY ZONE VIOLATION", "TO MAKE ROOM FOR APPLICANT", "TO ALLOW COMMS TO CORRECT VIOLATIONS", "OTHER")
    TSDLRFW.ListIndex = 0
    BSLRFW.list = Array("TO CORRECT 40"" STREETLIGHT SEPARATION VIOLATION", "TO MAKE ROOM FOR APPLICANT", "TO ALLOW COMMS TO CORRECT VIOLATIONS", "OTHER")
    BSLRFW.ListIndex = 0
    RSLRFW.list = Array("OTHER", "TO MAKE ROOM FOR APPLICANT", "TO ALLOW COMMS TO CORRECT VIOLATIONS")
    RSLRFW.ListIndex = 0
    RSLD.list = Array("24""", "23""", "22""", "21""", "20""", "19""", "18""", "17""", "16""", "15""", "14""", "13""", "12""", "11""", "10""", "9""", "8""", "7""", "6""", "5""", "4""")
    RSLD.ListIndex = 0
    SLDLRFW.list = Array("TO CORRECT 12"" STREETLIGHT DRIP LOOP SEPARATION VIOLATION", "TO MAKE ROOM FOR APPLICANT", "TO ALLOW COMMS TO CORRECT VIOLATIONS", "OTHER")
    SLDLRFW.ListIndex = 0
End Sub

Private Sub Initialize_Reconductor()
    RSPANS.list = Array(0, 1, 2, 3, 4, 5)
    RSPANS.ListIndex = 1
    
    Dim spanAray() As Variant
    ReDim Preserve spanAray(0)
    spanAray(0) = ""
    Dim name As Variant
    For i = 1 To 12
        For Each name In pds.Names
            If name.name = "'" & pds.name & "'" & "!" & "TOPOLE" & i Then
                If Trim(Replace(pds.Range("TOPOLE" & i).value, "-", "")) <> "" Then
                    For j = 1 To 50
                        If pds.Range("UTTYPE").offSet(j, 0).Interior.color = 16777215 Then Exit For
                        If pds.Range("UTTYPE").offSet(j, 0) = "" Then Exit For
                        If pds.Range("UTTYPE").offSet(j, 0).text = "STLT. BOTTOM BRKT." Then Exit For
                        If InStr(pds.Range("UTTYPE").offSet(j, 0).text, "OW") > 0 Then
                            If Trim(Replace(pds.Range("UTMIDSPAN" & i).offSet(j, 0).value, "-", "")) <> "" Then
                                ReDim Preserve spanAray(UBound(spanAray) + 1)
                                spanAray(UBound(spanAray)) = Trim(pds.Range("TOPOLE" & i).value)
                                Exit For
                            End If
                        End If
                    Next j
                End If
            End If
        Next name
    Next i
    
    OWSPAN1.list = spanAray()
    OWSPAN1.ListIndex = 0
    OWSPAN2.list = spanAray()
    OWSPAN2.ListIndex = 0
    OWSPAN3.list = spanAray()
    OWSPAN3.ListIndex = 0
    OWSPAN4.list = spanAray()
    OWSPAN4.ListIndex = 0
    OWSPAN5.list = spanAray()
    OWSPAN5.ListIndex = 0
    
    Dim secSizes() As Variant
    ReDim secSizes(2)
    secSizes = Array("4TX", "1/0 TX", "3/0 TX")
    
    OWNEWSIZE1.list = secSizes
    OWNEWSIZE1.ListIndex = 2
    OWNEWSIZE2.list = secSizes
    OWNEWSIZE2.ListIndex = 2
    OWNEWSIZE3.list = secSizes
    OWNEWSIZE3.ListIndex = 2
    OWNEWSIZE4.list = secSizes
    OWNEWSIZE4.ListIndex = 2
    OWNEWSIZE5.list = secSizes
    OWNEWSIZE5.ListIndex = 2
    
    Dim countArray As Variant: countArray = Array(1, 2, 3, 4)
    RECA1.list = countArray
    RECA1.ListIndex = 0
    RECA2.list = countArray
    RECA2.ListIndex = 0
    RECA3.list = countArray
    RECA3.ListIndex = 0
    RECA4.list = countArray
    RECA4.ListIndex = 0
End Sub

Private Sub Initialize_Downguys()
    IAC.list = Array(0, 1, 2, 3)
    IAC.ListIndex = 0
    
    For i = 1 To 3
        Me.Controls("IA" & i & "1").list = Array("RS", "RT", "STE")
        Me.Controls("IA" & i & "4").list = Array("NORTH", "NORTHEAST", "EAST", "SOUTHEAST", "SOUTH", "SOUTHWEST", "WEST", "NORTHWEST")
    Next i
    
    anchorCount = 0
    For i = 0 To 10
        If InStr(CStr(pds.Range("ANCHOROWNER").offSet(i, 0).value), "Consumers Energy") > 0 Then
            anchorCount = anchorCount + 1
            Me.Controls("EA" & anchorCount & 3).value = pds.Range("ANCHOROWNER").offSet(i, 0).offSet(0, 1).value
            Me.Controls("RA" & anchorCount & 3).value = pds.Range("ANCHOROWNER").offSet(i, 0).offSet(0, 1).value
            Me.Controls("EA" & anchorCount & 4).value = getDirection(pds.Range("ANCHOROWNER").offSet(i, 0).offSet(0, 1).offSet(0, 1).value)
            For j = 1 To 4
                Me.Controls("DGFIG" & anchorCount).visible = True
                Me.Controls("DGRFW" & anchorCount).visible = True
                Me.Controls("EA" & anchorCount & j).visible = True
                If j <> 4 Then Me.Controls("RA" & anchorCount & j).visible = True
            Next j
        If anchorCount = 4 Then Exit For
        End If
    Next i
    
    If anchorCount > 0 Then
        ANCHORL.visible = True
        ANCHORBAR.visible = True
        EL.visible = True
        RL.visible = True
        EL1.visible = True
        EL2.visible = True
        EL3.visible = True
        EL4.visible = True
        RL1.visible = True
        RL2.visible = True
        RL3.visible = True
    End If
    
    For i = 1 To 4
        Me.Controls("EA" & i & "1").list = Array("RS", "RT", "STE")
        Me.Controls("EA" & i & "4").list = Array("NORTH", "NORTHEAST", "EAST", "SOUTHEAST", "SOUTH", "SOUTHWEST", "WEST", "NORTHWEST")
        Me.Controls("RA" & i & "1").list = Array("RS", "RT", "STE")
    Next i
    
    Dim dgFigArray(17) As String
    dgFigArray(0) = "OTHER"
    dgFigArray(1) = "FIGURE 22-405-1 DETAIL [A/B/C/D], 1 GUY SECONDARY"
    dgFigArray(2) = "FIGURE 22-101-1, 1 GUY PTP"
    dgFigArray(3) = "FIGURE 22-101-2, 1 GUY VDE"
    dgFigArray(4) = "FIGURE 22-101-3, 1 GUY FG VDE"
    dgFigArray(5) = "FIGURE 22-101-4, 2 GUY FG VDE"
    dgFigArray(6) = "FIGURE 22-101-5, 1 GUY 1Ø XARM ANGLE"
    dgFigArray(7) = "FIGURE 22-101-6, 1 GUY 1Ø XARM DEADEND"
    dgFigArray(8) = "FIGURE 22-101-7, 2 GUY FG VPO/VDE"
    dgFigArray(9) = "FIGURE 22-101-8, 3 GUY FG VPO/VDE"
    dgFigArray(10) = "FIGURE 22-101-9, 1 GUY 3Ø XARM ANGLE GAP"
    dgFigArray(11) = "FIGURE 22-101-10, 1 GUY 2Ø XARM ANGLE"
    dgFigArray(12) = "FIGURE 22-101-11 DETAIL [/A], 1 GUY 3Ø XARM ANGLE"
    dgFigArray(13) = "FIGURE 22-101-12, 2 GUY 3Ø FG XARM ANGLE"
    dgFigArray(14) = "FIGURE 22-101-13, 1 GUY 2Ø XARM DEADEND"
    dgFigArray(15) = "FIGURE 22-101-14, 1 GUY 3Ø FG XARM DEADEND"
    dgFigArray(16) = "FIGURE 22-101-15, 2 GUY 3Ø FG XARM DEADEND"
    dgFigArray(17) = "FIGURE 22-101-16, 3 GUY 3Ø (2)FG XARM DEADEND"
    
    IGFIG1.list = dgFigArray
    IGFIG1.ListIndex = 0
    IGFIG2.list = dgFigArray
    IGFIG2.ListIndex = 0
    IGFIG3.list = dgFigArray
    IGFIG3.ListIndex = 0
    DGFIG1.list = dgFigArray
    DGFIG1.ListIndex = 0
    DGFIG2.list = dgFigArray
    DGFIG2.ListIndex = 0
    DGFIG3.list = dgFigArray
    DGFIG3.ListIndex = 0
    DGFIG4.list = dgFigArray
    DGFIG4.ListIndex = 0
    
    DGFV.list = dgFigArray
    DGFV.ListIndex = 0
    
    Dim dgRFWArray(3) As String
    dgRFWArray(0) = "OTHER"
    dgRFWArray(1) = "TO CORRECT DG/A LOADING FAILURE @[XX]% [AS-IS/WITH APPLICANT]"
    dgRFWArray(2) = "TO MAKE ROOM FOR APPLICANT"
    dgRFWArray(3) = "TO BACK UNSUPPORTED SPAN"
    
    IGRFW1.list = dgRFWArray
    IGRFW1.ListIndex = 0
    IGRFW2.list = dgRFWArray
    IGRFW2.ListIndex = 0
    IGRFW3.list = dgRFWArray
    IGRFW3.ListIndex = 0
    DGRFW1.list = dgRFWArray
    DGRFW1.ListIndex = 0
    DGRFW2.list = dgRFWArray
    DGRFW2.ListIndex = 0
    DGRFW3.list = dgRFWArray
    DGRFW3.ListIndex = 0
    DGRFW4.list = dgRFWArray
    DGRFW4.ListIndex = 0
End Sub

Private Sub SNT1_Change()
    If SNT1.value = "RP" Then
        SN1.value = snRP1
        SNC1.value = PreviousValueRPSNC1
    ElseIf SNT1.value = "IN" Then
        SN1.value = snIN1
        SNC1.value = PreviousValueINSNC1
    ElseIf SNT1.value = "RM" Then
        SN1.value = snRM1
        SNC1.value = PreviousValueRMSNC1
    End If
End Sub

Private Sub SN1_Click()
    SNC1.visible = SN1.value
    
    If SNT1.value = "RP" Then
        If snRP1 <> SN1.value Then
            snRP1 = SN1.value
            If SN1.value Then
                UASN.value = CDbl(UASN.value) - (CInt(SNC1.value) * 2)
            Else
                UASN.value = CDbl(UASN.value) + (CInt(SNC1.value) * 2)
            End If
        End If
    ElseIf SNT1.value = "IN" Then
        If snIN1 <> SN1.value Then
            snIN1 = SN1.value
            If SN1.value Then
                UASN.value = CDbl(UASN.value) - (CInt(SNC1.value))
            Else
                UASN.value = CDbl(UASN.value) + (CInt(SNC1.value))
            End If
        End If
    ElseIf SNT1.value = "RM" Then
        If snRM1 <> SN1.value Then
            snRM1 = SN1.value
            If SN1.value Then
                UASN.value = CDbl(UASN.value) - (CInt(SNC1.value))
            Else
                UASN.value = CDbl(UASN.value) + (CInt(SNC1.value))
            End If
        End If
    End If
End Sub

Private Sub SNC1_Change()
    If SNT1.value = "RP" Then
        If PreviousValueRPSNC1 <> SNC1.value Then
            UASN.value = CDbl(UASN.value) - (2 * (SNC1.value - PreviousValueRPSNC1))
            PreviousValueRPSNC1 = SNC1.value
        End If
    ElseIf SNT1.value = "IN" Then
        If PreviousValueINSNC1 <> SNC1.value Then
            UASN.value = CDbl(UASN.value) - (SNC1.value - PreviousValueINSNC1)
            PreviousValueINSNC1 = SNC1.value
        End If
    ElseIf SNT1.value = "RM" Then
        If PreviousValueRMSNC1 <> SNC1.value Then
            UASN.value = CDbl(UASN.value) - (SNC1.value - PreviousValueRMSNC1)
            PreviousValueRMSNC1 = SNC1.value
        End If
    End If
End Sub

Private Sub SNT2_Change()
    If SNT2.value = "RP" Then
        SN2.value = snRP2
        SNC2.value = PreviousValueRPSNC2
    ElseIf SNT2.value = "IN" Then
        SN2.value = snIN2
        SNC2.value = PreviousValueINSNC2
    ElseIf SNT2.value = "RM" Then
        SN2.value = snRM2
        SNC2.value = PreviousValueRMSNC2
    End If
End Sub

Private Sub SN2_Click()
    SNC2.visible = SN2.value
    
    If SNT2.value = "RP" Then
        If snRP2 <> SN2.value Then
            snRP2 = SN2.value
            If SN2.value Then
                UASN.value = CDbl(UASN.value) - (CInt(SNC2.value) * 2)
            Else
                UASN.value = CDbl(UASN.value) + (CInt(SNC2.value) * 2)
            End If
        End If
    ElseIf SNT2.value = "IN" Then
        If snIN2 <> SN2.value Then
            snIN2 = SN2.value
            If SN2.value Then
                UASN.value = CDbl(UASN.value) - (CInt(SNC2.value))
            Else
                UASN.value = CDbl(UASN.value) + (CInt(SNC2.value))
            End If
        End If
    ElseIf SNT2.value = "RM" Then
        If snRM2 <> SN2.value Then
            snRM2 = SN2.value
            If SN2.value Then
                UASN.value = CDbl(UASN.value) - (CInt(SNC2.value))
            Else
                UASN.value = CDbl(UASN.value) + (CInt(SNC2.value))
            End If
        End If
    End If
End Sub

Private Sub SNC2_Change()
    If SNT2.value = "RP" Then
        If PreviousValueRPSNC2 <> SNC2.value Then
            UASN.value = CDbl(UASN.value) - (2 * (SNC2.value - PreviousValueRPSNC2))
            PreviousValueRPSNC2 = SNC2.value
        End If
    ElseIf SNT2.value = "IN" Then
        If PreviousValueINSNC2 <> SNC2.value Then
            UASN.value = CDbl(UASN.value) - (SNC2.value - PreviousValueINSNC2)
            PreviousValueINSNC2 = SNC2.value
        End If
    ElseIf SNT2.value = "RM" Then
        If PreviousValueRMSNC2 <> SNC2.value Then
            UASN.value = CDbl(UASN.value) - (SNC2.value - PreviousValueRMSNC2)
            PreviousValueRMSNC2 = SNC2.value
        End If
    End If
End Sub

Private Sub SNT3_Change()
    If SNT3.value = "RP" Then
        SN3.value = snRP3
        SNC3.value = PreviousValueRPSNC3
    ElseIf SNT3.value = "IN" Then
        SN3.value = snIN3
        SNC3.value = PreviousValueINSNC3
    ElseIf SNT3.value = "RM" Then
        SN3.value = snRM3
        SNC3.value = PreviousValueRMSNC3
    End If
End Sub

Private Sub SN3_Click()
    SNC3.visible = SN3.value
    
    If SNT3.value = "RP" Then
        If snRP3 <> SN3.value Then
            snRP3 = SN3.value
            If SN3.value Then
                UASN.value = CDbl(UASN.value) - (CInt(SNC3.value) * 2)
            Else
                UASN.value = CDbl(UASN.value) + (CInt(SNC3.value) * 2)
            End If
        End If
    ElseIf SNT3.value = "IN" Then
        If snIN3 <> SN3.value Then
            snIN3 = SN3.value
            If SN3.value Then
                UASN.value = CDbl(UASN.value) - (CInt(SNC3.value))
            Else
                UASN.value = CDbl(UASN.value) + (CInt(SNC3.value))
            End If
        End If
    ElseIf SNT3.value = "RM" Then
        If snRM3 <> SN3.value Then
            snRM3 = SN3.value
            If SN3.value Then
                UASN.value = CDbl(UASN.value) - (CInt(SNC3.value))
            Else
                UASN.value = CDbl(UASN.value) + (CInt(SNC3.value))
            End If
        End If
    End If
End Sub

Private Sub SNC3_Change()
    If SNT3.value = "RP" Then
        If PreviousValueRPSNC3 <> SNC3.value Then
            UASN.value = CDbl(UASN.value) - (2 * (SNC3.value - PreviousValueRPSNC3))
            PreviousValueRPSNC3 = SNC3.value
        End If
    ElseIf SNT3.value = "IN" Then
        If PreviousValueINSNC3 <> SNC3.value Then
            UASN.value = CDbl(UASN.value) - (SNC3.value - PreviousValueINSNC3)
            PreviousValueINSNC3 = SNC3.value
        End If
    ElseIf SNT3.value = "RM" Then
        If PreviousValueRMSNC3 <> SNC3.value Then
            UASN.value = CDbl(UASN.value) - (SNC3.value - PreviousValueRMSNC3)
            PreviousValueRMSNC3 = SNC3.value
        End If
    End If
End Sub

Private Sub SNT4_Change()
    If SNT4.value = "RP" Then
        SN4.value = snRP4
        SNC4.value = PreviousValueRPSNC4
    ElseIf SNT4.value = "IN" Then
        SN4.value = snIN4
        SNC4.value = PreviousValueINSNC4
    ElseIf SNT4.value = "RM" Then
        SN4.value = snRM4
        SNC4.value = PreviousValueRMSNC4
    End If
End Sub

Private Sub SN4_Click()
    SNC4.visible = SN4.value
    
    If SNT4.value = "RP" Then
        If snRP4 <> SN4.value Then
            snRP4 = SN4.value
            If SN4.value Then
                UASN.value = CDbl(UASN.value) - (CInt(SNC4.value))
            Else
                UASN.value = CDbl(UASN.value) + (CInt(SNC4.value))
            End If
        End If
    ElseIf SNT4.value = "IN" Then
        If snIN4 <> SN4.value Then
            snIN4 = SN4.value
            If SN4.value Then
                UASN.value = CDbl(UASN.value) - (CInt(SNC4.value) * 0.5)
            Else
                UASN.value = CDbl(UASN.value) + (CInt(SNC4.value) * 0.5)
            End If
        End If
    ElseIf SNT4.value = "RM" Then
        If snRM4 <> SN4.value Then
            snRM4 = SN4.value
            If SN4.value Then
                UASN.value = CDbl(UASN.value) - (CInt(SNC4.value) * 0.5)
            Else
                UASN.value = CDbl(UASN.value) + (CInt(SNC4.value) * 0.5)
            End If
        End If
    End If
End Sub

Private Sub SNC4_Change()
    If SNT4.value = "RP" Then
        If PreviousValueRPSNC4 <> SNC4.value Then
            UASN.value = CDbl(UASN.value) - (SNC4.value - PreviousValueRPSNC4)
            PreviousValueRPSNC4 = SNC4.value
        End If
    ElseIf SNT4.value = "IN" Then
        If PreviousValueINSNC4 <> SNC4.value Then
            UASN.value = CDbl(UASN.value) - (0.5 * (SNC4.value - PreviousValueINSNC4))
            PreviousValueINSNC4 = SNC4.value
        End If
    ElseIf SNT4.value = "RM" Then
        If PreviousValueRMSNC4 <> SNC4.value Then
            UASN.value = CDbl(UASN.value) - (0.5 * (SNC4.value - PreviousValueRMSNC4))
            PreviousValueRMSNC4 = SNC4.value
        End If
    End If
End Sub

Private Sub SNT5_Change()
    If SNT5.value = "RP" Then
        SN5.value = snRP5
        SNC5.value = PreviousValueRPSNC5
    ElseIf SNT5.value = "IN" Then
        SN5.value = snIN5
        SNC5.value = PreviousValueINSNC5
    ElseIf SNT5.value = "RM" Then
        SN5.value = snRM5
        SNC5.value = PreviousValueRMSNC5
    End If
End Sub

Private Sub SN5_Click()
    SNC5.visible = SN5.value
    
    If SNT5.value = "RP" Then
        If snRP5 <> SN5.value Then
            snRP5 = SN5.value
            If SN5.value Then
                UASN.value = CDbl(UASN.value) - (CInt(SNC5.value))
            Else
                UASN.value = CDbl(UASN.value) + (CInt(SNC5.value))
            End If
        End If
    ElseIf SNT5.value = "IN" Then
        If snIN5 <> SN5.value Then
            snIN5 = SN5.value
            If SN5.value Then
                UASN.value = CDbl(UASN.value) - (CInt(SNC5.value) * 0.5)
            Else
                UASN.value = CDbl(UASN.value) + (CInt(SNC5.value) * 0.5)
            End If
        End If
    ElseIf SNT5.value = "RM" Then
        If snRM5 <> SN5.value Then
            snRM5 = SN5.value
            If SN5.value Then
                UASN.value = CDbl(UASN.value) - (CInt(SNC5.value) * 0.5)
            Else
                UASN.value = CDbl(UASN.value) + (CInt(SNC5.value) * 0.5)
            End If
        End If
    End If
End Sub

Private Sub SNC5_Change()
    If SNT5.value = "RP" Then
        If PreviousValueRPSNC5 <> SNC5.value Then
            UASN.value = CDbl(UASN.value) - (SNC5.value - PreviousValueRPSNC5)
            PreviousValueRPSNC5 = SNC5.value
        End If
    ElseIf SNT5.value = "IN" Then
        If PreviousValueINSNC5 <> SNC5.value Then
            UASN.value = CDbl(UASN.value) - (0.5 * (SNC5.value - PreviousValueINSNC5))
            PreviousValueINSNC5 = SNC5.value
        End If
    ElseIf SNT5.value = "RM" Then
        If PreviousValueRMSNC5 <> SNC5.value Then
            UASN.value = CDbl(UASN.value) - (0.5 * (SNC5.value - PreviousValueRMSNC5))
            PreviousValueRMSNC5 = SNC5.value
        End If
    End If
End Sub

Private Sub SNT6_Change()
    If SNT6.value = "RP" Then
        SN6.value = snRP6
        SNC6.value = PreviousValueRPSNC6
    ElseIf SNT6.value = "IN" Then
        SN6.value = snIN6
        SNC6.value = PreviousValueINSNC6
    ElseIf SNT6.value = "RM" Then
        SN6.value = snRM6
        SNC6.value = PreviousValueRMSNC6
    End If
End Sub

Private Sub SN6_Click()
    SNC6.visible = SN6.value
    
    If SNT6.value = "RP" Then
        If snRP6 <> SN6.value Then
            snRP6 = SN6.value
            If SN6.value Then
                UASN.value = CDbl(UASN.value) - (CInt(SNC6.value))
            Else
                UASN.value = CDbl(UASN.value) + (CInt(SNC6.value))
            End If
        End If
    ElseIf SNT6.value = "IN" Then
        If snIN6 <> SN6.value Then
            snIN6 = SN6.value
            If SN6.value Then
                UASN.value = CDbl(UASN.value) - (CInt(SNC6.value) * 0.5)
            Else
                UASN.value = CDbl(UASN.value) + (CInt(SNC6.value) * 0.5)
            End If
        End If
    ElseIf SNT6.value = "RM" Then
        If snRM6 <> SN6.value Then
            snRM6 = SN6.value
            If SN6.value Then
                UASN.value = CDbl(UASN.value) - (CInt(SNC6.value) * 0.5)
            Else
                UASN.value = CDbl(UASN.value) + (CInt(SNC6.value) * 0.5)
            End If
        End If
    End If
End Sub

Private Sub SNC6_Change()
    If SNT6.value = "RP" Then
        If PreviousValueRPSNC6 <> SNC6.value Then
            UASN.value = CDbl(UASN.value) - (SNC6.value - PreviousValueRPSNC6)
            PreviousValueRPSNC6 = SNC6.value
        End If
    ElseIf SNT6.value = "IN" Then
        If PreviousValueINSNC6 <> SNC6.value Then
            UASN.value = CDbl(UASN.value) - (0.5 * (SNC6.value - PreviousValueINSNC6))
            PreviousValueINSNC6 = SNC6.value
        End If
    ElseIf SNT6.value = "RM" Then
        If PreviousValueRMSNC6 <> SNC6.value Then
            UASN.value = CDbl(UASN.value) - (0.5 * (SNC6.value - PreviousValueRMSNC6))
            PreviousValueRMSNC6 = SNC6.value
        End If
    End If
End Sub

Private Sub SNT7_Change()
    If SNT7.value = "RP" Then
        SN7.value = snRP7
        SNC7.value = PreviousValueRPSNC7
    ElseIf SNT7.value = "IN" Then
        SN7.value = snIN7
        SNC7.value = PreviousValueINSNC7
    ElseIf SNT7.value = "RM" Then
        SN7.value = snRM7
        SNC7.value = PreviousValueRMSNC7
    End If
End Sub

Private Sub SN7_Click()
    SNC7.visible = SN7.value
    
    If SNT7.value = "RP" Then
        If snRP7 <> SN7.value Then
            snRP7 = SN7.value
            If SN7.value Then
                UASN.value = CDbl(UASN.value) - (CInt(SNC7.value) * 2)
            Else
                UASN.value = CDbl(UASN.value) + (CInt(SNC7.value) * 2)
            End If
        End If
    ElseIf SNT7.value = "IN" Then
        If snIN7 <> SN7.value Then
            snIN7 = SN7.value
            If SN7.value Then
                UASN.value = CDbl(UASN.value) - (CInt(SNC7.value))
            Else
                UASN.value = CDbl(UASN.value) + (CInt(SNC7.value))
            End If
        End If
    ElseIf SNT7.value = "RM" Then
        If snRM7 <> SN7.value Then
            snRM7 = SN7.value
            If SN7.value Then
                UASN.value = CDbl(UASN.value) - (CInt(SNC7.value))
            Else
                UASN.value = CDbl(UASN.value) + (CInt(SNC7.value))
            End If
        End If
    End If
End Sub

Private Sub SNC7_Change()
    If SNT7.value = "RP" Then
        If PreviousValueRPSNC7 <> SNC7.value Then
            UASN.value = CDbl(UASN.value) - (2 * (SNC7.value - PreviousValueRPSNC7))
            PreviousValueRPSNC7 = SNC7.value
        End If
    ElseIf SNT7.value = "IN" Then
        If PreviousValueINSNC7 <> SNC7.value Then
            UASN.value = CDbl(UASN.value) - (SNC7.value - PreviousValueINSNC7)
            PreviousValueINSNC7 = SNC7.value
        End If
    ElseIf SNT7.value = "RM" Then
        If PreviousValueRMSNC7 <> SNC7.value Then
            UASN.value = CDbl(UASN.value) - (SNC7.value - PreviousValueRMSNC7)
            PreviousValueRMSNC7 = SNC7.value
        End If
    End If
End Sub

Private Sub TRLETTER_Change()
    Call determineRecDSpace
End Sub

Private Sub TRT_Click()
    If TRT.value And RPT.value Then RPT.value = False
    XSL.visible = RPP.value And (TRT.value Or RPT.value)
    TRSIZE.visible = RPP.value And (TRT.value Or RPT.value)
    CLL.visible = RPP.value And (TRT.value Or RPT.value)
    COLA.visible = RPP.value And (TRT.value Or RPT.value)
    XDL.visible = RPP.value And (TRT.value Or RPT.value)
    TRLETTER.visible = RPP.value And (TRT.value Or RPT.value)
    Call determineRecDSpace
End Sub

Private Sub RPT_Click()
    If RPT.value And TRT.value Then TRT.value = False
    XSL.visible = RPP.value And (TRT.value Or RPT.value)
    TRSIZE.visible = TRT.value Or RPT.value
    XRL.visible = RPP.value And RPT.value
    TRRPSIZE.visible = RPP.value And RPT.value
    CLL.visible = RPP.value And (TRT.value Or RPT.value)
    COLA.visible = RPP.value And (TRT.value Or RPT.value)
    XDL.visible = RPP.value And (TRT.value Or RPT.value)
    TRLETTER.visible = RPP.value And (TRT.value Or RPT.value)
    Call determineRecDSpace
End Sub

Private Sub UASN_Change()
    UASN2.value = UASN.value
End Sub
