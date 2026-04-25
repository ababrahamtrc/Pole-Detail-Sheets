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
Dim CrewNotes As String, installNotes As String, removeNotes As String, replaceNotes As String, transferNotes, notes As String
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
Dim mhRP1 As Boolean, mhIN1 As Boolean, mhRM1 As Boolean, mhRP2 As Boolean, mhIN2 As Boolean, mhRM2 As Boolean
Dim figuresUsed As Scripting.Dictionary
Dim comms As Scripting.Dictionary
Dim commTransfers As Collection

Public Sub Initialize(sheet As Worksheet)
    Set pds = sheet
    Set pole = New pole
    Call pole.extractFromSheet(sheet)
    Me.MultiPage1.Value = 0
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.height) - (0.5 * Me.height)
    
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
        If pds.Range("CMOWNER").offset(i, 0).Interior.color = 16777215 Then Exit For
        If pds.Range("CMOWNER").offset(i, 0) = "" Then Exit For
        If pds.Range("CMOWNER").offset(i, 0).text = "Clearance Requirement" Then Exit For
        If InStr(pds.Range("CMSIZE").offset(i, 0).text, """") > 0 Then
            owner = Replace(pds.Range("CMOWNER").offset(i, 0).text, " SVC", "")
            If Not comms.Exists(owner) Then comms.Add owner, Nothing
        End If
    Next i
    
    For i = 0 To 50
        If pds.Range("UTTYPE").offset(i, 0).Interior.color = 16777215 Then Exit For
        If pds.Range("UTTYPE").offset(i, 0) = "" Then Exit For
        If pds.Range("UTTYPE").offset(i, 0).text = "STLT. BOTTOM BRKT." Then Exit For
        If InStr(pds.Range("UTTYPE").offset(i, 0).text, "NEUT") > 0 Then neutBool = True
        If InStr(pds.Range("UTTYPE").offset(i, 0).text, "SEC") > 0 Then secBool = True
        If InStr(pds.Range("UTTYPE").offset(i, 0).text, "OW") > 0 Then owBool = True
        If InStr(pds.Range("UTTYPE").offset(i, 0).text, "RISER") > 0 Then riserBool = True
        If InStr(pds.Range("UTTYPE").offset(i, 0).text, "DG") > 0 And InStr(pds.Range("UTSIZE").offset(i, 0).text, "11K") Then dg11KCount = dg11KCount + 1
        If InStr(pds.Range("UTTYPE").offset(i, 0).text, "DG") > 0 And InStr(pds.Range("UTSIZE").offset(i, 0).text, "20K") Then dg20KCount = dg20KCount + 1
        If InStr(pds.Range("UTTYPE").offset(i, 0).text, "SVC") > 0 Then
            For j = 1 To 12
                For Each name In pds.names
                    If name.name = "'" & pds.name & "'" & "!" & "TOPOLE" & j Then
                        If Trim(Replace(pds.Range("UTMIDSPAN" & j).offset(i, 0), "-", "")) <> "" Then
                            ohsCount = ohsCount + 1
                        End If
                    End If
                Next name
            Next j
        End If
        If InStr(pds.Range("UTTYPE").offset(i, 0).text, "SPG") > 0 Then
            For j = 1 To 12
                For Each name In pds.names
                    If name.name = "'" & pds.name & "'" & "!" & "TOPOLE" & j Then
                        If Trim(Replace(pds.Range("UTMIDSPAN" & j).offset(i, 0), "-", "")) <> "" And InStr(pds.Range("UTSIZE").offset(i, 0).text, "11K") Then spg11KCount = spg11KCount + 1
                        If Trim(Replace(pds.Range("UTMIDSPAN" & j).offset(i, 0), "-", "")) <> "" And InStr(pds.Range("UTSIZE").offset(i, 0).text, "20K") Then spg20KCount = spg20KCount + 1
                    End If
                Next name
            Next j
        End If
        If InStr(pds.Range("UTTYPE").offset(i, 0).text, "OW") > 0 Or InStr(pds.Range("UTTYPE").offset(i, 0).text, "SEC") > 0 Or InStr(pds.Range("UTTYPE").offset(i, 0).text, "NEUT") > 0 Then
            For j = 1 To 12
                For Each name In pds.names
                    If name.name = "'" & pds.name & "'" & "!" & "TOPOLE" & j Then
                        If Trim(Replace(pds.Range("UTMIDSPAN" & j).offset(i, 0), "-", "")) <> "" Then snC = snC + 1
                    End If
                Next name
            Next j
        End If
    Next i
    
    For i = 0 To 3
        If InStr(pds.Range("DESC").offset(i, 0).Value, "3") > 0 Then priC = priC + (3 * (CountCharInString(pds.Range("DIRECTION").offset(i, 0), "/") + 1))
        If InStr(pds.Range("DESC").offset(i, 0).Value, "2") > 0 Then priC = priC + (2 * (CountCharInString(pds.Range("DIRECTION").offset(i, 0), "/") + 1))
        If InStr(pds.Range("DESC").offset(i, 0).Value, "1") > 0 Then priC = priC + (CountCharInString(pds.Range("DIRECTION").offset(i, 0), "/") + 1)
    Next i
    
    UAP.Value = priC
    UASN.Value = snC
    
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
    If PSV.Value = "" And BAFIG.Value Then PSV.Value = "4'0"""
End Sub

Private Sub BSL_Click()
    BSLRFW.visible = BSL.Value
    RPBSL.Value = BSL.Value
    RPBSLRFW.visible = RPP.Value And BSL.Value
    SLRFWL.visible = RPP.Value And BSL.Value
    RFWL2.visible = BSL.Value
End Sub

Private Sub BSLRFW_Change()
    RPBSLRFW.Value = BSLRFW.Value
End Sub

Private Sub MHT1_Change()
    If MHT1.Value = "RP" Then
        MH1.Value = mhRP1
        MHC1.Value = previousValueRPMHC1
    ElseIf MHT1.Value = "IN" Then
        MH1.Value = mhIN1
        MHC1.Value = previousValueINMHC1
    ElseIf MHT1.Value = "RM" Then
        MH1.Value = mhRM1
        MHC1.Value = previousValueRMMHC1
    End If
End Sub

Private Sub MH1_Click()
    MHC1.visible = MH1.Value
    If MHT1.Value = "RP" Then
        If mhRP1 <> MH1.Value Then mhRP1 = MH1.Value
    ElseIf MHT1.Value = "IN" Then
        If mhIN1 <> MH1.Value Then mhIN1 = MH1.Value
    ElseIf MHT1.Value = "RM" Then
        If mhRM1 <> MH1.Value Then mhRM1 = MH1.Value
    End If
End Sub

Private Sub MHC1_Change()
    If MHT1.Value = "RP" Then
        previousValueRPMHC1 = MHC1.Value
    ElseIf MHT1.Value = "IN" Then
        previousValueINMHC1 = MHC1.Value
    ElseIf MHT1.Value = "RM" Then
        previousValueRMMHC1 = MHC1.Value
    End If
End Sub

Private Sub MHT2_Change()
    If MHT2.Value = "RP" Then
        MH2.Value = mhRP2
        MHC2.Value = previousValueRPMHC2
    ElseIf MHT2.Value = "IN" Then
        MH2.Value = mhIN2
        MHC2.Value = previousValueINMHC2
    ElseIf MHT2.Value = "RM" Then
        MH2.Value = mhRM2
        MHC2.Value = previousValueRMMHC2
    End If
End Sub

Private Sub MH2_Click()
    MHC2.visible = MH2.Value
    If MHT2.Value = "RP" Then
        If mhRP2 <> MH2.Value Then mhRP2 = MH2.Value
    ElseIf MHT2.Value = "IN" Then
        If mhIN2 <> MH2.Value Then mhIN2 = MH2.Value
    ElseIf MHT2.Value = "RM" Then
        If mhRM2 <> MH2.Value Then mhRM2 = MH2.Value
    End If
End Sub

Private Sub MHC2_Change()
    If MHT2.Value = "RP" Then
        previousValueRPMHC2 = MHC2.Value
    ElseIf MHT2.Value = "IN" Then
        previousValueINMHC2 = MHC2.Value
    ElseIf MHT2.Value = "RM" Then
        previousValueRMMHC2 = MHC2.Value
    End If
End Sub

Private Sub FUT1_Change()
    If FUT1.Value = "RP" Then
        FU1.Value = fuRP1
        FUC1.Value = PreviousValueRPFUC1
    ElseIf FUT1.Value = "IN" Then
        FU1.Value = fuIN1
        FUC1.Value = PreviousValueINFUC1
    ElseIf FUT1.Value = "RM" Then
        FU1.Value = fuRM1
        FUC1.Value = PreviousValueRMFUC1
    End If
End Sub

Private Sub FU1_Click()
    FUC1.visible = FU1.Value
    If FUT1.Value = "RP" Then
        If fuRP1 <> FU1.Value Then fuRP1 = FU1.Value
    ElseIf FUT1.Value = "IN" Then
        If fuIN1 <> FU1.Value Then fuIN1 = FU1.Value
    ElseIf FUT1.Value = "RM" Then
        If fuRM1 <> FU1.Value Then fuRM1 = FU1.Value
    End If
End Sub

Private Sub FUC1_Change()
    If FUT1.Value = "RP" Then
        PreviousValueRPFUC1 = FUC1.Value
    ElseIf FUT1.Value = "IN" Then
        PreviousValueINFUC1 = FUC1.Value
    ElseIf FUT1.Value = "RM" Then
        PreviousValueRMFUC1 = FUC1.Value
    End If
End Sub

Private Sub FUT2_Change()
    If FUT2.Value = "RP" Then
        FU2.Value = fuRP2
        FUC2.Value = PreviousValueRPFUC2
    ElseIf FUT2.Value = "IN" Then
        FU2.Value = fuIN2
        FUC2.Value = PreviousValueINFUC2
    ElseIf FUT2.Value = "RM" Then
        FU2.Value = fuRM2
        FUC2.Value = PreviousValueRMFUC2
    End If
End Sub

Private Sub FU2_Click()
    FUC2.visible = FU2.Value
    If FUT2.Value = "RP" Then
        If fuRP2 <> FU2.Value Then fuRP2 = FU2.Value
    ElseIf FUT2.Value = "IN" Then
        If fuIN2 <> FU2.Value Then fuIN2 = FU2.Value
    ElseIf FUT2.Value = "RM" Then
        If fuRM2 <> FU2.Value Then fuRM2 = FU2.Value
    End If
End Sub

Private Sub FUC2_Change()
    If FUT2.Value = "RP" Then
        PreviousValueRPFUC2 = FUC2.Value
    ElseIf FUT2.Value = "IN" Then
        PreviousValueINFUC2 = FUC2.Value
    ElseIf FUT2.Value = "RM" Then
        PreviousValueRMFUC2 = FUC2.Value
    End If
End Sub

Private Sub FUT3_Change()
    If FUT3.Value = "RP" Then
        FU3.Value = fuRP3
        FUC3.Value = PreviousValueRPFUC3
    ElseIf FUT3.Value = "IN" Then
        FU3.Value = fuIN3
        FUC3.Value = PreviousValueINFUC3
    ElseIf FUT3.Value = "RM" Then
        FU3.Value = fuRM3
        FUC3.Value = PreviousValueRMFUC3
    End If
End Sub

Private Sub FU3_Click()
    FUC3.visible = FU3.Value
    If FUT3.Value = "RP" Then
        If fuRP3 <> FU3.Value Then fuRP3 = FU3.Value
    ElseIf FUT3.Value = "IN" Then
        If fuIN3 <> FU3.Value Then fuIN3 = FU3.Value
    ElseIf FUT3.Value = "RM" Then
        If fuRM3 <> FU3.Value Then fuRM3 = FU3.Value
    End If
End Sub

Private Sub FUC3_Change()
    If FUT3.Value = "RP" Then
        PreviousValueRPFUC3 = FUC3.Value
    ElseIf FUT3.Value = "IN" Then
        PreviousValueINFUC3 = FUC3.Value
    ElseIf FUT3.Value = "RM" Then
        PreviousValueRMFUC3 = FUC3.Value
    End If
End Sub

Private Sub IA11_Change()
    If IA11.Value = "RS" Then
        IA12.list = Array("11K")
        IA12.ListIndex = 0
    ElseIf IA11.Value = "RT" Then
        IA12.list = Array("(2)11K", "20K")
        IA12.ListIndex = 0
    ElseIf IA11.Value = "STE" Then
        IA12.list = Array("(3)11K", "20K + 11K", "(2)20K")
        IA12.ListIndex = 0
    End If
End Sub

Private Sub IA21_Change()
    If IA21.Value = "RS" Then
        IA22.list = Array("11K")
        IA22.ListIndex = 0
    ElseIf IA21.Value = "RT" Then
        IA22.list = Array("(2)11K", "20K")
        IA22.ListIndex = 0
    ElseIf IA21.Value = "STE" Then
        IA22.list = Array("(3)11K", "20K + 11K", "(2)20K")
        IA22.ListIndex = 0
    End If
End Sub

Private Sub IA31_Change()
    If IA31.Value = "RS" Then
        IA32.list = Array("11K")
        IA32.ListIndex = 0
    ElseIf IA31.Value = "RT" Then
        IA32.list = Array("(2)11K", "20K")
        IA32.ListIndex = 0
    ElseIf IA31.Value = "STE" Then
        IA32.list = Array("(3)11K", "20K + 11K", "(2)20K")
        IA32.ListIndex = 0
    End If
End Sub

Private Sub EA11_Change()
    If EA11.Value = "RS" Then
        EA12.list = Array("11K")
        EA12.ListIndex = 0
    ElseIf EA11.Value = "RT" Then
        EA12.list = Array("(2)11K", "20K")
        EA12.ListIndex = 0
    ElseIf EA11.Value = "STE" Then
        EA12.list = Array("(3)11K", "20K + 11K", "(2)20K")
        EA12.ListIndex = 0
    End If
    RA11.Value = EA11.Value
End Sub

Private Sub EA21_Change()
    If EA21.Value = "RS" Then
        EA22.list = Array("11K")
        EA22.ListIndex = 0
    ElseIf EA21.Value = "RT" Then
        EA22.list = Array("(2)11K", "20K")
        EA22.ListIndex = 0
    ElseIf EA21.Value = "STE" Then
        EA22.list = Array("(3)11K", "20K + 11K", "(2)20K")
        EA22.ListIndex = 0
    End If
    RA21.Value = EA21.Value
End Sub

Private Sub EA31_Change()
    If EA31.Value = "RS" Then
        EA32.list = Array("11K")
        EA32.ListIndex = 0
    ElseIf EA31.Value = "RT" Then
        EA32.list = Array("(2)11K", "20K")
        EA32.ListIndex = 0
    ElseIf EA31.Value = "STE" Then
        EA32.list = Array("(3)11K", "20K + 11K", "(2)20K")
        EA32.ListIndex = 0
    End If
    RA31.Value = EA31.Value
End Sub

Private Sub EA41_Change()
    If EA41.Value = "RS" Then
        EA42.list = Array("11K")
        EA42.ListIndex = 0
    ElseIf EA41.Value = "RT" Then
        EA42.list = Array("(2)11K", "20K")
        EA42.ListIndex = 0
    ElseIf EA41.Value = "STE" Then
        EA42.list = Array("(3)11K", "20K + 11K", "(2)20K")
        EA42.ListIndex = 0
    End If
    RA41.Value = EA41.Value
End Sub

Private Sub PSV_Change()
    Call determineRecDSpace
End Sub

Private Sub RA11_Change()
    If RA11.Value = "RS" Then
        RA12.list = Array("11K")
        RA12.ListIndex = 0
    ElseIf RA11.Value = "RT" Then
        RA12.list = Array("(2)11K", "20K")
        RA12.ListIndex = 0
    ElseIf RA11.Value = "STE" Then
        RA12.list = Array("(3)11K", "20K + 11K", "(2)20K")
        RA12.ListIndex = 0
    End If
End Sub

Private Sub RA21_Change()
    If RA21.Value = "RS" Then
        RA22.list = Array("11K")
        RA22.ListIndex = 0
    ElseIf RA21.Value = "RT" Then
        RA22.list = Array("(2)11K", "20K")
        RA22.ListIndex = 0
    ElseIf RA21.Value = "STE" Then
        RA22.list = Array("(3)11K", "20K + 11K", "(2)20K")
        RA22.ListIndex = 0
    End If
End Sub

Private Sub RA31_Change()
    If RA31.Value = "RS" Then
        RA32.list = Array("11K")
        RA32.ListIndex = 0
    ElseIf RA31.Value = "RT" Then
        RA32.list = Array("(2)11K", "20K")
        RA32.ListIndex = 0
    ElseIf RA31.Value = "STE" Then
        RA32.list = Array("(3)11K", "20K + 11K", "(2)20K")
        RA32.ListIndex = 0
    End If
End Sub

Private Sub RA41_Change()
    If RA41.Value = "RS" Then
        RA42.list = Array("11K")
        RA42.ListIndex = 0
    ElseIf RA41.Value = "RT" Then
        RA42.list = Array("(2)11K", "20K")
        RA42.ListIndex = 0
    ElseIf RA41.Value = "STE" Then
        RA42.list = Array("(3)11K", "20K + 11K", "(2)20K")
        RA42.ListIndex = 0
    End If
End Sub

Private Sub EA12_Change()
    RA12.Value = EA12.Value
End Sub

Private Sub EA22_Change()
    RA22.Value = EA22.Value
End Sub

Private Sub EA32_Change()
    RA32.Value = EA32.Value
End Sub

Private Sub EA42_Change()
    RA42.Value = EA42.Value
End Sub

Private Sub IAC_Change()
    For i = 1 To 3
        If CInt(IAC.Value) >= i Then
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
    OWSIZE1.visible = OWSPAN1.Value <> ""
    OWNEWSIZE1.visible = OWSPAN1.Value <> ""
    OWLENGTH1.visible = OWSPAN1.Value <> ""
    OW1L2.visible = OWSPAN1.Value <> ""
    OW1L3.visible = OWSPAN1.Value <> ""
    OW1L4.visible = OWSPAN1.Value <> ""
    If OWSPAN1.Value <> "" Then
        OWSIZE1.Value = getOpenWireSizeFromSpan(OWSPAN1.Value)
        Set cell = pds.UsedRange.find(what:=OWSPAN1.Value, LookIn:=xlValues, lookat:=xlWhole, MatchCase:=True)
        If Not cell Is Nothing Then
            OWLENGTH1.Value = cell.offset(-1, 0).Value
        End If
    End If
End Sub

Private Sub OWSPAN2_Change()
    OWSIZE2.visible = OWSPAN2.Value <> ""
    OWNEWSIZE2.visible = OWSPAN2.Value <> ""
    OWLENGTH2.visible = OWSPAN2.Value <> ""
    OW2L2.visible = OWSPAN2.Value <> ""
    OW2L3.visible = OWSPAN2.Value <> ""
    OW2L4.visible = OWSPAN2.Value <> ""
    If OWSPAN2.Value <> "" Then
        OWSIZE2.Value = getOpenWireSizeFromSpan(OWSPAN2.Value)
        Set cell = pds.UsedRange.find(what:=OWSPAN2.Value, LookIn:=xlValues, lookat:=xlWhole, MatchCase:=True)
        If Not cell Is Nothing Then
            OWLENGTH2.Value = cell.offset(-1, 0).Value
        End If
    End If
End Sub

Private Sub OWSPAN3_Change()
    OWSIZE3.visible = OWSPAN3.Value <> ""
    OWNEWSIZE3.visible = OWSPAN3.Value <> ""
    OWLENGTH3.visible = OWSPAN3.Value <> ""
    OW3L2.visible = OWSPAN3.Value <> ""
    OW3L3.visible = OWSPAN3.Value <> ""
    OW3L4.visible = OWSPAN3.Value <> ""
    If OWSPAN3.Value <> "" Then
        OWSIZE3.Value = getOpenWireSizeFromSpan(OWSPAN3.Value)
        Set cell = pds.UsedRange.find(what:=OWSPAN3.Value, LookIn:=xlValues, lookat:=xlWhole, MatchCase:=True)
        If Not cell Is Nothing Then
            OWLENGTH3.Value = cell.offset(-1, 0).Value
        End If
    End If
End Sub

Private Sub OWSPAN4_Change()
    OWSIZE4.visible = OWSPAN4.Value <> ""
    OWNEWSIZE4.visible = OWSPAN4.Value <> ""
    OWLENGTH4.visible = OWSPAN4.Value <> ""
    OW4L2.visible = OWSPAN4.Value <> ""
    OW4L3.visible = OWSPAN4.Value <> ""
    OW4L4.visible = OWSPAN4.Value <> ""
    If OWSPAN4.Value <> "" Then
        OWSIZE4.Value = getOpenWireSizeFromSpan(OWSPAN4.Value)
        Set cell = pds.UsedRange.find(what:=OWSPAN4.Value, LookIn:=xlValues, lookat:=xlWhole, MatchCase:=True)
        If Not cell Is Nothing Then
            OWLENGTH4.Value = cell.offset(-1, 0).Value
        End If
    End If
End Sub

Private Sub OWSPAN5_Change()
    OWSIZE5.visible = OWSPAN5.Value <> ""
    OWNEWSIZE5.visible = OWSPAN5.Value <> ""
    OWLENGTH5.visible = OWSPAN5.Value <> ""
    OW5L2.visible = OWSPAN5.Value <> ""
    OW5L3.visible = OWSPAN5.Value <> ""
    OW5L4.visible = OWSPAN5.Value <> ""
    If OWSPAN5.Value <> "" Then
        OWSIZE5.Value = getOpenWireSizeFromSpan(OWSPAN5.Value)
        Set cell = pds.UsedRange.find(what:=OWSPAN5.Value, LookIn:=xlValues, lookat:=xlWhole, MatchCase:=True)
        If Not cell Is Nothing Then
            OWLENGTH5.Value = cell.offset(-1, 0).Value
        End If
    End If
End Sub

Private Function getOpenWireSizeFromSpan(span As String) As String
    Dim found As Integer: found = 0
    Dim name As Variant
    Dim openWireSizes As String: Output = ""
    For i = 1 To 12
        For Each name In pds.names
            If name.name = "'" & pds.name & "'" & "!" & "TOPOLE" & i Then
                If pds.Range("TOPOLE" & i).Value = span Then found = i
                Exit For
            End If
        Next name
        If found > 0 Then Exit For
    Next i
    
    For i = 0 To 50
        If pds.Range("UTTYPE").offset(i, 0).Interior.color = 16777215 Then Exit For
        If pds.Range("UTTYPE").offset(i, 0) = "" Then Exit For
        If pds.Range("UTTYPE").offset(i, 0).text = "STLT. BOTTOM BRKT." Then Exit For
        If InStr(pds.Range("UTTYPE").offset(i, 0).text, "OW") > 0 Then
            If Trim(Replace(pds.Range("UTMIDSPAN" & found).offset(i, 0).Value, "-", "")) <> "" Then
                openWireSizes = openWireSizes & ExtractNumbers(pds.Range("UTSIZE").offset(i, 0).Value) & "-"
            End If
        End If
    Next i
    
    If Len(openWireSizes) > 0 Then openWireSizes = Left(openWireSizes, Len(openWireSizes) - 1)
    
    getOpenWireSizeFromSpan = openWireSizes
    
End Function

Private Function ExtractNumbers(ByVal InputString As String) As String
    Dim i As Long
    Dim NumericString As String

    For i = 1 To Len(InputString)
        If IsNumeric(Mid(InputString, i, 1)) Then
            NumericString = NumericString & Mid(InputString, i, 1)
        End If
    Next i

    ExtractNumbers = NumericString
End Function

Private Sub PRIT1_Change()
    If PRIT1.Value = "RP" Then
        PRI1.Value = priRP1
        PRIC1.Value = PreviousValueRPPRIC1
    ElseIf PRIT1.Value = "IN" Then
        PRI1.Value = priIN1
        PRIC1.Value = PreviousValueINPRIC1
    ElseIf PRIT1.Value = "RM" Then
        PRI1.Value = priRM1
        PRIC1.Value = PreviousValueRMPRIC1
    End If
End Sub

Private Sub PRI1_Click()
    PRIC1.visible = PRI1.Value
    
    If PRIT1.Value = "RP" Then
        If priRP1 <> PRI1.Value Then
            priRP1 = PRI1.Value
            If PRI1.Value Then
                UAP.Value = CDbl(UAP.Value) - (CInt(PRIC1.Value) * 2)
            Else
                UAP.Value = CDbl(UAP.Value) + (CInt(PRIC1.Value) * 2)
            End If
        End If
    ElseIf PRIT1.Value = "IN" Then
        If priIN1 <> PRI1.Value Then
            priIN1 = PRI1.Value
            If PRI1.Value Then
                UAP.Value = CDbl(UAP.Value) - (CInt(PRIC1.Value))
            Else
                UAP.Value = CDbl(UAP.Value) + (CInt(PRIC1.Value))
            End If
        End If
    ElseIf PRIT1.Value = "RM" Then
        If priRM1 <> PRI1.Value Then
            priRM1 = PRI1.Value
            If PRI1.Value Then
                UAP.Value = CDbl(UAP.Value) - (CInt(PRIC1.Value))
            Else
                UAP.Value = CDbl(UAP.Value) + (CInt(PRIC1.Value))
            End If
        End If
    End If
End Sub

Private Sub PRIC1_Change()
    If PRIT1.Value = "RP" Then
        If PreviousValueRPPRIC1 <> PRIC1.Value Then
            UAP.Value = CDbl(UAP.Value) - (2 * (PRIC1.Value - PreviousValueRPPRIC1))
            PreviousValueRPPRIC1 = PRIC1.Value
        End If
    ElseIf PRIT1.Value = "IN" Then
        If PreviousValueINPRIC1 <> PRIC1.Value Then
            UAP.Value = CDbl(UAP.Value) - (PRIC1.Value - PreviousValueINPRIC1)
            PreviousValueINPRIC1 = PRIC1.Value
        End If
    ElseIf PRIT1.Value = "RM" Then
        If PreviousValueRMPRIC1 <> PRIC1.Value Then
            UAP.Value = CDbl(UAP.Value) - (PRIC1.Value - PreviousValueRMPRIC1)
            PreviousValueRMPRIC1 = PRIC1.Value
        End If
    End If
End Sub

Private Sub PRIT2_Change()
    If PRIT2.Value = "RP" Then
        PRI2.Value = priRP2
        PRIC2.Value = PreviousValueRPPRIC2
    ElseIf PRIT2.Value = "IN" Then
        PRI2.Value = priIN2
        PRIC2.Value = PreviousValueINPRIC2
    ElseIf PRIT2.Value = "RM" Then
        PRI2.Value = priRM2
        PRIC2.Value = PreviousValueRMPRIC2
    End If
End Sub

Private Sub PRI2_Click()
    PRIC2.visible = PRI2.Value
    
    If PRIT2.Value = "RP" Then
        If priRP2 <> PRI2.Value Then
            priRP2 = PRI2.Value
            If PRI2.Value Then
                UAP.Value = CDbl(UAP.Value) - (CInt(PRIC2.Value) * 2)
            Else
                UAP.Value = CDbl(UAP.Value) + (CInt(PRIC2.Value) * 2)
            End If
        End If
    ElseIf PRIT2.Value = "IN" Then
        If priIN2 <> PRI2.Value Then
            priIN2 = PRI2.Value
            If PRI2.Value Then
                UAP.Value = CDbl(UAP.Value) - (CInt(PRIC2.Value))
            Else
                UAP.Value = CDbl(UAP.Value) + (CInt(PRIC2.Value))
            End If
        End If
    ElseIf PRIT2.Value = "RM" Then
        If priRM2 <> PRI2.Value Then
            priRM2 = PRI2.Value
            If PRI2.Value Then
                UAP.Value = CDbl(UAP.Value) - (CInt(PRIC2.Value))
            Else
                UAP.Value = CDbl(UAP.Value) + (CInt(PRIC2.Value))
            End If
        End If
    End If
End Sub

Private Sub PRIC2_Change()
    If PRIT2.Value = "RP" Then
        If PreviousValueRPPRIC2 <> PRIC2.Value Then
            UAP.Value = CDbl(UAP.Value) - (2 * (PRIC2.Value - PreviousValueRPPRIC2))
            PreviousValueRPPRIC2 = PRIC2.Value
        End If
    ElseIf PRIT2.Value = "IN" Then
        If PreviousValueINPRIC2 <> PRIC2.Value Then
            UAP.Value = CDbl(UAP.Value) - (PRIC2.Value - PreviousValueINPRIC2)
            PreviousValueINPRIC2 = PRIC2.Value
        End If
    ElseIf PRIT2.Value = "RM" Then
        If PreviousValueRMPRIC2 <> PRIC2.Value Then
            UAP.Value = CDbl(UAP.Value) - (PRIC2.Value - PreviousValueRMPRIC2)
            PreviousValueRMPRIC2 = PRIC2.Value
        End If
    End If
End Sub

Private Sub PRIT3_Change()
    If PRIT3.Value = "RP" Then
        PRI3.Value = priRP3
        PRIC3.Value = PreviousValueRPPRIC3
    ElseIf PRIT3.Value = "IN" Then
        PRI3.Value = priIN3
        PRIC3.Value = PreviousValueINPRIC3
    ElseIf PRIT3.Value = "RM" Then
        PRI3.Value = priRM3
        PRIC3.Value = PreviousValueRMPRIC3
    End If
End Sub

Private Sub PRI3_Click()
    PRIC3.visible = PRI3.Value
    
    If PRIT3.Value = "RP" Then
        If priRP3 <> PRI3.Value Then
            priRP3 = PRI3.Value
            If PRI3.Value Then
                UAP.Value = CDbl(UAP.Value) - (CInt(PRIC3.Value) * 2)
            Else
                UAP.Value = CDbl(UAP.Value) + (CInt(PRIC3.Value) * 2)
            End If
        End If
    ElseIf PRIT3.Value = "IN" Then
        If priIN3 <> PRI3.Value Then
            priIN3 = PRI3.Value
            If PRI3.Value Then
                UAP.Value = CDbl(UAP.Value) - (CInt(PRIC3.Value))
            Else
                UAP.Value = CDbl(UAP.Value) + (CInt(PRIC3.Value))
            End If
        End If
    ElseIf PRIT3.Value = "RM" Then
        If priRM3 <> PRI3.Value Then
            priRM3 = PRI3.Value
            If PRI3.Value Then
                UAP.Value = CDbl(UAP.Value) - (CInt(PRIC3.Value))
            Else
                UAP.Value = CDbl(UAP.Value) + (CInt(PRIC3.Value))
            End If
        End If
    End If
End Sub

Private Sub PRIC3_Change()
    If PRIT3.Value = "RP" Then
        If PreviousValueRPPRIC3 <> PRIC3.Value Then
            UAP.Value = CDbl(UAP.Value) - (2 * (PRIC3.Value - PreviousValueRPPRIC3))
            PreviousValueRPPRIC3 = PRIC3.Value
        End If
    ElseIf PRIT3.Value = "IN" Then
        If PreviousValueINPRIC3 <> PRIC3.Value Then
            UAP.Value = CDbl(UAP.Value) - (PRIC3.Value - PreviousValueINPRIC3)
            PreviousValueINPRIC3 = PRIC3.Value
        End If
    ElseIf PRIT3.Value = "RM" Then
        If PreviousValueRMPRIC3 <> PRIC3.Value Then
            UAP.Value = CDbl(UAP.Value) - (PRIC3.Value - PreviousValueRMPRIC3)
            PreviousValueRMPRIC3 = PRIC3.Value
        End If
    End If
End Sub

Private Sub PRIT4_Change()
    If PRIT4.Value = "RP" Then
        PRI4.Value = priRP4
        PRIC4.Value = PreviousValueRPPRIC4
    ElseIf PRIT4.Value = "IN" Then
        PRI4.Value = priIN4
        PRIC4.Value = PreviousValueINPRIC4
    ElseIf PRIT4.Value = "RM" Then
        PRI4.Value = priRM4
        PRIC4.Value = PreviousValueRMPRIC4
    End If
End Sub

Private Sub PRI4_Click()
    PRIC4.visible = PRI4.Value
    
    If PRIT4.Value = "RP" Then
        If priRP4 <> PRI4.Value Then
            priRP4 = PRI4.Value
            If PRI4.Value Then
                UAP.Value = CDbl(UAP.Value) - (CInt(PRIC4.Value) * 2)
            Else
                UAP.Value = CDbl(UAP.Value) + (CInt(PRIC4.Value) * 2)
            End If
        End If
    ElseIf PRIT4.Value = "IN" Then
        If priIN4 <> PRI4.Value Then
            priIN4 = PRI4.Value
            If PRI4.Value Then
                UAP.Value = CDbl(UAP.Value) - (CInt(PRIC4.Value))
            Else
                UAP.Value = CDbl(UAP.Value) + (CInt(PRIC4.Value))
            End If
        End If
    ElseIf PRIT4.Value = "RM" Then
        If priRM4 <> PRI4.Value Then
            priRM4 = PRI4.Value
            If PRI4.Value Then
                UAP.Value = CDbl(UAP.Value) - (CInt(PRIC4.Value))
            Else
                UAP.Value = CDbl(UAP.Value) + (CInt(PRIC4.Value))
            End If
        End If
    End If
End Sub

Private Sub PRIC4_Change()
    If PRIT4.Value = "RP" Then
        If PreviousValueRPPRIC4 <> PRIC4.Value Then
            UAP.Value = CDbl(UAP.Value) - (2 * (PRIC4.Value - PreviousValueRPPRIC4))
            PreviousValueRPPRIC4 = PRIC4.Value
        End If
    ElseIf PRIT4.Value = "IN" Then
        If PreviousValueINPRIC4 <> PRIC4.Value Then
            UAP.Value = CDbl(UAP.Value) - (PRIC4.Value - PreviousValueINPRIC4)
            PreviousValueINPRIC4 = PRIC4.Value
        End If
    ElseIf PRIT4.Value = "RM" Then
        If PreviousValueRMPRIC4 <> PRIC4.Value Then
            UAP.Value = CDbl(UAP.Value) - (PRIC4.Value - PreviousValueRMPRIC4)
            PreviousValueRMPRIC4 = PRIC4.Value
        End If
    End If
End Sub

Private Sub PRIT5_Change()
    If PRIT5.Value = "RP" Then
        PRI5.Value = priRP5
        PRIC5.Value = PreviousValueRPPRIC5
    ElseIf PRIT5.Value = "IN" Then
        PRI5.Value = priIN5
        PRIC5.Value = PreviousValueINPRIC5
    ElseIf PRIT5.Value = "RM" Then
        PRI5.Value = priRM5
        PRIC5.Value = PreviousValueRMPRIC5
    End If
End Sub

Private Sub PRI5_Click()
    PRIC5.visible = PRI5.Value
    
    If PRIT5.Value = "RP" Then
        If priRP5 <> PRI5.Value Then
            priRP5 = PRI5.Value
            If PRI5.Value Then
                UAP.Value = CDbl(UAP.Value) - (CInt(PRIC5.Value) * 2)
            Else
                UAP.Value = CDbl(UAP.Value) + (CInt(PRIC5.Value) * 2)
            End If
        End If
    ElseIf PRIT5.Value = "IN" Then
        If priIN5 <> PRI5.Value Then
            priIN5 = PRI5.Value
            If PRI5.Value Then
                UAP.Value = CDbl(UAP.Value) - (CInt(PRIC5.Value))
            Else
                UAP.Value = CDbl(UAP.Value) + (CInt(PRIC5.Value))
            End If
        End If
    ElseIf PRIT5.Value = "RM" Then
        If priRM5 <> PRI5.Value Then
            priRM5 = PRI5.Value
            If PRI5.Value Then
                UAP.Value = CDbl(UAP.Value) - (CInt(PRIC5.Value))
            Else
                UAP.Value = CDbl(UAP.Value) + (CInt(PRIC5.Value))
            End If
        End If
    End If
End Sub

Private Sub PRIC5_Change()
    If PRIT5.Value = "RP" Then
        If PreviousValueRPPRIC5 <> PRIC5.Value Then
            UAP.Value = CDbl(UAP.Value) - (2 * (PRIC5.Value - PreviousValueRPPRIC5))
            PreviousValueRPPRIC5 = PRIC5.Value
        End If
    ElseIf PRIT5.Value = "IN" Then
        If PreviousValueINPRIC5 <> PRIC5.Value Then
            UAP.Value = CDbl(UAP.Value) - (PRIC5.Value - PreviousValueINPRIC5)
            PreviousValueINPRIC5 = PRIC5.Value
        End If
    ElseIf PRIT5.Value = "RM" Then
        If PreviousValueRMPRIC5 <> PRIC5.Value Then
            UAP.Value = CDbl(UAP.Value) - (PRIC5.Value - PreviousValueRMPRIC5)
            PreviousValueRMPRIC5 = PRIC5.Value
        End If
    End If
End Sub

Private Sub PRIT6_Change()
    If PRIT6.Value = "RP" Then
        PRI6.Value = priRP6
        PRIC6.Value = PreviousValueRPPRIC6
    ElseIf PRIT6.Value = "IN" Then
        PRI6.Value = priIN6
        PRIC6.Value = PreviousValueINPRIC6
    ElseIf PRIT6.Value = "RM" Then
        PRI6.Value = priRM6
        PRIC6.Value = PreviousValueRMPRIC6
    End If
End Sub

Private Sub PRI6_Click()
    PRIC6.visible = PRI6.Value
    
    If PRIT6.Value = "RP" Then
        If priRP6 <> PRI6.Value Then
            priRP6 = PRI6.Value
            If PRI6.Value Then
                UAP.Value = CDbl(UAP.Value) - (CInt(PRIC6.Value))
            Else
                UAP.Value = CDbl(UAP.Value) + (CInt(PRIC6.Value))
            End If
        End If
    ElseIf PRIT6.Value = "IN" Then
        If priIN6 <> PRI6.Value Then
            priIN6 = PRI6.Value
            If PRI6.Value Then
                UAP.Value = CDbl(UAP.Value) - (CInt(PRIC6.Value) * 0.5)
            Else
                UAP.Value = CDbl(UAP.Value) + (CInt(PRIC6.Value) * 0.5)
            End If
        End If
    ElseIf PRIT6.Value = "RM" Then
        If priRM6 <> PRI6.Value Then
            priRM6 = PRI6.Value
            If PRI6.Value Then
                UAP.Value = CDbl(UAP.Value) - (CInt(PRIC6.Value) * 0.5)
            Else
                UAP.Value = CDbl(UAP.Value) + (CInt(PRIC6.Value) * 0.5)
            End If
        End If
    End If
End Sub

Private Sub PRIC6_Change()
    If PRIT6.Value = "RP" Then
        If PreviousValueRPPRIC6 <> PRIC6.Value Then
            UAP.Value = CDbl(UAP.Value) - (PRIC6.Value - PreviousValueRPPRIC6)
            PreviousValueRPPRIC6 = PRIC6.Value
        End If
    ElseIf PRIT6.Value = "IN" Then
        If PreviousValueINPRIC6 <> PRIC6.Value Then
            UAP.Value = CDbl(UAP.Value) - (0.5 * (PRIC6.Value - PreviousValueINPRIC6))
            PreviousValueINPRIC6 = PRIC6.Value
        End If
    ElseIf PRIT6.Value = "RM" Then
        If PreviousValueRMPRIC6 <> PRIC6.Value Then
            UAP.Value = CDbl(UAP.Value) - (0.5 * (PRIC6.Value - PreviousValueRMPRIC6))
            PreviousValueRMPRIC6 = PRIC6.Value
        End If
    End If
End Sub

Private Sub PRIT7_Change()
    If PRIT7.Value = "RP" Then
        PRI7.Value = priRP7
        PRIC7.Value = PreviousValueRPPRIC7
    ElseIf PRIT7.Value = "IN" Then
        PRI7.Value = priIN7
        PRIC7.Value = PreviousValueINPRIC7
    ElseIf PRIT7.Value = "RM" Then
        PRI7.Value = priRM7
        PRIC7.Value = PreviousValueRMPRIC7
    End If
End Sub

Private Sub PRI7_Click()
    PRIC7.visible = PRI7.Value
    
    If PRIT7.Value = "RP" Then
        If priRP7 <> PRI7.Value Then
            priRP7 = PRI7.Value
            If PRI7.Value Then
                UAP.Value = CDbl(UAP.Value) - (CInt(PRIC7.Value) * 2)
            Else
                UAP.Value = CDbl(UAP.Value) + (CInt(PRIC7.Value) * 2)
            End If
        End If
    ElseIf PRIT7.Value = "IN" Then
        If priIN7 <> PRI7.Value Then
            priIN7 = PRI7.Value
            If PRI7.Value Then
                UAP.Value = CDbl(UAP.Value) - (CInt(PRIC7.Value))
            Else
                UAP.Value = CDbl(UAP.Value) + (CInt(PRIC7.Value))
            End If
        End If
    ElseIf PRIT7.Value = "RM" Then
        If priRM7 <> PRI7.Value Then
            priRM7 = PRI7.Value
            If PRI7.Value Then
                UAP.Value = CDbl(UAP.Value) - (CInt(PRIC7.Value))
            Else
                UAP.Value = CDbl(UAP.Value) + (CInt(PRIC7.Value))
            End If
        End If
    End If
End Sub

Private Sub PRIC7_Change()
    If PRIT7.Value = "RP" Then
        If PreviousValueRPPRIC7 <> PRIC7.Value Then
            UAP.Value = CDbl(UAP.Value) - (2 * (PRIC7.Value - PreviousValueRPPRIC7))
            PreviousValueRPPRIC7 = PRIC7.Value
        End If
    ElseIf PRIT7.Value = "IN" Then
        If PreviousValueINPRIC7 <> PRIC7.Value Then
            UAP.Value = CDbl(UAP.Value) - (PRIC7.Value - PreviousValueINPRIC7)
            PreviousValueINPRIC7 = PRIC7.Value
        End If
    ElseIf PRIT7.Value = "RM" Then
        If PreviousValueRMPRIC7 <> PRIC7.Value Then
            UAP.Value = CDbl(UAP.Value) - (PRIC7.Value - PreviousValueRMPRIC7)
            PreviousValueRMPRIC7 = PRIC7.Value
        End If
    End If
End Sub

Private Sub REC1_Click()
    RECA1.visible = REC1.Value
    RECO1.visible = REC1.Value
    
    If REC1.Value Then
        UASN.Value = CInt(UASN.Value) - (CInt(RECA1.Value) * 6)
    Else
        UASN.Value = CInt(UASN.Value) + (CInt(RECA1.Value) * 6)
    End If
End Sub

Private Sub REC2_Click()
    RECA2.visible = REC2.Value
    RECO2.visible = REC2.Value
    
    If REC2.Value Then
        UASN.Value = CInt(UASN.Value) - (CInt(RECA2.Value) * 3)
    Else
        UASN.Value = CInt(UASN.Value) + (CInt(RECA2.Value) * 3)
    End If
End Sub

Private Sub REC3_Click()
    RECA3.visible = REC3.Value
    RECO3.visible = REC3.Value
    
    If REC3.Value Then
        UASN.Value = CInt(UASN.Value) - (CInt(RECA3.Value) * 6)
    Else
        UASN.Value = CInt(UASN.Value) + (CInt(RECA3.Value) * 6)
    End If
End Sub

Private Sub REC4_Click()
    RECA4.visible = REC4.Value
    RECO4.visible = REC4.Value
    
    If REC4.Value Then
        UASN.Value = CInt(UASN.Value) - (CInt(RECA4.Value) * 6)
    Else
        UASN.Value = CInt(UASN.Value) + (CInt(RECA4.Value) * 6)
    End If
End Sub

Private Sub RPBSL_Click()
    BSL.Value = RPBSL.Value
    RPBSLRFW.visible = BSL.Value
    BSLRFW.visible = BSL.Value
    SLRFWL.visible = BSL.Value
End Sub

Private Sub RPBSLRFW_Change()
    BSLRFW.Value = RPBSLRFW.Value
End Sub

Private Sub RPP_Click()
    MHL.visible = RPP.Value
    MH1.visible = RPP.Value
    MHC1.visible = MH1.Value And RPP.Value
    MHT1.visible = RPP.Value
    MH2.visible = RPP.Value
    MHC2.visible = MH2.Value And RPP.Value
    MHT2.visible = RPP.Value
    RPPSIZE.visible = RPP.Value
    RPPCLASS.visible = RPP.Value
    RPPRFW.visible = RPP.Value
    RDST.visible = RPP.Value
    RDSV.visible = RPP.Value
    DGFV.visible = RPP.Value And RPDG.Value
    DGFL.visible = RPP.Value And RPDG.Value
    RPSWB.visible = RPP.Value And RPDG.Value
    DSL.visible = RPP.Value
    DSV.visible = RPP.Value
    PSL.visible = RPP.Value
    PSV.visible = RPP.Value
    BAFIG.visible = RPP.Value
    FIGL.visible = priBool And RPP.Value
    FIG.visible = priBool And RPP.Value
    TRT.visible = RPP.Value
    RPT.visible = RPP.Value
    RPR.visible = RPP.Value
    TRSL.visible = RPP.Value
    SLM1.visible = RPP.Value And TRSL
    SLM2.visible = RPP.Value And TRSL
    SLM3.Value = SLM1.Value And RPP.Value
    SLM4.Value = SLM2.Value And RPP.Value
    TRSPG.visible = RPP.Value
    RPR.visible = RPP.Value
    RPDG.visible = RPP.Value
    TRSIZE.visible = RPP.Value And (TRT.Value Or RPT.Value)
    TRRPSIZE.visible = RPP.Value And RPT.Value
    STLTBTMBRKT.visible = RPP.Value And TRSL.Value
    COLA.visible = RPP.Value And (TRT.Value Or RPT.Value)
    TRLETTER.visible = RPP.Value And (TRT.Value Or RPT.Value)
    RPBSL.visible = RPP.Value
    RPBSLRFW.visible = RPP.Value And RPBSL.Value
    CETL.visible = RPP.Value
    CETV.visible = RPP.Value
    UAPL.visible = RPP.Value
    UAP.visible = RPP.Value
    FUT1.visible = RPP.Value
    FUT2.visible = RPP.Value
    FUT3.visible = RPP.Value
    PRIT1.visible = RPP.Value
    PRIT2.visible = RPP.Value
    PRIT3.visible = RPP.Value
    PRIT4.visible = RPP.Value
    PRIT5.visible = RPP.Value
    PRIT6.visible = RPP.Value
    PRIT7.visible = RPP.Value
    PRI1.visible = RPP.Value
    PRI2.visible = RPP.Value
    PRI3.visible = RPP.Value
    PRI4.visible = RPP.Value
    PRI5.visible = RPP.Value
    PRI6.visible = RPP.Value
    PRI7.visible = RPP.Value
    PRIC1.visible = PRI1.Value And RPP.Value
    PRIC2.visible = PRI2.Value And RPP.Value
    PRIC3.visible = PRI3.Value And RPP.Value
    PRIC4.visible = PRI4.Value And RPP.Value
    PRIC5.visible = PRI5.Value And RPP.Value
    PRIC6.visible = PRI6.Value And RPP.Value
    PRIC7.visible = PRI7.Value And RPP.Value
    USNL.visible = RPP.Value
    UASN.visible = RPP.Value
    USNL2.visible = RPP.Value
    UASN2.visible = RPP.Value
    SNT1.visible = RPP.Value
    SNT2.visible = RPP.Value
    SNT3.visible = RPP.Value
    SNT4.visible = RPP.Value
    SNT5.visible = RPP.Value
    SNT6.visible = RPP.Value
    SNT7.visible = RPP.Value
    SN1.visible = RPP.Value
    SN2.visible = RPP.Value
    SN3.visible = RPP.Value
    SN4.visible = RPP.Value
    SN5.visible = RPP.Value
    SN6.visible = RPP.Value
    SN7.visible = RPP.Value
    SNC1.visible = SN1.Value And RPP.Value
    SNC2.visible = SN2.Value And RPP.Value
    SNC3.visible = SN3.Value And RPP.Value
    SNC4.visible = SN4.Value And RPP.Value
    SNC5.visible = SN5.Value And RPP.Value
    SNC6.visible = SN6.Value And RPP.Value
    SNC7.visible = SN7.Value And RPP.Value
    FUL.visible = RPP.Value
    FU1.visible = RPP.Value
    FU2.visible = RPP.Value
    FU3.visible = RPP.Value
    FUC1.visible = FU1.Value And RPP.Value
    FUC2.visible = FU2.Value And RPP.Value
    FUC3.visible = FU3.Value And RPP.Value
    NPL.visible = RPP.Value
    RFWL.visible = RPP.Value
    XRL.visible = RPP.Value And RPT.Value
    XSL.visible = RPP.Value And (TRT.Value Or RPT.Value)
    CLL.visible = RPP.Value And (TRT.Value Or RPT.Value)
    XDL.visible = RPP.Value And (TRT.Value Or RPT.Value)
    SLL.visible = RPP.Value And TRSL.Value
    SLRFWL.visible = RPP.Value And RPBSL.Value
End Sub

Private Sub RPPSIZE_Change()
    Call determineRecDSpace
End Sub

Private Sub determineRecDSpace()
    If pole.primaries.count > 0 Then
        If RPPSIZE.Value >= 50 Then
            If TRT.Value Or RPT.Value Then
                RDSV.caption = "[" & Utilities.inchesToFeetInches(((((RPPSIZE * 0.9) - 2) - 32) * 12) - 11 - IIf(Utilities.convertToInches(PSV.Value) > 0, Utilities.convertToInches(PSV.Value), 0)) & "]"
                If Utilities.convertToInches(RDSV.caption) < 60 Then RDSV.caption = "[5'0""]"
            Else
                RDSV.caption = "[" & Utilities.inchesToFeetInches(((((RPPSIZE * 0.9) - 2) - 35) * 12) - 11 - IIf(Utilities.convertToInches(PSV.Value) > 0, Utilities.convertToInches(PSV.Value), 0)) & "]"
                If Utilities.convertToInches(RDSV.caption) < 60 Then RDSV.caption = "[5'0""]"
            End If
        Else
            If TRT.Value Or RPT.Value Then
                If InStr(TRLETTER.Value, "C") = 1 Then
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
    RSLD.visible = RSL.Value
    RSLRFW.visible = RSL.Value
    RFWL4.visible = RSL.Value
End Sub

Private Sub RSPANS_Change()
    
    If RSPANS.Value >= 1 Then
        OWSPAN1.visible = True
        OW1L1.visible = True
    Else
        OWSPAN1.visible = False
        OW1L1.visible = False
    End If
    
    If RSPANS.Value >= 2 Then
        OWSPAN2.visible = True
        OW2L1.visible = True
    Else
        OWSPAN2.visible = False
        OW2L1.visible = False
    End If
    
    If RSPANS.Value >= 3 Then
        OWSPAN3.visible = True
        OW3L1.visible = True
    Else
        OWSPAN3.visible = False
        OW3L1.visible = False
    End If
    
    If RSPANS.Value >= 4 Then
        OWSPAN4.visible = True
        OW4L1.visible = True
    Else
        OWSPAN4.visible = False
        OW4L1.visible = False
    End If
    
    If RSPANS.Value >= 5 Then
        OWSPAN5.visible = True
        OW5L1.visible = True
    Else
        OWSPAN5.visible = False
        OW5L1.visible = False
    End If
    
    OWSIZE1.visible = OWSPAN1.Value <> "" And OWSPAN1.visible
    OWNEWSIZE1.visible = OWSPAN1.Value <> "" And OWSPAN1.visible
    OWLENGTH1.visible = OWSPAN1.Value <> "" And OWSPAN1.visible
    OW1L2.visible = OWSPAN1.Value <> "" And OWSPAN1.visible
    OW1L3.visible = OWSPAN1.Value <> "" And OWSPAN1.visible
    OW1L4.visible = OWSPAN1.Value <> "" And OWSPAN1.visible
    
    OWSIZE2.visible = OWSPAN2.Value <> "" And OWSPAN2.visible
    OWNEWSIZE2.visible = OWSPAN2.Value <> "" And OWSPAN2.visible
    OWLENGTH2.visible = OWSPAN2.Value <> "" And OWSPAN2.visible
    OW2L2.visible = OWSPAN2.Value <> "" And OWSPAN2.visible
    OW2L3.visible = OWSPAN2.Value <> "" And OWSPAN2.visible
    OW2L4.visible = OWSPAN2.Value <> "" And OWSPAN2.visible
    
    OWSIZE3.visible = OWSPAN3.Value <> "" And OWSPAN3.visible
    OWNEWSIZE3.visible = OWSPAN3.Value <> "" And OWSPAN3.visible
    OWLENGTH3.visible = OWSPAN3.Value <> "" And OWSPAN3.visible
    OW3L2.visible = OWSPAN3.Value <> "" And OWSPAN3.visible
    OW3L3.visible = OWSPAN3.Value <> "" And OWSPAN3.visible
    OW3L4.visible = OWSPAN3.Value <> "" And OWSPAN3.visible
    
    OWSIZE4.visible = OWSPAN4.Value <> "" And OWSPAN4.visible
    OWNEWSIZE4.visible = OWSPAN4.Value <> "" And OWSPAN4.visible
    OWLENGTH4.visible = OWSPAN4.Value <> "" And OWSPAN4.visible
    OW4L2.visible = OWSPAN4.Value <> "" And OWSPAN4.visible
    OW4L3.visible = OWSPAN4.Value <> "" And OWSPAN4.visible
    OW4L4.visible = OWSPAN4.Value <> "" And OWSPAN4.visible
    
    OWSIZE5.visible = OWSPAN5.Value <> "" And OWSPAN5.visible
    OWNEWSIZE5.visible = OWSPAN5.Value <> "" And OWSPAN5.visible
    OWLENGTH5.visible = OWSPAN5.Value <> "" And OWSPAN5.visible
    OW5L2.visible = OWSPAN5.Value <> "" And OWSPAN5.visible
    OW5L3.visible = OWSPAN5.Value <> "" And OWSPAN5.visible
    OW5L4.visible = OWSPAN5.Value <> "" And OWSPAN5.visible
    
End Sub

Private Sub SLDL_Click()
    SLDLRFW.visible = SLDL.Value
    RFWL3.visible = SLDL.Value
End Sub

Private Sub SLM1_Click()
    If RPP.Value Then SLM3.Value = SLM1.Value
End Sub

Private Sub SLM2_Click()
    If RPP.Value Then SLM4.Value = SLM2.Value
End Sub

Private Sub SLM3_Click()
    SLM1.Value = SLM3.Value
End Sub

Private Sub SLM4_Click()
    SLM2.Value = SLM4.Value
End Sub

Private Sub TRSL_Click()
    SLL.visible = RPP.Value And TRSL.Value
    STLTBTMBRKT.visible = RPP.Value And TRSL.Value
End Sub

Private Sub TSDL_Click()
    TSDLRFW.visible = TSDL.Value
    RFWL1.visible = TSDL.Value
End Sub

Private Sub CommandButton1_Click()

    Set figuresUsed = New Scripting.Dictionary

    CrewNotes = ""
    installNotes = ""
    removeNotes = ""
    replaceNotes = ""
    transferNotes = ""
    notes = ""
    replace11kCount = 0
    replace20kCount = 0
    Set commTransfers = New Collection

    Call generateDownguysSection
    If RPP.Value Then Call generateReplacePoleSection
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
        answer = MsgBox(warnings & vbLf & "Do you to generate crew notes even with the warnings? Select no or cancel to fix the issues first.", vbYesNoCancel + vbQuestion, "Confirm")
        If answer = vbNo Or answer = vbCancel Then Exit Sub
    End If
    
    If CrewNotes <> "" Then
        CrewNotes = Left(CrewNotes, Len(CrewNotes) - 1)

        Dim DataObj As DataObject: Set DataObj = New DataObject
        DataObj.SetText CrewNotes
        DataObj.PutInClipboard
    
        If Trim(pds.Range("ALTONE")) = "" Then
            pds.Range("ALTONE").Value = CrewNotes
            Call Figures.getSheetFigures(pds)
        End If
    
        MsgBox CrewNotes
    Else
        MsgBox "No work selected"
    End If
    
End Sub

Private Function findErrors() As String
    Errors = ""
    
    If RPP.Value And PSV.Value = "" And BAFIG.Value Then
        Errors = Errors & "Error: P space shouldn't be blank if there's a buckarm on the pole." & vbLf
    End If
    
    If RPP.Value And comms.count > CInt(CETV.Value) And Not RPPNOTE1.Value Then
        Errors = Errors & "Error: Not all comms accounted for on pole replace, either have CE transfer more comms or add top pole note." & vbLf
    End If
    
    If RPP.Value And priBool And DSV.Value = "" Then
        Errors = Errors & "Error: D space shouldn't be blank if there's primary on the pole." & vbLf
    End If
    
    If Not RPPNOTE2.Value And RPP.Value And (TRT.Value Or RPT) Then
        Errors = Errors & "Error: Outage note not selected, outage needs to be selected when replacing a pole with a transformer." & vbLf
    ElseIf RPP.Value And (ohsCount > 0) And Not RPPNOTE2.Value Then
        Errors = Errors & "Error: Outage note not selected, outage needs to be selected when replacing services." & vbLf
    ElseIf Not RPPNOTE2.Value And (REC1.Value Or REC2.Value Or REC3.Value Or REC4.Value Or (OWSPAN1.Value <> "" And OWSPAN1.visible) Or (OWSPAN2.Value <> "" And OWSPAN2.visible) Or (OWSPAN3.Value <> "" And OWSPAN3.visible) Or (OWSPAN4.Value <> "" And OWSPAN4.visible) Or (OWSPAN5.Value <> "" And OWSPAN5.visible)) Then
        Errors = Errors & "Error: Outage note not selected, outage needs to be selected when reconductoring." & vbLf
    End If
    
    If RPP.Value And (TRT.Value Or RPT.Value) And COLA.Value = "" Then
        Errors = Errors & "Error: If transfering or replacing transformer, CO/LA configuration needs to be chosen. Choose Other if nothing fits." & vbLf
    End If
    
    If RPP.Value And TRSL.Value And STLTBTMBRKT.Value = "" Then
        Errors = Errors & "Error: If transfering streetlight, streetlight attach height needs to be filled in." & vbLf
    End If
    
    findErrors = Errors
End Function

Private Function findWarnings() As String
    warnings = ""
    
    If TSDL.Value And Not RPPNOTE2.Value Then
        warnings = warnings & "Warning: Outage required not selected when trimming driploop." & vbLf
    ElseIf Not RPPNOTE2.Value And RPP.Value Then
        warnings = warnings & "Warning: Outage required not selected replacing pole." & vbLf
    End If
    
    If RPP.Value And PSV.Value = "" And Not BAFIG.Value Then
        warnings = warnings & "Warning: P space is blank, make sure there's no P space needed on this pole." & vbLf
    End If
    
    If RPP.Value And priBool And Not (FU1.Value Or FU2.Value Or FU3.Value) Then
        warnings = warnings & "Warning: Make sure the pole doesn't need any framing hardware replaced." & vbLf
    End If
    
    If RPP.Value And UAP <> 0 Then
        warnings = warnings & "Warning: Unaccounted Primary should be 0 to account for all primary spans." & vbLf
    End If
    
    If RPP.Value And UASN <> 0 Then
        warnings = warnings & "Warning: Unaccounted OW/Sec/Neut should be 0 to account for all OW/Sec/Neut spans." & vbLf
    End If
    
    If RPP.Value And ohsCount = 0 And (TRT.Value Or RPT.Value Or secBool Or owBool) And Not RPPNOTE2.Value Then
        warnings = warnings & "Warning: Outage note not selected, check if the transformer/secondaries power a house or streetlight." & vbLf
    End If
    
    If RPP.Value And Not RPPNOTE3.Value And Not RPPNOTE4.Value Then
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
    
    replacePoleSection = replacePoleSection & poleHeight & pds.Range("CLASS") & "/" & RPPSIZE.Value & "-" & RPPCLASS.Value & " " & Replace(RPPRFW.Value, "OTHER", "[REASON FOR WORK]") & vbLf
    If Trim(PSV.Value) <> "N/A" And Trim(PSV.Value) <> "" Then replacePoleSection = replacePoleSection & "P = " & Trim(PSV.Value) & vbLf
    If RDSV.caption <> "" Then
        If Trim(DSV.Value) <> "N/A" And Trim(DSV.Value) <> "" And Trim(DSV.Value) <> "[X']" Then
            replacePoleSection = replacePoleSection & "D = " & Trim(DSV.Value) & vbLf
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
    If BAFIG.Value Then replacePoleSection = replacePoleSection & "FIGURE 23-229-1" & vbLf
    
    
    Dim insulators As String: insulators = "[XP-FG]"
    If pole.primaries.count = 0 Then insulators = "[XP]"
    If RPDG.Value Then
        If dg11KCount - replace11kCount > 0 Or dg20KCount - replace20kCount > 0 Then
            If dg11KCount - replace11kCount > 0 Then
                replacePoleSection = replacePoleSection & IIf(dg11KCount - replace11kCount > 1, "(" & dg11KCount - replace11kCount & ")", "") & "11K-" & insulators & " GUY STRAND" & IIf(RPPSIZE.Value > poleHeight, " AND EXTEND TO NEW HEIGHT", "") & vbLf
            End If
            If dg20KCount - replace20kCount > 0 Then
                replacePoleSection = replacePoleSection & IIf(dg20KCount - replace20kCount > 1, "(" & dg20KCount - replace20kCount & ")", "") & "20K-" & insulators & " GUY STRAND" & IIf(RPPSIZE.Value > poleHeight, " AND EXTEND TO NEW HEIGHT", "") & vbLf
            End If
            If Me.Controls("DGFV").Value = "" Or Me.Controls("DGFV" & i).Value = "OTHER" Then
                replacePoleSection = replacePoleSection & "FIGURE [INSERT DOWNGUY FIGURE]" & vbLf
            Else
                replacePoleSection = replacePoleSection & Split(Me.Controls("DGFV").Value, ",")(0) & vbLf
            End If
        End If
        If RPSWB.Value Then
            replacePoleSection = replacePoleSection & "SIDEWALK BRACE" & vbLf
            replacePoleSection = replacePoleSection & "FIGURE 22-450-1" & vbLf
        End If
    End If
    
    If RPT.Value Then
        If UCase(TRSIZE.Value) <> UCase(TRRPSIZE.Value) Then
            replacePoleSection = replacePoleSection & UCase(TRSIZE.Value) & "/" & UCase(TRRPSIZE.Value) & " TRANSFORMER"
        ElseIf UCase(TRSIZE.Value) = UCase(TRRPSIZE.Value) Then
            replacePoleSection = replacePoleSection & UCase(TRRPSIZE.Value) & " TRANSFORMER"
        End If
        If RPPSIZE.Value >= 50 Then replacePoleSection = replacePoleSection & " [@35'0""]"
        replacePoleSection = replacePoleSection & vbLf
        If COLA.ListIndex = 0 Or COLA.Value = "" Then
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
        If TRLETTER.Value = "3 PHASE" Then
            replacePoleSection = replacePoleSection & "FIGURE 26-301-1" & vbLf
        Else
            replacePoleSection = replacePoleSection & "FIGURE 26-101-1 DETAIL " & IIf(Trim(TRLETTER.text) = "", "[A/B/C/D]", Left(TRLETTER.Value, 1)) & vbLf
        End If
    End If
    
    If RPR.Value Then
        If TRT.Value Or RPT.Value Then
            replacePoleSection = replacePoleSection & "[X]/C-[XX] AL SERVICE [OR SECONDARY] RISER & EXTEND MOLDING TO 12"" BELOW BOTTOM OF TRANSFORMER" & vbLf
            replacePoleSection = replacePoleSection & "FIGURE 63-20-1" & vbLf
        Else
            replacePoleSection = replacePoleSection & "[X]/C-[XX]AL SERVICE [OR SECONDARY] RISER & EXTEND MOLDING TO 4"" BELOW [OR ABOVE?] SECONDARY ATTACH HEIGHT" & vbLf
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
    
    Call extractFromRPINRMDicts(replacePoleSection, rpDict, inDict, rmDict)
    
    If priRP1 Or priRP2 Or snRP1 Or snRP2 Or mhRP1 Then
        spinCount = 0
        If priRP1 Then spinCount = spinCount + PreviousValueRPPRIC1
        If priRP2 Then spinCount = spinCount + (PreviousValueRPPRIC2 * 2)
        If snRP1 Then spinCount = spinCount + PreviousValueRPSNC1
        If snRP2 Then spinCount = spinCount + (PreviousValueRPSNC2 * 2)
        If mhRP1 Then spinCount = spinCount + previousValueRPMHC1
        
        rpDict("SPINS") = spinCount
    End If
    
    If priIN1 Or priIN2 Or snIN1 Or snIN2 Or mhIN1 Then
        spinCount = 0
        If priIN1 Then spinCount = spinCount + PreviousValueINPRIC1
        If priIN2 Then spinCount = spinCount + (PreviousValueINPRIC2 * 2)
        If snIN1 Then spinCount = spinCount + PreviousValueINSNC1
        If snIN2 Then spinCount = spinCount + (PreviousValueINSNC2 * 2)
        If mhIN1 Then spinCount = spinCount + previousValueINMHC1
        
        inDict("SPINS") = spinCount
    End If
    
    If priRM1 Or priRM2 Or snRM1 Or snRM2 Or mhRM1 Then
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
        conductors = conductors & "SERVICE" & IIf(ohsCount > 1, "S", "")
        Dim serviceSizes As Scripting.Dictionary: Set serviceSizes = New Scripting.Dictionary
        For Each service In pole.services
            For Each midspan In service.midspans
                If Not serviceSizes.Exists(service.size) Then serviceSizes(service.size) = 0
                serviceSizes(service.size) = serviceSizes(service.size) + 1
            Next midspan
        Next service
        
        For Each size In serviceSizes
            If serviceSizes(size) > 1 Then
                replacePoleSection = replacePoleSection & "(" & serviceSizes(size) & ")" & Replace(size, " ", "") & " SERVICE DEADEND" & vbLf
            Else
                replacePoleSection = replacePoleSection & Replace(size, " ", "") & " SERVICE DEADEND" & vbLf
            End If
        Next size
        replacePoleSection = replacePoleSection & "FIGURE 23-302-1 DETAIL A"
    End If
    
    If conductors <> "" Then transferNotes = transferNotes & conductors & vbLf
    If priBool = False Then transferNotes = transferNotes & "@11"" FROM TOP OF POLE" & vbLf
    
    i = 0
    For Each key In comms
        If i >= CInt(CETV.Value) Then Exit For
        For j = 1 To 8
            If UCase(Trim(pds.Range("COMM" & j).Value)) = UCase(Trim(key)) Then
                If pds.Range("COMM" & j).offset(2, 0).offset(0, 1).Value <> "" Then
                    transferNotes = transferNotes & key & " @" & pds.Range("COMM" & j).offset(2, 0).offset(0, 1).Value & vbLf
                Else
                    transferNotes = transferNotes & key & " @" & pds.Range("COMM" & j).offset(2, 0).Value & vbLf
                End If
                commTransfers.Add key
                Exit For
            End If
        Next j
        i = i + 1
    Next key
    
    If TRSL.Value Then
        transferNotes = transferNotes & "STREETLIGHT BOTTOM BRACKET @" & STLTBTMBRKT.Value & vbLf
        transferNotes = transferNotes & "FIGURE 43-116-1" & vbLf
    End If
    If TRSPG.Value Then
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
    If TRT.Value Then
        transferNotes = transferNotes & UCase(TRSIZE.Value) & " TRANSFORMER"
        If RPPSIZE.Value >= 50 Then transferNotes = transferNotes & " [@35'0""]"
        transferNotes = transferNotes & vbLf
        If COLA.ListIndex = 0 Or COLA.Value = "" Then
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
        If TRLETTER.Value = "3 PHASE" Then
            transferNotes = transferNotes & "FIGURE 26-301-1" & vbLf
        Else
            transferNotes = transferNotes & "FIGURE 26-101-1 DETAIL " & IIf(Trim(TRLETTER.text) = "", "[A/B/C/D]", Left(TRLETTER.Value, 1)) & vbLf
        End If
    End If
    
    transferNotes = transferNotes & "CE ID TAG" & vbLf
    
    replaceNotes = replacePoleSection & replaceNotes
End Sub

Private Sub extractFromRPINRMDicts(ByRef replacePoleSection As String, rpDict As Scripting.Dictionary, inDict As Scripting.Dictionary, rmDict As Scripting.Dictionary)
    For Each key In inDict
        If rmDict.Exists(key) Then
            If inDict(key) >= rmDict(key) Then
                If Not rpDict.Exists(key) Then rpDict(key) = 0
                rpDict(key) = rpDict(key) + rmDict(key)
                inDict(key) = inDict(key) - rmDict(key)
                rmDict(key) = rmDict(key) - rmDict(key)
            Else
                If Not rpDict.Exists(key) Then rpDict(key) = 0
                rpDict(key) = rpDict(key) + inDict(key)
                rmDict(key) = rmDict(key) - inDict(key)
                inDict(key) = inDict(key) - inDict(key)
            End If
            If inDict(key) = 0 Then Call inDict.Remove(key)
            If rmDict(key) = 0 Then Call rmDict.Remove(key)
        End If
    Next key
    
    For Each key In rpDict
        If rmDict.Exists(key) And inDict.count = 0 Then
            replacePoleSection = replacePoleSection & "(" & rpDict(key) + rmDict(key) & ")" & key & "/" & IIf(rpDict(key) <> 1, "(" & rpDict(key) & ")", "") & key & vbLf
            replacePoleSection = replacePoleSection & secondaryFigures(key)
            Call rmDict.Remove(key)
        ElseIf inDict.Exists(key) And rmDict.count = 0 Then
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
        If Not figuresUsed.Exists("[FIGURE 23-301-2 or FIGURE 23-303-1 DETAIL A (FOR OPEN WIRE)]") Then
            figuresUsed.Add "[FIGURE 23-301-2 or FIGURE 23-303-1 DETAIL A (FOR OPEN WIRE)]", Nothing
            figure = "[FIGURE 23-301-2 or FIGURE 23-303-1 DETAIL A (FOR OPEN WIRE)]" & vbLf
        End If
    ElseIf hardware = "SECONDARY DEADEND" Then
        If Not figuresUsed.Exists("FIGURE 23-302-1 DETAIL [A/B]") Then
            figuresUsed.Add "FIGURE 23-302-1 DETAIL [A/B]", Nothing
            figure = "FIGURE 23-302-1 DETAIL [A/B]" & vbLf
        End If
    ElseIf hardware = "OPEN WIRE DEADEND" Then
        If Not figuresUsed.Exists("FIGURE 23-303-1 DETAIL B") Then
            figuresUsed.Add "FIGURE 23-303-1 DETAIL B", Nothing
            figure = "FIGURE 23-303-1 DETAIL B" & vbLf
        End If
    ElseIf hardware = "TANGENT CLAMP" Then
        If Not figuresUsed.Exists("FIGURE 23-301-1") Then
            figuresUsed.Add "FIGURE 23-301-1", Nothing
            figure = "FIGURE 23-301-1" & vbLf
        End If
    ElseIf hardware = "SECONDARY AWAC DEADEND" Then
        If Not figuresUsed.Exists("FIGURE 23-302-1 DETAIL B") Then
            figuresUsed.Add "FIGURE 23-302-1 DETAIL B", Nothing
            figure = "FIGURE 23-302-1 DETAIL B" & vbLf
        End If
    End If
    
    secondaryFigures = figure
End Function

Private Sub generateReconductorSection()
    If OWSPAN1.Value <> "" And OWSPAN1.visible Then
        direction = getDirection(OWSPAN1)
        replaceNotes = replaceNotes & OWLENGTH1.Value & " " & OWSIZE1.Value & " OPEN WIRE / " & OWNEWSIZE1.Value & " SECONDARY TO THE " & direction & vbLf
    End If
    
    If OWSPAN2.Value <> "" And OWSPAN2.visible Then
        direction = getDirection(OWSPAN2)
        replaceNotes = replaceNotes & OWLENGTH2.Value & " " & OWSIZE2.Value & " OPEN WIRE / " & OWNEWSIZE2.Value & " SECONDARY TO THE " & direction & vbLf
    End If
    
    If OWSPAN3.Value <> "" And OWSPAN3.visible Then
        direction = getDirection(OWSPAN3)
        replaceNotes = replaceNotes & OWLENGTH3.Value & " " & OWSIZE3.Value & " OPEN WIRE / " & OWNEWSIZE3.Value & " SECONDARY TO THE " & direction & vbLf
    End If
    
    If OWSPAN4.Value <> "" And OWSPAN4.visible Then
        direction = getDirection(OWSPAN4)
        replaceNotes = replaceNotes & OWLENGTH4.Value & " " & OWSIZE4.Value & " OPEN WIRE / " & OWNEWSIZE4.Value & " SECONDARY TO THE " & direction & vbLf
    End If
    
    If OWSPAN5.Value <> "" And OWSPAN5.visible Then
        direction = getDirection(OWSPAN5)
        replaceNotes = replaceNotes & OWLENGTH5.Value & " " & OWSIZE5.Value & " OPEN WIRE / " & OWNEWSIZE5.Value & " SECONDARY TO THE " & direction & vbLf
    End If
    
    If REC1.Value Then
        replaceNotes = replaceNotes & "(" & 3 * RECA1.Value & ")WR/" & IIf(RECA1.Value > 1, "(" & RECA1.Value & ")", "")
        If RECO1.Value Then
            replaceNotes = replaceNotes & "TANGENT CLAMP" & vbLf
            replaceNotes = replaceNotes & secondaryFigures("TANGENT CLAMP")
        Else
            replaceNotes = replaceNotes & "WR" & vbLf
            replaceNotes = replaceNotes & secondaryFigures("WR")
        End If
    End If
    
    If REC2.Value Then
        replaceNotes = replaceNotes & "(" & 3 * RECA2.Value & ")OPEN WIRE DEADEND/" & IIf(RECA2.Value > 1, "(" & RECA2.Value & ")", "")
        If RECO2.Value Then
            replaceNotes = replaceNotes & "SECONDARY AWAC DEADEND" & vbLf
            replaceNotes = replaceNotes & secondaryFigures("SECONDARY AWAC DEADEND")
        Else
            replaceNotes = replaceNotes & "SECONDARY DEADEND" & vbLf
            replaceNotes = replaceNotes & "FIGURE 23-302-1 DETAIL A" & secondaryFigures("SECONDARY DEADEND")
        End If
    End If
    
    If REC3.Value Then
        replaceNotes = replaceNotes & "(" & 3 * RECA3.Value & ")WR/" & "(" & 3 * RECA3.Value & ")OPEN WIRE DEADEND" & IIf(RECA3.Value > 1, "(" & RECA3.Value & ")", "")
        If RECO3.Value Then
            replaceNotes = replaceNotes & "+ SECONDARY AWAC DEADEND" & vbLf
            replaceNotes = replaceNotes & secondaryFigures("SECONDARY AWAC DEADEND")
        Else
            replaceNotes = replaceNotes & "+ SECONDARY DEADEND" & vbLf
            replaceNotes = replaceNotes & "FIGURE 23-302-1 DETAIL A" & vbLf
        End If
    End If
    
    If REC4.Value Then
        replaceNotes = replaceNotes & "(" & 6 * RECA4.Value & ")OPEN WIRE DEADEND/" & IIf(RECA4.Value > 1, "(" & RECA4.Value & ")", "")
        If RECO4.Value Then
            replaceNotes = replaceNotes & "TANGENT CLAMP" & vbLf
            replaceNotes = replaceNotes & secondaryFigures("TANGENT CLAMP")
        Else
            replaceNotes = replaceNotes & "WR" & vbLf
            replaceNotes = replaceNotes & secondaryFigures("WR")
        End If
    End If
    
    If REC1.Value Or REC2.Value Or REC3.Value Or REC4.Value Or (OWSPAN1.Value <> "" And OWSPAN1.visible) Or (OWSPAN2.Value <> "" And OWSPAN2.visible) Or (OWSPAN3.Value <> "" And OWSPAN3.visible) Or (OWSPAN4.Value <> "" And OWSPAN4.visible) Or (OWSPAN5.Value <> "" And OWSPAN5.visible) Then
        replaceNotes = replaceNotes & "TO REPLACE OPEN WIRE" & vbLf
    End If
    
End Sub

Private Function getDirection(span As String) As String

    StartPos = InStr(span, "(")
    EndPos = InStr(StartPos + 1, span, ")")
    
    If StartPos > 0 And EndPos > 0 Then
        degrees = CInt(Mid(span, StartPos + 1, EndPos - StartPos - 1))
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
    If BSL.Value Then
        installNotes = installNotes & "BOND STREETLIGHT BRACKET TO NEUTRAL " & Replace(BSLRFW.Value, "OTHER", "[REASON FOR WORK]") & vbLf
        installNotes = installNotes & "FIGURE 42-105-1" & vbLf
    End If
    
    If SLM3.Value Then
        replaceNotes = replaceNotes & "STREETLIGHT MOLDING" & vbLf
    End If
    
    If SLM4.Value Then
        installNotes = installNotes & "STREETLIGHT MOLDING" & vbLf
    End If
    
    If TSDL.Value Then
        transferNotes = transferNotes & "TRIM SECONDARY DRIPLOOP TO 4"" BELOW SECONDARY ATTACHMENT HEIGHT " & Replace(TSDLRFW.Value, "OTHER", "[REASON FOR WORK]") & vbLf
    End If
    
    If RSL.Value Then
        transferNotes = transferNotes & "RAISE STREETLIGHT BRACKET " & RSLD.Value & " " & Replace(RSLRFW.Value, "OTHER", "[REASON FOR WORK]") & vbLf
        transferNotes = transferNotes & "FIGURE 43-116-1" & vbLf
    End If
    
    If SLDL.Value Then
        transferNotes = transferNotes & "TRIM STREETLIGHT DRIP LOOP TO 1"" BELOW STREETLIGHT BRACKET" & vbLf
        transferNotes = transferNotes & Replace(SLDLRFW.Value, "OTHER", "[REASON FOR WORK]") & vbLf
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

    If RPPNOTE1.Value Then
        notes = notes & "NOTE: TOP POLE ABOVE COMMS, TOPPED POLE TO BE REMOVED AT A LATER DATE" & vbLf
    End If
    If RPPNOTE2.Value Then
        notes = notes & "NOTE: OUTAGE REQUIRED, SCHEDULING TO NOTIFY" & vbLf
    End If
    If RPPNOTE3.Value Then
        notes = notes & "NOTE: TREE WORK REQUIRED" & vbLf
    End If
    If RPPNOTE4.Value Then
        notes = notes & "NOTE: BRUSH WORK REQUIRED" & vbLf
    End If
    If RPPNOTE5.Value Then
        notes = notes & "NOTE: POLE LOCATED IN ACTIVE FARMFIELD" & vbLf
    End If
    If RPPNOTE6.Value Then
        notes = notes & "NOTE: FENCE REMOVAL REQUIRED" & vbLf
    End If
    If RPPNOTE7.Value Then
        notes = notes & "NOTE: BACKYARD MACHINE REQUIRED" & vbLf
    End If
    If RPPNOTE8.Value Then
        notes = notes & "NOTE: [INSERT POLE ACCESS NOTES]" & vbLf
    End If
End Sub

Private Sub generateDownguysSection()
    Dim insulators As String: insulators = "[XP-FG]"
    If pole.primaries.count = 0 Then insulators = "[XP]"
    For i = 1 To 3
        If Me.Controls("IA" & i & "1").Value <> "" And Me.Controls("IA" & i & "1").visible Then
            If Me.Controls("IA" & i & "1").Value = "RS" Then
                installNotes = installNotes & "11K-" & insulators & "-" & IIf(Me.Controls("IA" & i & "3").Value = "", [lead], Me.Controls("IA" & i & "3").Value) & "-RS " & IIf(Me.Controls("IA" & i & "4").Value = "", "[DIRECTION]", Me.Controls("IA" & i & "4").Value) & vbLf
            ElseIf Me.Controls("IA" & i & "1").Value = "RT" Then
                If Me.Controls("IA" & i & "2").Value = "20K" Then
                    installNotes = installNotes & "20K-" & insulators & "-" & IIf(Me.Controls("IA" & i & "3").Value = "", [lead], Me.Controls("IA" & i & "3").Value) & "-RT " & IIf(Me.Controls("IA" & i & "4").Value = "", "[DIRECTION]", Me.Controls("IA" & i & "4").Value) & vbLf
                ElseIf Me.Controls("IA" & i & "2").Value = "(2)11K" Then
                    installNotes = installNotes & "11K-" & insulators & "-" & IIf(Me.Controls("IA" & i & "3").Value = "", [lead], Me.Controls("IA" & i & "3").Value) & "-RT " & IIf(Me.Controls("IA" & i & "4").Value = "", "[DIRECTION]", Me.Controls("IA" & i & "4").Value) & vbLf
                    installNotes = installNotes & "11K-" & insulators & vbLf
                End If
            ElseIf Me.Controls("IA" & i & "1").Value = "STE" Then
                If Me.Controls("IA" & i & "2").Value = "20K + 11K" Then
                    installNotes = installNotes & "20K-" & insulators & "-" & IIf(Me.Controls("IA" & i & "3").Value = "", [lead], Me.Controls("IA" & i & "3").Value) & "-STE " & IIf(Me.Controls("IA" & i & "4").Value = "", "[DIRECTION]", Me.Controls("IA" & i & "4").Value) & vbLf
                    installNotes = installNotes & "11K-" & insulators & vbLf
                ElseIf Me.Controls("IA" & i & "2").Value = "(3)11K" Then
                    installNotes = installNotes & "11K-" & insulators & "-" & IIf(Me.Controls("IA" & i & "3").Value = "", [lead], Me.Controls("IA" & i & "3").Value) & "-STE " & IIf(Me.Controls("IA" & i & "4").Value = "", "[DIRECTION]", Me.Controls("IA" & i & "4").Value) & vbLf
                    installNotes = installNotes & "11K-" & insulators & vbLf
                    installNotes = installNotes & "11K-" & insulators & vbLf
                ElseIf Me.Controls("IA" & i & "2").Value = "(2)20K" Then
                    installNotes = installNotes & "20K-" & insulators & "-" & IIf(Me.Controls("IA" & i & "3").Value = "", [lead], Me.Controls("IA" & i & "3").Value) & "-STE " & IIf(Me.Controls("IA" & i & "4").Value = "", "[DIRECTION]", Me.Controls("IA" & i & "4").Value) & vbLf
                    installNotes = installNotes & "20K-" & insulators & vbLf
                End If
            End If
            
            If Me.Controls("IA" & i & "5").Value Then
                installNotes = installNotes & "SIDEWALK BRACE" & vbLf
                installNotes = installNotes & "FIGURE 22-450-1" & vbLf
            End If

            If Me.Controls("IGRFW" & i).Value = "" Or Me.Controls("IGRFW" & i).Value = "OTHER" Then
                installNotes = installNotes & "[REASON FOR WORK]" & vbLf
            Else
                installNotes = installNotes & Me.Controls("IGRFW" & i) & vbLf
            End If
            
            If Me.Controls("IGFIG" & i).Value = "" Or Me.Controls("IGFIG" & i).Value = "OTHER" Then
                installNotes = installNotes & "FIGURE [INSERT DOWNGUY FIGURE]" & vbLf
            Else
                installNotes = installNotes & Split(Me.Controls("IGFIG" & i).Value, ",")(0) & vbLf
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
        If Me.Controls("EA" & i & "1").Value <> Me.Controls("RA" & i & "1").Value Or Me.Controls("EA" & i & "2").Value <> Me.Controls("RA" & i & "2").Value Or Me.Controls("EA" & i & "3").Value <> Me.Controls("RA" & i & "3").Value Then
            If Me.Controls("EA" & i & "1").Value <> "" Then
                If Me.Controls("EA" & i & "1").Value = "RS" Then
                    eLine1 = "11K-" & insulators & "-" & IIf(Me.Controls("EA" & i & "3").Value = "", [lead], Me.Controls("EA" & i & "3").Value) & "-RS"
                    replace11kCount = replace11kCount + 1
                ElseIf Me.Controls("EA" & i & "1").Value = "RT" Then
                    If Me.Controls("EA" & i & "2").Value = "20K" Then
                        eLine1 = "20K-" & insulators & "-" & IIf(Me.Controls("EA" & i & "3").Value = "", [lead], Me.Controls("EA" & i & "3").Value) & "-RT"
                        replace20kCount = replace20kCount + 1
                    ElseIf Me.Controls("EA" & i & "2").Value = "(2)11K" Then
                        eLine1 = "11K-" & insulators & "-" & IIf(Me.Controls("EA" & i & "3").Value = "", [lead], Me.Controls("EA" & i & "3").Value) & "-RT"
                        eLine2 = "11K-" & insulators
                        replace11kCount = replace11kCount + 2
                    End If
                ElseIf Me.Controls("EA" & i & "1").Value = "STE" Then
                    If Me.Controls("EA" & i & "2").Value = "20K + 11K" Then
                        eLine1 = "20K-" & insulators & "-" & IIf(Me.Controls("EA" & i & "3").Value = "", [lead], Me.Controls("EA" & i & "3").Value) & "-STE"
                        eLine2 = "11K-" & insulators
                        replace20kCount = replace20kCount + 1
                        replace11kCount = replace11kCount + 1
                    ElseIf Me.Controls("EA" & i & "2").Value = "(3)11K" Then
                        eLine1 = "11K-" & insulators & "-" & IIf(Me.Controls("EA" & i & "3").Value = "", [lead], Me.Controls("EA" & i & "3").Value) & "-STE"
                        eLine2 = "11K-" & insulators
                        eLine3 = "11K-" & insulators
                        replace11kCount = replace11kCount + 3
                    ElseIf Me.Controls("EA" & i & "2").Value = "(2)20K" Then
                        eLine1 = "20K-" & insulators & "-" & IIf(Me.Controls("EA" & i & "3").Value = "", [lead], Me.Controls("EA" & i & "3").Value) & "-STE"
                        eLine2 = "20K-" & insulators
                        replace20kCount = replace20kCount + 2
                    End If
                End If
            End If
            If Me.Controls("RA" & i & "1").Value <> "" Then
                If Me.Controls("RA" & i & "1").Value = "RS" Then
                    rLine1 = "11K-" & insulators & "-" & IIf(Me.Controls("RA" & i & "3").Value = "", [lead], Me.Controls("RA" & i & "3").Value) & "-RS"
                ElseIf Me.Controls("RA" & i & "1").Value = "RT" Then
                    If Me.Controls("RA" & i & "2").Value = "20K" Then
                        rLine1 = "20K-" & insulators & "-" & IIf(Me.Controls("RA" & i & "3").Value = "", [lead], Me.Controls("RA" & i & "3").Value) & "-RT"
                    ElseIf Me.Controls("RA" & i & "2").Value = "(2)11K" Then
                        rLine1 = "11K-" & insulators & "-" & IIf(Me.Controls("RA" & i & "3").Value = "", [lead], Me.Controls("RA" & i & "3").Value) & "-RT"
                        rLine2 = "11K-" & insulators
                    End If
                ElseIf Me.Controls("RA" & i & "1").Value = "STE" Then
                    If Me.Controls("RA" & i & "2").Value = "20K + 11K" Then
                        rLine1 = "20K-" & insulators & "-" & IIf(Me.Controls("RA" & i & "3").Value = "", [lead], Me.Controls("RA" & i & "3").Value) & "-STE"
                        rLine2 = "11K-" & insulators
                    ElseIf Me.Controls("RA" & i & "2").Value = "(3)11K" Then
                        rLine1 = "11K-" & insulators & "-" & IIf(Me.Controls("RA" & i & "3").Value = "", [lead], Me.Controls("RA" & i & "3").Value) & "-STE"
                        rLine2 = "11K-" & insulators
                        rLine3 = "11K-" & insulators
                    ElseIf Me.Controls("RA" & i & "2").Value = "(2)20K" Then
                        rLine1 = "20K-" & insulators & "-" & IIf(Me.Controls("RA" & i & "3").Value = "", [lead], Me.Controls("RA" & i & "3").Value) & "-STE"
                        rLine2 = "20K-" & insulators
                    End If
                End If
            End If
            replaceNotes = replaceNotes & eLine1 & " / " & rLine1 & " " & IIf(Me.Controls("EA" & i & "4").Value = "", "[DIRECTION]", Me.Controls("EA" & i & "4").Value) & vbLf
            If eLine2 <> "" Or rLine2 <> "" Then
                replaceNotes = replaceNotes & eLine2 & Space((Len(eLine1) * 2) - Len(eLine2) - 4) & rLine2 & vbLf
            End If
            If eLine3 <> "" Or rLine3 <> "" Then
                replaceNotes = replaceNotes & eLine3 & Space((Len(eLine1) * 2) - Len(eLine3) - 4) & rLine3 & vbLf
            End If
            
            If Me.Controls("DGRFW" & i).Value = "" Or Me.Controls("DGRFW" & i).Value = "OTHER" Then
                replaceNotes = replaceNotes & "[REASON FOR WORK]" & vbLf
            Else
                replaceNotes = replaceNotes & Me.Controls("DGRFW" & i) & vbLf
            End If
            
            If Me.Controls("DGFIG" & i).Value = "" Or Me.Controls("DGFIG" & i).Value = "OTHER" Then
                replaceNotes = replaceNotes & "FIGURE [INSERT DOWNGUY FIGURE]" & vbLf
            Else
                replaceNotes = replaceNotes & Split(Me.Controls("DGFIG" & i).Value, ",")(0) & vbLf
            End If
        End If
    Next i
    
End Sub

Private Sub compileCrewNotes()
    If installNotes <> "" Then
        CrewNotes = CrewNotes & "INSTALL" & vbLf & installNotes & vbLf
    End If
    If removeNotes <> "" Then
        CrewNotes = CrewNotes & "REMOVE" & vbLf & removeNotes & vbLf
    End If
    If replaceNotes <> "" Then
        CrewNotes = CrewNotes & "REPLACE" & vbLf & replaceNotes & vbLf
    End If
    If transferNotes <> "" Then
        CrewNotes = CrewNotes & "TRANSFER" & vbLf & transferNotes & vbLf
    End If
    If notes <> "" Then
        CrewNotes = CrewNotes & notes
    End If
End Sub

Private Sub Initialize_ReplacePole()
    RPPSIZE.list = Array("35", "40", "45", "50", "55", "60", "65", "70")
    RPPCLASS.list = Array("1", "2", "3", "4")
    If InStr(pds.Range("HEIGHT"), "(") > 0 Then
        If RPPSIZE.Value = "" Then
            RPPSIZE.Value = Left(Trim(pds.Range("HEIGHT")), InStr(Trim(pds.Range("HEIGHT")), "(") - 1)
        ElseIf CInt(Left(Trim(pds.Range("HEIGHT")), InStr(Trim(pds.Range("HEIGHT")), "(") - 1)) > 40 Then
            RPPSIZE.Value = Left(Trim(pds.Range("HEIGHT")), InStr(Trim(pds.Range("HEIGHT")), "(") - 1)
        End If
    Else
        If RPPSIZE.Value = "" Then
            RPPSIZE.Value = Trim(pds.Range("HEIGHT"))
        ElseIf CInt(Trim(pds.Range("HEIGHT"))) > 40 Then
            RPPSIZE.Value = Trim(pds.Range("HEIGHT"))
        End If
    End If
    If pds.Range("SECONLY").Value = True Then
        RPPCLASS.Value = 4
        DSV.Value = "N/A"
        PSV.Value = "N/A"
    ElseIf InStr(pds.Range("DESC"), "3") > 0 Or InStr(pds.Range("DESC").offset(1, 0), "3") > 0 Or InStr(pds.Range("DESC").offset(2, 0), "3") > 0 Or InStr(pds.Range("DESC").offset(3, 0), "3") > 0 Then
        RPPCLASS.Value = 2
    Else
        RPPCLASS.Value = 3
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
        TRT.Value = True
        TRSIZE.Value = Replace(pds.Range("EQUIPMENTSIZE"), " ", "")
        TRRPSIZE.Value = Replace(pds.Range("EQUIPMENTSIZE"), " ", "")
    End If
    
    If Trim(pds.Range("STLTBRKT").Value) <> "N/A" And Trim(pds.Range("STLTBRKT").Value) <> "" Then
        TRSL.Value = True
        STLTBTMBRKT.Value = Trim(Split(pds.Range("STLTBRKT").Value, vbLf)(0))
        If pds.Range("MBSM").offset(0, 1).Value = "YES" Then
            SLM1.Value = True
        ElseIf pds.Range("MBSM").offset(0, 1).Value = "NO" Then
            SLM2.Value = True
        End If
    End If
    
    If riserBool Then
        RPR.Value = True
    End If
    
    If dg11KCount > 0 Or dg20KCount > 0 Then
        RPDG.Value = True
    End If
    
    If spg11KCount > 0 Or spg20KCount > 0 Then
        TRSPG.Value = True
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
    SLDLRFW.list = Array("TO CORRECT 12 STREETLIGHT DRIP LOOP SEPARATION VIOLATION", "TO MAKE ROOM FOR APPLICANT", "TO ALLOW COMMS TO CORRECT VIOLATIONS", "OTHER")
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
        For Each name In pds.names
            If name.name = "'" & pds.name & "'" & "!" & "TOPOLE" & i Then
                If Trim(Replace(pds.Range("TOPOLE" & i).Value, "-", "")) <> "" Then
                    For j = 1 To 50
                        If pds.Range("UTTYPE").offset(j, 0).Interior.color = 16777215 Then Exit For
                        If pds.Range("UTTYPE").offset(j, 0) = "" Then Exit For
                        If pds.Range("UTTYPE").offset(j, 0).text = "STLT. BOTTOM BRKT." Then Exit For
                        If InStr(pds.Range("UTTYPE").offset(j, 0).text, "OW") > 0 Then
                            If Trim(Replace(pds.Range("UTMIDSPAN" & i).offset(j, 0).Value, "-", "")) <> "" Then
                                ReDim Preserve spanAray(UBound(spanAray) + 1)
                                spanAray(UBound(spanAray)) = Trim(pds.Range("TOPOLE" & i).Value)
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
        If pds.Range("ANCHOROWNER").offset(i, 0).Value = "Consumers Energy" Then
            anchorCount = anchorCount + 1
            Me.Controls("EA" & anchorCount & 3).Value = pds.Range("ANCHOROWNER").offset(i, 0).offset(0, 1).Value
            Me.Controls("RA" & anchorCount & 3).Value = pds.Range("ANCHOROWNER").offset(i, 0).offset(0, 1).Value
            Me.Controls("EA" & anchorCount & 4).Value = getDirection(pds.Range("ANCHOROWNER").offset(i, 0).offset(0, 1).offset(0, 1).Value)
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
    If SNT1.Value = "RP" Then
        SN1.Value = snRP1
        SNC1.Value = PreviousValueRPSNC1
    ElseIf SNT1.Value = "IN" Then
        SN1.Value = snIN1
        SNC1.Value = PreviousValueINSNC1
    ElseIf SNT1.Value = "RM" Then
        SN1.Value = snRM1
        SNC1.Value = PreviousValueRMSNC1
    End If
End Sub

Private Sub SN1_Click()
    SNC1.visible = SN1.Value
    
    If SNT1.Value = "RP" Then
        If snRP1 <> SN1.Value Then
            snRP1 = SN1.Value
            If SN1.Value Then
                UASN.Value = CDbl(UASN.Value) - (CInt(SNC1.Value) * 2)
            Else
                UASN.Value = CDbl(UASN.Value) + (CInt(SNC1.Value) * 2)
            End If
        End If
    ElseIf SNT1.Value = "IN" Then
        If snIN1 <> SN1.Value Then
            snIN1 = SN1.Value
            If SN1.Value Then
                UASN.Value = CDbl(UASN.Value) - (CInt(SNC1.Value))
            Else
                UASN.Value = CDbl(UASN.Value) + (CInt(SNC1.Value))
            End If
        End If
    ElseIf SNT1.Value = "RM" Then
        If snRM1 <> SN1.Value Then
            snRM1 = SN1.Value
            If SN1.Value Then
                UASN.Value = CDbl(UASN.Value) - (CInt(SNC1.Value))
            Else
                UASN.Value = CDbl(UASN.Value) + (CInt(SNC1.Value))
            End If
        End If
    End If
End Sub

Private Sub SNC1_Change()
    If SNT1.Value = "RP" Then
        If PreviousValueRPSNC1 <> SNC1.Value Then
            UASN.Value = CDbl(UASN.Value) - (2 * (SNC1.Value - PreviousValueRPSNC1))
            PreviousValueRPSNC1 = SNC1.Value
        End If
    ElseIf SNT1.Value = "IN" Then
        If PreviousValueINSNC1 <> SNC1.Value Then
            UASN.Value = CDbl(UASN.Value) - (SNC1.Value - PreviousValueINSNC1)
            PreviousValueINSNC1 = SNC1.Value
        End If
    ElseIf SNT1.Value = "RM" Then
        If PreviousValueRMSNC1 <> SNC1.Value Then
            UASN.Value = CDbl(UASN.Value) - (SNC1.Value - PreviousValueRMSNC1)
            PreviousValueRMSNC1 = SNC1.Value
        End If
    End If
End Sub

Private Sub SNT2_Change()
    If SNT2.Value = "RP" Then
        SN2.Value = snRP2
        SNC2.Value = PreviousValueRPSNC2
    ElseIf SNT2.Value = "IN" Then
        SN2.Value = snIN2
        SNC2.Value = PreviousValueINSNC2
    ElseIf SNT2.Value = "RM" Then
        SN2.Value = snRM2
        SNC2.Value = PreviousValueRMSNC2
    End If
End Sub

Private Sub SN2_Click()
    SNC2.visible = SN2.Value
    
    If SNT2.Value = "RP" Then
        If snRP2 <> SN2.Value Then
            snRP2 = SN2.Value
            If SN2.Value Then
                UASN.Value = CDbl(UASN.Value) - (CInt(SNC2.Value) * 2)
            Else
                UASN.Value = CDbl(UASN.Value) + (CInt(SNC2.Value) * 2)
            End If
        End If
    ElseIf SNT2.Value = "IN" Then
        If snIN2 <> SN2.Value Then
            snIN2 = SN2.Value
            If SN2.Value Then
                UASN.Value = CDbl(UASN.Value) - (CInt(SNC2.Value))
            Else
                UASN.Value = CDbl(UASN.Value) + (CInt(SNC2.Value))
            End If
        End If
    ElseIf SNT2.Value = "RM" Then
        If snRM2 <> SN2.Value Then
            snRM2 = SN2.Value
            If SN2.Value Then
                UASN.Value = CDbl(UASN.Value) - (CInt(SNC2.Value))
            Else
                UASN.Value = CDbl(UASN.Value) + (CInt(SNC2.Value))
            End If
        End If
    End If
End Sub

Private Sub SNC2_Change()
    If SNT2.Value = "RP" Then
        If PreviousValueRPSNC2 <> SNC2.Value Then
            UASN.Value = CDbl(UASN.Value) - (2 * (SNC2.Value - PreviousValueRPSNC2))
            PreviousValueRPSNC2 = SNC2.Value
        End If
    ElseIf SNT2.Value = "IN" Then
        If PreviousValueINSNC2 <> SNC2.Value Then
            UASN.Value = CDbl(UASN.Value) - (SNC2.Value - PreviousValueINSNC2)
            PreviousValueINSNC2 = SNC2.Value
        End If
    ElseIf SNT2.Value = "RM" Then
        If PreviousValueRMSNC2 <> SNC2.Value Then
            UASN.Value = CDbl(UASN.Value) - (SNC2.Value - PreviousValueRMSNC2)
            PreviousValueRMSNC2 = SNC2.Value
        End If
    End If
End Sub

Private Sub SNT3_Change()
    If SNT3.Value = "RP" Then
        SN3.Value = snRP3
        SNC3.Value = PreviousValueRPSNC3
    ElseIf SNT3.Value = "IN" Then
        SN3.Value = snIN3
        SNC3.Value = PreviousValueINSNC3
    ElseIf SNT3.Value = "RM" Then
        SN3.Value = snRM3
        SNC3.Value = PreviousValueRMSNC3
    End If
End Sub

Private Sub SN3_Click()
    SNC3.visible = SN3.Value
    
    If SNT3.Value = "RP" Then
        If snRP3 <> SN3.Value Then
            snRP3 = SN3.Value
            If SN3.Value Then
                UASN.Value = CDbl(UASN.Value) - (CInt(SNC3.Value) * 2)
            Else
                UASN.Value = CDbl(UASN.Value) + (CInt(SNC3.Value) * 2)
            End If
        End If
    ElseIf SNT3.Value = "IN" Then
        If snIN3 <> SN3.Value Then
            snIN3 = SN3.Value
            If SN3.Value Then
                UASN.Value = CDbl(UASN.Value) - (CInt(SNC3.Value))
            Else
                UASN.Value = CDbl(UASN.Value) + (CInt(SNC3.Value))
            End If
        End If
    ElseIf SNT3.Value = "RM" Then
        If snRM3 <> SN3.Value Then
            snRM3 = SN3.Value
            If SN3.Value Then
                UASN.Value = CDbl(UASN.Value) - (CInt(SNC3.Value))
            Else
                UASN.Value = CDbl(UASN.Value) + (CInt(SNC3.Value))
            End If
        End If
    End If
End Sub

Private Sub SNC3_Change()
    If SNT3.Value = "RP" Then
        If PreviousValueRPSNC3 <> SNC3.Value Then
            UASN.Value = CDbl(UASN.Value) - (2 * (SNC3.Value - PreviousValueRPSNC3))
            PreviousValueRPSNC3 = SNC3.Value
        End If
    ElseIf SNT3.Value = "IN" Then
        If PreviousValueINSNC3 <> SNC3.Value Then
            UASN.Value = CDbl(UASN.Value) - (SNC3.Value - PreviousValueINSNC3)
            PreviousValueINSNC3 = SNC3.Value
        End If
    ElseIf SNT3.Value = "RM" Then
        If PreviousValueRMSNC3 <> SNC3.Value Then
            UASN.Value = CDbl(UASN.Value) - (SNC3.Value - PreviousValueRMSNC3)
            PreviousValueRMSNC3 = SNC3.Value
        End If
    End If
End Sub

Private Sub SNT4_Change()
    If SNT4.Value = "RP" Then
        SN4.Value = snRP4
        SNC4.Value = PreviousValueRPSNC4
    ElseIf SNT4.Value = "IN" Then
        SN4.Value = snIN4
        SNC4.Value = PreviousValueINSNC4
    ElseIf SNT4.Value = "RM" Then
        SN4.Value = snRM4
        SNC4.Value = PreviousValueRMSNC4
    End If
End Sub

Private Sub SN4_Click()
    SNC4.visible = SN4.Value
    
    If SNT4.Value = "RP" Then
        If snRP4 <> SN4.Value Then
            snRP4 = SN4.Value
            If SN4.Value Then
                UASN.Value = CDbl(UASN.Value) - (CInt(SNC4.Value))
            Else
                UASN.Value = CDbl(UASN.Value) + (CInt(SNC4.Value))
            End If
        End If
    ElseIf SNT4.Value = "IN" Then
        If snIN4 <> SN4.Value Then
            snIN4 = SN4.Value
            If SN4.Value Then
                UASN.Value = CDbl(UASN.Value) - (CInt(SNC4.Value) * 0.5)
            Else
                UASN.Value = CDbl(UASN.Value) + (CInt(SNC4.Value) * 0.5)
            End If
        End If
    ElseIf SNT4.Value = "RM" Then
        If snRM4 <> SN4.Value Then
            snRM4 = SN4.Value
            If SN4.Value Then
                UASN.Value = CDbl(UASN.Value) - (CInt(SNC4.Value) * 0.5)
            Else
                UASN.Value = CDbl(UASN.Value) + (CInt(SNC4.Value) * 0.5)
            End If
        End If
    End If
End Sub

Private Sub SNC4_Change()
    If SNT4.Value = "RP" Then
        If PreviousValueRPSNC4 <> SNC4.Value Then
            UASN.Value = CDbl(UASN.Value) - (SNC4.Value - PreviousValueRPSNC4)
            PreviousValueRPSNC4 = SNC4.Value
        End If
    ElseIf SNT4.Value = "IN" Then
        If PreviousValueINSNC4 <> SNC4.Value Then
            UASN.Value = CDbl(UASN.Value) - (0.5 * (SNC4.Value - PreviousValueINSNC4))
            PreviousValueINSNC4 = SNC4.Value
        End If
    ElseIf SNT4.Value = "RM" Then
        If PreviousValueRMSNC4 <> SNC4.Value Then
            UASN.Value = CDbl(UASN.Value) - (0.5 * (SNC4.Value - PreviousValueRMSNC4))
            PreviousValueRMSNC4 = SNC4.Value
        End If
    End If
End Sub

Private Sub SNT5_Change()
    If SNT5.Value = "RP" Then
        SN5.Value = snRP5
        SNC5.Value = PreviousValueRPSNC5
    ElseIf SNT5.Value = "IN" Then
        SN5.Value = snIN5
        SNC5.Value = PreviousValueINSNC5
    ElseIf SNT5.Value = "RM" Then
        SN5.Value = snRM5
        SNC5.Value = PreviousValueRMSNC5
    End If
End Sub

Private Sub SN5_Click()
    SNC5.visible = SN5.Value
    
    If SNT5.Value = "RP" Then
        If snRP5 <> SN5.Value Then
            snRP5 = SN5.Value
            If SN5.Value Then
                UASN.Value = CDbl(UASN.Value) - (CInt(SNC5.Value))
            Else
                UASN.Value = CDbl(UASN.Value) + (CInt(SNC5.Value))
            End If
        End If
    ElseIf SNT5.Value = "IN" Then
        If snIN5 <> SN5.Value Then
            snIN5 = SN5.Value
            If SN5.Value Then
                UASN.Value = CDbl(UASN.Value) - (CInt(SNC5.Value) * 0.5)
            Else
                UASN.Value = CDbl(UASN.Value) + (CInt(SNC5.Value) * 0.5)
            End If
        End If
    ElseIf SNT5.Value = "RM" Then
        If snRM5 <> SN5.Value Then
            snRM5 = SN5.Value
            If SN5.Value Then
                UASN.Value = CDbl(UASN.Value) - (CInt(SNC5.Value) * 0.5)
            Else
                UASN.Value = CDbl(UASN.Value) + (CInt(SNC5.Value) * 0.5)
            End If
        End If
    End If
End Sub

Private Sub SNC5_Change()
    If SNT5.Value = "RP" Then
        If PreviousValueRPSNC5 <> SNC5.Value Then
            UASN.Value = CDbl(UASN.Value) - (SNC5.Value - PreviousValueRPSNC5)
            PreviousValueRPSNC5 = SNC5.Value
        End If
    ElseIf SNT5.Value = "IN" Then
        If PreviousValueINSNC5 <> SNC5.Value Then
            UASN.Value = CDbl(UASN.Value) - (0.5 * (SNC5.Value - PreviousValueINSNC5))
            PreviousValueINSNC5 = SNC5.Value
        End If
    ElseIf SNT5.Value = "RM" Then
        If PreviousValueRMSNC5 <> SNC5.Value Then
            UASN.Value = CDbl(UASN.Value) - (0.5 * (SNC5.Value - PreviousValueRMSNC5))
            PreviousValueRMSNC5 = SNC5.Value
        End If
    End If
End Sub

Private Sub SNT6_Change()
    If SNT6.Value = "RP" Then
        SN6.Value = snRP6
        SNC6.Value = PreviousValueRPSNC6
    ElseIf SNT6.Value = "IN" Then
        SN6.Value = snIN6
        SNC6.Value = PreviousValueINSNC6
    ElseIf SNT6.Value = "RM" Then
        SN6.Value = snRM6
        SNC6.Value = PreviousValueRMSNC6
    End If
End Sub

Private Sub SN6_Click()
    SNC6.visible = SN6.Value
    
    If SNT6.Value = "RP" Then
        If snRP6 <> SN6.Value Then
            snRP6 = SN6.Value
            If SN6.Value Then
                UASN.Value = CDbl(UASN.Value) - (CInt(SNC6.Value))
            Else
                UASN.Value = CDbl(UASN.Value) + (CInt(SNC6.Value))
            End If
        End If
    ElseIf SNT6.Value = "IN" Then
        If snIN6 <> SN6.Value Then
            snIN6 = SN6.Value
            If SN6.Value Then
                UASN.Value = CDbl(UASN.Value) - (CInt(SNC6.Value) * 0.5)
            Else
                UASN.Value = CDbl(UASN.Value) + (CInt(SNC6.Value) * 0.5)
            End If
        End If
    ElseIf SNT6.Value = "RM" Then
        If snRM6 <> SN6.Value Then
            snRM6 = SN6.Value
            If SN6.Value Then
                UASN.Value = CDbl(UASN.Value) - (CInt(SNC6.Value) * 0.5)
            Else
                UASN.Value = CDbl(UASN.Value) + (CInt(SNC6.Value) * 0.5)
            End If
        End If
    End If
End Sub

Private Sub SNC6_Change()
    If SNT6.Value = "RP" Then
        If PreviousValueRPSNC6 <> SNC6.Value Then
            UASN.Value = CDbl(UASN.Value) - (SNC6.Value - PreviousValueRPSNC6)
            PreviousValueRPSNC6 = SNC6.Value
        End If
    ElseIf SNT6.Value = "IN" Then
        If PreviousValueINSNC6 <> SNC6.Value Then
            UASN.Value = CDbl(UASN.Value) - (0.5 * (SNC6.Value - PreviousValueINSNC6))
            PreviousValueINSNC6 = SNC6.Value
        End If
    ElseIf SNT6.Value = "RM" Then
        If PreviousValueRMSNC6 <> SNC6.Value Then
            UASN.Value = CDbl(UASN.Value) - (0.5 * (SNC6.Value - PreviousValueRMSNC6))
            PreviousValueRMSNC6 = SNC6.Value
        End If
    End If
End Sub

Private Sub SNT7_Change()
    If SNT7.Value = "RP" Then
        SN7.Value = snRP7
        SNC7.Value = PreviousValueRPSNC7
    ElseIf SNT7.Value = "IN" Then
        SN7.Value = snIN7
        SNC7.Value = PreviousValueINSNC7
    ElseIf SNT7.Value = "RM" Then
        SN7.Value = snRM7
        SNC7.Value = PreviousValueRMSNC7
    End If
End Sub

Private Sub SN7_Click()
    SNC7.visible = SN7.Value
    
    If SNT7.Value = "RP" Then
        If snRP7 <> SN7.Value Then
            snRP7 = SN7.Value
            If SN7.Value Then
                UASN.Value = CDbl(UASN.Value) - (CInt(SNC7.Value) * 2)
            Else
                UASN.Value = CDbl(UASN.Value) + (CInt(SNC7.Value) * 2)
            End If
        End If
    ElseIf SNT7.Value = "IN" Then
        If snIN7 <> SN7.Value Then
            snIN7 = SN7.Value
            If SN7.Value Then
                UASN.Value = CDbl(UASN.Value) - (CInt(SNC7.Value))
            Else
                UASN.Value = CDbl(UASN.Value) + (CInt(SNC7.Value))
            End If
        End If
    ElseIf SNT7.Value = "RM" Then
        If snRM7 <> SN7.Value Then
            snRM7 = SN7.Value
            If SN7.Value Then
                UASN.Value = CDbl(UASN.Value) - (CInt(SNC7.Value))
            Else
                UASN.Value = CDbl(UASN.Value) + (CInt(SNC7.Value))
            End If
        End If
    End If
End Sub

Private Sub SNC7_Change()
    If SNT7.Value = "RP" Then
        If PreviousValueRPSNC7 <> SNC7.Value Then
            UASN.Value = CDbl(UASN.Value) - (2 * (SNC7.Value - PreviousValueRPSNC7))
            PreviousValueRPSNC7 = SNC7.Value
        End If
    ElseIf SNT7.Value = "IN" Then
        If PreviousValueINSNC7 <> SNC7.Value Then
            UASN.Value = CDbl(UASN.Value) - (SNC7.Value - PreviousValueINSNC7)
            PreviousValueINSNC7 = SNC7.Value
        End If
    ElseIf SNT7.Value = "RM" Then
        If PreviousValueRMSNC7 <> SNC7.Value Then
            UASN.Value = CDbl(UASN.Value) - (SNC7.Value - PreviousValueRMSNC7)
            PreviousValueRMSNC7 = SNC7.Value
        End If
    End If
End Sub

Private Sub TRLETTER_Change()
    Call determineRecDSpace
End Sub

Private Sub TRT_Click()
    If TRT.Value And RPT.Value Then RPT.Value = False
    XSL.visible = RPP.Value And (TRT.Value Or RPT.Value)
    TRSIZE.visible = RPP.Value And (TRT.Value Or RPT.Value)
    CLL.visible = RPP.Value And (TRT.Value Or RPT.Value)
    COLA.visible = RPP.Value And (TRT.Value Or RPT.Value)
    XDL.visible = RPP.Value And (TRT.Value Or RPT.Value)
    TRLETTER.visible = RPP.Value And (TRT.Value Or RPT.Value)
    Call determineRecDSpace
End Sub

Private Sub RPT_Click()
    If RPT.Value And TRT.Value Then TRT.Value = False
    XSL.visible = RPP.Value And (TRT.Value Or RPT.Value)
    TRSIZE.visible = TRT.Value Or RPT.Value
    XRL.visible = RPP.Value And RPT.Value
    TRRPSIZE.visible = RPP.Value And RPT.Value
    CLL.visible = RPP.Value And (TRT.Value Or RPT.Value)
    COLA.visible = RPP.Value And (TRT.Value Or RPT.Value)
    XDL.visible = RPP.Value And (TRT.Value Or RPT.Value)
    TRLETTER.visible = RPP.Value And (TRT.Value Or RPT.Value)
    Call determineRecDSpace
End Sub

Private Sub UASN_Change()
    UASN2.Value = UASN.Value
End Sub
