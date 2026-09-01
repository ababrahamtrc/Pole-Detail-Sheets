Attribute VB_Name = "CUNameMapping"
Private CUNameCorrecting As Scripting.Dictionary
Private CUNameMapping As Scripting.Dictionary
Private CUExtraCUsNeeded As Scripting.Dictionary
Private CUOpenWireNameMapping As Scripting.Dictionary
Private CUSecNameMapping As Scripting.Dictionary
Private CUPriNameMapping As Scripting.Dictionary
Private CUAssemblyMapping As Scripting.Dictionary

Public Function getCUAssemblyMapping() As Scripting.Dictionary
    Call InitializeCUAssemblyMapping
    
    Set getCUAssemblyMapping = CUAssemblyMapping
End Function

Sub InitializeCUAssemblyMapping()
    If CUAssemblyMapping Is Nothing Then
        Set CUAssemblyMapping = New Scripting.Dictionary
    Else
        Exit Sub
    End If
    
    CUAssemblyMapping("505002") = Array("100131", "100194", "100421") 'DOWN GUY 11K 1P W/RS ASSEMBLY
    CUAssemblyMapping("505004") = Array("100133", "100194", "100421") 'DOWN GUY 11K 1P W/RT
    CUAssemblyMapping("505006") = Array("100136", "100194", "100421") 'DOWN GUY 11K 1P W/STE
    CUAssemblyMapping("505013") = Array("100133", "100195", "100422") 'DOWN GUY 20K 1P W/RT
    CUAssemblyMapping("504995") = Array("100133", "100425")  'DOWN GUY 20K FG W/RT
    CUAssemblyMapping("504997") = Array("100136", "100425")  'DOWN GUY 20K FG W/STE
End Sub

Public Function getCUNameMapping(ByVal key As String) As String
    cleansedKey = cleanseKey(key)
    
    If CUNameCorrecting Is Nothing Then
        Set CUNameCorrecting = New Scripting.Dictionary
        Call InitializeCUNameCorrecting
    End If
    
    If CUNameMapping Is Nothing Then
        Set CUNameMapping = New Scripting.Dictionary
        Call InitializeCUNameMapping
    End If
    
    If CUNameCorrecting.exists(cleansedKey) Then cleansedKey = CUNameCorrecting(cleansedKey)
    If CUNameCorrecting.exists(key) Then key = CUNameCorrecting(key)
    
    If CUNameMapping.exists(cleansedKey) Then
        getCUNameMapping = CUNameMapping(cleansedKey)
    ElseIf CUNameMapping.exists(key) Then
        getCUNameMapping = CUNameMapping(key)
    Else
        getCUNameMapping = ""
    End If
End Function

Public Function CheckForAdditionalCUs(ByVal key As String) As Boolean
    key = cleanseKey(key)
    
    If CUNameCorrecting Is Nothing Then
        Set CUNameCorrecting = New Scripting.Dictionary
        Call InitializeCUNameCorrecting
    End If
    
    If CUExtraCUsNeeded Is Nothing Then
        Set CUExtraCUsNeeded = New Scripting.Dictionary
        Call InitializeCUExtraCUsNeeded
    End If
    
    If CUNameCorrecting.exists(key) Then key = CUNameCorrecting(key)
    CheckForAdditionalCUs = CUExtraCUsNeeded.exists(key)
End Function

Public Function getOWNameMapping(ByVal key As String) As String
    key = Replace(key, " ", "")
    
    If CUOpenWireNameMapping Is Nothing Then
        Set CUOpenWireNameMapping = New Scripting.Dictionary
        Call InitializeCUOpenWireNameMapping
    End If
    
    If CUOpenWireNameMapping.exists(key) Then
        getOWNameMapping = CUOpenWireNameMapping(key)
    Else
        getOWNameMapping = ""
    End If
End Function

Public Function getSecNameMapping(ByVal key As String) As String
    key = Replace(key, " ", "")
    key = Replace(key, "OF", "")
    
    If CUSecNameMapping Is Nothing Then
        Set CUSecNameMapping = New Scripting.Dictionary
        Call InitializeCUSecNameMapping
    End If
    
    If CUSecNameMapping.exists(key) Then
        getSecNameMapping = CUSecNameMapping(key)
    Else
        getSecNameMapping = ""
    End If
End Function

Public Function getPriNameMapping(ByVal key As String) As String
    key = Replace(key, " ", "")
    
    If CUPriNameMapping Is Nothing Then
        Set CUPriNameMapping = New Scripting.Dictionary
        Call InitializeCUPriNameMapping
    End If
    
    If CUPriNameMapping.exists(key) Then
        getPriNameMapping = CUPriNameMapping(key)
    Else
        getPriNameMapping = ""
    End If
End Function

Private Function cleanseKey(key As String) As String
    key = Trim(UCase(key))
    key = Replace(key, " ", "")
    cleansedKey = Replace(key, """", "")
    cleansedKey = Replace(cleansedKey, "#", "")
    cleansedKey = Replace(cleansedKey, "/0", "|0")
    cleansedKey = Replace(cleansedKey, "-RISER", "RISER")
    cleansedKey = Replace(cleansedKey, "DEADEND", "DE")
    cleansedKey = Replace(cleansedKey, "DEADENDS", "DE")
    cleansedKey = Replace(cleansedKey, "PRIMARY", "PRI")
    cleansedKey = Replace(cleansedKey, "NEUTRAL", "NEUT")
    cleansedKey = Replace(cleansedKey, "OPENWIRE", "OW")
    cleansedKey = Replace(cleansedKey, "SECONDARY", "SEC")
    cleansedKey = Replace(cleansedKey, "SERVICE", "SVC")
    
    cleansedKey = ThisWorkbook.RemoveParentheses(cleansedKey)
    
    cleanseKey = cleansedKey
End Function

Private Sub InitializeCUExtraCUsNeeded()
    CUExtraCUsNeeded("WR") = True
    CUExtraCUsNeeded("PRIDE") = True
    CUExtraCUsNeeded("NEUTDE") = True
    CUExtraCUsNeeded("SECDE") = True
    CUExtraCUsNeeded("PTP") = True
    CUExtraCUsNeeded("SPIN") = True
    CUExtraCUsNeeded("SCOR") = True
    CUExtraCUsNeeded("1VPO") = True
    CUExtraCUsNeeded("2VPO") = True
    CUExtraCUsNeeded("3VPO") = True
End Sub

Private Sub InitializeCUNameCorrecting()
    CUNameCorrecting("DEVICEARM") = "S8S"
    CUNameCorrecting("LCPTAGS") = "LCPTAG"
    CUNameCorrecting("VDE") = "PRIDE"
    CUNameCorrecting("CLUSTERMOUNTBRACKET") = "CLUSTERMOUNT"
End Sub

Private Sub InitializeCUNameMapping()
    'Framing Hardware
    CUNameMapping("PTP") = "100080"
    CUNameMapping("JUMPERPTP") = "100080"
    CUNameMapping("S8S") = "100020"
    CUNameMapping("S8M") = "100018"
    CUNameMapping("S8L") = "100013"
    CUNameMapping("S10S") = "100017"
    CUNameMapping("D8S") = "100006"
    CUNameMapping("D10S") = "100002"
    CUNameMapping("D8M") = "100004"
    CUNameMapping("D8L") = "100011"
    CUNameMapping("S8FGDE") = "107000"
    CUNameMapping("S10FGDE") = "107001"
    CUNameMapping("S11FGDE") = "107002"
    CUNameMapping("S11FGTAN") = "107003"
    CUNameMapping("SSA") = "100027"
    CUNameMapping("DSA") = "100007"
    CUNameMapping("SPIN") = "100022"
    CUNameMapping("PPI") = "200495"
    CUNameMapping("JUMPERSPIN") = "100022"
    CUNameMapping("SCOR") = "100105"
    CUNameMapping("RISERPIN") = "100016"
    CUNameMapping("WR") = "100052"
    CUNameMapping("TANCLAMP") = "100056"
    CUNameMapping("SECTANCLAMP") = "100056"
    CUNameMapping("AWACTANCLAMP") = "100036"
    CUNameMapping("TANCLAMPAWAC") = "100036"
    CUNameMapping("SECAWACTANCLAMP") = "100036"
    CUNameMapping("1VPO") = "290019"
    CUNameMapping("2VPO") = "290029"
    CUNameMapping("3VPO") = "290030"
    CUNameMapping("FGSTANDOFFBRACKETWITHSPIN") = "100029"
    CUNameMapping("18""FGSTANDOFFBRKTW|INS") = "100029"
    CUNameMapping("18FGSTANDOFF") = "100029"
    CUNameMapping("18""FGARM") = "100068"
    CUNameMapping("FGSTANDOFF") = "100029"
    CUNameMapping("FGSTANDOFFBRACKET") = "100029"
    CUNameMapping("FGSTANDOFFBRKT") = "100029"
    CUNameMapping("FGSB") = "100029"
    
    CUNameMapping("11KGRIP") = "100159"
    CUNameMapping("SPANGUYGRIP") = "100159"
    CUNameMapping("SPGGRIP") = "100159"
    
    CUNameMapping("TEMPORARYJUMPERTOCOMMPOWERSUPPLY") = "106277"
    CUNameMapping("TEMPORARYJUMPER") = "106277"
    CUNameMapping("TEMPJUMPER") = "106277"
    
    CUNameMapping("FGINS") = "100191"
    CUNameMapping("FG") = "100191"
    CUNameMapping("FGLINK") = "100191"
    
    CUNameMapping("CLUSTERMOUNT") = "200548"
    CUNameMapping("SIDEWALKBRACE") = "100432"
    CUNameMapping("SINGLE-PHASETERMINATIONBRACKET") = "106124"
    CUNameMapping("TERMINATIONBRACKET") = "106124"
    CUNameMapping("BONDWIRE") = "100144"
    
    'Misc
    CUNameMapping("CEIDTAG") = "200047"
    CUNameMapping("LCPTAG") = "200046"
    
    'Pole sizes
    CUNameMapping("STUBPOLE") = "100911"
    CUNameMapping("25-1") = "100911"
    CUNameMapping("25-2") = "100911"
    CUNameMapping("25-3") = "100911"
    CUNameMapping("25-4") = "100911"
    CUNameMapping("25-5") = "100911"
    CUNameMapping("25-6") = "100911"
    CUNameMapping("25-7") = "100911"
    CUNameMapping("30-1") = "100936"
    CUNameMapping("30-2") = "100936"
    CUNameMapping("30-3") = "100936"
    CUNameMapping("30-4") = "100936"
    CUNameMapping("30-5") = "100936"
    CUNameMapping("30-6") = "100936"
    CUNameMapping("30-7") = "100936"
    CUNameMapping("30-7") = "100936"
    CUNameMapping("35-7") = "100927"
    CUNameMapping("35-6") = "100927"
    CUNameMapping("35-5") = "100928"
    CUNameMapping("35-4") = "200054"
    CUNameMapping("40-6") = "100949"
    CUNameMapping("40-5") = "200051"
    CUNameMapping("40-4") = "100948"
    CUNameMapping("40-3") = "200055"
    CUNameMapping("40-2") = "200056"
    CUNameMapping("45-5") = "200049"
    CUNameMapping("45-4") = "100943"
    CUNameMapping("45-3") = "100942"
    CUNameMapping("45-2") = "200057"
    CUNameMapping("50-4") = "100909"
    CUNameMapping("50-3") = "100908"
    CUNameMapping("50-2") = "201706"
    CUNameMapping("50-1") = "201707"
    CUNameMapping("55-4") = "100884"
    CUNameMapping("55-3") = "100883"
    CUNameMapping("55-2") = "201705"
    CUNameMapping("55-1") = "201708"
    CUNameMapping("60-4") = "100958"
    CUNameMapping("60-3") = "100957"
    CUNameMapping("60-2") = "100956"
    CUNameMapping("65-4") = "100953"
    CUNameMapping("65-3") = "100952"
    CUNameMapping("65-2") = "100951"
    CUNameMapping("70-3") = "100964"
    CUNameMapping("70-2") = "200058"
    CUNameMapping("75-3") = "100960"
    CUNameMapping("75-2") = "200059"
    CUNameMapping("80-3") = "100969"
    CUNameMapping("80-2") = "200060"
    CUNameMapping("85-3") = "100967"
    CUNameMapping("85-2") = "200061"

    'Primary/Neutral deadend + grips
    CUNameMapping("PRIDE") = "290014"
    CUNameMapping("NEUTDE") = "290034"
    CUNameMapping("6CUDEGRIP") = "100710"
    CUNameMapping("4ACSRDEGRIP") = "100635"
    CUNameMapping("2ACSRDEGRIP") = "100619"
    CUNameMapping("1|0ACSRDEGRIP") = "100609"
    CUNameMapping("3|0ACSRDEGRIP") = "100626"
    CUNameMapping("336ACSRDEGRIP") = "100694"
    CUNameMapping("336ALDEGRIP") = "100696"
    
    CUNameMapping("6CUPRIDE") = "100710"
    CUNameMapping("4ACSRPRIDE") = "100635"
    CUNameMapping("2ACSRPRIDE") = "100619"
    CUNameMapping("1|0ACSRPRIDE") = "100609"
    CUNameMapping("3|0ACSRPRIDE") = "100626"
    CUNameMapping("336ACSRPRIDE") = "100694"
    CUNameMapping("336ALDPRIDE") = "100696"
    
    CUNameMapping("6PRIDE") = "100710"
    CUNameMapping("4PRIDE") = "100635"
    CUNameMapping("2PRIDE") = "100619"
    CUNameMapping("1|0PRIDE") = "100609"
    CUNameMapping("3|0PRIDE") = "100626"
    CUNameMapping("336PRIDE") = "100694"
    
    CUNameMapping("6NEUTDE") = "100710"
    CUNameMapping("4NEUTDE") = "100635"
    CUNameMapping("2NEUTDE") = "100619"
    CUNameMapping("1|0NEUTDE") = "100609"
    CUNameMapping("3|0NEUTDE") = "100626"
    CUNameMapping("336NEUTDE") = "100694"
    
    CUNameMapping("6CUNEUTDE") = "100710"
    CUNameMapping("4ACSRNEUTDE") = "100635"
    CUNameMapping("2ACSRNEUTDE") = "100619"
    CUNameMapping("1|0ACSRNEUTDE") = "100609"
    CUNameMapping("3|0ACSRNEUTDE") = "100626"
    CUNameMapping("336ACSRNEUTDE") = "100694"
    CUNameMapping("336ALDNEUTDE") = "100696"
    
    CUNameMapping("1|0ASCDEGRIP") = "201043"
    CUNameMapping("336ASCDEGRIP") = "200725"
    CUNameMapping("795ASCDEGRIP") = "201044"
    CUNameMapping("052(1/0-1Ø)DEGRIP") = "201043"
    CUNameMapping("052(1/0-3Ø)DEGRIP") = "201043"
    CUNameMapping("052(3/0-1Ø)DEGRIP") = "100626"
    CUNameMapping("052(3/0-3Ø)DEGRIP") = "100626"
    
    'Secondary deadends
    CUNameMapping("OWDE") = "101036"
    CUNameMapping("OWSECDE") = "101036"
    CUNameMapping("OWDESEC") = "101036"
    CUNameMapping("3|0OWDE") = "101036"
    CUNameMapping("1|0OWDE") = "101036"
    CUNameMapping("6OWDE") = "101036"
    CUNameMapping("4OWDE") = "101036"
    CUNameMapping("3OWDE") = "101036"
    CUNameMapping("2OWDE") = "101036"
    CUNameMapping("3|0ACSROWDE") = "101036"
    CUNameMapping("1|0ACSROWDE") = "101036"
    CUNameMapping("2ACSROWDE") = "101036"
    CUNameMapping("4ACSROWDE") = "101036"
    CUNameMapping("6CUOWDE") = "101036"
    CUNameMapping("3CUWDE") = "101036"
    CUNameMapping("1ACSR-COVDE") = "101036"
    CUNameMapping("1|0ACSR-COVDE") = "101036"
    CUNameMapping("2ACSR-COVDE") = "101036"
    CUNameMapping("3CUDE") = "101036"
    CUNameMapping("4ACSR-COVDE") = "101036"
    CUNameMapping("4CUDE") = "101036"
    CUNameMapping("6ACSR-COVDE") = "101036"
    CUNameMapping("6CU-COVDE") = "101036"
    
    CUNameMapping("6DXDE") = "101031"
    CUNameMapping("4TXDE") = "101030"
    CUNameMapping("4QXDE") = "101050"
    CUNameMapping("2TXDE") = "101033"
    CUNameMapping("1|0TXDE") = "101026"
    CUNameMapping("1|0QXDE") = "101046"
    CUNameMapping("3|0TXDE") = "101029"
    CUNameMapping("3|0QXDE") = "101049"

    CUNameMapping("2TXAWACDE") = "101032"
    CUNameMapping("2QXAWACDE") = "101032"
    CUNameMapping("4TXAWACDE") = "101032"
    CUNameMapping("4QXAWACDE") = "101032"
    CUNameMapping("1|0TXAWACDE") = "101032"
    CUNameMapping("1|0QXAWACDE") = "101032"
    CUNameMapping("3|0TXAWACDE") = "101032"
    CUNameMapping("3|0QXAWACDE") = "101032"
    
    CUNameMapping("2TXSECAWACDE") = "101032"
    CUNameMapping("2QXSECAWACDE") = "101032"
    CUNameMapping("4TXSECAWACDE") = "101032"
    CUNameMapping("4QXSECAWACDE") = "101032"
    CUNameMapping("1|0TXSECAWACDE") = "101032"
    CUNameMapping("1|0QXSECAWACDE") = "101032"
    CUNameMapping("3|0TXSECAWACDE") = "101032"
    CUNameMapping("3|0QXSECAWACDE") = "101032"

    CUNameMapping("6DXSECDE") = "101031"
    CUNameMapping("4TXSECDE") = "101030"
    CUNameMapping("4QXSECDE") = "101050"
    CUNameMapping("2TXSECDE") = "101033"
    CUNameMapping("1|0TXSECDE") = "101026"
    CUNameMapping("1|0QXSECDE") = "101046"
    CUNameMapping("3|0TXSECDE") = "101029"
    CUNameMapping("3|0QXSECDE") = "101049"
    
    CUNameMapping("2TXAWACSECDE") = "101032"
    CUNameMapping("2QXAWACSECDE") = "101032"
    CUNameMapping("4TXAWACSECDE") = "101032"
    CUNameMapping("4QXAWACSECDE") = "101032"
    CUNameMapping("1|0TXAWACSECDE") = "101032"
    CUNameMapping("1|0QXAWACSECDE") = "101032"
    CUNameMapping("3|0TXAWACSECDE") = "101032"
    CUNameMapping("3|0QXAWACSECDE") = "101032"
    
    CUNameMapping("SEC6DXDE") = "101031"
    CUNameMapping("SEC4TXDE") = "101030"
    CUNameMapping("SEC4QXDE") = "101050"
    CUNameMapping("SEC2TXDE") = "101033"
    CUNameMapping("SEC1|0TXDE") = "101026"
    CUNameMapping("SEC1|0QXDE") = "101046"
    CUNameMapping("SEC3|0TXDE") = "101029"
    CUNameMapping("SEC3|0QXDE") = "101049"
    
    CUNameMapping("6DXSVCDE") = "101031"
    CUNameMapping("4TXSVCDE") = "101030"
    CUNameMapping("4QXSVCDE") = "101050"
    CUNameMapping("2TXSVCDE") = "101033"
    CUNameMapping("1|0TXSVCDE") = "101026"
    CUNameMapping("1|0QXSVCDE") = "101046"
    CUNameMapping("3|0TXSVCDE") = "101029"
    CUNameMapping("3|0QXSVCDE") = "101049"
    
    CUNameMapping("SVC6DXDE") = "101031"
    CUNameMapping("SVC4TXDE") = "101030"
    CUNameMapping("SVC4QXDE") = "101050"
    CUNameMapping("SVC2TXDE") = "101033"
    CUNameMapping("SVC1|0TXDE") = "101026"
    CUNameMapping("SVC1|0QXDE") = "101046"
    CUNameMapping("SVC3|0TXDE") = "101029"
    CUNameMapping("SVC3|0QXDE") = "101049"
    
    CUNameMapping("1|0TXAWACDE") = "101032"
    CUNameMapping("1|0QXAWACDE") = "101032"
    CUNameMapping("3|0TXAWACDE") = "101032"
    CUNameMapping("3|0QXAWACDE") = "101032"
    CUNameMapping("1|0AWACDE") = "101032"
    CUNameMapping("3|0AWACDE") = "101032"
    
    CUNameMapping("1|0TXAWACSECDE") = "101032"
    CUNameMapping("1|0QXAWACSECDE") = "101032"
    CUNameMapping("3|0TXAWACSECDE") = "101032"
    CUNameMapping("3|0QXAWACSECDE") = "101032"
    CUNameMapping("1|0AWACSECDE") = "101032"
    CUNameMapping("3|0AWACSECDE") = "101032"
    
    CUNameMapping("SEC1|0TXAWACDE") = "101032"
    CUNameMapping("SEC1|0QXAWACDE") = "101032"
    CUNameMapping("SEC3|0TXAWACDE") = "101032"
    CUNameMapping("SEC3|0QXAWACDE") = "101032"
    CUNameMapping("SEC1|0AWACDE") = "101032"
    CUNameMapping("SEC3|0AWACDE") = "101032"
    
    'Tanget Clamps
    CUNameMapping("6DXTANCLAMP") = "100056"
    CUNameMapping("4TXTANCLAMP") = "100056"
    CUNameMapping("4QXTANCLAMP") = "100056"
    CUNameMapping("2TXTANCLAMP") = "100056"
    CUNameMapping("1|0TXTANTCLAMP") = "100056"
    CUNameMapping("1|0QXTANCLAMP") = "100056"
    CUNameMapping("3|0TXTANCLAMP") = "100056"
    CUNameMapping("3|0QXTANCLAMP") = "100056"
    
    CUNameMapping("6DXSECTANCLAMP") = "100056"
    CUNameMapping("4TXSECTANCLAMP") = "100056"
    CUNameMapping("4QXSECTANCLAMP") = "100056"
    CUNameMapping("2TXSECTANCLAMP") = "100056"
    CUNameMapping("1|0TXSECTANCLAMP") = "100056"
    CUNameMapping("1|0QXSECTANCLAMP") = "100056"
    CUNameMapping("3|0TXSECTANCLAMP") = "100056"
    CUNameMapping("3|0QXSECTANCLAMP") = "100056"
    
    'Top/Side/Spool ties
    CUNameMapping("6TOPTIE") = "100183"
    CUNameMapping("4TOPTIE") = "100183"
    CUNameMapping("2TOPTIE") = "100168"
    CUNameMapping("10TOPTIE") = "100154"
    CUNameMapping("30TOPTIE") = "100173"
    CUNameMapping("336TOPTIE") = "100180"
    CUNameMapping("795TOPTIE") = "100188"
    CUNameMapping("6SIDETIE") = "100185"
    CUNameMapping("4SIDETIE") = "100185"
    CUNameMapping("2SIDETIE") = "100170"
    CUNameMapping("10SIDETIE") = "100156"
    CUNameMapping("30SIDETIE") = "100175"
    CUNameMapping("336SIDETIE") = "100179"
    CUNameMapping("6SPOOLTIE") = "100145"
    CUNameMapping("4SPOOLTIE") = "100145"
    CUNameMapping("2SPOOLTIE") = "100142"
    CUNameMapping("10SPOOLTIE") = "100119"
    CUNameMapping("30SPOOLTIE") = "100119"
    
    'Riser
    CUNameMapping("3|C-350RISER") = "704160"
    CUNameMapping("4|C-350RISER") = "704170"
    CUNameMapping("3|C-3|0RISER") = "708160"
    CUNameMapping("4|C-3|0RISER") = "708170"
    CUNameMapping("3|C-1|0RISER") = "710160"
    CUNameMapping("4|C-1|0RISER") = "710170"
    CUNameMapping("3|C-350ALRISER") = "704160"
    CUNameMapping("4|C-350ALRISER") = "704170"
    CUNameMapping("3|C-3|0ALRISER") = "708160"
    CUNameMapping("4|C-3|0ALRISER") = "708170"
    CUNameMapping("3|C-1|0ALRISER") = "710160"
    CUNameMapping("4|C-1|0ALRISER") = "710170"
    

    CUNameMapping("3|C-350RISERSPLICE") = "101853"
    CUNameMapping("4|C-350RISERSPLICE") = "101855"
    CUNameMapping("3|C-3|0RISERSPLICE") = "101852"
    CUNameMapping("4|C-3|0RISERSPLICE") = "201569"
    CUNameMapping("3|C-1|0RISERSPLICE") = "101850"
    CUNameMapping("4|C-1|0RISERSPLICE") = "200670"
    CUNameMapping("3|C-350ALRISERSPLICE") = "101853"
    CUNameMapping("4|C-350ALRISERSPLICE") = "101855"
    CUNameMapping("3|C-3|0ALRISERSPLICE") = "101852"
    CUNameMapping("4|C-3|0ALRISERSPLICE") = "201569"
    CUNameMapping("3|C-1|0ALRISERSPLICE") = "101850"
    CUNameMapping("4|C-1|0ALRISERSPLICE") = "200670"
End Sub

Private Sub InitializeCUOpenWireNameMapping()
    CUOpenWireNameMapping("6") = "714113"
    CUOpenWireNameMapping("4") = "713013"
    CUOpenWireNameMapping("3") = "712113"
    CUOpenWireNameMapping("2") = "711013"
    CUOpenWireNameMapping("1|0") = "710013"
End Sub

Private Sub InitializeCUSecNameMapping()
    CUSecNameMapping("4TX") = "713033"
    CUSecNameMapping("1|0TX") = "710033"
    CUSecNameMapping("3|0TX") = "708033"
    CUSecNameMapping("4ACSR") = "713033"
    CUSecNameMapping("1|0ACSR") = "710033"
    CUSecNameMapping("3|0ACSR") = "708033"
    
    CUSecNameMapping("4TX") = "713033"
    CUSecNameMapping("4ACSR") = "713033"
    
    CUSecNameMapping("3|0TXAWAC") = "708053"
    CUSecNameMapping("3|0AWACTX") = "708053"
    CUSecNameMapping("3|0AWACACSR") = "708053"
    CUSecNameMapping("3|0ACSRAWAC") = "708053"
    CUSecNameMapping("3|0AWAC") = "708053"
    
    CUSecNameMapping("1|0TXAWAC") = "710053"
    CUSecNameMapping("1|0AWACTX") = "710053"
    CUSecNameMapping("1|0AWACACSR") = "710053"
    CUSecNameMapping("1|0ACSRAWAC") = "710053"
    CUSecNameMapping("1|0AWAC") = "710053"
End Sub

Private Sub InitializeCUPriNameMapping()
    CUPriNameMapping("14ACSR") = "713011"
    CUPriNameMapping("24ACSR") = "013011"
    CUPriNameMapping("34ACSR") = "913011"
    CUPriNameMapping("14") = "713011"
    CUPriNameMapping("24") = "013011"
    CUPriNameMapping("34") = "913011"
    
    CUPriNameMapping("13CU") = "712091"
    CUPriNameMapping("23CU") = "012091"
    CUPriNameMapping("33CU") = "912091"
    CUPriNameMapping("13") = "712091"
    CUPriNameMapping("23") = "012091"
    CUPriNameMapping("33") = "912091"
    
    CUPriNameMapping("12ACSR") = "711011"
    CUPriNameMapping("22ACSR") = "011011"
    CUPriNameMapping("32ACSR") = "911012"
    CUPriNameMapping("12") = "711011"
    CUPriNameMapping("22") = "011011"
    CUPriNameMapping("32") = "911012"
    
    CUPriNameMapping("11|0ACSR") = "710011"
    CUPriNameMapping("21|0ACSR") = "010011"
    CUPriNameMapping("31|0ACSR") = "910011"
    CUPriNameMapping("11|0") = "710011"
    CUPriNameMapping("21|0") = "010011"
    CUPriNameMapping("31|0") = "910011"
    
    CUPriNameMapping("13|0ACSR") = "708011"
    CUPriNameMapping("23|0ACSR") = "008011"
    CUPriNameMapping("33|0ACSR") = "908011"
    CUPriNameMapping("13|0") = "708011"
    CUPriNameMapping("23|0") = "008011"
    CUPriNameMapping("33|0") = "908011"
    
    CUPriNameMapping("1336ACSR") = "705011"
    CUPriNameMapping("2336ACSR") = "005011"
    CUPriNameMapping("3336ACSR") = "905011"
    CUPriNameMapping("1336KCMILACSR") = "705011"
    CUPriNameMapping("2336KCMILACSR") = "005011"
    CUPriNameMapping("3336KCMILACSR") = "905011"
    CUPriNameMapping("1336") = "705011"
    CUPriNameMapping("2336") = "005011"
    CUPriNameMapping("3336") = "905011"
    
    CUPriNameMapping("1795ACSR") = "701011"
    CUPriNameMapping("2795ACSR") = "001011"
    CUPriNameMapping("3795ACSR") = "901011"
    CUPriNameMapping("1795") = "701011"
    CUPriNameMapping("2795") = "001011"
    CUPriNameMapping("3795") = "901011"
End Sub
