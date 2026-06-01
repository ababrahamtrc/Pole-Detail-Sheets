Attribute VB_Name = "UtilitiesKatapultNameMapping"
Private KatapultNameMapping As Scripting.Dictionary

Public Function getKatapultNameMapping(ByVal key As String) As String
    key = Trim(UCase(key))
    
    If KatapultNameMapping Is Nothing Then
        Set KatapultNameMapping = New Scripting.Dictionary
        Call InitializeKatapultNameMapping
    End If
    
    If KatapultNameMapping.Exists(key) Then
        getKatapultNameMapping = KatapultNameMapping(key)
    Else
        getKatapultNameMapping = key
    End If
        
End Function

Private Sub InitializeKatapultNameMapping()
    'General
    KatapultNameMapping("PRIMARY") = "PRI"
    KatapultNameMapping("NEUTRAL") = "NEUT"
    KatapultNameMapping("SECONDARY") = "SEC"
    KatapultNameMapping("POWER DROP") = "SVC"
    KatapultNameMapping("POWER GUY") = "SPG"
    KatapultNameMapping("SECONDARY") = "SEC"
    KatapultNameMapping("COMM DROP") = "DROP"
    KatapultNameMapping("TELCO COM") = "COM"
    KatapultNameMapping("CATV COM") = "COM"
    KatapultNameMapping("TRANSFORMER") = "XFMR"
    KatapultNameMapping("STREET_LIGHT") = "SL"
    KatapultNameMapping("STREET LIGHT") = "SL"
    KatapultNameMapping("DRIP_LOOP") = "DL"
    
    'Pole species
    KatapultNameMapping("WESTERN RED CEDAR") = "WRC"
    KatapultNameMapping("SOUTHERN PINE") = "SP"

    'Primary and Neutral sizes
    KatapultNameMapping("4 ACSR (7/1)") = "4 ACSR"
    KatapultNameMapping("2 ACSR (7/1)") = "2 ACSR"
    KatapultNameMapping("2 ACSR (7/1) HDPE (45)") = "2 ACSR"
    KatapultNameMapping("1/0 ACSR (6/1)") = "1/0 ACSR"
    KatapultNameMapping("1/0 ACSR (6/1) HDPE (60)") = "1/0 ACSR"
    KatapultNameMapping("3/0 ACSR (6/1)") = "3/0 ACSR"
    KatapultNameMapping("3/0 ACSR (6/1) HDPE (60)") = "3/0 ACSR"
    KatapultNameMapping("336 ACSR (26/7)") = "336 ACSR"
    KatapultNameMapping("795 ACSR (26/7)") = "795"
    KatapultNameMapping("795 AL (37)") = "795 AL"
    KatapultNameMapping("80 KCMIL ACSR (8/1)") = "80 KCMIL ACSR"
    KatapultNameMapping("1/0 AERIAL SPACER CABLE") = "1/0 ASC"
    KatapultNameMapping("336 AERIAL SPACER CABLE") = "336 ASC"
    KatapultNameMapping("795 AERIAL SPACER CABLE") = "795 ASC"
    KatapultNameMapping("4 WP ACSR") = "4 ACSR"
    KatapultNameMapping("350 AAC (19)") = "350 AAC"
    KatapultNameMapping("6 CU (1)") = "6 CU"
    KatapultNameMapping("4 CU (1)") = "4 CU"
    KatapultNameMapping("3 CU (1)") = "3 CU"
    KatapultNameMapping("2 CU (1)") = "2 CU"
    KatapultNameMapping("4 ACSR (7/1) - SLACK 3PH") = "4 ACSR SLACK"
    KatapultNameMapping("2 ACSR (7/1) - SLACK 3PH") = "2 ACSR SLACK"
    KatapultNameMapping("1/0 ACSR (6/1) - SLACK 3PH") = "1/0 ACSR SLACK"
    KatapultNameMapping("3/0 ACSR (6/1) - SLACK 3PH") = "3/0 ACSR SLACK"
    KatapultNameMapping("336 ACSR (26/7) - SLACK 3PH") = "336 ACSR SLACK"
    KatapultNameMapping("795 ACSR (26/7) - SLACK 3PH") = "795 SLACK"
    KatapultNameMapping("795 AL (37) - SLACK 3PH") = "795 AL SLACK"
    KatapultNameMapping("1/0 ACSR (6/1) TREE WIRE") = "1/0 ACSR TREE WIRE"
    KatapultNameMapping("336 AAC (19) TREE WIRE") = "336 AAC TREE WIRE"
    KatapultNameMapping("4 ACSR (7/1) - SLACK 1PH") = "4 ACSR SLACK"
    KatapultNameMapping("2 ACSR (7/1) - SLACK 1PH") = "2 ACSR SLACK"
    KatapultNameMapping("1/0 ACSR (6/1) - SLACK 1PH") = "1/0 ACSR SLACK"
    KatapultNameMapping("3/0 ACSR (6/1) - SLACK 1PH") = "3/0 ACSR SLACK"
    KatapultNameMapping("336 ACSR (26/7) - SLACK 1PH") = "336 ACSR SLACK"
    KatapultNameMapping("795 ACSR (26/7) - SLACK 1PH") = "795 SLACK"
    KatapultNameMapping("795 AL (37) - SLACK 1PH") = "795 AL SLACK"
    
    'Neutral sizes
    KatapultNameMapping("052 (1/0-1Ø)") = "052 (1/0-1Ø)"
    KatapultNameMapping("052 (336-1Ø)") = "052 (336-1Ø)"
    KatapultNameMapping("052 (1/0-3Ø)") = "052 (1/0-3Ø)"
    KatapultNameMapping("052 (336-3Ø)") = "052 (336-3Ø)"
    KatapultNameMapping("7#6 (795-3Ø)") = "7#6 (795-3Ø)"
    KatapultNameMapping("052 AWA Shield") = "052 AWA Shield"
    KatapultNameMapping("7#6") = "795"
    KatapultNameMapping("350 AAC") = "350 AAC"
    
    'Secondary sizes
    KatapultNameMapping("4-4-4-4 ACSR QX") = "4 QX"
    KatapultNameMapping("2-2-2-2 ACSR QX") = "2 QX"
    KatapultNameMapping("1/0-1/0-1/0-1/0 ACSR QX") = "1/0 QX"
    KatapultNameMapping("3/0-3/0-3/0-3/0 ACSR QX") = "3/0 QX"
    'KatapultNameMapping("4 WP ACSR") = "4 ACSR" Duplicate
    KatapultNameMapping("2 WP ACSR") = "2 ACSR"
    KatapultNameMapping("1/0 HDPE ACSR") = "1/0 ACSR"
    KatapultNameMapping("3/0 HDPE ACSR") = "3/0 ACSR"
    KatapultNameMapping("6-6 ACSR DX SERV") = "6 DX"
    KatapultNameMapping("2-2 ACSR DX SERV") = "2 DX"
    KatapultNameMapping("4-4-4 ACSR TX SERV") = "4 TX"
    KatapultNameMapping("1/0-2-1/0 ACSR TX SERV") = "1/0 TX"
    KatapultNameMapping("3/0-1/0-3/0 ACSR TX SERV") = "3/0 TX"
    KatapultNameMapping("4-4-4-4 ACSR QX SERV") = "4 QX"
    KatapultNameMapping("2-2-2-2 ACSR QX SERV") = "2 QX"
    KatapultNameMapping("1/0-1/0-1/0-1/0 ACSR QX SERV") = "1/0 QX"
    KatapultNameMapping("3/0-3/0-3/0-3/0 ACSR QX SERV") = "3/0 QX"
    KatapultNameMapping("6-6 ACSR DX") = "6 DX"
    KatapultNameMapping("2-2 ACSR DX") = "2 DX"
    KatapultNameMapping("4-4-4 ACSR TX") = "4 TX"
    KatapultNameMapping("1/0-2-1/0 ACSR TX") = "1/0 TX"
    KatapultNameMapping("3/0-1/0-3/0 ACSR TX") = "3/0 TX"
    KatapultNameMapping("2-4-2 AWAC TX") = "2 TX AWAC"
    KatapultNameMapping("1/0-4-1/0 AWAC TX") = "1/0 TX AWAC"
    KatapultNameMapping("3/0-2-3/0 AWAC TX") = "3/0 TX AWAC"
    'KatapultNameMapping("6 CU (1)") = "6 CU" DUPLICATE
    
    'Com sizes
    KatapultNameMapping("6.6M (1/4) + 0.50"" TELCO") = "0.5"""
    KatapultNameMapping("6.6M (1/4) + 0.75"" TELCO") = "0.75"""
    KatapultNameMapping("6M (5/16) + 0.50"" TELCO") = "0.5"""
    KatapultNameMapping("6M (5/16) + 0.75"" TELCO") = "0.75"""
    KatapultNameMapping("10M (3/8) + 1.00"" TELCO") = "1"""
    KatapultNameMapping("10M (3/8) + 1.50"" TELCO") = "1.5"""
    KatapultNameMapping("16M (7/16) + 1.75"" TELCO") = "1.75"""
    KatapultNameMapping("16M (7/16) + 2.00"" TELCO") = "2"""
    KatapultNameMapping("25M (1/2) + 2.50"" TELCO") = "2.5"""
    KatapultNameMapping("25M (1/2) + 3.00"" TELCO") = "3"""
    KatapultNameMapping("6.6M (1/4) + 0.50"" CATV") = "0.5"""
    KatapultNameMapping("6.6M (1/4) + 0.75"" CATV") = "0.75"""
    KatapultNameMapping("6.6M (1/4) + 1.00"" CATV") = "1"""
    KatapultNameMapping("10M (3/8) + 1.50"" CATV") = "1.5"""
    KatapultNameMapping("10M (3/8) + 2.00"" CATV") = "2"""
    KatapultNameMapping("0.50"" - Slack") = "0.5"""
    KatapultNameMapping("1.00"" - Slack") = "1"""
    KatapultNameMapping("2.00"" - Slack") = "2"""
    KatapultNameMapping("3.00"" - Slack") = "3"""
    
    'Guy sizes
    KatapultNameMapping("11K 5/16"" EHS") = "11K"
    KatapultNameMapping("20K 7/16"" EHS") = "20K"

    'Transformer sizes
    KatapultNameMapping("3-10 KVA (3Ã˜ CL MT)") = "30 KVA"
    KatapultNameMapping("3-25 KVA (3Ã˜ CL MT)") = "75 KVA"
    KatapultNameMapping("3-50 KVA (3Ã˜ CL MT)") = "150 KVA"
    KatapultNameMapping("3-100 KVA (3Ã˜ CL MT)") = "300 KVA"
    KatapultNameMapping("3-167 KVA (3Ã˜ CL MT)") = "500 KVA"
    
    'Riser sizes
    KatapultNameMapping("1"" RISER - PRIMARY") = "1"" PRI"
    KatapultNameMapping("2"" RISER - PRIMARY") = "2"" PRI"
    KatapultNameMapping("3"" RISER - PRIMARY") = "3"" PRI"
    KatapultNameMapping("4"" RISER - PRIMARY") = "4"" PRI"
    KatapultNameMapping("5"" RISER - PRIMARY") = "5"" PRI"
    KatapultNameMapping("6"" RISER - PRIMARY") = "6"" PRI"
    KatapultNameMapping("1"" RISER - SECONDARY") = "1"" SEC"
    KatapultNameMapping("2"" RISER - SECONDARY") = "2"" SEC"
    KatapultNameMapping("3"" RISER - SECONDARY") = "3"" SEC"
    KatapultNameMapping("4"" RISER - SECONDARY") = "4"" SEC"
    KatapultNameMapping("5"" RISER - SECONDARY") = "5"" SEC"
    KatapultNameMapping("6"" RISER - SECONDARY") = "6"" SEC"

    'Streetlight sizes
    KatapultNameMapping("150-400W COBRA") = "LARGE"
    KatapultNameMapping("1000W COBRA") = ""
    KatapultNameMapping("150-400W FLOOD") = ""
    KatapultNameMapping("1000W FLOOD") = ""
    KatapultNameMapping("100W OPEN BOTTOM") = ""
    KatapultNameMapping("LED FLOOD") = ""
    KatapultNameMapping("LED SMALL STREET LIGHT") = ""
    KatapultNameMapping("LED LARGE STREET LIGHT") = ""
End Sub
