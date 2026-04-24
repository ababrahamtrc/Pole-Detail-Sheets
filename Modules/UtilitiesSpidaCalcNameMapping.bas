Attribute VB_Name = "UtilitiesSpidaCalcNameMapping"
Private SpidaCalcNameMapping As Scripting.Dictionary

Public Function getSpidaCalcNameMapping(ByVal key As String) As String
    key = Trim(UCase(key))
    
    If SpidaCalcNameMapping Is Nothing Then
        Set SpidaCalcNameMapping = New Scripting.Dictionary
        Call InitializeSpidaCalcNameMapping
    End If
    
    If SpidaCalcNameMapping.Exists(key) Then
        getSpidaCalcNameMapping = SpidaCalcNameMapping(key)
    Else
        getSpidaCalcNameMapping = key
    End If
        
End Function

Private Sub InitializeSpidaCalcNameMapping()
    'General
    SpidaCalcNameMapping("PRIMARY") = "PRI"
    SpidaCalcNameMapping("SECONDARY") = "SEC"
    SpidaCalcNameMapping("NEUTRAL") = "NEUT"
    SpidaCalcNameMapping("OPEN_WIRE") = "OW"
    SpidaCalcNameMapping("UTILITY_SERVICE") = "SVC"
    SpidaCalcNameMapping("COMMUNICATION_SERVICE") = "DROP"
    SpidaCalcNameMapping("COMMUNICATION") = "COM"
    SpidaCalcNameMapping("GUY") = "MSG"
    
    SpidaCalcNameMapping("SOUTHERN PINE") = "SP"
    SpidaCalcNameMapping("WESTERN RED CEDAR") = "WRC"
    
    SpidaCalcNameMapping("STREET LIGHT") = "SL"
    SpidaCalcNameMapping("STREET_LIGHT") = "SL"
    SpidaCalcNameMapping("DRIP_LOOP") = "DL"
    
    SpidaCalcNameMapping("FLOODLIGHT") = "FL"
    SpidaCalcNameMapping("LARGE") = "LARGE"
    SpidaCalcNameMapping("MEDIUM") = "MED"
    SpidaCalcNameMapping("SMALL") = "SMALL"
    
    'Transformer
    SpidaCalcNameMapping("TRANSFORMER") = "XFMR"
    SpidaCalcNameMapping("10 KVA SINGLE-PHASE") = "10 KVA"
    SpidaCalcNameMapping("100 KVA SINGLE-PHASE") = "100 KVA"
    SpidaCalcNameMapping("15 KVA SINGLE-PHASE") = "15 KVA"
    SpidaCalcNameMapping("167 KVA SINGLE-PHASE") = "167 KVA"
    SpidaCalcNameMapping("25 KVA SINGLE-PHASE") = "25 KVA"
    SpidaCalcNameMapping("250 KVA SINGLE-PHASE") = "250 KVA"
    SpidaCalcNameMapping("333 KVA SINGLE-PHASE") = "333 KVA"
    SpidaCalcNameMapping("37.5 KVA SINGLE-PHASE") = "37.5 KVA"
    SpidaCalcNameMapping("5 KVA SINGLE-PHASE") = "5 KVA"
    SpidaCalcNameMapping("50 KVA SINGLE-PHASE") = "50 KVA"
    SpidaCalcNameMapping("500 KVA SINGLE-PHASE") = "500 KVA"
    SpidaCalcNameMapping("75 KVA SINGLE-PHASE") = "75 KVA"
    
    'CO
    SpidaCalcNameMapping("CUTOUT_ARRESTOR") = "CO"
    SpidaCalcNameMapping("1 ARRESTOR") = "1 CO"
    SpidaCalcNameMapping("1 CUTOUT") = "1 CO"
    SpidaCalcNameMapping("2 ARRESTOR") = "2 CO"
    SpidaCalcNameMapping("2 CUTOUT") = "2 CO"
    SpidaCalcNameMapping("3 ARRESTOR") = "3 CO"
    SpidaCalcNameMapping("3 CUTOUT") = "3 CO"

    'Com sizes
    SpidaCalcNameMapping("CABLE 1/4""M - 0.5""") = "0.5"""
    SpidaCalcNameMapping("CABLE 1/4""M - 1""") = "1"""
    SpidaCalcNameMapping("CABLE 5/16""M - 1.5""") = "1.5"""
    SpidaCalcNameMapping("CABLE 5/16""M - 2""") = "2"""
    SpidaCalcNameMapping("CABLE 5/16""M - 3""") = "3"""
    SpidaCalcNameMapping("CABLE 5/16""M - 4""") = "4"""
    SpidaCalcNameMapping("CABLE SERVICE - 0.25""") = "0.25"""
    SpidaCalcNameMapping("FIBER SERVICE - 0.25""") = "0.25"""
    SpidaCalcNameMapping("SELF-SUPPORT - 0.5""") = "0.5"""
    SpidaCalcNameMapping("SELF-SUPPORT - 1""") = "1"""
    SpidaCalcNameMapping("SELF-SUPPORT - 1.5""") = "1.5"""
    SpidaCalcNameMapping("TELE 1/4""M - 0.5""") = "0.5"""
    SpidaCalcNameMapping("TELE 1/4""M - 1""") = "1"""
    SpidaCalcNameMapping("TELE 3/8""M - 2""") = "2"""
    SpidaCalcNameMapping("TELE 3/8""M - 3""") = "3"""
    SpidaCalcNameMapping("TELE 3/8""M - 4""") = "4"""
    SpidaCalcNameMapping("TELE 5/16""M - 0.5""") = "0.5"""
    SpidaCalcNameMapping("TELE 5/16""M - 1""") = "1"""
    SpidaCalcNameMapping("TELE SERVICE - 0.25""") = "0.25"""
End Sub
