VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PrintOptions 
   Caption         =   "Print Options"
   ClientHeight    =   6060
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9885.001
   OleObjectBlob   =   "PrintOptions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PrintOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public ShowDrawing As Boolean

Public DrawConductors As Boolean
Public DrawPrimary As Boolean
Public DrawSecondary As Boolean
Public DrawOpenWire As Boolean
Public DrawDeadends As Boolean
Public ConductorInitOffset As Long
Public ConductorOffsetAmount As Long

Public DrawROW As Boolean
Public DrawCenterLines As Boolean
Public DrawCenterLineDistances As Boolean

Public DrawDownGuys As Boolean
Public DrawSpanGuys As Boolean
Public DrawCrewNotes As Boolean
Public DrawTrees As Boolean

Public DrawServices As Boolean
Public DrawUGServices As Boolean
Public DrawAdjacentServices As Boolean

Public DrawEquipment As Boolean
Public DrawAdjacentPoleEquipment As Boolean
Public DrawTransformers As Boolean
Public DrawStreetlights As Boolean
Public DrawCapacitors As Boolean
Public DrawRegulators As Boolean
Public DrawIsolators As Boolean
Public DrawFuses As Boolean
Public DrawReclosers As Boolean
Public DrawSectionalizers As Boolean
Public DrawSwitches As Boolean
Public DrawSensors As Boolean
Public DrawSecondaryRisers As Boolean
Public DrawPrimaryRisers As Boolean

Public Sub Initialize()
    Set json = PrintGenerator.ReadJSON("settings")
    
    If Not json Is Nothing Then
        ShowDraw.value = json("ShowDrawing")
        
        Conductor.value = json("DrawConductors")
        Conductor1.value = json("DrawPrimary")
        Conductor2.value = json("DrawSecondary")
        Conductor3.value = json("DrawOpenWire")
        Conductor4.value = json("DrawDeadends")
        Conductor5.value = json("ConductorInitOffset")
        Conductor6.value = json("ConductorOffsetAmount")
        
        ROW.value = json("DrawROW")
        ROW1.value = json("DrawCenterLines")
        ROW1a.value = json("DrawCenterLineDistances")
        
        DG.value = json("DrawDownGuys")
        SPG.value = json("DrawSpanGuys")
        CN.value = json("DrawCrewNotes")
        Tree.value = json("DrawTrees")
        
        Service.value = json("DrawServices")
        Service1.value = json("DrawUGServices")
        Service2.value = json("DrawAdjacentServices")
        
        Equipment.value = json("DrawEquipment")
        Equipment1.value = json("DrawAdjacentPoleEquipment")
        Equipment2.value = json("DrawTransformers")
        Equipment3.value = json("DrawStreetlights")
        Equipment4.value = json("DrawCapacitors")
        Equipment5.value = json("DrawRegulators")
        Equipment6.value = json("DrawIsolators")
        Equipment7.value = json("DrawFuses")
        Equipment8.value = json("DrawReclosers")
        Equipment9.value = json("DrawSectionalizers")
        Equipment10.value = json("DrawSwitches")
        Equipment11.value = json("DrawSensors")
        Equipment12.value = json("DrawSecondaryRisers")
        Equipment13.value = json("DrawPrimaryRisers")
    End If
End Sub

Private Sub CommandButton1_Click()
    userResponse = MsgBox("Are you sure you want to generate the print?", _
                          vbYesNoCancel + vbQuestion, _
                          "Confirm Action")
    
    If userResponse <> vbYes Then Exit Sub
    
    ShowDrawing = ShowDraw.value
    
    DrawConductors = Conductor.value
    DrawPrimary = Conductor1.value And Conductor1.Enabled
    DrawSecondary = Conductor2.value And Conductor2.Enabled
    DrawOpenWire = Conductor3.value And Conductor3.Enabled
    DrawDeadends = Conductor4.value And Conductor4.Enabled
    ConductorInitOffset = Conductor5.value
    ConductorOffsetAmount = Conductor6.value
    
    DrawROW = ROW.value
    DrawCenterLines = ROW1.value And ROW1.Enabled
    DrawCenterLineDistances = ROW1a.value And ROW1a.Enabled
    
    DrawDownGuys = DG.value
    DrawSpanGuys = SPG.value
    DrawCrewNotes = CN.value
    DrawTrees = Tree.value
    
    DrawServices = Service.value
    DrawUGServices = Service1.value And Service1.Enabled
    DrawAdjacentServices = Service2.value And Service2.Enabled
    
    DrawEquipment = Equipment.value
    DrawAdjacentPoleEquipment = Equipment1.value And Equipment1.Enabled
    DrawTransformers = Equipment2.value And Equipment2.Enabled
    DrawStreetlights = Equipment3.value And Equipment3.Enabled
    DrawCapacitors = Equipment4.value And Equipment4.Enabled
    DrawRegulators = Equipment5.value And Equipment5.Enabled
    DrawIsolators = Equipment6.value And Equipment6.Enabled
    DrawFuses = Equipment7.value And Equipment7.Enabled
    DrawReclosers = Equipment8.value And Equipment8.Enabled
    DrawSectionalizers = Equipment9.value And Equipment9.Enabled
    DrawSwitches = Equipment10.value And Equipment10.Enabled
    DrawSensors = Equipment11.value And Equipment11.Enabled
    DrawSecondaryRisers = Equipment12.value And Equipment12.Enabled
    DrawPrimaryRisers = Equipment13.value And Equipment13.Enabled
    
    Dim varDict As Object
    Set varDict = CreateObject("Scripting.Dictionary")
    
    varDict.Add "ShowDrawing", ShowDraw.value
    
    varDict.Add "DrawConductors", Conductor.value
    varDict.Add "DrawPrimary", Conductor1.value
    varDict.Add "DrawSecondary", Conductor2.value
    varDict.Add "DrawOpenWire", Conductor3.value
    varDict.Add "DrawDeadends", Conductor4.value
    varDict.Add "ConductorInitOffset", Conductor5.value
    varDict.Add "ConductorOffsetAmount", Conductor6.value
    
    varDict.Add "DrawROW", ROW.value
    varDict.Add "DrawCenterLines", ROW1.value
    varDict.Add "DrawCenterLineDistances", ROW1a.value
    
    varDict.Add "DrawDownGuys", DG.value
    varDict.Add "DrawSpanGuys", SPG.value
    varDict.Add "DrawCrewNotes", CN.value
    varDict.Add "DrawTrees", Tree.value
    
    varDict.Add "DrawServices", Service.value
    varDict.Add "DrawUGServices", Service1.value
    varDict.Add "DrawAdjacentServices", Service2.value
    
    varDict.Add "DrawEquipment", Equipment.value
    varDict.Add "DrawAdjacentPoleEquipment", Equipment1.value
    varDict.Add "DrawTransformers", Equipment2.value
    varDict.Add "DrawStreetlights", Equipment3.value
    varDict.Add "DrawCapacitors", Equipment4.value
    varDict.Add "DrawRegulators", Equipment5.value
    varDict.Add "DrawIsolators", Equipment6.value
    varDict.Add "DrawFuses", Equipment7.value
    varDict.Add "DrawReclosers", Equipment8.value
    varDict.Add "DrawSectionalizers", Equipment9.value
    varDict.Add "DrawSwitches", Equipment10.value
    varDict.Add "DrawSensors", Equipment11.value
    varDict.Add "DrawSecondaryRisers", Equipment12.value
    varDict.Add "DrawPrimaryRisers", Equipment13.value
    
    JsonString = JsonConverter.ConvertToJson(varDict, Whitespace:=4)
    
    jsonFile = "C:\ProgramData\Bentley\OpenUtilities Map Connect Edition\Configuration\WorkSpaces\ConsumersEnergy\Standards\vba\settings.json"
    fileNum = FreeFile
    
    Open jsonFile For Output As #fileNum
    Print #fileNum, JsonString
    Close #fileNum
    
    Me.Hide
    Call PrintGenerator.DrawPrint
End Sub

Private Sub Conductor_Click()
    Dim ctrl As control
    For Each ctrl In Me.Controls
        If InStr(ctrl.name, "Conductor") > 0 And ctrl.name <> "Conductor" Then
            ctrl.Enabled = Conductor.value
        End If
    Next ctrl
End Sub

Private Sub Service_Click()
    Dim ctrl As control
    For Each ctrl In Me.Controls
        If InStr(ctrl.name, "Service") > 0 And ctrl.name <> "Service" Then
            ctrl.Enabled = Service.value
        End If
    Next ctrl
End Sub

Private Sub ROW_Click()
    For Each ctrl In Me.Controls
        If InStr(ctrl.name, "ROW") > 0 And ctrl.name <> "ROW" Then
            ctrl.Enabled = ROW.value
        End If
    Next ctrl
End Sub

Private Sub ROW1_Click()
    For Each ctrl In Me.Controls
        If InStr(ctrl.name, "ROW1") > 0 And ctrl.name <> "ROW1" Then
            ctrl.Enabled = ROW1.value
        End If
    Next ctrl
End Sub

Private Sub Equipment_Click()
    For Each ctrl In Me.Controls
        If InStr(ctrl.name, "Equipment") > 0 And ctrl.name <> "Equipment" Then
            ctrl.Enabled = Equipment.value
        End If
    Next ctrl
End Sub

Private Sub conductor5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub conductor5_Change()
    If InStr(Conductor5.text, "-") > 0 Then
        Conductor5.text = Replace(Conductor5.text, "-", "")
    End If
    
    If Conductor5.text <> "" And Not IsNumeric(Conductor5.text) Then
        MsgBox "Only positive numbers are allowed.", vbCritical, "Invalid Input"
        Conductor5.text = ""
    End If
End Sub

Private Sub conductor6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub conductor6_Change()
    If InStr(Conductor6.text, "-") > 0 Then
        Conductor6.text = Replace(Conductor6.text, "-", "")
    End If
    
    If Conductor6.text <> "" And Not IsNumeric(Conductor6.text) Then
        MsgBox "Only positive numbers are allowed.", vbCritical, "Invalid Input"
        Conductor6.text = ""
    End If
End Sub
