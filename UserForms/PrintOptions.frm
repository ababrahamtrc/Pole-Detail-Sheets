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
        ShowDraw.Value = json("ShowDrawing")
        
        Conductor.Value = json("DrawConductors")
        Conductor1.Value = json("DrawPrimary")
        Conductor2.Value = json("DrawSecondary")
        Conductor3.Value = json("DrawOpenWire")
        Conductor4.Value = json("DrawDeadends")
        Conductor5.Value = json("ConductorInitOffset")
        Conductor6.Value = json("ConductorOffsetAmount")
        
        row.Value = json("DrawROW")
        ROW1.Value = json("DrawCenterLines")
        ROW1a.Value = json("DrawCenterLineDistances")
        
        DG.Value = json("DrawDownGuys")
        SPG.Value = json("DrawSpanGuys")
        CN.Value = json("DrawCrewNotes")
        Tree.Value = json("DrawTrees")
        
        Service.Value = json("DrawServices")
        Service1.Value = json("DrawUGServices")
        Service2.Value = json("DrawAdjacentServices")
        
        equipment.Value = json("DrawEquipment")
        Equipment1.Value = json("DrawAdjacentPoleEquipment")
        Equipment2.Value = json("DrawTransformers")
        Equipment3.Value = json("DrawStreetlights")
        Equipment4.Value = json("DrawCapacitors")
        Equipment5.Value = json("DrawRegulators")
        Equipment6.Value = json("DrawIsolators")
        Equipment7.Value = json("DrawFuses")
        Equipment8.Value = json("DrawReclosers")
        Equipment9.Value = json("DrawSectionalizers")
        Equipment10.Value = json("DrawSwitches")
        Equipment11.Value = json("DrawSensors")
        Equipment12.Value = json("DrawSecondaryRisers")
        Equipment13.Value = json("DrawPrimaryRisers")
    End If
End Sub

Private Sub CommandButton1_Click()
    userResponse = MsgBox("Are you sure you want to generate the print?", _
                          vbYesNoCancel + vbQuestion, _
                          "Confirm Action")
    
    If userResponse <> vbYes Then Exit Sub
    
    ShowDrawing = ShowDraw.Value
    
    DrawConductors = Conductor.Value
    DrawPrimary = Conductor1.Value And Conductor1.Enabled
    DrawSecondary = Conductor2.Value And Conductor2.Enabled
    DrawOpenWire = Conductor3.Value And Conductor3.Enabled
    DrawDeadends = Conductor4.Value And Conductor4.Enabled
    ConductorInitOffset = Conductor5.Value
    ConductorOffsetAmount = Conductor6.Value
    
    DrawROW = row.Value
    DrawCenterLines = ROW1.Value And ROW1.Enabled
    DrawCenterLineDistances = ROW1a.Value And ROW1a.Enabled
    
    DrawDownGuys = DG.Value
    DrawSpanGuys = SPG.Value
    DrawCrewNotes = CN.Value
    DrawTrees = Tree.Value
    
    DrawServices = Service.Value
    DrawUGServices = Service1.Value And Service1.Enabled
    DrawAdjacentServices = Service2.Value And Service2.Enabled
    
    DrawEquipment = equipment.Value
    DrawAdjacentPoleEquipment = Equipment1.Value And Equipment1.Enabled
    DrawTransformers = Equipment2.Value And Equipment2.Enabled
    DrawStreetlights = Equipment3.Value And Equipment3.Enabled
    DrawCapacitors = Equipment4.Value And Equipment4.Enabled
    DrawRegulators = Equipment5.Value And Equipment5.Enabled
    DrawIsolators = Equipment6.Value And Equipment6.Enabled
    DrawFuses = Equipment7.Value And Equipment7.Enabled
    DrawReclosers = Equipment8.Value And Equipment8.Enabled
    DrawSectionalizers = Equipment9.Value And Equipment9.Enabled
    DrawSwitches = Equipment10.Value And Equipment10.Enabled
    DrawSensors = Equipment11.Value And Equipment11.Enabled
    DrawSecondaryRisers = Equipment12.Value And Equipment12.Enabled
    DrawPrimaryRisers = Equipment13.Value And Equipment13.Enabled
    
    Dim varDict As Object
    Set varDict = CreateObject("Scripting.Dictionary")
    
    varDict.Add "ShowDrawing", ShowDraw.Value
    
    varDict.Add "DrawConductors", Conductor.Value
    varDict.Add "DrawPrimary", Conductor1.Value
    varDict.Add "DrawSecondary", Conductor2.Value
    varDict.Add "DrawOpenWire", Conductor3.Value
    varDict.Add "DrawDeadends", Conductor4.Value
    varDict.Add "ConductorInitOffset", Conductor5.Value
    varDict.Add "ConductorOffsetAmount", Conductor6.Value
    
    varDict.Add "DrawROW", row.Value
    varDict.Add "DrawCenterLines", ROW1.Value
    varDict.Add "DrawCenterLineDistances", ROW1a.Value
    
    varDict.Add "DrawDownGuys", DG.Value
    varDict.Add "DrawSpanGuys", SPG.Value
    varDict.Add "DrawCrewNotes", CN.Value
    varDict.Add "DrawTrees", Tree.Value
    
    varDict.Add "DrawServices", Service.Value
    varDict.Add "DrawUGServices", Service1.Value
    varDict.Add "DrawAdjacentServices", Service2.Value
    
    varDict.Add "DrawEquipment", equipment.Value
    varDict.Add "DrawAdjacentPoleEquipment", Equipment1.Value
    varDict.Add "DrawTransformers", Equipment2.Value
    varDict.Add "DrawStreetlights", Equipment3.Value
    varDict.Add "DrawCapacitors", Equipment4.Value
    varDict.Add "DrawRegulators", Equipment5.Value
    varDict.Add "DrawIsolators", Equipment6.Value
    varDict.Add "DrawFuses", Equipment7.Value
    varDict.Add "DrawReclosers", Equipment8.Value
    varDict.Add "DrawSectionalizers", Equipment9.Value
    varDict.Add "DrawSwitches", Equipment10.Value
    varDict.Add "DrawSensors", Equipment11.Value
    varDict.Add "DrawSecondaryRisers", Equipment12.Value
    varDict.Add "DrawPrimaryRisers", Equipment13.Value
    
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
            ctrl.Enabled = Conductor.Value
        End If
    Next ctrl
End Sub

Private Sub Service_Click()
    Dim ctrl As control
    For Each ctrl In Me.Controls
        If InStr(ctrl.name, "Service") > 0 And ctrl.name <> "Service" Then
            ctrl.Enabled = Service.Value
        End If
    Next ctrl
End Sub

Private Sub ROW_Click()
    For Each ctrl In Me.Controls
        If InStr(ctrl.name, "ROW") > 0 And ctrl.name <> "ROW" Then
            ctrl.Enabled = row.Value
        End If
    Next ctrl
End Sub

Private Sub ROW1_Click()
    For Each ctrl In Me.Controls
        If InStr(ctrl.name, "ROW1") > 0 And ctrl.name <> "ROW1" Then
            ctrl.Enabled = ROW1.Value
        End If
    Next ctrl
End Sub

Private Sub Equipment_Click()
    For Each ctrl In Me.Controls
        If InStr(ctrl.name, "Equipment") > 0 And ctrl.name <> "Equipment" Then
            ctrl.Enabled = equipment.Value
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
