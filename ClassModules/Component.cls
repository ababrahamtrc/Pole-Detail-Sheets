VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Component"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public midspans As Scripting.Dictionary
Public height As Integer
Public bottomHeight  As Integer
Public size As String
Public owner As String
Public componentType As String

Private Sub Class_Initialize()
    Set midspans = New Scripting.Dictionary
End Sub

Public Function getWire() As wire
    Dim wire As wire: Set wire = New wire
    Set wire.midspans = Me.midspans
    wire.height = Me.height
    wire.size = Me.size
    wire.owner = Me.owner
    wire.componentType = Me.componentType
    
    Set getWire = wire
End Function

Public Function getDownGuy() As guy
    Dim guy As guy: Set guy = New guy
    guy.height = Me.height
    guy.size = Me.size
    guy.owner = Me.owner
    guy.componentType = Me.componentType
    
    Set getDownGuy = guy
End Function

Public Function getEquipment() As equipment
    Dim equipment As equipment: Set equipment = New equipment
    equipment.height = Me.height
    equipment.bottomHeight = Me.bottomHeight
    equipment.size = Me.size
    equipment.owner = Me.owner
    equipment.componentType = Me.componentType
    
    Set getEquipment = equipment
End Function
