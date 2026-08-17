VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoadingBar_Form2 
   Caption         =   "Progress Bar"
   ClientHeight    =   1800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6780
   OleObjectBlob   =   "LoadingBar_Form2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LoadingBar_Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gTotal As Long
Public gTotal2 As Long
Public gCurrent As Long
Public gCurrent2 As Long

Sub InitProgress(Total As Long, Optional starting As Boolean, Optional Total2 As Integer)
    gTotal = Total
    gCurrent = 0
    If starting Then
        gCurrent2 = 0
        If Total2 = 0 Then
            gTotal2 = 10
        Else
            gTotal2 = Total2
        End If
    Else
        gCurrent2 = gCurrent2 + 1
    End If
    
    If Not Me.visible Then
        Me.StartUpPosition = 0
        Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
        Me.top = Application.top + (0.5 * Application.height) - (0.5 * Me.height)
    End If

    LoadingBar_Form.barFill.top = 0
    LoadingBar_Form.barFill.height = LoadingBar_Form.frmBar.height

    With LoadingBar_Form
        .barFill.Width = 0
        .Show vbModeless
    End With
End Sub

Sub UpdateProgress(fileName As String, progressType As String, Optional noFiles As Boolean)
    Dim pct As Double
    Dim maxWidth As Long

    gCurrent = gCurrent + 1

    maxWidth = LoadingBar_Form.frmBar.Width
    pct = gCurrent / gTotal

    LoadingBar_Form.Label1.caption = progressType & IIf(noFiles, "", " files") & "..." & gCurrent & "/" & gTotal
    LoadingBar_Form.Label2.caption = fileName
    LoadingBar_Form.Label3.caption = gCurrent2 & "/" & gTotal2
    
    If pct > 1 Then pct = 1
    LoadingBar_Form.barFill.Width = maxWidth * pct
    DoEvents
End Sub
 
Sub FinishProgress()
    Unload LoadingBar_Form
    gTotal = 0
    gTotal2 = 0
    gCurrent = 0
    gCurrent2 = 0
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Call LoadingBar_Form.FinishProgress
        MsgBox "Operation Canceled"
    End If
End Sub
 
