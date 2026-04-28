Attribute VB_Name = "NJUNS"
Public Sub createNJUNS()
        Call LogMessage.SendLogMessage("createNJUNS")
        
        Set sheet = Application.ActiveSheet
        If sheet.name = "4 Spans" Or sheet.name = "8 Spans" Or sheet.name = "12 Spans" Or sheet.Cells(2, 2).Value <> "Notification:" Then
            MsgBox "You need to have a pole detail sheet active to run this script."
            Exit Sub
        End If

        Unload NJUNS_Form
        Call NJUNS_Form.Initialize
        NJUNS_Form.Show vbModeless
End Sub

Public Function moveComms(ByVal index As Long, ByVal comms As Collection, ByVal movements As Collection, applicant As Boolean, ApplyAbove As Boolean) As String
    Dim movement As String
    If index > comms.count Then
        moveComms = ""
        Exit Function
    End If
    
    Dim Comm As Comm
    Set Comm = comms(index)
    
    Dim drops As String
    If Comm.drops Then
        drops = " and attach drops to mainline "
    End If
    
    movement = Comm.owner & vbCrLf
    If Comm.modification = "" Then Comm.modification = Comm.height
    If movements(index) = "Raise" Or movements(index) = "Lower" Then
        If Comm.Bracket Then
            movement = movement & "to remove [from] standoff bracket and attach to pole to correct illegal attachment violation. "
        End If
        If Comm.Boxed Then
            movement = movement & "to correct boxing violation. "
        End If
    End If
    If movements(index) = "Lower" Then
        If Comm.Mainline Then
            movement = movement & "Lower " & Comm.orientation & " mainline a minimum of " & findDifference(Comm.height, Comm.modification) & " on the pole" & IIf(Comm.drops, drops, " ") & "to " & Comm.reasonForMovement
        Else
            movement = movement & "Lower drop a minimum of " & findDifference(Comm.height, Comm.modification) & " on the pole" & IIf(Comm.drops, drops, " ") & "to " & Comm.reasonForMovement
        End If
        If (index + 1 > comms.count) Then
            movement = movement & "Maintain minimum " & IIf(Not applicant Or ApplyAbove, "15'6""", "16'0""") & " midspan ground clearance." & vbCrLf
        Else
            movement = movement & "Maintain minimum 12"" comm separation on the pole and 6"" separation at the midspan above " & IIf(comms(index + 1).owner = comms(index).owner, "other " & comms(index + 1).owner & " mainline", comms(index + 1).owner) & "." & vbCrLf
        End If
    ElseIf movements(index) = "Raise" Then
        If Comm.Mainline Then
            movement = movement & "Raise " & Comm.orientation & " mainline a minimum of " & findDifference(Comm.height, Comm.modification) & " on the pole" & IIf(Comm.drops, drops, " ") & "to " & Comm.reasonForMovement
        Else
            movement = movement & "Raise drop a minimum of " & findDifference(Comm.height, Comm.modification) & " on the pole" & IIf(Comm.drops, drops, " ") & "to " & Comm.reasonForMovement
        End If
        If (index = 1) Then
            movement = movement & "Maintain minimum " & IIf(applicant And ApplyAbove, "52""", "40""") & " pole separation and " & IIf(applicant And ApplyAbove, "36""", "30""") & " midspan separation below lowest power." & vbCrLf
        Else
            movement = movement & "Maintain minimum 12"" comm separation on the pole and 6"" separation at the midspan below " & IIf(comms(index - 1).owner = comms(index).owner, "other " & comms(index - 1).owner & " mainline", comms(index - 1).owner) & "." & vbCrLf
        End If
    ElseIf movements(index) = "Attach" Then
        movement = movement & "Attach to pole. "
        If (index = 1) Then
            movement = movement & "Maintain minimum " & IIf(applicant And ApplyAbove, "52""", "40""") & " pole separation and " & IIf(applicant And ApplyAbove, "36""", "30""") & " midspan separation below lowest power. "
        End If
        If (index > 1) Then
            movement = movement & "Maintain minimum 12"" comm separation on the pole and 6"" separation at the midspan below " & IIf(comms(index - 1).owner = comms(index).owner, "other " & comms(index - 1).owner & " mainline", comms(index - 1).owner) & ". "
        End If
        movement = movement & "Maintain minimum " & Utilities.inchesToFeetInches(((movements.count - index) * 6) + 186 + IIf(Not applicant Or ApplyAbove, 0, 6)) & " midspan ground clearance." & vbCrLf
    End If
    If movements(index) = "Raise" Or movements(index) = "Lower" Then
        Dim othercomm As Comm
        For Each othercomm In comms
            If Not Comm.Equals(othercomm) And Comm.Mainline And othercomm.Mainline And Not othercomm.NotAttached Then
                If Utilities.convertToInches(Comm.height) < convertToInches(othercomm.height) And Utilities.convertToInches(Comm.modification) > Utilities.convertToInches(othercomm.modification) Then
                    movement = Left(movement, Len(movement) - 2) & " Coordinate with " & othercomm.owner & " to change attach orientation." & vbCrLf
                ElseIf Utilities.convertToInches(Comm.height) > convertToInches(othercomm.height) And Utilities.convertToInches(Comm.modification) < Utilities.convertToInches(othercomm.modification) Then
                    movement = Left(movement, Len(movement) - 2) & " Coordinate with " & othercomm.owner & " to change attach orientation." & vbCrLf
                End If
            End If
        Next othercomm
    End If
    
    If movement = Comm.owner & vbCrLf And Comm.drops Then
        movement = movement & "Attach drops to mainline to " & Comm.reasonForMovement & vbCrLf
    End If
    
    If Comm.Bracket And (movements(index) <> "Raise" And movements(index) <> "Lower") Then
        If Right(movement, 3) = "." & vbCrLf Then
            movement = Left(movement, Len(movement) - 2) & " "
        End If
        movement = movement & "to remove [from] standoff bracket and attach to pole to correct illegal attachment violation. "
        If Comm.Boxed Then
            movement = movement & "to correct boxing violation. "
        End If
        If (index < comms.count) Then
            movement = movement & "Maintain minimum 12"" comm separation on the pole and 6"" separation at the midspan above " & IIf(comms(index + 1).owner = comms(index).owner, "other " & comms(index + 1).owner & " mainline", comms(index + 1).owner) & "." & vbCrLf
        Else
            movement = movement & "Maintain minimum " & IIf(Not applicant Or ApplyAbove, "15'6""", "16'0""") & " midspan ground clearance." & vbCrLf
        End If
    ElseIf Comm.Boxed And (movements(index) <> "Raise" And movements(index) <> "Lower") Then
        If Right(movement, 3) = "." & vbCrLf Then
            movement = Left(movement, Len(movement) - 2) & " "
        End If
        movement = movement & "to correct boxing violation. "
        If (index < comms.count) Then
            movement = movement & "Maintain minimum 12"" comm separation on the pole and 6"" separation at the midspan above " & IIf(comms(index + 1).owner = comms(index).owner, "other " & comms(index + 1).owner & " mainline", comms(index + 1).owner) & "." & vbCrLf
        Else
            movement = movement & "Maintain minimum " & IIf(Not applicant Or ApplyAbove, "15'6""", "16'0""") & " midspan ground clearance." & vbCrLf
        End If
    End If
    
    If Comm.UpgradeGuy Then
        If Right(movement, 3) = "." & vbCrLf Then
            movement = Left(movement, Len(movement) - 2) & " "
        End If
        movement = movement & "Replace 6M guy with 10M guy to correct pole loading failure." & vbCrLf
    End If
    
    If Comm.InstallAnchor Then
        If Right(movement, 3) = "." & vbCrLf Then
            movement = Left(movement, Len(movement) - 2) & " "
        End If
        movement = movement & "Install [6M/10M] guy/anchor " & IIf(Comm.NextAnchor > 0, Comm.NextAnchor & "'", "[INSERT ANCHOR DISTANCE]") & " to the [DIRECTION] to support unsupported [span/angle]. Maintain a minimum 3' anchor separation and 5' pole separation. Do not cross other down guys." & vbCrLf
    End If
    
    If Comm.MoveAnchor Then
        If Comm.NextAnchor <> Comm.PreviousAnchor Then
            If Right(movement, 3) = "." & vbCrLf Then
                movement = Left(movement, Len(movement) - 2) & " "
            End If
            If Comm.NextAnchor > Comm.PreviousAnchor Then
                movement = movement & "Extend guy/anchor to a minimum of " & Comm.NextAnchor & "' away from the pole to " & Comm.reasonForAnchorMovement & "Maintain a minimum 3' anchor separation. Do not cross other down guys." & vbCrLf
            ElseIf Comm.NextAnchor < Comm.PreviousAnchor Then
                movement = movement & "Retract guy/anchor to a maximum of " & Comm.NextAnchor & "' away from the pole to " & Comm.reasonForAnchorMovement & "Maintain a minimum 3' anchor separation and 5' pole separation. Do not cross other down guys." & vbCrLf
            End If
        End If
    End If
    
    If movements(index) = "Raise" Then
        moveComms = movement & vbCrLf & moveComms(index + 1, comms, movements, applicant, ApplyAbove)
    ElseIf movements(index) = "Lower" Then
        If index < movements.count Then
            If movements(index + 1) = "Raise" Then
                moveComms = movement & vbCrLf & moveComms(index + 1, comms, movements, applicant, ApplyAbove)
            Else
                moveComms = moveComms(index + 1, comms, movements, applicant, ApplyAbove) & movement & vbCrLf
            End If
        Else
            moveComms = moveComms(index + 1, comms, movements, applicant, ApplyAbove) & movement & vbCrLf
        End If
    ElseIf movements(index) = "Nothing" And movement <> Comm.owner & vbCrLf Then
        moveComms = movement & vbCrLf & moveComms(index + 1, comms, movements, applicant, ApplyAbove)
    ElseIf movements(index) = "Nothing" Then
        moveComms = moveComms(index + 1, comms, movements, applicant, ApplyAbove)
    ElseIf movements(index) = "Attach" Then
        moveComms = movement & vbCrLf & moveComms(index + 1, comms, movements, applicant, ApplyAbove)
    End If

End Function

Public Function topPole(ByVal comms As Collection, applicant As Boolean, ApplyAbove As Boolean) As String
    If comms.count < 1 Then
        topPole = ""
        Exit Function
    End If
    Dim movement As String
    movement = "Consumers to complete required work." & vbCrLf & vbCrLf
    
    Dim clearanceDict As Object
    Set clearanceDict = CreateObject("Scripting.Dictionary")
    clearanceDict.Add 1, "15'6"""
    clearanceDict.Add 2, "16'0"""
    clearanceDict.Add 3, "16'6"""
    clearanceDict.Add 4, "17'0"""
    clearanceDict.Add 5, "17'6"""
    clearanceDict.Add 6, "18'0"""
    clearanceDict.Add 7, "18'6"""
    clearanceDict.Add 8, "19'0"""
    clearanceDict.Add 9, "19'6"""
    
    Dim previousOwner As String
    previousOwner = ""
    For i = 1 To comms.count
        Dim Comm As Comm
        Set Comm = comms(i)
        Dim drops As String
        If Comm.drops Then
            drops = " and attach drops to mainline "
        End If
        If i = 1 Then
            If applicant Then
                If ApplyAbove Then
                    movement = movement & Comm.owner & vbCrLf & "To transfer " & Comm.orientation & " mainline to new pole" & IIf(Comm.drops, drops, " ") & "with a minimum 52"" safety zone separation on the pole and 36"" separation at the midspan below lowest power. Maintain minimum " & clearanceDict(comms.count) & " midspan ground clearance." & vbCrLf
                Else
                    movement = movement & Comm.owner & vbCrLf & "To transfer " & Comm.orientation & " mainline to new pole" & IIf(Comm.drops, drops, " ") & "with a minimum 40"" safety zone separation on the pole and 30"" separation at the midspan below lowest power. Maintain minimum " & clearanceDict(comms.count + 1) & " midspan ground clearance." & vbCrLf
                End If
                previousOwner = Comm.owner
            Else
                movement = movement & Comm.owner & vbCrLf & "To transfer " & Comm.orientation & " mainline to new pole" & IIf(Comm.drops, drops, " ") & "with a minimum 40"" safety zone separation on the pole and 30"" separation at the midspan below lowest power. Maintain minimum " & clearanceDict(comms.count - i + IIf(applicant And Not ApplyAbove, 2, 1)) & " midspan ground clearance." & vbCrLf
                previousOwner = Comm.owner
            End If
        Else
            movement = movement & Comm.owner & vbCrLf & "To transfer " & Comm.orientation & " mainline to new pole" & IIf(Comm.drops, drops, " ") & "with a minimum 12"" comm separation on the pole and 6"" separation at the midspan below " & previousOwner & ". Maintain minimum " & clearanceDict(comms.count - i + IIf(applicant And Not ApplyAbove, 2, 1)) & " midspan ground clearance." & vbCrLf
            previousOwner = Comm.owner
        End If
        If Not Comm.Mainline Then
            movement = Replace(movement, "mainline", "drops")
        End If
        
        If Comm.UpgradeGuy Then
            If Right(movement, 3) = "." & vbCrLf Then
                movement = Left(movement, Len(movement) - 2) & " "
            End If
        movement = movement & " Replace 6M guy with 10M guy to correct pole loading failure." & vbCrLf
        End If
    
        If Comm.MoveAnchor Then
            If Comm.NextAnchor <> Comm.PreviousAnchor Then
                If Right(movement, 3) = "." & vbCrLf Then
                    movement = Left(movement, Len(movement) - 2) & " "
                End If
                If Comm.NextAnchor > Comm.PreviousAnchor Then
                    movement = movement & "Extend guy/anchor to a minimum of " & Comm.NextAnchor & "' away from the pole to " & Comm.reasonForAnchorMovement & "Maintain a minimum 3' anchor separation. Do not cross other down guys." & vbCrLf
                ElseIf Comm.NextAnchor < Comm.PreviousAnchor Then
                    movement = movement & "Retract guy/anchor to a maximum of " & Comm.NextAnchor & "' away from the pole to " & Comm.reasonForAnchorMovement & "Maintain a minimum 3' anchor separation and 5' pole separation. Do not cross other down guys." & vbCrLf
                End If
            End If
        End If
        
        If Comm.InstallAnchor Then
            If Right(movement, 3) = "." & vbCrLf Then
                movement = Left(movement, Len(movement) - 2) & " "
            End If
            movement = movement & "Install [6M/10M] guy/anchor to the [DIRECTION] to support unsupported [span/angle]." & vbCrLf
        End If
        movement = movement & vbCrLf
    Next i
    
    movement = movement & "Consumers after comms transfer to new pole, pull topped pole."

    topPole = movement
End Function

Public Function findDifference(ByVal Height1 As String, ByVal Height2 As String) As String
    Dim inches1 As Integer
    Dim inches2 As Integer
    
    inches1 = Utilities.convertToInches(Height1)
    inches2 = Utilities.convertToInches(Height2)
    
    
    findDifference = Utilities.inchesToFeetInches(Abs(inches1 - inches2))
End Function

Public Function sortComms(ByVal comms As Collection) As Collection
    Dim sortedcomms As New Collection
    Dim itemToInsert As Variant
    Dim inserted As Boolean
    Dim i As Long
 
    For Each itemToInsert In comms
        inserted = False
        If Not itemToInsert.Ignored Then
           For i = 1 To sortedcomms.count
               Dim sortedHeight As Integer
               Dim itemToInsertHeight As Integer
               If Utilities.convertToInches(sortedcomms(i).modification) > 0 Then
                   sortedHeight = Utilities.convertToInches(sortedcomms(i).modification)
               Else
                   sortedHeight = Utilities.convertToInches(sortedcomms(i).height)
               End If
               If Utilities.convertToInches(itemToInsert.modification) > 0 Then
                   itemToInsertHeight = Utilities.convertToInches(itemToInsert.modification)
               Else
                   itemToInsertHeight = Utilities.convertToInches(itemToInsert.height)
               End If
               If itemToInsertHeight > sortedHeight Then
                   sortedcomms.Add itemToInsert, , i
                   inserted = True
                   Exit For
               End If
           Next i
    
           If Not inserted Then
               sortedcomms.Add itemToInsert
           End If
        End If
    Next itemToInsert
 
    Set sortComms = sortedcomms
End Function
