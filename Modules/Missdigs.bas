Attribute VB_Name = "Missdigs"
Public Sub MissDigsExport()
    Dim project As project: Set project = New project
    Call project.extractFromSheets
    Dim pole As pole
    Dim miss_dig_counter As Long
    miss_dig_counter = 0
    Dim output As String
    Dim DataObj As DataObject
    Set DataObj = New DataObject
    
    Call SendLogMessage("MissDigsExport")
    
    userChoice = MsgBox("Click YES for Normal Mode" & vbCrLf & _
                        "Click NO for Demo Mode", _
                        vbYesNo + vbQuestion, "Select Mode")
    
    'output = output & "Contact_EngineerName.value = " & Chr(34) & project.CoordinatorName & Chr(34) & vbLf
    
    'output = output & "EngineerPhone.value = " & Chr(34) & project.CoordinatorPhone & Chr(34) & vbLf
    'output = output & "Contact_FieldContact.value = " & Chr(34) & project.DesignerName & Chr(34) & vbLf
    'output = output & "Phone.value = " & Chr(34) & project.DesignerNumber & Chr(34) & vbLf
    
    For Each pole In project.poles
        If userChoice = vbNo Then
            pole.ReplaceAnchor = False
            pole.InstallAnchor = False
            pole.ReplacePole = False
            pole.InstallPole = False
            pole.ReplaceRiser = False
            pole.RemoveAnchor = False
            pole.RemovePole = False
            
            If pole.topped Then
                pole.RemovePole = True
            End If
        End If
    
        If (pole.ReplaceAnchor Or pole.InstallAnchor Or pole.ReplacePole Or pole.InstallPole Or pole.ReplaceRiser Or pole.RemoveAnchor Or pole.RemovePole) Then
            If (miss_dig_counter <> 0) Then
                output = output & "addjob.click()" & vbLf
            End If
            If (pole.ReplaceAnchor And pole.ReplacePole) Then
                output = output & "document.getElementById('work-type-adder-input-" & miss_dig_counter & "').value = ""REPLACE ANCHOR""" & vbLf
                output = output & "document.querySelector('input#work-type-adder-input-" & miss_dig_counter & ".form-control.form-control-sm.work-type-adder.ui-autocomplete-input').focus()" & vbLf
                output = output & "document.querySelector('input#work-type-adder-input-" & miss_dig_counter & ".form-control.form-control-sm.work-type-adder.ui-autocomplete-input').dispatchEvent(new Event('blur',{bubbles:true,cancelable:true}))" & vbLf
                output = output & "document.getElementById('work-type-adder-input-" & miss_dig_counter & "').value = ""REPLACE POLE""" & vbLf
                output = output & "JobLocation_" & miss_dig_counter & "__StakingInformation.value = ""The pole and guy/anchor at " & pole.address & " will be replaced. Please locate with paint and flags a 40' radius around existing pole to ensure safe installations of new pole and guy/anchor. The pole has been tagged with \""Consumers " & pole.existingCEID & ".\""""" & vbLf
            ElseIf (pole.InstallAnchor And pole.ReplacePole) Then
                output = output & "document.getElementById('work-type-adder-input-" & miss_dig_counter & "').value = ""INSTALL ANCHOR""" & vbLf
                output = output & "document.querySelector('input#work-type-adder-input-" & miss_dig_counter & ".form-control.form-control-sm.work-type-adder.ui-autocomplete-input').focus()" & vbLf
                output = output & "document.querySelector('input#work-type-adder-input-" & miss_dig_counter & ".form-control.form-control-sm.work-type-adder.ui-autocomplete-input').dispatchEvent(new Event('blur',{bubbles:true,cancelable:true}))" & vbLf
                output = output & "document.getElementById('work-type-adder-input-" & miss_dig_counter & "').value = ""REPLACE POLE""" & vbLf
                output = output & "JobLocation_" & miss_dig_counter & "__StakingInformation.value = ""The pole at " & pole.address & " will be replaced and have a new guy/anchor installed. Please locate with paint and flags a 40' radius around existing pole to ensure safe installations of new pole and guy/anchor. The pole has been tagged with \""Consumers " & pole.existingCEID & ".\""""" & vbLf
            ElseIf ((pole.InstallAnchor Or pole.ReplaceAnchor) And pole.InstallPole) Then
                output = output & "document.getElementById('work-type-adder-input-" & miss_dig_counter & "').value = ""INSTALL ANCHOR""" & vbLf
                output = output & "document.querySelector('input#work-type-adder-input-" & miss_dig_counter & ".form-control.form-control-sm.work-type-adder.ui-autocomplete-input').focus()" & vbLf
                output = output & "document.querySelector('input#work-type-adder-input-" & miss_dig_counter & ".form-control.form-control-sm.work-type-adder.ui-autocomplete-input').dispatchEvent(new Event('blur',{bubbles:true,cancelable:true}))" & vbLf
                output = output & "document.getElementById('work-type-adder-input-" & miss_dig_counter & "').value = ""INSTL POLE(S)""" & vbLf
                output = output & "JobLocation_" & miss_dig_counter & "__StakingInformation.value = ""A new pole and anchor at " & pole.address & " will be installed. Please locate with paint and flags a 40' radius around stake to ensure safe installations of new pole. The pole has not been tagged. Located between poles [A and B]""" & vbLf
            ElseIf (pole.InstallAnchor And pole.ReplaceRiser) Then
                output = output & "document.getElementById('work-type-adder-input-" & miss_dig_counter & "').value = ""INSTALL ANCHOR""" & vbLf
                output = output & "document.querySelector('input#work-type-adder-input-" & miss_dig_counter & ".form-control.form-control-sm.work-type-adder.ui-autocomplete-input').focus()" & vbLf
                output = output & "document.querySelector('input#work-type-adder-input-" & miss_dig_counter & ".form-control.form-control-sm.work-type-adder.ui-autocomplete-input').dispatchEvent(new Event('blur',{bubbles:true,cancelable:true}))" & vbLf
                output = output & "document.getElementById('work-type-adder-input-" & miss_dig_counter & "').value = ""REPLACE RISER""" & vbLf
                output = output & "JobLocation_" & miss_dig_counter & "__StakingInformation.value = ""The pole at " & pole.address & " will have its riser replaced and have a new guy/anchor installed. Please locate with paint and flags a 40' radius around existing pole to ensure safe installations of new riser and guy/anchor. The pole has been tagged with \""Consumers " & pole.existingCEID & ".\""""" & vbLf
            ElseIf (pole.ReplaceAnchor And pole.ReplaceRiser) Then
                output = output & "document.getElementById('work-type-adder-input-" & miss_dig_counter & "').value = ""REPLACE ANCHOR""" & vbLf
                output = output & "document.querySelector('input#work-type-adder-input-" & miss_dig_counter & ".form-control.form-control-sm.work-type-adder.ui-autocomplete-input').focus()" & vbLf
                output = output & "document.querySelector('input#work-type-adder-input-" & miss_dig_counter & ".form-control.form-control-sm.work-type-adder.ui-autocomplete-input').dispatchEvent(new Event('blur',{bubbles:true,cancelable:true}))" & vbLf
                output = output & "document.getElementById('work-type-adder-input-" & miss_dig_counter & "').value = ""REPLACE RISER""" & vbLf
                output = output & "JobLocation_" & miss_dig_counter & "__StakingInformation.value = ""The pole at " & pole.address & " will have its riser and guy/anchor replaced. Please locate with paint and flags a 40' radius around existing pole to ensure safe installations of new riser and guy/anchor. The pole has been tagged with \""Consumers " & pole.existingCEID & ".\""""" & vbLf
            ElseIf (pole.RemoveAnchor And pole.RemovePole) Then
                output = output & "document.getElementById('work-type-adder-input-" & miss_dig_counter & "').value = ""REMOVE ANCHOR""" & vbLf
                output = output & "document.querySelector('input#work-type-adder-input-" & miss_dig_counter & ".form-control.form-control-sm.work-type-adder.ui-autocomplete-input').focus()" & vbLf
                output = output & "document.querySelector('input#work-type-adder-input-" & miss_dig_counter & ".form-control.form-control-sm.work-type-adder.ui-autocomplete-input').dispatchEvent(new Event('blur',{bubbles:true,cancelable:true}))" & vbLf
                output = output & "document.getElementById('work-type-adder-input-" & miss_dig_counter & "').value = ""REMOVE POLE""" & vbLf
                output = output & "JobLocation_" & miss_dig_counter & "__StakingInformation.value = ""The pole and guy/anchor at " & pole.address & " will be removed. Please locate with paint and flags a 40' radius around existing pole and anchor to ensure safe removal of pole and guy/anchor. The pole has been tagged with \""Consumers " & pole.existingCEID & ".\""""" & vbLf
            ElseIf (pole.RemoveAnchor And pole.ReplacePole) Then
                output = output & "document.getElementById('work-type-adder-input-" & miss_dig_counter & "').value = ""REMOVE ANCHOR""" & vbLf
                output = output & "document.querySelector('input#work-type-adder-input-" & miss_dig_counter & ".form-control.form-control-sm.work-type-adder.ui-autocomplete-input').focus()" & vbLf
                output = output & "document.querySelector('input#work-type-adder-input-" & miss_dig_counter & ".form-control.form-control-sm.work-type-adder.ui-autocomplete-input').dispatchEvent(new Event('blur',{bubbles:true,cancelable:true}))" & vbLf
                output = output & "document.getElementById('work-type-adder-input-" & miss_dig_counter & "').value = ""REPLACE POLE""" & vbLf
                output = output & "JobLocation_" & miss_dig_counter & "__StakingInformation.value = ""The pole at " & pole.address & " will be replaced and have an existing guy/anchor removed. Please locate with paint & flags, a 40' radius around the existing pole and anchor to ensure a safe installation of pole and removal of guy/anchor. The pole is tagged with \""Consumers " & pole.existingCEID & ".\""""" & vbLf
            ElseIf (pole.RemoveAnchor And pole.ReplaceRiser) Then
                output = output & "document.getElementById('work-type-adder-input-" & miss_dig_counter & "').value = ""REMOVE ANCHOR""" & vbLf
                output = output & "document.querySelector('input#work-type-adder-input-" & miss_dig_counter & ".form-control.form-control-sm.work-type-adder.ui-autocomplete-input').focus()" & vbLf
                output = output & "document.querySelector('input#work-type-adder-input-" & miss_dig_counter & ".form-control.form-control-sm.work-type-adder.ui-autocomplete-input').dispatchEvent(new Event('blur',{bubbles:true,cancelable:true}))" & vbLf
                output = output & "document.getElementById('work-type-adder-input-" & miss_dig_counter & "').value = ""REPLACE RISER""" & vbLf
                output = output & "JobLocation_" & miss_dig_counter & "__StakingInformation.value = ""The underground riser on the pole at " & pole.address & " will be replaced and an existing guy/anchor will be removed. Please locate with paint and flags a 40' radius around existing pole to ensure safe installation of new riser and removal of guy/anchor. The pole has been tagged with \""Consumers " & pole.existingCEID & ".\""""" & vbLf
            ElseIf (pole.ReplaceAnchor) Then
                output = output & "document.getElementById('work-type-adder-input-" & miss_dig_counter & "').value = ""REPLACE ANCHOR""" & vbLf
                output = output & "JobLocation_" & miss_dig_counter & "__StakingInformation.value = ""The pole at " & pole.address & " will have its guy/anchor replaced. Please locate with paint and flags a 40' radius around existing anchor to ensure safe installations new of new guy/anchor. The pole has been tagged with \""Consumers " & pole.existingCEID & ".\""""" & vbLf
            ElseIf (pole.InstallAnchor) Then
                output = output & "document.getElementById('work-type-adder-input-" & miss_dig_counter & "').value = ""INSTALL ANCHOR""" & vbLf
                output = output & "JobLocation_" & miss_dig_counter & "__StakingInformation.value = ""The pole at " & pole.address & " will have a new guy/anchor installed. Please locate with paint and flags a 40' radius around stake to ensure safe installations of new guy/anchor. The pole has been tagged with \""Consumers " & pole.existingCEID & ".\""""" & vbLf
            ElseIf (pole.ReplacePole) Then
                output = output & "document.getElementById('work-type-adder-input-" & miss_dig_counter & "').value = ""REPLACE POLE""" & vbLf
                output = output & "JobLocation_" & miss_dig_counter & "__StakingInformation.value = ""The pole at " & pole.address & " will be replaced. Please locate with paint and flags a 40' radius around existing pole to ensure safe installations of new pole. The pole has been tagged with \""Consumers " & pole.existingCEID & ".\""""" & vbLf
            ElseIf (pole.InstallPole) Then
                output = output & "document.getElementById('work-type-adder-input-" & miss_dig_counter & "').value = ""INSTL POLE(S)""" & vbLf
                output = output & "JobLocation_" & miss_dig_counter & "__StakingInformation.value = ""A new pole at " & pole.address & " will be installed. Please locate with paint and flags a 40' radius around stake to ensure safe installations of new pole. The pole has not been tagged. Located between poles [A and B]""" & vbLf
            ElseIf (pole.ReplaceRiser) Then
                output = output & "document.getElementById('work-type-adder-input-" & miss_dig_counter & "').value = ""REPLACE RISER""" & vbLf
                output = output & "JobLocation_" & miss_dig_counter & "__StakingInformation.value = ""The pole at " & pole.address & " will have its riser replaced. Please locate with paint and flags a 40' radius around existing pole to ensure safe installations of new riser. The pole has been tagged with \""Consumers " & pole.existingCEID & ".\""""" & vbLf
            ElseIf (pole.RemovePole) Then
                output = output & "document.getElementById('work-type-adder-input-" & miss_dig_counter & "').value = ""REMOVE POLE""" & vbLf
                output = output & "JobLocation_" & miss_dig_counter & "__StakingInformation.value = ""The pole at " & pole.address & " will be removed. Please locate with paint & flags, a 40' radius around the existing pole to ensure a safe removal. The pole is tagged with \""Consumers " & pole.existingCEID & ".\""""" & vbLf
            ElseIf (pole.RemoveAnchor) Then
                output = output & "document.getElementById('work-type-adder-input-" & miss_dig_counter & "').value = ""REMOVE ANCHOR""" & vbLf
                output = output & "JobLocation_" & miss_dig_counter & "__StakingInformation.value = ""A guy/anchor for the pole at " & pole.address & " will be removed. Please Locate with paint & flags, a 40' radius around the existing pole and anchor to ensure a safe removal. The pole is tagged with \""Consumers " & pole.existingCEID & ".\""""" & vbLf
            End If
   
            If pole.latitude <> 0 And pole.longitude <> 0 Then
                output = Left(output, Len(output) - 2)
                output = output & "\nGPS: Latitude: " & pole.latitude & ", Longitude: " & pole.longitude & """" & vbLf
            End If
            
            output = output & "document.querySelector('input#work-type-adder-input-" & miss_dig_counter & ".form-control.form-control-sm.work-type-adder.ui-autocomplete-input').focus()" & vbLf
            output = output & "document.querySelector('input#work-type-adder-input-" & miss_dig_counter & ".form-control.form-control-sm.work-type-adder.ui-autocomplete-input').dispatchEvent(new Event('blur',{bubbles:true,cancelable:true}))" & vbLf
            output = output & "PrintLocations" & miss_dig_counter & ".value = " & Chr(34) & pole.location & Chr(34) & vbLf
            output = output & "JobLocation_" & miss_dig_counter & "__FromAddress.value = " & Chr(34) & pole.address & Chr(34) & vbLf
            output = output & "JobLocation_" & miss_dig_counter & "__ToAddress.value = " & Chr(34) & pole.address & Chr(34) & vbLf
            output = output & "JobLocation_" & miss_dig_counter & "__FirstCrossStreet.value = " & Chr(34) & pole.firstCrossStreet & Chr(34) & vbLf
            output = output & "JobLocation_" & miss_dig_counter & "__SecondCrossStreet.value = ""n/a""" & vbLf
            
            miss_dig_counter = miss_dig_counter + 1
        End If
    Next pole
    
    DataObj.SetText output
    DataObj.PutInClipboard
    
    MsgBox "Copied code, now go to Missdigs and press f12 for console and paste it in."
    
End Sub

