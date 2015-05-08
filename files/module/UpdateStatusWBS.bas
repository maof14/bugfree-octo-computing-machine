Attribute VB_Name = "UpdateStatusWBS"
Dim sBar
Dim oName
Dim oName2

Sub UpdateStatusWBSScript(ByRef WBS As String, ByRef status As String, ByRef trx As String, ByRef SAPConnection As Variant, ByRef CMail As Variant)
    On Error GoTo ErrHandler

        Set sBar = SAPConnection.session.findById("wnd[0]/sbar")
        
        SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").topNode = "         23"
        SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell").pressButton "OPEN"
        SAPConnection.session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PROJ_EXT").Text = ""
        SAPConnection.session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PRPS_EXT").Text = WBS
        SAPConnection.session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-AUFNR").Text = ""
        SAPConnection.session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PRPS_EXT").SetFocus
        SAPConnection.session.findById("wnd[1]").sendVKey 0
        
        If sBar.Text = "Not all objects were locked (see lock log)" Then
            ActiveCell.Offset(0, 3) = "WBS is being processed by another user, try later."
            SAPConnection.session.findById("wnd[0]/tbar[0]/btn[3]").press
            GoTo Redanstangdhoppaover
        End If
        
        ' Check if change mode is active
        If Not SAPConnection.session.ActiveWindow.FindByName("PRPS-POST1", GuiTextField).Changeable = True Then
            SAPConnection.session.findById("wnd[0]/tbar[1]/btn[13]").press
        End If

        Set oName = SAPConnection.session.ActiveWindow.FindByName("CNJ_STAT-STTXT_INT", "GuiTextField")
        Set oName2 = SAPConnection.session.ActiveWindow.FindByName("CNJ_STAT-STTXT_EXT", "GuiTextField")
                
        If status = "Set CLSD" Then
            If Not InStr(oName.Text, "CLSD") > 0 Then
                SAPConnection.session.findById("wnd[0]/mbar/menu[1]/menu[2]/menu[6]/menu[0]").Select ' <-- Set CLSD
                If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
                    GoTo StatusErr
                Else
                    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[11]").press ' Save if no popup window.
                    ActiveCell.Offset(0, 3).value = sBar.Text
                End If
            Else
                SAPConnection.session.findById("wnd[0]/tbar[0]/btn[3]").press ' back
                ActiveCell.Offset(0, 3).value = "Status already set"
                GoTo Redanstangdhoppaover
            End If
        ElseIf status = "Set TECO" Then
            If Not InStr(oName.Text, "TECO") > 0 And InStr(oName.Text, "CLSD") = 0 Then
                SAPConnection.session.findById("wnd[0]/mbar/menu[1]/menu[2]/menu[4]/menu[0]").Select
                If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
                    
                    SAPConnection.session.findById("wnd[1]/usr/btnOPTION1").press
                    ActiveCell.Offset(0, 3).value = "ERROR: " & SAPConnection.session.findById("wnd[2]/usr/lbl[9,3]").Text
                    SAPConnection.session.findById("wnd[2]/tbar[0]/btn[0]").press
                    SAPConnection.session.findById("wnd[1]/usr/btnOPTION2").press
                    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[3]").press
                    ActiveCell.value = 0

                    Exit Sub
                    
                Else
                    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[11]").press
                    ActiveCell.Offset(0, 3).value = sBar.Text
                End If
            Else
                If InStr(oName.Text, "CLSD") > 0 Then
                ActiveCell.Offset(0, 3).value = "WBS Already closed"
                Else
                If InStr(oName.Text, "TECO") > 0 Then
                ActiveCell.Offset(0, 3).value = "Status already set"
                End If
                End If
                SAPConnection.session.findById("wnd[0]/tbar[0]/btn[3]").press ' back
                GoTo Redanstangdhoppaover
            End If
        ElseIf status = "Set FNBL" Then
            If Not InStr(oName.Text, "FNBL") > 0 Then
                SAPConnection.session.findById("wnd[0]/mbar/menu[1]/menu[2]/menu[5]/menu[0]").Select
                If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
                    Do
                        SAPConnection.session.findById("wnd[1]").Close
                    Loop While Not SAPConnection.session.findById("wnd[1]") Is Nothing
                    ActiveCell.Offset(0, 3).value = "Information window upon status change attempt, skipping this WBS."
                    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[3]").press ' Back
                Else
                    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[11]").press
            If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
                Do
                    If InStr(SAPConnection.session.findById("wnd[1]").Text, "Cost Calculation") > 0 Then
                        SAPConnection.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                    ElseIf InStr(SAPConnection.session.findById("wnd[1]").Text, "Scheduling") > 0 Then
                        SAPConnection.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                    ElseIf InStr(SAPConnection.session.findById("wnd[1]").Text, "Budget") > 0 Then
                        SAPConnection.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                    Else
                        SAPConnection.session.findById("wnd[1]").sendVKey 0
                    End If
                Loop While Not SAPConnection.session.findById("wnd[1]", False) Is Nothing
            End If
                    ActiveCell.Offset(0, 3).value = sBar.Text
                End If
            Else
                SAPConnection.session.findById("wnd[0]/tbar[0]/btn[3]").press ' back
                ActiveCell.Offset(0, 3).value = "Status already set"
                GoTo Redanstangdhoppaover
            End If
        ElseIf status = "Set Release" Then
                If InStr(oName2.Text, "ZLIQ") > 0 Then
                ActiveCell.Offset(0, 3).value = "ERROR: No settlement is created on WBS, cannot be released"
                Exit Sub
                Else
                End If
            If Not InStr(oName.Text, "REL") > 0 Then
                SAPConnection.session.findById("wnd[0]/mbar/menu[1]/menu[2]/menu[0]").Select
                If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
                    Do
                        SAPConnection.session.findById("wnd[1]").Close
                    Loop While Not SAPConnection.session.findById("wnd[1]") Is Nothing
                    ActiveCell.Offset(0, 3).value = "Information window upon status change attempt, skipping this WBS."
                    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[3]").press ' Back
                Else
                    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[11]").press
                    ActiveCell.Offset(0, 3).value = sBar.Text
                End If
            Else
                SAPConnection.session.findById("wnd[0]/tbar[0]/btn[3]").press ' back
                ActiveCell.Offset(0, 3).value = "Status already set"
                GoTo Redanstangdhoppaover
            End If
        ElseIf status = "Remove CLSD" Then
            If Not InStr(oName.Text, "CLSD") < 1 Then
                SAPConnection.session.findById("wnd[0]/mbar/menu[1]/menu[2]/menu[6]/menu[1]").Select ' <-- Set CLSD
                If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
                    ActiveCell.Offset(0, 3) = "Information window upon status change attempt, skipping this WBS."
                    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[3]").press ' Back
                Else
                    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[11]").press ' Save if no popup window.
                    ActiveCell.Offset(0, 3).value = sBar.Text
                End If
            Else
                SAPConnection.session.findById("wnd[0]/tbar[0]/btn[3]").press ' back
                ActiveCell.Offset(0, 3).value = "Status already set"
                GoTo Redanstangdhoppaover
            End If
        ElseIf status = "Remove TECO" Then
            If Not InStr(oName.Text, "TECO") < 1 Then
                SAPConnection.session.findById("wnd[0]/mbar/menu[1]/menu[2]/menu[4]/menu[1]").Select ' <-- Remove TECO
                If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
                    Do
                        SAPConnection.session.findById("wnd[1]").Close
                    Loop While Not SAPConnection.session.findById("wnd[1]", False) Is Nothing
                    ActiveCell.Offset(0, 3) = "Information window upon status change attempt, skipping this WBS."
                    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[3]").press ' Back
                Else
                    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[11]").press ' Save if no popup window.
                    ActiveCell.Offset(0, 3).value = sBar.Text
                End If
            Else
                SAPConnection.session.findById("wnd[0]/tbar[0]/btn[3]").press ' back
                ActiveCell.Offset(0, 3).value = "Status already set"
                GoTo Redanstangdhoppaover
            End If
        ElseIf status = "Remove FNBL" Then
            If Not InStr(oName.Text, "FNBL") < 1 Then
                SAPConnection.session.findById("wnd[0]/mbar/menu[1]/menu[2]/menu[5]/menu[1]").Select ' <-- REMOVE FNBL
                If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
                    Do
                        SAPConnection.session.findById("wnd[1]").Close
                    Loop While Not SAPConnection.session.findById("wnd[1]", False) Is Nothing
                    ActiveCell.Offset(0, 3) = "Information window upon status change attempt, skipping this WBS."
                    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[3]").press ' Back
                Else
                    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[11]").press ' Save if no popup window.
                    ActiveCell.Offset(0, 3).value = sBar.Text
                End If
            Else
                SAPConnection.session.findById("wnd[0]/tbar[0]/btn[3]").press ' back
                ActiveCell.Offset(0, 3).value = "Status already set"
                GoTo Redanstangdhoppaover
            End If
            Else
        SAPConnection.session.findById("wnd[0]/tbar[0]/btn[3]").press
        ActiveCell.Offset(0, 3).value = "No valid action chosen, no change"
        End If
        GoTo Redanstangdhoppaover:
        
    
Redanstangdhoppaover:
        ActiveCell.value = 1
        Exit Sub
StatusErr:
    If SAPConnection.session.findById("wnd[1]/usr/btnOPTION1").Text = "Status informatio" Then
    SAPConnection.session.findById("wnd[1]/usr/btnOPTION1").press ' press button to see error list
    Else
    SAPConnection.session.findById("wnd[1]/usr/btnOPTION3").press
    End If
    Dim str As String
    Dim counter As Integer

    counter = 0
    str = "Errors on this WBS: "
    For Each E In SAPConnection.session.findById("wnd[2]/usr/").Children ' error row
        If (counter Mod 4 = 2) And (counter > 4) Then
            str = str & E.Text & ", "
        End If
        counter = counter + 1
    Next
    str = Replace(str, WBS & " ", "")
    str = Replace(str, WBS, "")
       
    SAPConnection.session.findById("wnd[2]/tbar[0]/btn[0]").press ' close err list windows
    SAPConnection.session.findById("wnd[1]/usr/btnOPTION2").press ' Close "errors have occured" window
    ' only wnd[0] open from here
    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[3]").press
        ActiveCell.Offset(0, 3).value = str
    Exit Sub

Fortsaett:
    Exit Sub

ErrHandler:
    If Err.Number <> 0 Then
        GoTo ErrGoNext
    End If
    Exit Sub
ErrGoNext:
    SAPConnection.ErrorCounter = SAPConnection.ErrorCounter + 1
    'SAPConnection.reportError "UpdateStatusWBS", ActiveCell.Offset(0, 1), Err.Number, Err.Description, Err.Source, sBar.Text
    CMail.BuildErrorList ActiveCell.Offset(0, 1), "UpdateStatusWBS", Err.Number, Err.Description, Err.Source, sBar.Text
    Err.Clear
    SAPConnection.errorContinueNextItem (trx)
End Sub
