Attribute VB_Name = "WBSCC"
Dim sBar
Dim oName
Dim oName2

Sub UpdateWBSCCScript(ByRef WBS As String, ByRef status As String, ByRef trx As String, ByRef SAPConnection As Variant, ByRef CMail As Variant)
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
        
'        SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpGRND/ssubSUBSCR1:SAPLCJWB:1210/ctxtPRPS-FKOKR").Text = "1000"
        SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpGRND/ssubSUBSCR1:SAPLCJWB:1210/ctxtPRPS-FKSTL").Text = "2800001050"
        SAPConnection.session.findById("wnd[0]").sendVKey 0

                    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[11]").press ' Save if no popup window.
                    ActiveCell.Offset(0, 3).value = sBar.Text
                
    ' If there is a popup - loop through popup checks for Availability control, Scheduling, budget and any else pupop until there are no more popups.
    If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
        Do
            If InStr(SAPConnection.session.findById("wnd[1]").Text, "Availability Control") > 0 Then
                SAPConnection.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
            ElseIf InStr(SAPConnection.session.findById("wnd[1]").Text, "Scheduling") > 0 Then
                SAPConnection.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
            ElseIf InStr(SAPConnection.session.findById("wnd[1]").Text, "Commit") > 0 Then
                SAPConnection.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
            ElseIf InStr(SAPConnection.session.findById("wnd[1]").Text, "Cost") > 0 Then
                SAPConnection.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
            ElseIf InStr(SAPConnection.session.findById("wnd[1]").Text, "Budget") > 0 Then
                SAPConnection.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                budgetWarning = True
            Else
                SAPConnection.session.findById("wnd[1]").sendVKey 0
            End If
        Loop While Not SAPConnection.session.findById("wnd[1]", False) Is Nothing
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

