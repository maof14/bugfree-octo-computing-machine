Attribute VB_Name = "POC_Milestone_Creation_Update"
Sub POC_Milestone_Creation_UpdateScript(ByRef WBS As String, ByRef MilestoneDesc As String, ByRef ActDate As String, ByRef POC As String, ByRef Supplink As String, ByRef trx As String, ByRef SAPConnection As Variant, ByRef CMail As Variant)
    Dim sBar
    Dim oName
    Dim MileDesc
    Set sBar = SAPConnection.session.findById("wnd[0]/sbar")
    
    On Error GoTo ErrHandler
    
    If MilestoneDesc = "" Then
    MileDesc = POC & "% POC as per POD " & "- " & Format(Now(), "YYMM") & "A"
    ElseIf MilestoneDesc = "POD" Then
    MileDesc = POD & "% POC as per POD " & "- " & Format(Now(), "YYMM") & "A"
    Else
    MileDesc = POC & "% POC requested by " & UCase(MilestoneDesc) & " - " & Format(Now(), "YYMM") & "A"
    End If
    
    SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell").pressButton "OPEN"
    SAPConnection.session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PROJ_EXT").Text = ""
    SAPConnection.session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PRPS_EXT").Text = WBS
    SAPConnection.session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-AUFNR").Text = ""
    SAPConnection.session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PRPS_EXT").SetFocus
    SAPConnection.session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PRPS_EXT").caretPosition = 18
    SAPConnection.session.findById("wnd[1]/tbar[0]/btn[0]").press
    If Not SAPConnection.session.ActiveWindow.FindByName("PRPS-POST1", GuiTextField).Changeable = True Then
        SAPConnection.session.findById("wnd[0]/tbar[1]/btn[13]").press
    End If
        If InStr(sBar.Text, "locked") > 0 Then
            ActiveCell.value = 0
            Cells(ActiveCell.Row, 8).value = "WBS BLOCKED by another user, try again later."
            Exit Sub
        End If
            
            Set oName = SAPConnection.session.ActiveWindow.FindByName("CNJ_STAT-STTXT_INT", "GuiTextField")

            If InStr(oName.Text, "CLSD") > 0 Then
                        ActiveCell.value = 1
            Cells(ActiveCell.Row, 8).value = "WBS is closed, cannot maintain."
            Exit Sub
            End If

    SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").SelectedNode = "         19"
    SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").DoubleClickNode "         19"
    SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1020/ssubVIEW_AREA:SAPLCJWB:3801/ctxtMLST-MLSTN").Text = "ZPOC1"
    SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1020/ssubVIEW_AREA:SAPLCJWB:3801/ctxtMLST-MLSTN").SetFocus
    SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1020/ssubVIEW_AREA:SAPLCJWB:3801/ctxtMLST-MLSTN").caretPosition = 5
    SAPConnection.session.findById("wnd[0]").sendVKey 0
    SAPConnection.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    SAPConnection.session.findById("wnd[1]/usr/lbl[18,3]").SetFocus
    SAPConnection.session.findById("wnd[1]/usr/lbl[18,3]").caretPosition = 7
    SAPConnection.session.findById("wnd[1]").sendVKey 2
    SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1020/ssubVIEW_AREA:SAPLCJWB:3801/ctxtMLST-LST_FERTG").Text = Round(POC, 0)
    SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1020/ssubVIEW_AREA:SAPLCJWB:3801/ctxtMLST-LST_ACTDT").Text = ActDate
    SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1020/ssubVIEW_AREA:SAPLCJWB:3801/ctxtMLST-LST_ACTDT").SetFocus
    SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1020/ssubVIEW_AREA:SAPLCJWB:3801/ctxtMLST-LST_ACTDT").caretPosition = 9
    SAPConnection.session.findById("wnd[0]").sendVKey 0
    SAPConnection.session.findById("wnd[0]").sendVKey 0
    If InStr(sBar.Text, "working") Then
        ActDate = (Replace(Right(sBar.Text, 11), ")", ""))
        SAPConnection.session.findById("wnd[0]").sendVKey 0
        SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1020/ssubVIEW_AREA:SAPLCJWB:3801/ctxtMLST-LST_ACTDT").Text = ActDate
        SAPConnection.session.findById("wnd[0]").sendVKey 0

    End If
    SAPConnection.session.findById("wnd[0]").sendVKey 0
    SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1020/subIDENTIFICATION:SAPLCJWB:3992/txtMLTX-KTEXT").Text = MileDesc
    SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1020/subIDENTIFICATION:SAPLCJWB:3992/txtMLTX-KTEXT").caretPosition = 3
    SAPConnection.session.findById("wnd[0]").sendVKey 0
    If Supplink <> "" Then
    SAPConnection.session.findById("wnd[0]/titl/shellcont/shell").PressContextButton "%GOS_TOOLBOX"
    SAPConnection.session.findById("wnd[0]/titl/shellcont/shell").SelectContextMenuItem "%GOS_URL_CREA"
    SAPConnection.session.findById("wnd[1]/usr/txtDOCUMENT_TITLE").Text = POC & "% POC Evidence " & Format(Now(), "YYMM") & "A"
    SAPConnection.session.findById("wnd[1]/usr/txtURL").Text = Supplink
    SAPConnection.session.findById("wnd[1]/usr/txtURL").SetFocus
    SAPConnection.session.findById("wnd[1]/usr/txtURL").caretPosition = 21
    SAPConnection.session.findById("wnd[1]/tbar[0]/btn[0]").press
    Else
    End If
    SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").SelectedNode = "000002"
    SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpEVAN").Select
    
    If SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpEVAN/ssubSUBSCR1:SAPLCJWB:1491/subSUBSCR_EVA:SAPLCNEV_06_DIALOG:0200/txtRCNEV-EVGEW").Text = "" Then
        SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpEVAN/ssubSUBSCR1:SAPLCJWB:1491/subSUBSCR_EVA:SAPLCNEV_06_DIALOG:0200/txtRCNEV-EVGEW").Text = "100"
    Else
    End If
    If SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpEVAN/ssubSUBSCR1:SAPLCJWB:1491/subSUBSCR_EVA:SAPLCNEV_06_DIALOG:0200/tblSAPLCNEV_06_DIALOGTCTRL_EVOPD/ctxtEVOPD-VERSN_EV[0,0]").Text = "" Then
        SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpEVAN/ssubSUBSCR1:SAPLCJWB:1491/subSUBSCR_EVA:SAPLCNEV_06_DIALOG:0200/tblSAPLCNEV_06_DIALOGTCTRL_EVOPD/ctxtEVOPD-VERSN_EV[0,0]").Text = "200"
        SAPConnection.session.findById("wnd[0]").sendVKey 0
    Else
    End If

    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[11]").press
    SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").topNode = "         23"
    
    'slut
    
    ActiveCell.Offset(0, 6).value = sBar.Text
    ActiveCell.value = 1
    Exit Sub
    
ErrHandler:
    If Err.Number <> 0 Then
        GoTo ErrGoNext
    End If
    Exit Sub
ErrGoNext:
    SAPConnection.ErrorCounter = SAPConnection.ErrorCounter + 1
    'SAPConnection.reportError "POC_Milestone_Creation_Update", ActiveCell.Offset(0, 1), Err.Number, Err.Description, Err.Source, sBar.Text
    CMail.BuildErrorList ActiveCell.Offset(0, 1), "POC_Milestone_Creation_Update", Err.Number, Err.Description, Err.Source, sBar.Text
    Err.Clear
    SAPConnection.errorContinueNextItem (trx)
End Sub
