Attribute VB_Name = "ENOPlannedCostUpdateMarzenna"
Option Explicit
Dim sBar
Dim nscroll, startingCellRow, startingAct As Integer
Dim Description, Amount, Curr, CostEl As String
Sub ENOPlannedCostUpdateMarzennaScript(ByRef WBS As String, ByRef trx As String, ByRef SAPConnection As Variant, ByRef CMail As Variant) ' amount, curr, etc
    Dim sBar
    Set sBar = SAPConnection.session.findById("wnd[0]/sbar")

    On Error GoTo ErrHandler

    startingCellRow = ActiveCell.Row
    startingAct = 9900
    nscroll = 1
    SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").topNode = "         23"
    SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell").pressButton "OPEN"
    SAPConnection.session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PRPS_EXT").Text = WBS
    SAPConnection.session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PRPS_EXT").SetFocus
    SAPConnection.session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PRPS_EXT").caretPosition = 18
    SAPConnection.session.findById("wnd[1]").sendVKey 0
    
    If Not SAPConnection.session.ActiveWindow.FindByName("PRPS-POST1", "GuiTextField").Changeable = True Then
        SAPConnection.session.findById("wnd[0]/tbar[1]/btn[13]").press
    End If
    
    SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
    
    ' Kolla om det inte finns något nätverk!!!
    If SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,0]").Text = "" Then
        SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").ExpandNode "          1"
        SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").SelectedNode = "          4"
        SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").topNode = "          1"
        SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").DoubleClickNode "          4"
        If SAPConnection.session.ActiveWindow.FindByName("CAUFVD-PROFID", "GuiComboBox").Changeable = True Then
            SAPConnection.session.ActiveWindow.FindByName("CAUFVD-PROFID", "GuiComboBox").Key = "ZEAB001"
        End If
        SAPConnection.session.ActiveWindow.FindByName("CAUFVD-DISPO", "GuiCTextField").Text = "001"
        SAPConnection.session.findById("wnd[0]").sendVKey 0
        SAPConnection.session.findById("wnd[0]").sendVKey 0
        SAPConnection.session.findById("wnd[0]").sendVKey 0
        SAPConnection.session.findById("wnd[0]").sendVKey 0
        SAPConnection.session.findById("wnd[0]").sendVKey 0
        SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").ExpandNode "          5"
        SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").SelectedNode = "         11"
        SAPConnection.session.findById("wnd[0]/mbar/menu[1]/menu[1]/menu[0]").Select
    End If
    
    SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,0]").SetFocus
    SAPConnection.session.findById("wnd[0]").sendVKey 2
    SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "ACTY_OVW"
    SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpPSPV").Select
        ' If 9900 is not the first activity
        If Not SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpPSPV/ssubSUBSCR_2000:SAPLCOVG:2095/tblSAPLCOVGTCTRL_2095/txtAFVGD-VORNR[0,0]").Text = startingAct Then
            Do Until SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpPSPV/ssubSUBSCR_2000:SAPLCOVG:2095/tblSAPLCOVGTCTRL_2095/txtAFVGD-VORNR[0,0]").Text = startingAct
            ' Om första raden är aktivitet 9900 så skiter det sig ser jag nu.
                If SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpPSPV/ssubSUBSCR_2000:SAPLCOVG:2095/tblSAPLCOVGTCTRL_2095/txtAFVGD-LTXA1[5,1]").Text = "" Then Exit Do
                ' scrolla ned
                SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpPSPV/ssubSUBSCR_2000:SAPLCOVG:2095/tblSAPLCOVGTCTRL_2095").verticalScrollbar.Position = nscroll
                ' Öka scrollens position värde med en rad
                If SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpPSPV/ssubSUBSCR_2000:SAPLCOVG:2095/tblSAPLCOVGTCTRL_2095/txtAFVGD-VORNR[0,1]").Text = startingAct Then
                    startingAct = startingAct + 1
                End If
                nscroll = nscroll + 1
            Loop
        ' If 9900 is the first activity
        Else
            Do Until SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpPSPV/ssubSUBSCR_2000:SAPLCOVG:2095/tblSAPLCOVGTCTRL_2095/txtAFVGD-LTXA1[5,1]").Text = ""
                If SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpPSPV/ssubSUBSCR_2000:SAPLCOVG:2095/tblSAPLCOVGTCTRL_2095/txtAFVGD-VORNR[0,0]").Text = startingAct Then
                    startingAct = startingAct + 1
                End If
                SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOVG:2001/tabsTABSTRIP_2000/tabpPSPV/ssubSUBSCR_2000:SAPLCOVG:2095/tblSAPLCOVGTCTRL_2095").verticalScrollbar.Position = nscroll
                nscroll = nscroll + 1
            Loop
            startingAct = startingAct + 1
        End If
    Do
        Description = ActiveCell.Offset(0, 2)
        Amount = ActiveCell.Offset(0, 3)
        Curr = ActiveCell.Offset(0, 4)
        CostEl = ActiveCell.Offset(0, 5)
        SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").ExpandNode "          1"
        SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").SelectedNode = "          5"
        SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").topNode = "          1"
        SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").DoubleClickNode "         11"
        SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subIDENTIFICATION:SAPLCONW:0110/txtAFVGD-VORNR").Text = startingAct
        SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subIDENTIFICATION:SAPLCONW:0110/txtAFVGM-LTXA1").Text = Description
        SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpKOSD/ssubSUBSCR_1000:SAPLCONW:1550/txtAFVGD-PRKST").Text = Round(Amount, 2)
        SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpKOSD/ssubSUBSCR_1000:SAPLCONW:1550/ctxtAFVGD-SAKTO").Text = CostEl
        SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpKOSD/ssubSUBSCR_1000:SAPLCONW:1550/ctxtAFVGD-WAERS").Text = Curr
        SAPConnection.session.findById("wnd[0]").sendVKey 0
        ActiveCell.Offset(1, 0).Select
        startingAct = startingAct + 1
    Loop Until ActiveCell.Offset(0, 2) = ""
    'SAPConnection.Session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = "000003"
    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[11]").press ' save

    ' If there is a popup - loop through popup checks for Availability control, Scheduling, budget and any else pupop until there are no more popups.
    If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
        Do
            If InStr(SAPConnection.session.findById("wnd[1]").Text, "Availability Control") > 0 Then
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
    
'Check if the status bar is blank just to make sure it didnt save, if it didnt then it will add another "ENTER"
    If sBar.Text = "" Then
        SAPConnection.session.findById("wnd[0]").sendVKey 0
    End If
        
    SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").topNode = "         23"
    
    Cells(startingCellRow, 8) = sBar.Text
    Cells(startingCellRow, 2).value = 1
    Exit Sub
    
ErrHandler:
    If Err.Number <> 0 Then
        GoTo ErrGoNext
    End If
    Exit Sub
ErrGoNext:
    Stop
    Resume
    'SAPConnection.reportError "ENOPlannedCostUpdateMarzenna", ActiveCell.Offset(0, 1), Err.Number, Err.Description, Err.Source, sBar.Text
    CMail.BuildErrorList ActiveCell.Offset(0, 1), "ENOPlannedCostUpdateMarzenna", Err.Number, Err.Description, Err.Source, sBar.Text
    SAPConnection.errorContinueNextItem (trx)
    SAPConnection.ErrorCounter = SAPConnection.ErrorCounter + 1
    Err.Clear
End Sub
