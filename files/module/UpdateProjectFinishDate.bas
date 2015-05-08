Attribute VB_Name = "UpdateProjectFinishDate"
Option Explicit
Dim sBar
Sub UpdateProjectFinishDateScript(ByRef Prjct As String, ByRef FinDat As String, ByRef trx As String, ByRef SAPConnection As Variant, ByRef CMail As Variant)
    On Error GoTo ErrHandler
    Set sBar = SAPConnection.session.findById("wnd[0]/sbar")
    
    SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").topNode = "         23"
    SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell").pressButton "OPEN"
        SAPConnection.session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PROJ_EXT").Text = Prjct
        SAPConnection.session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PRPS_EXT").Text = ""
        SAPConnection.session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-AUFNR").Text = ""
    SAPConnection.session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PROJ_EXT").caretPosition = 12
    SAPConnection.session.findById("wnd[1]").sendVKey 0
    Application.Wait (Now + TimeValue("00:00:01"))
    If Not SAPConnection.session.ActiveWindow.FindByName("PROJ-POST1", "GuiTextField").Changeable = True Then
        SAPConnection.session.ActiveWindow.findById("wnd[0]/tbar[1]/btn[13]").press
    End If
    SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3998/tabsPTABSCR/tabpPGND/ssubSUBSCR2:SAPLCJWB:1205/ctxtPROJ-PLSEZ").Text = FinDat
    SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3998/tabsPTABSCR/tabpPGND/ssubSUBSCR2:SAPLCJWB:1205/ctxtPROJ-PLSEZ").SetFocus
    SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3998/tabsPTABSCR/tabpPGND/ssubSUBSCR2:SAPLCJWB:1205/ctxtPROJ-PLSEZ").caretPosition = 10
    SAPConnection.session.findById("wnd[0]").sendVKey 0
    SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3998/tabsPTABSCR/tabpPGND/ssubSUBSCR2:SAPLCJWB:1205/ctxtPROJ-PLSEZ").SetFocus
    SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3998/tabsPTABSCR/tabpPGND/ssubSUBSCR2:SAPLCJWB:1205/ctxtPROJ-PLSEZ").caretPosition = 6
    SAPConnection.session.findById("wnd[0]").sendVKey 0
    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[11]").press
    
    If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
        Do
            SAPConnection.session.findByUd("wnd[1]").sendVKey 0
        Loop While Not SAPConnection.session.findById("wnd[1]", False) Is Nothing
    End If
    
    SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").topNode = "         23"
    
    ActiveCell.Offset(0, 3).value = sBar.Text
    ActiveCell.value = 1
Exit Sub

ErrHandler:
    If Err.Number <> 0 Then
        GoTo ErrGoNext
    End If
    Exit Sub
ErrGoNext:
    SAPConnection.ErrorCounter = SAPConnection.ErrorCounter + 1
    CMail.BuildErrorList ActiveCell.Offset(0, 1), "UpdateProjectFinishDate", Err.Number, Err.Description, Err.Source, sBar.Text
    Err.Clear
    SAPConnection.errorContinueNextItem (trx)
End Sub
