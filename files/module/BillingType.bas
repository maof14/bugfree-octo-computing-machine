Attribute VB_Name = "BillingType"
Option Explicit
    Dim sBar
    Dim ScriptSetting As String

Sub UpdateBillingTypeScript(ByRef SO As String, ByRef BilDate As String, ByRef trx As String, ByRef SAPConnection As Variant, ByRef CMail As Variant)
    Set sBar = SAPConnection.session.findById("wnd[0]/sbar")
    
    On Error GoTo ErrHandler
    
    SAPConnection.session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = SO
    SAPConnection.session.findById("wnd[0]").sendVKey 0
    
    If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
        SAPConnection.session.findById("wnd[1]").sendVKey 0
    End If
    
        
    
SAPConnection.session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04").Select
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA").GetAbsoluteRow(0).Selected = True
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-AFDAT[0,0]").SetFocus
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-AFDAT[0,0]").caretPosition = 0
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/btnBT_KOLO").press
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-AFDAT[0,0]").Text = BilDate
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-TETXT[1,0]").Text = "z002"
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/txtFPLT-FPROZ[4,0]").Text = "100"
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-FAREG[9,0]").Text = "1"
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-FPTTP[12,0]").Text = "21"
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-FKARV[13,0]").Text = "zf11"
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-FKARV[13,0]").SetFocus
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-FKARV[13,0]").caretPosition = 4
SAPConnection.session.findById("wnd[0]").sendVKey 0
SAPConnection.session.findById("wnd[0]/tbar[0]/btn[3]").press
SAPConnection.session.findById("wnd[0]/tbar[0]/btn[11]").press
            
    If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
        If SAPConnection.session.findById("wnd[1]").Text = "Save Incomplete Document" Then
           SAPConnection.session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
        Else
           SAPConnection.session.findById("wnd[1]").sendVKey 0
        End If
    End If
    
    If InStr(sBar.Text, "Master cost") > 0 Then
        Do
        SAPConnection.session.findById("wnd[0]").sendVKey 0
        Loop While InStr(sBar.Text, "Master cost") > 0
    End If
    ActiveCell.Offset(0, 3) = sBar.Text & ", " & Format(Now(), "yyyy/mm/dd | hh:mm")
SAPConnection.session.findById("wnd[0]/tbar[0]/btn[3]").press

    
    ActiveCell.value = 1
    
    Exit Sub
ErrHandler:
    If Err.Number <> 0 Then
        GoTo ErrGoNext
    End If
    Exit Sub
ErrGoNext:
    SAPConnection.ErrorCounter = SAPConnection.ErrorCounter + 1
    'SAPConnection.reportError "UpdateBillingType", ActiveCell.Offset(0, 1), Err.Number, Err.Description, Err.Source, sBar.Text
    CMail.BuildErrorList ActiveCell.Offset(0, 1), "UpdateBillingType", Err.Number, Err.Description, Err.Source, sBar.Text
    Err.Clear
    SAPConnection.errorContinueNextItem (trx)
End Sub

