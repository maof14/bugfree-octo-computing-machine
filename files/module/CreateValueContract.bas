Attribute VB_Name = "CreateValueContract"
Option Explicit

Sub Create_Value_ContractScript(ByRef ContTyp As String, ByRef Sarea As String, ByRef DistCha As String, ByRef Divi As String, ByRef SOff As String, ByRef SGroup As String, ByRef Soldto As String, ByRef ShipTo As String, ByRef ValidFrom As String, ByRef Validto As String, ByRef CustPo As String, ByRef VCDescr As String, ByRef Curr As String, ByRef Material As String, ByRef MatDescri As String, ByRef Amount As String, ByRef pCode As String, ByRef BilDate As String, ByRef DtD As String, ByRef BilPerc As String, ByRef Ordreason As String, ByRef CLMID As String, ByRef AssigID As String, ByRef LACemail1 As String, ByRef CEMNo As String, ByRef SponsorNo As String, ByRef CPMNo As String, ByRef PSPNo As String, ByRef trx As String, ByRef SAPConnection As Variant, ByRef CMail As Variant)
    Dim sBar
    Dim nscroll As String
    Dim startingCellRow As String
    Dim newrow As String
        
    Set sBar = SAPConnection.session.findById("wnd[0]/sbar")
        nscroll = 0
        startingCellRow = ActiveCell.Row
        On Error GoTo ErrHandler

SAPConnection.session.findById("wnd[0]/usr/ctxtVBAK-AUART").Text = ContTyp
SAPConnection.session.findById("wnd[0]/usr/ctxtVBAK-VKORG").Text = Sarea
SAPConnection.session.findById("wnd[0]/usr/ctxtVBAK-VTWEG").Text = DistCha
SAPConnection.session.findById("wnd[0]/usr/ctxtVBAK-SPART").Text = Divi
SAPConnection.session.findById("wnd[0]/usr/ctxtVBAK-VKBUR").Text = SOff
SAPConnection.session.findById("wnd[0]/usr/ctxtVBAK-VKGRP").Text = SGroup
SAPConnection.session.findById("wnd[0]/usr/ctxtVBAK-VKGRP").caretPosition = 3
SAPConnection.session.findById("wnd[0]").sendVKey 0
SAPConnection.session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").Text = VCDescr
SAPConnection.session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").Text = Soldto
SAPConnection.session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUWEV-KUNNR").Text = ShipTo
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4431/txtVBAK-KTEXT").Text = VCDescr
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4431/ctxtVBAK-GUEBG").Text = ValidFrom
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4431/ctxtVBAK-GUEEN").Text = Validto
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4431/ctxtVBAK-GUEEN").SetFocus
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4431/ctxtVBAK-GUEEN").caretPosition = 10
SAPConnection.session.findById("wnd[0]").sendVKey 0
SAPConnection.session.findById("wnd[1]/tbar[0]/btn[0]").press
SAPConnection.session.findById("wnd[1]/usr/lbl[36,6]").SetFocus
SAPConnection.session.findById("wnd[1]/usr/lbl[36,6]").caretPosition = 4
SAPConnection.session.findById("wnd[1]").sendVKey 2
SAPConnection.session.findById("wnd[0]").sendVKey 0
SAPConnection.session.findById("wnd[0]").sendVKey 0
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4431/subSUBSCREEN_TC:SAPMV45A:4909/tblSAPMV45ATCTRL_U_ERF_WERTKONTRAKT/txtVBAP-ZWERT[1,0]").Text = Amount
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4431/subSUBSCREEN_TC:SAPMV45A:4909/tblSAPMV45ATCTRL_U_ERF_WERTKONTRAKT/ctxtRV45A-MABNR[8,0]").Text = Material
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4431/subSUBSCREEN_TC:SAPMV45A:4909/tblSAPMV45ATCTRL_U_ERF_WERTKONTRAKT/txtVBAP-ARKTX[9,0]").Text = MatDescri
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4431/subSUBSCREEN_TC:SAPMV45A:4909/tblSAPMV45ATCTRL_U_ERF_WERTKONTRAKT/cmbVBAP-MVGR1[16,0]").Key = pCode
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4431/subSUBSCREEN_TC:SAPMV45A:4909/tblSAPMV45ATCTRL_U_ERF_WERTKONTRAKT/cmbVBAP-MVGR1[16,0]").SetFocus
SAPConnection.session.findById("wnd[0]").sendVKey 0
SAPConnection.session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBAK-WAERK").Text = Curr
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBAK-WAERK").SetFocus
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBAK-WAERK").caretPosition = 3
SAPConnection.session.findById("wnd[0]").sendVKey 0
SAPConnection.session.findById("wnd[1]").sendVKey 0
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/cmbVBAK-AUGRU").Key = Ordreason
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/cmbVBAK-ABRUF_PART").Key = ""
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/cmbVBAK-AUGRU").SetFocus
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04").Select
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-AFDAT[0,0]").Text = BilDate
'SAPConnection.Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-AFDAT[0,1]").Text = "28.07.2015"
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-TETXT[1,0]").Text = DtD
'SAPConnection.Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-TETXT[1,1]").Text = "0020"
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/txtFPLT-FPROZ[4,0]").Text = BilPerc
'SAPConnection.Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/txtFPLT-FPROZ[4,1]").Text = "85"
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-FAREG[9,0]").Text = "1"
'SAPConnection.Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-FAREG[9,1]").Text = "1"
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-FAREG[9,1]").SetFocus
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-FAREG[9,1]").caretPosition = 1
SAPConnection.session.findById("wnd[0]").sendVKey 0
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\08").Select
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,2]").Text = SponsorNo
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,6]").SetFocus
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,6]").caretPosition = 0
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\11").Select
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\11/ssubSUBSCREEN_BODY:SAPMV45A:4305/btnBT_KSTC").press
'Contract Signed
SAPConnection.session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,0]").Selected = False

SAPConnection.session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,2]").Selected = True
SAPConnection.session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,2]").SetFocus
SAPConnection.session.findById("wnd[0]/tbar[0]/btn[3]").press
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13").Select
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/txtVBAK-ZZEMAIL1").Text = LACemail1
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBAK-ZZ_CCLMID").Text = CLMID
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/cmbVBAK-ZZGOVSTREAM").Key = "01"
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBAK-ZZASSIGNID").Text = AssigID
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/txtVBAK-ZZEMAIL1").SetFocus
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/txtVBAK-ZZEMAIL1").caretPosition = 24
SAPConnection.session.findById("wnd[0]").sendVKey 0
SAPConnection.session.findById("wnd[0]/tbar[0]/btn[3]").press
SAPConnection.session.findById("wnd[0]/tbar[0]/btn[11]").press
SAPConnection.session.findById("wnd[1]/usr/btnBUTTON_1").press

Range(Cells(startingCellRow, 31), Cells(ActiveCell.Row, 31)) = Mid(sBar.Text, 19, 8)

        Range(Cells(startingCellRow, 2), Cells(ActiveCell.Row, 2)) = 1
    ActiveCell.value = 1
    SAPConnection.session.findById("wnd[0]/tbar[0]/okcd").Text = "/n" & trx
    SAPConnection.session.findById("wnd[0]").sendVKey 0
    Exit Sub
    
ErrHandler:
    If Err.Number <> 0 Then
        GoTo ErrGoNext
    End If
    Exit Sub
ErrGoNext:
    SAPConnection.ErrorCounter = SAPConnection.ErrorCounter + 1
    'SAPConnection.reportError "Create_Value_Contract", ActiveCell.Offset(0, 1), Err.Number, Err.Description, Err.Source, sBar.Text
    CMail.BuildErrorList ActiveCell.Offset(0, 1), "Create_Value_Contract", Err.Number, Err.Description, Err.Source, sBar.Text
    Err.Clear
    SAPConnection.errorContinueNextItem (trx)
End Sub



