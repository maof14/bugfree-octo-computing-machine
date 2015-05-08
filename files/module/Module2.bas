Attribute VB_Name = "Module2"
Option Explicit
Dim sBar
Dim VCNumber, SoldToParty, Mtrl, Qty, Amnt, pCode As String
Dim startingCellRow, nscroll As Integer
Dim crtdByStp As Boolean
Dim todaySAP As String

Sub C2reateQTCMSalesOrderScript(ByRef PONumber As String, ByRef Curr As String, ByRef WBSElement As String, ByRef Descr As String, ByRef RAKey As String, ByRef trx As String, ByRef SAPConnection As Variant, ByRef CMail As Variant)
crtdByStp = False
On Error GoTo ErrHandler

'If GetDATEFORMAT = "" Then
'MsgBox "You must enter your SAP date format on the settings menu of the SmartApp"
'End
'End If

' todaySAP = Format(Now(), GetDATEFORMAT)
 

nscroll = 1
    startingCellRow = ActiveCell.Row
' Man skulle kunna ha som setting på förstasidan att "jag har VC / jag har inte VC"
    Set sBar = SAPConnection.session.findById("wnd[0]/sbar")
    ' static inputs VA01 start screen
    
    SAPConnection.session.findById("wnd[0]/usr/ctxtVBAK-AUART").Text = "ZOR3"
    SAPConnection.session.findById("wnd[0]/usr/ctxtVBAK-VKORG").Text = "1337"
    SAPConnection.session.findById("wnd[0]/usr/ctxtVBAK-VTWEG").Text = "XX"
    SAPConnection.session.findById("wnd[0]/usr/ctxtVBAK-SPART").Text = "XX"
    SAPConnection.session.findById("wnd[0]/usr/ctxtVBAK-VKBUR").Text = CreateQTCMfrm.SalesOfficetxt.value
    SAPConnection.session.findById("wnd[0]/usr/ctxtVBAK-VKGRP").Text = CreateQTCMfrm.SalesGrouptxt.value
    
    ' Om man har "I have VC for reference" så gör nedan.
    ' Om man har den andra så ska man fylla i StP och BtP. En ny variabel behövs.
    
    If ActiveCell.Offset(0, 1).value <> "" Then
        VCNumber = ActiveCell.Offset(0, 1)
    ' ********* Click "Create with reference"
    ' *********
'        SAPConnection.session.findById("wnd[0]/tbar[1]/btn[8]").press
        
        ' ********* Tab Contract
        ' *********
        
'        SAPConnection.session.findById("wnd[1]/usr/tabsMYTABSTRIP/tabpRKON").Select
        
        ' ********* Enter VC at Contract field
        ' *********
'        SAPConnection.session.findById("wnd[1]/usr/tabsMYTABSTRIP/tabpRKON/ssubSUB1:SAPLV45C:0302/ctxtLV45C-VBELN").Text = VCNumber ' Value Contract (variable)
        'SAPConnection.Session.FindById("wnd[1]/usr/tabsMYTABSTRIP/tabpRKON/ssubSUB1:SAPLV45C:0302/ctxtLV45C-VBELN").caretPosition = 8
        
        ' ********* Copy-button twice (first time it gets Sold to pt.
        ' *********
    SAPConnection.session.findById("wnd[0]").sendVKey 0
'        SAPConnection.session.findById("wnd[1]/tbar[0]/btn[5]").press
'        If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
'        If InStr(SAPConnection.session.findById("wnd[1]").Text, "Create") > 0 Then
'            SAPConnection.session.findById("wnd[1]/tbar[0]/btn[5]").press
'        End If
'        End If
'        If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
'            Do
'                ' In test this handles the following:
'                ' if currency chosen is same as default no window comes up
'                SAPConnection.session.findById("wnd[1]").sendVKey 0
'            Loop While Not SAPConnection.session.findById("wnd[1]", False) Is Nothing
'        End If
'
'        ' ********* Enter PO-number
'        ' *********
    End If
    
SAPConnection.session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").Text = VCNumber

SAPConnection.session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").Text = "UBS Revaluation" ' PO-number (variable)
            SAPConnection.session.findById("wnd[0]").sendVKey 0 ' (?)
            
            
        If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
            Do
                ' In test this handles the following:
                ' if currency chosen is same as default no window comes up
                SAPConnection.session.findById("wnd[1]").Close
            Loop While Not SAPConnection.session.findById("wnd[1]", False) Is Nothing
        End If
        
        Do
        ' sBar PO already exist..
            SAPConnection.session.findById("wnd[0]").sendVKey 0 ' (?)
        Loop Until Not sBar.MessageType = "W"
'
    ' ********* Goto header and enter SEK as currency followed by a enter on the field and on a popup
    ' *********
    
    SAPConnection.session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBAK-WAERK").Text = Curr ' Currency (Variable ?)
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBAK-WAERK").SetFocus

    SAPConnection.session.findById("wnd[0]").sendVKey 0

SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04").Select
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04").Select

SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-AFDAT[0,0]").Text = "30.05.2015"
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-TETXT[1,0]").Text = "0020"
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/txtFPLT-FPROZ[4,0]").Text = "100"
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-FAREG[9,0]").Text = "1"
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-FAREG[9,0]").SetFocus
SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-FAREG[9,0]").caretPosition = 1
SAPConnection.session.findById("wnd[0]").sendVKey 0

    
'    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04").Select
'    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04").Select
'    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-AFDAT[0,0]").Text = todaySAP
'    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-TETXT[1,0]").Text = "Z002"
'    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/txtFPLT-FPROZ[4,0]").Text = "100"
'    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-FAREG[9,0]").Text = "1"
'    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-FPTTP[12,0]").Text = "21"
'    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-FKARV[13,0]").Text = "ZF11"
'    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-FKARV[13,0]").SetFocus
'    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\04/ssubSUBSCREEN_BODY:SAPLV60F:4203/tblSAPLV60FTCTRL_FPLAN_TEILFA/ctxtFPLT-FKARV[13,0]").caretPosition = 4
'    SAPConnection.session.findById("wnd[0]").sendVKey 0

        If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
            Do
                ' In test this handles the following:
                ' if currency chosen is same as default no window comes up
                SAPConnection.session.findById("wnd[1]").sendVKey 0
            Loop While Not SAPConnection.session.findById("wnd[1]", False) Is Nothing
        End If
        '********* Set value "Correction Project" on customer group 5
    If CreateQTCMfrm.Custgrp5box.value = True Then
    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[3]").press
        SAPConnection.session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press

    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\12").Select
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR5").Key = "COP"
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR5").SetFocus
    End If
    
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08").Select
    
        nscroll = 0

        If SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,0]").Key <> "ZG" Then  ' Check if there are more items (PCODE)
            Do
                SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW").verticalScrollbar.Position = nscroll
                
                nscroll = nscroll + 1
            Loop While SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,0]").Text <> "" And SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,0]").Key <> "ZG" ' Loop while the VC is the same as the current
        End If
    
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,0]").Key = "ZG"
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,0]").Text = PONumber

' ********* Back (from the header, to the "start" screen)
    ' *********
    
    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[3]").press
    
    ' ********* Enter first material, necessary before making connection to WBS...
    ' *********
    
    Mtrl = ActiveCell.Offset(0, 7)
    Amnt = ActiveCell.Offset(0, 8)
    pCode = ActiveCell.Offset(0, 9)
    
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,0]").Text = Mtrl ' Material (variable)
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,0]").Text = "1" ' Quantity (variable(?))
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtKOMV-KBETR[15,0]").Text = Round(Amnt, 2) ' Amount (variable)
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/cmbVBAP-MVGR1[37,0]").Key = pCode ' P-code (variable)
    
    ' ********* Mark material cell and hit enter twice (first generates warning)
    ' *********
    
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,0]").SetFocus
    SAPConnection.session.findById("wnd[0]").sendVKey 0
        If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
            Do
                ' In test this handles the following:
                ' if currency chosen is same as default no window comes up
                SAPConnection.session.findById("wnd[1]").sendVKey 0
            Loop While Not SAPConnection.session.findById("wnd[1]", False) Is Nothing
        End If
    SAPConnection.session.findById("wnd[0]").sendVKey 0
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,0]").SetFocus
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,0]").caretPosition = 4
    ActiveCell.Offset(0, 12).value = SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0,0]").Text

    ' ********* Doubleclick the first material to create WBS connection
    ' *********
    
    SAPConnection.session.findById("wnd[0]").sendVKey 2
    
    ' ********* Goto Additional data B
    ' *********
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15").Select
    
    ' ********* Fill in Additional data B
    ' *********
    
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/ctxt*PROJ-PSPNR").Text = Left(WBSElement, 12) ' Project definition (variable (eller left(wbs, 12))
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/ctxt*PRHI-POSNR").Text = WBSElement ' Superior WBS (variable)
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/txt*PRPS-POST1").Text = Descr ' Description (variable)
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/ctxt*PRPS-ABGSL").Text = RAKey ' (RA-key, variable ?)
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/txt*PRPS-POST1").SetFocus
    
    ' ********* Create and make assignment-button
    ' *********
    
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/btnOK89541").press
    
    ' ********* Press no on copying WBS-prompt
    ' *********
    
    SAPConnection.session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
    
    ' ********* Go back to header
    ' *********
        If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
            Do
                ' In test this handles the following:
                ' if currency chosen is same as default no window comes up
                SAPConnection.session.findById("wnd[1]").sendVKey 0
            Loop While Not SAPConnection.session.findById("wnd[1]", False) Is Nothing
        End If
    Range(Cells(startingCellRow, 13), Cells(ActiveCell.Row, 13)).value = Mid(sBar.Text, 13, 18)
    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[3]").press
    
    ' ********* If more items per WBS is needed, Loop this.
    ' *********
    If ActiveCell.Offset(0, 5) = ActiveCell.Offset(1, 5) Then ' Check for a different WBS
        Do
            ActiveCell.Offset(1, 0).Select
            Mtrl = ActiveCell.Offset(0, 7)
            Amnt = ActiveCell.Offset(0, 8)
            pCode = ActiveCell.Offset(0, 9)
            
            SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG").verticalScrollbar.Position = nscroll
            SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,0]").Text = Mtrl
            SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,0]").Text = "1"
            SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtKOMV-KBETR[15,0]").Text = Round(Amnt, 2)
            SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/cmbVBAP-MVGR1[37,0]").Key = pCode
            SAPConnection.session.findById("wnd[0]").sendVKey 0
        If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
            Do
                ' In test this handles the following:
                ' if currency chosen is same as default no window comes up
                SAPConnection.session.findById("wnd[1]").sendVKey 0
            Loop While Not SAPConnection.session.findById("wnd[1]", False) Is Nothing
        End If
            SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,0]").SetFocus
            ActiveCell.Offset(0, 11).value = SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-PS_PSP_PNR[51,0]").Text
            ActiveCell.Offset(0, 12).value = SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0,0]").Text
            nscroll = nscroll + 1
        Loop While ActiveCell.Offset(0, 5) = ActiveCell.Offset(1, 5) ' Loop while the WBS Description is the same as the current
    End If
    
    ' ********* Final enter click
    ' *********
    
        SAPConnection.session.findById("wnd[0]").sendVKey 0
                    

    ' ********* Save Sales order
    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[11]").press
    
    If crtdByStp = True Then ' If created with stp then check for document incomplete log
        If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
            If InStr(SAPConnection.session.findById("wnd[1]", False).Text, "Save Incomplete Document") > 0 Then SAPConnection.session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
        End If
    End If
    
'    Do
'        SAPConnection.session.findById("wnd[0]").sendVKey 0 ' (?)
        If InStr(sBar.Text, "Partner") > 0 Then
        Range(Cells(startingCellRow, 2), Cells(ActiveCell.Row, 2)).value = 0
        Range(Cells(startingCellRow, 12), Cells(ActiveCell.Row, 12)).value = sBar.Text
        SAPConnection.session.findById("wnd[0]/tbar[0]/okcd").Text = "/n" & trx
        SAPConnection.session.findById("wnd[0]").sendVKey 0
        Exit Sub
        End If
'        'If (crtdByStp = True) And (InStr(SAPConnection.Session.FindById("wnd[1]", False).Text, "Save Incomplete Document") > 0) Then SAPConnection.Session.FindById("wnd[1]/usr/btnSPOP-VAROPTION1").Press
'    Loop Until sBar.MessageType = "S"
    ' Stop ' for later debugging
    
    ' ********* Here is PROBABLY where SAP sBar outputs the new sales order number. This needs to be captured !!!
    ' *********
    
    Range(Cells(startingCellRow, 12), Cells(ActiveCell.Row, 12)).value = Mid(sBar.Text, 13, 8)
    Range(Cells(startingCellRow, 2), Cells(ActiveCell.Row, 2)).value = 1
    'ActiveCell.Offset(0, 11).Value = sBar.Text
    
    ' ********* Create settlement rule prompt upon save
    
    SAPConnection.session.findById("wnd[0]/tbar[1]/btn[6]").press
    
    ' ********* Check some checkbox (select WBS / SO to create settlement rule for?
    ' *********
    If CreateQTCMfrm.Settlementbox.value = True And CreateQTCMfrm.MonthBox.value = "Current" Then
    SAPConnection.session.findById("wnd[0]/usr/chk[0,0]").Selected = True
    SAPConnection.session.findById("wnd[0]/tbar[1]/btn[6]").press ' (?)
    Range(Cells(startingCellRow, 15), Cells(ActiveCell.Row, 15)).value = sBar.Text
    Else
    
    Range(Cells(startingCellRow, 15), Cells(ActiveCell.Row, 15)).value = "WBS created without settlement rule"

    End If
    
    If IsEmpty(ActiveCell.Offset(1, 1)) Then
        If CreateQTCMfrm.MonthBox.value <> "Current" Then
        Call createsettlementrule(SAPConnection)
        End If
        If CreateQTCMfrm.Releasebox.value = True Then
        Call ReleaseWBS(SAPConnection)
        Call removerfr(SAPConnection)
        End If
    End If
    
    
    
    SAPConnection.session.findById("wnd[0]/tbar[0]/okcd").Text = "/n" & trx
    SAPConnection.session.findById("wnd[0]").sendVKey 0
    ' ********* When looping the script, it needs to finish with going back to transaction /nva01
    ' *********

    Exit Sub
ErrHandler:
    If Err.Number <> 0 Then
        GoTo ErrGoNext
    End If
    Exit Sub
ErrGoNext:
    Stop
    Resume
    'SAPConnection.reportError "CreateQTCMSalesOrder", ActiveCell.Offset(0, 1), Err.Number, Err.Description, Err.Source, sBar.Text
    CMail.BuildErrorList ActiveCell.Offset(0, 1), "CreateQTCMSalesOrder", Err.Number, Err.Description, Err.Source, sBar.Text
    SAPConnection.errorContinueNextItem (trx)
    SAPConnection.ErrorCounter = SAPConnection.ErrorCounter + 1
    Err.Clear
End Sub
Sub ReleaseWBS(ByRef SAPConnection As Variant)
    
    Dim WBS As String
    Dim oName
    Dim oName2
    
    SAPConnection.session.findById("wnd[0]/tbar[0]/okcd").Text = "/ncj20n"
    SAPConnection.session.findById("wnd[0]").sendVKey 0
    Range("P8").Select
    ActiveCell.FormulaR1C1 = "Release Status (SAP)"
    Range("Q8").Select
    ActiveCell.FormulaR1C1 = "RFR Status (SAP)"
    Cells(9, 2).Select
    
    Do Until IsEmpty(ActiveCell.Offset(0, 1))
        WBS = ActiveCell.Offset(0, 11)
        SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").topNode = "         23"
        SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell").pressButton "OPEN"
        SAPConnection.session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PROJ_EXT").Text = ""
        SAPConnection.session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PRPS_EXT").Text = WBS
        SAPConnection.session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-AUFNR").Text = ""
        SAPConnection.session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PRPS_EXT").SetFocus
        SAPConnection.session.findById("wnd[1]").sendVKey 0
        
        If sBar.Text = "Not all objects were locked (see lock log)" Then
            ActiveCell.Offset(0, 14) = "WBS is being processed by another user, try later."
            SAPConnection.session.findById("wnd[0]/tbar[0]/btn[3]").press
        End If
        
        ' Check if change mode is active
        If Not SAPConnection.session.ActiveWindow.FindByName("PRPS-POST1", GuiTextField).Changeable = True Then
            SAPConnection.session.findById("wnd[0]/tbar[1]/btn[13]").press
        End If

        Set oName = SAPConnection.session.ActiveWindow.FindByName("CNJ_STAT-STTXT_INT", "GuiTextField")
        Set oName2 = SAPConnection.session.ActiveWindow.FindByName("CNJ_STAT-STTXT_EXT", "GuiTextField")
                
                If InStr(oName2.Text, "ZLIQ") > 0 Then
                ActiveCell.Offset(0, 14).value = "ERROR: No settlement is created on WBS, cannot be released"
                Exit Sub
                Else
                End If
            If Not InStr(oName.Text, "REL") > 0 Then
                SAPConnection.session.findById("wnd[0]/mbar/menu[1]/menu[2]/menu[0]").Select
                If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
                    Do
                        SAPConnection.session.findById("wnd[1]").Close
                    Loop While Not SAPConnection.session.findById("wnd[1]") Is Nothing
                    ActiveCell.Offset(0, 14).value = "Information window upon status change attempt, skipping this WBS."
                    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[3]").press ' Back
                Else
                    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[11]").press
                    ActiveCell.Offset(0, 14).value = sBar.Text
                End If
            Else
                SAPConnection.session.findById("wnd[0]/tbar[0]/btn[3]").press ' back
                ActiveCell.Offset(0, 14).value = "Status already set"
            End If
    
       
        Do
        
        ActiveCell.Offset(1, 0).Select
    If IsEmpty(ActiveCell.Offset(0, 1)) Then
    Exit Sub
    End If
        ActiveCell.Offset(0, 14).value = "=R[-1]C"
        Loop While ActiveCell.Offset(0, 11) = ActiveCell.Offset(-1, 11) And ActiveCell.Offset(0, 11) <> ""
        
        Loop
        

End Sub

Sub removerfr(ByRef SAPConnection As Variant)
On Error GoTo Errorhandler
Dim nscroll
Dim SO

    SAPConnection.session.findById("wnd[0]/tbar[0]/okcd").Text = "/nva02"
    SAPConnection.session.findById("wnd[0]").sendVKey 0
    
    Cells(9, 2).Select
    
    Do Until IsEmpty(ActiveCell.Offset(0, 1))
    nscroll = 1

        SO = ActiveCell.Offset(0, 10)
    SAPConnection.session.findById("wnd[0]/tbar[0]/okcd").Text = "/nva02"
    SAPConnection.session.findById("wnd[0]").sendVKey 0
    SAPConnection.session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = SO
    SAPConnection.session.findById("wnd[0]").sendVKey 0
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\07").Select
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4409/subSUBSCREEN_TC:SAPMV45A:4922/tblSAPMV45ATCTRL_UPOS_ABSAGE/cmbVBAP-ABGRU[2,0]").Key = ""
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4409/subSUBSCREEN_TC:SAPMV45A:4922/tblSAPMV45ATCTRL_UPOS_ABSAGE/cmbVBAP-ABGRU[2,0]").SetFocus

If SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4409/subSUBSCREEN_TC:SAPMV45A:4922/tblSAPMV45ATCTRL_UPOS_ABSAGE/cmbVBAP-ABGRU[2,1]").Key = "YQ" Then  ' Check if there are more items (PCODE)
        Do
            SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4409/subSUBSCREEN_TC:SAPMV45A:4922/tblSAPMV45ATCTRL_UPOS_ABSAGE").verticalScrollbar.Position = nscroll
            If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
            Do
                ' In test this handles the following:
                ' if currency chosen is same as default no window comes up
                SAPConnection.session.findById("wnd[1]").sendVKey 0
            Loop While Not SAPConnection.session.findById("wnd[1]", False) Is Nothing
            End If
            SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4409/subSUBSCREEN_TC:SAPMV45A:4922/tblSAPMV45ATCTRL_UPOS_ABSAGE/cmbVBAP-ABGRU[2,0]").Key = ""
            nscroll = nscroll + 1
        Loop While SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4409/subSUBSCREEN_TC:SAPMV45A:4922/tblSAPMV45ATCTRL_UPOS_ABSAGE/cmbVBAP-ABGRU[2,1]").Key = "YQ" ' Loop while the VC is the same as the current
End If

    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[11]").press
    ActiveCell.Offset(0, 15).value = sBar.Text
    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[3]").press
    
        Do
        
        ActiveCell.Offset(1, 0).Select
    If IsEmpty(ActiveCell.Offset(0, 1)) Then
    Exit Sub
    End If
        ActiveCell.Offset(0, 15).value = "=R[-1]C"
        Loop While ActiveCell.Offset(0, 10) = ActiveCell.Offset(-1, 10)
        
        Loop
        Exit Sub
Errorhandler:
Stop
Resume

End Sub
Sub createsettlementrule(ByRef SAPConnection As Variant)
    Dim WBS As String
    Dim YearN
    Dim MonthN
    
    Range("R8").Select
    ActiveCell.FormulaR1C1 = "Settlement Rule Status (SAP)"
    
    If CreateQTCMfrm.YearBox.value = "Current" Then
    YearN = Format(Now(), "yyyy")
    Else
    YearN = CreateQTCMfrm.YearBox.value
    End If
    MonthN = CreateQTCMfrm.MonthBox.value
    
    SAPConnection.session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzsetc"
    SAPConnection.session.findById("wnd[0]").sendVKey 0
    
    Cells(9, 2).Select
    
    Do Until IsEmpty(ActiveCell.Offset(0, 1))
    
        WBS = ActiveCell.Offset(0, 11)
        
    SAPConnection.session.findById("wnd[0]/usr/chkPA_MESS").Selected = True
    SAPConnection.session.findById("wnd[0]/usr/ctxtPA_PROJ").Text = WBS
    SAPConnection.session.findById("wnd[0]/usr/txtP_GJAHR").Text = YearN
    SAPConnection.session.findById("wnd[0]/usr/txtP_PERIO").Text = MonthN
    SAPConnection.session.findById("wnd[0]").sendVKey 8
    ActiveCell.Offset(0, 16).value = sBar.Text
        Do
        
        ActiveCell.Offset(1, 0).Select
    If IsEmpty(ActiveCell.Offset(0, 1)) Then
    Exit Sub
    End If
        ActiveCell.Offset(0, 16).value = "=R[-1]C"
        Loop While ActiveCell.Offset(0, 11) = ActiveCell.Offset(-1, 11)
    Loop

End Sub

