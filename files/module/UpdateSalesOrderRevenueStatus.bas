Attribute VB_Name = "UpdateSalesOrderRevenueStatus"
Option Explicit
    Dim sBar
    Dim ScriptSetting As String
    Dim Oldlink As String

Sub UpdateSalesOrderRevenueStatusScript(ByRef SO As String, ByRef ScriptSetting As String, ByRef Supplink As String, ByRef trx As String, ByRef SAPConnection As Variant, ByRef CMail As Variant)
    Set sBar = SAPConnection.session.findById("wnd[0]/sbar")
    
    On Error GoTo ErrHandler
    
    SAPConnection.session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = SO
    SAPConnection.session.findById("wnd[0]").sendVKey 0
    
    If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
        SAPConnection.session.findById("wnd[1]").sendVKey 0
    End If
        
    If Supplink <> "" Then
        SAPConnection.session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
        SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09").Select
        SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").SelectItem "Z002", "Column1"
        SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").EnsureVisibleHorizontalItem "Z002", "Column1"
        SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").DoubleClickItem "Z002", "Column1"
        Oldlink = SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").Text
        SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").Text = Supplink + vbCr + "" + vbCr + Oldlink
        SAPConnection.session.findById("wnd[0]/tbar[0]/btn[3]").press
    Else
    End If
    
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0,0]").SetFocus
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0,0]").caretPosition = 3
    Application.Wait (Now + TimeValue("00:00:01"))
    SAPConnection.session.findById("wnd[0]").sendVKey 2
        
    If InStr(sBar.Text, "Master cost missing") Then
        ActiveCell.Offset(0, 4) = "Master cost fail."
        SAPConnection.session.findById("wnd[0]/tbar[0]/btn[11]").press
        GoTo Sejvaodra
    End If
    
    If SAPConnection.session.ActiveWindow.FindByName("T\07", "GuiTab").Text = "Account assignment" Then
        SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\07").Select
    Else
        SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06").Select
    End If

    Do
    ' Checks if Result Analysis Key is "ZBAN"
        If Not SAPConnection.session.ActiveWindow.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4457/ctxtVBAP-ABGRS", False) Is Nothing Then
            If SAPConnection.session.ActiveWindow.FindByName("VBAP-ABGRS", "GuiCTextField").Text = "ZBAN" Then
            
                SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\12").Select
                SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4456/btnBT_PSTC").press
            
                Select Case ScriptSetting
                Case "Set RREC"
                    SAPConnection.session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,1]").Selected = False
                    SAPConnection.session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,0]").Selected = True
                    SAPConnection.session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,0]").SetFocus
                Case "Remove All Status"
                    SAPConnection.session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,1]").Selected = False
                    SAPConnection.session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,0]").Selected = False
                    SAPConnection.session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,0]").SetFocus
                Case "Set RSUR"
                    SAPConnection.session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,0]").Selected = False
                    SAPConnection.session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,1]").Selected = True
                    SAPConnection.session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,1]").SetFocus
                Case Else
                    MsgBox "You have not selected an action...", , "Error"
                    Exit Sub
                End Select
                SAPConnection.session.findById("wnd[0]/tbar[0]/btn[3]").press
                ' Kanske ska vara utanför loopen men detta borde vara mer effektivt!
                ' Go (back) to SAP sheet "Account assignment"
                If SAPConnection.session.ActiveWindow.FindByName("T\07", "GuiTab").Text = "Account assignment" Then
                    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\07").Select
                Else
                    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06").Select
                End If
                'tjo
            End If
        Else
            GoTo Fortsaett
        End If
    
Fortsaett:

    'Next item
    SAPConnection.session.findById("wnd[0]/tbar[1]/btn[19]").press
    Loop While sBar.Text <> "There are no more items to be displayed"
    
Sejvaodra:
    'Save
    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[11]").press
            
    If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
        If SAPConnection.session.findById("wnd[1]").Text = "Save Incomplete Document" Then
           SAPConnection.session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
        Else
           SAPConnection.session.findById("wnd[1]").sendVKey 0
        End If
    End If
    
    If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
        Do
            If InStr(SAPConnection.session.findById("wnd[1]").Text, "Copy") > 0 Then
                SAPConnection.session.findById("wnd[1]/tbar[0]/btn[6]").press
            Else
                SAPConnection.session.findById("wnd[1]").sendVKey 0
            End If
        Loop While Not SAPConnection.session.findById("wnd[1]", False) Is Nothing
    End If
    
    If InStr(sBar.Text, "Master cost") > 0 Then
        Do
        SAPConnection.session.findById("wnd[0]").sendVKey 0
        Loop While InStr(sBar.Text, "Master cost") > 0
    End If

    If InStr(sBar.Text, "Error") > 0 Then
    ActiveCell.Offset(0, 4) = sBar.Text & ", " & Format(Now(), "yyyy/mm/dd | hh:mm")
    ActiveCell.value = 0
    SAPConnection.session.findById("wnd[0]/tbar[0]/okcd").Text = "/n" & trx
    SAPConnection.session.findById("wnd[0]").sendVKey 0
    ElseIf sBar.Text = "" Then
    GoTo ErrHandler
    ElseIf sBar.MessageType = "E" Then
    
    ActiveCell.Offset(0, 4) = "ERROR: " & sBar.Text & ", " & Format(Now(), "yyyy/mm/dd | hh:mm")
    ActiveCell.value = 0
    Else
    ActiveCell.Offset(0, 4) = sBar.Text & ", " & Format(Now(), "yyyy/mm/dd | hh:mm")
    ActiveCell.value = 1
    End If
    
    Exit Sub
ErrHandler:
    If Err.Number <> 0 Then
        GoTo ErrGoNext
    End If
    Exit Sub
ErrGoNext:
    SAPConnection.ErrorCounter = SAPConnection.ErrorCounter + 1
    CMail.BuildErrorList ActiveCell.Offset(0, 1), "UpdateSalesOrderRevenueStatus", Err.Number, Err.Description, Err.Source, sBar.Text
    Err.Clear
    SAPConnection.errorContinueNextItem (trx)
End Sub
