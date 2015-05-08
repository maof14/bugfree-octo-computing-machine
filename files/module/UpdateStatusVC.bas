Attribute VB_Name = "UpdateStatusVC"
Option Explicit
Dim sBar

Sub UpdateStatusVCScript(ByRef VC As String, ByRef Nstatus As String, ByRef trx As String, ByRef SAPConnection As Variant, ByRef CMail As Variant)
    On Error GoTo ErrHandler
    Set sBar = SAPConnection.session.findById("wnd[0]/sbar")
    
    SAPConnection.session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = VC ' = "40039503"
    SAPConnection.session.findById("wnd[0]").sendVKey 0
    On Error Resume Next
    SAPConnection.session.findById("wnd[1]").sendVKey 0 ' Om ruta "consider subsequent documents" kommer upp, vilket inte alltid händer. Ovan kod borde kunna hantera den, och hoppa till raden efter bara.
    Application.Wait (Now + TimeValue("00:00:01"))
    SAPConnection.session.ActiveWindow.FindByName("BT_HEAD", "GuiButton").press
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\11").Select
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\11/ssubSUBSCREEN_BODY:SAPMV45A:4305/btnBT_KSTC").press
    
    Application.Wait (Now + TimeValue("00:00:01"))
    If Nstatus = "SIGN to CLOS (NONF)" Then
        SAPConnection.session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,2]").Selected = False
        SAPConnection.session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO").verticalScrollbar.Position = 1
        SAPConnection.session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,3]").Selected = True
    ElseIf Nstatus = "SIGN to COMP (FIXD)" Then
        SAPConnection.session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,2]").Selected = False
        SAPConnection.session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO").verticalScrollbar.Position = 1
        SAPConnection.session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,2]").Selected = True
    ElseIf Nstatus = "COMP to CLOS (FIXD)" Then
        SAPConnection.session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,3]").Selected = False
        SAPConnection.session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO").verticalScrollbar.Position = 1
        SAPConnection.session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,3]").Selected = True
    Else
        ActiveCell.Offset(0, 3).value = "No action chosen"
    End If

    Application.Wait (Now + TimeValue("00:00:01"))
    
    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[3]").press ' back
    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[11]").press ' spara
        
    SAPConnection.session.findById("wnd[1]/usr/btnBUTTON_1").press
    SAPConnection.session.findById("wnd[1]/usr/btnBUTTON_1").press
    
    'On Error Resume Next
    If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
        Do
                SAPConnection.session.findById("wnd[1]").sendVKey 0
        Loop While Not SAPConnection.session.findById("wnd[1]", False) Is Nothing
    End If
    
    ActiveCell.Offset(0, 3).value = sBar.Text
    ActiveCell.value = 1
    
NextVC:

Exit Sub
ErrHandler:
    If Err.Number <> 0 Then
        GoTo ErrGoNext
    End If
    Exit Sub
ErrGoNext:
    SAPConnection.ErrorCounter = SAPConnection.ErrorCounter + 1
    'SAPConnection.reportError "UpdateStatusVC", ActiveCell.Offset(0, 1), Err.Number, Err.Description, Err.Source, sBar.Text
    CMail.BuildErrorList ActiveCell.Offset(0, 1), "UpdateStatusVC", Err.Number, Err.Description, Err.Source, sBar.Text
    Err.Clear
    SAPConnection.errorContinueNextItem (trx)
End Sub
