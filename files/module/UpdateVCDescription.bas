Attribute VB_Name = "UpdateVCDescription"
Sub UpdateVCDescriptionScript(ByRef VC As String, ByRef Desc As String, ByRef trx As String, ByRef SAPConnection As Variant, ByRef CMail As Variant)
    Dim sBar
    sBar = SAPConnection.session.findById("wnd[0]/sbar")
    
    SAPConnection.session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = VC
    SAPConnection.session.findById("wnd[0]").sendVKey 0
    ActiveCell.Offset(0, 3).value = SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4431/txtVBAK-KTEXT").Text
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4431/txtVBAK-KTEXT").Text = Desc
    SAPConnection.session.findById("wnd[0]").sendVKey 0
    SAPConnection.session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
    newtxt = ActiveCell.Offset(0, 3).value
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09").Select
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").SelectItem "0001", "Column1"
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").EnsureVisibleHorizontalItem "0001", "Column1"
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").Text = newtxt + vbCr + ""
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").SetSelectionIndexes 26, 26
    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[11]").press
        If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
            Do
                SAPConnection.session.findById("wnd[1]/usr/btnBUTTON_1").press
            Loop While Not SAPConnection.session.findById("wnd[1]", False) Is Nothing
        End If
    ActiveCell.Offset(0, 4) = sBar.Text
    ActiveCell.value = 1
Exit Sub
    
ErrHandler:
    If Err.Number <> 0 Then
        GoTo ErrGoNext
    End If
    Exit Sub
ErrGoNext:
    Stop
    Resume
    CMail.BuildErrorList ActiveCell.Offset(0, 1), "UpdateVCDescription", Err.Number, Err.Description, Err.Source, sBar.Text
    SAPConnection.errorContinueNextItem (trx)
    SAPConnection.ErrorCounter = SAPConnection.ErrorCounter + 1
    Err.Clear
    
End Sub
