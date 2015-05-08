Attribute VB_Name = "RebateUpdate"
Sub RebateUpdateScript(ByRef Rebate As String, ByRef Percen As String, ByRef trx As String, ByRef SAPConnection As Variant, ByRef CMail As Variant)
    Dim sBar
    Dim nscroll
    Set sBar = SAPConnection.session.findById("wnd[0]/sbar")
    
    On Error GoTo ErrHandler
    
    nscroll = 1
SAPConnection.session.findById("wnd[0]").resizeWorkingPane 99, 11, False
startingCellRow = ActiveCell.Row
SAPConnection.session.findById("wnd[0]/usr/ctxtRV13A-KNUMA_BO").Text = Rebate
SAPConnection.session.findById("wnd[0]").sendVKey 0
    If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
        Do
            SAPConnection.session.findById("wnd[1]").sendVKey 0
        Loop While Not SAPConnection.session.findById("wnd[1]", False) Is Nothing
    End If
SAPConnection.session.findById("wnd[0]/tbar[1]/btn[9]").press
SAPConnection.session.findById("wnd[1]/usr/cntlCUSTOM_CONTAINER/shellcont/shell").selectedRows = "0"
SAPConnection.session.findById("wnd[1]/usr/cntlCUSTOM_CONTAINER/shellcont/shell").doubleClickCurrentCell
SAPConnection.session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/txtKONP-KBETR[2,0]").Text = Round(Percen, 3)
SAPConnection.session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/txtKONP-KBRUE[6,0]").Text = Round(Percen, 3)
SAPConnection.session.findById("wnd[0]").resizeWorkingPane 99, 11, False
If SAPConnection.session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/txtKOMG-MVGR1[0,1]").Text <> "" Then  ' Check if there are more items (PCODE)
        Do
            SAPConnection.session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY").verticalScrollbar.Position = nscroll
            SAPConnection.session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/txtKONP-KBETR[2,0]").Text = Round(Percen, 3)
            SAPConnection.session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/txtKONP-KBRUE[6,0]").Text = Round(Percen, 3)
            SAPConnection.session.findById("wnd[0]").sendVKey 0
            nscroll = nscroll + 1
        Loop While SAPConnection.session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/txtKOMG-MVGR1[0,1]").Text <> "" ' Loop while the VC is the same as the current
    End If
SAPConnection.session.findById("wnd[0]").sendVKey 0
SAPConnection.session.findById("wnd[0]/tbar[0]/btn[11]").press
        Range(Cells(startingCellRow, 5), Cells(ActiveCell.Row, 5)) = sBar.Text
        Range(Cells(startingCellRow, 2), Cells(ActiveCell.Row, 2)) = 1
    Exit Sub
    
ErrHandler:
    If Err.Number <> 0 Then
        GoTo ErrGoNext
    End If
    Exit Sub
ErrGoNext:
    SAPConnection.ErrorCounter = SAPConnection.ErrorCounter + 1
    CMail.BuildErrorList ActiveCell.Offset(0, 1), "Rebate_Update", Err.Number, Err.Description, Err.Source, sBar.Text
    Err.Clear
    SAPConnection.errorContinueNextItem (trx)
End Sub

