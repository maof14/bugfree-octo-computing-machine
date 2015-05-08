Attribute VB_Name = "RebateDescriptionUpdate"
Sub Rebate_Description_UpdateScript(ByRef Rebate As String, ByRef Descrip As String, ByRef trx As String, ByRef SAPConnection As Variant, ByRef CMail As Variant)
    Dim sBar
    Dim startingCellRow
    Set sBar = SAPConnection.session.findById("wnd[0]/sbar")
    
    On Error GoTo ErrHandler
    
startingCellRow = ActiveCell.Row
SAPConnection.session.findById("wnd[0]/usr/ctxtRV13A-KNUMA_BO").Text = Rebate
SAPConnection.session.findById("wnd[0]/usr/ctxtRV13A-KNUMA_BO").caretPosition = 3
SAPConnection.session.findById("wnd[0]").sendVKey 0
    If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
        Do
            SAPConnection.session.findById("wnd[1]").sendVKey 0
        Loop While Not SAPConnection.session.findById("wnd[1]", False) Is Nothing
    End If
SAPConnection.session.findById("wnd[0]/usr/txtKONA-BOTEXT").Text = Descrip
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
    'SAPConnection.reportError "Rebate_Description_Update", ActiveCell.Offset(0, 1), Err.Number, Err.Description, Err.Source, sBar.Text
    CMail.BuildErrorList ActiveCell.Offset(0, 1), "Rebate_Description_Update", Err.Number, Err.Description, Err.Source, sBar.Text
    Err.Clear
    SAPConnection.errorContinueNextItem (trx)
End Sub


