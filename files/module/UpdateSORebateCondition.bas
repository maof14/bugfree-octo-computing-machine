Attribute VB_Name = "UpdateSORebateCondition"
Option Explicit

Sub SO_Rebate_Condition_UpdateScript(ByRef Soldto As String, ByRef SalesO As String, ByRef FromDate As String, ByRef trx As String, ByRef SAPConnection As Variant, ByRef CMail As Variant)
    Dim sBar
    Dim startingCellRow As Integer
On Error GoTo ErrHandler

    Set sBar = SAPConnection.session.findById("wnd[0]/sbar")
    
    startingCellRow = ActiveCell.Row
    
        SAPConnection.session.findById("wnd[0]/usr/ctxtVBCOM-KUNDE").Text = Soldto
        SAPConnection.session.findById("wnd[0]/usr/ctxtVBCOM-AUDAT").Text = FromDate
        SAPConnection.session.findById("wnd[0]/usr/txtVBCOM-BSTKD").SetFocus
        SAPConnection.session.findById("wnd[0]/usr/txtVBCOM-BSTKD").caretPosition = 0
        SAPConnection.session.findById("wnd[0]/tbar[1]/btn[20]").press
        SAPConnection.session.findById("wnd[1]/usr/sub:SAPLKAB1:0400/chkRKAB1-XSUCH[6,0]").Selected = True
        SAPConnection.session.findById("wnd[1]/usr/sub:SAPLKAB1:0400/chkRKAB1-XSUCH[6,0]").SetFocus
        SAPConnection.session.findById("wnd[1]").sendVKey 82
        SAPConnection.session.findById("wnd[1]/tbar[0]/btn[0]").press
        SAPConnection.session.findById("wnd[2]/usr/sub:SAPLKAB1:0410/ctxtRKAB1-VONSL[0,23]").Text = SalesO
        If ActiveCell.Offset(1, 1).value = ActiveCell.Offset(0, 1).value Then
        ActiveCell.Offset(1, 0).Select
        SalesO = ActiveCell.Offset(0, 2).value
        SAPConnection.session.findById("wnd[2]/usr/sub:SAPLKAB1:0410/ctxtRKAB1-VONSL[1,23]").Text = SalesO
        End If
        If ActiveCell.Offset(1, 1).value = ActiveCell.Offset(0, 1).value Then
        
        ActiveCell.Offset(1, 0).Select
        SalesO = ActiveCell.Offset(0, 2).value
        SAPConnection.session.findById("wnd[2]/usr/sub:SAPLKAB1:0410/ctxtRKAB1-VONSL[2,23]").Text = SalesO
        End If
        If ActiveCell.Offset(1, 1).value = ActiveCell.Offset(0, 1).value Then
        
        ActiveCell.Offset(1, 0).Select
        SalesO = ActiveCell.Offset(0, 2).value
        SAPConnection.session.findById("wnd[2]/usr/sub:SAPLKAB1:0410/ctxtRKAB1-VONSL[3,23]").Text = SalesO
        End If
        If ActiveCell.Offset(1, 1).value = ActiveCell.Offset(0, 1).value Then
        
        ActiveCell.Offset(1, 0).Select
        SalesO = ActiveCell.Offset(0, 2).value
        SAPConnection.session.findById("wnd[2]/usr/sub:SAPLKAB1:0410/ctxtRKAB1-VONSL[4,23]").Text = SalesO
        End If
        If ActiveCell.Offset(1, 1).value = ActiveCell.Offset(0, 1).value Then
        
        ActiveCell.Offset(1, 0).Select
        SalesO = ActiveCell.Offset(0, 2).value
        SAPConnection.session.findById("wnd[2]/usr/sub:SAPLKAB1:0410/ctxtRKAB1-VONSL[5,23]").Text = SalesO
        End If
        
        SAPConnection.session.findById("wnd[2]/tbar[0]/btn[0]").press
        SAPConnection.session.findById("wnd[0]").sendVKey 0
        If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
        SAPConnection.session.findById("wnd[1]/usr/ctxtVBCOM-VKORG").Text = "1337"
        SAPConnection.session.findById("wnd[1]/usr/ctxtVBCOM-VTWEG").Text = "XX"
        SAPConnection.session.findById("wnd[1]/usr/ctxtVBCOM-SPART").Text = "XX"
        SAPConnection.session.findById("wnd[1]/usr/ctxtVBCOM-SPART").SetFocus
        SAPConnection.session.findById("wnd[1]/usr/ctxtVBCOM-SPART").caretPosition = 2
        SAPConnection.session.findById("wnd[1]/tbar[0]/btn[0]").press
        Else
        End If
        SAPConnection.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").SetCurrentCell -1, "NETPR"
        SAPConnection.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").SelectColumn "NETPR"
        SAPConnection.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").ContextMenu
        SAPConnection.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").SelectContextMenuItem "&FILTER"
        SAPConnection.session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN002-LOW").Text = "0" + Application.International(xlDecimalSeparator) + "01"
        SAPConnection.session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN002-HIGH").Text = "999999999"
        SAPConnection.session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN002-HIGH").caretPosition = 9
        SAPConnection.session.findById("wnd[1]/tbar[0]/btn[0]").press
        SAPConnection.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").SetCurrentCell -1, ""
        SAPConnection.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").SelectAll
        SAPConnection.session.findById("wnd[0]/mbar/menu[1]/menu[2]/menu[2]").Select
        SAPConnection.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        SAPConnection.session.findById("wnd[1]/usr/lbl[6,8]").SetFocus
        SAPConnection.session.findById("wnd[1]/usr/lbl[6,8]").caretPosition = 15
        SAPConnection.session.findById("wnd[1]").sendVKey 2
        SAPConnection.session.findById("wnd[0]/tbar[0]/okcd").Text = "/NVA05"
        SAPConnection.session.findById("wnd[0]").sendVKey 0

    
    'slut
    Range(Cells(startingCellRow, 2), Cells(ActiveCell.Row, 2)) = 1
    ActiveCell.Offset(0, 5).value = sBar.Text
    ActiveCell.value = 1
    Exit Sub
    
ErrHandler:
    If Err.Number <> 0 Then
        GoTo ErrGoNext
        Range(Cells(startingCellRow, 2), Cells(ActiveCell.Row, 2)) = 0
        Range(Cells(startingCellRow, 7), Cells(ActiveCell.Row, 7)) = "ERROR: " & sBar.Text
    End If
    Exit Sub
ErrGoNext:
    SAPConnection.ErrorCounter = SAPConnection.ErrorCounter + 1
    'SAPConnection.reportError "SO_Rebate_Condition_Update", ActiveCell.Offset(0, 1), Err.Number, Err.Description, Err.Source, sBar.Text
    CMail.BuildErrorList ActiveCell.Offset(0, 1), "SO_Rebate_Condition_Update", Err.Number, Err.Description, Err.Source, sBar.Text
    Err.Clear
    SAPConnection.errorContinueNextItem (trx)
End Sub

