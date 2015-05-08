Attribute VB_Name = "RunSettlementQTC"
Option Explicit

Sub RunSettlementWBSScript(ByRef WBS As String, ByRef yMonth As String, ByRef yYear As String, ByRef trx As String, ByRef SAPConnection As Variant, ByRef CMail As Variant)
    Dim startingCellRow As Integer
    Dim CJ88result As String
    Dim sBar
    Set sBar = SAPConnection.session.findById("wnd[0]/sbar")
    On Error GoTo ErrHandler

    startingCellRow = ActiveCell.Row
    
    If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
        SAPConnection.session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").Text = "1000"
        SAPConnection.session.findById("wnd[1]").sendVKey 0
    End If
    SAPConnection.session.findById("wnd[0]/usr/chkRKAUF-TEST").Selected = False
    SAPConnection.session.findById("wnd[0]/usr/subBLOCK1:SAPLKAOP:0570/ctxtLKP74-VBELN").Text = ""
    SAPConnection.session.findById("wnd[0]/usr/subBLOCK1:SAPLKAOP:0570/ctxtLKP74-PSPID").Text = Left(WBS, 12)
    SAPConnection.session.findById("wnd[0]/usr/subBLOCK1:SAPLKAOP:0570/ctxtLKP74-POSID").Text = ""
    SAPConnection.session.findById("wnd[0]/usr/subBLOCK2:SAPLKAZB:2100/txtRKAUF-FROM").Text = yMonth
    SAPConnection.session.findById("wnd[0]/usr/subBLOCK2:SAPLKAZB:2100/txtRKAUF-GJAHR").Text = yYear
    SAPConnection.session.findById("wnd[0]/usr/chkRKAUF-TEST").SetFocus
    SAPConnection.session.findById("wnd[0]").sendVKey 8
    ActiveCell.Offset(0, 4).value = SAPConnection.session.findById("wnd[0]/usr/lbl[53,21]").Text & " New objects for incoming orders"
    SAPConnection.session.findById("wnd[0]/tbar[0]/okcd").Text = "/NCNE1"
    SAPConnection.session.findById("wnd[0]").sendVKey 0
    SAPConnection.session.findById("wnd[0]/usr/subBLOCK1:SAPLCNEV_01_MASTER_DATA:0550/chkLKP74-INCL_HIER").Selected = True
    SAPConnection.session.findById("wnd[0]/usr/subBLOCK1:SAPLCNEV_01_MASTER_DATA:0550/chkLKP74-INCL_AUF").Selected = True
    SAPConnection.session.findById("wnd[0]/usr/chkRKAUF-TEST").Selected = False
    SAPConnection.session.findById("wnd[0]/usr/subBLOCK1:SAPLCNEV_01_MASTER_DATA:0550/ctxtLKP74-PSPID").Text = Left(WBS, 12)
    SAPConnection.session.findById("wnd[0]/usr/subBLOCK1:SAPLCNEV_01_MASTER_DATA:0550/ctxtLKP74-POSID").Text = ""
    SAPConnection.session.findById("wnd[0]/usr/subBLOCK1:SAPLCNEV_01_MASTER_DATA:0550/ctxtLKP74-NPLNR").Text = ""
    SAPConnection.session.findById("wnd[0]/usr/subBLOCK2:SAPLKAZB:2550/ctxtRPSCO_X-VERSN_EV").Text = "200"
    SAPConnection.session.findById("wnd[0]/usr/subBLOCK2:SAPLKAZB:2550/txtRKAUF-TO").Text = yMonth
    SAPConnection.session.findById("wnd[0]/usr/chkRKAUF-TEST").SetFocus
    SAPConnection.session.findById("wnd[0]/tbar[1]/btn[8]").press
    ActiveCell.Offset(0, 5).value = sBar.Text & " Done"
    SAPConnection.session.findById("wnd[0]/tbar[0]/okcd").Text = "/NKKA2"
    SAPConnection.session.findById("wnd[0]").sendVKey 0
    SAPConnection.session.findById("wnd[0]/usr/ctxtKKA0100-POSID").Text = WBS
    SAPConnection.session.findById("wnd[0]/usr/txtKKA0100-BIS_ABGR_M").Text = yMonth
    SAPConnection.session.findById("wnd[0]/usr/txtKKA0100-BIS_ABGR_J").Text = yYear
    SAPConnection.session.findById("wnd[0]/usr/ctxtKKA0100-VERSN").Text = "0"
    SAPConnection.session.findById("wnd[0]/tbar[1]/btn[8]").press
    If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
        Do
            SAPConnection.session.findById("wnd[1]/tbar[0]/btn[0]").press
        Loop While Not SAPConnection.session.findById("wnd[1]", False) Is Nothing
    End If
    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[11]").press
    SAPConnection.session.findById("wnd[1]/tbar[0]/btn[0]").press
    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[11]").press
            ActiveCell.Offset(0, 0).value = 1
            ActiveCell.Offset(0, 6).value = sBar.Text
    If Left(ActiveCell.Offset(0, 1).value, 12) = Left(ActiveCell.Offset(1, 1).value, 12) Then ' Check for a different Project Definition
        Do
            ActiveCell.Offset(1, 0).Select
            ActiveCell.Offset(0, 0).value = 1
            ActiveCell.Offset(0, 4).value = ActiveCell.Offset(-1, 4).value
            ActiveCell.Offset(0, 5).value = ActiveCell.Offset(-1, 5).value
            WBS = ActiveCell.Offset(0, 1).value
            SAPConnection.session.findById("wnd[0]/tbar[0]/okcd").Text = "/NKKA2"
            SAPConnection.session.findById("wnd[0]").sendVKey 0
            SAPConnection.session.findById("wnd[0]/usr/ctxtKKA0100-POSID").Text = WBS
            SAPConnection.session.findById("wnd[0]/usr/txtKKA0100-BIS_ABGR_M").Text = yMonth
            SAPConnection.session.findById("wnd[0]/usr/txtKKA0100-BIS_ABGR_J").Text = yYear
            SAPConnection.session.findById("wnd[0]/usr/ctxtKKA0100-VERSN").Text = "0"
            SAPConnection.session.findById("wnd[0]/tbar[1]/btn[8]").press
    If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
        Do
            SAPConnection.session.findById("wnd[1]/tbar[0]/btn[0]").press
        Loop While Not SAPConnection.session.findById("wnd[1]", False) Is Nothing
    End If
            SAPConnection.session.findById("wnd[0]/tbar[0]/btn[11]").press
            SAPConnection.session.findById("wnd[1]/tbar[0]/btn[0]").press
            SAPConnection.session.findById("wnd[0]/tbar[0]/btn[11]").press
            ActiveCell.Offset(0, 6).value = sBar.Text
        Loop While Left(ActiveCell.Offset(0, 1).value, 12) = Left(ActiveCell.Offset(1, 1).value, 12) ' Loop while the Project Definition is the same as the current
    End If
    SAPConnection.session.findById("wnd[0]/tbar[0]/okcd").Text = "/NCJ88"
    SAPConnection.session.findById("wnd[0]").sendVKey 0
    SAPConnection.session.findById("wnd[0]/usr/subBLOCK1:SAPLKAOP:0550/chkLKP74-INCL_HIER").Selected = True
    SAPConnection.session.findById("wnd[0]/usr/subBLOCK1:SAPLKAOP:0550/chkLKP74-INCL_AUF").Selected = True
    SAPConnection.session.findById("wnd[0]/usr/chkLKO74-TESTLAUF").Selected = False
    SAPConnection.session.findById("wnd[0]/usr/subBLOCK1:SAPLKAOP:0550/ctxtLKP74-PSPID").Text = Left(WBS, 12)
    SAPConnection.session.findById("wnd[0]/usr/subBLOCK1:SAPLKAOP:0550/ctxtLKP74-POSID").Text = ""
    SAPConnection.session.findById("wnd[0]/usr/subBLOCK1:SAPLKAOP:0550/ctxtLKP74-NPLNR").Text = ""
    SAPConnection.session.findById("wnd[0]/usr/ctxtLKO74-PERIO").Text = yMonth
    SAPConnection.session.findById("wnd[0]/usr/txtLKO74-GJAHR").Text = yYear
    SAPConnection.session.findById("wnd[0]/usr/cmbLKO74-VAART").Key = "P"
    SAPConnection.session.findById("wnd[0]/usr/chkLKO74-TESTLAUF").SetFocus
    SAPConnection.session.findById("wnd[0]/tbar[1]/btn[8]").press
    SAPConnection.session.findById("wnd[0]/tbar[0]/okcd").Text = "/NCJA2"
    SAPConnection.session.findById("wnd[0]").sendVKey 0
    Range(Cells(startingCellRow, 9), Cells(ActiveCell.Row, 9)) = "Done"

Exit Sub
    
ErrHandler:
    If Err.Number <> 0 Then
        GoTo ErrGoNext
    End If
    Exit Sub
ErrGoNext:
    SAPConnection.ErrorCounter = SAPConnection.ErrorCounter + 1
    CMail.BuildErrorList ActiveCell.Offset(0, 1), "Run_Settlement", Err.Number, Err.Description, Err.Source, sBar.Text
    Err.Clear
    SAPConnection.errorContinueNextItem (trx)
End Sub

