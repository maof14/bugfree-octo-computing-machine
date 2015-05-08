Attribute VB_Name = "Gstream_FAS_Update"
Option Explicit
Dim sBar
Dim GstreamKey As String
Dim startingCellRow As Integer

Sub Update_Gstream_AssignmentIDScript(ByRef VC As String, ByRef FAS As String, ByRef Gstream As String, ByRef trx As String, ByRef SAPConnection As Variant, ByRef CMail As Variant)
    
    On Error GoTo ErrHandler
    Set sBar = SAPConnection.session.findById("wnd[0]/sbar")
    
    startingCellRow = ActiveCell.Row
    
    SAPConnection.session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = VC ' = "40039503"
    SAPConnection.session.findById("wnd[0]").sendVKey 0
    
    If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
        SAPConnection.session.findById("wnd[1]").sendVKey 0
    End If
    
    If Gstream = "Customer Project" Then
    GstreamKey = "01"
    ElseIf Gstream = "Product Delivery" Then
    GstreamKey = "02"
    ElseIf Gstream = "Customer Support" Then
    GstreamKey = "03"
    ElseIf Gstream = "Managed Operations" Then
    GstreamKey = "04"
    ElseIf Gstream = "Contract Financial Adjustments" Then
    GstreamKey = "99"
    Else
    GstreamKey = ""
    End If

    SAPConnection.session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13").Select
    
    If Gstream <> "" Then
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/cmbVBAK-ZZGOVSTREAM").Key = GstreamKey
    End If
    
    If FAS <> "" Then
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBAK-ZZASSIGNID").Text = FAS
    End If
    SAPConnection.session.findById("wnd[0]").sendVKey 0
    If sBar.Text <> "" Then
        GoTo finscript:
    End If
    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[11]").press
    If InStr(SAPConnection.session.findById("wnd[1]").Text, "Partner") > 0 Then
       SAPConnection.session.findById("wnd[1]/usr/btnBUTTON_1").press
    End If
    If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
        SAPConnection.session.findById("wnd[1]").sendVKey 0
    End If
    GoTo finscript:

finscript:

    Range(Cells(startingCellRow, 6), Cells(ActiveCell.Row, 6)) = sBar.Text
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
    'SAPConnection.reportError "UpdateGstreamFAS", ActiveCell.Offset(0, 1), Err.Number, Err.Description, Err.Source, sBar.Text
    CMail.BuildErrorList ActiveCell.Offset(0, 1), "UpdateGstreamFAS", Err.Number, Err.Description, Err.Source, sBar.Text
    Err.Clear
    SAPConnection.errorContinueNextItem (trx)
End Sub

