Attribute VB_Name = "UpdatePartnerVC"
Option Explicit
Dim sBar
Dim nscroll As String
Dim PartnerKey As String
Dim startingCellRow As Integer

Sub UpdatePartnerVCScript(ByRef VC As String, ByRef Partner As String, ByRef EmplNumber As String, ByRef trx As String, ByRef SAPConnection As Variant, ByRef CMail As Variant)
    
    On Error GoTo ErrHandler
    Set sBar = SAPConnection.session.findById("wnd[0]/sbar")
    
    nscroll = 0
    startingCellRow = ActiveCell.Row
    
    SAPConnection.session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = VC ' = "40039503"
    SAPConnection.session.findById("wnd[0]").sendVKey 0
    
    If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
        SAPConnection.session.findById("wnd[1]").sendVKey 0
    End If
    
    On Error Resume Next
    
    
    SAPConnection.session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08").Select
GoTo Change_Partner
Change_Partner:

    If Partner = "Contract Accountable" Then
    PartnerKey = "ZC"
    ElseIf Partner = "Contract Responsable" Then
    PartnerKey = "YM"
    ElseIf Partner = "Order Responsable" Then
    PartnerKey = "Z3"
    ElseIf Partner = "Exec Responsable" Then
    PartnerKey = "ZP"
    ElseIf Partner = "Sponsor" Then
    PartnerKey = "KM"
    ElseIf Partner = "PSP" Then
    PartnerKey = "Z8"
    End If

        nscroll = 0

        If SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,0]").Key <> PartnerKey Then  ' Check if there are more items (PCODE)
            Do
                SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW").verticalScrollbar.Position = nscroll
                
                nscroll = nscroll + 1
            Loop While SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,0]").Text <> "" And SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,0]").Key <> PartnerKey ' Loop while the VC is the same as the current
        End If
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,0]").Key = PartnerKey
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,0]").Text = EmplNumber
    
    If ActiveCell.Offset(0, 1) = ActiveCell.Offset(1, 1) Then
    
    ActiveCell.value = 1
    ActiveCell.Offset(1, 0).Select
    Partner = ActiveCell.Offset(0, 2).value
    EmplNumber = ActiveCell.Offset(0, 3).value
    
    GoTo Change_Partner
    End If
    
    SAPConnection.session.findById("wnd[0]").sendVKey 0
    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[11]").press
    If InStr(SAPConnection.session.findById("wnd[1]").Text, "Change") > 0 Then
       SAPConnection.session.findById("wnd[1]/usr/btnBUTTON_1").press
    End If
    If InStr(SAPConnection.session.findById("wnd[1]").Text, "Partner") > 0 Then
       SAPConnection.session.findById("wnd[1]/usr/btnBUTTON_1").press
    End If
    If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
        Do
                SAPConnection.session.findById("wnd[1]").sendVKey 0
        Loop While Not SAPConnection.session.findById("wnd[1]", False) Is Nothing
    End If
    
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
    CMail.BuildErrorList ActiveCell.Offset(0, 1), "UpdatePartnerVC", Err.Number, Err.Description, Err.Source, sBar.Text
    Err.Clear
    SAPConnection.errorContinueNextItem (trx)
End Sub
