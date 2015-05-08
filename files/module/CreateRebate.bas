Attribute VB_Name = "CreateRebate"
Option Explicit

Sub Create_RebateScript(ByRef VC As String, ByRef Description As String, ByRef ValidFrom As String, ByRef Validto As String, ByRef Soldto As String, ByRef Curr As String, ByRef Mat As String, ByRef pCode As String, ByRef Percen As String, ByRef trx As String, ByRef SAPConnection As Variant, ByRef CMail As Variant)
    Dim sBar
    Dim nscroll As String
    Dim startingCellRow As String
    Dim newrow As String
        
    Set sBar = SAPConnection.session.findById("wnd[0]/sbar")
        nscroll = 0
        startingCellRow = ActiveCell.Row
        On Error GoTo ErrHandler

SAPConnection.session.findById("wnd[0]").resizeWorkingPane 99, 11, False
SAPConnection.session.findById("wnd[0]/usr/ctxtRV13A-BOART_BO").Text = "Z007"
SAPConnection.session.findById("wnd[0]").sendVKey 9
        If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
SAPConnection.session.findById("wnd[1]/usr/ctxtKONA-VKORG").Text = "1337"
SAPConnection.session.findById("wnd[1]/usr/ctxtKONA-VTWEG").Text = "XX"
SAPConnection.session.findById("wnd[1]/usr/ctxtKONA-SPART").Text = "XX"
SAPConnection.session.findById("wnd[1]/usr/ctxtKONA-VKBUR").Text = ""
SAPConnection.session.findById("wnd[1]/usr/ctxtKONA-VKGRP").Text = ""
SAPConnection.session.findById("wnd[1]/usr/ctxtKONA-VKGRP").SetFocus
SAPConnection.session.findById("wnd[1]/usr/ctxtKONA-VKGRP").caretPosition = 3
SAPConnection.session.findById("wnd[1]").sendVKey 0
Else
End If
SAPConnection.session.findById("wnd[0]/usr/txtKONA-BOTEXT").Text = Description
SAPConnection.session.findById("wnd[0]/usr/ctxtKONA-BONEM").Text = Soldto
SAPConnection.session.findById("wnd[0]/usr/ctxtKONA-WAERS").Text = Curr
SAPConnection.session.findById("wnd[0]/usr/txtKONA-ABREX").Text = Description
SAPConnection.session.findById("wnd[0]/usr/ctxtKONA-IDENT3").Text = ""
SAPConnection.session.findById("wnd[0]/usr/ctxtKONA-DATAB").Text = ValidFrom
SAPConnection.session.findById("wnd[0]/usr/ctxtKONA-DATBI").Text = Validto
SAPConnection.session.findById("wnd[0]/usr/ctxtKONA-DATBI").SetFocus
SAPConnection.session.findById("wnd[0]/usr/ctxtKONA-DATBI").caretPosition = 10
SAPConnection.session.findById("wnd[0]").sendVKey 0
SAPConnection.session.findById("wnd[0]/usr/ctxtKONA-DATAB").Text = ValidFrom
SAPConnection.session.findById("wnd[0]/usr/ctxtKONA-DATAB").SetFocus
SAPConnection.session.findById("wnd[0]/usr/ctxtKONA-DATAB").caretPosition = 10
SAPConnection.session.findById("wnd[0]").sendVKey 0
SAPConnection.session.findById("wnd[0]/tbar[1]/btn[23]").press
SAPConnection.session.findById("wnd[0]/tbar[1]/btn[9]").press
SAPConnection.session.findById("wnd[0]/usr/txtKOMG-ZZCON").Text = VC
SAPConnection.session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/txtKOMG-MVGR1[0,0]").Text = pCode
SAPConnection.session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/txtKONP-KBETR[2,0]").Text = Round(Percen, 0)
SAPConnection.session.findById("wnd[0]").sendVKey 0
    If ActiveCell.Offset(0, 1) = ActiveCell.Offset(1, 1) Then ' Check for a different VC
        Do
            ActiveCell.Offset(1, 0).Select
            pCode = ActiveCell.Offset(0, 8).value
            Percen = ActiveCell.Offset(0, 9).value
            
            SAPConnection.session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY").verticalScrollbar.Position = nscroll
            SAPConnection.session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/txtKOMG-MVGR1[0,1]").Text = pCode
            SAPConnection.session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/txtKONP-KBETR[2,1]").Text = Round(Percen, 0)
            SAPConnection.session.findById("wnd[0]").sendVKey 0
            nscroll = nscroll + 1
        Loop While ActiveCell.Offset(0, 1) = ActiveCell.Offset(1, 1) ' Loop while the VC is the same as the current
    End If
SAPConnection.session.findById("wnd[0]").sendVKey 0
SAPConnection.session.findById("wnd[0]/tbar[1]/btn[38]").press
newrow = startingCellRow - ActiveCell.Row
            ActiveCell.Offset(newrow, 0).Select
            Mat = ActiveCell.Offset(0, 7).value
            nscroll = 1
            SAPConnection.session.findById("wnd[0]/usr/sub:SAPMV13A:0411/ctxtKONP-BOMAT[2,16]").Text = Mat
    If ActiveCell.Offset(0, 1) = ActiveCell.Offset(1, 1) Then ' Check for a different VC
        Do
            ActiveCell.Offset(1, 0).Select
            Mat = ActiveCell.Offset(0, 7).value
            
            SAPConnection.session.findById("wnd[0]/usr").verticalScrollbar.Position = nscroll
            SAPConnection.session.findById("wnd[0]/usr/sub:SAPMV13A:0411/ctxtKONP-BOMAT[2,16]").Text = Mat
            nscroll = nscroll + 1
        Loop While ActiveCell.Offset(0, 1) = ActiveCell.Offset(1, 1) ' Loop while the VC is the same as the current
    End If
SAPConnection.session.findById("wnd[0]/tbar[0]/btn[11]").press
        Range(Cells(startingCellRow, 12), Cells(ActiveCell.Row, 12)) = Mid(sBar.Text, 21, 4)

        Range(Cells(startingCellRow, 2), Cells(ActiveCell.Row, 2)) = 1
    ActiveCell.value = 1
    Exit Sub
    
ErrHandler:
    If Err.Number <> 0 Then
        GoTo ErrGoNext
    End If
    Exit Sub
ErrGoNext:
    SAPConnection.ErrorCounter = SAPConnection.ErrorCounter + 1
    'SAPConnection.reportError "Create_Rebate", ActiveCell.Offset(0, 1), Err.Number, Err.Description, Err.Source, sBar.Text
    CMail.BuildErrorList ActiveCell.Offset(0, 1), "Create_Rebate", Err.Number, Err.Description, Err.Source, sBar.Text
    Err.Clear
    SAPConnection.errorContinueNextItem (trx)
End Sub


