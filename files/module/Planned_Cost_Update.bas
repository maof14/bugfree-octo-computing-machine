Attribute VB_Name = "Planned_Cost_Update"
Option Explicit
Dim startingCellRow As Integer
Dim Description, Amount, Curr, CostEl As String
Dim sBar
Dim budgetWarning, tecoRemovedNW, tecoRemovedWBS As Boolean

Sub Planned_Cost_UpdateScript(ByRef WBS As String, ByRef trx As String, ByRef SAPConnection As Variant, ByRef CMail As Variant)
    budgetWarning = False
    tecoRemovedNW = False
    tecoRemovedWBS = False
    Set sBar = SAPConnection.session.findById("wnd[0]/sbar")
    
    On Error GoTo ErrHandler
    
    startingCellRow = ActiveCell.Row
        
    SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell").pressButton "OPEN"
    SAPConnection.session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PROJ_EXT").Text = ""
    SAPConnection.session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PRPS_EXT").Text = WBS
    SAPConnection.session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-AUFNR").Text = ""
    SAPConnection.session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PRPS_EXT").SetFocus
    SAPConnection.session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PRPS_EXT").caretPosition = 19
    SAPConnection.session.findById("wnd[1]").sendVKey 0
    
    If Not SAPConnection.session.ActiveWindow.FindByName("PRPS-POST1", "GuiTextField").Changeable = True Then
        SAPConnection.session.findById("wnd[0]/tbar[1]/btn[13]").press
    End If
    'If object is blocked by another user proceed to next one
        If InStr(sBar.Text, "locked") > 0 Then
            ActiveCell.value = 0
            Cells(startingCellRow, 8).value = "WBS BLOCKED by another user, try again later."
            ActiveCell.Offset(-1, 0).Select
            Exit Sub
        End If
    
    If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
        Do
                SAPConnection.session.findById("wnd[1]").sendVKey 0
        Loop While Not SAPConnection.session.findById("wnd[1]", False) Is Nothing
    End If
    
    'check if WBS is in TECO, and release, will be put back in TECO in line 109
    If InStr(SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpGRND/ssubSUBSCR1:SAPLCJWB:1210/subSTATUS:SAPLCJWB:0700/txtCNJ_STAT-STTXT_INT").Text, "TECO") > 0 Then
        SAPConnection.session.findById("wnd[0]/mbar/menu[1]/menu[2]/menu[4]/menu[1]").Select ' remove TECO on chosen WBS
        tecoRemovedWBS = True
    End If
    'check if WBS is in CLSD, and continue to next item
    If InStr(SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCJWB:3999/tabsTABCJWB/tabpGRND/ssubSUBSCR1:SAPLCJWB:1210/subSTATUS:SAPLCJWB:0700/txtCNJ_STAT-STTXT_INT").Text, "CLSD") > 0 Then
            ActiveCell.value = 1
            Cells(startingCellRow, 8).value = "WBS in status CLSD, update not possible."
            ActiveCell.Offset(-1, 0).Select
            Exit Sub
    End If
    
    SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
    ' Nedan if-rad kontrollerar om det INTE finns ett nätverk.
    ' Om det inte finns så skapar den ett.
    ' Om det finns, så kollar den om det är i CLSD eller TECO.
    ' Men det funkar tydligen inte om det är två som finns. Kontrollera hur man kan lösa detta.
    ' Ser ut som att den lägger in på första..
    If SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,0]").Text = "" Then
        SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").ExpandNode "          1"
        SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").SelectedNode = "          4"
        SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").topNode = "          1"
        SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").DoubleClickNode "          4"
        If SAPConnection.session.ActiveWindow.FindByName("CAUFVD-PROFID", "GuiComboBox").Changeable = True Then
            SAPConnection.session.ActiveWindow.FindByName("CAUFVD-PROFID", "GuiComboBox").Key = "ZEAB001"
        End If
        SAPConnection.session.ActiveWindow.FindByName("CAUFVD-DISPO", "GuiCTextField").Text = "001"
        SAPConnection.session.findById("wnd[0]").sendVKey 0
        SAPConnection.session.findById("wnd[0]").sendVKey 0
        SAPConnection.session.findById("wnd[0]").sendVKey 0
        SAPConnection.session.findById("wnd[0]").sendVKey 0
        SAPConnection.session.findById("wnd[0]").sendVKey 0
        SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").ExpandNode "          5"
        SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").SelectedNode = "         11"
        SAPConnection.session.findById("wnd[0]/mbar/menu[1]/menu[2]/menu[0]").Select
    Else
        SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,0]").SetFocus
        SAPConnection.session.findById("wnd[0]").sendVKey 2
        
        If InStr(SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN/ssubSUBSCR_2100:SAPLCOKO:2110/txtCAUFVD-STTXT").Text, "TECO") > 0 Then
            SAPConnection.session.findById("wnd[0]/mbar/menu[1]/menu[2]/menu[4]/menu[1]").Select ' remove TECO on NW
            tecoRemovedNW = True
        ElseIf InStr(SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOKO:2101/tabsTABSTR_2100/tabpTRMN/ssubSUBSCR_2100:SAPLCOKO:2110/txtCAUFVD-STTXT").Text, "CLSD") > 0 Then
            ActiveCell.value = 1
            Cells(startingCellRow, 8).value = "Network in status CLSD, update not possible."
            ActiveCell.Offset(-1, 0).Select
            Exit Sub
        End If
    End If
        
    Do
        CostEl = ActiveCell.Offset(0, 5)
        'Call checkCostElement(CostEl, trx, SAPConnection)
        If Left(CostEl, 1) <> 4 Then
                If ActiveCell.Offset(0, 1) <> ActiveCell.Offset(1, 1) And ActiveCell.Offset(0, 1) = ActiveCell.Offset(-1, 1) Then
                    ActiveCell.value = 0
                    ActiveCell.Offset(0, 6).value = "Cost element " & CostEl & " is not a primary cost element"
                    ActiveCell.Offset(2, 0).Select
                    GoTo SaveWBS
                ElseIf ActiveCell.Offset(0, 1) <> ActiveCell.Offset(1, 1) And ActiveCell.Offset(0, 1) <> ActiveCell.Offset(-1, 1) Then
                    ActiveCell.value = 0
                    ActiveCell.Offset(0, 6).value = "Cost element " & CostEl & " is not a primary cost element"
                    'ActiveCell.Offset(-1, 0).Select
                    SAPConnection.session.findById("wnd[0]/tbar[0]/okcd").Text = "/NCJ20N"
                    SAPConnection.session.findById("wnd[0]").sendVKey 0
                    Exit Sub
                Else
                    ActiveCell.value = 0
                    ActiveCell.Offset(0, 6).value = "Cost element " & CostEl & " is not a primary cost element"
                    ActiveCell.Offset(1, 0).Select
                End If
        Else
        End If
        Description = ActiveCell.Offset(0, 2)
        Amount = ActiveCell.Offset(0, 3)
        Curr = ActiveCell.Offset(0, 4)
        CostEl = ActiveCell.Offset(0, 5)
        SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").topNode = "          1"
        SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").DoubleClickNode "         11"
        SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subIDENTIFICATION:SAPLCONW:0110/txtAFVGM-LTXA1").Text = Description
        SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subIDENTIFICATION:SAPLCONW:0110/txtAFVGM-LTXA1").SetFocus
        SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subIDENTIFICATION:SAPLCONW:0110/txtAFVGM-LTXA1").caretPosition = 18
        SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpKOSD/ssubSUBSCR_1000:SAPLCONW:1550/txtAFVGD-PRKST").Text = Round(Amount, 2)
        SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpKOSD/ssubSUBSCR_1000:SAPLCONW:1550/txtAFVGD-PRKST").SetFocus
        SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpKOSD/ssubSUBSCR_1000:SAPLCONW:1550/txtAFVGD-PRKST").caretPosition = 15
        SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpKOSD/ssubSUBSCR_1000:SAPLCONW:1550/ctxtAFVGD-WAERS").Text = Curr
        SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpKOSD/ssubSUBSCR_1000:SAPLCONW:1550/ctxtAFVGD-WAERS").SetFocus
        SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpKOSD/ssubSUBSCR_1000:SAPLCONW:1550/ctxtAFVGD-WAERS").caretPosition = 3
        SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpKOSD/ssubSUBSCR_1000:SAPLCONW:1550/ctxtAFVGD-SAKTO").Text = CostEl
        SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCONW:1001/tabsTABSTRIP_1000/tabpKOSD/ssubSUBSCR_1000:SAPLCONW:1550/ctxtAFVGD-SAKTO").caretPosition = 6
        SAPConnection.session.findById("wnd[0]").sendVKey 0
        ActiveCell.value = 1
        ActiveCell.Offset(1, 0).Select
    Loop Until ActiveCell.Offset(0, 1) <> ActiveCell.Offset(-1, 1)

SaveWBS:
    
    SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").SelectedNode = "000003"
    
    If tecoRemovedNW = True Then ' check if TECO was removed
        SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").SelectedNode = "000002"
        SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_OVERVIEW/shellcont/shell").pressButton "NETW_OVW"
        SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010").GetAbsoluteRow(0).Selected = True
        SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,0]").SetFocus
        SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCNPB_M:2010/tblSAPLCNPB_MTCTRL_2010/txtNETW_OVW-AUFNR[0,0]").caretPosition = 0
        SAPConnection.session.findById("wnd[0]/mbar/menu[1]/menu[2]/menu[4]/menu[0]").Select
    End If
    
    If tecoRemovedWBS = True Then ' check if TECO was removed
            SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").SelectedNode = "000002" ' Choose the WBS (to set TECO)
            SAPConnection.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/cntlTOOLBAR_CONTAINER_DETAIL/shellcont/shell").pressButton "WBSE_DET"
            SAPConnection.session.findById("wnd[0]/mbar/menu[1]/menu[2]/menu[4]/menu[0]").Select ' Set TECO back on the WBS
    End If
    
    SAPConnection.session.findById("wnd[0]/tbar[0]/btn[11]").press

    ' If there is a popup - loop through popup checks for Availability control, Scheduling, budget and any else pupop until there are no more popups.
    If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
        Do
            If InStr(SAPConnection.session.findById("wnd[1]").Text, "Availability Control") > 0 Then
                SAPConnection.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
            ElseIf InStr(SAPConnection.session.findById("wnd[1]").Text, "Scheduling") > 0 Then
                SAPConnection.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
            ElseIf InStr(SAPConnection.session.findById("wnd[1]").Text, "Commit") > 0 Then
                SAPConnection.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
            ElseIf InStr(SAPConnection.session.findById("wnd[1]").Text, "Cost") > 0 Then
                SAPConnection.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
            ElseIf InStr(SAPConnection.session.findById("wnd[1]").Text, "Budget") > 0 Then
                SAPConnection.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                budgetWarning = True
            Else
                SAPConnection.session.findById("wnd[1]").sendVKey 0
            End If
        Loop While Not SAPConnection.session.findById("wnd[1]", False) Is Nothing
    End If
    
'Check if the status bar is blank just to make sure it didnt save, if it didnt then it will add another "ENTER"
    If sBar.Text = "" Then
        SAPConnection.session.findById("wnd[0]").sendVKey 0
    End If
    
    SAPConnection.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").topNode = "         23"
    
    If InStr(sBar.Text, "budget") > 0 Then
    budgetWarning = True
    Do
            SAPConnection.session.findById("wnd[0]").sendVKey 0
    
    Loop While InStr(sBar.Text, "budget") > 0
    
    End If
  
    
    If budgetWarning = True Then
        Range(Cells(startingCellRow, 8), Cells(ActiveCell.Row - 1, 8)) = "=IF(LEFT(RC[-1],1)<>""4"",""Cost element ""&RC[-1]&"" is not a primary cost element, Use another CE"",""Budget exceeded, "" & """ & sBar.Text & """)"
        Range(Cells(startingCellRow, 2), Cells(ActiveCell.Row - 1, 2)) = 1

    Else
        Range(Cells(startingCellRow, 8), Cells(ActiveCell.Row - 1, 8)) = "=IF(LEFT(RC[-1],1)<>""4"",""Cost element ""&RC[-1]&"" is not a primary cost element, Use another CE"",""" & sBar.Text & """)"
    End If
    
    ActiveCell.Offset(-1, 0).Select
    SAPConnection.session.findById("wnd[0]/tbar[0]/okcd").Text = "/n" & trx
    SAPConnection.session.findById("wnd[0]").sendVKey 0

    Exit Sub
    
ErrHandler:
    If Err.Number <> 0 Then
        GoTo ErrGoNext
    End If
    Exit Sub
ErrGoNext:
    Stop
    Resume
    SAPConnection.ErrorCounter = SAPConnection.ErrorCounter + 1
    CMail.BuildErrorList ActiveCell.Offset(0, 1), "Planned_Cost_Update", Err.Number, Err.Description, Err.Source, sBar.Text
    Err.Clear
    SAPConnection.errorContinueNextItem (trx)
End Sub

