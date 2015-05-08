Attribute VB_Name = "UpdateStatusSO"
Option Explicit
Dim sBar

Sub UpdateStatusSOScript(ByRef SO As String, ByRef Statusx As String, ByRef trx As String, ByRef SAPConnection As Variant, ByRef CMail As Variant)
    On Error GoTo Errorhandler
    Set sBar = SAPConnection.session.findById("wnd[0]/sbar")
    
    SAPConnection.session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = SO
    SAPConnection.session.findById("wnd[0]").sendVKey 0

    If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
        Do
            SAPConnection.session.findById("wnd[1]").sendVKey 0
        Loop While Not SAPConnection.session.findById("wnd[1]", False) Is Nothing
    End If
    
    If sBar.Text = "Fin ext cust and/or Selling BU missing, please check your entries!" Then
        SAPConnection.session.findById("wnd[0]").sendVKey 0
    End If
    
    SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0,0]").SetFocus
    SAPConnection.session.findById("wnd[0]").sendVKey 2
    
    Dim eleventhTab, twelfthTab As String
    eleventhTab = SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\11").Text
    twelfthTab = SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\12").Text
    
    ' kolla vilken tab den ska gå till (11:e eller 12:e)
    Select Case "Status"
        Case eleventhTab
        SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\11").Select
        Case twelfthTab
        SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\12").Select
    End Select
    
    ' Line item loop start
    
    Do
        ' Definera materialtyper som den kan hoppa över
        Dim itemCat
        Set itemCat = SAPConnection.session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4013/ctxtVBAP-PSTYV")
        
        If itemCat.Text = "ZVCO" Or itemCat.Text = "ZHSS" Then GoTo Fortsaett
        
        Dim chTech
        Set chTech = SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY:SAPMV45A:4456/txtRV45A-STTXT")
        
        
        ' Om man ska sätta TECO ..
        If Statusx = "Set TECO" Then
            If InStr(chTech.Text, "TECO") > 0 Then GoTo Fortsaett
        ElseIf Statusx = "Remove TECO" Then ' Om man ska ta bort TECO eller CLSD ..
            If InStr(chTech.Text, "REL") > 0 Then
                GoTo Fortsaett
            ElseIf InStr(chTech.Text, "CLSD") > 0 Then
                ActiveCell.Offset(0, 3).value = "Found CLSD item"
                GoTo Fortsaett
            End If
        ElseIf Statusx = "Remove CLSD" Then
            If InStr(chTech.Text, "TECO") > 0 Or InStr(chTech.Text, "REL") > 0 Then GoTo Fortsaett
        ElseIf Statusx = "Set CLSD" Then
            If InStr(chTech.Text, "CLSD") > 0 Then GoTo Fortsaett
        End If

        If InStr(chTech.Text, "NoMP") > 0 Then GoTo Fortsaett
        
        If (SAPConnection.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4456/btnBT_STAE")) Is Nothing Then GoTo Fortsaett
        If Statusx = "Set TECO" Then
            SAPConnection.session.ActiveWindow.FindByName("BT_STAE", "GuiButton").press
            SAPConnection.session.findById("wnd[1]/usr/btnFCODE_BTAB").press ' original SÄTT teco
        ElseIf Statusx = "Remove TECO" Then
            SAPConnection.session.ActiveWindow.FindByName("BT_STAE", "GuiButton").press
            SAPConnection.session.ActiveWindow.FindByName("FCODE_BUTA", "GuiButton").press ' Ta BORT TECO
        ElseIf Statusx = "Remove CLSD" Then
            SAPConnection.session.ActiveWindow.FindByName("BT_STAE", "GuiButton").press
            SAPConnection.session.ActiveWindow.FindByName("FCODE_BUAB", "GuiButton").press ' Ta BORT CLSD
        ElseIf Statusx = "Set CLSD" Then
            SAPConnection.session.ActiveWindow.FindByName("BT_STAE", "GuiButton").press
            SAPConnection.session.ActiveWindow.FindByName("FCODE_STAB", "GuiButton").press ' Sätt CLSD
         ElseIf Statusx = "Set FNBL" Then
            SAPConnection.session.ActiveWindow.FindByName("BT_STAE", "GuiButton").press
            SAPConnection.session.ActiveWindow.FindByName("FCODE_STEF", "GuiButton").press ' Sätt CLSD
       End If
        
Fortsaett:         ' Till nästa line item
        SAPConnection.session.findById("wnd[0]/tbar[1]/btn[19]").press
        
        On Error Resume Next ' om det inte dyker upp en ruta att trycka enter på (nedan)
        SAPConnection.session.findById("wnd[1]").sendVKey 0
        Loop While sBar.Text <> "There are no more items to be displayed"
        'slut
        
        SAPConnection.session.findById("wnd[0]/tbar[0]/btn[11]").press
        
        If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
            If SAPConnection.session.findById("wnd[1]").Text = "Information" Then
                SAPConnection.session.findById("wnd[1]").sendVKey 0
            ElseIf SAPConnection.session.findById("wnd[1]").Text = "Save Incomplete Document" Then
                SAPConnection.session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
            ElseIf SAPConnection.session.findById("wnd[1]").Text = "Copy Text" Then
                Do
                    SAPConnection.session.findById("wnd[1]/tbar[0]/btn[6]").press
                Loop While Not SAPConnection.session.findById("wnd[1]", False) Is Nothing And SAPConnection.session.findById("wnd[1]").Text = "Copy Text"
            Else
                SAPConnection.session.findById("wnd[1]").sendVKey 0
        End If
            
        End If

        If InStr(sBar.Text, "Master cost missing") > 0 Then
            SAPConnection.session.ActiveWindow.FindByName("btn[12]", "GuiButton").press
            SAPConnection.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
            ActiveCell.Offset(0, 3) = "Master cost missing, no change"
            GoTo NextVC
        End If
        
        If InStr(sBar.Text, "Employee 23154149 is marked for deletion") > 0 Then
            SAPConnection.session.findById("wnd[0]").sendVKey 0
        End If
        
        If InStr(sBar.Text, "Fin ext cust and/or Selling BU missing, please check your entries!") > 0 Then
            SAPConnection.session.findById("wnd[0]").sendVKey 0
        End If
        
        If InStr(sBar.Text, "The cost estimate is being saved") > 0 Then
            GoTo NextVC
        End If
        
        If InStr(sBar.Text, "Partner ZG 911132 not maintained for partner 651537") > 0 Then
            SAPConnection.session.findById("wnd[0]/tbar[0]/btn[15]").press
            SAPConnection.session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
            ActiveCell.Offset(0, 3).value = "Partner ZG 911132 not maintained for partner 651537"
            ActiveCell.value = 1
            SAPConnection.session.findById("wnd[0]/tbar[0]/okcd").Text = "/nva02"
            SAPConnection.session.findById("wnd[0]").sendVKey 0
            GoTo NextVC
        End If
    
    If sBar.Text = "" Then
    GoTo Errorhandler
    
    ElseIf sBar.MessageType = "E" Then
    
    ActiveCell.Offset(0, 3) = "ERROR: " & sBar.Text & ", " & Format(Now(), "yyyy/mm/dd | hh:mm")
    ActiveCell.value = 0
    Else
    ActiveCell.Offset(0, 3) = sBar.Text & ", " & Format(Now(), "yyyy/mm/dd | hh:mm")
    ActiveCell.value = 1
    End If
NextVC:
    Exit Sub

Errorhandler:
    ' Debug.Print Err.Number & " " & Err.Description & " " & Err.Source
    If sBar.Text = "Fin ext cust and/or Selling BU missing, please check your entries!" Then
        ActiveCell.Offset(0, 3).value = sBar.Text
        GoTo NextVC
    ElseIf InStr(sBar.Text, "Master cost missing") > 0 Then
        SAPConnection.session.findById("wnd[0]").sendVKey 0
    Else
        GoTo ErrGoNext
    End If
    Exit Sub
ErrGoNext:
    SAPConnection.ErrorCounter = SAPConnection.ErrorCounter + 1
    CMail.BuildErrorList ActiveCell.Offset(0, 1), "UpdateStatusSO", Err.Number, Err.Description, Err.Source, sBar.Text
    Err.Clear
    SAPConnection.errorContinueNextItem (trx)
End Sub
