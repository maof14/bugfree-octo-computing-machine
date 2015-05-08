VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EBWfrm 
   Caption         =   "EBW Portal Reports"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11715
   OleObjectBlob   =   "EBWfrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EBWfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Dim i As Long
Dim j As Long
Dim msg As String
Dim arrItems() As String
Dim iRet As Integer

On Error GoTo Errorhandler

    ReDim arrItems(0 To PopularBox.ColumnCount - 1)
    For j = 0 To PopularBox.ListCount - 1
        If PopularBox.Selected(j) Then

            For i = 0 To PopularBox.ColumnCount - 1
                arrItems(i) = PopularBox.column(i, j)
            Next i
            msg = msg & Join(arrItems, ",") & vbCrLf
        End If
    Next j
    ReDim arrItems(0 To CleanUpBox.ColumnCount - 1)
    For j = 0 To CleanUpBox.ListCount - 1
        If CleanUpBox.Selected(j) Then

            For i = 0 To CleanUpBox.ColumnCount - 1
                arrItems(i) = CleanUpBox.column(i, j)
            Next i
            msg = msg & Join(arrItems, ",") & vbCrLf
        End If
    Next j
    ReDim arrItems(0 To AgingBox.ColumnCount - 1)
    For j = 0 To AgingBox.ListCount - 1
        If AgingBox.Selected(j) Then

            For i = 0 To AgingBox.ColumnCount - 1
                arrItems(i) = AgingBox.column(i, j)
            Next i
            msg = msg & Join(arrItems, ",") & vbCrLf
        End If
    Next j
    ReDim arrItems(0 To DeploymentBox.ColumnCount - 1)
    For j = 0 To DeploymentBox.ListCount - 1
        If DeploymentBox.Selected(j) Then

            For i = 0 To DeploymentBox.ColumnCount - 1
                arrItems(i) = DeploymentBox.column(i, j)
            Next i
            msg = msg & Join(arrItems, ",") & vbCrLf
        End If
    Next j
    ReDim arrItems(0 To KPIBox.ColumnCount - 1)
    For j = 0 To KPIBox.ListCount - 1
        If KPIBox.Selected(j) Then

            For i = 0 To KPIBox.ColumnCount - 1
                arrItems(i) = KPIBox.column(i, j)
            Next i
            msg = msg & Join(arrItems, ",") & vbCrLf
        End If
    Next j
    ReDim arrItems(0 To OtherBox.ColumnCount - 1)
    For j = 0 To OtherBox.ListCount - 1
        If OtherBox.Selected(j) Then

            For i = 0 To OtherBox.ColumnCount - 1
                arrItems(i) = OtherBox.column(i, j)
            Next i
            msg = msg & Join(arrItems, ",") & vbCrLf
        End If
    Next j
    
    With Workbooks
        If .Count = 0 Then
        .Add
        End If
    End With
      
    
    If msg = "" Then
    
    MsgBox "You must select at least one report"
    Exit Sub
    
    Else
      iRet = MsgBox("Are you sure you want to run these reports?" & vbCrLf & vbCrLf & msg, vbYesNo, "Execute EBW Portal repors")
            If iRet = vbNo Then
            Exit Sub
            Else
            End If
            
    End If
    'AgingBOX
    
    If InStr(msg, "Closing Backlog") > 0 Then
        ActiveWorkbook.FollowHyperlink Address:="https://dbcip50.al.sw.ericsson.se:8310/sap/bw/BEx?sap-language=EN&bsplanguage=EN&cmd=ldoc&INFOCUBE=FIN_MPA1&QUERY=P_STD_FIN_MPA1_12210", NewWindow:=True
    End If
    If InStr(msg, "Work in Progress (WIP)") > 0 Then
        ActiveWorkbook.FollowHyperlink Address:="https://dbcip50.al.sw.ericsson.se:8310/sap/bw/BEx?sap-language=EN&bsplanguage=EN&cmd=ldoc&INFOCUBE=FIN_MSTDR&QUERY=P_STD_FIN_MSTDR_12241", NewWindow:=True
    End If
    If InStr(msg, "Reserve Unrealized Costs (RUC)") > 0 Then
        ActiveWorkbook.FollowHyperlink Address:="https://dbcip50.al.sw.ericsson.se:8310/sap/bw/BEx?sap-language=EN&bsplanguage=EN&cmd=ldoc&INFOCUBE=FIN_MSTDR&QUERY=P_STD_FIN_MSTDR_12240", NewWindow:=True
    End If
    If InStr(msg, "Unbilled Sales (UBS)") > 0 Then
        ActiveWorkbook.FollowHyperlink Address:="https://dbcip48-new.al.sw.ericsson.se:53801/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fcom.sap.pct!2fplatform_add_ons!2fcom.sap.ip.bi!2fiViews!2fcom.sap.ip.bi.bex?BOOKMARK=90L58IGAXEA9P7BK1CY13PI2K", NewWindow:=True
    End If
    If InStr(msg, "Deferred Revenue") > 0 Then
        ActiveWorkbook.FollowHyperlink Address:="https://dbcip48-new.al.sw.ericsson.se:53801/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fcom.sap.pct!2fplatform_add_ons!2fcom.sap.ip.bi!2fiViews!2fcom.sap.ip.bi.bex?BOOKMARK=90L58IGAXEA7Z08GSW9DL0AHX", NewWindow:=True
    End If
    
    'PopularBOX
    
    If InStr(msg, "Copa Details & Analysis (OSF4)") > 0 Then
        ActiveWorkbook.FollowHyperlink Address:="https://dbcip48-new.al.sw.ericsson.se:53801/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fcom.sap.pct!2fplatform_add_ons!2fcom.sap.ip.bi!2fiViews!2fcom.sap.ip.bi.bex?BOOKMARK=90L58IGAXEA7YXWAYX6E5AO5N", NewWindow:=True
    End If
    If InStr(msg, "Copa Details & Analysis (OSF1)") > 0 Then
        ActiveWorkbook.FollowHyperlink Address:="https://dbcip48-new.al.sw.ericsson.se:53801/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fcom.sap.pct!2fplatform_add_ons!2fcom.sap.ip.bi!2fiViews!2fcom.sap.ip.bi.bex?BOOKMARK=90L58IGAXEA7YXWAYX6E5AO5N", NewWindow:=True
    End If
    If InStr(msg, "Copa FI") > 0 Then
        ActiveWorkbook.FollowHyperlink Address:="https://dbcip50.al.sw.ericsson.se:8310/sap/bw/BEx?sap-language=EN&bsplanguage=EN&cmd=ldoc&INFOCUBE=FIN_MSTD4&QUERY=P_STD_FIN_MSTD4_12208", NewWindow:=True
    End If
    If InStr(msg, "Assignment Status Card") > 0 Then
        ActiveWorkbook.FollowHyperlink Address:="https://dbcip48-new.al.sw.ericsson.se:53801/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fcom.sap.pct!2fplatform_add_ons!2fcom.sap.ip.bi!2fiViews!2fcom.sap.ip.bi.bex?BOOKMARK=90L58IGAXEA9P7BKBKQMNXJ9N", NewWindow:=True
    End If
    If InStr(msg, "SO & WBS Analysis") > 0 Then
        ActiveWorkbook.FollowHyperlink Address:="https://dbcip50.al.sw.ericsson.se:8310/sap/bw/BEx?sap-language=EN&bsplanguage=EN&cmd=ldoc&INFOCUBE=FIN_MSTD2&QUERY=P_STD_FIN_MSTDR_12104", NewWindow:=True
    End If
    If InStr(msg, "Proof of Delivery Check") > 0 Then
        ActiveWorkbook.FollowHyperlink Address:="https://dbcip48-new.al.sw.ericsson.se:53801/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fcom.sap.pct!2fplatform_add_ons!2fcom.sap.ip.bi!2fiViews!2fcom.sap.ip.bi.bex?BOOKMARK=90L58IGAXEA9P7BKF2BYVF8B4", NewWindow:=True
    End If
    
    'CleanUpBOX
    
    If InStr(msg, "Overdue Not Closed") > 0 Then
        ActiveWorkbook.FollowHyperlink Address:="https://dbcip48-new.al.sw.ericsson.se:53801/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fcom.sap.pct!2fplatform_add_ons!2fcom.sap.ip.bi!2fiViews!2fcom.sap.ip.bi.bex?BOOKMARK=90L58IGAXEA9P7BLP6NUGS7WW", NewWindow:=True
    End If
    If InStr(msg, "Fully Invoice not Closed") > 0 Then
        ActiveWorkbook.FollowHyperlink Address:="https://dbcip48-new.al.sw.ericsson.se:53801/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fcom.sap.pct!2fplatform_add_ons!2fcom.sap.ip.bi!2fiViews!2fcom.sap.ip.bi.bex?BOOKMARK=90L58IGAXEA9P7BLN056IFX1O", NewWindow:=True
    End If
    
    'DeploymentBox
    
    If InStr(msg, "Value Contract") > 0 Then
        ActiveWorkbook.FollowHyperlink Address:="https://dbcip48-new.al.sw.ericsson.se:53801/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fcom.sap.pct!2fplatform_add_ons!2fcom.sap.ip.bi!2fiViews!2fcom.sap.ip.bi.bex?BOOKMARK=90L58IGAXEA9P7BL8NH1JG0S8", NewWindow:=True
    End If
    If InStr(msg, "CCLM ID") > 0 Then
        ActiveWorkbook.FollowHyperlink Address:="https://dbcip48-new.al.sw.ericsson.se:53801/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fcom.sap.pct!2fplatform_add_ons!2fcom.sap.ip.bi!2fiViews!2fcom.sap.ip.bi.bex?BOOKMARK=90L58IGAXEA9P7BLAJ3X2R01N", NewWindow:=True
    End If
    If InStr(msg, "Assignment ID") > 0 Then
        ActiveWorkbook.FollowHyperlink Address:="https://dbcip48-new.al.sw.ericsson.se:53801/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fcom.sap.pct!2fplatform_add_ons!2fcom.sap.ip.bi!2fiViews!2fcom.sap.ip.bi.bex?BOOKMARK=90L58IGAXEA9P7BL9DVJ3O54P", NewWindow:=True
    End If

    'KPIBOX
    
    If InStr(msg, "Enhanced Plan Cost Quality") > 0 Then
        ActiveWorkbook.FollowHyperlink Address:="https://dbcip48-new.al.sw.ericsson.se:53801/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fcom.sap.pct!2fplatform_add_ons!2fcom.sap.ip.bi!2fiViews!2fcom.sap.ip.bi.bex?BOOKMARK=90L58IGAXEA9P7BNX9193PROQ", NewWindow:=True
    End If
    If InStr(msg, "RUC vs COS") > 0 Then
        ActiveWorkbook.FollowHyperlink Address:="https://dbcip48-new.al.sw.ericsson.se:53801/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fcom.sap.pct!2fplatform_add_ons!2fcom.sap.ip.bi!2fiViews!2fcom.sap.ip.bi.bex?BOOKMARK=90L58IGAXEA9P7BO0EXBBZV58", NewWindow:=True
    End If
    If InStr(msg, "Billing Forecast") > 0 Then
        ActiveWorkbook.FollowHyperlink Address:="https://dbcip48-new.al.sw.ericsson.se:53801/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fcom.sap.pct!2fplatform_add_ons!2fcom.sap.ip.bi!2fiViews!2fcom.sap.ip.bi.bex?BOOKMARK=90L58IGAXEA9P7BKHLEFSSQTV", NewWindow:=True
    End If
    If InStr(msg, "Plan Cost vs Budget") > 0 Then
        ActiveWorkbook.FollowHyperlink Address:="https://dbcip48-new.al.sw.ericsson.se:53801/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fcom.sap.pct!2fplatform_add_ons!2fcom.sap.ip.bi!2fiViews!2fcom.sap.ip.bi.bex?BOOKMARK=90L58IGAXEA9P7BL5A6PSP0J6", NewWindow:=True
    End If
    If InStr(msg, "Overdue not Closed") > 0 Then
        ActiveWorkbook.FollowHyperlink Address:="https://dbcip48-new.al.sw.ericsson.se:53801/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fcom.sap.pct!2fplatform_add_ons!2fcom.sap.ip.bi!2fiViews!2fcom.sap.ip.bi.bex?BOOKMARK=90L58IGAXEA9P7BLP6NUGS7WW", NewWindow:=True
    End If

    'OtherBOX
    
    If InStr(msg, "Project Follow-Up USD") > 0 Then
        ActiveWorkbook.FollowHyperlink Address:="https://dbcip50.al.sw.ericsson.se:8302/sap/bw/BEx?SAP-LANGUAGE=EN&BOOKMARK_ID=90L58IGR0WBJY1BVTNW69O3JM", NewWindow:=True
    End If
    If InStr(msg, "CPL Actuals") > 0 Then
        ActiveWorkbook.FollowHyperlink Address:="https://dbcip48-new.al.sw.ericsson.se:53801/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fcom.sap.pct!2fplatform_add_ons!2fcom.sap.ip.bi!2fiViews!2fcom.sap.ip.bi.bex?BOOKMARK=90L58IGAXEA7ZDSXBYRAO04LT", NewWindow:=True
    End If
    If InStr(msg, "ICRRB Report") > 0 Then
        ActiveWorkbook.FollowHyperlink Address:="https://dbcip48-new.al.sw.ericsson.se:53801/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fcom.sap.pct!2fplatform_add_ons!2fcom.sap.ip.bi!2fiViews!2fcom.sap.ip.bi.bex?BOOKMARK=90L58IGAXEA9P7BL2DXCY2CQ7", NewWindow:=True
    End If
    If InStr(msg, "Detail Cost & Hours Report") > 0 Then
        ActiveWorkbook.FollowHyperlink Address:="https://dbcip48-new.al.sw.ericsson.se:53801/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fcom.sap.pct!2fplatform_add_ons!2fcom.sap.ip.bi!2fiViews!2fcom.sap.ip.bi.bex?BOOKMARK=90L58IGAXEA9P7BKPTVZ8R9WE", NewWindow:=True
    End If
'    If InStr(msg, "Loss Making Projects") > 0 Then
'        ActiveWorkbook.FollowHyperlink Address:="", NewWindow:=True
'    End If
    
    'CLEAR SELECTION
    
    Me.PopularBox.MultiSelect = fmMultiSelectSingle
    Me.PopularBox.value = ""
    Me.PopularBox.MultiSelect = fmMultiSelectMulti
    Me.CleanUpBox.MultiSelect = fmMultiSelectSingle
    Me.CleanUpBox.value = ""
    Me.CleanUpBox.MultiSelect = fmMultiSelectMulti
    Me.AgingBox.MultiSelect = fmMultiSelectSingle
    Me.AgingBox.value = ""
    Me.AgingBox.MultiSelect = fmMultiSelectMulti
    Me.DeploymentBox.MultiSelect = fmMultiSelectSingle
    Me.DeploymentBox.value = ""
    Me.DeploymentBox.MultiSelect = fmMultiSelectMulti
    Me.KPIBox.MultiSelect = fmMultiSelectSingle
    Me.KPIBox.value = ""
    Me.KPIBox.MultiSelect = fmMultiSelectMulti
    Me.OtherBox.MultiSelect = fmMultiSelectSingle
    Me.OtherBox.value = ""
    Me.OtherBox.MultiSelect = fmMultiSelectMulti
    EBWfrm.Hide
    Exit Sub
Errorhandler:
    
    MsgBox "Seems the shortcut is incorrect, dont select multiple reports at once, try running one at a time"
    Exit Sub
    
End Sub

Private Sub UserForm_Initialize()
With PopularBox
     .AddItem "Copa Details & Analysis (OSF4)"
     .AddItem "Copa Details & Analysis (OSF1)"
     .AddItem "Copa FI"
     .AddItem "Assignment Status Card"
     .AddItem "SO & WBS Analysis"
     .AddItem "Proof of Delivery Check"
End With
With CleanUpBox
     .AddItem "Overdue Not Closed"
     .AddItem "Fully Invoice not Closed"
End With

With AgingBox
    .AddItem "Closing Backlog"
    .AddItem "Work in Progress (WIP)"
    .AddItem "Reserve Unrealized Costs (RUC)"
    .AddItem "Unbilled Sales (UBS)"
    .AddItem "Deferred Revenue"
End With
With DeploymentBox
    .AddItem "Value Contract"
    .AddItem "CCLM ID"
    .AddItem "Assignment ID"
End With

With KPIBox
    .AddItem "Enhanced Plan Cost Quality"
    .AddItem "RUC vs COS"
    .AddItem "Billing Forecast"
    .AddItem "Plan Cost vs Budget"
    .AddItem "Overdue not Closed"
End With

With OtherBox
    .AddItem "Project Follow-Up USD"
    .AddItem "CPL Actuals"
    .AddItem "ICRRB Report"
    .AddItem "Detail Cost & Hours Report"
    .AddItem "Loss Making Projects"
End With
End Sub

