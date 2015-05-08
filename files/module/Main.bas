Attribute VB_Name = "Main"
Option Explicit
Dim CLPuserid, CLPpasswd As String
Dim StartTime, StopTime, ElapsedTime, ElapsedTimePerObject, ElapsedTimeRemainderSeconds, ElapsedTimeMinutes, TimeAtLastProcessedObject, AvgTimeToNow, preTimeEstimate As Double
Dim n, sendMailResponse, nleft As Integer
Private trx As String
Dim Question As String
Dim iRet As Integer
Dim shtname As String
Dim Wsave As Integer
' Mail functionality
Dim CMail As CCMail
Dim CStatMail As CCMail
Dim Wsize As Integer

Sub MainScript()
    Dim SAPConnection As CSAPConnection
    Set SAPConnection = New CSAPConnection
    Set CMail = New CCMail
    CMail.Init (CErrorReport)
    Application.DisplayFullScreen = True
    With Workbooks
        If .Count = 0 Then
        .Add
        End If
    End With

    
    If Workbooks.Count > 1 Then
        Question = MsgBox("CAUTION, You have " & Workbooks.Count - 1 & " additional WORKBOOK OPEN, it's recommended to close them before continuing, do you wish to ignore this and CONTINUE to run the SCRIPT?", vbYesNo, "CAUTION!")
        If Question = vbNo Then
    Application.DisplayFullScreen = False
            End
        End If
    End If
    
    Dim chosenScript As String
    chosenScript = Sheets(1).Cells(3, 2).value
            
            If chosenScript = "Create_QTCM_Sales_Order" Then
                CreateQTCMfrm.Show
            End If
            
            If chosenScript = "Revenue recognition (QTC)" Then
            
            iRet = MsgBox("Are you sure you want to process this QTC revenue recognition?", vbYesNo, "Revenue Recognition request")
            If iRet = vbNo Then
    Application.DisplayFullScreen = False
            End
            Else
            End If
        End If

    Application.DisplayAlerts = False
    SAPConnection.absorbConnection
    
    chosenScript = Sheets(1).Cells(3, 2).value
    
shtname = ActiveSheet.Name
            If chosenScript = "Update_WBS_System_Status" Then
                GoTo Continuesub
            ElseIf chosenScript = "Airtel_Rebate_List" Then
                AirtelRebateListScript trx, SAPConnection, CMail
                GoTo Finish_Sub:

            ElseIf chosenScript = "Update_Sales_Order_System_Status" Then
                GoTo Continuesub
            ElseIf chosenScript = "Update_Value_Contract_System_Status" Then
                GoTo Continuesub
            ElseIf chosenScript = "Planned_Cost_Update" Then
                GoTo Continuesub
            ElseIf chosenScript = "POC_Milestone_Creation_Update" Then
                GoTo Continuesub
            ElseIf chosenScript = "Revenue recognition (QTC)" Then
                GoTo Continuesub
            
            ElseIf chosenScript = "Update_Sales_Order_Revenue_Status" Then
                GoTo Continuesub
            ElseIf chosenScript = "Update_Project_Finish_Date" Then
                GoTo Continuesub
            ElseIf chosenScript = "ENO_Planned_Cost_Update_Marzenna" Then
                GoTo Continuesub
            ElseIf chosenScript = "Update_BillingType" Then
      iRet = MsgBox("This Script is currently in BETA, do you wish to continue?", vbYesNo, "BETA Script (UNDER DEVELOPMENT)")
            If iRet = vbNo Then
            End
            Else
                GoTo Continuesub
                End If
            ElseIf chosenScript = "Create_QTCM_Sales_Order" Then
                
                GoTo Continuesub
            ElseIf chosenScript = "Update_Value_Contract_Description" Then
                GoTo Continuesub
            ElseIf chosenScript = "Run_Settlement_QTC" Then
                GoTo Continuesub
            ElseIf chosenScript = "Run_Settlement_PSF" Then
      iRet = MsgBox("This Script is currently in BETA, do you wish to continue?", vbYesNo, "BETA Script (UNDER DEVELOPMENT)")
            If iRet = vbNo Then
            End
            Else
                GoTo Continuesub
            End If
            ElseIf chosenScript = "Rebate_Percentage_Update" Then
                GoTo Continuesub
            ElseIf chosenScript = "Rebate_Description_Update" Then
                GoTo Continuesub
            ElseIf chosenScript = "SO_Rebate_Condition_Update" Then
                GoTo Continuesub
            ElseIf chosenScript = "Update_Partner_Value_Contract" Then
                GoTo Continuesub
            ElseIf chosenScript = "Update_Gstream_AssignmentID" Then
                GoTo Continuesub
            ElseIf chosenScript = "UpdateWBSCC" Then
                GoTo Continuesub
            ElseIf chosenScript = "Create_Value_Contract" Then
      iRet = MsgBox("This Script is currently in BETA, do you wish to continue?", vbYesNo, "BETA Script (UNDER DEVELOPMENT)")
            If iRet = vbNo Then
            End
            Else
                GoTo Continuesub
            End If
            ElseIf chosenScript = "Create_Rebate" Then
                GoTo Continuesub
            Else
            GoTo RunTemplate
            End If
                
Continuesub:

    
    On Error GoTo Finish_Sub
    Worksheets(shtname).Select
        
    Select Case chosenScript
        Case "Update_WBS_System_Status"
            preTimeEstimate = 6.484003522
        Case "Planned_Cost_Update"
            preTimeEstimate = 25.61350692
        Case "POC_Milestone_Creation_Update"
            preTimeEstimate = 9.190022676
        Case "Update_Project_Finish_Date"
            preTimeEstimate = 8.326732673
        Case "Update_Sales_Order_System_Status"
            preTimeEstimate = 16.14139009
        Case "Update_Value_Contract_System_Status"
            preTimeEstimate = 6.132616487
        Case Else
            preTimeEstimate = 0
    End Select
    
    If preTimeEstimate > 0 Then
        Sheets(shtname).Cells(6, 2) = "Estimated completion in " & Round(((Range(Cells(9, 3), Cells(9, 3).End(xlDown)).Count) * preTimeEstimate) / 60, 0) & " minutes."
    End If
    
    Application.ScreenUpdating = False
    
    Range("A9").Select
    ActiveWindow.FreezePanes = True
    Rows("1:5").Select
    Selection.EntireRow.Hidden = True
    Application.ScreenUpdating = True
    
    StartTime = Now()
    Application.StatusBar = "Running, calculating remaining time..."
    SAPConnection.session.TestToolMode = 1
    ' Hit enter on information windows upon login
    If Not SAPConnection.session.findById("wnd[1]", False) Is Nothing Then
        Do
            SAPConnection.session.findById("wnd[1]", False).sendVKey 0
        Loop While Not SAPConnection.session.findById("wnd[1]", False) Is Nothing
    End If
            If chosenScript = "Update_WBS_System_Status" Then
                trx = "CJ20N"
            ElseIf chosenScript = "Update_Sales_Order_System_Status" Then
                trx = "VA02"
            ElseIf chosenScript = "Update_Value_Contract_System_Status" Then
                trx = "VA42"
            ElseIf chosenScript = "Update_Partner_Value_Contract" Then
                trx = "VA42"
            ElseIf chosenScript = "Update_Gstream_AssignmentID" Then
                trx = "VA42"
            ElseIf chosenScript = "UpdateWBSCC" Then
                trx = "CJ20N"
            ElseIf chosenScript = "Planned_Cost_Update" Then
                trx = "CJ20N"
            ElseIf chosenScript = "POC_Milestone_Creation_Update" Then
                trx = "CJ20N"
            ElseIf chosenScript = "Revenue recognition (QTC)" Then
                trx = "CJ20N"
            ElseIf chosenScript = "Update_Sales_Order_Revenue_Status" Then
                trx = "VA02"
            ElseIf chosenScript = "Update_Project_Finish_Date" Then
                trx = "CJ20N"
            ElseIf chosenScript = "ENO_Planned_Cost_Update_Marzenna" Then
                trx = "CJ20N"
            ElseIf chosenScript = "Create_QTCM_Sales_Order" Then
                trx = "VA01"
            ElseIf chosenScript = "Update_Value_Contract_Description" Then
                trx = "VA42"
            ElseIf chosenScript = "Update_BillingType" Then
                trx = "VA02"
            ElseIf chosenScript = "Run_Settlement_QTC" Then
                trx = "CJA2"
            ElseIf chosenScript = "Run_Settlement_PSF" Then
                trx = "KKA3"
            ElseIf chosenScript = "Airtel_Rebate_List" Then
                trx = "VB(8"
            ElseIf chosenScript = "Rebate_Description_Update" Then
                trx = "VBO2"
                Wsize = True
            ElseIf chosenScript = "Rebate_Percentage_Update" Then
                trx = "VBO2"
                Wsize = True
            ElseIf chosenScript = "SO_Rebate_Condition_Update" Then
                trx = "VA05"
            ElseIf chosenScript = "Create_Value_Contract" Then
                trx = "VA41"
            ElseIf chosenScript = "Create_Rebate" Then
                trx = "VBO1"
                Wsize = True
               End If
    SAPConnection.session.findById("wnd[0]").Maximize
    SAPConnection.session.findById("wnd[0]/tbar[0]/okcd").Text = "/N" & trx
    SAPConnection.session.findById("wnd[0]").sendVKey 0
    
    Cells(9, 2).Select
    
    Do Until IsEmpty(ActiveCell.Offset(0, 1))
    
    'if the object has already been processed then dont Save workbook
    If ActiveCell.value = 1 Then
    Wsave = False
    Else
    Wsave = True
    End If
    
        If (ActiveCell.value <> 1) Then
        ' call void modules with arguments
            If chosenScript = "Update_WBS_System_Status" Then
                UpdateStatusWBSScript ActiveCell.Offset(0, 1), ActiveCell.Offset(0, 2), trx, SAPConnection, CMail
            ElseIf chosenScript = "Update_Sales_Order_System_Status" Then
                UpdateStatusSOScript ActiveCell.Offset(0, 1), ActiveCell.Offset(0, 2), trx, SAPConnection, CMail
            ElseIf chosenScript = "UpdateWBSCC" Then
                UpdateWBSCCScript ActiveCell.Offset(0, 1), ActiveCell.Offset(0, 2), trx, SAPConnection, CMail
            ElseIf chosenScript = "Update_Value_Contract_System_Status" Then
                UpdateStatusVCScript ActiveCell.Offset(0, 1), ActiveCell.Offset(0, 2), trx, SAPConnection, CMail
            ElseIf chosenScript = "Update_Partner_Value_Contract" Then
                UpdatePartnerVCScript ActiveCell.Offset(0, 1), ActiveCell.Offset(0, 2), ActiveCell.Offset(0, 3), trx, SAPConnection, CMail
            ElseIf chosenScript = "Update_Gstream_AssignmentID" Then
                Update_Gstream_AssignmentIDScript ActiveCell.Offset(0, 1), ActiveCell.Offset(0, 2), ActiveCell.Offset(0, 3), trx, SAPConnection, CMail
            ElseIf chosenScript = "Planned_Cost_Update" Then
                Planned_Cost_UpdateScript ActiveCell.Offset(0, 1), trx, SAPConnection, CMail
            ElseIf chosenScript = "POC_Milestone_Creation_Update" Then
                POC_Milestone_Creation_UpdateScript ActiveCell.Offset(0, 1), ActiveCell.Offset(0, 2), ActiveCell.Offset(0, 3), ActiveCell.Offset(0, 4), ActiveCell.Offset(0, 5), trx, SAPConnection, CMail
            ElseIf chosenScript = "Revenue recognition (QTC)" Then
                RG_POC_Milestone_CreationScript ActiveCell.Offset(0, 1), ActiveCell.Offset(0, 5), ActiveCell.Offset(0, 6), ActiveCell.Offset(0, 8), trx, SAPConnection, CMail
            ElseIf chosenScript = "Update_Sales_Order_Revenue_Status" Then
                UpdateSalesOrderRevenueStatusScript ActiveCell.Offset(0, 1), ActiveCell.Offset(0, 2), ActiveCell.Offset(0, 3), trx, SAPConnection, CMail
            ElseIf chosenScript = "Update_Project_Finish_Date" Then
                UpdateProjectFinishDateScript ActiveCell.Offset(0, 1), ActiveCell.Offset(0, 2), trx, SAPConnection, CMail
           ElseIf chosenScript = "ENO_Planned_Cost_Update_Marzenna" Then
                ENOPlannedCostUpdateMarzennaScript ActiveCell.Offset(0, 1), trx, SAPConnection, CMail
            ElseIf chosenScript = "Create_QTCM_Sales_Order" Then
                CreateQTCMSalesOrderScript ActiveCell.Offset(0, 2), ActiveCell.Offset(0, 3), ActiveCell.Offset(0, 4), ActiveCell.Offset(0, 5), ActiveCell.Offset(0, 6), trx, SAPConnection, CMail
            ElseIf chosenScript = "Update_Value_Contract_Description" Then
                UpdateVCDescriptionScript ActiveCell.Offset(0, 1), ActiveCell.Offset(0, 2), trx, SAPConnection, CMail
            ElseIf chosenScript = "Run_Settlement_QTC" Then
                RunSettlementWBSScript ActiveCell.Offset(0, 1), ActiveCell.Offset(0, 2), ActiveCell.Offset(0, 3), trx, SAPConnection, CMail
            ElseIf chosenScript = "Run_Settlement_PSF" Then
                RunSettlementWBSScript ActiveCell.Offset(0, 1), ActiveCell.Offset(0, 2), ActiveCell.Offset(0, 3), trx, SAPConnection, CMail
            ElseIf chosenScript = "Update_BillingType" Then
                UpdateBillingTypeScript ActiveCell.Offset(0, 1), ActiveCell.Offset(0, 2), trx, SAPConnection, CMail
            ElseIf chosenScript = "Rebate_Percentage_Update" Then
                RebateUpdateScript ActiveCell.Offset(0, 1), ActiveCell.Offset(0, 2), trx, SAPConnection, CMail
            ElseIf chosenScript = "SO_Rebate_Condition_Update" Then
                SO_Rebate_Condition_UpdateScript ActiveCell.Offset(0, 1), ActiveCell.Offset(0, 2), ActiveCell.Offset(0, 3), trx, SAPConnection, CMail
            ElseIf chosenScript = "Rebate_Description_Update" Then
                Rebate_Description_UpdateScript ActiveCell.Offset(0, 1), ActiveCell.Offset(0, 2), trx, SAPConnection, CMail
            ElseIf chosenScript = "Create_Value_Contract" Then
                Create_Value_ContractScript ActiveCell.Offset(0, 1), ActiveCell.Offset(0, 2), ActiveCell.Offset(0, 3), ActiveCell.Offset(0, 4), ActiveCell.Offset(0, 5), ActiveCell.Offset(0, 6), ActiveCell.Offset(0, 7), ActiveCell.Offset(0, 8), ActiveCell.Offset(0, 9), ActiveCell.Offset(0, 10), ActiveCell.Offset(0, 11), ActiveCell.Offset(0, 12), ActiveCell.Offset(0, 13), ActiveCell.Offset(0, 14), ActiveCell.Offset(0, 15), ActiveCell.Offset(0, 16), ActiveCell.Offset(0, 17), ActiveCell.Offset(0, 18), ActiveCell.Offset(0, 19), ActiveCell.Offset(0, 20), ActiveCell.Offset(0, 21), ActiveCell.Offset(0, 22), ActiveCell.Offset(0, 23), ActiveCell.Offset(0, 24), ActiveCell.Offset(0, 25), ActiveCell.Offset(0, 26), ActiveCell.Offset(0, 27), ActiveCell.Offset(0, 28), trx, SAPConnection, CMail
            ElseIf chosenScript = "Create_Rebate" Then
                Create_RebateScript ActiveCell.Offset(0, 1), ActiveCell.Offset(0, 2), ActiveCell.Offset(0, 3), ActiveCell.Offset(0, 4), ActiveCell.Offset(0, 5), ActiveCell.Offset(0, 6), ActiveCell.Offset(0, 7), ActiveCell.Offset(0, 8), ActiveCell.Offset(0, 9), trx, SAPConnection, CMail
            Else
                GoTo RunTemplate
            End If
            
            n = n + 1
        End If
        TimeAtLastProcessedObject = Now() ' time taken until this 'n'
        If Not n = 0 Then ' if no objects processed, do not attempt division by zero
            AvgTimeToNow = (TimeAtLastProcessedObject - StartTime) / n
        End If
        
        If IsEmpty(ActiveCell.Offset(1, 1)) Then ' if this is the last object (no left)
            nleft = 0
        Else
            nleft = Range(ActiveCell.Offset(0, 1), ActiveCell.Offset(0, 1).End(xlDown)).Count - 1 ' else, count remaining
        End If
        
        If Not nleft = 0 Then ' if no objects left to update, skip updating statusbar
            Application.StatusBar = "Running, estimated completion in " & Round(((AvgTimeToNow * 1400) * nleft), 0) & " minutes and " & Round(((AvgTimeToNow * 86400) * nleft) Mod 60, 0) & " seconds. (" & nleft & " objects left.)"
        End If
        
        If n = 0 And nleft = 0 Then
        Progressfrm.Text.Caption = "100% Processed"
        Progressfrm.Bar.Width = 306
        Progressfrm.status.Caption = "No records updated"
        Else
        Progressfrm.Text.Caption = Round(n / (n + nleft) * 100, 0) & "% Processed"
        Progressfrm.Bar.Width = n / (n + nleft) * 306
        Progressfrm.status.Caption = "Estimated completion in " & Round(((AvgTimeToNow * 1400) * nleft), 0) & " minutes and " & Round(((AvgTimeToNow * 86400) * nleft) Mod 60, 0) & " seconds. (" & nleft & " objects left.)"
        End If
        
        If Wsave = True Then
             If Not ActiveWorkbook.ReadOnly Then
                    ActiveWorkbook.Save
             End If
            SAPConnection.session.findById("wnd[0]/tbar[0]/okcd").Text = "/N" & trx
            SAPConnection.session.findById("wnd[0]").sendVKey 0
        End If
                
        ActiveCell.Offset(1, 0).Select
        
    
        If Wsize = False Then
            SAPConnection.session.findById("wnd[0]").Maximize
        End If
    Wsave = True
    Loop
      
    SAPConnection.session.TestToolMode = 0
    
    StopTime = Now()
    ElapsedTime = (StopTime - StartTime) * 86400
    If Not n = 0 Then
        ElapsedTimePerObject = ElapsedTime / n
    End If
        
    If ElapsedTime <= 60 Then
        Sheets(shtname).Cells(6, 2) = Format(Now(), "dd-mmm-yyyy | hh:mm AM/PM") & " -The work task was finished in " & Round(ElapsedTime, 0) & " second(s), and processed " & n & " object(s). Which is about " & Round(ElapsedTimePerObject, 0) & " seconds per object."
    ElseIf ElapsedTime > 60 Then
        ElapsedTimeRemainderSeconds = ElapsedTime Mod 60
        ElapsedTimeMinutes = (ElapsedTime - ElapsedTimeMinutes) / 60
        Sheets(shtname).Cells(6, 2) = Format(Now(), "dd-mmm-yyyy | hh:mm AM/PM") & " -The work task was finished in " & Round(ElapsedTimeMinutes, 0) & " minute(s) and " & Round(ElapsedTimeRemainderSeconds, 0) & " second(s), and processed " & n & " object(s). Which is about " & Round(ElapsedTimePerObject, 0) & " second(s) per object."
    End If
    
          'Update Format
        Application.ScreenUpdating = False
  If ActiveSheet.ProtectContents = True Then
  ActiveSheet.Unprotect
  End If
        Range("B6").Select
    With ActiveCell.Characters(Start:=1, Length:=0).Font
        .Name = "Calibri"
        .FontStyle = "Bold"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With ActiveCell.Characters(Start:=1, Length:=22).Font
        .Name = "Calibri"
        .FontStyle = "Italic"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .Color = -16776961
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With ActiveCell.Characters(Start:=23, Length:=109).Font
        .Name = "Calibri"
        .FontStyle = "Bold"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Range("B7").Select
    Columns("C:C").EntireColumn.AutoFit
    Columns("D:D").EntireColumn.AutoFit
    Columns("E:E").EntireColumn.AutoFit
    Columns("F:F").EntireColumn.AutoFit
    Columns("G:G").EntireColumn.AutoFit
    Columns("H:H").EntireColumn.AutoFit
    Columns("I:I").EntireColumn.AutoFit
    Columns("J:J").EntireColumn.AutoFit
    Columns("K:K").EntireColumn.AutoFit
    Columns("L:L").EntireColumn.AutoFit
    Columns("M:M").EntireColumn.AutoFit
    Columns("N:N").EntireColumn.AutoFit
    Columns("O:O").EntireColumn.AutoFit
    Columns("P:P").EntireColumn.AutoFit
    Columns("Q:Q").EntireColumn.AutoFit
    Columns("R:R").EntireColumn.AutoFit
    
    
    Range("A9").Select
    ActiveWindow.FreezePanes = False
    Rows("1:5").Select
    Selection.EntireRow.Hidden = False
    Rows("8:8").Select
    Selection.AutoFilter
    Rows("9:9").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Locked = False
    Selection.FormulaHidden = False
    Range("A9").Select
    
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingColumns:=True, AllowFormattingRows:=True, _
        AllowInsertingRows:=True, AllowInsertingHyperlinks:=True, _
        AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
        AllowUsingPivotTables:=True

    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    If (CMail.CheckIfErrorListExists) Then
        CMail.SendMail
    End If
    
    Set CMail = Nothing
    Application.Wait (Now + TimeValue("00:00:03"))
    Progressfrm.Hide
    
GoTo Finish_Sub:
RunTemplate:

                MsgBox "Error, current workbook is NOT an authorized Script Form! Use the Create Script Form button.", vbInformation, "Error"
                    Application.DisplayFullScreen = False
    
                    End
    
Finish_Sub:
    If Err = 619 Then
        Resume Next
        MsgBox "Unknown error, try rebooting the macro file", , "Error"
        If sendMailResponse = vbYes Then
            SAPConnection.reportCrash ActiveCell.Offset(0, 1)
        End If
    End If
         
    ' Nollställa variabler start
    CLPuserid = 0
    CLPpasswd = 0
    If n > 0 Then
        SAPConnection.logStatitics chosenScript, n, Now(), (ElapsedTime + ElapsedTimeMinutes)
    End If
    
    If GetSAPOption <> "Yes" Then
    SAPConnection.session.findById("wnd[0]").Close
    End If

    SAPConnection.killConnection
    Set SAPConnection = Nothing
    
    StartTime = 0
    StopTime = 0
    ElapsedTime = 0
    ElapsedTimePerObject = 0
    ElapsedTimeRemainderSeconds = 0
    ElapsedTimeMinutes = 0
    nleft = 0
    
    n = 0
    ' Nollställa variabler slut
    Application.DisplayAlerts = True
    
    If Workbooks.Count > 0 Then
             If Not ActiveWorkbook.ReadOnly Then
                    ActiveWorkbook.Save
             End If
    End If
    
    Application.StatusBar = False
    Application.DisplayFullScreen = False
    
    
    If GetTurnOffOption = "Yes" Then
        Shell "shutdown -s -t 120"
    End If
    
     End
End Sub
