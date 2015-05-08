Attribute VB_Name = "Template"
Option Explicit
Sub Create_Template(ByRef scriptname As String)
'
' Macro6 Macro
'

'
    Application.ScreenUpdating = False

        'Enter code for Header
            If scriptname = "Update_WBS_System_Status" Then
                Call Update_WBS_System_StatusTemplate

            ElseIf scriptname = "Update_Sales_Order_System_Status" Then
                Call Update_Sales_Order_System_StatusTemplate

            ElseIf scriptname = "Update_Value_Contract_System_Status" Then
                Call Update_Value_Contract_System_StatusTemplate

            ElseIf scriptname = "Planned_Cost_Update" Then
                Call Planned_Cost_UpdateTemplate

            ElseIf scriptname = "POC_Milestone_Creation_Update" Then
                Call POC_Milestone_Creation_UpdateTemplate

            ElseIf scriptname = "Update_Sales_Order_Revenue_Status" Then
                Call Update_Sales_Order_Revenue_StatusTemplate

            ElseIf scriptname = "Update_Project_Finish_Date" Then
                Call Update_Project_Finish_DateTemplate

            ElseIf scriptname = "ENO_Planned_Cost_Update_Marzenna" Then
                Call ENO_Planned_Cost_Update_MarzennaTemplate

            ElseIf scriptname = "Create_QTCM_Sales_Order" Then
                Call Create_QTCM_Sales_OrderTemplate

            ElseIf scriptname = "Update_BillingType" Then
                Call Update_BillingTypeTemplate
            
            ElseIf scriptname = "Update_Value_Contract_Description" Then
                Call Update_Value_Contract_DescriptionTemplate

            ElseIf scriptname = "Run_Settlement_QTC" Then
                Call Run_Settlement_QTCTemplate

            ElseIf scriptname = "Run_Settlement_PSF" Then
                Call Run_Settlement_PSFTemplate
            
            ElseIf scriptname = "Update_Partner_Value_Contract" Then
                Call Update_Partner_VCTemplate
            
            ElseIf scriptname = "Update_Gstream_AssignmentID" Then
                Call Update_Gstream_AssignmentIDTemplate
            
            ElseIf scriptname = "Airtel_Rebate_List" Then
                Call Airtel_Rebate_ListTemplate
            
            ElseIf scriptname = "Rebate_Percentage_Update" Then
                Call Update_RebateTemplate
            
            ElseIf scriptname = "Rebate_Description_Update" Then
                Call Rebate_DescriptionTemplate
                
            ElseIf scriptname = "Create_Value_Contract" Then
                Call Create_Value_ContractTemplate
                
            ElseIf scriptname = "SO_Rebate_Condition_Update" Then
                Call Update_SO_Rebate_ConditionTemplate
                
            ElseIf scriptname = "Create_Rebate" Then
                Call Create_RebateTemplate
          
          Else
                MsgBox "Oh Oh, Looks like you forgot to select an option from the Script list dropdown, try again...", , "OOOPS!!!"
                End
                End If
    'Stop
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "Template for Script:"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = scriptname
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "Description"
    Range("B2:D2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("B3:D4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("B2:D2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("B3:D4").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13762046
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("F2:J2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("F3:J4").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("F2:J2").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("F3:J4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Range("B6:J6").Select
    Selection.Font.Bold = False
    Selection.Font.Bold = True
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("C11").Select
    Cells.FormatConditions.Delete
    Range("B9:B10000").Select
    Range("B9").Activate
    Selection.FormatConditions.AddIconSetCondition
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .ReverseOrder = False
        .ShowIconOnly = True
        .IconSet = ActiveWorkbook.IconSets(xl3Symbols2)
    End With
    Selection.FormatConditions(1).IconCriteria(1).Icon = xlIconRedCross
    With Selection.FormatConditions(1).IconCriteria(2)
        .Type = xlConditionValueNumber
        .value = 0.5
        .Operator = 5
        .Icon = xlIconRedCross
    End With
    With Selection.FormatConditions(1).IconCriteria(3)
        .Type = xlConditionValueNumber
        .value = 1
        .Operator = 7
        .Icon = xlIconGreenCheck
    End With
    Range("B9:R10000").Select
    Range("B9").Activate
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
        Formula1:="="""""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Borders(xlLeft)
        .LineStyle = xlDot
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.FormatConditions(1).Borders(xlRight)
        .LineStyle = xlDot
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.FormatConditions(1).Borders(xlTop)
        .LineStyle = xlDot
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.FormatConditions(1).Borders(xlBottom)
        .LineStyle = xlDot
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("B8:R8").Select
    Range("B8").Activate
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
        Formula1:="="""""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.FormatConditions(1).Borders(xlLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.FormatConditions(1).Borders(xlRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.FormatConditions(1).Borders(xlTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.FormatConditions(1).Borders(xlBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = 0
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    Range("C9:R10000").Select
    Range("C11").Activate
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = False
        .Italic = True
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.FormatConditions(1).Borders(xlLeft)
        .LineStyle = xlDot
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.FormatConditions(1).Borders(xlRight)
        .LineStyle = xlDot
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.FormatConditions(1).Borders(xlTop)
        .LineStyle = xlDot
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.FormatConditions(1).Borders(xlBottom)
        .LineStyle = xlDot
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.FormatConditions(1).Interior
        .Pattern = xlSolid
        .PatternColorIndex = 0
        .Color = 10418422
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("C9:R10000").Select
    Range("C11").Activate
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=ISNUMBER(SEARCH(""SAP""" & Excel.Application.International(xlListSeparator) & "C$8" & Excel.Application.International(xlListSeparator) & "1))"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = False
        .Italic = True
        .Color = -16776961
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Borders(xlLeft)
        .LineStyle = xlDot
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.FormatConditions(1).Borders(xlRight)
        .LineStyle = xlDot
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.FormatConditions(1).Borders(xlTop)
        .LineStyle = xlDot
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.FormatConditions(1).Borders(xlBottom)
        .LineStyle = xlDot
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.FormatConditions(1).Interior
        .Pattern = xlNone
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("C9:R10000").Select
    Range("C11").Activate
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).StopIfTrue = True
    ActiveWindow.LargeScroll Down:=-1
    Range("B8").Select
    ActiveCell.FormulaR1C1 = " "
    Range("B9").Select

'Second Macro
    
    ActiveWindow.Zoom = 85
    Columns("B:Q").Select
    Selection.ColumnWidth = 12.29
    Range("B7").Select
    Columns("B:B").ColumnWidth = 9.43
    Range("B2:D2").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
    Selection.Font.Size = 12
    Range("F2:J2").Select
    Selection.Font.Size = 12
    Selection.Font.Bold = True
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Range("B2:D2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Range("F3:J4").Select
    ActiveWindow.DisplayGridlines = False
    Range("B3:D4").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("B2:D2").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Columns("T:T").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    Cells.Select
    Selection.Locked = False
    Selection.FormulaHidden = False
    Range("B3:D4").Select
    Selection.Locked = True
    Selection.FormulaHidden = False
    Range("C10").Select
    Columns("A:A").ColumnWidth = 1.86
    Columns("T:T").ColumnWidth = 1.71
    Range("B8").Select
    Range("F3:J4").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = True
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Rows("4:4").EntireRow.AutoFit
    Rows("4:4").EntireRow.AutoFit
    Columns("J:J").EntireColumn.AutoFit
    Columns("J:J").ColumnWidth = 21.14
    Columns("J:J").ColumnWidth = 28.71
    Selection.Font.Size = 10
    Selection.Font.Size = 9
    Selection.Font.Size = 8
    Rows("4:4").EntireRow.AutoFit
    Rows("4:4").RowHeight = 21.75
    Rows("3:3").EntireRow.AutoFit
    Rows("3:3").RowHeight = 18
    Rows("3:3").RowHeight = 23.25
    Rows("8:8").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
        Range("B8").Select
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
        
            Columns("B:B").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    Range("B6").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("B8").Select
    Columns("B:B").ColumnWidth = 3.86
    Range("B8").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range(Selection, Selection.End(xlToRight)).Select


        ActiveSheet.EnableSelection = xlUnlockedCells
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingColumns:=True, AllowFormattingRows:=True, _
        AllowInsertingRows:=True, AllowInsertingHyperlinks:=True, _
        AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
        AllowUsingPivotTables:=True
        
        If Not GetReportFilePath = "" Then
        ActiveWorkbook.SaveAs (GetReportFilePath & "\" & scriptname & "_" & Format(Now(), "mmddhhmmss") & ".xlsx"), FileFormat:=51
        Exit Sub
        End If
      
        Dim path As String
        path = "C:\Users\" & Environ("Username") & "\Documents\PFM SmartApp"
        If Len(Dir(path, vbDirectory)) = 0 Then
        MkDir (path)
        End If
        If Len(Dir(path & "/" & Format(Now(), "yyyy"), vbDirectory)) = 0 Then
        MkDir (path & "/" & Format(Now(), "yyyy"))
        End If
        
        ActiveWorkbook.SaveAs (path & "/" & Format(Now(), "yyyy") & "\" & scriptname & "_" & Format(Now(), "mmddhhmmss") & ".xlsx"), FileFormat:=51
        
        

    'ActiveWorkbook.SaveAs Filename:="C:\Users\" & Environ("Username") & "\Documents\PFM SmartApp\" & Format(Now(), "yyyy") & "\" & scriptname & Format(Now(), "ddmmyyhhmm"), FileFormat:=xlWorkbookNormal



End Sub

    


Sub Planned_Cost_UpdateTemplate()
    Workbooks.Add
    Range("C8").Select
    ActiveCell.FormulaR1C1 = "WBS Element"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "Description"
    Range("E8").Select
    ActiveCell.FormulaR1C1 = "Amount"
    Range("F8").Select
    ActiveCell.FormulaR1C1 = "Curr"
    Range("G8").Select
    ActiveCell.FormulaR1C1 = "Cost Element"
    Range("H8").Select
    ActiveCell.FormulaR1C1 = "Status (SAP)"
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "This template is ment to update plan cost on list of WBS, the script will create a NWA on the existing NW of the WBS, in the case that the WBS does not have an existing NW, then the script will create a NW for it. Please keep in mind the script will not work on WBS with status TECO or CLSD, or WBS that have more than one NW. Please make sure that the amounts do not have more than 2 decimals."


End Sub
Sub Update_WBS_System_StatusTemplate()
    Workbooks.Add
    
    Application.ScreenUpdating = False
    
    Range("C8").Select
    ActiveCell.FormulaR1C1 = "WBS Element"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "Choose NEW status"
    Range("E8").Select
    ActiveCell.FormulaR1C1 = "Status (SAP)"
    Sheets("Sheet2").Select
    ActiveCell.FormulaR1C1 = "Set Release"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "Set CLSD"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Set TECO"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "Set FNBL"
    Range("A5").Select
    ActiveCell.FormulaR1C1 = "Remove CLSD"
    Range("A6").Select
    ActiveCell.FormulaR1C1 = "Remove TECO"
    Range("A7").Select
    ActiveCell.FormulaR1C1 = "Remove FNBL"

    Sheets("Sheet1").Select
    Range("D9").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=Sheet2!$A$1:$A$7"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "This template is ment to update status on the WBS e.g.(Final Billed, Techinically Completed, Closed)."

    Range("B8").Select

End Sub
Sub Update_Sales_Order_System_StatusTemplate()
    Workbooks.Add
    Range("C8").Select
    ActiveCell.FormulaR1C1 = "Sales Order"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "Choose NEW status"
    Range("E8").Select
    ActiveCell.FormulaR1C1 = "Status (SAP)"
    Sheets("Sheet2").Select
    ActiveCell.FormulaR1C1 = "Set CLSD"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "Set TECO"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Remove CLSD"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "Remove TECO"
    Range("A5").Select
    ActiveCell.FormulaR1C1 = "Set FNBL"
    Sheets("Sheet1").Select
    Range("D9").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=Sheet2!$A$1:$A$5"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("B8").Select
        Range("F3").Select
    ActiveCell.FormulaR1C1 = "This template is ment to update status on the SO e.g.(Technically Completed, Closed)."

End Sub
Sub POC_Milestone_Creation_UpdateTemplate()
    Workbooks.Add
    Range("C8").Select
    ActiveCell.FormulaR1C1 = "WBS Element"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "SIGNUM of requestor"
    Range("E8").Select
    ActiveCell.FormulaR1C1 = "Actual Date"
    Range("F8").Select
    ActiveCell.FormulaR1C1 = "POC"
    Range("G8").Select
    ActiveCell.FormulaR1C1 = "Link to supporting Documents"
    Range("H8").Select
    ActiveCell.FormulaR1C1 = "Status (SAP)"
        Range("F3").Select
    ActiveCell.FormulaR1C1 = "This template is ment to create a Milestone on the WBS for percentage of completion update (POC), it will also add the 100-200 POC parameter on the WBS if it does not have it. Please keep in mind that the DATE FORMAT needs to be exactly the same as the one you have in SAP."
    Columns("E:E").Select
    Selection.NumberFormat = "@"
        Range("B8").Select


End Sub
Sub Update_Value_Contract_System_StatusTemplate()
    Workbooks.Add
    Range("C8").Select
    ActiveCell.FormulaR1C1 = "Value Contract"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "Choose NEW status"
    Range("E8").Select
    ActiveCell.FormulaR1C1 = "Status (SAP)"
    Sheets("Sheet2").Select
    ActiveCell.FormulaR1C1 = "SIGN to CLOS (NONF)"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "SIGN to COMP (FIXD)"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "COMP to CLOS (FIXD)"
    Sheets("Sheet1").Select
    Range("D9").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=Sheet2!$A$1:$A$3"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
        Range("F3").Select
    ActiveCell.FormulaR1C1 = "This template is ment to update the value contract system status."
        Range("B8").Select

End Sub
Sub Update_Sales_Order_Revenue_StatusTemplate()
    Workbooks.Add
    Range("C8").Select
    ActiveCell.FormulaR1C1 = "Sales Order"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "Choose revenue status"
    Range("E8").Select
    ActiveCell.FormulaR1C1 = "Link to supporting document"
    Range("F8").Select
    ActiveCell.FormulaR1C1 = "Status (SAP)"
    Sheets("Sheet2").Select
    ActiveCell.FormulaR1C1 = "Set RREC"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "Set RSUR"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Remove All Status"
    Sheets("Sheet1").Select
    Range("D9").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=Sheet2!$A$1:$A$3"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
        Range("F3").Select
    ActiveCell.FormulaR1C1 = "This template is ment to update the PSF Sales Order revenue status, in order to take revenue (RREC) or defer it (RSUR)."
        Range("B8").Select

End Sub
Sub Update_Project_Finish_DateTemplate()
    Workbooks.Add
    Range("C8").Select
    ActiveCell.FormulaR1C1 = "Project Definition"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "New Finish Date"
    Range("E8").Select
    ActiveCell.FormulaR1C1 = "Status (SAP)"
            Range("F3").Select
    ActiveCell.FormulaR1C1 = "This template is ment to update the finish date for a list of projects. Please keep in mind that the DATE FORMAT needs to be exactly the same as the one you have in SAP."
        Range("B8").Select


End Sub
Sub ENO_Planned_Cost_Update_MarzennaTemplate()
    Workbooks.Add
    Range("C8").Select
    ActiveCell.FormulaR1C1 = "WBS Element"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "NWA Description"
    Range("E8").Select
    ActiveCell.FormulaR1C1 = "Amount"
    Range("F8").Select
    ActiveCell.FormulaR1C1 = "Curr"
    Range("G8").Select
    ActiveCell.FormulaR1C1 = "Cost Element"
    Range("H8").Select
    ActiveCell.FormulaR1C1 = "Status (SAP)"

End Sub
Sub Create_QTCM_Sales_OrderTemplate()
    Workbooks.Add
    Range("C8").Select
    ActiveCell.FormulaR1C1 = "Value Contract"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "PO Number"
    Range("E8").Select
    ActiveCell.FormulaR1C1 = "Curr"
    Range("F8").Select
    ActiveCell.FormulaR1C1 = "WBS Element"
    Range("G8").Select
    ActiveCell.FormulaR1C1 = "Description"
    Range("H8").Select
    ActiveCell.FormulaR1C1 = "RA KEY"
    Range("I8").Select
    ActiveCell.FormulaR1C1 = "Material"
    Range("J8").Select
    ActiveCell.FormulaR1C1 = "Amount"
    Range("K8").Select
    ActiveCell.FormulaR1C1 = "PCODE (KEY)"
    Range("L8").Select
    ActiveCell.FormulaR1C1 = "NEW SO (SAP)"
    Range("M8").Select
    ActiveCell.FormulaR1C1 = "NEW WBS (SAP)"
    Range("N8").Select
    ActiveCell.FormulaR1C1 = "NEW ITEM (SAP)"
    Range("O8").Select
    ActiveCell.FormulaR1C1 = "Status (SAP)"
    Sheets("Sheet2").Select
    ActiveCell.FormulaR1C1 = "EAB-DM_HW"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "EAB-DM_SW"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "EAB-DM_CS"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "EAB-ZNET_SERVICE"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "ZPS003"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "ZPS009"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "ZPS006"
    Sheets("Sheet1").Select
    Range("I9").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=Sheet2!$A$1:$A$4"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("H9").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=Sheet2!$B$1:$B$3"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "This template is ment to create a QTC-M sales order and its WBS, the script will create one object for each Value Contract and every variance in WBS Description column, please make sure to use the SAP PCODE and not the fire code, if the pcode starts with a zero please do not include it e.g. pcode 014 should be placed as pcode 14."
    Range("B8").Select



End Sub
Sub Update_Value_Contract_DescriptionTemplate()
    Workbooks.Add
    Range("C8").Select
    ActiveCell.FormulaR1C1 = "Value Contract"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "Description"
    Range("E8").Select
    ActiveCell.FormulaR1C1 = "Old text (SAP)"
    Range("F8").Select
    ActiveCell.FormulaR1C1 = "Status (SAP)"
        Range("F3").Select
    ActiveCell.FormulaR1C1 = "This template is ment to change the description on each Value Contract, and to retrieve the current one."
        Range("B8").Select


End Sub
Sub Run_Settlement_PSFTemplate()
    Workbooks.Add
    Range("C8").Select
    ActiveCell.FormulaR1C1 = "SO"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "Month"
    Range("E8").Select
    ActiveCell.FormulaR1C1 = "Year"
    Range("F8").Select
    ActiveCell.FormulaR1C1 = "KKA3 Status (SAP)"
    Range("G8").Select
    ActiveCell.FormulaR1C1 = "VA88 Status (SAP)"
        Range("F3").Select
    ActiveCell.FormulaR1C1 = "This template is ment run settlements for PSF Sales Orders, the script runs KKA3-VA88."
        Range("B8").Select

End Sub
Sub Run_Settlement_QTCTemplate()
    Workbooks.Add
    Range("C8").Select
    ActiveCell.FormulaR1C1 = "WBS Element"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "Month"
    Range("E8").Select
    ActiveCell.FormulaR1C1 = "Year"
    Range("F8").Select
    ActiveCell.FormulaR1C1 = "CJA2 Status (SAP)"
    Range("G8").Select
    ActiveCell.FormulaR1C1 = "CNE1 Status (SAP)"
    Range("H8").Select
    ActiveCell.FormulaR1C1 = "KKA2 Status (SAP)"
    Range("I8").Select
    ActiveCell.FormulaR1C1 = "CJ88 Status (SAP)"
        Range("F3").Select
    ActiveCell.FormulaR1C1 = "This template is ment run settlements for WBS list, the script runs CJA2-CNE1-CJ88 at a Project Definition level."
        Range("B8").Select

End Sub
Sub Create_RebateTemplate()
    Workbooks.Add
    Range("C8").Select
    ActiveCell.FormulaR1C1 = "Value Contract"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "Rebate Description"
    Range("E8").Select
    ActiveCell.FormulaR1C1 = "Valid From"
    Range("F8").Select
    ActiveCell.FormulaR1C1 = "Valid To"
    Range("G8").Select
    ActiveCell.FormulaR1C1 = "Sold to"
    Range("H8").Select
    ActiveCell.FormulaR1C1 = "Curr"
    Range("I8").Select
    ActiveCell.FormulaR1C1 = "Material"
    Range("J8").Select
    ActiveCell.FormulaR1C1 = "PCODE (KEY)"
    Range("K8").Select
    ActiveCell.FormulaR1C1 = "%"
    Range("L8").Select
    ActiveCell.FormulaR1C1 = "New Rebate # (SAP)"
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "This template is ment to create a rebate agreement for each Value Contract on the list, please make sure to use the SAP PCODE and not the fire code, if the pcode starts with a zero please do not include it e.g. pcode 014 should be placed as pcode 14. Please keep in mind that the DATE FORMAT needs to be exactly the same as the one you have in SAP."
    Sheets("Sheet2").Select
    ActiveCell.FormulaR1C1 = "EAB-DM_HW"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "EAB-DM_SW"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "EAB-DM_CS"
    Sheets("Sheet1").Select
    Range("I9").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=Sheet2!$A$1:$A$3"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("B8").Select


End Sub
Sub Update_RebateTemplate()
    Workbooks.Add
    Range("C8").Select
    ActiveCell.FormulaR1C1 = "Rebate#"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "New%"
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "This template is ment to update the percentage on a list of Rebate agreements."
    Range("B8").Select

End Sub
Sub Rebate_DescriptionTemplate()
    Workbooks.Add
    Range("C8").Select
    ActiveCell.FormulaR1C1 = "Rebate#"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "New Description"
    Range("E8").Select
    ActiveCell.FormulaR1C1 = "Status (SAP)"
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "This template is ment to change the description on a list of Rebate agreements."
    Range("B8").Select

End Sub
Sub VC_DescriptionTemplate()
    Workbooks.Add
    Range("C8").Select
    ActiveCell.FormulaR1C1 = "Value Contract"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "New Description"
    Range("E8").Select
    ActiveCell.FormulaR1C1 = "Old Description (SAP)"
    Range("F8").Select
    ActiveCell.FormulaR1C1 = "Status (SAP)"
        Range("F3").Select
    ActiveCell.FormulaR1C1 = "This template is ment to change the description on each Value Contract, and to retrieve the current one."
        Range("B8").Select
End Sub
Sub Update_Gstream_AssignmentIDTemplate()
    Workbooks.Add
    Range("C8").Select
    ActiveCell.FormulaR1C1 = "Value Contract"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "Assignment ID"
    Range("E8").Select
    ActiveCell.FormulaR1C1 = "Governance Stream"
    Range("F8").Select
    ActiveCell.FormulaR1C1 = "Status (SAP)"
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "This template is ment to update the Governance Stream or the Assignment ID on each Value Contract."
    Sheets("Sheet2").Select
    ActiveCell.FormulaR1C1 = "Customer Project"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "Product Delivery"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Customer Support"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "Managed Operations"
    Range("A5").Select
    ActiveCell.FormulaR1C1 = "Contract Financial Adjustments"
    Sheets("Sheet1").Select
    Range("E9").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=Sheet2!$A$1:$A$5"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("B8").Select
End Sub
Sub Update_SO_Rebate_ConditionTemplate()
    Workbooks.Add
    Range("C8").Select
    ActiveCell.FormulaR1C1 = "Sold To"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "Sales Order"
    Range("E8").Select
    ActiveCell.FormulaR1C1 = "From Date"
    Range("F8").Select
    ActiveCell.FormulaR1C1 = "Currency"
        Range("F3").Select
    ActiveCell.FormulaR1C1 = "This template is ment to change the rebate condition on the pricing items of a SO, it runs 5 SO at a time."
        Range("B8").Select
End Sub
Sub Create_Value_ContractTemplate()
    Range("C8").Select
    ActiveCell.FormulaR1C1 = "Contract type"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "Sales Area"
    Range("E8").Select
    ActiveCell.FormulaR1C1 = "Dist Channel"
    Range("F8").Select
    ActiveCell.FormulaR1C1 = "Division"
    Range("G8").Select
    ActiveCell.FormulaR1C1 = "Sales Office"
    Range("H8").Select
    ActiveCell.FormulaR1C1 = "Sales Group"
    Range("I8").Select
    ActiveCell.FormulaR1C1 = "Sold-to-Party"
    Range("J8").Select
    ActiveCell.FormulaR1C1 = "Ship to Party"
    Range("K8").Select
    ActiveCell.FormulaR1C1 = "Valid From:"
    Range("L8").Select
    ActiveCell.FormulaR1C1 = "Valid to:"
    Range("M8").Select
    ActiveCell.FormulaR1C1 = "PO Number"
    Range("N8").Select
    ActiveCell.FormulaR1C1 = "Description"
    Range("O8").Select
    ActiveCell.FormulaR1C1 = "Curr"
    Range("P8").Select
    ActiveCell.FormulaR1C1 = "Material"
    Range("Q8").Select
    ActiveCell.FormulaR1C1 = "Material Description"
    Range("R8").Select
    ActiveCell.FormulaR1C1 = "Amount"
    Range("S8").Select
    ActiveCell.FormulaR1C1 = "PCODE (KEY)"
    Range("T8").Select
    ActiveCell.FormulaR1C1 = "Billing Date"
    Range("U8").Select
    ActiveCell.FormulaR1C1 = "DtD"
    Range("V8").Select
    ActiveCell.FormulaR1C1 = "%"
    Range("W8").Select
    ActiveCell.FormulaR1C1 = "Order Reason"
    Range("X8").Select
    ActiveCell.FormulaR1C1 = "CCLM ID"
    Range("Y8").Select
    ActiveCell.FormulaR1C1 = "Assignment"
    Range("Z8").Select
    ActiveCell.FormulaR1C1 = "LAC Email"
    Range("AA8").Select
    ActiveCell.FormulaR1C1 = "CEM #"
    Range("AB8").Select
    ActiveCell.FormulaR1C1 = "Sponsor #"
    Range("AC8").Select
    ActiveCell.FormulaR1C1 = "CPM #"
    Range("AD8").Select
    ActiveCell.FormulaR1C1 = "PSP #"
    Range("AE8").Select
    ActiveCell.FormulaR1C1 = "NEW VC # (SAP)"

    Range("F3").Select
    ActiveCell.FormulaR1C1 = "This template is ment to create a Value Contract"


End Sub

Sub Update_BillingTypeTemplate()
    Workbooks.Add
    Range("C8").Select
    ActiveCell.FormulaR1C1 = "Sales Order"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "Billing Date"
    Range("E8").Select
    ActiveCell.FormulaR1C1 = "Status (SAP)"
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "This template is ment to delete billing Type and change it to Delivery with acceptance"

End Sub
Sub Update_Partner_VCTemplate()
    Workbooks.Add
    Range("C8").Select
    ActiveCell.FormulaR1C1 = "VC"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "Partner"
    Range("E8").Select
    ActiveCell.FormulaR1C1 = "Employee Number"
    Range("F8").Select
    ActiveCell.FormulaR1C1 = "Status (SAP)"
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "This template is ment update or add partners on the Value Contract"
    Sheets("Sheet2").Select
    ActiveCell.FormulaR1C1 = "Contract Accountable"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "Contract Responsable"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Order Responsable"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "Exec Responsable"
    Range("A5").Select
    ActiveCell.FormulaR1C1 = "Sponsor"
    Range("A6").Select
    ActiveCell.FormulaR1C1 = "PSP"
    Sheets("Sheet1").Select
    Range("D9").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=Sheet2!$A$1:$A$6"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
        Range("F3").Select
End Sub

Sub Airtel_Rebate_ListTemplate()
    Workbooks.Add
    Range("C8").Select
    ActiveCell.FormulaR1C1 = "Sold To"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "Status (SAP)"
End Sub


