Attribute VB_Name = "AirtelRebateList"
Sub AirtelRebateListScript(ByRef trx As String, ByRef SAPConnection As Variant, ByRef CMail As Variant)
    
    On Error GoTo ErrHandler
    Dim sBar
    Set sBar = SAPConnection.session.findById("wnd[0]/sbar")
SAPConnection.session.findById("wnd[0]").Maximize
SAPConnection.session.findById("wnd[0]/tbar[0]/okcd").Text = "/NVB(8"
SAPConnection.session.findById("wnd[0]").sendVKey 0
SAPConnection.session.findById("wnd[0]/tbar[1]/btn[17]").press
SAPConnection.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 1
SAPConnection.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "1"
SAPConnection.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell
SAPConnection.session.findById("wnd[0]/usr/chkBCHECK").Selected = False
SAPConnection.session.findById("wnd[0]/usr/chkBREADY").Selected = False
SAPConnection.session.findById("wnd[0]/usr/chkBREADY").SetFocus
SAPConnection.session.findById("wnd[0]/usr/chkBCHECK").Selected = True
SAPConnection.session.findById("wnd[0]/usr/chkBCREDI").Selected = True
SAPConnection.session.findById("wnd[0]/usr/chkBREADY").Selected = True
SAPConnection.session.findById("wnd[0]/usr/chkBSETTLE").Selected = True
SAPConnection.session.findById("wnd[0]/usr/chkBOPEN").Selected = True
SAPConnection.session.findById("wnd[0]/usr/ctxtKNUMA-LOW").Text = ""
SAPConnection.session.findById("wnd[0]/tbar[1]/btn[8]").press
        Progressfrm.Text.Caption = Round((0.3) * 100, 0) & "% Processed"
        Progressfrm.Bar.Width = (0.3) * 306
SAPConnection.session.findById("wnd[0]/usr/lbl[0,1]").SetFocus
SAPConnection.session.findById("wnd[0]/usr/lbl[0,1]").caretPosition = 16
SAPConnection.session.findById("wnd[0]/mbar/menu[4]/menu[5]/menu[2]/menu[2]").Select
SAPConnection.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
SAPConnection.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus
SAPConnection.session.findById("wnd[1]/tbar[0]/btn[0]").press
Application.StatusBar = False
    Workbooks.Add
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "X"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "X"
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "Original"
    Sheets("Sheet2").Select
    ActiveSheet.PasteSpecial Format:="Unicode Text", Link:=False, _
        DisplayAsIcon:=False, NoHTMLFormatting:=True

    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(3, 1), Array(14, 1), Array(25, 1), Array(34, 1), _
        Array(45, 1), Array(70, 1), Array(80, 1), Array(91, 1), Array(101, 1)), _
        TrailingMinusNumbers:=True
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Rows("1:13278").Select
    Range(Selection, Selection.End(xlDown)).Select
    Cells.Select
    Selection.AutoFilter
    Rows("2:13468").Select
    Range(Selection, Selection.End(xlDown)).Select
        Progressfrm.Text.Caption = Round((0.4) * 100, 0) & "% Processed"
        Progressfrm.Bar.Width = (0.4) * 306
    Rows("2:65536").Select
    ActiveSheet.Range("$A$1:$J$65536").AutoFilter field:=2, Criteria1:= _
        "-----------"
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("$A$1:$J$65536").AutoFilter field:=2
    ActiveSheet.Range("$A$1:$J$65536").AutoFilter field:=1, Criteria1:= _
        "@There are"
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("$A$1:$J$65536").AutoFilter field:=1, Criteria1:= _
        "Agreement"
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("$A$1:$J$65536").AutoFilter field:=1, Criteria1:= _
        "Condition k"
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("$A$1:$J$65536").AutoFilter field:=1, Criteria1:= _
        "CTyp Name"
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("$A$1:$J$65536").AutoFilter field:=1, Criteria1:= _
        "ZBOP /// co"
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("$A$1:$J$65536").AutoFilter field:=1, Criteria1:= _
        "Sales org."
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("$A$1:$J$65536").AutoFilter field:=1, Criteria1:="="
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("$A$1:$J$65536").AutoFilter field:=1
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A6").Select
    ActiveWindow.DisplayGridlines = False
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Sales Org"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Distr Chan"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Sales Group"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Sold to"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Rebate #"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Description"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Value Contract"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "PCODE"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Valid From"
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Valid to"
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Customer"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=IF(ISBLANK(RC[3]),RC[1]&RC[2],R[-1]C)"
    Range("E2").Select
        Progressfrm.Text.Caption = Round((0.5) * 100, 0) & "% Processed"
        Progressfrm.Bar.Width = (0.5) * 306
    Selection.Copy
    Selection.End(xlDown).Select
    Range("F65536").Select
    Selection.End(xlUp).Select
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("F1").Select
    Columns("E:E").EntireColumn.AutoFit
    Range("E21").Select
    Range("B2:C2").Select
    Selection.NumberFormat = "m/d/yyyy"
    Columns("A:A").EntireColumn.AutoFit
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=IF(ISBLANK(RC[7]),PROPER(RC[9]&RC[10]),R[-1]C)"
    Range("M26").Select
    Range("E2").Select
    ActiveSheet.Range("$F$1:$O$65536").AutoFilter field:=3
    ActiveCell.FormulaR1C1 = "=IF(ISBLANK(RC[8]),R[-1]C,RC[1]&RC[2]&RC[3])"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=IF(ISBLANK(RC[10]),R[-1]C,RC[9])"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=IF(ISBLANK(RC[11]),R[-1]C,RC[11])"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=IF(ISBLANK(RC[12]),R[-1]C,PROPER(RC[9]&RC[10]))"
    Range("A2:E2").Select
    Selection.Copy
    Range("F2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, -5).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("C3425").Select
    Selection.End(xlUp).Select
    'SELECT COLUMNS
    Columns("A:E").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveCell.Rows("1:1").EntireRow.Select
    Selection.AutoFilter
    Selection.AutoFilter
    Columns("A:E").Select
    Range("E1").Activate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B7").Select
    Application.CutCopyMode = False
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Rows("2:65536").Select
    Selection.AutoFilter
    Rows("1:1").Select
    Selection.AutoFilter
    Rows("2:65536").Select
    ActiveSheet.Range("$A$1:$N$65536").AutoFilter field:=13, Criteria1:="<>"
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("$A$1:$N$65536").AutoFilter field:=13
    Columns("N:N").Select
    Selection.ClearContents
        Progressfrm.Text.Caption = Round((0.6) * 100, 0) & "% Processed"
        Progressfrm.Bar.Width = (0.6) * 306
    Columns("L:L").Select
    Columns("K:K").Select
    Selection.Replace What:=" Condit", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Columns("L:L").Select
    Selection.Replace What:="ion deleted", Replacement:="Condition Deleted", _
        LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:= _
        False, ReplaceFormat:=False
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("E:E").Select
    Selection.TextToColumns Destination:=Range("E1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(10, 2)), TrailingMinusNumbers:=True
    Range("D1").Select
    Selection.Cut
    Range("F1").Select
    ActiveSheet.Paste
    Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft
    Columns("B:C").Select
    Range("C1").Activate
    Selection.Cut
    Columns("L:L").Select
    Selection.Insert Shift:=xlToRight
    Columns("A:A").Select
    Selection.Cut
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight
    Columns("A:A").Select
    Columns("A:A").EntireColumn.AutoFit
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
    Columns("C:E").Select
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
    Columns("H:H").Select
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
    Selection.ColumnWidth = 23.29
    Columns("H:H").EntireColumn.AutoFit
    Columns("I:I").Select
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
        Progressfrm.Text.Caption = Round((0.7) * 100, 0) & "% Processed"
        Progressfrm.Bar.Width = (0.7) * 306
    Columns("I:I").EntireColumn.AutoFit
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Status"
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Rows("1:1").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
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
    Selection.Font.Bold = True
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    ActiveWindow.Zoom = 85
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
        Progressfrm.Text.Caption = Round((0.8) * 100, 0) & "% Processed"
        Progressfrm.Bar.Width = (0.8) * 306
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
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
    Range("L1").Select
    Columns("L:L").EntireColumn.AutoFit
    Rows("1:4").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A6").Select
    Columns("A:A").ColumnWidth = 0.92
    Range("B2").Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 28
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
    ActiveCell.FormulaR1C1 = "REBATE LIST"
    Range("B2").Select
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
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "=TODAY()"
    Range("B3").Select
    Selection.NumberFormat = "[$-F800]dddd, mmmm dd, yyyy"
    Range("B3").Copy
        Progressfrm.Text.Caption = Round((0.9) * 100, 0) & "% Processed"
        Progressfrm.Bar.Width = (0.9) * 306
    Range("B3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B5").Select
    Columns("B:B").ColumnWidth = 11.57
    Columns("B:B").ColumnWidth = 12.86
    Columns("B:B").EntireColumn.AutoFit
    Columns("B:B").ColumnWidth = 20.43
    Columns("B:B").ColumnWidth = 11.71
    Range("B3:C3").Select
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
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Columns("H:H").EntireColumn.AutoFit
    Range("N5").Select
        Progressfrm.Text.Caption = Round((1) * 100, 0) & "% Processed"
        Progressfrm.Bar.Width = (1) * 306
        
        'Save File
        Dim path As String
        path = "C:\Users\" & Environ("Username") & "\Documents\PFM SmartApp"
        If Len(Dir(path, vbDirectory)) = 0 Then
        MkDir (path)
        End If
        If Len(Dir(path & "/" & Format(Now(), "yyyy"), vbDirectory)) = 0 Then
        MkDir (path & "/" & Format(Now(), "yyyy"))
        End If
        
        ActiveWorkbook.SaveAs (path & "/" & Format(Now(), "yyyy") & "\" & "Rebate List " & "_" & Format(Now(), "mmddhhmmss") & ".xlsx"), FileFormat:=51

Exit Sub
ErrHandler:
    If Err.Number <> 0 Then
        GoTo ErrGoNext
    End If
    Exit Sub
ErrGoNext:
    SAPConnection.ErrorCounter = SAPConnection.ErrorCounter + 1
    'SAPConnection.reportError "Airtel_Rebate_List", ActiveCell.Offset(0, 1), Err.Number, Err.Description, Err.Source, sBar.Text
    CMail.BuildErrorList ActiveCell.Offset(0, 1), "Airtel_Rebate_List", Err.Number, Err.Description, Err.Source, sBar.Text
    Err.Clear
    SAPConnection.errorContinueNextItem (trx)
End Sub
