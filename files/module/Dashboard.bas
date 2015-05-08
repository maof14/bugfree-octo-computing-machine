Attribute VB_Name = "Dashboard"
Sub Create_Dashboard()
Dim sName As String

Application.DisplayAlerts = False
Application.ScreenUpdating = False
    Workbooks.Open fileName:= _
        "C:\Users\" & Environ("USERNAME") & "\Documents\PFM SmartApp\Dashboards\COPA FI ISO.xlsm", _
        UpdateLinks:=0
    Cells.Select
    Selection.Copy
    Workbooks.Add
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Windows("COPA FI ISO.xlsm").Activate
    Application.CutCopyMode = False
    ActiveWindow.Close
    Range("A1").Select
    sName = "ISO"
Call FormatReport(sName)
    wName = ActiveWorkbook.Name
    Workbooks.Open fileName:= _
        "C:\Users\" & Environ("USERNAME") & "\Documents\PFM SmartApp\Dashboards\COPA FI YTD.xlsm", _
        UpdateLinks:=0
    Cells.Select
    Selection.Copy
    Windows(wName).Activate
    Sheets("Sheet2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Windows("COPA FI YTD.xlsm").Activate
    Application.CutCopyMode = False
    ActiveWindow.Close
    Range("A1").Select
    sName = "Accumulated"
Call FormatReport(sName)
    wName = ActiveWorkbook.Name
    Sheets("Accumulated").Select
    Rows("4:4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    ActiveSheet.Previous.Select
    Range("A3").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Rows("4:4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Workbooks.Open fileName:= _
        "C:\Users\" & Environ("USERNAME") & "\Documents\PFM SmartApp\Dashboards\Dashboard.xlsm"
    Sheets("Data").Select
    
    Dim rngFilter As Range
    Dim r As Long, f As Long
    Set rngFilter = ActiveSheet.AutoFilter.Range
    r = rngFilter.Rows.Count
    f = rngFilter.SpecialCells(xlCellTypeVisible).Count
    If r > f Then
    ActiveSheet.ShowAllData
    Else
    End If
    
    Rows("5:5").Select
    Selection.Insert Shift:=xlDown
    Rows("4:4").Select
    Rows("5:5").Select
    Selection.Insert Shift:=xlDown
    Rows("4:4").Select
    Windows(wName).Activate
    Selection.Copy
    Windows("Dashboard.xlsm").Activate
    Rows("4:4").Select
    Selection.Insert Shift:=xlDown
    Range("A15").Select
    Selection.End(xlDown).Select
    Rows(ActiveCell.Offset(0, 0).Row & ":" & ActiveCell.Offset(0, 0).Row).Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Selection.Delete Shift:=xlUp
    Selection.Delete Shift:=xlUp
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    Range("A4").Select
    
    'change RUC and DR to positive
            Range("AV3").Select
            ActiveCell.FormulaR1C1 = "-1"
            Range("AV3").Select
            Selection.Copy
            Range("AQ4").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlMultiply, _
                SkipBlanks:=False, Transpose:=False
            Range("AV3").Select
            Application.CutCopyMode = False
            Selection.Copy
            Range("AS4").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlMultiply, _
                SkipBlanks:=False, Transpose:=False
            Range("AV3").Select
            Application.CutCopyMode = False
            Selection.ClearContents
            Range("A3").Select
            Range("A4").Select
            
    Windows("Dashboard.xlsm").Activate
    ActiveWorkbook.RefreshAll
    Windows(wName).Activate
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Sheets("Dashboard").Select


End Sub
Sub FormatReport(ByRef sName As String)

Dim topleftcorner As Integer
Dim toprightcorner As Integer
Dim lastrow As Long
Dim valtable As Range
Dim pivrange As Range
Dim pivName As String
Dim Pivsheet As String
Dim iRet As Integer
Dim RepLayout As Integer
Dim c As Range
Dim newvalue As String


'      iRet = MsgBox("This Script is currently in BETA, do you wish to continue?", vbYesNo, "BETA Script (UNDER DEVELOPMENT)")
'            If iRet = vbNo Then
'            End
'            Else
'            End If


Application.ScreenUpdating = False
    Range("A1").Select
    Rows("1:14").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Columns("A:F").Select
    Range("F1").Activate
    Selection.Delete Shift:=xlToLeft
    Range("D1").Select
    Selection.ClearContents
    Range("G1").Select
    Selection.ClearContents
    Range("N1").Select
    Selection.ClearContents
    Range("U1").Select
    Selection.ClearContents
    Range("W1").Select
    Selection.ClearContents
    Range("Z1").Select
    Selection.ClearContents
    Range("AB1").Select
    Selection.ClearContents
    Rows("1:1").Select
    Range("AC1").Activate
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AG2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range("AG1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("AG2").Select
    ActiveCell.FormulaR1C1 = "SEK"
    Range("AG2").Select
    Selection.Copy
    Range(Selection, Selection.End(xlToRight)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("A2").Select
    Pivsheet = ActiveSheet.Name
    Sheets(Pivsheet).Select
    Sheets(Pivsheet).Name = sName
    Dim obj As Shape
     
    For Each obj In ActiveWorkbook.ActiveSheet.Shapes
        obj.Delete
    Next
    Range("A1").Select
    If ActiveCell.Offset(6, 5).value = "Table" And ActiveCell.Offset(7, 5).value = "" Then
    Rows("1:7").Select
    Selection.Delete Shift:=xlUp
    Columns("A:E").Select
    Range("E1").Activate
    Selection.Delete Shift:=xlToLeft
    ElseIf ActiveCell.Offset(6, 3).value = "Table" And ActiveCell.Offset(7, 3).value = "" Then
    Rows("1:7").Select
    Selection.Delete Shift:=xlUp
    Columns("A:C").Select
    Range("C1").Activate
    Selection.Delete Shift:=xlToLeft
    ElseIf ActiveCell.value = "COPA Detail Analysis" Then
    Rows("1:30").Select
    Selection.Delete Shift:=xlUp
    Else
    End If
    
    If ActiveCell.value <> "" And ActiveCell.Offset(1, 0).value <> "" Then
    RepLayout = True
    Else
    End If
    
    Cells.Select
    Selection.NumberFormat = "General"
    With Selection
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection.Font
        .Name = "Calibri"
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    ActiveWindow.Zoom = 85
'    If RepLayout = False Then
    Rows("2:2").Select
    Selection.Copy
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Application.CutCopyMode = False
    Range("XFD2").Select
    Selection.End(xlToLeft).Select
    toprightcorner = ActiveCell.column
    Range("A2").Select
    Selection.End(xlToRight).Select
    topleftcorner = ActiveCell.column
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Cells(ActiveCell.Row, 1)).Select
    Selection.ClearContents
    Range(Cells(1, topleftcorner), Cells(2, toprightcorner)).Select
    Selection.Cut
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    Range(Cells(3, 1), Cells(3, toprightcorner)).Select
    
    Progressfrm.Text.Caption = "10% Completion"
    Progressfrm.Bar.Width = 0.1 * 306
Application.ScreenUpdating = False
    Dim blnkrng As Range
    
        Set blnkrng = Selection
        If WorksheetFunction.CountIf(blnkrng, "") > 0 Then
            Selection.SpecialCells(xlCellTypeBlanks).Select
            Selection.FormulaR1C1 = "=RC[-1]&"" Description"""
        Else
        'if there are no blanks in the selection then it wont do anything
        End If
        
        'remove the space on the header
        For Each c In Selection
            newvalue = Replace(c.value, Chr(10), " ")
            c.value = newvalue
        Next c


    Progressfrm.Text.Caption = "20% Completion"
    Progressfrm.Bar.Width = 0.2 * 306
    
    Selection.Replace What:=" SAP", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Range("A1048576").Select
    Selection.End(xlUp).Select
    lastrow = ActiveCell.Row
    Selection.End(xlUp).Select
   
    Range(Cells(4, topleftcorner), Cells(lastrow, toprightcorner)).Select
    Set valtable = Selection
        If WorksheetFunction.CountIf(valtable, "") > 0 Then
            Selection.SpecialCells(xlCellTypeBlanks).Select
            Selection.FormulaR1C1 = "0"
        Else
        'if there are no blanks in the selection then it wont do anything
        End If
    Range("A3").Select
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
    valtable.Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlDash
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlDash
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDash
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlDash
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlDash
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDash
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13434879
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
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
        .LineStyle = xlDash
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDash
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    

    Progressfrm.Text.Caption = "50% Completion"
    Progressfrm.Bar.Width = 0.5 * 306
    Application.Wait (Now + TimeValue("00:00:02"))
    Range("C3").Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
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
    Range("A3").Select
    Range(Selection, Selection.End(xlToRight)).Select
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
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Font.Bold = True
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
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
    Rows("2:2").Select
    Selection.Font.Bold = True
    With Selection
        .HorizontalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    

    Progressfrm.Text.Caption = "70% Completion"
    Progressfrm.Bar.Width = 0.7 * 306
    Application.Wait (Now + TimeValue("00:00:02"))
 
    Range(Cells(2, topleftcorner), Cells(2, toprightcorner)).Select
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
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Cells.Select
    Cells.EntireColumn.AutoFit
    ActiveWindow.DisplayGridlines = False
    Cells.EntireRow.AutoFit
    
    Progressfrm.Text.Caption = "80% Completion"
    Progressfrm.Bar.Width = 0.8 * 306
    
    
    Selection.Replace What:="_22", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    Range(Cells(4, 1), Cells(lastrow, topleftcorner - 1)).Select
        For Each c In Selection
        newvalue = Replace(c.value, "?", " ")
        c.value = newvalue
        Next c
        
    Progressfrm.Text.Caption = "90% Completion"
    Progressfrm.Bar.Width = 0.9 * 306
        
    Range("A3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Set pivrange = Selection
    
    Columns("Y:Z").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("X3").Select
    Selection.Copy
    Range("Y3:Z3").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("Z4").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-2]=""#"",R[-1]C[-2],RC[-2])"
    Range("Z4").Select
    Selection.Copy
    Range("AA4").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Columns("Z:Z").Select
    Range("Z429").Activate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("Z430").Select
    Selection.End(xlUp).Select
    Range("Y4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-16]=R[-1]C[-16],IF(RC[-1]<>R[-1]C[-1],R[-1]C&""; ""&RC[-1],R[-1]C),RC[-1])"
    Range("Y4").Select
    Selection.Copy
    Range("X4").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Application.CutCopyMode = False
    Range(Selection, Selection.End(xlUp)).Select
    Range("Y4").Select
    Selection.Copy
    Range("X4").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Columns("Y:Y").Select
    Range("Y429").Activate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("Y431").Select
    Application.CutCopyMode = False
    Range("Y430").Select
    Selection.End(xlUp).Select
    Columns("Y:Y").Select
    Range("Y3").Activate
    Selection.Replace What:="; #", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("Y178").Select
    Selection.End(xlUp).Select
    ActiveCell.FormulaR1C1 = "Commodity Group"
    Columns("X:X").Select
    Range("X3").Activate
    Selection.Delete Shift:=xlToLeft
    Range("X3").Select
    
    Range("AT4").Select
    ActiveCell.FormulaR1C1 = sName
    Range("AT4").Select
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("AN3").Select


End Sub

