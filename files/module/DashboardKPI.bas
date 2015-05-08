Attribute VB_Name = "DashboardKPI"
Sub Create_DashboardKPI()
Dim sName As String

Application.DisplayAlerts = False
Application.ScreenUpdating = False
    Workbooks.Open fileName:= _
        "C:\Users\" & Environ("USERNAME") & "\Documents\PFM SmartApp\Dashboards\ASC KPI.xlsm" _
        , UpdateLinks:=0
    Cells.Select
    Selection.Copy
    Workbooks.Add
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A2").Select
    sName = "PTD"
Call FormatReport(sName)
    wName = ActiveWorkbook.Name
    Windows("ASC KPI.xlsm").Activate
    ActiveWindow.Close
    Workbooks.Open fileName:= _
        "C:\Users\" & Environ("USERNAME") & "\Documents\PFM SmartApp\Dashboards\ASC PSF-M KPI.xlsm" _
        , UpdateLinks:=0
    Cells.Select
    Selection.Copy
    Workbooks.Add
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A2").Select
    sName = "PTD"
Call FormatReport(sName)
    wName2 = ActiveWorkbook.Name
    Windows("ASC PSF-M KPI.xlsm").Activate
    ActiveWindow.Close
    Range("A4:BE4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows(wName).Activate
    Range("A4").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    Range("A4").Select
    Windows(wName2).Activate
    ActiveWorkbook.Close
    Workbooks.Open fileName:= _
        "C:\Users\" & Environ("USERNAME") & "\Documents\PFM SmartApp\Dashboards\Dashboard KPI.xlsm"
    Sheets("Data").Select
    Sheets("Data").Copy After:=Sheets(9)
    Sheets("Data (2)").Select
    Sheets("Data (2)").Name = "Old"
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
    Rows("4:4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("Dashboard KPI.xlsm").Activate
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
    Range("BF4").Select
    ActiveCell.FormulaR1C1 = "=IF(ISERROR(SEARCH(""QTC"",RC[-25],1)>0),0,1)"
    Range("BG4").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-13]<=RC[-17],""Good"",""Needs update"")"
    Range("BH4").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(RC[-12]>0,RC[-15]<100),""Yes"",""No"")"
    Range("BI4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-43]=""Completely processed"",""Invoiced"",""Not Invoiced"")"
    Range("BF4:BI4").Select
    Selection.Copy
    Range("BE4").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    'Copy previous comments
    Range("BJ4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(VLOOKUP(RC[-45],old!C[-45]:C,46,FALSE)),"""",IF(VLOOKUP(RC[-45],old!C[-45]:C,46,FALSE)=0,"""",VLOOKUP(RC[-45],old!C[-45]:C,46,FALSE)))"
    Range("BJ4").Select
    Selection.Copy
    Range("BI4").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.End(xlUp).Select
    Columns("BJ:BJ").Select
    Range("BJ3").Activate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    'Age
    Range("BK4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF((TODAY()-RC[-43])<61,""<60"",IF(AND((TODAY()-RC[-43])>60,(TODAY()-RC[-43])<121),""61-120"",IF(AND((TODAY()-RC[-43])>120,(TODAY()-RC[-43])<181),""121-180"",IF((TODAY()-RC[-43])>180,"">181"",))))"
    Range("BK4").Select
    Selection.Copy
    Range("BI4").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveCell.Columns("A:A").EntireColumn.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

    Range("BJ3").Select
    Sheets("old").Select
    ActiveWindow.SelectedSheets.Delete
    Windows("Dashboard KPI.xlsm").Activate
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
    Columns("A:F").Select
    Range("F1").Activate
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Rows("1:14").Select
    Selection.Delete Shift:=xlUp
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AL2:BG2").Select
    Selection.Cut
    Range("AL1").Select
    ActiveSheet.Paste
    Range("AL2").Select
    ActiveCell.FormulaR1C1 = "SEK"
    Range("AL2").Select
    Selection.Copy
    Range("AL2:BG2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("E2").Select
    Selection.ClearContents
    Range("G2").Select
    Selection.ClearContents
    Range("I2").Select
    Selection.ClearContents
    Range("L2").Select
    Selection.ClearContents
    Range("V2").Select
    Selection.ClearContents
    Range("X2").Select
    Selection.ClearContents
    Range("AB2").Select
    Selection.ClearContents
    Range("AG2").Select
    Selection.ClearContents
    Range("AN2").Select
    Range("I2").Select
    ActiveWindow.DisplayGridlines = False
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
    
    Application.StatusBar = "10% Completion"
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


    Application.StatusBar = "20% Completion"
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
    

    Application.StatusBar = "50% Completion"
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
    

    Application.StatusBar = "70% Completion"
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
    
    Application.StatusBar = "80% Completion"
    Progressfrm.Bar.Width = 0.8 * 306
    
    
    Selection.Replace What:="_22", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
'    Range(Cells(4, 1), Cells(lastrow, topleftcorner - 1)).Select
'        For Each c In Selection
'        newvalue = Replace(c.Value, "?", " ")
'        c.Value = newvalue
'        Next c
        
    Application.StatusBar = "90% Completion"
    Progressfrm.Bar.Width = 0.9 * 306
        
    Range("A3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Set pivrange = Selection
    
    
        Columns("AL:AM").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AK3").Select
    Selection.Copy
    Range("AL3").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Range("AM3").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("AL3").Select
    ActiveCell.FormulaR1C1 = "Commodity2"
    Range("AM3").Select
    ActiveCell.FormulaR1C1 = "Commodity3"
    Range("AL4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-21]=R[-1]C[-21],IF(RC[-1]=R[-1]C[-1],R[-1]C,RC[-1]&""; ""&R[-1]C),RC[-1])"
    Range("AL4").Select
    Range("AM4").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-22]=R[1]C[-22],R[1]C,RC[-1])"
    Range("AL4:AM4").Select
    Selection.Copy
    Range("AK4").Select
    Selection.End(xlDown).Select
    ActiveWindow.SmallScroll Down:=3
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.End(xlUp).Select
    Rows("3:3").Select
    Range("AA3").Activate
    Range("AK6").Select
    Range("A3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Set pivrange = Selection
    Selection.AutoFilter

    pivrange.AutoFilter field:=37, Criteria1:="#"
    Range("AK3").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C"
    Range("AK3").Select
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.ShowAllData
    Columns("AK:AK").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("AK:AM").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("AK3").Select
    Application.CutCopyMode = False

    Range("A3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Set pivrange = Selection
        
    pivName = "Pivot" & Format(Now(), "mmddhhmmss")
    Sheets.Add
    Pivsheet = ActiveSheet.Name
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        pivrange, Version:=xlPivotTableVersion14). _
        CreatePivotTable TableDestination:=Pivsheet & "!R3C1", tablename:=pivName _
        , DefaultVersion:=xlPivotTableVersion14
    Cells(3, 1).Select
    Range("B7").Select
    With ActiveSheet.PivotTables(pivName)
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
    End With
    Sheets(Pivsheet).Select
    Sheets(Pivsheet).Name = "Pivot"
    ActiveSheet.PivotTables(pivName).TableStyle2 = "PivotStyleDark13"
    Rows("1:2").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A2").Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 18
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
    ActiveCell.FormulaR1C1 = "PIVOT"
    Range("A2").Select
    With Selection.Font
        .Name = "Ericsson Capital TT"
        .Size = 18
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Range("A2").Select
    With Selection.Font
        .Name = "Ericsson Capital TT"
        .Size = 22
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Sheets("Pivot").Select
    With ActiveWorkbook.Sheets("Pivot").Tab
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0
    End With

    Range("B9").Select
    With ActiveSheet.PivotTables(pivName).PivotFields("Customer Unit")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields("CRG")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields("Country")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields("Assignment ID")
        .Orientation = xlRowField
        .Position = 4
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields( _
        "Assignment ID Description")
        .Orientation = xlRowField
        .Position = 5
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields( _
        "Project Definition")
        .Orientation = xlRowField
        .Position = 6
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields( _
        "Project Definition Description")
        .Orientation = xlRowField
        .Position = 7
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields("WBS Element")
        .Orientation = xlRowField
        .Position = 8
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields( _
        "WBS Element Description")
        .Orientation = xlRowField
        .Position = 9
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields("PSP Comments")
        .Orientation = xlRowField
        .Position = 10
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields("WBS RA Key")
        .Orientation = xlRowField
        .Position = 11
    End With
    ActiveSheet.PivotTables(pivName).PivotFields("WBS RA Key"). _
        Orientation = xlHidden
    With ActiveSheet.PivotTables(pivName).PivotFields( _
        "WBS RA Key Description")
        .Orientation = xlRowField
        .Position = 11
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields("WBS RA Key")
        .Orientation = xlRowField
        .Position = 12
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields( _
        "FNBL stat chg date")
        .Orientation = xlRowField
        .Position = 13
    End With
    ActiveWindow.SmallScroll Down:=-6
    With ActiveSheet.PivotTables(pivName).PivotFields( _
        "System stat chg date")
        .Orientation = xlRowField
        .Position = 14
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields("WBS System Status" _
        )
        .Orientation = xlRowField
        .Position = 15
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields("Customer PO")
        .Orientation = xlRowField
        .Position = 16
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields("Sales Order")
        .Orientation = xlRowField
        .Position = 17
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields("Billing status")
        .Orientation = xlRowField
        .Position = 18
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields( _
        "Assignment Execution")
        .Orientation = xlRowField
        .Position = 19
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields("Created on (SO)")
        .Orientation = xlRowField
        .Position = 20
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields("Delivery Status")
        .Orientation = xlRowField
        .Position = 21
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields( _
        "Delivery Status Description")
        .Orientation = xlRowField
        .Position = 22
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields("Reason for order")
        .Orientation = xlRowField
        .Position = 23
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields( _
        "Reason for order Description")
        .Orientation = xlRowField
        .Position = 24
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields("Project Manager")
        .Orientation = xlRowField
        .Position = 25
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields("ESTA Flag")
        .Orientation = xlRowField
        .Position = 26
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields("Customer group 5")
        .Orientation = xlRowField
        .Position = 27
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields( _
        "Customer group 5 Description")
        .Orientation = xlRowField
        .Position = 28
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields("Changed on")
        .Orientation = xlRowField
        .Position = 29
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields("Governance Stream" _
        )
        .Orientation = xlRowField
        .Position = 30
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields( _
        "Text LC -> E/// BU")
        .Orientation = xlRowField
        .Position = 31
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields("Sales doc. type")
        .Orientation = xlRowField
        .Position = 32
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields( _
        "Sales doc. type Description")
        .Orientation = xlRowField
        .Position = 33
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields("PSP")
        .Orientation = xlRowField
        .Position = 34
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields("WBS RA Key2")
        .Orientation = xlRowField
        .Position = 35
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields("Value Contract")
        .Orientation = xlRowField
        .Position = 36
    End With
    With ActiveSheet.PivotTables(pivName).PivotFields("Commodity3")
        .Orientation = xlRowField
        .Position = 37
    End With
    Range("A6").Select
    ActiveSheet.PivotTables(pivName).PivotFields("Customer Unit"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    Range("B6").Select
    ActiveSheet.PivotTables(pivName).PivotFields("CRG").Subtotals = Array _
        (False, False, False, False, False, False, False, False, False, False, False, False)
    Range("C6").Select
    ActiveSheet.PivotTables(pivName).PivotFields("Country").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    Range("D6").Select
    ActiveSheet.PivotTables(pivName).PivotFields("Assignment ID"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveWorkbook.ShowPivotTableFieldList = False
    Range("E6").Select
    ActiveSheet.PivotTables(pivName).PivotFields( _
        "Assignment ID Description").Subtotals = Array(False, False, False, False, False, _
        False, False, False, False, False, False, False)
    Columns("E:E").ColumnWidth = 23.86
    Range("F6").Select
    ActiveSheet.PivotTables(pivName).PivotFields("Project Definition"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    Range("G6").Select
    ActiveSheet.PivotTables(pivName).PivotFields( _
        "Project Definition Description").Subtotals = Array(False, False, False, False, _
        False, False, False, False, False, False, False, False)
    Range("H6").Select
    ActiveSheet.PivotTables(pivName).PivotFields("WBS Element"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    Range("I6").Select
    ActiveSheet.PivotTables(pivName).PivotFields( _
        "WBS Element Description").Subtotals = Array(False, False, False, False, False, _
        False, False, False, False, False, False, False)
    Range("J6").Select
    ActiveSheet.PivotTables(pivName).PivotFields("PSP Comments"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    Range("K6").Select
    ActiveSheet.PivotTables(pivName).PivotFields("WBS RA Key Description" _
        ).Subtotals = Array(False, False, False, False, False, False, False, False, False, False _
        , False, False)
    Range("L6").Select
    ActiveSheet.PivotTables(pivName).PivotFields("WBS RA Key").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    Range("M6").Select
    ActiveSheet.PivotTables(pivName).PivotFields("FNBL stat chg date"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    Range("N6").Select
    ActiveSheet.PivotTables(pivName).PivotFields("System stat chg date"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables(pivName).PivotFields("WBS System Status"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    Range("P6").Select
    ActiveSheet.PivotTables(pivName).PivotFields("Customer PO"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    Range("Q6").Select
    ActiveSheet.PivotTables(pivName).PivotFields("Sales Order"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    Range("R6").Select
    ActiveSheet.PivotTables(pivName).PivotFields("Billing status"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    Range("S6").Select
    ActiveSheet.PivotTables(pivName).PivotFields("Assignment Execution"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables(pivName).PivotFields("Created on (SO)"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    Range("U6").Select
    ActiveSheet.PivotTables(pivName).PivotFields("Delivery Status"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    Range("V6").Select
    ActiveSheet.PivotTables(pivName).PivotFields( _
        "Delivery Status Description").Subtotals = Array(False, False, False, False, False, _
        False, False, False, False, False, False, False)
    ActiveSheet.PivotTables(pivName).PivotFields("Reason for order"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    Range("X6").Select
    ActiveSheet.PivotTables(pivName).PivotFields( _
        "Reason for order Description").Subtotals = Array(False, False, False, False, False _
        , False, False, False, False, False, False, False)
    Range("Y6").Select
    ActiveSheet.PivotTables(pivName).PivotFields("Project Manager"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    Range("Z6").Select
    ActiveSheet.PivotTables(pivName).PivotFields("ESTA Flag").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables(pivName).PivotFields("Customer group 5"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    Range("AB6").Select
    ActiveSheet.PivotTables(pivName).PivotFields( _
        "Customer group 5 Description").Subtotals = Array(False, False, False, False, False _
        , False, False, False, False, False, False, False)
    Range("AC6").Select
    ActiveSheet.PivotTables(pivName).PivotFields("Changed on").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    Range("AD6").Select
    ActiveSheet.PivotTables(pivName).PivotFields("Governance Stream"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    Range("AE6").Select
    ActiveSheet.PivotTables(pivName).PivotFields("Text LC -> E/// BU"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables(pivName).PivotFields("Sales doc. type"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    Range("AG6").Select
    ActiveSheet.PivotTables(pivName).PivotFields( _
        "Sales doc. type Description").Subtotals = Array(False, False, False, False, False, _
        False, False, False, False, False, False, False)
    Range("AH6").Select
    ActiveSheet.PivotTables(pivName).PivotFields("PSP").Subtotals = Array _
        (False, False, False, False, False, False, False, False, False, False, False, False)
    Range("AI6").Select
    ActiveSheet.PivotTables(pivName).PivotFields("WBS RA Key2"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    Range("AJ6").Select
    ActiveSheet.PivotTables(pivName).PivotFields("Value Contract"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables(pivName).PivotSelect "Commodity3[All]", _
        xlLabelOnly, True
    Range("AK8").Select
    ActiveWorkbook.ShowPivotTableFieldList = True
    ActiveSheet.PivotTables(pivName).AddDataField ActiveSheet.PivotTables _
        (pivName).PivotFields("" & Chr(10) & "Orders Booked (SEK)"), _
        "Sum of " & Chr(10) & "Orders Booked (SEK)", xlSum
    ActiveSheet.PivotTables(pivName).AddDataField ActiveSheet.PivotTables _
        (pivName).PivotFields("" & Chr(10) & "Net Sales (SEK)"), "Sum of " & Chr(10) & "Net Sales (SEK)" _
        , xlSum
    ActiveSheet.PivotTables(pivName).AddDataField ActiveSheet.PivotTables _
        (pivName).PivotFields("" & Chr(10) & "Planned Billing (SEK)"), _
        "Sum of " & Chr(10) & "Planned Billing (SEK)", xlSum
    ActiveSheet.PivotTables(pivName).AddDataField ActiveSheet.PivotTables _
        (pivName).PivotFields("" & Chr(10) & "Actual Billing (SEK)"), _
        "Sum of " & Chr(10) & "Actual Billing (SEK)", xlSum
    ActiveSheet.PivotTables(pivName).AddDataField ActiveSheet.PivotTables _
        (pivName).PivotFields("" & Chr(10) & "Planned Cost (SEK)"), _
        "Sum of " & Chr(10) & "Planned Cost (SEK)", xlSum
    ActiveSheet.PivotTables(pivName).AddDataField ActiveSheet.PivotTables _
        (pivName).PivotFields("" & Chr(10) & "Actual Cost (SEK)"), _
        "Sum of " & Chr(10) & "Actual Cost (SEK)", xlSum
    ActiveSheet.PivotTables(pivName).AddDataField ActiveSheet.PivotTables _
        (pivName).PivotFields("" & Chr(10) & "Committed Cost (SEK)"), _
        "Sum of " & Chr(10) & "Committed Cost (SEK)", xlSum
    ActiveSheet.PivotTables(pivName).AddDataField ActiveSheet.PivotTables _
        (pivName).PivotFields("" & Chr(10) & "Closing Backlog (SEK)"), _
        "Sum of " & Chr(10) & "Closing Backlog (SEK)", xlSum
    ActiveSheet.PivotTables(pivName).AddDataField ActiveSheet.PivotTables _
        (pivName).PivotFields("" & Chr(10) & "Assigned Cost (SEK)"), _
        "Sum of " & Chr(10) & "Assigned Cost (SEK)", xlSum
    ActiveSheet.PivotTables(pivName).AddDataField ActiveSheet.PivotTables _
        (pivName).PivotFields("" & Chr(10) & "COS (SEK)"), "Sum of " & Chr(10) & "COS (SEK)", xlSum
    ActiveSheet.PivotTables(pivName).AddDataField ActiveSheet.PivotTables _
        (pivName).PivotFields("" & Chr(10) & "WIP (SEK)"), "Sum of " & Chr(10) & "WIP (SEK)", xlSum
    ActiveSheet.PivotTables(pivName).AddDataField ActiveSheet.PivotTables _
        (pivName).PivotFields("" & Chr(10) & "RUC (SEK)"), "Sum of " & Chr(10) & "RUC (SEK)", xlSum
    ActiveSheet.PivotTables(pivName).AddDataField ActiveSheet.PivotTables _
        (pivName).PivotFields("" & Chr(10) & "Approved Project Budget (SEK)"), _
        "Sum of " & Chr(10) & "Approved Project Budget (SEK)", xlSum
    ActiveSheet.PivotTables(pivName).AddDataField ActiveSheet.PivotTables _
        (pivName).PivotFields("" & Chr(10) & "Remaining Project Budget (SEK)"), _
        "Sum of " & Chr(10) & "Remaining Project Budget (SEK)", xlSum
    ActiveSheet.PivotTables(pivName).AddDataField ActiveSheet.PivotTables _
        (pivName).PivotFields("" & Chr(10) & "UM (SEK)"), "Sum of " & Chr(10) & "UM (SEK)", xlSum
    With ActiveSheet.PivotTables(pivName).PivotFields("" & Chr(10) & "UM%")
        .Orientation = xlRowField
        .Position = 38
    End With
    ActiveSheet.PivotTables(pivName).AddDataField ActiveSheet.PivotTables _
        (pivName).PivotFields("" & Chr(10) & "Unbilled Sales (SEK)"), _
        "Sum of " & Chr(10) & "Unbilled Sales (SEK)", xlSum
    ActiveSheet.PivotTables(pivName).AddDataField ActiveSheet.PivotTables _
        (pivName).PivotFields("" & Chr(10) & "Deferred Revenue (SEK)"), _
        "Sum of " & Chr(10) & "Deferred Revenue (SEK)", xlSum
    ActiveSheet.PivotTables(pivName).AddDataField ActiveSheet.PivotTables _
        (pivName).PivotFields("" & Chr(10) & "Planned Hours"), "Sum of " & Chr(10) & "Planned Hours", _
        xlSum
    ActiveSheet.PivotTables(pivName).AddDataField ActiveSheet.PivotTables _
        (pivName).PivotFields("" & Chr(10) & "Actual Hours"), "Sum of " & Chr(10) & "Actual Hours", xlSum
    ActiveSheet.PivotTables(pivName).AddDataField ActiveSheet.PivotTables _
        (pivName).PivotFields("" & Chr(10) & "Unapproved Hours"), _
        "Sum of " & Chr(10) & "Unapproved Hours", xlSum
    ActiveWorkbook.ShowPivotTableFieldList = False
    Range("AK6").Select
    ActiveSheet.PivotTables(pivName).PivotFields("Commodity3").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    Range("AL7").Select
    ActiveSheet.PivotTables(pivName).ColumnGrand = False
    ActiveWorkbook.ShowPivotTableFieldList = True
    ActiveSheet.PivotTables(pivName).AddDataField ActiveSheet.PivotTables _
        (pivName).PivotFields("" & Chr(10) & "UM%"), "Count of " & Chr(10) & "UM%", xlCount
    With ActiveSheet.PivotTables(pivName).PivotFields("Count of " & Chr(10) & "UM%")
        .Caption = "Sum of "
        .Function = xlSum
    End With
    ActiveSheet.PivotTables(pivName).PivotFields("Sum of ").Orientation _
        = xlHidden
    ActiveWorkbook.ShowPivotTableFieldList = True
    ActiveWorkbook.ShowPivotTableFieldList = False
    ActiveWindow.Zoom = 85
    Cells.Select
    Range("AD1").Activate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Rows("6:6").Select
    Range("AD6").Activate
    Selection.Font.Bold = True
    Range("AL7").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, -37).Range("A1").Select
    Range(Selection, Cells(6, 37)).Select
    
            Dim selrange As Range
            Set selrange = Selection
            Selection.SpecialCells(xlCellTypeBlanks).Select
            Selection.FormulaR1C1 = "=R[-1]C"
            selrange.Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Application.CutCopyMode = False
    
    Rows("2:4").Select
    Selection.Delete Shift:=xlUp
    Range("B3").Select
    Range("AX3").Select
    Range("A3").Select
    Application.StatusBar = False
End Sub



