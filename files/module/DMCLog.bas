Attribute VB_Name = "DMCLog"
Sub PFMLogScript()
Attribute PFMLogScript.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
Dim lastrow As Long
Dim valtable As Range
Dim pivrange As Range
Dim pivName As String
Dim Pivsheet As String

Application.ScreenUpdating = False
    Workbooks.OpenText fileName:= _
        "\\esekina005\groupfbs\SmartApp\Excel\LOG\scriptruns.log", Origin:= _
        xlWindows, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
        Comma:=False, Space:=False, Other:=True, OtherChar:="|", FieldInfo:= _
        Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7 _
        , 1), Array(8, 1)), TrailingMinusNumbers:=True
        Dim path As String
        path = "C:\Users\" & Environ("Username") & "\Documents\PFM SmartApp"
        If Len(Dir(path, vbDirectory)) = 0 Then
        MkDir (path)
        End If
        ActiveWorkbook.SaveAs (path & "\" & "PFM SmartApp Log " & "_" & Format(Now(), "mmddhhmmss") & ".xlsx"), FileFormat:=51
    Cells.Select
    Cells.EntireColumn.AutoFit
    With Selection.Font
        .Name = "Calibri"
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
    Rows("1:1").Select
    Selection.Font.Bold = True
    Range("A4").Select
    ActiveWindow.Zoom = 85
    ActiveWindow.DisplayGridlines = False
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
    Range("A1").Select
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
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
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
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Columns("B:B").Select
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
    Rows("1:4").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A5").Select
    Columns("A:A").ColumnWidth = 1.57
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "PFM SmartApp Logs"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "=TODAY()"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "=NOW()"
    Range("B3").Select
    Selection.NumberFormat = "[$-409]m/d/yy h:mm AM/PM;@"
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
    Range("B3").Copy
    Range("B3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B2").Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 15
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
    Selection.Font.Bold = True
    Selection.Font.Size = 16
    Range("B5").Select
    Sheets("scriptruns").Select
    Sheets("scriptruns").Name = "Data"
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
    Sheets("Data").Select
    With ActiveWorkbook.Sheets("Data").Tab
        .Color = 255
        .TintAndShade = 0
    End With
    Sheets("Pivot").Select
    ActiveWindow.DisplayGridlines = False
    Range("C8").Select
End Sub
