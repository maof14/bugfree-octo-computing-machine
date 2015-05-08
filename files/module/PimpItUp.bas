Attribute VB_Name = "PimpItUp"
Sub PimpItUpScript()

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




Application.ScreenUpdating = False
    Pivsheet = ActiveSheet.Name
    Sheets(Pivsheet).Select
    Sheets(Pivsheet).Name = "Original"
    Sheets("Original").Select
    Sheets("Original").Copy After:=Sheets(1)
    Pivsheet = ActiveSheet.Name
    Sheets(Pivsheet).Select
    Sheets(Pivsheet).Name = "Data"
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
    Range("A1").Select
    ElseIf ActiveCell.Offset(6, 3).value = "Table" And ActiveCell.Offset(7, 3).value = "" Then
    Rows("1:7").Select
    Selection.Delete Shift:=xlUp
    Columns("A:C").Select
    Range("C1").Activate
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    ElseIf ActiveCell.value = "COPA Detail Analysis" Then
    Rows("1:30").Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
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

'If it has 2 row header then move the 2nd row to the top
    Rows("2:2").Select
    Selection.Copy
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Application.CutCopyMode = False
    Range("XFD2").Select
    Selection.End(xlToLeft).Select
    
    'Determine the top RIGHT corner of the table
    toprightcorner = ActiveCell.column
    Range("A2").Select
    Selection.End(xlToRight).Select
    
    'Determine the top LEFT corner of the first level of the header

    topleftcorner = ActiveCell.column
    
    'Clear the contents that are already on the 2nd level of the header
    
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Cells(ActiveCell.Row, 1)).Select
    Selection.ClearContents
    
    'Move the header one row down
    
    Range(Cells(1, topleftcorner), Cells(2, toprightcorner)).Select
    Selection.Cut
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    
    'Select 2nd row header
    Range(Cells(3, 1), Cells(3, toprightcorner)).Select
    
'Application.ScreenUpdating = False

    'Fill in the blanks on HEADER level
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
    Sheets("Original").Select
    With ActiveWorkbook.Sheets("Original").Tab
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.599993896298105
    End With
    Sheets("Pivot").Select
    Progressfrm.Text.Caption = "100% Completion"
    Progressfrm.Bar.Width = 306
    Range("A5").Select
    ActiveWindow.DisplayGridlines = False
        Application.Wait (Now + TimeValue("00:00:02"))
    
    Progressfrm.Hide

End Sub
