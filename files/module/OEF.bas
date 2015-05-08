Attribute VB_Name = "OEF"
Sub OEF_CHC()
Attribute OEF_CHC.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

Dim iRet As Integer
Dim WorkName As String
Dim contactwname As String



      iRet = MsgBox("This Script is currently in BETA, do you wish to continue?", vbYesNo, "BETA Script (UNDER DEVELOPMENT)")
            If iRet = vbNo Then
            End
            Else
            End If
    
    If GetASCFilePath = "" Or GetContactsFilePath = "" Then
        If GetASCFilePath = "" Then
    MsgBox "You need to select the file for your Assignment Status Card (ASC), on the SmartApp Settings"
        End If
        If GetContactsFilePath = "" Then
    MsgBox "You need to select the file for your Contacts, on the SmartApp Settings"
        End If
        Exit Sub
    End If
    
    If GetContactsFilePath = "" Then
    MsgBox "You need to select the file for your Contacts, on the SmartApp Settings"
    Else
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Workbooks.Add
    ActiveSheet.Paste
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "Original"
    Sheets("Sheet2").Select
    ActiveSheet.PasteSpecial Format:="Unicode Text", Link:=False, _
    DisplayAsIcon:=False, NoHTMLFormatting:=True
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    Columns("A:A").ColumnWidth = 41.71
    ActiveCell.FormulaR1C1 = "=MID(RC[1],SEARCH("":"",RC[1],1)+2,LEN(RC[1]))"
    Range("A1").Select
    Selection.AutoFill Destination:=Range("A1:A49")
    Range("A1:A49").Select
    Columns("A:A").Select
    Selection.Copy
    Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=LEFT(RC[1],SEARCH("":"",RC[1],1))"
    Range("A1").Select
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("A1").Select
    ActiveWindow.SmallScroll Down:=-12
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=RIGHT(RC[1],LEN(RC[1])-SEARCH(RC[1],"":"",1))"
    Range("B2").Select
    Selection.Copy
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RIGHT(RC[1],LEN(RC[1])-SEARCH("":"",RC[1],1))"
    Range("B2").Select
    Selection.AutoFill Destination:=Range("B2:B49")
    Range("B2:B49").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B5").Select
    Application.CutCopyMode = False
    ActiveWindow.SmallScroll Down:=-6
    Rows("1:1").Select
    Selection.AutoFilter
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Range("$A$1:$C$49").AutoFilter field:=1, Criteria1:="#VALUE!"
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("$A$1:$C$35").AutoFilter field:=1
    ActiveSheet.Range("$A$1:$C$35").AutoFilter field:=1, Criteria1:="By:"
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("$A$1:$C$33").AutoFilter field:=1, Criteria1:= _
        "Bonds and Guarantees:"
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("$A$1:$C$32").AutoFilter field:=1, Criteria1:= _
        "Contract/Agreement:"
    ActiveSheet.Range("$A$1:$C$32").AutoFilter field:=1
    ActiveSheet.Range("$A$1:$C$32").AutoFilter field:=2, Criteria1:="="
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("$A$1:$C$15").AutoFilter field:=2
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.DisplayGridlines = False
    Rows("2:2").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "Governing Law:"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=RIGHT(R[-1]C,LEN(R[-1]C)-SEARCH("":"",R[-1]C,1)-2)"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=RIGHT(R[-1]C,LEN(R[-1]C)-SEARCH("":"",R[-1]C,1)-1)"
    Rows("2:2").Select
    Selection.Copy
    Rows("4:4").Select
    Selection.Insert Shift:=xlDown
    Range("A4").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Contract Language:"
    Range("B3").Select
    Range("B4").Select
    Range("B4").Select
    Range("B3").Select
    Range("B4").Copy
    Range("B4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B3").Select
    Range("B2").Select
    Range("B2").Copy
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[1],11,3)"
    Range("B1").Select
    Range("B1").Copy
    Range("B1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B3").Select
    ActiveCell.FormulaR1C1 = _
        "=MID(RC[1],SEARCH("":"",RC[1],1)+2,SEARCH(""Contract"",RC[1],1)-SEARCH("":"",RC[1],1))"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = _
        "=MID(RC[1],SEARCH("":"",RC[1],1)+2,SEARCH(""Contract"",RC[1],1)-SEARCH("":"",RC[1],1)-3)"
    Range("B3").Select
    Range("B3").Copy
    Range("B3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("B:B").Select
    Selection.Replace What:=" Contract Administrator:", Replacement:="", _
        LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:= _
        False, ReplaceFormat:=False
    Range("B10").Select
    Columns("A:A").ColumnWidth = 22.29
    Columns("B:B").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B6").Select
    Application.CutCopyMode = False
    Columns("C:C").Select
    Columns("A:A").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("A1").Select
    
    'Create OEF Value Contract on sheet3
    
    Sheets("Sheet3").Select
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A3").Select
    Columns("A:A").ColumnWidth = 17.43
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "ONE Entry Form - Value Contract"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = _
        "Click here for work instruction on how to use the ONE Entry Form"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "Governance stream"
    Range("A5").Select
    ActiveCell.FormulaR1C1 = "Sales track"
    Range("A7").Select
    ActiveCell.FormulaR1C1 = "Partner data"
    Range("A8").Select
    ActiveCell.FormulaR1C1 = "Execution responsible"
    Range("A9").Select
    ActiveCell.FormulaR1C1 = "Contract accountable"
    Range("A10").Select
    ActiveCell.FormulaR1C1 = "Sponsor"
    Range("A11").Select
    ActiveCell.FormulaR1C1 = "PSP"
    Range("B7").Select
    ActiveCell.FormulaR1C1 = "Employee number"
    Range("C7").Select
    ActiveCell.FormulaR1C1 = "Employee name"
    Range("A13").Select
    ActiveCell.FormulaR1C1 = "Fulfillment Assignment (FAS) ID"
    Range("A14").Select
    ActiveCell.FormulaR1C1 = "FAS start date"
    Range("A15").Select
    ActiveCell.FormulaR1C1 = "FAS end date"
    Range("A17").Select
    ActiveCell.FormulaR1C1 = "Value Contract description"
    Range("B17").Select
    ActiveCell.FormulaR1C1 = "Value Contract number from ONE"
    Range("C17").Select
    ActiveCell.FormulaR1C1 = "Contract number on customer side:"
    Range("D17").Select
    ActiveCell.FormulaR1C1 = "Sales " & Chr(10) & "organisation"
    Range("E17").Select
    ActiveCell.FormulaR1C1 = "Sales " & Chr(10) & "office"
    Range("F17").Select
    ActiveCell.FormulaR1C1 = "Sales" & Chr(10) & "group"
    Range("G17").Select
    ActiveCell.FormulaR1C1 = "Sold to party"
    Range("H17").Select
    ActiveCell.FormulaR1C1 = "Ship to party"
    Range("I17").Select
    ActiveCell.FormulaR1C1 = "Transfer to " & Chr(10) & "Global Chronos?"
    Range("J17").Select
    ActiveCell.FormulaR1C1 = "Customer Contract ID (CC ID - CRM360)"
    Range("K17").Select
    ActiveCell.FormulaR1C1 = "Currency"
    Range("K18").Select
    ActiveWindow.DisplayGridlines = False
    Range("A17:K18").Select
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
    Range("A13:B15").Select
    Range("B15").Activate
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
    Range("A7:C11").Select
    Range("C11").Activate
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
    Range("A4:B5").Select
    Range("B5").Activate
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
    With Selection.Font
        .Name = "Arial"
        .Size = 20
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Font
        .Name = "Arial"
        .Size = 19
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Font
        .Name = "Arial"
        .Size = 20
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Range("A4:A5").Select
    Selection.Font.Bold = False
    Selection.Font.Bold = True
    Range("A7:C7").Select
    Selection.Font.Bold = False
    Selection.Font.Bold = True
    Range("A13:A15").Select
    Selection.Font.Bold = False
    Selection.Font.Bold = True
    Range("A17:K17").Select
    Selection.Font.Bold = False
    Selection.Font.Bold = True
    Range("A19").Select
    ActiveCell.FormulaR1C1 = _
        "NOTE: for allowed exceptions, where multiple VC's for one FAS shall be created, please add row(s) to this table by copying and inserting existing row."
    Range("A20").Select
    ActiveCell.FormulaR1C1 = _
        "Allowed exeptions are described in work instructions for Value Contracts, which can be found via below link:"
    Range("A21").Select
    ActiveCell.FormulaR1C1 = "Link to document"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = _
        "Click here for work instruction on how to use the ONE Entry Form"
    Range("A2").Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:= _
        "http://anon.ericsson.se/eridoc/component/eriurl?docno=&objectId=09004cff849b22aa&action=current&format=msw8" _
        , TextToDisplay:= _
        "Click here for work instruction on how to use the ONE Entry Form"
    Range("A5").Select
    Columns("A:A").ColumnWidth = 21.86
    Columns("B:B").ColumnWidth = 19.43
    Columns("C:C").ColumnWidth = 21.43
    Columns("D:D").ColumnWidth = 16.14
    Columns("E:E").ColumnWidth = 10.29
    Columns("F:F").ColumnWidth = 10.14
    Columns("I:I").ColumnWidth = 17.14
    Columns("J:J").ColumnWidth = 32.86
    Columns("K:K").ColumnWidth = 13.14
    Rows("17:17").EntireRow.AutoFit
    Columns("I:I").ColumnWidth = 19.71
    Columns("G:G").ColumnWidth = 15
    Columns("H:H").ColumnWidth = 15.57
    Rows("17:17").EntireRow.AutoFit
    'Range("A8").Select
    'Range("A8").Comment.Text Text:= _
    '    "Enter the responsible person for execution (e.g. CPM / SDM / ASR etc.) to be entered in SAP ONE partner field 'ZP Exec Responsible'."
    'Range("A9").Select
    'Range("A9").Comment.Text Text:= _
    '    "Enter contract accountable (e.g. MSCOO/COM);to be entered in SAP partner field 'ZC Contract accountable'."
    'Range("A10").Select
    'Range("A10").Comment.Text Text:= _
    '    "Sponsor of the Assignment, to be entered in the Partner Field as 'Sponsor'"
    'Range("A11").Select
    'Range("A11").Comment.Text Text:= _
    '    "Name of PSP to be entered in SAP ONE Partner Field 'PSP'"
    'Range("A13").Select
    'Range("A13").Comment.Text Text:= _
    '    "This number should be entered in SAP field 'Assignment ID'"
    'Range("A14").Select
    'Range("A14").Comment.Text Text:= _
    '    "This date should be entered in SAP field 'Contract start'."
    'Range("A15").Select
    'Range("A15").Comment.Text Text:= _
    '    "This date should be entered in SAP field 'Contract end'."
    'Range("B17").Select
    'Range("B17").Comment.Text Text:= _
    '    "Value contract number in ONE (leave blank if still to be created in ONE)."
    'Range("C17").Select
    'Range("C17").Comment.Text Text:= _
    '    "Customer's identification number for the contract. In ONE this number is to be filled in the value contract header field 'PO number'."
    'Range("J17").Select
    'Range("J17").Comment.Text Text:= _
    '    "This number should be entered in SAP field 'CCLM ID'."
    'Range("K17").Select
    'Range("K17").Comment.Text Text:= _
    '    "Choose from the drop down menu the applicable currency for the customer contract."
    Range("A4").Select
    Sheets("Sheet3").Select
    Sheets("Sheet3").Name = "Value Contract"
        Cells.Select
    With Selection.Font
        .Name = "Arial"
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Range("A4").Select
    Columns("A:A").ColumnWidth = 33
    Rows("13:15").Select
    Rows("13:15").EntireRow.AutoFit
    Rows("13:15").EntireRow.AutoFit
    Range("A14").Select
    Rows("13:18").Select
    Selection.RowHeight = 18
    Rows("7:12").Select
    Range("A12").Activate
    Selection.RowHeight = 19.5
    Range("B8:C11").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10092543
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("B13:B15").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10092543
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("A18:K18").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10092543
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("E10").Select
    Columns("J:J").ColumnWidth = 40.71
    Columns("C:C").ColumnWidth = 27.14
    Rows("7:7").Select
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
    Range("A7").Select
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
    Rows("4:5").Select
    Selection.RowHeight = 19.5
    Range("A8").Select
    Columns("B:B").ColumnWidth = 24.57
    Columns("C:C").ColumnWidth = 28
    Rows("7:7").Select
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
    Rows("3:6").Select
    Selection.RowHeight = 17.25
    Range("A7").Select
    Range("A1").Select
    Selection.Font.Bold = True
    Range("A19:A20").Select
    With Selection.Font
        .Name = "Arial"
        .Size = 8
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
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.349986266670736
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.499984740745262
    End With
    Range("A17").Select
    Range("A19").Select
    ActiveCell.FormulaR1C1 = _
        "NOTE: for allowed exceptions, where multiple VC's for one FAS shall be created, please add row(s) to this table by copying and inserting existing row."
    With ActiveCell.Characters(Start:=1, Length:=0).Font
        .Name = "Arial"
        .FontStyle = "Regular"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.499984740745262
        .ThemeFont = xlThemeFontNone
    End With
    With ActiveCell.Characters(Start:=1, Length:=6).Font
        .Name = "Arial"
        .FontStyle = "Bold"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.499984740745262
        .ThemeFont = xlThemeFontNone
    End With
    With ActiveCell.Characters(Start:=7, Length:=144).Font
        .Name = "Arial"
        .FontStyle = "Regular"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.499984740745262
        .ThemeFont = xlThemeFontNone
    End With
    Range("A20").Select
    Range("B4:B5").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10092543
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("B9").Select
    Columns("A:A").ColumnWidth = 32.71
    Rows("1:1").RowHeight = 24
    Range("A2").Select
    Range("A2").Select
    With Selection.Font
        .Name = "Arial"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleSingle
        .ThemeColor = xlThemeColorHyperlink
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Range("A19:A20").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.499984740745262
    End With
        Range("A8").Select
    Rows("17:17").EntireRow.AutoFit
    Range("K18").Select
    ActiveCell.FormulaR1C1 = "=Sheet2!R[-17]C[-9]"
    Range("B4").Select
    ActiveCell.FormulaR1C1 = "Customer Project"
    Range("J18").Select
    ActiveCell.FormulaR1C1 = "=Sheet2!R[-8]C[-8]"
    Range("B4").Select
    Sheets("Value Contract").Select
    Range("G18").Select
    ActiveCell.FormulaR1C1 = "=LEFT(Sheet2!R[-4]C[-5],6)"
    Range("H18").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]"
    Range("D18").Select
    ActiveCell.FormulaR1C1 = "1337"
    Range("E18").Select
    ActiveCell.FormulaR1C1 = "6014"
    Range("F18").Select
    ActiveCell.FormulaR1C1 = "614"
    Range("B4").Select
    ActiveCell.FormulaR1C1 = _
        "=PROPER(MID(Sheet2!R[10]C,SEARCH("":"",Sheet2!R[10]C,1)+2,SEARCH(""="",Sheet2!R[10]C,SEARCH("":"",Sheet2!R[10]C,1))-SEARCH("":"",Sheet2!R[10]C,1)-3))"
    Range("B5").Select
    Sheets("Value Contract").Select
    
    WorkName = ActiveWorkbook.Name
    Range("B8").Select
    Workbooks.Open fileName:=GetContactsFilePath
    contactwname = ActiveWorkbook.Name
    
    Sheets("Contacts").Select
    Sheets("Contacts").Copy After:=Workbooks(WorkName).Sheets(3)
    Sheets("Value Contract").Select
    Range("B9").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(R[9]C[5],'Contacts'!R3C4:R14C22,15,FALSE)"
    Range("B10").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(R[8]C[5],'Contacts'!R3C4:R14C22,19,FALSE)"
    Range("B11").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(R[7]C[5],'Contacts'!R3C4:R14C8,5,FALSE)"
    Range("C8").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(R[10]C[4],'Contacts'!R3C4:R14C24,20,FALSE)"
    Range("C9").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(R[9]C[4],'Contacts'!R3C4:R14C22,14,FALSE)"
    Range("C10").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(R[8]C[4],'Contacts'!R3C4:R14C22,18,FALSE)"
    Range("C11").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(R[7]C[4],'Contacts'!R3C4:R14C8,4,FALSE)"
    Range("B8").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(R[10]C[5],'Contacts'!R3C4:R14C24,21,FALSE)"

    Range("B9").Select
    Range("G18").Select
    ActiveCell.FormulaR1C1 = "=VALUE(MID(Sheet2!R[-4]C[-5],1,7))"
        Cells.Select
    Windows(contactwname).Activate
    ActiveWindow.Close
    'Selection.Copy
    'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'Range("B8").Select
    'Application.CutCopyMode = False
                Question = MsgBox("Have you also copied info from SiteHandler", vbYesNo, "Paste SiteHandler data!")
                If Question = vbNo Then
                Exit Sub
                End If
    
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Paste

Dim o As OLEObject

    For Each o In Application.ActiveSheet.OLEObjects
        If o.progID = "Forms.HTML:Hidden.1" Then
            Debug.Print o.Name, o.Object.value, o.TopLeftCell.Address()
            'sometimes merged cells result from a HTML copy/paste,
            '  so don't just use .TopLeftCell to set the Value
            o.TopLeftCell.MergeArea.value = o.Object.value
            o.Delete
        End If
    Next o

    For Each o In Application.ActiveSheet.OLEObjects
        If o.progID = "Forms.HTML:Text.1" Then
            Debug.Print o.Name, o.Object.value, o.TopLeftCell.Address()
            'sometimes merged cells result from a HTML copy/paste,
            '  so don't just use .TopLeftCell to set the Value
            o.TopLeftCell.MergeArea.value = o.Object.value
            o.Delete
        End If
    Next o
    
    Dim obj As Shape
     
    For Each obj In ActiveWorkbook.ActiveSheet.Shapes
        obj.Delete
    Next
    
    Sheets("Value Contract").Select
    Range("B13").Select
    ActiveCell.FormulaR1C1 = "=Sheet5!R[14]C[6]"
    Range("B4").Select
    
    'CHECK IF THERE ARE OTHER VC WITH SAME ASSIGNMENT
    Range("A19:A21").Select
    Selection.ClearContents
    Range("A18").Select

    
Dim Assing As String
Dim tablename As String
Dim tablerange As Range

    Assing = Sheets(3).Cells(13, 2).value
    tablename = "Table_" & Format(Now(), "ddmmmyy_hhmmss")
        
        
    Range("B19").Select
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:=Array(Array( _
        "ODBC;DSN=Excel Files;DBQ=" & GetASCFilePath & ";DefaultDir=C:\Users\" & Environ("USERNAME") & "\Documents;DriverId=1046;MaxBufferSiz" _
        ), Array("e=2048;PageTimeout=5;")), Destination:=Range("$A$19")).QueryTable
        .CommandText = Array( _
        "SELECT `Data$`.`Value Contract Description`, `Data$`.`Value Contract`, `Data$`.`Assignment ID`, `Data$`.`Sales Organization`, `Data$`.`Sales Office`, `Data$`.`Sales group`, `Data$`.`Sold-to party`, `D" _
        , _
        "ata$`.`Ship-To Party`, `Data$`.`Governance Stream`, `Data$`.`CCLM ID`" & Chr(13) & "" & Chr(10) & "FROM `C:\Users\" & Environ("USERNAME") & "\Documents\CU MTN ASC.xlsx`.`Data$` `Data$`" & Chr(13) & "" & Chr(10) & "WHERE (`Data$`.`Assignment ID`='" & Assing & "')" & Chr(13) & "" & Chr(10) & "ORDER BY `Data$`" _
        , ".`Value Contract`")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = tablename
        .Refresh BackgroundQuery:=False
    End With
    Range("A19").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    ActiveSheet.ListObjects(tablename).Unlink
    Range("C8").Select
    ActiveCell.FormulaR1C1 = "=Sheet5!R[11]C[1]"
    Range("C9").Select
    ActiveCell.FormulaR1C1 = "=Sheet5!R[13]C[5]"
    Range("C10").Select
    ActiveCell.FormulaR1C1 = "=Sheet5!R[11]C[5]"
    Range("A18").Select
    ActiveCell.FormulaR1C1 = "=Sheet5!R[-16]C"
    Range("B8").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[1],Contacts!C[21]:C[22],2,FALSE)"
    Range("B9").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[1],Contacts!C[15]:C[16],2,FALSE)"
    Range("B10").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[1],Contacts!C[17]:C[18],2,FALSE)"
    Range("B19").Select
    ActiveWindow.Zoom = 85
    Range("B4").Select
    Selection.ClearContents
    Range("G18").Select
    ActiveCell.FormulaR1C1 = _
        "=VALUE(PROPER(MID(Sheet2!R[-4]C[-5],SEARCH("":"",Sheet2!R[-4]C[-5],1)+2,SEARCH(""="",Sheet2!R[-4]C[-5],SEARCH("":"",Sheet2!R[-4]C[-5],1))-SEARCH("":"",Sheet2!R[-4]C[-5],1)-3)))"
    Rows("20:27").Select
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    Range("B4").Select

End Sub


