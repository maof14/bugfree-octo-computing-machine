Attribute VB_Name = "CNS41"
Option Explicit

Dim Exit_Check
Dim Full, path, FileN, RowN, Proj_Color, Sel(10), Num_Opt

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Option Explicit
'////////////////////////////////////////////////////////////////////
'Password masked inputbox
'Allows you to hide characters entered in a VBA Inputbox.
'
'Code written by Daniel Klann
'March 2003
'////////////////////////////////////////////////////////////////////


'API functions to be used
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, _
ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" _
(ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, _
ByVal dwThreadId As Long) As Long

Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Private Declare Function SendDlgItemMessage Lib "user32" Alias "SendDlgItemMessageA" _
(ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wMsg As Long, _
ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, _
ByVal lpClassName As String, _
ByVal nMaxCount As Long) As Long

Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long

'Constants to be used in our API functions
Private Const EM_SETPASSWORDCHAR = &HCC
Private Const WH_CBT = 5
Private Const HCBT_ACTIVATE = 5
Private Const HC_ACTION = 0

Private hHook As Long

    Dim CLPdirName As String
    Dim CLPrptName As String
    Dim CLPsapName As String
    Dim Cancel_Flag






Sub ChangeFormA()

    Dim a, b, c, d, a1, b1, c1, d1_cnf, d1_wbs
    Dim Dum, DirN, i
    
    Application.Calculation = xlCalculationAutomatic

    Sel(0) = "No"
    Sel(1) = "Yes"
    Sel(2) = "Yes"
    Sel(3) = "Yes"
    Sel(4) = "Yes"
  
    
    Exit_Check = False

    a = Cells(1, 1)
    If Left(a, 22) = "No. of Project object:" Then
        FormatChange
    End If
    
    a = Cells(1, 1)
    b = Cells(1, 2)
    c = Cells(1, 3)
    
   
    a1 = "."
    b1 = "Level"
    c1 = "Project object"
    d1_cnf = "Projektelm"
    d1_wbs = "Short text"

    If a = a1 And b = b1 And c = c1 Then
    'Prevent screen redraws until the macro is finished.
        With Application
            .EnableEvents = False
            .ScreenUpdating = False
        End With
        
        
        a = Cells(1, 19)
        
        Application.Run ("DelColumn2")
        
        Application.Run ("Insert_Info")
        
        b = Cells(1, 3)
        If Sel(4) = "Yes" And b = "Projektelm" Then
            Application.Run ("Option_AddGroup")
        End If
        
        If a = "Finish (A)" Then
            Application.Run ("RowTypeB") 'ZNRJCNS41CNF
        Else
            Application.Run ("RowTypeA") 'ZNRJCNS41WBS
        End If
        
        'Application.Run ("Warnings")
        
        Application.Run ("Formatting")
        
        ActiveSheet.Outline.ShowLevels RowLevels:=2
        
        Application.Run ("Warnings")
        Columns("U:AC").Delete
        
        For i = 7 To 100
            a = Cells(1, i)
            If a = "" Then Exit For
        Next i
        Range(Cells(1, 1), Cells(1, i + 1)).AutoFilter
        Range(Cells(1, 1), Cells(1, i + 1)).AutoFilter

        'Application.Run ("Page_Setup")

        With Application
            .EnableEvents = True
            .ScreenUpdating = True
        End With
        
    Else
        If a = b1 And b = c1 And (c = d1_cnf Or c = d1_wbs) Then
            MsgBox ("This sheet is already formatted.")
        Else
            MsgBox ("This sheet does not contain exported CNS41 table.")
            Exit_Check = True
        End If
    End If
    
End Sub



Sub DelColumn() '*** 不要な列を削除 ***
'変数の定義
    Dim i As Integer

    Dim DelClm      '削除する列
    'DelClm = Array(".", "Finish (B)", "Finish (A)", "Pers.resp.")
    DelClm = Array(".", "Finish (B)", "Finish (A)")
        
    Dim DelCLM_JPY  '2列目が「JPY」のとき、削除する列
    DelCLM_JPY = Array("PrRevPl000", "Act. rev.", "Budget", "PrCstSc000", "Act. costs", "TtlCstComm")
    
    Dim DelCLM_H    '2列目が「H」のとき、削除する列
    DelCLM_H = Array("Work", "Work (A)")
        
    Dim var As Variant

    For i = 1 To 256 Step 1     '最大列は256
        If Cells(1, i).value = "" Then Exit For ''空のセルになったら抜ける.
        
        Select Case Cells(2, i).value '2行目の値で分岐
            Case "JPY"
                For Each var In DelCLM_JPY '2列目が「JPY」のとき、削除する列
                    If Cells(1, i).value = var Then
                        Cells(1, i).EntireColumn.Delete
                        i = i - 1
                        Exit For
                    End If
                Next
            Case "KRW"
                For Each var In DelCLM_JPY '2列目が「KRW」のとき、削除する列
                    If Cells(1, i).value = var Then
                        Cells(1, i).EntireColumn.Delete
                        i = i - 1
                        Exit For
                    End If
                Next
            Case "H"
                For Each var In DelCLM_H '2列目が「H」のとき、削除する列
                    If Cells(1, i).value = var Then
                        Cells(1, i).EntireColumn.Delete
                        i = i - 1
                        Exit For
                    End If
                Next
            Case Else
                For Each var In DelClm '削除する列
                    If Cells(1, i).value = var Then
                        Cells(1, i).EntireColumn.Delete
                        i = i - 1
                        Exit For
                    End If
                Next
        End Select
    Next
    
End Sub
Sub DelColumn2()
Dim Dum

Dum = Cells(1, 4)
If Dum = "Projektelm" Then
    Range("A:A,G:G,I:I,K:K,M:M,O:O,Q:S,V:V,X:X").Select
    Range("X1").Activate
    Selection.Delete Shift:=xlToLeft

    Range("A1").Select
End If
If Dum = "Short text" Then
    Range("A:A,G:G,I:I,K:K,M:M,O:O,Q:R,U:U,W:W").Select
    Range("W1").Activate
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft

    Range("A1").Select
End If

End Sub

Sub Insert_Info()
    Dim i, j, a, b, c, Dum, Dum1, dum2
    Dim Level, Project, Prj_Code, CLSD, Budget, PrCst, ActCst, Commit
    Dim Group_St, Group_End, Group_Found, ZNRJCNS41WBS, Color_count, Right_End, ColRange As Range
    Dim Color_Choice, Col_Head, Col_Head_Let, Col_00, Col_01, Col_02, Col_03, Col_04, Col_05, Col_06, Col_FAC
    Dim WBSL3, WBS_Found
    
    Color_Choice = "Summer" 'Summer or Winter
    
    Select Case Color_Choice
        Case "Winter"
            Col_Head = 13
            Col_Head_Let = 2
            Col_00 = 46
            Col_01 = 46
            Col_02 = 45
            Col_03 = 44
            Col_04 = 38
            Col_05 = 40
            Col_FAC = 39
'            Col_WBSL3 = 37
'        Case "Summer"
'            Col_Head = 33
'            Col_Head_Let = 2
'            Col_00 = 4
'            Col_01 = 4
'            Col_02 = 35
'            Col_03 = 33
'            Col_04 = 37
'            Col_05 = 34
'            Col_FAC = 39
'            Col_WBSL3 = 41
        Case "Summer"
            Col_Head = 33
            Col_Head_Let = 2
            Col_00 = 4
            Col_01 = 4
            Col_02 = 35
            Col_03 = 36 '33
            Col_04 = 8
            Col_05 = 34
            Col_06 = 37
            Col_FAC = 39

    End Select
    
    ZNRJCNS41WBS = False

    a = Cells(1, 3)
    If a = "Short text" Then ZNRJCNS41WBS = True
    
    ' Insert New column for Network Number and CLSD status
    Range("D:E").Insert
    If Not ZNRJCNS41WBS Then
        Cells(1, 4) = "NW#"
    End If
    Cells(1, 5) = "Status"
    Columns("D:E").NumberFormat = "General"
    ' Put Formulas
    If Not ZNRJCNS41WBS Then
        With Range("C2")
            'Range(.Cells(1, 1), .End(xlDown)).Offset(25, 12).FormulaR1C1 = "=IF(RC[-12]<>"""",CONCATENATE(RC[-12],""-"",RC[-11]),"""")"
            'Range(.Cells(1, 1), .End(xlDown)).Offset(0, 12).FormulaR1C1 = "=IF(RC[-12]<>"""",CONCATENATE(RC[-12],""-"",RC[-11]),"""")"
'            Range(.Cells(1, 1), .End(xlDown)).Offset(0, 1).FormulaR1C1 = "=IF(R[0]C[-3]=""02"",IF(AND(LEFT(R[1]C[-1],1)=""9"",LEN(R[1]C[-1])=8),R[1]C[-1],IF(AND(LEFT(R[2]C[-1],1)=""9"",LEN(R[2]C[-1])=8),R[2]C[-1],"""")),"""")"
            'Range(.Cells(1, 1), .End(xlDown)).Offset(0, 1).FormulaR1C1 = "=IF(OR(R[0]C[-3]=""02"",R[0]C[-3]=""03""),IF(AND(LEFT(R[1]C[-1],1)=""9"",LEN(R[1]C[-1])=8),R[1]C[-1],IF(AND(LEFT(R[2]C[-1],1)=""9"",LEN(R[2]C[-1])=8),R[2]C[-1],"""")),"""")"
            Range(.Cells(1, 1), .End(xlDown)).Offset(0, 1).FormulaR1C1 = "=IF(OR(R[0]C[-3]=""02"",R[0]C[-3]=""03""),IF(AND(LEFT(R[1]C[-1],1)=""9"",LEN(R[1]C[-1])=8),R[1]C[-1],IF(AND(LEFT(R[2]C[-1],1)=""9"",LEN(R[2]C[-1])=8,code(LEFT(R[1]C[-1],1))<65),R[2]C[-1],"""")),"""")"
            '=IF(OR(A2="02",A2="03"),IF(AND(LEFT(C3,1)="9",LEN(C3)=8),C3,IF(AND(LEFT(C4,1)="9",LEN(C4)=8,ASC(LEFT(C3,1))<65),C4,"")),"")
        End With '=IF(AND(LEFT(D3,1)="9",LEN(D3)=8),D3,"")               ' at D4
                                                                         '=IF(A4="02",IF(AND(LEFT(C5,1)="9",LEN(C5)=8),C5,IF(AND(LEFT(C6,1)="9",LEN(C6)=8),C6,"")),"")
    End If
    With Range("C2")
        'Range(.Cells(1, 1), .End(xlDown)).Offset(0, 2).FormulaR1C1 = "=IF(ISERROR(SEARCH(""CLSD"",RC[1],1)),""open"",""CLSD"")"
        'Range(.Cells(1, 1), .End(xlDown)).Offset(0, 2).FormulaR1C1 = "=IF(NOT(ISERROR(SEARCH(""FAC"",RC[-3],1))),R[-1]C,IF(ISERROR(SEARCH(""CLSD"",RC[1],1)),""open"",""CLSD""))"
        'Range(.Cells(1, 1), .End(xlDown)).Offset(0, 2).FormulaR1C1 = "=IF(NOT(ISERROR(SEARCH(""FAC"",RC[-3],1))),IF(SEARCH(""FAC"",RC[-3],1)=1,R[-1]C,IF(ISERROR(SEARCH(""CLSD"",RC[1],1)),""open"",""CLSD"")),IF(ISERROR(SEARCH(""CLSD"",RC[1],1)),""open"",""CLSD""))"
        Range(.Cells(1, 1), .End(xlDown)).Offset(0, 2).FormulaR1C1 = "=IF(NOT(ISERROR(SEARCH(""FAC"",RC[-3],1))),IF(SEARCH(""FAC"",RC[-3],1)=1,R[-1]C,left(RC[1],4)),left(RC[1],4))"
    End With '=IF(NOT(ISERROR(SEARCH("FAC",B110,1))),E109,IF(ISERROR(SEARCH("CLSD",F110,1)),"open","CLSD"))
    '=IF(NOT(ISERROR(SEARCH("FAC",B2,1))),IF(SEARCH("FAC",B2,1)=1,E1,IF(ISERROR(SEARCH("CLSD",F2,1)),"open","CLSD")),IF(ISERROR(SEARCH("CLSD",F2,1)),"open","CLSD"))
    
    With Range("C2")
        Range(.Cells(1, 1), .End(xlDown)).Offset(0, 14).FormulaR1C1 = "=IF(R[0]C[-10]=0,"""",IF(CODE(LEFT(R[0]C[-14],1))>64,((-1*R[0]C[-10])-R[0]C[-8])/(-1*R[0]C[-10]),""""))"
        Range(.Cells(1, 1), .End(xlDown)).Offset(0, 15).FormulaR1C1 = "=IF(R[0]C[-11]=0,"""",IF(CODE(LEFT(R[0]C[-15],1))>64,((-1*R[0]C[-11])-R[0]C[-8])/(-1*R[0]C[-11]),""""))"
        Range(.Cells(1, 1), .End(xlDown)).Offset(0, 16).FormulaR1C1 = "=IF(R[0]C[-12]=0,"""",IF(CODE(LEFT(R[0]C[-16],1))>64,((-1*R[0]C[-12])-(R[0]C[-8]+R[0]C[-7]))/(-1*R[0]C[-12]),""""))"
    End With '=IF(CODE(LEFT(C2,1))>64,((-1*G2)-I2)/(-1*G2),"")
    
    Range("Q1") = "UM_Budget"
    Range("R1") = "UM_Plan"
    Range("S1") = "UM_Act+Commit"
    
    Columns("Q:S").NumberFormat = "0.0%"
    
    Range("D:E").EntireColumn.AutoFit
    Range("Q:S").EntireColumn.AutoFit
    

    Columns("D:E").Copy
    Columns("D:E").PasteSpecial Paste:=xlPasteValues
    
    Columns("Q:S").Copy
    Columns("Q:S").PasteSpecial Paste:=xlPasteValues
    
    
    For i = 7 To 50
        a = Cells(1, i)
        If a = "" Then
            Right_End = i - 1
            Exit For
        End If
    Next i
    
   
    ' Coloring
    Color_count = 0
    For i = 1 To 65535
        Level = Cells(i, 1)
        Project = Cells(i, 2)
        Prj_Code = Cells(i, 3)
        CLSD = Cells(i, 5)
        Budget = Cells(i, 9)
        PrCst = Cells(i, 10)
        ActCst = Cells(i, 11)
        Commit = Cells(i, 12)
            
        If Level = "" Then Exit For
                
        Set ColRange = Range(Cells(i, 1), Cells(i, Right_End))
'        Set ColRange = Range(Cells(i, 1), Cells(i, 1).End(xlToRight))
                
        If ZNRJCNS41WBS Then
        
            If Level = "Level" Then
                'Range(Cells(1, 1), Cells(1, Right_End)).Interior.ColorIndex = 55
                ColRange.Interior.ColorIndex = 55 '33 ' 13 '54 '33
                ColRange.Font.ColorIndex = 2
            End If
            
            
            If Level = "00" Then
                'WBSL3 = Check_WBSL3(i)
                Color_count = Color_count + 1
                If Color_count Mod 2 = 1 Then
                    ColRange.Interior.ColorIndex = 41 '33 '46
                Else
                    ColRange.Interior.ColorIndex = 50 '10
                End If
            End If
            If Level = "01" Then
                If Color_count Mod 2 = 1 Then
                    ColRange.Interior.ColorIndex = 37 '33 '8 '45
                Else
                    ColRange.Interior.ColorIndex = 35 '42 '50 '34 '4
                End If

            End If
            If Level = "02" Then
                If Color_count Mod 2 = 1 Then
                    ColRange.Interior.ColorIndex = 36 '33 '8 '45
                Else
                    ColRange.Interior.ColorIndex = 36 '42 '50 '34 '4
                End If

            End If
        Else
            If Level = "Level" Then
                ColRange.Interior.ColorIndex = Col_Head '54 '33
                ColRange.Font.ColorIndex = Col_Head_Let
            End If
            
            If Level = "00" Then
                ColRange.Interior.ColorIndex = Col_00
            End If
            If Level = "01" Then
                ColRange.Interior.ColorIndex = Col_01
            End If
            If Level = "02" Then
                ColRange.Interior.ColorIndex = Col_02
            End If
            If Level = "03" Then
                If Prj_Code Like "9*" Then
                    WBSL3 = False
                    ColRange.Interior.ColorIndex = Col_04
                Else
                    WBSL3 = True
                    ColRange.Interior.ColorIndex = Col_03
                End If
            End If
            If Level = "04" Then
                If WBSL3 Then
                    ColRange.Interior.ColorIndex = Col_04
                Else
                    ColRange.Interior.ColorIndex = Col_05
                End If
            End If
            If Level = "05" Then
                If WBSL3 Then
                    ColRange.Interior.ColorIndex = Col_05
                Else
                    ColRange.Interior.ColorIndex = Col_06
                End If
            End If
            If Project Like "FAC*" Or (Prj_Code Like "0*" And Len(Prj_Code) = 12) Then
                ColRange.Interior.ColorIndex = Col_FAC
            End If
            If Level = "06" Then
                ColRange.Interior.ColorIndex = Col_06
            End If
            If i > 2 And Level = "00" Then
                Dum = Cells(i - 1, 1)
'                If Dum <> "Dummy" Then
                If Dum <> "Level" Then
                    Rows(i).Insert
'                    For j = 1 To 15
'                        Cells(i, j) = "Dummy"
'                    Next
                    Rows("1:1").Copy
                    Dum = i & ":" & i
                    Rows(Dum).PasteSpecial Paste:=xlPasteAll
'                    ColRange.Font.ColorIndex = 2 '13 '54 '33
'                    ColRange.Interior.ColorIndex = 13 '54 '33
                    CLSD = ""
                End If
            End If
        End If
        
        If ZNRJCNS41WBS Then
            If CLSD = "CLSD" And Level = "00" Then
                ColRange.Interior.ColorIndex = 16 ' light grey
            End If
            If CLSD = "CLSD" And Level = "01" Then
                ColRange.Interior.ColorIndex = 48 ' light grey
            End If
        Else
            If CLSD = "CLSD" And (Level = "00" Or Level = "01") Then
                ColRange.Interior.ColorIndex = 16 ' light grey
            End If
            If CLSD = "CLSD" And Level = "02" Then
                ColRange.Interior.ColorIndex = 48 ' light grey
            End If
            If CLSD = "CLSD" And (Level = "03" Or Level = "04" Or Level = "05") Then
                ColRange.Interior.ColorIndex = 15 ' light grey
            End If
        End If
    Next

    If ZNRJCNS41WBS Then
    
        ' Grouping (Project)
        Group_Found = False
        For i = 1 To 65535
            Level = Cells(i, 1)
            Project = Cells(i, 2)
            If Level = "" Then Exit For
            
            If Level = "01" Then
                If Group_Found = False Then
                    Group_Found = True
                    Group_St = i
                End If
            End If
            If Level = "00" Then
                If Group_Found = True Then
                    Group_Found = False
                    Group_End = i - 1
        
                    Range(Rows(Group_St), Rows(Group_End)).Select
                    Selection.Rows.Group
                    
                    Group_Found = True
                    Group_St = i + 1
                End If
            End If
        Next i
        If Group_Found = True Then
            Group_Found = False
            Group_End = i
    
            Range(Rows(Group_St), Rows(Group_End)).Select
            Selection.Rows.Group
        End If
    
    Else
'        ' Grouping (WBS-L3)
'        Group_Found = False
'        For i = 1 To 65535
'            Level = Cells(i, 1)
'            Prj_Code = Cells(i, 3)
'            If Level = "" Then Exit For
'
'            If Level = "03" And Not (Prj_Code Like "9*") Then
'                If Group_Found = False Then
'                    Group_Found = True
'                    Group_St = i + 1
'                Else
'                    Group_Found = False
'                    Group_End = i - 1
'
'                    Range(Rows(Group_St), Rows(Group_End)).Select
'                    Selection.Rows.Group
'
'                    Group_Found = True
'                    Group_St = i + 1
'                End If
'            Else
'                If Level = "03" Or Level = "02" Then
'                    If Group_Found = True Then
'                        Group_Found = False
'                        Group_End = i - 1
'
'                        Range(Rows(Group_St), Rows(Group_End)).Select
'                        Selection.Rows.Group
'
'                        Group_Found = True
'                        Group_St = i + 1
'                    End If
'                End If
'            End If
'
''            If Level = "Dummy" Then
'            If Level = "Level" Then
'                If Group_Found = True Then
'                    Group_Found = False
'                    Group_End = i - 1
'
'                    Range(Rows(Group_St), Rows(Group_End)).Select
'                    Selection.Rows.Group
'                End If
'            End If
'        Next i
'        If Group_Found = True Then
'            Group_Found = False
'            Group_End = i - 1
'
'            Range(Rows(Group_St), Rows(Group_End)).Select
'            Selection.Rows.Group
'        End If
        ' Grouping (WBS)
        Group_Found = False
        WBS_Found = False
        Application.Run ("Delete_MS")
        For i = 1 To 65535
            
            Level = Cells(i, 1)
            Prj_Code = Cells(i, 3)
            If Level = "" Then Exit For


            If WBS_Found = True Then
                WBS_Found = False
                If Group_Found = False Then
                    Group_Found = True
                    Group_St = i + 1
                End If
            End If
            If (Level = "03" And Not (Prj_Code Like "9*") And Not (Prj_Code Like "0*")) Or (Level = "02") Then
                If Group_Found = True Then
                    If Group_St < i Then
                        Group_Found = False
                        Group_End = i - 1
            
                        Range(Rows(Group_St), Rows(Group_End)).Select
                        Selection.Rows.Group
                        
                        Group_Found = True
                        Group_St = i + 1
                    Else
                        Group_St = i + 1
                    End If
                End If
            End If
            
'            If Level = "Dummy" Then
            If Level = "Level" Then
                If Group_Found = True Then
                    Group_Found = False
                    Group_End = i - 1
        
                    Range(Rows(Group_St), Rows(Group_End)).Select
                    Selection.Rows.Group
                End If
            End If
            
            If Level = "01" Then
                    WBS_Found = True
            End If
            
            
        Next i
        If Group_Found = True Then
            Group_Found = False
            Group_End = i - 1
    
            Range(Rows(Group_St), Rows(Group_End)).Select
            Selection.Rows.Group
        End If
    
        ' Grouping (Project)
        Group_Found = False
        For i = 1 To 65535
            Level = Cells(i, 1)
            Project = Cells(i, 2)
            If Level = "" Then Exit For
            
            If Level = "01" Then
                If Group_Found = False Then
                    Group_Found = True
                    'Group_St = i + 1
                    Group_St = i
                End If
            End If
            If Level = "00" Then
                If Group_Found = True Then
                    Group_Found = False
                    Group_End = i - 1
        
                    Range(Rows(Group_St), Rows(Group_End)).Select
                    Selection.Rows.Group
                    
                    Group_Found = True
                    Group_St = i + 1
                End If
            End If
            
'            If Level = "Dummy" Then
            If Level = "Level" Then
                If Group_Found = True Then
                    Group_Found = False
                    Group_End = i
        
                    Range(Rows(Group_St), Rows(Group_End)).Select
                    Selection.Rows.Group
                End If
            End If
        Next i
        If Group_Found = True Then
            Group_Found = False
            Group_End = i
    
            Range(Rows(Group_St), Rows(Group_End)).Select
            Selection.Rows.Group
        End If
    End If

    ' Hide Original Status
    If ZNRJCNS41WBS Then
        Columns("D:F").Select
        Selection.Columns.Group
        
    Else
        Columns(6).Select
        Selection.Columns.Group
        ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
    End If

    ' Insert Actual + Commitment
    Range("M:M").Insert
    Range("M1") = "Act+Commit"
    With Range("C2")
        Range(.Cells(1, 1), .End(xlDown)).Offset(0, 10).FormulaR1C1 = "=IF(not(isnumber(R[0]C[-1])),""Act+Commit"",R[0]C[-2]+R[0]C[-1])"
    End With '=IF(NOT(ISNUMBER(L2)),"Act+Commit",K2+L2)
    
    Columns("M:M").Copy
    Columns("M:M").PasteSpecial Paste:=xlPasteValues
    
    Columns(13).AutoFit


End Sub
Sub Warnings()
    Dim i, j, a, b, c, Dum, Dum1, dum2
    Dim Level, Project, Prj_Code, Network, CLSD, Budget, PrCst, ActCst, Commit, Plan_Rev, Act_Rev, Act_Com, status
    Dim Level_col, Project_col, CLSD_col, Status_col, Budget_col, PrCst_col, ActCst_col, Commit_Col, Plan_Rev_col, Act_Rev_col, Act_Com_col
    
    Dim ZNRJCNS41WBS
    Dim Pl_overBud, Pl_less85, Pl_lessAct, Act_overBud, Assign_overPlBu, Rev_Rem, TECO, Issue, CNF, Issue_col
    Dim Plan_Low_Act_Col, Plan_Low_Ass_Col, Plan_Over_Budget_Col, Plan_less85_Col, Ass_Over_Budget_Col, WBSL2_3_Col, CNF_Col
    Dim TECO_Col, Rev_Rem_Col, Com2_Col



'    Open Full For Input As #1
'        Input #1, Proj_Color
'    Close #1
    Proj_Color = Sel(0)

    a = Cells(1, 3)
    If a = "Short text" Then ZNRJCNS41WBS = True
    If ZNRJCNS41WBS Then
        Level_col = 1
        Project_col = 2
        CLSD_col = 5
        Status_col = 6
        Plan_Rev_col = 7
        Act_Rev_col = 8
        Budget_col = 9
        PrCst_col = 10
        ActCst_col = 11
        Commit_Col = 12
        Act_Com_col = 13
    
        Issue_col = 30
        WBSL2_3_Col = 31
        Plan_Low_Act_Col = 32
        Plan_Low_Ass_Col = 33
        Plan_Over_Budget_Col = 34
        Plan_less85_Col = 35
        Ass_Over_Budget_Col = 36
        Rev_Rem_Col = 37
        Com2_Col = 38
        TECO_Col = 39
    
        Cells(1, Issue_col) = "Issue"
        Cells(1, WBSL2_3_Col) = "WBSL2/3"
        Cells(1, Plan_Low_Act_Col) = "Pl<Act"
        Cells(1, Plan_Low_Ass_Col) = "Pl<Ass"
        Cells(1, Plan_less85_Col) = "Pl<85%"
        Cells(1, Plan_Over_Budget_Col) = "Pl>Bud"
        Cells(1, Ass_Over_Budget_Col) = "Ass>Bud"
        Cells(1, Rev_Rem_Col) = "Rev_Rem"
        Cells(1, Com2_Col) = "Commit"
        Cells(1, TECO_Col) = "TECO"
    
    Else
        Level_col = 1
        Project_col = 3
        CLSD_col = 5
        Status_col = 6
        Plan_Rev_col = 7
        Act_Rev_col = 8
        Budget_col = 9
        PrCst_col = 10
        ActCst_col = 11
        Commit_Col = 12
        Act_Com_col = 13

        Issue_col = 30
        
        WBSL2_3_Col = 31
        Plan_Low_Act_Col = 32
        Plan_Low_Ass_Col = 33
        Plan_Over_Budget_Col = 34
        Plan_less85_Col = 35
        Ass_Over_Budget_Col = 36
        Rev_Rem_Col = 37
        Com2_Col = 38
        TECO_Col = 39
        CNF_Col = 40
        
    Cells(1, Issue_col) = "Issue"
    Cells(1, WBSL2_3_Col) = "WBSL2/3"
    Cells(1, Plan_Low_Act_Col) = "Pl<Act"
    Cells(1, Plan_Low_Ass_Col) = "Pl<Ass"
    Cells(1, Plan_less85_Col) = "Pl<85%"
    Cells(1, Plan_Over_Budget_Col) = "Pl>Bud"
    Cells(1, Ass_Over_Budget_Col) = "Ass>Bud"
    Cells(1, Rev_Rem_Col) = "Rev_Rem"
    Cells(1, Com2_Col) = "Commit"
    Cells(1, TECO_Col) = "TECO"
    Cells(1, CNF_Col) = "CNF"

    End If
    

    For i = 1 To 65535
        
        Issue = False
                
        Level = Cells(i, Level_col)
        Project = Cells(i, Project_col)
        CLSD = Cells(i, CLSD_col)
        status = Cells(i, Status_col)
        Plan_Rev = Cells(i, Plan_Rev_col)
        Act_Rev = Cells(i, Act_Rev_col)
        Budget = Cells(i, Budget_col)
        PrCst = Cells(i, PrCst_col)
        ActCst = Cells(i, ActCst_col)
        Commit = Cells(i, Commit_Col)
        Act_Com = Cells(i, Act_Com_col)
            
        If Level = "" Then Exit For
        
        If Level = "00" Or Level = "01" Then    'Project Def
            If Commit > 0 Then
                Cells(i, Commit_Col).Interior.ColorIndex = 6
            End If
        End If
        Pl_overBud = False
        Pl_less85 = False
        Pl_lessAct = False
        Act_overBud = False
        Assign_overPlBu = False
        Rev_Rem = False
        TECO = False
        CNF = False
        
        Dum = ""
        Dum1 = ""
        
        If ZNRJCNS41WBS And CLSD <> "CLSD" Then
            If Level = "01" Or Level = "02" Then                    'WBS
                If Commit > 0 Then
                    Cells(i, Commit_Col).Interior.ColorIndex = 6
                    
                    If ActCst + Commit > PrCst Then
                        Cells(i, Commit_Col).Interior.ColorIndex = 8
                    
                        With Cells(i, Commit_Col)
                            .AddComment
                            With .Comment
                                .Visible = True
                                '.Text Text:="Actual+Commitment > Plan Cost" & Chr(10) & Format(ActCst, "##,##0") & "+" & Format(Commit, "##,##0") & "=" & Format(ActCst + Commit, "##,##0") & ">" & Format(PrCst, "##,##0")
                                .Text Text:="Plan Cost < Actual+Commitment" & Chr(10) & Format(PrCst, "##,##0") & "<" & Format(ActCst + Commit, "##,##0") & "=" & Format(ActCst, "##,##0") & "+" & Format(Commit, "##,##0")
                                .Shape.TextFrame.AutoSize = True
                                .Visible = False
                            End With
                        End With
                    
                    End If
                
                End If
                If ActCst > Budget Then
                    Act_overBud = True
                    Cells(i, ActCst_col).Interior.ColorIndex = 3
                    With Cells(i, ActCst_col)
                        .AddComment
                        With .Comment
                            .Visible = True
                            .Text Text:="Actual Cost > Budget" & Chr(10) & "(" & Format(ActCst - Budget, "##,##0") & " Over)" & Chr(10) & "You may need to issue CR."
                            .Shape.TextFrame.AutoSize = True
                            .Visible = False
                        End With
                    End With
                End If
                
                If PrCst > Budget Then
                    Pl_overBud = True
                    Cells(i, PrCst_col).Interior.ColorIndex = 3
                    If Budget <> 0 Then
                        Dum = "Plan Cost (" & Format(PrCst / Budget, "0.0%") & ") > Budget." '& Chr(10) & "Plan Cost should be < " & Format(Budget, "##,##0")
                    Else
                        Dum = "Plan Cost > Budget. (Budget=0)" '& Chr(10) & "Plan Cost should be < " & Format(Budget, "##,##0")
                    End If
                End If
                If PrCst < 0.85 * Budget And Budget <> 0 Then
                    Pl_less85 = True
                    Cells(i, PrCst_col).Interior.ColorIndex = 6
                    Dum = "Plan Cost (" & Format(PrCst / Budget, "0.0%") & ") < 85% of Budget." '& Chr(10) & "Plan Cost should be > " & Format(0.85 * Budget, "##,##0")
                End If
                If ActCst > PrCst Then
                    Pl_lessAct = True
                    Cells(i, PrCst_col).Interior.ColorIndex = 7
                    If Dum <> "" Then Dum = Dum & Chr(10)
                    If ActCst < 0.85 * Budget Then
                        Dum = Dum & "Plan Cost < Actual Cost." '& Chr(10) & "Plan Cost should be > " & Format(0.85 * Budget, "##,##0")
                    Else
                        Dum = Dum & "Plan Cost < Actual Cost." '& Chr(10) & "Plan Cost should be > " & Format(ActCst, "##,##0")
                    End If
                End If
                
                If ActCst < 0.85 * Budget Then Dum1 = "Plan Cost should be:" & Chr(10) & Format(0.85 * Budget, "##,##0") & " < Plan Cost =< " & Format(Budget, "##,##0")
                If ActCst > 0.85 * Budget And ActCst < Budget Then Dum1 = "Plan Cost should be:" & Chr(10) & Format(ActCst, "##,##0") & " =< Plan Cost =< " & Format(Budget, "##,##0")
                If ActCst > Budget Then Dum1 = "Plan Cost should be >= " & Format(ActCst, "##,##0")
                
                If Dum <> "" Then
                    With Cells(i, PrCst_col)
                        .AddComment
                        With .Comment
                            .Visible = True
                            .Text Text:=Dum & Chr(10) & Dum1
                            .Shape.TextFrame.AutoSize = True
                            .Visible = False
                        End With
                    End With
                End If
            
                If Abs(Plan_Rev) > Abs(Act_Rev) Then
                    Rev_Rem = True
                    Cells(i, Act_Rev_col).Interior.ColorIndex = 38
                    With Cells(i, Act_Rev_col)
                        .AddComment
                        With .Comment
                            .Visible = True
                            .Text Text:="Revenue remaining : " & Chr(10) & Format(Abs(Plan_Rev) - Abs(Act_Rev), "##,##0")
                            .Shape.TextFrame.AutoSize = True
                            .Visible = False
                        End With
                    End With
                End If
                '--Actual + Commitment
                Dum = ""
                If Act_Com > PrCst And Act_Com < Budget Then
                    Assign_overPlBu = False
                    Cells(i, Act_Com_col).Interior.ColorIndex = 8
                    Dum = "Actual Cost + Commitment > Plan Cost" & Chr(10) & "Actual Cost + Commitment < Budget"
                End If
                If Act_Com < PrCst And Act_Com > Budget Then
                    Cells(i, Act_Com_col).Interior.ColorIndex = 45
                    Dum = "Actual Cost + Commitment < Plan Cost" & Chr(10) & "Actual Cost + Commitment > Budget"
                End If
                If Act_Com > PrCst And Act_Com > Budget Then
                    Cells(i, Act_Com_col).Interior.ColorIndex = 3
                    Dum = "Actual Cost + Commitment > Plan Cost" & Chr(10) & "Actual Cost + Commitment > Budget"
                End If
                If Dum <> "" Then
                    With Cells(i, Act_Com_col)
                        .AddComment
                        With .Comment
                            .Visible = True
                            .Text Text:=Dum
                            .Shape.TextFrame.AutoSize = True
                            .Visible = False
                        End With
                    End With
                End If
                
            End If
            If Level <> "Level" And Level <> "00" Then
                If Not (IsNumeric(Left(Project, 1))) Then
                    Cells(i, WBSL2_3_Col) = 1
                    Issue = Assign_overPlBu Or Act_overBud Or Pl_overBud Or Pl_less85 Or Pl_lessAct
                    If Issue Then Cells(i, Issue_col) = 1
                
                    If PrCst < ActCst Then Cells(i, Plan_Low_Act_Col) = 1
                    If PrCst < Act_Com Then Cells(i, Plan_Low_Ass_Col) = 1
                    If PrCst > Budget Then Cells(i, Plan_Over_Budget_Col) = 1
                    If PrCst < 0.85 * Budget Then Cells(i, Plan_less85_Col) = 1
                    If Act_Com > Budget Then Cells(i, Ass_Over_Budget_Col) = 1
                    If Rev_Rem Then Cells(i, Rev_Rem_Col) = 1
                    If Commit > 0 Then Cells(i, Com2_Col) = 1
                    If Left(status, 4) = "TECO" Then TECO = True
                    If TECO Then Cells(i, TECO_Col) = 1
                                                            
                End If
            End If

        Else
            If ((Level = "00" And Proj_Color = "Yes") Or Level = "02" Or (Level = "03" And Not (Project Like "9*"))) And CLSD <> "CLSD" Then             'CNF
                If Commit > 0 Then
                    Cells(i, Commit_Col).Interior.ColorIndex = 6
                    If ActCst + Commit > PrCst Then
                        Cells(i, Commit_Col).Interior.ColorIndex = 8
                    
                        With Cells(i, Commit_Col)
                            .AddComment
                            With .Comment
                                .Visible = True
                                '.Text Text:="Actual+Commitment > Plan Cost" & Chr(10) & Format(ActCst, "##,##0") & "+" & Format(Commit, "##,##0") & "=" & Format(ActCst + Commit, "##,##0") & ">" & Format(PrCst, "##,##0")
                                .Text Text:="Plan Cost < Actual+Commitment" & Chr(10) & Format(PrCst, "##,##0") & "<" & Format(ActCst + Commit, "##,##0") & "=" & Format(ActCst, "##,##0") & "+" & Format(Commit, "##,##0")
                                .Shape.TextFrame.AutoSize = True
                                .Visible = False
                            End With
                        End With
                    
                    End If
                End If
                If ActCst > Budget Then
                    Act_overBud = True
                    Cells(i, ActCst_col).Interior.ColorIndex = 3
                    With Cells(i, ActCst_col)
                        .AddComment
                        With .Comment
                            .Visible = True
                            .Text Text:="Actual Cost > Budget" & Chr(10) & "(" & Format(ActCst - Budget, "##,##0") & " Over)" & Chr(10) & "You may need to issue CR."
                            .Shape.TextFrame.AutoSize = True
                            .Visible = False
                        End With
                    End With
                End If
                If PrCst > Budget Then
                    Pl_overBud = True
                    Cells(i, PrCst_col).Interior.ColorIndex = 3
                    If Budget <> 0 Then
                        Dum = "Plan Cost (" & Format(PrCst / Budget, "0.0%") & ") > Budget."  '& Chr(10) & "Plan Cost should be < " & Format(Budget, "##,##0")
                    Else
                        Dum = "Plan Cost > Budget. (Budget=0)" '& Chr(10) & "Plan Cost should be < " & Format(Budget, "##,##0")
                    End If
                End If
                If PrCst < 0.85 * Budget And Budget <> 0 Then
                    Pl_less85 = True
                    Cells(i, PrCst_col).Interior.ColorIndex = 6
                    Dum = "Plan Cost (" & Format(PrCst / Budget, "0.0%") & ") < 85% of Budget." '& Chr(10) & "Plan Cost should be > " & Format(0.85 * Budget, "##,##0")
                End If
                If ActCst > PrCst Then
                    Pl_lessAct = True
                    Cells(i, PrCst_col).Interior.ColorIndex = 7
                    If Dum <> "" Then Dum = Dum & Chr(10)
                    If ActCst < 0.85 * Budget Then
                        Dum = Dum & "Plan Cost < Actual Cost." '& Chr(10) & "Plan Cost should be > " & Format(0.85 * Budget, "##,##0")
                    Else
                        Dum = Dum & "Plan Cost < Actual Cost." '& Chr(10) & "Plan Cost should be > " & Format(ActCst, "##,##0")
                    End If
                End If
                
                If ActCst < 0.85 * Budget Then Dum1 = "Plan Cost should be:" & Chr(10) & Format(0.85 * Budget, "##,##0") & " < Plan Cost =< " & Format(Budget, "##,##0")
                If ActCst > 0.85 * Budget And ActCst < Budget Then Dum1 = "Plan Cost should be:" & Chr(10) & Format(ActCst, "##,##0") & " =< Plan Cost =< " & Format(Budget, "##,##0")
                If ActCst > Budget Then Dum1 = "Plan Cost should be > " & Format(ActCst, "##,##0")
                If ActCst > Budget And PrCst > Budget Then
                    Dum = "Actual Cost and Plan cost > Budget"
                    Dum1 = "Please issue CR to increase Budget."
                End If
                If Dum <> "" Then
                    With Cells(i, PrCst_col)
                        .AddComment
                        With .Comment
                            .Visible = True
                            .Text Text:=Dum & Chr(10) & Dum1
                            .Shape.TextFrame.AutoSize = True
                            .Visible = False
                        End With
                    End With
                End If
                 If Abs(Plan_Rev) > Abs(Act_Rev) Then
                    Rev_Rem = True
                    Cells(i, Act_Rev_col).Interior.ColorIndex = 38
                    With Cells(i, Act_Rev_col)
                        .AddComment
                        With .Comment
                            .Visible = True
                            .Text Text:="Revenue remaining : " & Chr(10) & Format(Abs(Plan_Rev) - Abs(Act_Rev), "##,##0")
                            .Shape.TextFrame.AutoSize = True
                            .Visible = False
                        End With
                    End With
                End If
                '--Actual + Commitment
                Dum = ""
                If Act_Com > PrCst And Act_Com < Budget Then
                    Assign_overPlBu = True
                    Cells(i, Act_Com_col).Interior.ColorIndex = 8
                    Dum = "Actual Cost + Commitment > Plan Cost" & Chr(10) & "Actual Cost + Commitment < Budget"
                End If
                If Act_Com < PrCst And Act_Com > Budget Then
                    Assign_overPlBu = True
                    Cells(i, Act_Com_col).Interior.ColorIndex = 45
                    Dum = "Actual Cost + Commitment < Plan Cost" & Chr(10) & "Actual Cost + Commitment > Budget"
                End If
                If Act_Com > PrCst And Act_Com > Budget Then
                    Assign_overPlBu = True
                    Cells(i, Act_Com_col).Interior.ColorIndex = 3
                    Dum = "Actual Cost + Commitment > Plan Cost" & Chr(10) & "Actual Cost + Commitment > Budget"
                End If
                If Dum <> "" Then
                    With Cells(i, Act_Com_col)
                        .AddComment
                        With .Comment
                            .Visible = True
                            .Text Text:=Dum
                            .Shape.TextFrame.AutoSize = True
                            .Visible = False
                        End With
                    End With
                End If
           
            End If
            
            If Level <> "Level" And Level <> "00" And Level <> "01" Then
                If Not (IsNumeric(Left(Project, 1))) Then
                    Cells(i, WBSL2_3_Col) = 1
                    Issue = Assign_overPlBu Or Act_overBud Or Pl_overBud Or Pl_less85 Or Pl_lessAct
                    If Issue Then Cells(i, Issue_col) = 1
                
                    If PrCst < ActCst Then Cells(i, Plan_Low_Act_Col) = 1
                    If PrCst < Act_Com Then Cells(i, Plan_Low_Ass_Col) = 1
                    If PrCst > Budget Then Cells(i, Plan_Over_Budget_Col) = 1
                    If PrCst < 0.85 * Budget Then Cells(i, Plan_less85_Col) = 1
                    If Act_Com > Budget Then Cells(i, Ass_Over_Budget_Col) = 1
                    If Rev_Rem Then Cells(i, Rev_Rem_Col) = 1
                    If Commit > 0 Then Cells(i, Com2_Col) = 1
                    If Left(status, 4) = "TECO" Then TECO = True
                    If TECO Then Cells(i, TECO_Col) = 1
                    
                    For j = 1 To 5
                    
                        Dum = Cells(i + j, Status_col)
                        If Left(Dum, 3) = "CNF" Then CNF = True
                        
                    Next
                                        
                    If CNF Then Cells(i, CNF_Col) = 1
                                        
                End If
            End If
            
            
            If Level = "03" And CLSD <> "CLSD" Then                     'Network
                If Commit > 0 Then
                    Cells(i, Commit_Col).Interior.ColorIndex = 6
                    If ActCst + Commit > PrCst Then
                        Cells(i, Commit_Col).Interior.ColorIndex = 8
                    End If
                End If
'                '--Actual + Commitment
'                Dum = ""
'                If Act_Com > PrCst And Act_Com < Budget Then
'                    Cells(i, Act_Com_col).Interior.ColorIndex = 8
'                    Dum = "Actual Cost + Commitment > Plan Cost" & Chr(10) & "Actual Cost + Commitment < Budget"
'                End If
'                If Act_Com < PrCst And Act_Com > Budget Then
'                    Cells(i, Act_Com_col).Interior.ColorIndex = 45
'                    Dum = "Actual Cost + Commitment < Plan Cost" & Chr(10) & "Actual Cost + Commitment > Budget"
'                End If
'                If Act_Com > PrCst And Act_Com > Budget Then
'                    Cells(i, Act_Com_col).Interior.ColorIndex = 3
'                    Dum = "Actual Cost + Commitment > Plan Cost" & Chr(10) & "Actual Cost + Commitment > Budget"
'                End If
'                If Dum <> "" Then
'                    With Cells(i, Act_Com_col)
'                        .AddComment
'                        With .Comment
'                            .Visible = True
'                            .Text Text:=Dum
'                            .Shape.TextFrame.AutoSize = True
'                            .Visible = False
'                        End With
'                    End With
'                End If
            End If
            If Level = "04" And CLSD <> "CLSD" Then                     'Activity
                If Commit > 0 Then
                    Cells(i, Commit_Col).Interior.ColorIndex = 36
                End If
            End If
            If Level = "05" And CLSD <> "CLSD" Then                    'Sub Activity
                If Commit > 0 Then
                    Cells(i, Commit_Col).Interior.ColorIndex = 36
                End If
            End If
        End If
    Next

End Sub
Sub Formatting()
    Dim i, a
'書式設定
  '定数(数値)が含まれているセルの書式　「0」は「-」と表示　（Category;Acounting,  Decimal places;0, Symbol;none）-> Just custom
   ' Cells.Select
    'Cells.SpecialCells(xlCellTypeConstants, xlNumbers).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
    'Cells.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* """"-""""_);_(@_)"
    Range("G:O").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* """"-""""_);_(@_)"


'オートフィルタ
'    If Not ActiveSheet.AutoFilterMode Then
'      Rows("1:1").AutoFilter
'    End If
    If Not ActiveSheet.AutoFilterMode Then
        For i = 7 To 100
            a = Cells(1, i)
            If a = "" Then Exit For
        Next i
        Range(Cells(1, 1), Cells(1, i + 1)).AutoFilter
    End If
        
'ウィンドウ枠の固定
    Range("D2").Select
    ActiveWindow.FreezePanes = True
End Sub

Sub RowTypeA()  'TypeA
'変数の定義
    Dim j As Long, Level, Prj_Code

    For j = 2 To 65536 Step 1     '2列目から最大行(65536)まで
        Level = Cells(j, 1)
        Prj_Code = Cells(j, 2)
        If Cells(j, 1).value = "" Then Exit For ''空のセルになったら抜ける
        
        If Prj_Code Like "9*" Then
            Cells(j, 1).EntireRow.Delete
            j = j - 1
        End If

'        Select Case Cells(j, 1).Value  'Levelのセルの値ごとに処理
'            Case "00"                  '"00"の行はBoldに
'                'Cells(j, 1).EntireRow.Font.Bold = True
'            Case "02"                  '"02"の行は削除
'                Cells(j, 1).EntireRow.Delete
'                j = j - 1
'        End Select
    Next
End Sub

Sub RowTypeB()  'TypeB
'変数の定義
    Dim j As Long
    
    For j = 2 To 65536 Step 1     '2列目から最大行(65536)まで
        If Cells(j, 1).value = "" Then Exit For ''空のセルになったら抜ける

        Select Case Cells(j, 1).value  'Levelのセルの値ごとに処理
            Case "02", "03", "04"                '"02", "03", "04"の行はBoldに
                'Cells(j, 1).EntireRow.Font.Bold = True
        End Select
    Next
End Sub

Sub Page_Setup()
    
    
    ' Page Setup "Sheet"
'    With ActiveSheet.PageSetup
'        .PrintTitleRows = "$2:$2"
'        .PrintTitleColumns = ""
'    End With
    
    With ActiveSheet.PageSetup
        ' Page Setup "Header/Footer"
'        .LeftHeader = ""
'        .CenterHeader = "&F"
'        .RightHeader = ""
'        .LeftFooter = ""
'        .CenterFooter = "&P ページ"
'        .RightFooter = ""

        ' Page Setup "Margin"
'        .LeftMargin = Application.InchesToPoints(0.7874)
'        .RightMargin = Application.InchesToPoints(0.7874)
'        .TopMargin = Application.InchesToPoints(0.984251)
'        .BottomMargin = Application.InchesToPoints(0.984251)
'        .HeaderMargin = Application.InchesToPoints(0.511811)
'        .FooterMargin = Application.InchesToPoints(0.511811)
        
        ' Page Setup "Page"
'        .PrintHeadings = False
'        .PrintGridlines = True
'        .PrintComments = xlPrintNoComments
'        .PrintQuality = -3
'        .CenterHorizontally = False
'        .CenterVertically = False
        
        .Orientation = xlLandscape
'        .Draft = True
        .PaperSize = xlPaperA4
'        .FirstPageNumber = xlAutomatic
        
'        .Order = xlOverThenDown
'        .BlackAndWhite = True
        ' Select either Zoom or FitToPages
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    
    End With
    
    ' Page Setup "Print Area"
    'ActiveSheet.PageSetup.PrintArea = "$A$2:$G$200"

End Sub
Sub Check_WBSL3(Line_N)
'変数の定義
    Dim j As Long, Dum, a, b
    a = Cells(Line_N, 1)
    b = Cells(Line_N, 3)
    For j = Line_N To 65536 Step 1     '2列目から最大行(65536)まで
        If Cells(j, 1).value = "" Then Exit For ''空のセルになったら抜ける

        Select Case Cells(j, 1).value  'Levelのセルの値ごとに処理
            Case "02", "03", "04"                '"02", "03", "04"の行はBoldに
                'Cells(j, 1).EntireRow.Font.Bold = True
        End Select
    Next
End Sub


Function FormatChange()
    
    Dim DBPath As String, S_Select As String, S_Connection As String, S1 As String, S2 As String, S3 As String
    Dim a, b, i As Integer
    Dim WBS_N As String, WBS_Sel As String
    Dim ws_MAIN As Worksheet, ws_Prj As Worksheet, ws_SUB As Worksheet, ws_Chart As Worksheet, ws As Worksheet
    Dim Dum, H_level
    
    'Set ws_headers = ActiveWorkbook.Sheets("Headers")
    
        Set ws = ThisWorkbook.ActiveSheet
        
            For i = 0 To 10
                Dum = Cells(3, 4 + i)
                If Dum <> "" Then
                    H_level = i
                    Exit For
                End If
            Next
            a = "=""0""&"
            b = "=CONCATENATE("
            For i = 0 To H_level
                If i > 0 Then
                    a = a & "+"
                    b = b & ","
                End If
            a = a & i & "*NOT(ISBLANK(RC[" & i + 2 & "]))"
            b = b & "RC[" & i + 1 & "]"
            Next i
            b = b & ")"
           
            'ws.Select
            Columns("C:C").Insert Shift:=xlToRight


            With Range("A5")
            
                Range(.Cells(1, 1), .End(xlDown)).Offset(0, 1).FormulaR1C1 = a
                Range(.Cells(1, 1), .End(xlDown)).Offset(0, 2).FormulaR1C1 = b
            
            End With
            
            Columns("B:C").Copy
            Columns("B:C").PasteSpecial Paste:=xlPasteValues
            
            Dum = Chr(68 + H_level)
            
            Columns("D:" & Dum).Delete
            
            'ws_headers.Range("2:2").Copy Destination:=ws.Range("1:1")
            Cells(1, 1) = "."
            Cells(1, 2) = "Level"
            Cells(1, 3) = "Project object"
            Cells(1, 4) = "Projektelm"
            Cells(1, 5) = "Status"
            Cells(1, 6) = "PrRevPl000"
            Cells(1, 7) = "PrRevPl000"
            Cells(1, 8) = "Act. rev."
            Cells(1, 9) = "Act. rev."
            Cells(1, 10) = "Budget"
            Cells(1, 11) = "Budget"
            Cells(1, 12) = "PrCstSc000"
            Cells(1, 13) = "PrCstSc000"
            Cells(1, 14) = "Act. costs"
            Cells(1, 15) = "Act. costs"
            Cells(1, 16) = "TtlCstComm"
            Cells(1, 17) = "TtlCstComm"
            Cells(1, 18) = "Finish (B)"
            Cells(1, 19) = "Finish (A)"
            Cells(1, 20) = "Pers.resp."
            Cells(1, 21) = "Work"
            Cells(1, 22) = "Work"
            Cells(1, 23) = "Work (A)"
            Cells(1, 24) = "Work (A)"
            Cells(1, 25) = "Work ctr"
            
            
            
            Rows("2:4").Delete
            
            Columns("A:Y").AutoFit
            

End Function

Sub Option_AddGroup()

    Dim i, a, b, c, Prj_Code, Level, Group_St, Group_Found, Group_End, Header_R
    Dim Header As Range
        For i = 2 To 65535
            Level = Cells(i, 1)
            Prj_Code = Cells(i, 3)
            If Level = "" Then
                Rows(i & ":" & i).Insert
                Cells(i, 1) = "x"
                Exit For
            End If
            a = Cells(i - 1, 3)
            
            If (Level = "03" And Not (Prj_Code Like "9*") And Not (Prj_Code Like "0*")) Or (Level = "02") Then
                If (a Like "9*") Or (a Like "0*") Or (a Like "5*") Or (a Like "3*") Or (a Like "4*") Or (a Like "6*") Then
                    Rows(i & ":" & i).Insert
                    Cells(i, 1) = "x"
                    i = i + 1
                End If
            End If
        
        Next i
        
        For i = 5 To 100
            a = Cells(1, i)
            If a = "" Then
                Set Header = Range(Cells(1, 1), Cells(1, i))
                Header_R = i
                Exit For
            End If
        Next i
                
        For i = 2 To 65535
            a = Cells(i, 1)
            If a = "x" Then
                Header.Copy Range(Cells(i, 1), Cells(i, Header_R))
            End If
        Next i
        
        ' Grouping (WBS)
        Group_Found = False
        For i = 1 To 65535
            Level = Cells(i, 1)
            Prj_Code = Cells(i, 3)
            If Level = "" Then Exit For
            
            If Len(Prj_Code) = 17 Then
                If Group_Found = False Then
                    Group_Found = True
                    Group_St = i + 1
                Else
                    If Group_St < i Then
                        Group_Found = False
                        Group_End = i - 1
            
                        Range(Rows(Group_St), Rows(Group_End)).Select
                        Selection.Rows.Group
                        
                        Group_Found = True
                        Group_St = i + 1
                    Else
                        Group_St = i + 1
                    End If
                End If
            End If
            
'            If Level = "Dummy" Then
            If Level = "Level" Then
                If Group_Found = True Then
                    Group_Found = False
                    If Group_St < i Then
                        Group_End = i - 1
            
                        Range(Rows(Group_St), Rows(Group_End)).Select
                        Selection.Rows.Group
                    End If
                End If
            End If
        Next i
        If Group_Found = True Then
            Group_Found = False
            Group_End = i - 2
    
            Range(Rows(Group_St), Rows(Group_End)).Select
            Selection.Rows.Group
        End If


End Sub


Sub Delete_MS()
    Dim i, Prj_Obj, Prj_Code
        For i = 2 To 65535

            Prj_Obj = Cells(i, 2)
            Prj_Code = Cells(i, 3)
            If Left(Prj_Obj, 18) = "Standard Milestone" And Left(Prj_Code, 1) = "0" Then
                Rows(i & ":" & i).Delete
                i = i - 1
            End If
        Next


End Sub
Function Add_Issue(Pos_X As Integer)







End Function
Sub Quick_CNS41G_WBS()
    Quick_CNS41 ("WBS")
End Sub
Sub Quick_CNS41G_PD()
    Quick_CNS41 ("PD")
End Sub

Public Sub Quick_CNS41(InMode)
    Dim flgSapIsOpen As Boolean
    Dim SAPApp As Object
    Dim SapGuiApp As Object
    Dim Connection As Object
    Dim session As Object
    Dim infobox As Object
    Dim CB As New DataObject

    Dim c As Range
    Dim WBS_Check, a, b, d, i, j, k, Dum, b2, b3
    Dim PS_CNF, PS_WBS, PS_WBS2
    Dim wbRpt As Workbook
    Dim NewWB, CLPuserid, CLPpasswd, WBS_Ct, NW_Ct, PD_Ct
    
    CLPrptName = "Worksheet in*"
    CLPsapName = "P12 ONE"
    WBS_Ct = 0
    NW_Ct = 0
    PD_Ct = 0
    b = ""
    b2 = ""
    b3 = ""
    Application.StatusBar = False
    If InMode = "WBS" Then
        For Each c In Selection
            If Not (c.EntireColumn.Hidden Or c.EntireRow.Hidden) Then
            
                WBS_Check = True
                a = c.value
                
                If Left(a, 4) = "WBS " Or Left(a, 4) = "WBS:" Then a = Right(a, Len(a) - 4)
                If Left(a, 3) = "NW " Or Left(a, 3) = "NW:" Then a = Right(a, Len(a) - 3)
                If Left(a, 3) = "PD " Or Left(a, 3) = "PD:" Then a = Right(a, Len(a) - 3)
                
                Dum = InStr(a, ":")
                If Dum > 0 Then a = Left(a, Dum - 1)
                Dum = InStr(a, " ")
                If Dum > 0 Then a = Left(a, Dum - 1)
                
                For i = 1 To 3
                    If IsNumeric(Mid(a, i, i)) = True Then
                        WBS_Check = False
                        Exit For
                    End If
                Next
'                If IsNumeric(Mid(a, 4, 2)) = False Then
'                    WBS_Check = False
'
'                End If
                If Mid(a, 6, 1) <> "." Then
                    WBS_Check = False
                    
                End If
'                If IsNumeric(Mid(a, 7, 6)) = False Then
'                    WBS_Check = False
'
'                End If
                If WBS_Check Then
                    b = b & a & Chr(13) & Chr(10)
                    WBS_Ct = WBS_Ct + 1
                End If
                
                If Len(a) = 12 Then
                    PD_Ct = PD_Ct + 1
                    WBS_Ct = WBS_Ct - 1
                    b3 = b3 & a & Chr(13) & Chr(10)
                End If
                
                WBS_Check = True
                If IsNumeric(a) = False Or Left(a, 1) <> "9" Then
                    WBS_Check = False
                    
                End If
                If WBS_Check Then
                    a = Left(a, 8)
                    b2 = b2 & a & Chr(13) & Chr(10)
                    NW_Ct = NW_Ct + 1
                End If
            End If
        Next c
        If WBS_Ct = 0 And NW_Ct > 0 Then
            InMode = "NW"
            b = b2
        Else
            If WBS_Ct = 0 And PD_Ct > 0 Then
                InMode = "PD"
                b = b3
            End If
        End If
    Else
        If InMode = "PD" Then
            For Each c In Selection
                If Not (c.EntireColumn.Hidden Or c.EntireRow.Hidden) Then
                
                    WBS_Check = True
                    a = c.value
                    
                    If Left(a, 3) = "PD " Or Left(a, 3) = "PD:" Then a = Right(a, Len(a) - 3)
                    Dum = InStr(a, ":")
                    If Dum > 0 Then a = Left(a, Dum - 1)
                    Dum = InStr(a, " ")
                    If Dum > 0 Then a = Left(a, Dum - 1)

                    For i = 1 To 3
                        If IsNumeric(Mid(a, i, i)) = True Then
                            WBS_Check = False
                            Exit For
                        End If
                    Next
                    If IsNumeric(Mid(a, 4, 2)) = False Then
                        WBS_Check = False
                        
                    End If
'                    If IsNumeric(Mid(a, 7, 6)) = False Then
'                        WBS_Check = False
'
'                    End If
                    If WBS_Check Then
                        a = Left(a, 12)
                        b = b & a & Chr(13) & Chr(10)
                    End If
                End If
            Next c
        Else
            For Each c In Selection
                
                WBS_Check = True
                a = c.value
                
                If Left(a, 3) = "NW " Or Left(a, 3) = "NW:" Then a = Right(a, Len(a) - 3)
                Dum = InStr(a, ":")
                If Dum > 0 Then a = Left(a, Dum - 1)
                If IsNumeric(a) = False Or Left(a, 1) <> "9" Then
                    WBS_Check = False
                    
                End If
                If WBS_Check Then
                    a = Left(a, 8)
                    b = b & a & Chr(13) & Chr(10)
                End If
            Next c
        End If
    End If
    
    If b = "" Then
        MsgBox "Please select WBS(s) or Project Definition(s) and Click Quick CNS41 button."
    Else
        ''ClipBoardに文字列をセットする
        CNS41AFGUserForm2.Label1 = ""
        CNS41AFGUserForm2.Show
        CNS41frm.Show
        
        
        If CNS41AFGUserForm2.Label1 = "" Or CNS41AFGUserForm2.Label1 = "1" Then GoTo Exit_Quick_CNS41
        
        For i = 0 To 10
            PS_CNF = CNS41AFGUserForm2.OptionButton1.value
            PS_WBS = CNS41AFGUserForm2.OptionButton2.value
        Next
        On Error Resume Next
        
        Set SapGuiApp = GetObject("SAPGUI")
        Set SAPApp = SapGuiApp.GetScriptingEngine
        If Connection Is Nothing Then
            Set Connection = SAPApp.Children(0)
        End If
        If session Is Nothing Then
            Set session = Connection.Children(0)
        End If
    
    
        If Err = 0 Then flgSapIsOpen = True
        Err.Clear
        
        On Error GoTo 0
        If flgSapIsOpen = False Then
            CNS41AFGUserForm3.Label5 = ""
            CNS41AFGUserForm3.Show
            If CNS41AFGUserForm3.Label5 = "" Or CNS41AFGUserForm3.Label5 = "1" Then GoTo Exit_Quick_CNS41
            CLPsapName = CNS41AFGUserForm3.TextBox1.value
            CLPuserid = CNS41AFGUserForm3.TextBox2.value
            CLPpasswd = CNS41AFGUserForm3.TextBox3.value
            Set SapGuiApp = CreateObject("Sapgui.ScriptingCtrl.1")
            Set Connection = SapGuiApp.OpenConnection(CLPsapName, True, False)
            If Connection Is Nothing Then GoTo Exit_AllReport
            Set session = Connection.Children(0)
    
            If Not IsObject(Connection) Then
                Set Connection = SAPApp.Children(0)
            End If
            If Not IsObject(session) Then
                Set session = Connection.Children(0)
            End If
            
            session.findById("wnd[0]").resizeWorkingPane 279, 29, False
            session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = CLPuserid
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = CLPpasswd
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").SetFocus
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 8
            session.findById("wnd[0]").sendVKey 0
            Set infobox = session.findById("wnd[1]/tbar[0]/btn[12]", False)
            If IsObject(infobox) And IsNull(infobox) = False And (infobox Is Nothing = False) Then
                session.findById("wnd[1]/tbar[0]/btn[12]", False).press
            End If
    
        End If
        
        'session.findById("wnd[0]").maximize
        
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
        session.findById("wnd[0]").sendVKey 0
        
        If CNS41frm.CN41btn.value = True Then
        session.findById("wnd[0]/tbar[0]/okcd").Text = "cn41"
        session.findById("wnd[0]").sendVKey 0
        Else
        session.findById("wnd[0]/tbar[0]/okcd").Text = "cns41"
        session.findById("wnd[0]").sendVKey 0
        End If
        
        If PS_CNF Then
            session.findById("wnd[0]/tbar[1]/btn[28]").press
            session.findById("wnd[1]/usr/ctxtTCNTT-PROFID").Text = "ZNRJCNS41cnf"
            session.findById("wnd[1]/usr/ctxtTCNTT-PROFID").caretPosition = 12
            session.findById("wnd[1]/tbar[0]/btn[0]").press
            
            session.findById("wnd[0]/usr/ctxtCN_MATNR-LOW").Text = ""
            session.findById("wnd[0]/usr/ctxtCN_ACTVT-LOW").Text = ""
            session.findById("wnd[0]/usr/ctxtCN_NETNR-LOW").Text = ""
            session.findById("wnd[0]/usr/ctxtCN_PROJN-LOW").Text = ""
            session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").Text = ""
            session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").SetFocus
            session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").caretPosition = 0
            session.findById("wnd[0]").sendVKey 0
            If InMode = "NW" Then
                session.findById("wnd[0]/usr/btn%_CN_NETNR_%_APP_%-VALU_PUSH").press
            Else
                If InMode = "PD" Then
                    session.findById("wnd[0]/usr/btn%_CN_PROJN_%_APP_%-VALU_PUSH").press
                Else
                    session.findById("wnd[0]/usr/btn%_CN_PSPNR_%_APP_%-VALU_PUSH").press
                End If
            End If
            
        
            ''ClipBoardに文字列をセットする
            With CB
                .SetText b        ''変数のデータをDataObjectに格納する
                .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
                .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
                Dum = .GetText     ''DataObjectのデータを変数に取得する
            End With
        
            
            session.findById("wnd[1]/tbar[0]/btn[24]").press
        
            session.findById("wnd[1]/tbar[0]/btn[8]").press
            session.findById("wnd[0]/tbar[1]/btn[8]").press
            session.findById("wnd[0]/mbar/menu[0]/menu[11]/menu[4]").Select
            session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").Select
            session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").SetFocus
            session.findById("wnd[1]/tbar[0]/btn[0]").press
            session.findById("wnd[1]/tbar[0]/btn[0]").press
        
        Else
            
            session.findById("wnd[0]/tbar[1]/btn[28]").press
            session.findById("wnd[1]/usr/ctxtTCNTT-PROFID").Text = "ZNRJCNS41wbs"
            session.findById("wnd[1]").sendVKey 0
            session.findById("wnd[0]/usr/ctxtCN_PROJN-LOW").Text = ""
            session.findById("wnd[0]/usr/ctxtCN_PROJN-LOW").caretPosition = 0
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").Text = ""
            session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").SetFocus
            session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").caretPosition = 0
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]/usr/ctxtCN_NETNR-LOW").Text = ""
            session.findById("wnd[0]/usr/ctxtCN_NETNR-LOW").SetFocus
            session.findById("wnd[0]/usr/ctxtCN_NETNR-LOW").caretPosition = 0
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]/usr/ctxtCN_ACTVT-LOW").Text = ""
            session.findById("wnd[0]/usr/ctxtCN_ACTVT-LOW").SetFocus
            session.findById("wnd[0]/usr/ctxtCN_ACTVT-LOW").caretPosition = 0
            session.findById("wnd[0]").sendVKey 0
            If InMode = "NW" Then
                session.findById("wnd[0]/usr/btn%_CN_NETNR_%_APP_%-VALU_PUSH").press
            Else
                If InMode = "PD" Then
                    session.findById("wnd[0]/usr/btn%_CN_PROJN_%_APP_%-VALU_PUSH").press
                Else
                    session.findById("wnd[0]/usr/btn%_CN_PSPNR_%_APP_%-VALU_PUSH").press
                End If
            End If
            'session.findById("wnd[0]/usr/btn%_CN_PSPNR_%_APP_%-VALU_PUSH").press
            
            ''ClipBoardに文字列をセットする
            With CB
                .SetText b        ''変数のデータをDataObjectに格納する
                .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
                .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
                Dum = .GetText     ''DataObjectのデータを変数に取得する
            End With

            session.findById("wnd[1]/tbar[0]/btn[24]").press
            session.findById("wnd[1]/tbar[0]/btn[8]").press
            session.findById("wnd[0]/tbar[1]/btn[8]").press
            session.findById("wnd[0]/mbar/menu[0]/menu[11]/menu[4]").Select
            session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").Select
            session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").SetFocus
            session.findById("wnd[1]/tbar[0]/btn[0]").press
            session.findById("wnd[1]/tbar[0]/btn[0]").press
        End If
        
            

        For Each wbRpt In Workbooks
            If wbRpt.Name Like CLPrptName Then
                
                NewWB = Workbooks.Add.Name
                wbRpt.Worksheets(1).Cells.Copy Workbooks(NewWB).Worksheets(1).Cells
                wbRpt.Close
                
                Exit For
            End If
        Next
        
        Workbooks(NewWB).Activate
    
    Sheets("Sheet1").Select
    Sheets("Sheet1").Copy Before:=Sheets(2)
    Sheets("Sheet1 (2)").Select
    Sheets("Sheet1 (2)").Name = "Original"
    Sheets("Original").Select
    With ActiveWorkbook.Sheets("Original").Tab
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0
    End With
    Sheets("Original").Select
    Sheets("Original").Move Before:=Sheets(1)
    Sheets("Sheet1").Select
        'Save WB
        Dim path As String
        path = "C:\Users\" & Environ("Username") & "\Documents\PFM SmartApp"
        If Len(Dir(path, vbDirectory)) = 0 Then
        MkDir (path)
        End If
        If Len(Dir(path & "/" & Format(Now(), "yyyy"), vbDirectory)) = 0 Then
        MkDir (path & "/" & Format(Now(), "yyyy"))
        End If
        
        ActiveWorkbook.SaveAs (path & "/" & Format(Now(), "yyyy") & "\" & "CNS41" & "_" & Format(Now(), "mmddhhmmss") & ".xlsx"), FileFormat:=51
        
            ChangeFormA
'            If PS_WBS Then
'                Add_BU
'                Columns("E:E").ColumnWidth = 5
'            End If
    End If
Exit_AllReport:

    Application.DisplayAlerts = True
    Set SAPApp = Nothing
    'Set SapGuiApp = Nothing
    Set Connection = Nothing
    Set session = Nothing
    Set infobox = Nothing
    Set CB = Nothing
    
Exit_Quick_CNS41:
    Exit Sub
    
Err_Quick_CNS41:
    MsgBox Err & Error()
    
    Resume Exit_Quick_CNS41
End Sub
Function Copy_Formulas(NewWB)
    Dim a, b, Dum, path, FileN, RowN, dum2
    path = "G:\Special\MUS_Project\MUS_Central_Database\CNS41AutoFormat_DB"
    FileN = "PSP_Formula_Template.xlsx"
    Workbooks.Open path & "\" & FileN
    a = Cells(5, 3)
    b = Cells(6, 3)
    Dum = a & "1:" & b & "2"
    Range(Dum).Copy Workbooks(NewWB).Worksheets(1).Range(Dum)
    Workbooks(FileN).Close
    Workbooks(NewWB).Worksheets(1).Activate
    RowN = ActiveSheet.Range("A1").End(xlDown).Row
    Dum = a & "2:" & b & "2"
    dum2 = a & "2:" & b & RowN
    
    Range(Dum).AutoFill Destination:=Range(dum2), Type:=xlFillDefault
End Function
Sub Add_BU()
    Dim i, a, b, BU
    
    b = "02"
    
    i = 2
    a = Cells(i, 1)
    Do While a <> ""
        If a = b Then
            BU = Cells(i, 4)
            If BU = "N/A" Or BU = 0 Then
                Range(Cells(i, 4), Cells(i, 4)).FormulaR1C1 = "=R[-1]C[0]"
                
            End If
        End If
        i = i + 1
        a = Cells(i, 1)
    Loop

        
End Sub
Sub test()
    Dim W1 As CommandBar
    Dim i As Long
    For Each W1 In Application.CommandBars
        
            i = i + 1
            Cells(i, 1).value = W1.index
            
            Cells(i, 2).value = W1.Name
    Next
End Sub
Function Check_G()
Dim PathG, Dum, DirN
    PathG = "G:\Special\MUS_Project\MUS_Central_Database\CNS41AutoFormat_DB"
    DirN = Dir(PathG, vbDirectory)
    If DirN = "" Then
        Check_G = 1
    Else
        Check_G = 2
    End If
End Function

