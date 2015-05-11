Attribute VB_Name = "RibbonCallbacks"
Option Explicit
Dim scriptname As String

Private Sub OnRibbonLoad(ribbon As IRibbonUI)
    ThisWorkbook.ribbonUI = ribbon
End Sub
Sub PIMP_Click(control As IRibbonControl)

Call PimpItUpScript

End Sub
Sub MoreGoodies_Click(control As IRibbonControl)
Dim Button As String
     
    Button = control.ID

    Call MoreGoodies(Button)

End Sub
Sub Dashboard_Click(control As IRibbonControl)

Call Create_Dashboard

End Sub
Sub DashboardOpen_Click(control As IRibbonControl)

Workbooks.Open fileName:="C:\Users\" & Environ("USERNAME") & "\Documents\PFM SmartApp\Dashboards\Dashboard.xlsm", ReadOnly:=False


End Sub
Sub DashboardKPI_Click(control As IRibbonControl)

Call Create_DashboardKPI

End Sub
Sub DashboardOpenKPI_Click(control As IRibbonControl)

Workbooks.Open fileName:="C:\Users\" & Environ("USERNAME") & "\Documents\PFM SmartApp\Dashboards\Dashboard KPI.xlsm", ReadOnly:=False


End Sub


Sub Plancost_Click(control As IRibbonControl)

MsgBox "Working hard to bring you this!", vbInformation, "Working"

End Sub


Sub HelpWI(ByVal control As IRibbonControl)

Helpfrm.Show

End Sub


Sub ProductGuide(ByVal control As IRibbonControl)
'On Error GoTo Errorhandler

With Workbooks
    If .Count = 0 Then
    .Add
    End If
End With

Searchfrm.intrasites.value = "Ericsson Intranet"
Searchfrm.searchtxt.value = ""
Searchfrm.searchtxt.SetFocus
Searchfrm.Show


    If Searchfrm.intrasites.value = "PFM Work Instructions" Then
    ActiveWorkbook.FollowHyperlink Address:="https://ericoll.internal.ericsson.com/Search/pages/results.aspx?k=" & Searchfrm.searchtxt.value & "&cs=This%20List&u=https%3A%2F%2Fericoll.internal.ericsson.com%2Fsites%2FCompany_Control_Stockholm%2FLists%2FWork%20Instructions", NewWindow:=True
    
    ElseIf Searchfrm.intrasites.value = "Ericsson Intranet" Then
    ActiveWorkbook.FollowHyperlink Address:="https://search.internal.ericsson.com/?query=" & Searchfrm.searchtxt.value & "&key=intranet", NewWindow:=True
    
    ElseIf Searchfrm.intrasites.value = "Global PSP Ericoll" Then
    ActiveWorkbook.FollowHyperlink Address:="https://ericoll.internal.ericsson.com/Search/pages/results.aspx?k=" & Searchfrm.searchtxt.value & "&cs=This%20Site&u=https%3A%2F%2Fericoll.internal.ericsson.com%2Fsites%2FContract_Accounting_Governance%2FPSPGlobalSupport", NewWindow:=True
    
    ElseIf Searchfrm.intrasites.value = "Supporting documents for NS" Then
    
    ActiveWorkbook.FollowHyperlink Address:="https://ericoll.internal.ericsson.com/Search/pages/results.aspx?k=" & Searchfrm.searchtxt.value & "&cs=This%20List&u=https%3A%2F%2Fericoll.internal.ericsson.com%2Fsites%2FEVS%2FEAB", NewWindow:=True
End If
End Sub

Sub BEX_Click(ByVal control As IRibbonControl)

Dim Button As String
     
    Button = control.ID

    Call EBW_BEXScript(Button)

End Sub
Sub EBW_Click(ByVal control As IRibbonControl)

If control.ID = "OpenEBW" Then

With Workbooks
    If .Count = 0 Then
    .Add
    End If
End With

    ActiveWorkbook.FollowHyperlink Address:="https://ep.ss.sw.ericsson.se/irj/portal", NewWindow:=True
Else
EBWfrm.Show
End If
End Sub
Sub GetItemID(control As IRibbonControl, index As Integer, _
           ByRef ID)
    ID = "bookmark" & index
End Sub

Sub GetTextShare(control As IRibbonControl, ByRef Text)
    Text = " " & HTTPGet("http://finance.yahoo.com/d/quotes.csv?s=EURSEK=X&f=l1", "") & " SEK"  'visible value in combo when initialized

End Sub
Sub GetTextShare2(control As IRibbonControl, ByRef Text)
    Text = " " & HTTPGet("http://finance.yahoo.com/d/quotes.csv?s=ERIC&f=l1", "") & " USD" 'visible value in combo when initialized

End Sub

Sub GetTextRate(control As IRibbonControl, ByRef Text)
    Text = " " & HTTPGet("http://finance.yahoo.com/d/quotes.csv?s=USDSEK=X&f=l1", "") & " SEK" 'visible value in combo when initialized

End Sub
Sub CreateTemplate_Click(control As IRibbonControl)
Dim ID

    SetChosenScript (control.ID)
    
    scriptname = Getchosenscript

Call Template.Create_Template(scriptname)
End Sub

Sub RunScript_Click(control As IRibbonControl)

Progressfrm.Show

End Sub
Sub OEF_CHC_Click(control As IRibbonControl)

    Call OEF.OEF_CHC
    
End Sub

Sub SettingsButton_Click(ByVal control As IRibbonControl)
    frmSettings.Show
End Sub

Sub FIRE_SAP_Click(ByVal control As IRibbonControl)
      On Error GoTo Errorhandler
    Dim ScriptN As String
    
    ScriptN = control.ID & "Script"

Run ScriptN

Errorhandler:
Exit Sub
 
 End Sub
 Sub Auto_Remove()
    AddIns("PFM_SmartApp").Installed = False
 End Sub
Sub FillwithZero(ByVal control As IRibbonControl)
     Dim selrange As Range

    On Error GoTo Errorhandler
  If ActiveSheet.ProtectContents = True Then
  ActiveSheet.Unprotect
  End If
    Set selrange = Selection

    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.FormulaR1C1 = "0"
    
Errorhandler:
    If WorksheetFunction.CountA(selrange) < 1 Then
    Selection.FormulaR1C1 = "0"
    Else
    End If
Exit Sub
 
 End Sub
 Sub FillSAPdate(ByVal control As IRibbonControl)
 Dim todaySAP As String
 
 todaySAP = Format(Now(), GetDATEFORMAT)
    
    Selection.FormulaR1C1 = "'" & todaySAP

 End Sub
 Sub Fillblanks(ByVal control As IRibbonControl)
 Dim selrange As Range
  On Error GoTo Errorhandler
  If ActiveSheet.ProtectContents = True Then
  ActiveSheet.Unprotect
  End If
    Application.ScreenUpdating = False
    Set selrange = Selection
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.FormulaR1C1 = "=R[-1]C"
    selrange.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    
Errorhandler:
Exit Sub

 End Sub
 Sub ADD000_22(ByVal control As IRibbonControl)
 Dim c As Range
 Dim newvalue As String
  On Error GoTo Errorhandler
  If ActiveSheet.ProtectContents = True Then
  ActiveSheet.Unprotect
  End If
    Application.ScreenUpdating = False
        For Each c In Selection
            newvalue = "00" & c.value & "_22"
            c.value = newvalue
        Next c
    Application.ScreenUpdating = True
    
Errorhandler:
Exit Sub

 End Sub

Sub SetTheScriptItemInDropDown(control As IRibbonControl, index As Integer, ByRef returnedVal)

 SetChosenScript (returnedVal)

 End Sub
 Public Function HTTPGet(sUrl As String, sQuery As String) As String
On Error GoTo Errorhandler:
Dim sResult As String
Dim xml As Object
Set xml = CreateObject("Microsoft.XMLHTTP")
xml.Open "GET", sUrl, False
xml.Send
sResult = xml.ResponseText
Set xml = Nothing
HTTPGet = sResult
Errorhandler:
End Function

Sub OneToOneRelation(control As IRibbonControl)
    Dim db As CDatabase
    Dim res As Collection
    Dim warning As frmWarning
    Dim pCodeMatch As Variant
    
    ' Warning not to overwrite
    If (getHideConvertWarning = "0" Or getHideConvertWarning = "") Then
        Set warning = New frmWarning
        warning.Init WConvertWarning
        warning.lblPrompt = "You will not be able to undo this action. If you just want to see the results, you can create a new column, copy the values you want to convert, and try the conversion there so nothing important gets overwritten."
        warning.Show
        If warning.response = False Then Exit Sub
        Set warning = Nothing
    End If
    
    Set db = New CDatabase
    db.Init
    
    Dim str, table, have, want, SQL As String
    str = Split(control.tag, ",")
    
    table = str(0)
    have = str(1)
    want = str(2)
    
    ' PCode special case scenario to be able to toggle between them.
    If (have = "AnyPCode" And want = "AnyPCode") Then
        If (Len(ActiveCell.value) = 3) Then
            have = "pCodeOne"
            want = "pCode"
        Else
            have = "pCode"
            want = "pCodeOne"
        End If
    ElseIf (have = "AnyPCode") Then
        If (Len(ActiveCell.value) = 3) Then
            have = "pCodeOne"
        Else
            have = "pCode"
        End If
    End If
    
    Dim c, r As Variant
    
    For Each c In Selection
        With db
            SQL = .selectQuery(table, want)
            SQL = SQL & .where(have, c)
            Set res = .fetchCollection(SQL)
        End With
        For Each r In res
            c.value = "'" & r
        Next r
    Next c
End Sub
