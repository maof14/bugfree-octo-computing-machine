Attribute VB_Name = "Goodies"
Sub MoreGoodies(ByRef Button As String)
With Workbooks
    If .Count = 0 Then
    .Add
    End If
End With

If Button = "MSTbtn" Then
ActiveWorkbook.FollowHyperlink Address:="https://esessmw2206.ss.sw.ericsson.se:444/MicroStrategy/asp/Main.aspx", NewWindow:=True

ElseIf Button = "Citrixbtn" Then
ActiveWorkbook.FollowHyperlink Address:="https://lighthouse.lmera.ericsson.se/", NewWindow:=True

ElseIf Button = "ECAbtn" Then
ActiveWorkbook.FollowHyperlink Address:="http://anon.ericsson.se/eridoc/component/eriurl?docno=&objectId=09004cff82271138&action=approved&format=excel8book", NewWindow:=True

ElseIf Button = "RGbtn" Then
Workbooks.Open fileName:="\\esekina005\groupfbs\SmartApp\Excel\Request_generator.xlsm", ReadOnly:=True

ElseIf Button = "CPLbtn" Then
ActiveWorkbook.FollowHyperlink Address:="http://anon.ericsson.se/eridoc/component/eriurl?docno=LME-10:000280Uen&objectId=09004cff839d627d&action=approved&format=excel8book", NewWindow:=True

ElseIf Button = "EBPbtn" Then
ActiveWorkbook.FollowHyperlink Address:="http://internal.ericsson.com/book/10492/ebp-ericsson-business-process/", NewWindow:=True

ElseIf Button = "ECBbtn" Then
ActiveWorkbook.FollowHyperlink Address:="http://internal.ericsson.com/book-pages/31271/list-reporting-objects?unit=30957987", NewWindow:=True

ElseIf Button = "ISAbtn" Then
ActiveWorkbook.FollowHyperlink Address:="https://erilink.ericsson.se/eridoc/erl/objectId/09004cff88f1e927?docno=39/00201-5/FEA101701Uen&action=approved&format=excel12mebook", NewWindow:=True

ElseIf Button = "JVbtn" Then
ActiveWorkbook.FollowHyperlink Address:="http://erilink.ericsson.se/eridoc/erl/objectId/09004cff80a36414?docno=21/00201-5/FEA101701Uen&action=current&format=excel8book", NewWindow:=True

ElseIf Button = "OEFbtn" Then
ActiveWorkbook.FollowHyperlink Address:="http://anon.ericsson.se/eridoc/component/eriurl?docno=5/00201-8/FEA101701Uen&objectId=09004cff81548121&action=approved&format=excel8book", NewWindow:=True

ElseIf Button = "FI_Invbtn" Then
ActiveWorkbook.FollowHyperlink Address:="http://anon.ericsson.se/eridoc/erl/objectId/09004cff8152d479?docno=3/00201-1/FEA101701Uen&action=approved&format=excel8book", NewWindow:=True

ElseIf Button = "PECRatesbtn" Then
ActiveWorkbook.FollowHyperlink Address:="http://erilink.ericsson.se/eridoc/erl/objectId/09004cff802da4ed?docno=&action=current&format=excel8book", NewWindow:=True

ElseIf Button = "WISearchbtn" Then
ActiveWorkbook.FollowHyperlink Address:="http://anon.ericsson.se/eridoc/erl/objectId/09004cff85c17941?docno=1/00151-FEA101701Uen&action=approved&format=excel8book", NewWindow:=True

ElseIf Button = "Winshuttlebtn" Then
ActiveWorkbook.FollowHyperlink Address:="https://ericoll.internal.ericsson.com/sites/Company_Control_Stockholm/ToolBox/Forms/AllItems.aspx?RootFolder=%2Fsites%2FCompany%5FControl%5FStockholm%2FToolBox%2FWinshuttle%2FTemplates", NewWindow:=True

ElseIf Button = "AD09btn" Then
ActiveWorkbook.FollowHyperlink Address:="http://anon.ericsson.se/eridoc/component/eriurl?docno=LME-09:000085Uen&objectId=09004cff82f1c3a1&action=approved&format=ppt8", NewWindow:=True

ElseIf Button = "AD09IMbtn" Then
ActiveWorkbook.FollowHyperlink Address:="https://erilink.ericsson.se/eridoc/erl/objectId/09004cff834827df?action=approved&format=ppt8", NewWindow:=True

ElseIf Button = "DMCbtn" Then
ActiveWorkbook.FollowHyperlink Address:="https://ericoll.internal.ericsson.com/sites/Direct_Market_Control_Ericoll_Site/Pages/EAB%20PSP%20Global%20Operations.aspx", NewWindow:=True

ElseIf Button = "PSPbtn" Then
ActiveWorkbook.FollowHyperlink Address:="https://ericoll.internal.ericsson.com/sites/Contract_Accounting_Governance/PSPGlobalSupport/Pages/home.aspx", NewWindow:=True
End If

End Sub

