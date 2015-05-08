Attribute VB_Name = "EBW_BEX"
Sub EBW_BEXScript(ByRef Button As String)
Dim FileN As String

If Button = "CheckDelivery" Then
FileN = "C:\Users\" & Environ("USERNAME") & "\Documents\PFM SmartApp\BEX Reports\Check Deliveries.xlsm"
ElseIf Button = "InvNotDeliv" Then
FileN = "C:\Users\" & Environ("USERNAME") & "\Documents\PFM SmartApp\BEX Reports\Invoiced not delivered.xlsm"
ElseIf Button = "ASC" Then
FileN = "C:\Users\" & Environ("USERNAME") & "\Documents\PFM SmartApp\BEX Reports\Assignment Status Card.xlsm"
ElseIf Button = "COPA" Then
FileN = "C:\Users\" & Environ("USERNAME") & "\Documents\PFM SmartApp\BEX Reports\COPA Details.xlsm"
ElseIf Button = "COPAFI" Then
FileN = "C:\Users\" & Environ("USERNAME") & "\Documents\PFM SmartApp\BEX Reports\COPA FI.xlsm"
ElseIf Button = "COPAFIYTD" Then
FileN = "C:\Users\" & Environ("USERNAME") & "\Documents\PFM SmartApp\Dashboards\COPA FI YTD.xlsm"
ElseIf Button = "COPAFIISO" Then
FileN = "C:\Users\" & Environ("USERNAME") & "\Documents\PFM SmartApp\Dashboards\COPA FI ISO.xlsm"
ElseIf Button = "Elis" Then
FileN = "C:\Users\" & Environ("USERNAME") & "\Documents\PFM SmartApp\BEX Reports\Elis Report.xlsm"
ElseIf Button = "EPQ" Then
FileN = "C:\Users\" & Environ("USERNAME") & "\Documents\PFM SmartApp\BEX Reports\EPQ.xlsm"
ElseIf Button = "ASCKPI" Then
FileN = "C:\Users\" & Environ("USERNAME") & "\Documents\PFM SmartApp\Dashboards\ASC KPI.xlsm"
ElseIf Button = "ASCKPIPSF" Then
FileN = "C:\Users\" & Environ("USERNAME") & "\Documents\PFM SmartApp\Dashboards\ASC PSF-M KPI.xlsm"
ElseIf Button = "OpenBEX" Then
Shell "C:\Program Files (x86)\SAP\Business Explorer\BI\BExAnalyzer.exe", vbMaximizedFocus
Exit Sub
ElseIf Button = "OpenNewBEX" Then
Shell "C:\Program Files (x86)\SAP BusinessObjects\Analysis\BiSharedAddinLauncher.exe", vbMaximizedFocus
Exit Sub
End If

Shell "C:\Program Files (x86)\SAP\Business Explorer\BI\BExAnalyzer.exe" & " " & FileN, vbMaximizedFocus

End Sub


