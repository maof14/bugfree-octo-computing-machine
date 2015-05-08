VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Helpfrm 
   Caption         =   "Help"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2910
   OleObjectBlob   =   "Helpfrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Helpfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FAQbtn_Click()
Call PFMLogScript
Me.Hide
End Sub

Private Sub Label3_Click()
ActiveWorkbook.FollowHyperlink Address:="mailto:eedgcan;qolsmat", NewWindow:=True

End Sub

Private Sub UpdateBtn_Click()
      Dim batchPath As String
      Dim iRet As Integer
      
      On Error GoTo Errorhandler
      iRet = MsgBox("You must have all other EXCEL sessions CLOSED (e.g. BEX) before continuing. Are you sure you want to proceed with update?", vbYesNo, "Close multiple EXCEL sessions")
    If iRet = vbNo Then
        End
    Else
With Workbooks
    If .Count = 0 Then
    .Add
    End If
End With

Workbooks.Open fileName:="\\esekina005\groupfbs\SmartApp\Excel\Update SmartApp.xlsm", ReadOnly:=True
    'Run batch file with argument of current directory
'    Shell batchPath, vbNormalFocus
    Call Auto_Remove
    End If
Errorhandler:
Exit Sub
 

End Sub


Private Sub UserForm_Initialize()
Me.Versiontxt.Caption = "4." & Format(FileDateTime("C:\Users\" & Environ("Username") & "\AppData\Roaming\Microsoft\AddIns\PFM_SmartApp.xlam"), "YYMMDDHHSS")

End Sub

Private Sub WIbtn_Click()
ActiveWorkbook.FollowHyperlink Address:="https://ericoll.internal.ericsson.com/sites/Company_Control_Stockholm/ToolBox/PFM%20Smart%20App/WI%20PFM%20SmartApp.docx", NewWindow:=True
End Sub
