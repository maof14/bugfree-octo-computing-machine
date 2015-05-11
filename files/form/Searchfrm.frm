VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Searchfrm 
   Caption         =   "Search"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4470
   OleObjectBlob   =   "Searchfrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Searchfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Image1_Click()
Searchfrm.Hide
If Me.intrasites.value = "Supporting documents for NS" Then
'POCdatefrm.Show

End If
End Sub


Private Sub searchtxt_Enter()
Me.searchtxt.value = ""

End Sub

Private Sub searchtxt_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = vbKeyReturn Then
Call Image1_Click
End If


End Sub

'Private Sub searchtxt_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'If KeyAscii = 13 Then
'Call Image1_Click
'End If
'End Sub

    Private Sub UserForm_Initialize()
    'Dim SAPServ As Range
    'Set ws = Worksheets("Sheet2")
    
    'For Each SAPServ In ws.Range("SAPServer")
    With Me.intrasites
        .AddItem "Supporting documents for NS"
        .AddItem "Global PSP Ericoll"
        .AddItem "Ericsson Intranet"
    End With
    'Next SAPServ
End Sub

