VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CNS41AFGUserForm2 
   Caption         =   "Select Report Format"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4650
   OleObjectBlob   =   "CNS41AFGUserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CNS41AFGUserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
    CNS41AFGUserForm2.Hide
    CNS41AFGUserForm2.Label1 = "0"
End Sub

Private Sub CommandButton2_Click()
    CNS41AFGUserForm2.Hide
    CNS41AFGUserForm2.Label1 = "1"
End Sub
