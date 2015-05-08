VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CreateQTCMfrm 
   Caption         =   "Create QTC-M script options"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5310
   OleObjectBlob   =   "CreateQTCMfrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CreateQTCMfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()

If Me.Settlementbox.value = False And Me.Releasebox.value = True Then
MsgBox "Settlement box must be checked in order to be able to release the WBS.", vbExclamation
Exit Sub
End If

Me.Hide

End Sub

Private Sub UserForm_Initialize()
Me.Settlementbox.value = True
Me.Custgrp5box.value = True

    With Me.YearBox
                    .AddItem "Current"
                    .AddItem Format(Now(), "yyyy") - 1
                    .AddItem Format(Now(), "yyyy")
    End With
    
    With Me.MonthBox
                    .AddItem "Current"
                    .AddItem "01"
                    .AddItem "02"
                    .AddItem "03"
                    .AddItem "04"
                    .AddItem "05"
                    .AddItem "06"
                    .AddItem "07"
                    .AddItem "08"
                    .AddItem "09"
                    .AddItem "10"
                    .AddItem "11"
                    .AddItem "12"
    End With

End Sub
