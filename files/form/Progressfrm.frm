VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Progressfrm 
   Caption         =   "Processing..."
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8745
   OleObjectBlob   =   "Progressfrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Progressfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim scriptname As String

Private Sub UserForm_Activate()
'Application.Wait (Now + TimeValue("00:00:02"))
Me.SpecialEffect = fmSpecialEffectFlat
Me.Text.Caption = "1% Processed"
Me.Bar.Width = 0.01 * 306
'scriptname = Getchosenscript
Call MainScript
End Sub
