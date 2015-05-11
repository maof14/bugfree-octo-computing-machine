VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings 
   Caption         =   "SmartApp settings"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8655
   OleObjectBlob   =   "frmSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public exportPath As String
Public chosenWB As String
Public proceed As Boolean
Dim exportPrompt As frmExport

Private Sub btnCancel_Click()
    If GetTurnOffOption = "Yes" Then
        Me.chkTurnoff.value = True
    Else
        Me.chkTurnoff.value = False
    End If
        If GetEmailOption = "Yes" Then
        Me.chkEmail.value = True
    Else
        Me.chkEmail.value = False
    End If
    If GetSAPOption = "Yes" Then
        Me.chkSAP.value = True
    Else
        Me.chkSAP.value = False
    End If

    If Not GetReportFilePath = "" Then
        Me.txtFolderPath = GetReportFilePath
    End If
        If Not GetASCFilePath = "" Then
        Me.txtFilePathASC = GetASCFilePath
    End If
        If Not GetContactsFilePath = "" Then
        Me.txtFilePathContacts = GetContactsFilePath
    End If
        If Not GetDATEFORMAT = "" Then
        Me.Dateformattxt = GetDATEFORMAT
    End If

    Me.Hide
End Sub

Private Sub btnChoosePath_Click()
    Dim fldr As FileDialog
    Dim selectedFolder As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then
            'selectedFolder = Application.DefaultFilePath
            GoTo LeaveFolderPicker
        End If
        selectedFolder = .SelectedItems(1)
    End With

    ' Set value to folder picker textbox
    Me.txtFolderPath = selectedFolder
LeaveFolderPicker:
    Set fldr = Nothing
End Sub

Private Sub btnClearPath_Click()
    SetReportFilePath ("")
    Me.txtFolderPath.value = ""
End Sub
Private Sub btnChooseFileASC_Click()
    Dim fldr As FileDialog
    Dim selectedFile As String
    Set fldr = Application.FileDialog(msoFileDialogOpen)
    With fldr
        .Title = "Select a File"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then
            'selectedFolder = Application.DefaultFilePath
            GoTo LeaveFolderPicker
        End If
        selectedFile = .SelectedItems(1)
    End With

    ' Set value to folder picker textbox
    Me.txtFilePathASC = selectedFile
LeaveFolderPicker:
    Set fldr = Nothing
End Sub

Private Sub btnClearFileASC_Click()
    SetASCFilePath ("")
    Me.txtFilePathASC.value = ""
End Sub
Private Sub btnChooseFileContacts_Click()
    Dim fldr As FileDialog
    Dim selectedFile As String
    Set fldr = Application.FileDialog(msoFileDialogOpen)
    With fldr
        .Title = "Select File where you have your contacts"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then
            'selectedFolder = Application.DefaultFilePath
            GoTo LeaveFolderPicker
        End If
        selectedFile = .SelectedItems(1)
    End With

    ' Set value to folder picker textbox
    Me.txtFilePathContacts = selectedFile
LeaveFolderPicker:
    Set fldr = Nothing
End Sub

Private Sub btnClearFileContacts_Click()
    SetContactsFilePath ("")
    Me.txtFilePathContacts.value = ""
End Sub

' Settings form - Export button.
Private Sub btnExport_Click()
    Dim gr As CGitResource
    Set exportPrompt = New frmExport
    frmExport.Show
    Set gr = New CGitResource
    If frmExport.proceed = False Then Exit Sub
    gr.Init frmExport.chosenWB
    gr.ExportCode
    Set gr = Nothing
    Set frmExport = Nothing
End Sub

' Settings form - Import button.
Private Sub btnImport_Click()
    Dim gr As CGitResource
    Set gr = New CGitResource
    gr.Init
    gr.ImportCode
    Set gr = Nothing
End Sub

Private Sub btnOk_Click()
    If Me.chkTurnoff.value = True Then
        SetTurnOffOption ("Yes")
        MsgBox "When the macro is finished, the PC will shut down automatically after two minutes. If you wish to cancel the process at that time, type in ""shutdown -a"" into Run from the start menu to abort.", 48, "Automatic shutdown enabled"
    Else
        SetTurnOffOption ("No")
    End If
       
    If Me.chkEmail.value = True Then
        SetEmailOption ("Yes")
    Else
        SetEmailOption ("No")
    End If
    
    If Me.chkSAP.value = True Then
        SetSAPOption ("Yes")
        MsgBox "With this option ON, the script will overtake any of the SAP windows you have open", vbOKOnly, "Keep in mind..."
    Else
        SetSAPOption ("No")
    End If


    ' Kolla om det är något skrivet i textboxen folderpath. Om det inte är det - ändra inte i ini-filen.
    If Me.txtFolderPath <> "" Then
        SetReportFilePath (txtFolderPath)
    End If
    If Me.txtFilePathASC <> "" Then
        SetASCFilePath (txtFilePathASC)
    End If
    If Me.txtFilePathContacts <> "" Then
        SetContactsFilePath (txtFilePathContacts)
    End If
    If Me.Dateformattxt <> "" Then
        SetDATEFORMAT (Dateformattxt)
    End If

    Me.Hide
End Sub


Private Sub CommandButton1_Click()
    
End Sub

Private Sub UserForm_Activate()
    If GetTurnOffOption = "Yes" Then
        Me.chkTurnoff.value = True
    Else
        Me.chkTurnoff.value = False
    End If
        If GetEmailOption = "Yes" Then
        Me.chkEmail.value = True
    Else
        Me.chkEmail.value = False
    End If
    If GetSAPOption = "Yes" Then
        Me.chkSAP.value = True
    Else
        Me.chkSAP.value = False
    End If

    If Not GetReportFilePath = "" Then
        Me.txtFolderPath = GetReportFilePath
    End If
        If Not GetASCFilePath = "" Then
        Me.txtFilePathASC = GetASCFilePath
    End If
        If Not GetContactsFilePath = "" Then
        Me.txtFilePathContacts = GetContactsFilePath
    End If
            If Not GetDATEFORMAT = "" Then
        Me.Dateformattxt = GetDATEFORMAT
    End If
    
End Sub
