Attribute VB_Name = "SettingsHandler"
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
ByVal lpApplicationName As String, _
ByVal lpKeyName As String, _
ByVal lpDefault As String, _
ByVal lpReturnedString As String, _
ByVal nSize As Long, _
ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
ByVal lpApplicationName As String, _
ByVal lpKeyName As String, _
ByVal lpString As String, _
ByVal lpFileName As String) As Long



Private Function GetINIString(ByVal sApp As String, ByVal sKey As String, ByVal filepath As String) As String
    Dim sBuf As String * 256
    Dim lBuf As Long

    lBuf = GetPrivateProfileString(sApp, sKey, "", sBuf, Len(sBuf), filepath)
    GetINIString = Left$(sBuf, lBuf)
End Function

Private Function SetINIString(ByVal sApp As String, ByVal sKey As String, ByVal sString As String, lpFileName As String) As String
    Dim sBuf As String * 256
    Dim lBuf As Long
    SetINIString = WritePrivateProfileString(sApp, sKey, sString, lpFileName)
End Function

Public Function SetTurnOffOption(ByVal sVal As String) As String
    Dim path As String
    path = ThisWorkbook.path & "\" & "settings.ini"
    SetTurnOffOption = SetINIString("TurnOffPC", "Option", sVal, path)
End Function

Public Function GetTurnOffOption() As String
    Dim path As String
    path = ThisWorkbook.path & "\" & "settings.ini"
    GetTurnOffOption = GetINIString("TurnOffPC", "Option", path)
End Function
Public Function SetEmailOption(ByVal sVal As String) As String
    Dim path As String
    path = ThisWorkbook.path & "\" & "settings.ini"
    SetEmailOption = SetINIString("Email", "Option", sVal, path)
End Function

Public Function GetEmailOption() As String
    Dim path As String
    path = ThisWorkbook.path & "\" & "settings.ini"
    GetEmailOption = GetINIString("Email", "Option", path)
End Function
Public Function SetSAPOption(ByVal sVal As String) As String
    Dim path As String
    path = ThisWorkbook.path & "\" & "settings.ini"
    SetSAPOption = SetINIString("SAP", "Option", sVal, path)
End Function

Public Function GetSAPOption() As String
    Dim path As String
    path = ThisWorkbook.path & "\" & "settings.ini"
    GetSAPOption = GetINIString("SAP", "Option", path)
End Function
Public Function SetReportFilePath(ByVal sVal As String) As String
    Dim path As String
    path = ThisWorkbook.path & "\" & "settings.ini"
    SetReportFilePath = SetINIString("ReportPath", "Option", sVal, path)
End Function
Public Function SetASCFilePath(ByVal sVal As String) As String
    Dim path As String
    path = ThisWorkbook.path & "\" & "settings.ini"
    SetASCFilePath = SetINIString("ASCPath", "Option", sVal, path)
End Function
Public Function SetContactsFilePath(ByVal sVal As String) As String
    Dim path As String
    path = ThisWorkbook.path & "\" & "settings.ini"
    SetContactsFilePath = SetINIString("ContactsPath", "Option", sVal, path)
End Function
Public Function SetDATEFORMAT(ByVal sVal As String) As String
    Dim path As String
    path = ThisWorkbook.path & "\" & "settings.ini"
    SetDATEFORMAT = SetINIString("DATEformat", "Option", sVal, path)
End Function
Public Function SetChosenScript(ByVal sVal As String) As String
    Dim path As String
    path = ThisWorkbook.path & "\" & "settings.ini"
    SetChosenScript = SetINIString("ChosenScript", "Option", sVal, path)
End Function
Public Function GetReportFilePath() As String
    Dim path As String
    path = ThisWorkbook.path & "\" & "settings.ini"
    GetReportFilePath = GetINIString("ReportPath", "Option", path)
End Function
Public Function GetASCFilePath() As String
    Dim path As String
    path = ThisWorkbook.path & "\" & "settings.ini"
    GetASCFilePath = GetINIString("ASCPath", "Option", path)
End Function
Public Function GetContactsFilePath() As String
    Dim path As String
    path = ThisWorkbook.path & "\" & "settings.ini"
    GetContactsFilePath = GetINIString("ContactsPath", "Option", path)
End Function
Public Function GetDATEFORMAT() As String
    Dim path As String
    path = ThisWorkbook.path & "\" & "settings.ini"
    GetDATEFORMAT = GetINIString("DATEformat", "Option", path)
End Function
Public Function Getchosenscript() As String
    Dim path As String
    path = ThisWorkbook.path & "\" & "settings.ini"
    Getchosenscript = GetINIString("ChosenScript", "Option", path)
End Function
Public Function GetNetworkDrive() As String
    Dim path As String
    path = ThisWorkbook.path & "\" & "settings.ini"
    GetNetworkDrive = GetINIString("Ndrive", "Option", path)
End Function
Public Function SetNetworkDrive(ByVal sVal As String) As String
    Dim path As String
    path = ThisWorkbook.path & "\" & "settings.ini"
    SetNetworkDrive = SetINIString("Ndrive", "Option", sVal, path)
End Function

' To be copied from here downwards

Public Function setHideConvertWarning(ByVal sVal As Integer) As String
    Dim path As String
    path = ThisWorkbook.path & "\settings.ini"
    setHideConvertWarning = SetINIString("HideConvertWarning", "Option", sVal, path)
End Function

Public Function getHideConvertWarning() As String
    Dim path As String
    path = ThisWorkbook.path & "\settings.ini"
    getHideConvertWarning = GetINIString("HideConvertWarning", "Option", path)
End Function
