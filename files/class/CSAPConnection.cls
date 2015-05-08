VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSAPConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private SapGuiApp As Object
Private Connection As GuiConnection
Public WithEvents session As GuiSession
Attribute session.VB_VarHelpID = -1
Private Wscript As Variant
Private SapGuiAuto As Object
Private CLPMUSProfile As String
Private rotEntry As Object
Private WithEvents application2 As GuiApplication
Attribute application2.VB_VarHelpID = -1
Private sessionCount As Integer
Private pErrorCounter As Integer ' för property endast (finns redan)
Private pPrcObj As Integer ' för property endast
Private pElpTime As Double ' för property endast
Private trx As String


Private Sub Class_Initialize()
    pErrorCounter = 0
End Sub

Sub absorbConnection()
    On Error Resume Next
    Dim pSessionCount

If GetSAPOption <> "Yes" Then
'NEW
    Dim Session_number_max As String

    Session_number_max = 6                                  'the maximum number of possible SAP sessions
Dim session_number_(6)

Set SapGuiAuto = GetObject("SAPGUI")
Set SAP = SapGuiAuto.GetScriptingEngine
Set Connection = SAP.Children(0)
Set session = Connection.Children(0)
    Set rotEntry = CreateObject("SAPGUI")
    If Err.Number <> 0 Then
        Debug.Print Err.Number & ": " & Err.Description
        GoTo EndSub
    End If
    Set application2 = rotEntry.GetScriptingEngine
    Set Connection1 = application2.Connections.ElementAt(0)
    Set Session1 = Connection.Children.ElementAt(0)

'------------------------------------------- new session connect -------------------------------------------------------------
session_number_all = Connection.Children.Count - 1

For i = 1 To Session_number_max
     session_number_(i) = 0
Next

For session_number = 0 To session_number_all
     Set session = Connection.Children(Int(session_number))
     session_number_(session.Info.SessionNumber) = session.Info.SessionNumber
Next

If session_number_all < Session_number_max - 1 Then
     session.CreateSession
     Do
      Wscript.Sleep 500
      If Connection.Children.Count - session_number_all >= 2 Then Exit Do
     Loop
     On Error Resume Next
     Error_number = 1
     For session_number = 0 To session_number_all + 1
          Err.Clear
          Set session = Connection.Children(Int(session_number))
          If Err.Number > 0 Or Err.Number < 0 Then Exit For

         If session_number_(session.Info.SessionNumber) = 0 Then
            Error_number = 0
            Exit For
         End If
         'session.findById("wnd[0]").iconify                          'So you can send another SAP session in the task bar.
     Next
     On Error GoTo 0
Else
 MsgBox "You've reached the maximum number of SAP windows (6), you must close one window, after that click RUN again.", vbCritical
    Application.DisplayFullScreen = False
 End
End If
    pSessionCount = Connection1.Sessions.Count
    If pSessionCount = 0 Then
        killConnection
        MsgBox "You must log in to SAP", vbCritical
        Application.DisplayFullScreen = False
        End
    End If
    
  Else
    'OLD
    
    
    Set rotEntry = CreateObject("SAPGUI")
    If Err.Number <> 0 Then
        Debug.Print Err.Number & ": " & Err.Description
        GoTo EndSub
    End If
    Set application2 = rotEntry.GetScriptingEngine
    Set Connection = application2.Connections.ElementAt(0)
    Set session = Connection.Children.ElementAt(0)
    If Not IsObject(Application) Then
        Set SapGuiAuto = GetObject("SAPGUI")
    End If
    If Not IsObject(Connection) Then
        Set Connection = Application.Children(0)
    End If
    If Not IsObject(session) Then
        Set session = Connection.Children(0)
    End If
    If IsObject(Wscript) Then
        Wscript.ConnectObject session, "on"
        Wscript.ConnectObject Application, "on"
    End If
    pSessionCount = Connection.Sessions.Count
    If pSessionCount = 0 Then
        killConnection
        MsgBox "You must log in to SAP", vbCritical
        Application.DisplayFullScreen = False
        End
    End If
    End If
Exit Sub
EndSub:
        MsgBox "You must log in to SAP", vbCritical
Application.DisplayFullScreen = False

End
End Sub

Sub killConnection()
    Set SapGuiApp = Nothing
    Set Connection = Nothing
    Set session = Nothing
    Set Wscript = Nothing
    Set SapGuiAuto = Nothing
    Set rotEntry = Nothing
    Set application2 = Nothing
    Set SAP = Nothing
    Set session = Nothing
End Sub
' For stats-file (amount of errors)
Public Property Get ErrorCounter() As Integer
    ErrorCounter = pErrorCounter
End Property
Public Property Let ErrorCounter(value As Integer)
    pErrorCounter = value
End Property


Sub reportCrash(ByRef crashItem As String)
    On Error GoTo FinishCrashMail
    Dim oApp As Object
    Dim oMail As Object
    Set oApp = CreateObject("Outlook.Application")
    Set oMail = oApp.Createitem(0)

    oMail.subject = "Auto crash report from PFM SmartApp"
    oMail.To = Application.UserName
    oMail.body = "The macro has crashed at " & crashItem & "." & vbNewLine & _
    "/ PFM SmartApp file"
    oMail.Send
    
FinishCrashMail:
    Set oApp = Nothing
    Set oMail = Nothing
End Sub

Sub errorContinueNextItem(ByRef trx As String)
    ActiveCell.value = 0
    If Me.session.findById("wnd[0]/sbar").Text = "" Then
    Cells(ActiveCell.Row, ActiveCell.End(xlToRight).column + 1).value = "Error updating object, email sent to SmartApp Analyst."
    Else
    Cells(ActiveCell.Row, ActiveCell.End(xlToRight).column + 1).value = "ERROR: " & Me.session.findById("wnd[0]/sbar").Text & ", email sent to SmartApp Analyst."
    End If
    Me.session.findById("wnd[0]/tbar[0]/okcd").Text = "/n" & trx
    Me.session.findById("wnd[0]").sendVKey 0
End Sub

Sub logStatitics(ByRef chosenScript As String, ByVal prcObj As Integer, ByRef finishTime As Double, ByRef elpTime As Double)
    On Error GoTo ErrStats

    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile("\\esekina005\groupfbs\SmartApp\Excel\LOG\scriptruns.log", 8, 0)
    f.WriteLine chosenScript & " |" & prcObj & " |" & Application.UserName & " |" & Me.ErrorCounter & " |" & Format(finishTime, "yyyy-mm-dd") & " |" & Format(finishTime, "hh:mm") & " |" & Round(elpTime, 0) & " |" & "4." & Format(FileDateTime("C:\Users\" & Environ("Username") & "\AppData\Roaming\Microsoft\AddIns\PFM_SmartApp.xlam"), "YYMMDDHHSS") & " | EAB"
    f.Close
    Set fso = Nothing
    Set f = Nothing
    Exit Sub
ErrStats:
    Me.statsCannotBeWritten
End Sub

Public Sub statsCannotBeWritten()
    On Error GoTo ErrStatsError
    Dim oApp As Object
    Dim oMail As Object
    Set oApp = CreateObject("Outlook.Application")
    Set oMail = oApp.Createitem(0)
    
    oMail.subject = "Auto error report from PFM SmartApp"
    oMail.To = "eedgcan"
    oMail.body = Application.UserName & " is trying to update the statistics file but it is not possible." & vbNewLine & _
    "/ PFM SmartApp file"
    oMail.Send
    Set oApp = Nothing
    Set oMail = Nothing
    Exit Sub
ErrStatsError:
End Sub
