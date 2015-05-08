VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function FileToMD5Hex(sFileName As String) As String
    Dim enc
    Dim bytes
    Dim outstr As String
    Dim pos As Integer
    Set enc = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    'Convert the string to a byte array and hash it
    bytes = GetFileBytes(sFileName)
    bytes = enc.ComputeHash_2((bytes))
    'Convert the byte array to a hex string
    For pos = 1 To LenB(bytes)
        outstr = outstr & LCase(Right("0" & Hex(AscB(MidB(bytes, pos, 1))), 2))
    Next
    FileToMD5Hex = outstr
    Set enc = Nothing
End Function

Private Function GetFileBytes(ByVal path As String) As Byte()
    Dim lngFileNum As Long
    Dim bytRtnVal() As Byte
    lngFileNum = FreeFile
    If LenB(Dir(path)) Then ''// Does file exist?
        Open path For Binary Access Read As lngFileNum
        ReDim bytRtnVal(LOF(lngFileNum) - 1&) As Byte
        Get lngFileNum, , bytRtnVal
        Close lngFileNum
    Else
        Err.Raise 53
    End If
    GetFileBytes = bytRtnVal
    Erase bytRtnVal
End Function

