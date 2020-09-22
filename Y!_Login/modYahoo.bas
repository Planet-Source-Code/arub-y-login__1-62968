Attribute VB_Name = "modYahoo"
Option Explicit
Private Declare Function YMSG12_ScriptedMind_Encrypt Lib "YMSG12ENCRYPT.dll" (ByVal username As String, ByVal Password As String, ByVal Seed As String, ByVal result_6 As String, ByVal result_96 As String, intt As Long) As Boolean

Public Function EncryptPassword(strID As String, strPassword As String, strChallange As String)
    'encrypt password
    Dim strEncrypted1 As String, strEncrypted2 As String
    strEncrypted1 = String(50, vbNullChar): strEncrypted2 = String(50, vbNullChar)
    Call YMSG12_ScriptedMind_Encrypt(strID, strPassword, strChallange, strEncrypted1, strEncrypted2, 1)
    EncryptPassword = strEncrypted1 & ":" & strEncrypted2
End Function
Public Function ChrH(strString) As String
    'hex to ascii
    Dim A1
    A1 = Split(strString, " ")
    Dim i As Integer
    For i = 0 To UBound(A1)
        ChrH = ChrH & Chr("&H" & A1(i))
        DoEvents
    Next i
End Function
Public Function AscToHex(strString As String) As String
    'ascii to hex
    Dim i As Integer
    For i = 1 To Len(strString)
        If Len(Hex(Asc(Mid(strString, i, 1)))) = 1 Then
            AscToHex = AscToHex & "0" & Hex(Asc(Mid(strString, i, 1)))
        Else
            AscToHex = AscToHex & Hex(Asc(Mid(strString, i, 1)))
        End If
        If i <> Len(strString) Then AscToHex = AscToHex & " "
        DoEvents
    Next i
End Function

'````````````````````````````````````````````````````````````

'Put a value up to 65535 into this, and get a 2 byte integer
Public Function Word(ByVal lngVal As Long) As String
'by Xeon
    Dim Lo As Single
    Dim Hi As Single

    Lo = Fix(lngVal / 256)
    Hi = lngVal Mod 256

    Word = Chr(Lo) & Chr(Hi)
End Function
