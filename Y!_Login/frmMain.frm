VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Y_Login"
   ClientHeight    =   1815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2505
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   2505
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Sock 
      Left            =   2040
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSendMessage 
      Caption         =   "Send Message"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdStartLogin 
      Caption         =   "Start Login"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "PW"
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "ID"
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   2295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Yahoo YMSG Protocol example by Arub
'packets based off yahoo messenger 7.0.0.426
'www.arubs.net
Dim strSessionID As String


Private Sub cmdSendMessage_Click()
    Dim strID As String, strMessage As String
    strID = InputBox("ID to send to:")
    strMessage = InputBox("Message to send:")
    
    SendPacket 79, ChrH("00 00 00 00"), ChrH("31 C0 80") & txtID.Text & ChrH("C0 80 34 C0 80") & txtID.Text & _
    ChrH("C0 80 31 32 C0 80 4D 6A 41 77 4E 7A 59 31 4D 54 49 33 4E 51 3D 3D C0 80 36 31 C0 80 30 C0 80 32 C0 80") & _
    ChrH("C0 80 35 C0 80") & strID & ChrH("C0 80 31 33 C0 80 30 C0 80 34 39 C0 80 50 45 45 52 54 4F 50 45 45 52 C0 80")
    'packet detailing what kind of message
    
    SendPacket 6, ChrH("5A 55 AA 56"), ChrH("31 C0 80") & txtID.Text & _
    ChrH("C0 80 35 C0 80") & strID & ChrH("C0 80 31 34 C0 80") & strMessage & _
    ChrH("C0 80 39 37 C0 80 31 C0 80 36 33 C0 80 3B 30 C0 80 36 34 C0 80 30 C0 80 32 30 36 C0 80 30 C0 80")
    'actual message packet
    
End Sub
Public Function SendPacket(lngCommand As Long, strStatus As String, strData As String) 'send a packet
    
    If Sock.State <> sckConnected Then Exit Function
    Sock.SendData "YMSG" & _
    ChrH("00 0D 00 00") & _
    Word(Len(strData)) & Word(lngCommand) & _
    strStatus & strSessionID & _
    strData

End Function
Private Sub Sock_DataArrival(ByVal bytesTotal As Long)
    Dim strDatas As String
    Sock.GetData strDatas
    Call ParseData(strDatas)
End Sub

Private Function ParseData(strData As String)
    Dim A1, A2, A3, A4, A5, A6
    If Len(strData) < 20 Then Exit Function
    Select Case AscToHex(Mid(strData, 11, 2)) 'grab command
    
        Case "00 06" 'incoming message
            
            A1 = InStr(1, strData, ChrH("31 C0 80"))
            A2 = InStr(A1 + 3, strData, ChrH("C0 80"))
            A3 = Mid(strData, A1 + 3, A2 - A1 - 3) 'username
        
            A4 = InStr(1, strData, ChrH("C0 80 31 34 C0 80"))
            A5 = InStr(A4 + 6, strData, ChrH("C0 80"))
            A6 = Mid(strData, A4 + 6, A5 - A4 - 6) 'message
            
            MsgBox A3 & " says " & A6, vbInformation, "Message"
        
        Case "00 4C" 'replied to inital login command
        
            SendPacket 87, ChrH("00 00 00 00"), ChrH("31 C0 80") & txtID.Text & ChrH("C0 80")
        
        Case "00 55" 'logged in
                   
            lblStatus.Caption = "Status: Logged In"
            
        Case "00 57" 'replied with challange for login
        
            Dim strChallange As String, Encrypted As Variant
            A1 = InStr(1, strData, ChrH("39 34 C0 80")) 'challange start
            A2 = InStr(A1 + 4, strData, ChrH("C0 80")) 'challange finish
            strChallange = Mid$(strData, A1 + 4, A2 - A1 - 4)  'grab the challange
            
            Encrypted = Split(EncryptPassword(txtID.Text, txtPassword.Text, strChallange), ":", 2)  'encrypt password with challange
            
            SendPacket 84, ChrH("5A 55 AA 55"), ChrH("36 C0 80") & Encrypted(0) & ChrH("C0 80 39 36 C0 80") & _
            Encrypted(1) & ChrH("C0 80 30 C0 80") & txtID.Text & ChrH("C0 80 32 C0 80") & _
            txtID.Text & ChrH("C0 80 32 C0 80 31 C0 80 31 C0 80") & txtID.Text & _
            ChrH("C0 80 32 34 34 C0 80 35 32 34 32 32 33 C0 80 31 33 35 C0 80 37 2C 30 2C 30 2C 34 32 36 C0 80 31 34 38 C0 80 33 36 30 C0 80 35 39 C0 80 42 09 35 6C 72 66 63 63 6C 31 6B 68 65 68 39 26 62 3D 32 C0 80")
        
        Case Else
    End Select
End Function
Private Sub Sock_Connect()
    'initialise login
    SendPacket 76, ChrH("00 00 00 00"), ""
End Sub
Private Sub cmdStartLogin_Click()
    lblStatus.Caption = "Status: Logging in.."
    Sock.Close
    Sock.Connect "scs.msg.yahoo.com", 5050
    strSessionID = ChrH("00 00 00 00") 'apparently yahoo is retarded and you can just use 4 null bytes as the session id and it will still work
End Sub

Private Sub Sock_Close()
    lblStatus.Caption = "Status: Logged Off"
End Sub


