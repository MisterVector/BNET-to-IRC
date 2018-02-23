VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BNET To IRC Daemon"
   ClientHeight    =   6075
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   14265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   14265
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrReleaseQueue 
      Enabled         =   0   'False
      Interval        =   1250
      Left            =   2160
      Top             =   0
   End
   Begin MSWinsockLib.Winsock sckBNLS 
      Index           =   0
      Left            =   3120
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckBNET 
      Index           =   0
      Left            =   3600
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckIRC 
      Left            =   2640
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Internet Relay Chat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   7200
      TabIndex        =   1
      Top             =   120
      Width           =   6975
      Begin VB.TextBox txtIRCChannel 
         Height          =   285
         Left            =   1680
         TabIndex        =   27
         Top             =   720
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         Caption         =   "Display Mode"
         Height          =   555
         Left            =   3360
         TabIndex        =   20
         Top             =   240
         Width           =   2415
         Begin VB.OptionButton rcChat 
            Caption         =   "Chat"
            Height          =   255
            Left            =   1320
            TabIndex        =   22
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton rcConsole 
            Caption         =   "Console"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.CommandButton btnConnectIRC 
         Caption         =   "Connect!"
         Height          =   255
         Left            =   3360
         TabIndex        =   18
         Top             =   840
         Width           =   2415
      End
      Begin VB.CheckBox chkBtoBNET 
         Caption         =   "Broadcast to BNET"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   15
         Top             =   5520
         Width           =   2055
      End
      Begin VB.TextBox txtIRCChat 
         Height          =   285
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   5520
         Width           =   4575
      End
      Begin VB.TextBox txtIRCUsername 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtIRCServer 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   1080
         Width           =   1575
      End
      Begin RichTextLib.RichTextBox rtbChatIRCConsole 
         Height          =   3855
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   6800
         _Version        =   393217
         BackColor       =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"Form1.frx":0000
      End
      Begin RichTextLib.RichTextBox rtbChatIRCChat 
         Height          =   3855
         Left            =   120
         TabIndex        =   25
         Top             =   1560
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   6800
         _Version        =   393217
         BackColor       =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"Form1.frx":0082
      End
      Begin VB.Label Label5 
         Caption         =   "Channel:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "IRC Username:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Server: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Battle.Net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.TextBox txtChannel 
         Height          =   285
         Left            =   3840
         TabIndex        =   24
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton btnConnectBNET 
         Caption         =   "Connect!"
         Height          =   300
         Left            =   1200
         TabIndex        =   19
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CheckBox chkBtoIRC 
         Caption         =   "Broadcast to IRC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   16
         Top             =   5520
         Width           =   1815
      End
      Begin VB.TextBox txtBNETChat 
         Height          =   285
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   5520
         Width           =   4815
      End
      Begin VB.ComboBox cmbServer 
         Height          =   315
         ItemData        =   "Form1.frx":0104
         Left            =   3840
         List            =   "Form1.frx":0106
         TabIndex        =   10
         Text            =   "Select Server"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
      Begin RichTextLib.RichTextBox rtbChatBNET 
         Height          =   3855
         Left            =   120
         TabIndex        =   13
         Top             =   1560
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   6800
         _Version        =   393217
         BackColor       =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"Form1.frx":0108
      End
      Begin VB.Label Label1 
         Caption         =   "Channel:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   23
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Server:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   17
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Username: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuKeys 
         Caption         =   "Manage Keys"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkDisplayRTB_Click()
  If chkDisplayRTB.value = 1 Then
    rtbChatIRCConsole.Visible = False
    rtbChatIRCChat.Visible = True
    chkDisplayRTB.value = 0
  Else
    rtbChatIRCConsole.Visible = True
    rtbChatIRCChat.Visible = False
    chkDisplayRTB.value = 1
  End If
End Sub

Private Sub btnConnectBNET_Click()
  Dim socketsStillAlive As Boolean, username As String, password As String, channel As String, bnetServer As String

  If botCount = 0 Then
    MsgBox "Your keys are not configured. Go to File -> Manage Keys first."
    Exit Sub
  End If
  
  username = txtUsername.text
  password = txtPassword.text
  channel = txtChannel.text
  bnetServer = cmbServer.text
  
  If btnConnectBNET.Caption = "Connect!" Then
    
    AddChat rtbChatBNET, vbYellow, "Bot #0: [BNET] Connecting..."
    sckBNET(0).Connect cmbServer.text, 6112
    
    btnConnectBNET.Caption = "Disconnect!"
  Else
    btnConnectBNET.Caption = "Connect!"
    
    For i = 0 To sckBNET.Count - 1
      sckBNET(i).Close
      
      If sckBNLS(i).State <> sckClosed Then
        If Not socketsStillAlive Then socketsStillAlive = True
        sckBNLS(i).Close
      End If
    Next i
    
    If socketsStillAlive Then
      AddChat rtbChatBNET, vbRed, "[BNLS] All connections closed."
    End If
    
    AddChat rtbChatBNET, vbRed, "[BNET] All connections closed."
  End If
End Sub

Private Sub btnConnectIRC_Click()
  If txtIRCServer.text = vbNullString Then
    MsgBox "You have not entered a server name!"
    Exit Sub
  End If

  If txtIRCUsername.text = vbNullString Then
    MsgBox "No username was entered!"
    Exit Sub
  End If
  
  If btnConnectIRC.Caption = "Connect!" Then
    IRC.username = txtIRCUsername.text
    
    If InStr(txtIRCServer.text, ":") Then
      IRC.Server = Split(txtIRCServer.text, ":")(0)
      IRC.Port = Split(txtIRCServer.text, ":")(1)
    Else
      IRC.Server = txtIRCServer.text
      IRC.Port = 6667
    End If
  
    btnConnectIRC.Caption = "Disconnect!"
    AddChat rtbChatIRCConsole, vbYellow, "[IRC] Connecting to " & IRC.Server & ":" & IRC.Port & "..."
    sckIRC.Connect IRC.Server, IRC.Port
  Else
    AddChat rtbChatIRCConsole, vbRed, "[IRC] All connectiosn closed."
    btnConnectIRC.Caption = "Connect!"
    If sckIRC.State = sckConnected Then
      'SendToBNET "Disconnected from " & IRC.Server & "!"
      SendToBNET "Disconnected from IRC!"
      sckIRC.SendData "QUIT"
      DoEvents: DoEvents: DoEvents: DoEvents
    End If
    sckIRC.Close

  End If
End Sub

Private Sub chkBtoBNET_Click()
  If chkBtoBNET.value = 1 Then
    isBroadcastToBNET = True
  Else
    isBroadcastToBNET = False
  End If
End Sub

Private Sub chkBtoIRC_Click()
  If chkBtoIRC.value = 1 Then
    isBroadcastToIRC = True
  Else
    isBroadcastToIRC = False
  End If
End Sub

Private Sub Form_Load()
  Dim val As Variant, arrGateways() As Variant, gateway As String, IPs() As String

  val = ReadINI("Main", "Top", "Config.ini")

  If (IsNumeric(val)) Then
    Me.Top = val
  End If

  val = ReadINI("Main", "Left", "Config.ini")

  If (IsNumeric(val)) Then
    Me.Left = val
  End If

  txtUsername.text = ReadINI("Main", "Username", "Config.ini")
  txtPassword.text = ReadINI("Main", "Password", "Config.ini")
  txtChannel.text = ReadINI("Main", "Channel", "Config.ini")
  
  cmbServer.text = ReadINI("Main", "Server", "Config.ini")
  BNLSServer = ReadINI("Main", "BNLSServer", "Config.ini")

  If IsNumeric(ReadINI("Main", "BotCount", "Config.ini")) Then
    botCount = ReadINI("Main", "BotCount", "Config.ini")
    
    If (botCount > 0) Then
      ReDim pBNET(botCount - 1)
      ReDim pBNLS(botCount - 1)
      ReDim BNET(botCount - 1)
      
      For i = 0 To botCount - 1
        If i > 0 Then
          Load sckBNET(i)
          Load sckBNLS(i)
        End If

        Set pBNET(i) = New clsPacket
        Set pBNLS(i) = New clsPacket
        With BNET(i)
          .prodStr = ReadINI(i, "Product", "Config.ini")
          .CDKey = ReadINI(i, "CDKey", "Config.ini")
        End With
      Next i
    End If
  End If
  
  txtIRCUsername.text = ReadINI("IRC", "Username", "Config.ini")
  txtIRCServer.text = ReadINI("IRC", "Server", "Config.ini")
  txtIRCChannel.text = ReadINI("IRC", "Channel", "Config.ini")
  rcConsole.value = True

  arrGateways = Array("uswest.battle.net", "useast.battle.net", "europe.battle.net", "asia.battle.net")

  For i = 0 To 3
    gateway = arrGateways(i)
    cmbServer.AddItem gateway
    IPs = Split(Resolve(gateway))

    For j = 0 To UBound(IPs)
      cmbServer.AddItem IPs(j)
    Next j

    If (i < 3) Then
      cmbServer.AddItem ""
    End If
  Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Dir$(App.Path & "\Config.ini") <> vbNullString Then
    Kill App.Path & "\Config.ini"
  End If

  WriteINI "Main", "Top", Me.Top, "Config.ini"
  WriteINI "Main", "Left", Me.Left, "Config.ini"

  WriteINI "Main", "Username", txtUsername.text, "Config.ini"
  WriteINI "Main", "Password", txtPassword.text, "Config.ini"
  WriteINI "Main", "Channel", txtChannel.text, "Config.ini"
  WriteINI "Main", "Server", cmbServer.text, "Config.ini"
  WriteINI "Main", "BotCount", botCount, "Config.ini"

  For i = 0 To botCount - 1
    WriteINI i, "Product", BNET(i).prodStr, "Config.ini"
    WriteINI i, "CDKey", BNET(i).CDKey, "Config.ini"
  Next i
  
  WriteINI "IRC", "Username", txtIRCUsername.text, "Config.ini"
  WriteINI "IRC", "Server", txtIRCServer.text, "Config.ini"
  WriteINI "IRC", "Channel", txtIRCChannel.text, "Config.ini"

  Dim oFrm As Form

  For Each oFrm In Forms
    Unload oFrm
  Next
End Sub

Private Sub mnuKeys_Click()
  frmKeys.Show
End Sub

Private Sub rcChat_Click()
  rtbChatIRCConsole.Visible = False
  rtbChatIRCChat.Visible = True
End Sub

Private Sub rcConsole_Click()
  rtbChatIRCConsole.Visible = True
  rtbChatIRCChat.Visible = False
End Sub

Private Sub sckBNET_Connect(index As Integer)
  AddChat rtbChatBNET, vbGreen, "Socket #" & index & ": [BNET] Connected!"
  sckBNET(index).SendData Chr$(1)
  Send0x50 index
End Sub


Private Sub sckBNET_DataArrival(index As Integer, ByVal bytesTotal As Long)
  Dim data As String, pLen As Long, pID As Byte
  sckBNET(index).GetData data
  
  Do While Len(data) > 0
    pID = Asc(Mid(data, 2, 1))
    CopyMemory pLen, ByVal Mid$(data, 3, 2), 2
    pBNET(index).SetData Mid(data, 5)
    
    Select Case pID
      Case &HA: Recv0x0A index
      Case &HF: Recv0x0F index
      Case &H25: Send0x25 index
      Case &H3A: Recv0x3A index
      Case &H3D: Recv0x3D index
      Case &H50: Recv0x50 index
      Case &H51: Recv0x51 index
    End Select
    
    data = Mid(data, pLen + 1)
  Loop
End Sub

Private Sub sckBNET_Error(index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  AddChat rtbChatBNET, vbRed, "Bot #" & index & " error #" & Number & ": " & Description
End Sub

Private Sub sckBNLS_Connect(index As Integer)
  If newAccFlag Then
    With pBNLS(index)
      .InsertDWORD Len(password)
      .InsertDWORD &H4
      .InsertNonNTString password
      .InsertDWORD &H0
      .sendPacket &HB, True, index
    End With
  Else
    AddChat rtbChatBNET, vbGreen, "Bot #" & index & ": [BNLS] Connected!"
    Send_BNLS_0x01 index
  End If
End Sub

Public Sub Click_start()
  btnConnectBNET_Click
End Sub

Private Sub sckBNLS_DataArrival(index As Integer, ByVal bytesTotal As Long)
  Dim data As String, pID As Byte, pLen As Long
  sckBNLS(index).GetData data
  
  Do While Len(data) > 0
    CopyMemory pLen, ByVal Mid$(data, 1, 2), 2
    pID = Asc(Mid(data, 3, 1))
    pBNLS(index).SetData Mid(data, 4)
    
    Select Case pID
      Case &H1: Recv_BNLS_0x01 index
      Case &H9: Recv_BNLS_0x09 index
      Case &HB: Recv_BNLS_0x0B index
    End Select
    
    data = Mid(data, pLen + 1)
  Loop
End Sub

Private Sub sckIRC_Close()
  'SendToBNET "Disconnected from the IRC server at " & IRC.Server & "!"
  SendToBNET "Disconnected from IRC!"
End Sub

Private Sub sckIRC_Connect()
  AddChat rtbChatIRCConsole, vbGreen, "[IRC] Connected!"
  sckIRC.SendData "NICK " & IRC.username & vbCrLf
  sckIRC.SendData "USER " & IRC.username & " 0 0 " & IRC.username & vbCrLf
  'SendToBNET "Connected to the IRC server at " & IRC.Server & "!"
  SendToBNET "Connected to IRC!"
End Sub

Private Sub sckIRC_DataArrival(ByVal bytesTotal As Long)
  Dim data As String, arrData() As String, name As String, text As String, channel As String
  sckIRC.GetData data
  
  If InStr(data, IRC.username) Then
    data = Mid(data, InStr(data, IRC.username) + Len(IRC.username) + 1)
  End If
    
  If UBound(Split(data)) > 1 Then
    arrData = Split(data)
    
    Select Case UCase(arrData(1))
      Case "PRIVMSG"
        name = Mid(Split(arrData(0), "!")(0), 2)
        text = Split(data, arrData(1))(1)
        text = Replace(Mid(text, InStr(text, ":") + 1), vbCrLf, vbNullString)
        'text = Replace(Split(Split(data, arrData(1))(1), ":")(1), vbCrLf, vbNullString)
        channel = Split(arrData(2), ":")(0)
  
        If isBroadcastToBNET Then
          'SendToBNET "(" & IRC.Channel & " @ " & IRC.Server & ") " & getName & ": " & getText
          SendToBNET name & ": " & text
        End If
      Case Else
        AddChat rtbChatIRCConsole, vbYellow, data
    End Select
  Else
    AddChat rtbChatIRCConsole, vbYellow, data
  
    If Left(data, 5) = "PING " Then
      AddChat rtbChatIRCConsole, vbWhite, "PING has been PONG'D"
      sckIRC.SendData "PONG " & Mid(data, 6) & vbCrLf
      Exit Sub
    End If
  End If
End Sub

Private Sub tmrReleaseQueue_Timer()
  Dim qMsg As String
  qMsg = dicQueue.Item(dicIdx)
  
  With pBNET(bIdx)
    .InsertNTString qMsg
    .sendPacket &HE, False, bIdx
  End With
  
  bIdx = bIdx + 1
  If bIdx = sckBNET.Count Then bIdx = 0
  
  dicIdx = dicIdx + 1
  If dicIdx > dicQueue.Count Then
    dicIdx = 1
    
    dicQueue.RemoveAll
    tmrReleaseQueue.Enabled = False
  End If
End Sub

Private Sub txtBNETChat_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    If sckBNET(cIdx).State = sckConnected Then
      txtBNETChat.text = Replace(txtBNETChat.text, vbNewLine, "")
      AddChat rtbChatBNET, vbYellow, "Bot #" & cIdx & ": <" & username & "> ", vbWhite, txtBNETChat.text
      SendToBNET txtBNETChat.text
      txtBNETChat.text = vbNullString
    End If
    
    cIdx = cIdx + 1
    
    If cIdx = sckBNET.Count Then cIdx = 0
  End If
End Sub

Private Sub txtIRCChat_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim text As String, cmd() As String, cmdEx() As String
  
  If KeyCode = 13 Then
    text = txtIRCChat.text
    text = Replace(text, vbNewLine, vbNullString)
    txtIRCChat.text = vbNullString
    
    If Left(text, 1) = "/" Then
      cmd = Split(Mid(text, 2))
      
      Select Case LCase(cmd(0))
        Case "join"
          cmdEx = Split(text, " ", 2)
          IRC.channel = cmdEx(1)
          sckIRC.SendData "JOIN " & cmdEx(1) & vbCrLf
      End Select
    Else
      If sckIRC.State <> sckConnected Then Exit Sub

      sckIRC.SendData "PRIVMSG " & IRC.channel & " :" & text & vbCrLf
      AddChat rtbChatIRCChat, vbWhite, IRC.Server & " (", vbYellow, IRC.channel, vbWhite, ") " & text
    End If
  End If
End Sub
