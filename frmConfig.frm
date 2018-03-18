VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Battle.Net to IRC Configuration"
   ClientHeight    =   6075
   ClientLeft      =   1875
   ClientTop       =   1995
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ilIcons 
      Left            =   2400
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConfig.frx":0000
            Key             =   "W2BN"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConfig.frx":04EA
            Key             =   "D2DV"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConfig.frx":09D4
            Key             =   "WAR3"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   5175
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Battle.Net"
      TabPicture(0)   =   "frmConfig.frx":0EBE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label11"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label12"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label13"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label14"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label15"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtBNETChannel"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtBNLSServer"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmbBNETServer"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtBNETUsername"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtBNETPassword"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtW2BNVerByte"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtD2DVVerByte"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtWAR3VerByte"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Key Manager"
      TabPicture(1)   =   "frmConfig.frx":0EDA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "opWAR3"
      Tab(1).Control(1)=   "opD2DV"
      Tab(1).Control(2)=   "opW2BN"
      Tab(1).Control(3)=   "txtBNETKey"
      Tab(1).Control(4)=   "btnRemove"
      Tab(1).Control(5)=   "btnAdd"
      Tab(1).Control(6)=   "lvKeyList"
      Tab(1).Control(7)=   "Label2"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "IRC"
      TabPicture(2)   =   "frmConfig.frx":0EF6
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtIRCServer"
      Tab(2).Control(1)=   "txtIRCChannel"
      Tab(2).Control(2)=   "txtIRCUsername"
      Tab(2).Control(3)=   "Label10"
      Tab(2).Control(4)=   "Label7"
      Tab(2).Control(5)=   "Label8"
      Tab(2).Control(6)=   "Label9"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Miscellaneous"
      TabPicture(3)   =   "frmConfig.frx":0F12
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "chkRememberWindowPosition"
      Tab(3).ControlCount=   1
      Begin VB.TextBox txtWAR3VerByte 
         Height          =   375
         Left            =   1320
         TabIndex        =   36
         Top             =   4440
         Width           =   735
      End
      Begin VB.TextBox txtD2DVVerByte 
         Height          =   405
         Left            =   3240
         TabIndex        =   35
         Top             =   3960
         Width           =   735
      End
      Begin VB.TextBox txtW2BNVerByte 
         Height          =   375
         Left            =   1320
         TabIndex        =   34
         Top             =   3960
         Width           =   735
      End
      Begin VB.CheckBox chkRememberWindowPosition 
         Caption         =   "Remember Window Position"
         Height          =   375
         Left            =   -74760
         TabIndex        =   29
         Top             =   600
         Width           =   3615
      End
      Begin VB.OptionButton opWAR3 
         Caption         =   "Warcraft III"
         Height          =   255
         Left            =   -72600
         TabIndex        =   28
         Top             =   3720
         Width           =   1215
      End
      Begin VB.OptionButton opD2DV 
         Caption         =   "Diablo II"
         Height          =   255
         Left            =   -73560
         TabIndex        =   27
         Top             =   3720
         Width           =   975
      End
      Begin VB.OptionButton opW2BN 
         Caption         =   "Warcraft II"
         Height          =   255
         Left            =   -74640
         TabIndex        =   26
         Top             =   3720
         Width           =   1095
      End
      Begin VB.TextBox txtIRCServer 
         Height          =   345
         Left            =   -73320
         TabIndex        =   20
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox txtIRCChannel 
         Height          =   345
         Left            =   -73320
         TabIndex        =   19
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox txtIRCUsername 
         Height          =   345
         Left            =   -73320
         TabIndex        =   18
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox txtBNETKey 
         Height          =   345
         Left            =   -74760
         TabIndex        =   15
         Top             =   4080
         Width           =   5295
      End
      Begin VB.CommandButton btnRemove 
         Caption         =   "Remove"
         Height          =   255
         Left            =   -71640
         TabIndex        =   14
         Top             =   4560
         Width           =   2175
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "Add"
         Height          =   255
         Left            =   -74760
         TabIndex        =   13
         Top             =   4560
         Width           =   2175
      End
      Begin VB.TextBox txtBNETPassword 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox txtBNETUsername 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   6
         Top             =   1200
         Width           =   2535
      End
      Begin VB.ComboBox cmbBNETServer 
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Top             =   3120
         Width           =   2535
      End
      Begin VB.TextBox txtBNLSServer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   4
         Top             =   2640
         Width           =   2535
      End
      Begin VB.TextBox txtBNETChannel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   3
         Top             =   2160
         Width           =   2535
      End
      Begin MSComctlLib.ListView lvKeyList 
         Height          =   2415
         Left            =   -74760
         TabIndex        =   16
         Top             =   1200
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4260
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ilIcons"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "key"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label Label15 
         Caption         =   "Warcraft III"
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
         TabIndex        =   33
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Diablo II"
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
         Left            =   2280
         TabIndex        =   32
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Warcraft II"
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
         TabIndex        =   31
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Version Bytes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   30
         Top             =   3480
         Width           =   3615
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Battle.Net Settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Internet Relay Chat Settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   24
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Label7 
         Caption         =   "Username"
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
         Left            =   -74760
         TabIndex        =   23
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Channel"
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
         Left            =   -74760
         TabIndex        =   22
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Server"
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
         Left            =   -74760
         TabIndex        =   21
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Battle.Net Key Manager"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   17
         Top             =   600
         Width           =   4935
      End
      Begin VB.Label Label4 
         Caption         =   "Server"
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
         TabIndex        =   12
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Password"
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
         TabIndex        =   11
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Username"
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
         TabIndex        =   10
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "BNLS Server"
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
         TabIndex        =   9
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Channel"
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
         TabIndex        =   8
         Top             =   2160
         Width           =   855
      End
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   5400
      Width           =   1695
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private productChosen As String

Private Sub btnAdd_Click()
  Dim li As ListItem

  If (txtBNETKey.text = vbNullString) Then
    MsgBox "You must enter a CD-Key.", vbOKOnly, PROGRAM_TITLE
    Exit Sub
  End If
  
  If (productChosen = vbNullString) Then
    MsgBox "You must select a product first.", vbOKOnly, PROGRAM_TITLE
    Exit Sub
  End If
  
  If (Not isValidKey(txtBNETKey.text)) Then
    MsgBox "You did not enter a valid " & productChosen & " key.", vbOKOnly, PROGRAM_TITLE
    Exit Sub
  End If
  
  Set li = lvKeyList.ListItems.Add(, , txtBNETKey.text, , productChosen)
  li.Tag = productChosen
  
  txtBNETKey.text = vbNullString
End Sub

Private Sub btnCancel_Click()
  Unload Me
End Sub

Private Sub btnOk_Click()
  Dim oldKeyCount As Integer, li As ListItem, val As Variant
  
  oldKeyCount = config.bnetKeyCount

  config.bnetUsername = txtBNETUsername.text
  config.bnetPassword = txtBNETPassword.text
  config.bnetChannel = txtBNETChannel.text
  config.bnlsServer = txtBNLSServer.text
  config.bnetServer = cmbBNETServer.text
  config.bnetKeyCount = lvKeyList.ListItems.Count
  
  config.bnetW2BNVerByte = "&H" & txtW2BNVerByte.text
  config.bnetD2DVVerByte = "&H" & txtD2DVVerByte.text
  config.bnetWAR3VerByte = "&H" & txtWAR3VerByte.text
  
  setupSockets oldKeyCount, config.bnetKeyCount
  
  If (config.bnetKeyCount > 0) Then
    ReDim bnetData(config.bnetKeyCount - 1)
    
    For i = 0 To config.bnetKeyCount - 1
      With bnetData(i)
        Set li = lvKeyList.ListItems.Item(i + 1)
        .cdKey = li.text
        .product = li.Tag
      End With
    Next i
  End If
  
  config.ircUsername = txtIRCUsername.text
  config.ircChannel = txtIRCChannel.text
  
  val = txtIRCServer.text
  
  If (InStr(val, ":") > 0) Then
    parts = Split(val, ":")
  
    config.ircServer = parts(0)
    config.ircPort = parts(1)
  Else
    config.ircServer = val
    config.ircPort = 6667
  End If
  
  config.rememberWindowPosition = IIf(chkRememberWindowPosition.value = 1, True, False)
  
  saveConfig
  
  Unload Me
End Sub

Private Sub btnRemove_Click()
  lvKeyList.ListItems.Remove (lvKeyList.SelectedItem.index)
End Sub

Private Sub Form_Load()
  Dim arrGateways() As Variant, gateway As String, IPs() As String, key As String, productValue As Long, li As ListItem
  
  txtBNETUsername.text = config.bnetUsername
  txtBNETPassword.text = config.bnetPassword
  txtBNETChannel.text = config.bnetChannel
  txtBNLSServer.text = config.bnlsServer
  cmbBNETServer.text = config.bnetServer
  
  txtW2BNVerByte.text = Right("0" & Hex(config.bnetW2BNVerByte), 2)
  txtD2DVVerByte.text = Right("0" & Hex(config.bnetD2DVVerByte), 2)
  txtWAR3VerByte.text = Right("0" & Hex(config.bnetWAR3VerByte), 2)
  
  If (config.bnetKeyCount > 0) Then
    For i = 0 To config.bnetKeyCount - 1
      With bnetData(i)
        Set li = lvKeyList.ListItems.Add(, , .cdKey, , .product)
        li.Tag = .product
      End With
    Next i
  End If

  txtIRCUsername.text = config.ircUsername
  txtIRCChannel.text = config.ircChannel
  txtIRCServer.text = config.ircServer
  
  chkRememberWindowPosition.value = IIf(config.rememberWindowPosition = True, 1, 0)
  
  arrGateways = Array("uswest.battle.net", "useast.battle.net", "europe.battle.net", "asia.battle.net")

  For i = 0 To 3
    gateway = arrGateways(i)
    cmbBNETServer.AddItem gateway
    IPs = Split(Resolve(gateway))

    For j = 0 To UBound(IPs)
      cmbBNETServer.AddItem IPs(j)
    Next j

    If (i < 3) Then
      cmbBNETServer.AddItem vbNullString
    End If
  Next i
End Sub

Public Function isValidKey(key As String) As Boolean
  Dim productFound As String, productValue As Long
  
  decode_hash_cdkey key, 0, 0, 0, productValue, vbNullString

  Select Case productValue
    Case &H4
      productFound = "W2BN"
    Case &H6, &H7
      productFound = "D2DV"
    Case &HE, &HF
      productFound = "WAR3"
  End Select
  
  If (productFound = productChosen) Then
    isValidKey = True
  Else
    isValidKey = False
  End If
End Function

Private Sub opD2DV_Click()
  productChosen = "D2DV"
End Sub

Private Sub opW2BN_Click()
  productChosen = "W2BN"
End Sub

Private Sub opWAR3_Click()
  productChosen = "WAR3"
End Sub
