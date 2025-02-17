VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings for Battle.Net to IRC"
   ClientHeight    =   6645
   ClientLeft      =   1875
   ClientTop       =   1995
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ilIcons 
      Left            =   2280
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":0000
            Key             =   "W2BN"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":04EA
            Key             =   "D2DV"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   5535
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Battle.Net"
      TabPicture(0)   =   "frmSettings.frx":09D4
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
      Tab(0).Control(9)=   "Label18"
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
      Tab(0).Control(17)=   "txtBNETBroadcastPrefix"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "chkBNETLocalHashing"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "Key Manager"
      TabPicture(1)   =   "frmSettings.frx":09F0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "opD2DV"
      Tab(1).Control(1)=   "opW2BN"
      Tab(1).Control(2)=   "txtBNETKey"
      Tab(1).Control(3)=   "btnRemove"
      Tab(1).Control(4)=   "btnAdd"
      Tab(1).Control(5)=   "lvKeyList"
      Tab(1).Control(6)=   "Label2"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "IRC"
      TabPicture(2)   =   "frmSettings.frx":0A0C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtIRCBroadcastPrefix"
      Tab(2).Control(1)=   "chkUpdateChannelOnChannelJoin"
      Tab(2).Control(2)=   "txtIRCQuitMessage"
      Tab(2).Control(3)=   "txtIRCServer"
      Tab(2).Control(4)=   "txtIRCChannel"
      Tab(2).Control(5)=   "txtIRCUsername"
      Tab(2).Control(6)=   "Label19"
      Tab(2).Control(7)=   "Label15"
      Tab(2).Control(8)=   "Label10"
      Tab(2).Control(9)=   "Label7"
      Tab(2).Control(10)=   "Label8"
      Tab(2).Control(11)=   "Label9"
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "Miscellaneous"
      TabPicture(3)   =   "frmSettings.frx":0A28
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "chkMinimizeToTray"
      Tab(3).Control(1)=   "txtConnectionTimeout"
      Tab(3).Control(2)=   "chkCheckUpdateOnStartup"
      Tab(3).Control(3)=   "chkRememberWindowPosition"
      Tab(3).Control(4)=   "Label17"
      Tab(3).Control(5)=   "Label16"
      Tab(3).ControlCount=   6
      Begin VB.CheckBox chkMinimizeToTray 
         Caption         =   "Minimize to Tray"
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
         TabIndex        =   46
         Top             =   2400
         Width           =   3015
      End
      Begin VB.CheckBox chkBNETLocalHashing 
         Caption         =   "Use Local Hashing"
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
         TabIndex        =   45
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox txtIRCBroadcastPrefix 
         Height          =   375
         Left            =   -72720
         TabIndex        =   43
         Top             =   2640
         Width           =   2775
      End
      Begin VB.TextBox txtBNETBroadcastPrefix 
         Height          =   375
         Left            =   2280
         TabIndex        =   42
         Top             =   3960
         Width           =   2775
      End
      Begin VB.TextBox txtConnectionTimeout 
         Height          =   360
         Left            =   -72720
         TabIndex        =   38
         Top             =   1200
         Width           =   2775
      End
      Begin VB.CheckBox chkUpdateChannelOnChannelJoin 
         Caption         =   "Update On Channel Join"
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
         Left            =   -72720
         TabIndex        =   17
         Top             =   3120
         Width           =   2775
      End
      Begin VB.TextBox txtIRCQuitMessage 
         Height          =   1575
         Left            =   -74760
         TabIndex        =   18
         Top             =   3720
         Width           =   4815
      End
      Begin VB.CheckBox chkCheckUpdateOnStartup 
         Caption         =   "Check for Update on Startup"
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
         TabIndex        =   20
         Top             =   2040
         Width           =   3015
      End
      Begin VB.TextBox txtD2DVVerByte 
         Height          =   360
         Left            =   3360
         TabIndex        =   7
         Top             =   4920
         Width           =   735
      End
      Begin VB.TextBox txtW2BNVerByte 
         Height          =   375
         Left            =   1440
         TabIndex        =   6
         Top             =   4920
         Width           =   735
      End
      Begin VB.CheckBox chkRememberWindowPosition 
         Caption         =   "Remember Window Position"
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
         TabIndex        =   19
         Top             =   1680
         Width           =   2895
      End
      Begin VB.OptionButton opD2DV 
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
         Left            =   -73440
         TabIndex        =   11
         Top             =   4320
         Width           =   1095
      End
      Begin VB.OptionButton opW2BN 
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
         Left            =   -74760
         TabIndex        =   10
         Top             =   4320
         Width           =   1215
      End
      Begin VB.TextBox txtIRCServer 
         Height          =   345
         Left            =   -72720
         TabIndex        =   14
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txtIRCChannel 
         Height          =   345
         Left            =   -72720
         TabIndex        =   16
         Top             =   2160
         Width           =   2775
      End
      Begin VB.TextBox txtIRCUsername 
         Height          =   345
         Left            =   -72720
         TabIndex        =   15
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox txtBNETKey 
         Height          =   345
         Left            =   -74760
         TabIndex        =   9
         Top             =   3840
         Width           =   4815
      End
      Begin VB.CommandButton btnRemove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   -72360
         TabIndex        =   13
         Top             =   4800
         Width           =   2415
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   -74760
         TabIndex        =   12
         Top             =   4800
         Width           =   2415
      End
      Begin VB.TextBox txtBNETPassword 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   3000
         Width           =   2775
      End
      Begin VB.TextBox txtBNETUsername 
         Height          =   360
         Left            =   2280
         TabIndex        =   3
         Top             =   2520
         Width           =   2775
      End
      Begin VB.ComboBox cmbBNETServer 
         Height          =   315
         Left            =   2280
         TabIndex        =   1
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txtBNLSServer 
         Height          =   360
         Left            =   2280
         TabIndex        =   2
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox txtBNETChannel 
         Height          =   360
         Left            =   2280
         TabIndex        =   5
         Top             =   3480
         Width           =   2775
      End
      Begin MSComctlLib.ListView lvKeyList 
         Height          =   2535
         Left            =   -74760
         TabIndex        =   8
         Top             =   1200
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   4471
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
      Begin VB.Label Label19 
         Caption         =   "Broadcast Prefix"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   44
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label Label18 
         Caption         =   "Broadcast Prefix"
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
         TabIndex        =   41
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "Miscellaneous Settings"
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
         TabIndex        =   40
         Top             =   600
         Width           =   4815
      End
      Begin VB.Label Label16 
         Caption         =   "Connection Timeout"
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
         TabIndex        =   39
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label15 
         Caption         =   "Quit Message"
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
         TabIndex        =   37
         Top             =   3360
         Width           =   1455
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
         Left            =   2400
         TabIndex        =   36
         Top             =   4920
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
         Left            =   360
         TabIndex        =   35
         Top             =   4920
         Width           =   975
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
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
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   4440
         Width           =   4815
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
         TabIndex        =   33
         Top             =   600
         Width           =   4815
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
         TabIndex        =   32
         Top             =   600
         Width           =   4815
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
         TabIndex        =   31
         Top             =   1680
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
         TabIndex        =   30
         Top             =   2160
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
         TabIndex        =   29
         Top             =   1200
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
         TabIndex        =   28
         Top             =   600
         Width           =   4815
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
         TabIndex        =   0
         Top             =   1200
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
         TabIndex        =   25
         Top             =   3000
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
         TabIndex        =   24
         Top             =   2520
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
         TabIndex        =   23
         Top             =   1680
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
         TabIndex        =   26
         Top             =   3480
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
      TabIndex        =   22
      Top             =   5880
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
      Left            =   240
      TabIndex        =   21
      Top             =   5880
      Width           =   1695
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private productChosen As String

Private Sub btnAdd_Click()
    Dim li As ListItem

    If (txtBNETKey.text = vbNullString) Then
        MsgBox "You must enter a CD-Key.", vbOKOnly Or vbInformation, PROGRAM_TITLE
        Exit Sub
    End If
  
    If (productChosen = vbNullString) Then
        MsgBox "You must select a product first.", vbOKOnly Or vbInformation, PROGRAM_TITLE
        Exit Sub
    End If

    If (Not isValidKey(txtBNETKey.text)) Then
        MsgBox "You did not enter a valid " & productChosen & " key.", vbOKOnly Or vbInformation, PROGRAM_TITLE
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
  
    If (IsNumeric(txtConnectionTimeout.text)) Then
        If (txtConnectionTimeout.text < 1 Or txtConnectionTimeout.text > 30000) Then
            MsgBox "Connection timeout must be a value between 1 and 30000.", vbOKOnly, PROGRAM_NAME
            
            Exit Sub
        End If
    Else
        MsgBox "Connection timeout must be a value between 1 and 30000.", vbOKOnly, PROGRAM_NAME
        
        Exit Sub
    End If
  
    If (Not IsNumeric("&H" & txtW2BNVerByte.text)) Then
        MsgBox "The Warcraft II version byte must be numeric. (I.E. " & hexToString(VERBYTE_W2BN) & ")"
        Exit Sub
    End If
    
    If (Not IsNumeric("&H" & txtD2DVVerByte.text)) Then
        MsgBox "The Diablo II version byte must be numeric. (I.E. " & hexToString(VERBYTE_D2DV) & ")"
        Exit Sub
    End If
  
    oldKeyCount = config.bnetKeyCount

    config.bnetUsername = txtBNETUsername.text
    config.bnetPassword = txtBNETPassword.text
    config.bnetChannel = txtBNETChannel.text
    config.bnlsServer = txtBNLSServer.text
    config.bnetServer = cmbBNETServer.text
    config.bnetKeyCount = lvKeyList.ListItems.count
    config.bnetBroadcastPrefix = txtBNETBroadcastPrefix.text
  
    config.bnetLocalHashing = chkBNETLocalHashing.Value = 1
    config.bnetW2BNVerByte = "&H" & txtW2BNVerByte.text
    config.bnetD2DVVerByte = "&H" & txtD2DVVerByte.text
  
    setupSockets oldKeyCount, config.bnetKeyCount
  
    If (config.bnetKeyCount > 0) Then
        ReDim bnetData(config.bnetKeyCount - 1)
    
        For i = 0 To config.bnetKeyCount - 1
            With bnetData(i)
                Set li = lvKeyList.ListItems.Item(i + 1)
                .CDKey = li.text
                .product = li.Tag
            End With
        Next i
    End If
  
    config.ircServer = txtIRCServer.text
    config.ircUsername = txtIRCUsername.text
    config.ircChannel = txtIRCChannel.text
    config.ircQuitMessage = txtIRCQuitMessage.text
    config.ircUpdateChannelOnChannelJoin = IIf(chkUpdateChannelOnChannelJoin.Value = 1, True, False)
    config.ircBroadcastPrefix = txtIRCBroadcastPrefix.text
 
    config.rememberWindowPosition = IIf(chkRememberWindowPosition.Value = 1, True, False)
    config.checkUpdateOnStartup = IIf(chkCheckUpdateOnStartup.Value = 1, True, False)
    config.connectionTimeout = txtConnectionTimeout.text
    config.minimizeToTray = IIf(chkMinimizeToTray.Value = 1, True, False)

    frmMain.tmrIRCConnectionTimeout.Interval = config.connectionTimeout

    saveConfig
  
    Unload Me
End Sub

Private Sub btnRemove_Click()
    lvKeyList.ListItems.Remove (lvKeyList.SelectedItem.index)
End Sub

Private Sub Form_Load()
    Dim arrGateways() As Variant, gateway As String, IPs() As String, key As String, productValue As Long, li As ListItem
  
    Me.Icon = frmMain.Icon

    txtBNETUsername.text = config.bnetUsername
    txtBNETPassword.text = config.bnetPassword
    txtBNETChannel.text = config.bnetChannel
    txtBNLSServer.text = config.bnlsServer
    txtBNETBroadcastPrefix.text = config.bnetBroadcastPrefix
    cmbBNETServer.text = config.bnetServer
  
    chkBNETLocalHashing.Value = IIf(config.bnetLocalHashing, 1, 0)
    txtW2BNVerByte.text = Right("0" & Hex(config.bnetW2BNVerByte), 2)
    txtD2DVVerByte.text = Right("0" & Hex(config.bnetD2DVVerByte), 2)
  
    If (config.bnetKeyCount > 0) Then
        For i = 0 To config.bnetKeyCount - 1
            With bnetData(i)
                Set li = lvKeyList.ListItems.Add(, , .CDKey, , .product)
                li.Tag = .product
            End With
        Next i
    End If

    txtIRCUsername.text = config.ircUsername
    txtIRCChannel.text = config.ircChannel
    txtIRCServer.text = config.ircServer
    txtIRCQuitMessage.text = config.ircQuitMessage
    txtIRCBroadcastPrefix.text = config.ircBroadcastPrefix
    chkUpdateChannelOnChannelJoin.Value = IIf(config.ircUpdateChannelOnChannelJoin, 1, 0)
  
    chkRememberWindowPosition.Value = IIf(config.rememberWindowPosition = True, 1, 0)
    chkCheckUpdateOnStartup.Value = IIf(config.checkUpdateOnStartup = True, 1, 0)
    txtConnectionTimeout.text = config.connectionTimeout
    chkMinimizeToTray.Value = IIf(config.minimizeToTray = True, 1, 0)
  
    arrGateways = Array("uswest.battle.net", "useast.battle.net", "europe.battle.net", "asia.battle.net", _
                        "connect-eur.classic.blizzard.com", "connect-kor.classic.blizzard.com", _
                        "connect-use.classic.blizzard.com", "connect-usw.classic.blizzard.com")

    If (cmbBNETServer.text <> vbNullString) Then
        cmbBNETServer.AddItem cmbBNETServer.text
        cmbBNETServer.AddItem vbNullString
    End If

    For i = 0 To UBound(arrGateways)
        gateway = arrGateways(i)
        cmbBNETServer.AddItem gateway
        IPs = Split(Resolve(gateway))

        For j = 0 To UBound(IPs)
            cmbBNETServer.AddItem IPs(j)
        Next j

        If (i < UBound(arrGateways)) Then
            cmbBNETServer.AddItem vbNullString
        End If
    Next i
  
    Me.Top = Screen.Height / 4
    Me.Left = Screen.Width / 4
End Sub

Public Function isValidKey(key As String) As Boolean
    Dim productFound As String, productValue As Long
  
    kd_quick key, 0, 0, 0, productValue, vbNullString, 0

    Select Case productValue
        Case &H4
            productFound = "W2BN"
        Case &H6, &H7, &H18
            productFound = "D2DV"
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

