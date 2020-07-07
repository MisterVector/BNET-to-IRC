VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Battle.Net to IRC"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   2955
      TabIndex        =   3
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Battle.Net to IRC is a program that bridges the communication between a Battle.Net channel and an IRC channel."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   5
      Top             =   1320
      Width           =   4815
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      Caption         =   "By Vector"
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
      Left            =   3240
      TabIndex        =   4
      Top             =   720
      Width           =   4815
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Battle.Net to IRC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label lblLink 
      Caption         =   "www.codespeak.org"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   2760
      Width           =   2775
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmMain.Icon
    
    Me.Top = Screen.Height / 4
    Me.Left = Screen.Width / 4
End Sub

Private Sub lblLink_Click()
    ShellExecute 0, "open", "https://www.codespeak.org", 0, 0, 0
End Sub

Private Sub lblLink_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
End Sub

