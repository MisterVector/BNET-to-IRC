VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Battle.Net to IRC"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Battle.Net to IRC is a program that allows you to communicate between an IRC server and a Battle.Net channel."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   6135
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Battle.Net to IRC v%v by Vector"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6135
   End
   Begin VB.Label lblLink 
      Caption         =   "http://www.codespeak.org"
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
      Left            =   120
      TabIndex        =   0
      Top             =   2040
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
    lblTitle.Caption = Replace(lblTitle.Caption, "%v", PROGRAM_VERSION)

    Me.Top = Screen.Height / 4
    Me.Left = Screen.Width / 4
End Sub

Private Sub lblLink_Click()
    ShellExecute 0, "open", "http://www.codespeak.org", 0, 0, 0
End Sub

Private Sub lblLink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
End Sub

