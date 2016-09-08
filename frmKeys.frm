VERSION 5.00
Begin VB.Form frmKeys 
   Caption         =   "frmConnections"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbProduct 
      Height          =   315
      ItemData        =   "frmKeys.frx":0000
      Left            =   120
      List            =   "frmKeys.frx":0010
      TabIndex        =   7
      Text            =   "Choose product for this key"
      Top             =   3120
      Width           =   2895
   End
   Begin VB.CommandButton btnRemove 
      Caption         =   "Remove Key"
      Height          =   300
      Left            =   4080
      TabIndex        =   6
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   "Add Key"
      Height          =   300
      Left            =   3120
      TabIndex        =   5
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox txtKey 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   2895
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   300
      Left            =   4080
      TabIndex        =   3
      Top             =   3135
      Width           =   1095
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   300
      Left            =   3120
      TabIndex        =   2
      Top             =   3135
      Width           =   855
   End
   Begin VB.ListBox lstKeys 
      Height          =   2595
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "The number of keys you put in this list determines how many connectios you have to Battle.Net. Do not put more than 8."
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
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   5055
   End
End
Attribute VB_Name = "frmKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAdd_Click()
  If cmbProduct.Text = "Choose product for this key" Then
    MsgBox "Choose a product for this key."
    Exit Sub
  End If

  lstKeys.AddItem txtKey.Text & " -> " & getProduct(cmbProduct.ListIndex)
  cmbProduct.Text = "Choose product for this key"
End Sub

Private Sub btnOk_Click()
  Dim getKeyLine As String, getKey As String, getProduct As String

  If Dir$(App.Path & "\Config.ini") = vbNullString Then
    Kill App.Path & "\Config.ini"
  End If
  
  BotCount = lstKeys.ListCount
  ReDim BNET(BotCount - 1)
  
  For i = 0 To BotCount - 1
    getKeyLine = lstKeys.List(i)
    
    getKey = Split(getKeyLine, " -> ")(0)
    getProduct = Split(getKeyLine, " -> ")(1)
  
    With BNET(i)
      .CDKey = getKey
      .prodStr = getProduct
    End With
    
    WriteINI i, "Product", getProduct, "Config.ini"
    WriteINI i, "CDKey", getKey, "Config.ini"
  Next i
  
  Unload Me
End Sub

Private Sub btnRemove_Click()
  If lstKeys.List(lstKeys.ListIndex) = vbNullString Then Exit Sub
  
  lstKeys.RemoveItem (lstKeys.ListIndex)
End Sub

Public Function getProduct(ByVal prodIdx As Integer) As String
  Select Case prodIdx
    Case 0: getProduct = "STAR"
    Case 1: getProduct = "W2BN"
    Case 2: getProduct = "D2DV"
    Case 3: getProduct = "WAR3"
  End Select
End Function

Private Sub Form_Load()
  If BotCount > 0 Then
    For i = 0 To UBound(BNET)
      lstKeys.AddItem BNET(i).CDKey & " -> " & BNET(i).prodStr
    Next i
  End If
End Sub
