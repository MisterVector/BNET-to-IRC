Attribute VB_Name = "modVars"
Public Const PROGRAM_VERSION As String = "0.0.0"
Public Const PROGRAM_TITLE As String = "Battle.Net to IRC v" & PROGRAM_VERSION & " by Vector"

Public Const VERBYTE_W2BN As Long = &H4F
Public Const VERBYTE_D2DV As Long = &HE
Public Const VERBYTE_WAR3 As Long = &H1C

Public Const DEFAULT_BNLS_SERVER As String = "bnls.bnetdocs.org"
Public Const DEFAULT_REMEMBER_WINDOW_POSITION As Boolean = False

'// BNET SIDE
Public dicQueue As New Dictionary
Public dicQueueIndex As Integer

Public isBroadcastToIRC As Boolean
Public isBroadcastToBNET As Boolean
  
Public bnetSocketIndex As Integer
Public bnetQueueIndex As Integer

Public bnetPacketHandler() As clsPacketHandler
Public bnlsPacketHandler() As clsPacketHandler

Public Enum BNLSRequestType
  REQUEST_FILE_INFO
End Enum
Public bnlsType As BNLSRequestType

Public Enum PacketType
  BNCS
  BNLS
End Enum

Public Type ConfigStructure
  formTop As Integer
  formLeft As Integer
  rememberWindowPosition As Boolean

  bnlsServer As String
  bnetServer As String
  bnetUsername As String
  bnetPassword As String
  bnetChannel As String
  bnetKeyCount As Integer

  bnetW2BNVerByte As Long
  bnetD2DVVerByte As Long
  bnetWAR3VerByte As Long

  ircUsername As String
  ircPassword As String
  ircServer As String
  ircPort As Long
  ircChannel As String
End Type
Public config As ConfigStructure

Public Type BNETDataStructure
  accountName As String
  uniqueName As String

  product As String
  passwordHash As String
  verByte As Long
  clientToken As Long
  serverToken As Long
  valueString As String
  lockdownFile As String
  checksumFormula As String
  exeInfo As String
  exeVersion As String
  Checksum As String
  cdKey As String
  
  bnlsServerCode As Long
  
  nls_P As Long
End Type
Public bnetData() As BNETDataStructure
