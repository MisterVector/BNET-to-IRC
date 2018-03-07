Attribute VB_Name = "modVars"
Public Const PROGRAM_VERSION As String = "0.0.0"
Public Const PROGRAM_TITLE As String = "BNET to IRC v" & PROGRAM_VERSION & " by Vector"

'// BNET SIDE
Public dicQueue As New Dictionary
Public dicQueueIndex As Integer

Public isBroadcastToIRC As Boolean
Public isBroadcastToBNET As Boolean
  
Public bnetSocketIndex As Integer
Public bnetQueueIndex As Integer

Public bnetPacketBuffer() As clsPacketBuffer
Public bnlsPacketBuffer() As clsPacketBuffer

Public Type ConfigStructure
  bnlsServer As String
  bnetServer As String
  bnetUsername As String
  bnetPassword As String
  bnetChannel As String
  bnetKeyCount As Integer

  ircUsername As String
  ircPassword As String
  ircServer As String
  ircPort As Long
  ircChannel As String
End Type
Public config As ConfigStructure

Public Type bnetDataStructure
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
  
  nls_P As Long
End Type
Public bnetData() As bnetDataStructure
