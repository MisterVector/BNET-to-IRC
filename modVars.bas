Attribute VB_Name = "modVars"
Public Const PROGRAM_VERSION As String = "0.0.0"

Public Const joinTillFlood = 5
Public Const timeToWait = 3

Public BNLSServer As String
Public bnetServer As String

Public dicQueue As New Dictionary
Public dicIdx As Integer

Public isFlood As Boolean

Public isBroadcastToIRC As Boolean
Public isBroadcastToBNET As Boolean

Public username As String
Public password As String
Public channel As String
Public newAccFlag As Boolean
Public myChannel As String

Public cIdx As Integer
Public bIdx As Integer

Public pBNET() As clsPacket
Public pBNLS() As clsPacket
Public botCount As Integer

Public Type BNETData
  AccountName As String
  UniqueName As String

  prodStr As String
  PasswordHash As String
  NewAccPasswordHash As String
  VerByte As Long
  ClientToken As Long
  ValueString As String
  ServerToken As Long
  lockdownFile As String
  ChecksumFormula As String
  EXEInfo As String
  EXEVersion As String
  checksum As String
  CDKey As String
  CDKeyLength As Long
  CDKeyProductValue As Long
  CDKeyPublicValue As Long
  CDKeyHash As String
End Type
Public BNET() As BNETData

'// IRC SIDE
Public Type IRCData
  username As String
  password As String
  Server As String
  Port As Long
  channel As String
End Type
Public IRC As IRCData
