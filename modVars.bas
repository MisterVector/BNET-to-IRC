Attribute VB_Name = "modVars"
'// BNET SIDE
Public Const joinTillFlood = 5
Public Const timeToWait = 3

Public BNLSServer As String
Public BNETServer As String

Public dicQueue As New Dictionary
Public dicIdx As Integer

Public isFlood As Boolean

Public isBroadcastToIRC As Boolean
Public isBroadcastToBNET As Boolean

Public Username As String
Public Password As String
Public Channel As String
Public newAccFlag As Boolean
Public myChannel As String

Public cIdx As Integer
Public bIdx As Integer

Public pBNET() As clsPacket
Public pBNLS() As clsPacket
Public BotCount As Integer

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
  LockdownFile As String
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
  Username As String
  Password As String
  Server As String
  Port As Long
  Channel As String
End Type
Public IRC As IRCData
