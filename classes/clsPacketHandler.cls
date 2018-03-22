VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPacketHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private sIn As String, sOut As String, sLen As Long
Private sck As Winsock, pt As packetType

Public Sub setSocket(sock As Winsock, packetType As packetType)
  Set sck = sock
  pt = packetType
End Sub

Public Sub SetData(ByVal sData As String)
  sIn = sData
  sLen = 1
End Sub

Public Sub clearOut()
  sOut = vbNullString
End Sub

Public Sub clearIn()
  sIn = vbNullString
End Sub

Public Sub Skip(Bytes As Long)
  sLen = sLen + Bytes
End Sub

Public Sub InsertByte(ByVal value As Byte)
  sOut = sOut & Chr$(value)
End Sub

Public Sub InsertWORD(ByVal value As Integer)
  Dim NewValue As String * 2
  CopyMemory ByVal NewValue, value, 2
  sOut = sOut & NewValue
End Sub

Public Sub InsertDWORD(ByVal value As Long)
  Dim NewValue As String * 4
  CopyMemory ByVal NewValue, value, 4
  sOut = sOut & NewValue
End Sub

Public Sub InsertULONG(ByVal data As String)
  Dim finalLong As String * 8
  CopyMemory ByVal finalLong, data, 8
  sOut = sOut & finalLong
End Sub

Public Sub InsertNTString(ByVal value As String)
  sOut = sOut & value & vbNullChar
End Sub

Public Sub InsertNonNTString(ByVal value As String)
  sOut = sOut & value
End Sub

Public Function getPacket() As String
  getPacket = sIn
End Function

Public Sub sendPacket(ByVal ID As Byte)
  Dim finalString As String, NewValue As String * 2
  
  CopyMemory ByVal NewValue, (Len(sOut) + IIf(pt = packetType.BNLS, 3, 4)), 2
  
  If (sck.State = sckConnected) Then
    If (pt = packetType.BNLS) Then
      CopyMemory ByVal NewValue, (Len(sOut) + 3), 2
      finalString = NewValue & Chr$(ID) & sOut
    Else
      CopyMemory ByVal NewValue, (Len(sOut) + 4), 2
      finalString = Chr$(&HFF) & Chr$(ID) & NewValue & sOut
    End If
  
    sck.SendData finalString
  End If
  
  sOut = vbNullString
End Sub

Public Function GetByte() As Byte
  GetByte = Asc(Mid$(sIn, sLen, 1))
  sLen = sLen + 1
End Function

Public Function GetWORD() As Long
  Dim NewValue As Long
  CopyMemory NewValue, ByVal Mid$(sIn, sLen, 2), 2
  GetWORD = NewValue
  sLen = sLen + 2
End Function

Public Function GetDWORD() As Long
  Dim NewValue As Long
  CopyMemory NewValue, ByVal Mid$(sIn, sLen, 4), 4
  GetDWORD = NewValue
  sLen = sLen + 4
End Function

Public Function GetNonNTString(ByVal Length As Integer) As String
  GetNonNTString = Mid$(sIn, sLen, Length)
  sLen = sLen + Length
End Function

Public Function getNTString() As String
  On Error Resume Next
  
  getNTString = Split(Mid$(sIn, sLen), vbNullChar)(0)
  sLen = sLen + Len(getNTString) + 1
End Function

Public Function getULONG() As String
  getULONG = Mid$(sIn, sLen, 8)
  sLen = sLen + 8
End Function
