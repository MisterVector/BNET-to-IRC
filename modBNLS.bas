Attribute VB_Name = "modBNLS"
Public Sub Send_BNLS_0x01(index As Integer)
  AddChat frmMain.rtbChatBNET, vbYellow, "Bot #" & index & ": [BNLS] Hashing Key..."
  
  With bnlsPacketBuffer(index)
    .InsertDWORD bnetData(index).serverToken
    .InsertNTString bnetData(index).cdKey
    .sendPacket &H1, True, index
  End With
End Sub

Public Sub Recv_BNLS_0x01(index As Integer)
  If bnlsPacketBuffer(index).GetDWORD = &H0 Then
    AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNLS] Key hash failed!"
    
    frmMain.sckBNET(index).Close
    frmMain.sckBNLS(index).Close
    
    Exit Sub
  Else
    AddChat frmMain.rtbChatBNET, vbGreen, "Bot #" & index & ": [BNLS] Key hashed successfully!"
  End If
  
  With bnlsPacketBuffer(index)
    bnetData(index).clientToken = .GetDWORD
    bnetData(index).cdKeyLength = .GetDWORD
    bnetData(index).cdKeyProductValue = .GetDWORD
    bnetData(index).cdKeyPublicValue = .GetDWORD
    .Skip 4
    bnetData(index).cdKeyHash = .GetNonNTString(20)
  End With
  
  Send_BNLS_0x09 index
End Sub

Public Sub Send_BNLS_0x09(index As Integer)
  Dim lockdownFile As String
  lockdownFile = Mid(bnetData(index).lockdownFile, InStr(bnetData(index).lockdownFile, "mpq") - 3, 2)
  If Not IsNumeric(lockdownFile) Or Left(lockdownFile, 1) = "-" Then lockdownFile = Mid(lockdownFile, 2)

  AddChat frmMain.rtbChatBNET, vbYellow, "Bot #" & index & ": [BNLS] Requesting version info..."
  
  With bnlsPacketBuffer(index)
    .InsertDWORD getProdID(bnetData(index).product)
    .InsertDWORD lockdownFile
    .InsertNTString bnetData(index).valueString
    .sendPacket &H9, True, index
  End With
End Sub

Public Sub Recv_BNLS_0x09(index As Integer)
  If bnlsPacketBuffer(index).GetDWORD = &H0 Then
    AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNLS] Failed to get version info!"
  Else
    AddChat frmMain.rtbChatBNET, vbGreen, "Bot #" & index & ": [BNLS] Version info received!"
    With bnlsPacketBuffer(index)
      bnetData(index).exeVersion = .GetDWORD
      bnetData(index).checksum = .GetDWORD
      bnetData(index).exeInfo = .getNTString
    End With
    
    Send_BNLS_0x0B index
  End If
End Sub

Public Sub Send_BNLS_0x0B(index As Integer)
  With bnlsPacketBuffer(index)
    .InsertDWORD Len(config.bnetPassword)
    .InsertDWORD &H2
    .InsertNonNTString config.bnetPassword
    .InsertDWORD bnetData(index).clientToken
    .InsertDWORD bnetData(index).serverToken
    .sendPacket &HB, True, index
  End With
End Sub

Public Sub Recv_BNLS_0x0B(index As Integer)
  With bnlsPacketBuffer(index)
    If newAccFlag Then
      bnetData(index).newAccPasswordHash = .GetNonNTString(20)
    Else
      bnetData(index).passwordHash = .GetNonNTString(20)
    End If
  End With

  frmMain.sckBNLS(index).Close
  
  If newAccFlag Then
    newAccFlag = False
    
    With bnetPacketBuffer(index)
      .InsertNonNTString bnetData(index).newAccPasswordHash
      .InsertNTString config.bnetUsername
      .sendPacket &H3D, False, index
    End With
  Else
    Send0x51 index
  End If
End Sub


