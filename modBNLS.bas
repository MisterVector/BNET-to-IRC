Attribute VB_Name = "modBNLS"
Public Sub Send_BNLS_0x01(index As Integer)
  AddChat frmMain.rtbChatBNET, vbYellow, "Bot #" & index & ": [BNLS] Hashing Key..."
  
  With pBNLS(index)
    .InsertDWORD BNET(index).serverToken
    .InsertNTString BNET(index).CDKey
    .sendPacket &H1, True, index
  End With
End Sub

Public Sub Recv_BNLS_0x01(index As Integer)
  If pBNLS(index).GetDWORD = &H0 Then
    AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNLS] Key hash failed!"
    
    frmMain.sckBNET(index).Close
    frmMain.sckBNLS(index).Close
    
    Exit Sub
  Else
    AddChat frmMain.rtbChatBNET, vbGreen, "Bot #" & index & ": [BNLS] Key hashed successfully!"
  End If
  
  With pBNLS(index)
    BNET(index).clientToken = .GetDWORD
    BNET(index).CDKeyLength = .GetDWORD
    BNET(index).CDKeyProductValue = .GetDWORD
    BNET(index).CDKeyPublicValue = .GetDWORD
    .Skip 4
    BNET(index).CDKeyHash = .GetNonNTString(20)
  End With
  
  Send_BNLS_0x09 index
End Sub

Public Sub Send_BNLS_0x09(index As Integer)
  Dim lockdownFile As String
  lockdownFile = Mid(BNET(index).lockdownFile, InStr(BNET(index).lockdownFile, "mpq") - 3, 2)
  If Not IsNumeric(lockdownFile) Or Left(lockdownFile, 1) = "-" Then lockdownFile = Mid(lockdownFile, 2)

  AddChat frmMain.rtbChatBNET, vbYellow, "Bot #" & index & ": [BNLS] Requesting version info..."
  
  With pBNLS(index)
    .InsertDWORD getProdID(BNET(index).prodStr)
    .InsertDWORD lockdownFile
    .InsertNTString BNET(index).valueString
    .sendPacket &H9, True, index
  End With
End Sub

Public Sub Recv_BNLS_0x09(index As Integer)
  If pBNLS(index).GetDWORD = &H0 Then
    AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNLS] Failed to get version info!"
  Else
    AddChat frmMain.rtbChatBNET, vbGreen, "Bot #" & index & ": [BNLS] Version info received!"
    With pBNLS(index)
      BNET(index).exeVersion = .GetDWORD
      BNET(index).checksum = .GetDWORD
      BNET(index).exeInfo = .getNTString
    End With
    
    Send_BNLS_0x0B index
  End If
End Sub

Public Sub Send_BNLS_0x0B(index As Integer)
  With pBNLS(index)
    .InsertDWORD Len(password)
    .InsertDWORD &H2
    .InsertNonNTString password
    .InsertDWORD BNET(index).clientToken
    .InsertDWORD BNET(index).serverToken
    .sendPacket &HB, True, index
  End With
End Sub

Public Sub Recv_BNLS_0x0B(index As Integer)
  With pBNLS(index)
    If newAccFlag Then
      BNET(index).newAccPasswordHash = .GetNonNTString(20)
    Else
      BNET(index).passwordHash = .GetNonNTString(20)
    End If
  End With

  frmMain.sckBNLS(index).Close
  
  If newAccFlag Then
    newAccFlag = False
    
    With pBNET(index)
      .InsertNonNTString BNET(index).newAccPasswordHash
      .InsertNTString username
      .sendPacket &H3D, False, index
    End With
  Else
    Send0x51 index
  End If
End Sub


