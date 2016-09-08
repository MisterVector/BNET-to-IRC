Attribute VB_Name = "modBNLS"
Public Sub Send_BNLS_0x01(index As Integer)
  AddChat frmMain.rtbChatBNET, vbYellow, "Bot #" & index & ": [BNLS] Hashing Key..."
  With pBNLS(index)
    .InsertDWORD BNET(index).ServerToken
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
    BNET(index).ClientToken = .GetDWORD
    BNET(index).CDKeyLength = .GetDWORD
    BNET(index).CDKeyProductValue = .GetDWORD
    BNET(index).CDKeyPublicValue = .GetDWORD
    .Skip 4
    BNET(index).CDKeyHash = .GetNonNTString(20)
  End With
  
  Send_BNLS_0x09 index
End Sub

Public Sub Send_BNLS_0x09(index As Integer)
  Dim tmpLD As String
  tmpLD = Mid(BNET(index).LockdownFile, InStr(BNET(index).LockdownFile, "mpq") - 3, 2)
  If Not IsNumeric(tmpLD) Or Left(tmpLD, 1) = "-" Then tmpLD = Mid(tmpLD, 2)

  AddChat frmMain.rtbChatBNET, vbYellow, "Bot #" & index & ": [BNLS] Requesting version info..."
  With pBNLS(index)
    .InsertDWORD getProdID(BNET(index).prodStr)
    .InsertDWORD tmpLD
    .InsertNTString BNET(index).ValueString
    .sendPacket &H9, True, index
  End With
End Sub

Public Sub Recv_BNLS_0x09(index As Integer)
  If pBNLS(index).GetDWORD = &H0 Then
    AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNLS] Failed to get version info!"
  Else
    AddChat frmMain.rtbChatBNET, vbGreen, "Bot #" & index & ": [BNLS] Version info received!"
    With pBNLS(index)
      BNET(index).EXEVersion = .GetDWORD
      BNET(index).checksum = .GetDWORD
      BNET(index).EXEInfo = .getNTString
    End With
    
    Send_BNLS_0x0B index
  End If
End Sub

Public Sub Send_BNLS_0x0B(index As Integer)
  With pBNLS(index)
    .InsertDWORD Len(Password)
    .InsertDWORD &H2
    .InsertNonNTString Password
    .InsertDWORD BNET(index).ClientToken
    .InsertDWORD BNET(index).ServerToken
    .sendPacket &HB, True, index
  End With
End Sub

Public Sub Recv_BNLS_0x0B(index As Integer)
  With pBNLS(index)
    If newAccFlag Then
      BNET(index).NewAccPasswordHash = .GetNonNTString(20)
    Else
      BNET(index).PasswordHash = .GetNonNTString(20)
    End If
  End With

  frmMain.sckBNLS(index).Close
  If newAccFlag Then
    newAccFlag = False
    
    With pBNET(index)
      .InsertNonNTString BNET(index).NewAccPasswordHash
      .InsertNTString Username
      .sendPacket &H3D, False, index
    End With
  Else
    Send0x51 index
  End If
End Sub


