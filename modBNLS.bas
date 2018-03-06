Attribute VB_Name = "modBNLS"
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
    
    frmMain.Click_start
  Else
    AddChat frmMain.rtbChatBNET, vbGreen, "Bot #" & index & ": [BNLS] Version info received!"
    
    With bnlsPacketBuffer(index)
      bnetData(index).exeVersion = .GetDWORD
      bnetData(index).Checksum = .GetDWORD
      bnetData(index).exeInfo = .getNTString
    End With
    
    Send0x51 index
  End If
End Sub

