Attribute VB_Name = "modBNLS"
Public Sub Send_BNLS_0x09(index As Integer)
    Dim lockdownFile As String
    lockdownFile = Mid$(bnetData(index).lockdownFile, InStr(bnetData(index).lockdownFile, "mpq") - 3, 2)
  
    If Not IsNumeric(lockdownFile) Or Left(lockdownFile, 1) = "-" Then lockdownFile = Mid$(lockdownFile, 2)

    AddChat frmMain.rtbChatBNET, vbYellow, "Bot #" & index & ": [BNLS] Requesting version info..."
  
    With bnlsPacketHandler(index)
        .InsertDWORD getProdID(bnetData(index).product)
        .InsertDWORD lockdownFile
        .InsertNTString bnetData(index).valueString
        .sendPacket &H9
    End With
End Sub

Public Sub Recv_BNLS_0x09(index As Integer)
    If bnlsPacketHandler(index).GetDWORD = &H0 Then
        AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNLS] Failed to get version info!"
    
        disconnectAll
    Else
        AddChat frmMain.rtbChatBNET, vbGreen, "Bot #" & index & ": [BNLS] Version info received!"
    
        With bnlsPacketHandler(index)
            bnetData(index).exeVersion = .GetDWORD
            bnetData(index).Checksum = .GetDWORD
            bnetData(index).exeInfo = .getNTString
        End With
    
        frmMain.sckBNLS(index).Close
    
        Send0x51 index
    End If
End Sub

Public Sub Send_BNLS_0x0E(index As Integer)
    AddChat frmMain.rtbChatBNET, vbYellow, "Bot #" & index & ": [BNLS] Authorizing..."
  
    With bnlsPacketHandler(index)
        .InsertNTString "BNET to IRC"
        .sendPacket &HE
    End With
End Sub

Public Sub Recv_BNLS_0x0E(index As Integer)
    With bnlsPacketHandler(index)
        bnetData(index).bnlsServerCode = .GetDWORD
    End With
  
    Send_BNLS_0x0F index
End Sub

Public Sub Send_BNLS_0x0F(index As Integer)
    With bnlsPacketHandler(index)
        .InsertDWORD BNLSChecksum("password", bnetData(index).bnlsServerCode)
        .sendPacket &HF
    End With
End Sub

Public Sub Recv_BNLS_0x0F(index As Integer)
    Dim statusCode As Long
  
    statusCode = bnlsPacketHandler(index).GetDWORD
  
    If (statusCode = &H0) Then
        AddChat frmMain.rtbChatBNET, vbGreen, "Bot #" & index & ": [BNLS] Authorized!"
        Send_BNLS_0x09 index
    Else
        AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNLS] Failed to authorize!"
        disconnectAll
    End If
End Sub

Public Sub Send_BNLS_0x10(index As Integer, product As Long)
    With bnlsPacketHandler(index)
        .InsertDWORD product
        .sendPacket &H10
    End With
End Sub

Public Sub Recv_BNLS_0x10(index As Integer)
    Dim product As Long, versionByte As Long
  
    product = bnlsPacketHandler(index).GetDWORD
  
    If (product > &H0) Then
        versionByte = bnlsPacketHandler(index).GetDWORD
    
        Select Case product
            Case &H3: config.bnetW2BNVerByte = versionByte
            Case &H4: config.bnetD2DVVerByte = versionByte
        End Select
    
        AddChat frmMain.rtbChatBNET, vbGreen, "Bot #" & index & ": [BNLS] Version byte updated!"
        AddChat frmMain.rtbChatBNET, vbYellow, "Bot #" & index & ": [BNET] Reconnecting..."
    
        frmMain.sckBNLS(index).Close
        frmMain.sckBNET(index).Connect config.bnetServer, 6112
    Else
        AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNLS] Unable to update version byte!"
        frmMain.sckBNLS(index).Close
    End If
End Sub

