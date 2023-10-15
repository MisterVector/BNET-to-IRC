Attribute VB_Name = "modBNLS"
Public Sub Send_BNLS_0x1a(index As Integer)
    AddChat frmMain.rtbChatBNET, vbYellow, "Bot #" & index & ": [BNLS] Requesting version info..."
  
    With bnlsPacketHandler(index)
        .InsertDWORD getProdID(bnetData(index).product)
        .InsertDWORD &H0
        .InsertDWORD &H0
        .InsertDWORD bnetData(index).dwLowDateTime
        .InsertDWORD bnetData(index).dwHighDateTime
        .InsertNTString bnetData(index).archiveFileName
        .InsertNTString bnetData(index).valueString
        .sendPacket &H1A
    End With
End Sub

Public Sub Recv_BNLS_0x1a(index As Integer)
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
    
        saveConfig
    
        AddChat frmMain.rtbChatBNET, vbGreen, "Bot #" & index & ": [BNLS] Version byte updated!"
        AddChat frmMain.rtbChatBNET, vbYellow, "Bot #" & index & ": [BNET] Reconnecting..."
    
        frmMain.sckBNLS(index).Close
        frmMain.sckBNET(index).Connect config.bnetServer, 6112
        
        bnetData(index).bnetConnectionState = ConnectionTimeoutState.BNET_CONNECT
        frmMain.tmrBNETConnectionTimeout(index).Enabled = True
    Else
        AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNLS] Unable to update version byte!"
        frmMain.sckBNLS(index).Close
    End If
End Sub

