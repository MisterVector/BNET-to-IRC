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
    Dim versionByte As Long

    If bnlsPacketHandler(index).GetDWORD = &H0 Then
        AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNLS] Failed to get version info!"
    
        disconnectAll
    Else
        AddChat frmMain.rtbChatBNET, vbGreen, "Bot #" & index & ": [BNLS] Version info received!"
    
        With bnlsPacketHandler(index)
            bnetData(index).exeVersion = .GetDWORD
            bnetData(index).Checksum = .GetDWORD
            bnetData(index).exeInfo = .getNTString
            .Skip 4
            versionByte = .GetDWORD
            
            bnetData(index).oldVerByte = bnetData(index).verByte
            bnetData(index).verByte = versionByte
        End With
    
        Select Case bnetData(index).product
            Case "W2BN": config.bnetW2BNVerByte = versionByte
            Case "D2DV": config.bnetD2DVVerByte = versionByte
        End Select

        saveConfig

        frmMain.sckBNLS(index).Close
    
        Send0x51 index
    End If
End Sub

