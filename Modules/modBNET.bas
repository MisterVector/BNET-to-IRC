Attribute VB_Name = "modBNET"
Public Sub Recv0x25(index As Integer)
    Send0x25 index
End Sub

Public Sub Send0x25(index As Integer)
    With bnetPacketHandler(index)
        .InsertDWORD .GetDWORD
        .sendPacket &H25
    End With
End Sub

Public Sub Send0x50(index As Integer)
    bnetData(index).verByte = getVerByte(bnetData(index).product)

    With bnetPacketHandler(index)
        .InsertDWORD &H0
        .InsertNonNTString "68XI" & StrReverse(bnetData(index).product)
        .InsertDWORD bnetData(index).verByte
        .InsertDWORD &H0
        .InsertDWORD &H0
        .InsertDWORD &H0
        .InsertDWORD &H0
        .InsertDWORD &H0
        .InsertNTString "USA"
        .InsertNTString "United States"
        .sendPacket &H50
    End With
End Sub

Public Sub Recv0x50(index As Integer)
    Dim FT As FILETIME, archiveTime As String, exeVersion As Long, Checksum As Long, exeInfo As String, result As Long, error As String

    With bnetPacketHandler(index)
        .Skip 4
        bnetData(index).clientToken = GetTickCount
        bnetData(index).serverToken = .GetDWORD
       .Skip 4
        bnetData(index).dwLowDateTime = .GetDWORD
        bnetData(index).dwHighDateTime = .GetDWORD
        bnetData(index).archiveFileName = .getNTString
        bnetData(index).valueString = .getNTString
    End With

    If (config.bnetLocalHashing) Then
        AddChat frmMain.rtbChatBNET, vbYellow, "Bot #" & index & ": [HASH] Checking version..."

        exeInfo = String$(255, Chr$(0))

        With bnetData(index)
            FT.dwLowDateTime = .dwLowDateTime
            FT.dwHighDateTime = .dwHighDateTime
            archiveTime = GetFTTime(FT)

            result = check_revision(archiveTime, .archiveFileName, .valueString, App.Path & "\VersionCheck.ini", .product, exeVersion, Checksum, exeInfo)
        End With

        If (result > 0) Then
            error = String$(crev_max_result, Chr$(0))
            crev_error_description result, error, Len(error)

            AddChat frmMain.rtbChatBNET, vbRed, "[HASH] Version check failed! Error: " & error

            disconnectAll

            Exit Sub
        End If

        AddChat frmMain.rtbChatBNET, vbGreen, "Bot #" & index & ": [HASH] Version check passed!"

        exeInfo = KillNull(exeInfo)

        With bnetData(index)
            .exeVersion = exeVersion
            .Checksum = Checksum
            .exeInfo = exeInfo
        End With

        Send0x51 index
    Else
        AddChat frmMain.rtbChatBNET, vbYellow, "Bot #" & index & ": [BNLS] Connecting to " & config.bnlsServer & "..."

        bnetData(index).bnetConnectionState = ConnectionTimeoutState.BNLS_CONNECT
        frmMain.sckBNLS(index).Connect config.bnlsServer, 9367
        frmMain.tmrBNETConnectionTimeout(index).Enabled = True
    End If
End Sub

Public Sub Send0x51(index As Integer)
    Dim cdKeyLength As Long, cdKeyHash As String * 20, cdKeyProductValue As Long, cdKeyPublicValue As Long
    Dim result As Long

    cdKeyLength = Len(bnetData(index).CDKey)
    
    result = kd_quick(bnetData(index).CDKey, bnetData(index).clientToken, bnetData(index).serverToken, cdKeyPublicValue, cdKeyProductValue, cdKeyHash, Len(cdKeyHash))

    If (result = 0) Then
        AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] Could not decode CD-Key!"
  
        disconnectAll
        Exit Sub
    End If
  
    With bnetPacketHandler(index)
        .InsertDWORD bnetData(index).clientToken
        .InsertDWORD bnetData(index).exeVersion
        .InsertDWORD bnetData(index).Checksum
        .InsertDWORD &H1
        .InsertDWORD &H0
        .InsertDWORD cdKeyLength
        .InsertDWORD cdKeyProductValue
        .InsertDWORD cdKeyPublicValue
        .InsertDWORD &H0
        .InsertNonNTString cdKeyHash
        .InsertNTString bnetData(index).exeInfo
        .InsertNTString "BNET to IRC"
        .sendPacket &H51
    End With
End Sub

Public Sub Recv0x51(index As Integer)
    Dim results As Long, oldVerByte As String, verByte As String
  
    results = bnetPacketHandler(index).GetDWORD
  
    Select Case results
        Case &H0:    AddChat frmMain.rtbChatBNET, vbGreen, "Bot #" & index & ": [BNET] CDKey is accepted."
        Case &H100:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] Your game is out of date."
        Case &H101:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] Invalid game version."
        Case &H102:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] Downgrade game version."
        Case &H200:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] CDKey is invalid."
        Case &H201:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] CDKey is in use by " & bnetPacketHandler(index).getNTString & "."
        Case &H202:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] CDKey is banned."
        Case &H203:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] CDKey is for another product."
        Case &H210:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] Expansion CDKey is invalid."
        Case &H211:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] Expansion CDKey is in use by " & bnetPacketHandler(index).getNTString & "."
        Case &H212:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] Expansion CDKey is banned."
    End Select
  
    If results = &H0 Then
        Send0x3A index
    Else
        disconnectAll
        
        If ((results = &H100 Or results = &H102) And Not config.bnetLocalHashing) Then
            oldVerByte = "0x" & Right("0" & Hex(bnetData(index).oldVerByte), 2)
            verByte = "0x" & Right("0" & Hex(bnetData(index).verByte), 2)
            
            If (oldVerByte <> verByte) Then
                AddChat frmMain.rtbChatBNET, vbYellow, "[INFO] When communicating with the BNLS server, your"
                AddChat frmMain.rtbChatBNET, vbYellow, "[INFO] version byte was updated from " & oldVerByte & " to " & verByte & "."
                AddChat frmMain.rtbChatBNET, vbYellow, "[INFO] Attempting to reconnect..."
                AddChat frmMain.rtbChatBNET, vbYellow, "Bot #0: [BNET] Connecting..."
            
                bnetData(0).bnetConnectionState = ConnectionTimeoutState.BNET_CONNECT
                frmMain.sckBNET(0).Connect config.bnetServer, 6112
                frmMain.tmrBNETConnectionTimeout(0).Enabled = True
            End If
        End If
    End If
End Sub

Public Sub Send0x14(index As Integer)
    bnetPacketHandler(index).InsertNonNTString "tenb"
    bnetPacketHandler(index).sendPacket &H14
End Sub

Public Sub Send0x3A(index As Integer)
    Dim hashCode As String * 20
  
    AddChat frmMain.rtbChatBNET, vbYellow, "Bot #" & index & ": [BNET] Logging in..."

    hashCode = doubleHashPassword(config.bnetPassword, bnetData(index).clientToken, bnetData(index).serverToken)

    With bnetPacketHandler(index)
        .InsertDWORD bnetData(index).clientToken
        .InsertDWORD bnetData(index).serverToken
        .InsertNonNTString hashCode
        .InsertNTString config.bnetUsername
        .sendPacket &H3A
    End With
End Sub

Public Sub Recv0x3A(index As Integer)
    Select Case bnetPacketHandler(index).GetDWORD
        Case &H0:
            AddChat frmMain.rtbChatBNET, vbGreen, "Bot #" & index & ": [BNET] Password accepted!"
            ConnectOtherBots
            Send0x0A index
            Send0x0B index
            Send0x0C index
        Case &H1:
            AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] Account does not exist!"
            AddChat frmMain.rtbChatBNET, vbYellow, "Bot #" & index & ": [BNET] Creating account..."
            Send0x3D index
    Case &H2: AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] Password is invalid!"
    Case &H6: AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] Account is closed: " & bnetPacketHandler(index).getNTString
  End Select
End Sub

Public Sub Send0x3D(index As Integer)
    Dim passwordHash As String * 20
  
    passwordHash = hashPassword(config.bnetPassword)
  
    With bnetPacketHandler(index)
        .InsertNonNTString passwordHash
        .InsertNTString config.bnetUsername
        .sendPacket &H3D
    End With
End Sub

Public Sub Recv0x3D(index As Integer)
    Select Case bnetPacketHandler(index).GetDWORD
        Case &H0: AddChat frmMain.rtbChatBNET, vbGreen, "Bot #" & index & ": [BNET] Account created!"
            Send0x3A index
            Exit Sub
        Case &H2: AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] Account contained invalid characters."
        Case &H3: AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] Account contained a banned word."
        Case &H4: AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] Account already exists!"
        Case &H6: AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] Not enough characters!"
    End Select
  
    disconnectAll
End Sub

Public Sub Send0x0A(index As Integer)
    With bnetPacketHandler(index)
        .InsertNTString vbNullString
        .InsertNTString vbNullString
        .sendPacket &HA
    End With
End Sub

Public Sub Recv0x0A(index As Integer)
    bnetData(index).uniqueName = bnetPacketHandler(index).getNTString
    bnetPacketHandler(index).getNTString 'skip statstring
    bnetData(index).accountName = bnetPacketHandler(index).getNTString
  
    AddChat frmMain.rtbChatBNET, vbGreen, "Bot #" & index & ": [BNET] Logged in as " & bnetData(index).uniqueName
End Sub

Public Sub Send0x0B(index As Integer)
    With bnetPacketHandler(index)
        .InsertDWORD &H0
        .sendPacket &HB
    End With
End Sub

Public Sub Send0x0C(index As Integer)
    With bnetPacketHandler(index)
        .InsertDWORD &H2 'IIf(bnetData(index).product = "D2DV", &H5, &H1)
        .InsertNTString config.bnetChannel
        .sendPacket &HC
    End With
End Sub

Public Sub Recv0x0F(index As Integer)
    Dim user As String, text As String, ID As Long, flags As Long
    Dim supressEvent As Boolean, msgPart As String, currentChannel As String
    Dim message As String

    If index <> findFirstAliveBot Then supressEvent = True
  
    With bnetPacketHandler(index)
        ID = .GetDWORD
        flags = .GetDWORD
        .Skip 16
        user = .getNTString
        text = .getNTString
    
        Select Case ID
            Case &H2:
                If Not supressEvent Then
                    AddChat frmMain.rtbChatBNET, vbWhite, user, vbYellow, " joined " & text
                    SendToIRC user & " has joined " & text & "."
                End If
            Case &H3:
                If Not supressEvent Then
                    AddChat frmMain.rtbChatBNET, vbWhite, user, vbYellow, " left " & text
                    SendToIRC user & " has left " & text & "."
                End If
            Case &H5, &H17:
                If Not supressEvent Then
                    AddChat frmMain.rtbChatBNET, vbWhite, "<" & user & "> ", vbYellow, text
        
                    For i = 0 To UBound(bnetData)
                        If Left(LCase(user), Len(bnetData(i).accountName)) = LCase(bnetData(i).accountName) Then
                            Exit Sub
                        End If
                    Next i
                  
                    ' If broadcast prefix is set, check to see if it is present at the start of the message
                    If (config.bnetBroadcastPrefix <> vbNullString) Then
                        If (Len(text) < Len(config.bnetBroadcastPrefix)) Then
                            Exit Sub
                        End If

                        msgPart = Mid(text, 1, Len(config.bnetBroadcastPrefix))

                        If (msgPart <> config.bnetBroadcastPrefix) Then
                            Exit Sub
                        End If
                    End If

                    currentChannel = bnetData(index).currentChannel
                    message = user & " (" & currentChannel & "): "

                    If (ID = &H17) Then
                        message = message & Chr(&H1D) & user & " " & text & Chr(&H1D)
                    Else
                        message = message & text
                    End If

                    SendToIRC message
                End If
            Case &H7:
                bnetData(index).currentChannel = text
                AddChat frmMain.rtbChatBNET, vbYellow, "You joined the channel ", vbWhite, text
                SendToIRC bnetData(index).uniqueName & " has joined the Battle.Net channel " & text
        End Select
    End With
End Sub
