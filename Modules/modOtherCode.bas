Attribute VB_Name = "modOtherCode"
Private Sub AddQ(ByVal msg As String)
    dicQueue.Add dicQueue.count + 1, msg

    If Not frmMain.tmrReleaseQueue.Enabled Then
        dicQueueIndex = 1
        frmMain.tmrReleaseQueue.Enabled = True
    End If
End Sub

Public Function getVerByte(ByVal product As String) As Long
    Select Case product
        Case "W2BN": getVerByte = config.bnetW2BNVerByte
        Case "D2DV": getVerByte = config.bnetD2DVVerByte
    End Select
End Function

Public Function getProdID(ByVal product As String) As Long
    Select Case product
        Case "W2BN": getProdID = &H3
        Case "D2DV": getProdID = &H4
    End Select
End Function

Public Function findFirstAliveBot() As Integer
    For i = 0 To frmMain.sckBNET.count - 1
        If frmMain.sckBNET(i).State = sckConnected Then
            findFirstAliveBot = i
            Exit Function
        End If
    Next i

    findFirstAliveBot = -1
End Function

Public Sub SendToBNET(ByVal msg As String, Optional ByVal force As Boolean = False)
    If (Not isBroadcastToBNET And Not force) Then Exit Sub
    If (findFirstAliveBot() = -1) Then Exit Sub

    If Len(msg) > 140 Then
        Do While Len(msg) > 140
            AddQ Mid$(msg, 1, 140) & " [more]"
            msg = Mid$(msg, 141)
        Loop

        If Len(msg) > 0 Then
            AddQ msg
        End If
    Else
        AddQ msg
    End If
End Sub

Public Sub SendToIRC(ByVal msg As String)
    Dim currentJoinedChannel As String
    
    If (isBroadcastToIRC) Then
        If frmMain.sckIRC.State = sckConnected Then
            currentJoinedChannel = IRCData.joinedChannel
        
            If (currentJoinedChannel <> vbNullString) Then
                SendPRIVMSG currentJoinedChannel, msg
            End If
        End If
    End If
End Sub

Public Sub ConnectOtherBots()
    If config.bnetKeyCount > 1 Then
        For i = 1 To frmMain.sckBNET.count - 1
            If frmMain.sckBNET(i).State = sckClosed Then
                frmMain.sckBNET(i).Connect config.bnetServer, 6112
                
                bnetData(i).bnetConnectionState = ConnectionTimeoutState.BNET_CONNECT
                frmMain.tmrBNETConnectionTimeout(i).Enabled = True
            End If
        Next i
    End If
End Sub

Public Sub quitProgram()
    Dim oFrm As Form

    For Each oFrm In Forms
        Unload oFrm
    Next
End Sub

Public Function accountIdToReason(ByVal ID As Long) As String
    Dim reason As String

    Select Case ID
        Case &H1: reason = "username is too short"
        Case &H2: reason = "username contains invalid characters"
        Case &H3: reason = "username contained a banned word"
        Case &H4: reason = "username already exists"
        Case &H5: reason = "username is still being created"
        Case &H6: reason = "username does not contain enough alphanumeric characters"
        Case &H7: reason = "username contained adjacent punctuation characters"
        Case &H8: reason = "username contained too many punctuation characters"
    End Select

    accountIdToReason = reason
End Function

Public Sub setupSockets(previousConnectionCount As Integer, connectionCount As Integer)
    If (previousConnectionCount > 0) Then
        For i = 0 To previousConnectionCount - 1
            If (i > 0) Then
                Unload frmMain.sckBNLS(i)
                Unload frmMain.sckBNET(i)
                Unload frmMain.tmrBNETConnectionTimeout(i)
            End If
        Next i
    End If
  
    If (connectionCount > 0) Then
        ReDim bnlsPacketHandler(connectionCount - 1)
        ReDim bnetPacketHandler(connectionCount - 1)
    
        For i = 0 To connectionCount - 1
            If (i > 0) Then
                Load frmMain.sckBNLS(i)
                Load frmMain.sckBNET(i)
                Load frmMain.tmrBNETConnectionTimeout(i)
            End If
      
            Set bnlsPacketHandler(i) = New clsPacketHandler
            Set bnetPacketHandler(i) = New clsPacketHandler
  
            bnlsPacketHandler(i).setSocket frmMain.sckBNLS(i), packetType.BNLS
            bnetPacketHandler(i).setSocket frmMain.sckBNET(i), packetType.BNCS
            
            frmMain.tmrBNETConnectionTimeout(i).Interval = config.connectionTimeout
        Next i
    End If
End Sub

Public Sub killSocket(ByVal index As Integer)
    Dim ds As DisconnectStatus, activeConnections As Integer

    ds = disconnectSocket(index)
    activeConnections = countActiveConnections()
    
    If (activeConnections = 0) Then
        finishDisconnectAll
    End If
    
    showDisconnectMessage ds, activeConnections = 0, index
End Sub

Public Function disconnectSocket(ByVal index As Integer) As DisconnectStatus
    Dim ds As DisconnectStatus

    frmMain.tmrBNETConnectionTimeout(index).Enabled = False

    If frmMain.sckBNLS(index).State <> sckClosed Then
        frmMain.sckBNLS(index).Close

        ds.disconnectedBNLS = True
    End If
    
    If (frmMain.sckBNET(index).State <> sckClosed) Then
        frmMain.sckBNET(index).Close
    
        ds.disconnectedBNET = True
    End If

    disconnectSocket = ds
End Function

Public Sub disconnectAll()
    Dim ds As DisconnectStatus, dsAll As DisconnectStatus
  
    For i = 0 To frmMain.sckBNET.count - 1
        ds = disconnectSocket(i)
    
        If (ds.disconnectedBNLS) Then
            If (Not dsAll.disconnectedBNLS) Then
                dsAll.disconnectedBNLS = True
            End If
        End If
        
        If (ds.disconnectedBNET) Then
            If (Not dsAll.disconnectedBNET) Then
                dsAll.disconnectedBNET = True
            End If
        End If
    Next i
  
    finishDisconnectAll
    
    showDisconnectMessage dsAll, True
End Sub

Public Sub finishDisconnectAll()
    frmMain.mnuDisconnectBNET.Enabled = False
    frmMain.mnuConnectBNET.Enabled = True
End Sub

Public Sub showDisconnectMessage(ds As DisconnectStatus, ByVal allSocketsDisconnect As Boolean, Optional ByVal index As Integer = -1)
    If (ds.disconnectedBNLS) Then
        If (allSocketsDisconnect) Then
            AddChat frmMain.rtbChatBNET, vbRed, IIf(index > -1, "Bot #" & index & ": ", "") & "[BNLS] All connections closed."
        Else
            AddChat frmMain.rtbChatBNET, vbRed, IIf(index > -1, "Bot #" & index & ": ", "") & "[BNLS] Connection has been closed."
        End If
    End If
    
    If (ds.disconnectedBNET) Then
        If (allSocketsDisconnect) Then
            AddChat frmMain.rtbChatBNET, vbRed, IIf(index > -1, "Bot #" & index & ": ", "") & "[BNET] All connections closed."
        Else
            AddChat frmMain.rtbChatBNET, vbRed, IIf(index > -1, "Bot #" & index & ": ", "") & "[BNET] Connection has been closed."
        End If
    End If
End Sub

Public Function countActiveConnections() As Integer
    Dim stateBNLS As Integer, stateBNET As Integer
    Dim count As Integer

    If (hasLoadedConnections) Then
        For i = 0 To profiles.getCount() - 1
            stateBNLS = frmMain.sckBNLS(i).State
            stateBNCS = frmMain.sckBNET(i).State
            
            If (stateBNLS <> sckClosed Or stateBNCS <> sckClosed) Then
                count = count + 1
            End If
        Next i
    End If
    
    countActiveConnections = count
End Function

Public Function joinArrayAtIndex(arr() As String, index As Integer)
    Dim finalString As String

    For i = 0 To UBound(arr)
        If (i >= index) Then
            If (finalString <> vbNullString) Then
                finalString = finalString & " "
            End If
            
            finalString = finalString & arr(i)
        End If
    Next i
    
    joinArrayAtIndex = finalString
End Function

Public Function hexToString(ByVal val As Long) As String
    hexToString = Right("0" & hex(val), 2)
End Function

Public Function makeCompatibleDate(ByVal dateTimeString As String) As Date
    dateTimeString = Replace(dateTimeString, "T", " ")
    dateTimeString = Replace(dateTimeString, "Z", "")

    makeCompatibleDate = dateTimeString
End Function

Public Function KillNull(ByVal text As String) As String
    Dim pos As Integer
  
    pos = InStr(text, Chr$(0))
  
    KillNull = IIf(pos > 0, Mid$(text, 1, pos - 1), text)
End Function

Public Function checkProgramUpdate(ByVal manualUpdateCheck As Boolean) As Boolean
    On Error GoTo err
    
    Dim text As String, status As Integer, requestReleaseTime As Date, releaseTime As Date, requestVersion As String, Version As String
    Dim isoRequestReleaseTime As String, isoReleaseTime As String
    Dim jsonResponse As Dictionary, jsonContents As Dictionary
    Dim updateMsg As String, msgBoxResult As Integer
    Dim xml As Object
    
    Set xml = CreateObject("MSXML2.XMLHTTP")

    xml.Open "GET", PROGRAM_UPDATE_URL, False
    xml.setRequestHeader "User-Agent", "BattleNetToIRC/" & PROGRAM_VERSION
    xml.send
    
    text = xml.responseText
    Set jsonResponse = JSON.parse(text)
    status = jsonResponse.Item("status")
    
    If (status = 1) Then
        Set jsonContents = jsonResponse.Item("contents")
        
        isoRequestReleaseTime = jsonContents.Item("request_release_time")
        requestVeresion = jsonContents.Item("request_version")
        isoReleaseTime = jsonContents.Item("release_time")
        Version = jsonContents.Item("version")
        
        requestReleaseTime = makeCompatibleDate(isoRequestReleaseTime)
        releaseTime = makeCompatibleDate(isoReleaseTime)
        
        If (releaseTime > requestReleaseTime) Then
            updateMsg = "There is a new update for " & PROGRAM_NAME & "!" & vbNewLine & vbNewLine & "Your version: " & PROGRAM_VERSION & " new version: " & Version & vbNewLine & vbNewLine _
                      & "Would you like to view the changelog and download the latest update?"
        
            msgBoxResult = MsgBox(updateMsg, vbYesNo Or vbInformation, "New version for " & PROGRAM_NAME)
    
            If (msgBoxResult = vbYes) Then
                ShellExecute 0, "open", UPDATE_SUMMARY_URL, vbNullString, vbNullString, 4
            End If
        Else
            If (manualUpdateCheck) Then
                MsgBox "There is no new version at this time.", vbOKOnly Or vbInformation, PROGRAM_TITLE
            End If
        End If
        
        checkProgramUpdate = True
        Exit Function
    End If

err:
    Set xml = Nothing
End Function
