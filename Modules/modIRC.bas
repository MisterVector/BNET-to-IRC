Attribute VB_Name = "modIRC"
Public Sub handleIRCData(ByVal source As String, ByVal hostname As String, ByVal command As String, ByVal data As String)
    Dim connectedUsername As String
    
    connectedUsername = IRCData.connectedUsername
    
    Select Case command
        Case "353"
            AddChat frmMain.rtbChatIRCChat, vbYellow, data
        Case "376"
            SendJOIN config.ircChannel
        Case "JOIN"
            RecvJOIN data
        Case "PRIVMSG"
            RecvPRIVMSG source, hostname, data
        Case "PING"
             RecvPING data
        Case Else
            If (Len(data) >= Len(connectedUsername + " :")) Then
                If (Left$(data, Len(connectedUsername + " :")) = connectedUsername + " :") Then
                    data = Mid$(data, Len(connectedUsername + " :") + 1)
                End If
            End If
            
            If (Len(data) >= Len("* :*** ")) Then
                If (Left$(data, Len("* :*** "))) = "* :*** " Then
                    data = Mid$(data, Len("* :*** ") + 1)
                End If
            End If
            
            AddChat frmMain.rtbChatIRCConsole, vbYellow, data
    End Select
End Sub

Public Sub SendJOIN(ByVal channel As String)
    frmMain.sckIRC.SendData "JOIN " & channel & vbCrLf
End Sub

Public Sub RecvJOIN(ByVal channel As String)
    frmMain.IRCTab.TabCaption(1) = "Chat (" & channel & ")"
    IRCData.joinedChannel = channel
    
    If (config.ircUpdateChannelOnChannelJoin) Then
        If (channel <> config.ircChannel) Then
            config.ircChannel = channel
            
            saveConfig
        End If
    End If

    AddChat frmMain.rtbChatIRCChat, vbYellow, "Joined the channel ", vbWhite, channel, vbYellow, "."
    
    If (Not canSendQuit) Then
        canSendQuit = True
    End If
End Sub

Public Sub SendNICK(ByVal username As String)
    frmMain.sckIRC.SendData "NICK " & username & vbCrLf
End Sub

Public Sub SendPART(ByVal channel As String)
    frmMain.sckIRC.SendData "PART " & channel & vbCrLf
End Sub

Public Sub SendPING(ByVal data As String)
    frmMain.sckIRC.SendData "PONG " & data & vbCrLf
    AddChat frmMain.rtbChatIRCConsole, vbWhite, "PING has been PONG'D"
End Sub

Public Sub RecvPING(ByVal data As String)
    SendPING data
End Sub

Public Sub SendPRIVMSG(ByVal target As String, text As String)
    frmMain.sckIRC.SendData "PRIVMSG " & target & " :" & text & vbCrLf
End Sub

Public Sub RecvPRIVMSG(ByVal source As String, ByVal hostname As String, ByVal text As String)
    Dim arrTextData() As String, msgTarget As String, msg As String, msgPart As String
    
    arrTextData = Split(text, " ", 2)
    msgTarget = arrTextData(0)
    msg = Mid$(arrTextData(1), 2)
    
    If (msgTarget = config.ircChannel) Then
        AddChat frmMain.rtbChatIRCChat, vbYellow, source, vbWhite, ": ", vbYellow, msg
    Else
        AddChat frmMain.rtbChatIRCChat, vbYellow, source & " (", vbWhite, msgTarget, vbYellow, ")", vbWhite, ": ", vbYellow, msg
    End If
    
    ' If broadcast prefix is set, check to see if it is present at the start of the message
    If (config.ircBroadcastPrefix <> vbNullString) Then
        If (Len(msg) < Len(config.ircBroadcastPrefix)) Then
            Exit Sub
        End If

        msgPart = Mid(msg, 1, Len(config.ircBroadcastPrefix))

        If (msgPart <> config.ircBroadcastPrefix) Then
            Exit Sub
        End If
    End If

    SendToBNET source & ": " & text
End Sub

Public Sub SendUSER(ByVal username As String)
    frmMain.sckIRC.SendData "USER " & username & " 0 0 " & username & vbCrLf
End Sub
