Attribute VB_Name = "modIRC"
Public Sub handleIRCData(ByVal source As String, ByVal hostname As String, ByVal command As String, ByVal data As String)
    Dim connectedUsername As String, parts() As String, target As String
    
    connectedUsername = IRCData.connectedUsername
    
    Select Case command
        Case "353"
            AddChat frmMain.rtbChatIRCChat, vbYellow, data
        Case "372", "375" 'MOTD
            AddChat frmMain.rtbChatIRCConsole, vbWhite, data
        Case "376" 'End of MOTD, now join home channel
            SendJOIN config.ircChannel
        Case "JOIN"
            RecvJOIN data
        Case "PRIVMSG"
            parts = Split(data, " ", 2)
            target = parts(0)
            data = Mid$(parts(1), 2)
        
            RecvPRIVMSG source, hostname, target, data
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
End Sub

Public Sub RecvPING(ByVal data As String)
    SendPING data
End Sub

Public Sub SendPRIVMSG(ByVal target As String, text As String)
    frmMain.sckIRC.SendData "PRIVMSG " & target & " :" & text & vbCrLf
End Sub

Public Sub RecvPRIVMSG(ByVal source As String, ByVal hostname As String, ByVal target As String, ByVal text As String)
    Dim msgPart As String, parts() As String, command As String, message As String, isEmote As Boolean

    isEmote = False

    ' Check for CTCP
    If (Left$(text, 1) = Chr$(&H1) And Right$(text, 1) = Chr$(&H1)) Then
        text = Mid(text, 2, Len(text) - 2) ' Length 2 to skip last null character
        parts = Split(text, " ", 2)
        command = parts(0)
        text = parts(1)
        
        Select Case command
            Case "ACTION"
                isEmote = True
            Case Else
                AddChat frmMain.rtbChatIRCConsole, vbRed, "CTCP command not supported: " & command
                Exit Sub
        End Select
    End If
    
    If (LCase(target) = LCase(config.ircChannel)) Then
        AddChat frmMain.rtbChatIRCChat, vbYellow, source, vbWhite, ": ", vbYellow, text
    Else
        AddChat frmMain.rtbChatIRCChat, vbYellow, source & " (", vbWhite, target, vbYellow, ")", vbWhite, ": ", vbYellow, text
        
        'Don't broadcast non-channel messages to Battle.Net
        Exit Sub
    End If
    
    'If broadcast prefix is set, check to see if it is present at the start of the message
    If (config.ircBroadcastPrefix <> vbNullString) Then
        If (Len(text) < Len(config.ircBroadcastPrefix)) Then
            Exit Sub
        End If

        msgPart = Mid(text, 1, Len(config.ircBroadcastPrefix))

        If (msgPart <> config.ircBroadcastPrefix) Then
            Exit Sub
        End If
    End If

    message = source & " (" & target & "): "

    If (isEmote) Then
        message = "/me " & message & source & " " & text
    Else
        message = message & text
    End If

    SendToBNET message
End Sub

Public Sub SendUSER(ByVal username As String)
    frmMain.sckIRC.SendData "USER " & username & " 0 0 " & username & vbCrLf
End Sub
