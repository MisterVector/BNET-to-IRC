Attribute VB_Name = "modIRC"
Public Sub handleIRCData(ByVal source As String, ByVal hostname As String, ByVal command As String, ByVal data As String)
    Select Case command
        Case "PRIVMSG"
            RecvPRIVMSG source, hostname, data
        Case "PING"
             RecvPING data
        Case Else
            If (InStr(data, "End of /MOTD command.")) Then
                frmMain.sckIRC.SendData "JOIN " & config.ircChannel & vbCrLf
            End If

            AddChat frmMain.rtbChatIRCConsole, vbYellow, data
    End Select
End Sub

Public Sub SendPING(ByVal data As String)
    frmMain.sckIRC.SendData "PONG " & data & vbCrLf
    AddChat frmMain.rtbChatIRCConsole, vbWhite, "PING has been PONG'D"
End Sub

Public Sub RecvPING(ByVal data As String)
    SendPING data
End Sub

Public Sub RecvPRIVMSG(ByVal source As String, ByVal hostname As String, ByVal text As String)
    Dim arrTextData() As String, msgTarget As String, msg As String
    
    arrTextData = Split(text, " ", 2)
    msgTarget = arrTextData(0)
    msg = Mid$(arrTextData(1), 2)
    
    AddChat frmMain.rtbChatIRCChat, vbYellow, source & " (", vbWhite, msgTarget, vbYellow, ")", vbWhite, ": ", vbYellow, msg
    SendToBNET source & ": " & text
End Sub
