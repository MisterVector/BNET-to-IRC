Attribute VB_Name = "modOtherCode"
Private Sub AddQ(ByVal msg As String)
  dicQueue.Add dicQueue.Count + 1, msg

  If Not frmMain.tmrReleaseQueue.Enabled Then
    dicQueueIndex = 1
    frmMain.tmrReleaseQueue.Enabled = True
  End If
End Sub

Public Function getVerByte(ByVal product As String) As Long
  Select Case product
    Case "W2BN": getVerByte = VERBYTE_W2BN
    Case "D2DV": getVerByte = VERBYTE_D2DV
    Case "WAR3": getVerByte = VERBYTE_WAR3
  End Select
End Function

Public Function getProdID(ByVal product As String) As Long
  Select Case product
    Case "W2BN": getProdID = &H3
    Case "D2DV": getProdID = &H4
    Case "WAR3": getProdID = &H7
  End Select
End Function

Public Function findFirstAliveBot() As Integer
  For i = 0 To frmMain.sckBNET.Count - 1
    If frmMain.sckBNET(i).State = sckConnected Then
      findFirstAliveBot = i
      Exit Function
    End If
  Next i

  findFirstAliveBot = -1
End Function

Public Sub SendToBNET(ByVal msg As String)
  If (findFirstAliveBot() = -1) Then Exit Sub

  If Len(msg) > 140 Then
    Do While Len(msg) > 140
      AddQ Mid(msg, 1, 140) & " [more]"
      msg = Mid(msg, 141)
    Loop
  
    If Len(msg) > 0 Then
      AddQ msg
    End If
  Else
    AddQ msg
  End If
End Sub

Public Sub SendToIRC(ByVal msg As String)
  If frmMain.sckIRC.State = sckConnected Then
    frmMain.sckIRC.SendData "PRIVMSG " & config.ircChannel & " :" & msg & vbCrLf
  End If
End Sub

Public Sub ConnectOtherBots()
  If config.bnetKeyCount > 1 Then
    For i = 1 To frmMain.sckBNET.Count - 1
      If frmMain.sckBNET(i).State = sckClosed Then
        frmMain.sckBNET(i).Connect config.bnetServer, 6112
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

Public Function accountIdToReason(ByVal ID As Long, ByVal isWar3 As Boolean) As String
  Dim reason As String

  If isWar3 Then
    Select Case ID
      Case &H4: reason = "username already exists."
      Case &H7: reason = "username is too short or blank."
      Case &H8: reason = "username contains an illegal character."
      Case &H9: reason = "username contains an illegal word."
      Case &HA: reason = "username contains too few alphanumeric characters."
      Case &HB: reason = "username contains adjacent punctuation characters."
      Case &HC: reason = "username contains too many punctuation characters."
      Case Else: reason = "username already exists."
    End Select
  Else
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
  End If

  accountIdToReason = reason
End Function

Public Sub setupSockets(previousConnectionCount As Integer, connectionCount As Integer)
  If (previousConnectionCount > 0) Then
    For i = 0 To previousConnectionCount - 1
      If (i > 0) Then
        Unload frmMain.sckBNLS(i)
        Unload frmMain.sckBNET(i)
      End If
    Next i
  End If
  
  ReDim bnlsPacketHandler(connectionCount - 1)
  ReDim bnetPacketHandler(connectionCount - 1)
  
  For i = 0 To connectionCount - 1
    If (i > 0) Then
      Load frmMain.sckBNLS(i)
      Load frmMain.sckBNET(i)
    End If
    
    Set bnlsPacketHandler(i) = New clsPacketHandler
    Set bnetPacketHandler(i) = New clsPacketHandler

    bnlsPacketHandler(i).setSocket frmMain.sckBNLS(i), packetType.BNLS
    bnetPacketHandler(i).setSocket frmMain.sckBNET(i), packetType.BNCS
  Next i
End Sub

Public Sub loadConfig()
  Dim val As Variant, parts() As String

  val = ReadINI("Window", "Top", "Config.ini")
  
  If (IsNumeric(val) And val > 0) Then
    config.formTop = val
  End If
  
  val = ReadINI("Window", "Left", "Config.ini")
  
  If (IsNumeric(val) And val > 0) Then
    config.formLeft = val
  End If

  config.bnetUsername = ReadINI("BNET", "Username", "Config.ini")
  config.bnetPassword = ReadINI("BNET", "Password", "Config.ini")
  config.bnetChannel = ReadINI("BNET", "Channel", "Config.ini")
  config.bnetServer = ReadINI("BNET", "Server", "Config.ini")
  config.bnlsServer = ReadINI("BNET", "BNLSServer", "Config.ini")
  
  val = ReadINI("BNET", "KeyCount", "Config.ini")
  
  If (IsNumeric(val)) Then
    config.bnetKeyCount = val
    
    If (config.bnetKeyCount > 0) Then
      setupSockets 0, config.bnetKeyCount
      
      ReDim bnetData(config.bnetKeyCount - 1)
    
      For i = 0 To config.bnetKeyCount - 1
        With bnetData(i)
          .product = ReadINI(i, "Product", "Config.ini")
          .cdKey = ReadINI(i, "CDKey", "Config.ini")
        End With
      Next i
    End If
  End If
  
  config.ircUsername = ReadINI("IRC", "Username", "Config.ini")
  config.ircChannel = ReadINI("IRC", "Channel", "Config.ini")
  
  val = ReadINI("IRC", "Server", "Config.ini")
  
  If (InStr(val, ":") > 0) Then
    parts = Split(val, ":")
  
    config.ircServer = parts(0)
    config.ircPort = parts(1)
  Else
    config.ircServer = val
    config.ircPort = 6667
  End If
End Sub

Public Sub saveConfig()
  WriteINI "Window", "Top", frmMain.Top, "Config.ini"
  WriteINI "Window", "Left", frmMain.Left, "Config.ini"
  
  WriteINI "BNET", "Username", config.bnetUsername, "Config.ini"
  WriteINI "BNET", "Password", config.bnetPassword, "Config.ini"
  WriteINI "BNET", "Channel", config.bnetChannel, "Config.ini"
  WriteINI "BNET", "Server", config.bnetServer, "Config.ini"
  WriteINI "BNET", "BNLSServer", config.bnlsServer, "Config.ini"
  WriteINI "BNET", "KeyCount", config.bnetKeyCount, "Config.ini"
  WriteINI "IRC", "Username", config.ircUsername, "Config.ini"
  WriteINI "IRC", "Channel", config.ircChannel, "Config.ini"
  WriteINI "IRC", "Server", config.ircServer, "Config.ini"
  
  If (config.bnetKeyCount > 0) Then
    For i = 0 To config.bnetKeyCount - 1
      With bnetData(i)
        WriteINI i, "Product", .product, "Config.ini"
        WriteINI i, "CDKey", .cdKey, "Config.ini"
      End With
    Next i
  End If
End Sub
