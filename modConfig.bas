Attribute VB_Name = "modConfig"

Public Function readBooleanValue(ByVal section As String, ByVal key As String, Optional ByVal defaultValue As Boolean = False) As Boolean
    Dim val As String
    
    val = ReadINI(section, key, "Config.ini")
    
    If (val <> vbNullString And UCase(val) = "Y") Then
        readBooleanValue = True
    Else
        readBooleanValue = defaultValue
    End If
End Function

Public Sub writeBooleanValue(ByVal section As String, ByVal key As String, ByVal val As Boolean)
    WriteINI section, key, IIf(val, "Y", "N"), "Config.ini"
End Sub

Public Function readNumericValue(ByVal section As String, ByVal key As String, Optional ByVal defaultValue As Integer = 0) As Integer
    Dim val As String
    
    val = ReadINI(section, key, "Config.ini")
    
    If (IsNumeric(val)) Then
        If (val > 0 And val <= 32767) Then
            readNumericValue = val
            Exit Function
        End If
    End If
    
    readNumericValue = defaultValue
End Function

Public Sub writeNumericValue(ByVal section As String, ByVal key As String, ByVal val As Integer)
    WriteINI section, key, val, "Config.ini"
End Sub

Public Function readHexValue(ByVal section As String, ByVal key As String, Optional ByVal defaultValue As Long = 0) As Long
    Dim val As String
    
    val = ReadINI(section, key, "Config.ini")
    
    If (val <> vbNullString) Then
        val = "&H" & val
    End If

    If (IsNumeric(val)) Then
        readHexValue = val
    Else
        readHexValue = defaultValue
    End If
End Function

Public Sub writeHexValue(ByVal section As String, ByVal key As String, ByVal val As Long)
    WriteINI section, key, Right("0" & hex(val), 2), "Config.ini"
End Sub

Public Function readStringValue(ByVal section As String, ByVal key As String, Optional ByVal defaultValue As String = vbNullString) As String
    Dim val As String
    
    val = ReadINI(section, key, "Config.ini")
    
    If (val <> vbNullString) Then
        readStringValue = val
    Else
        readStringValue = defaultValue
    End If
End Function

Public Sub writeStringValue(ByVal section As String, ByVal key As String, ByVal val As String)
    WriteINI section, key, val, "Config.ini"
End Sub

Public Sub loadConfig()
    config.rememberWindowPosition = readBooleanValue("Window", "RememberWindowPosition", DEFAULT_REMEMBER_WINDOW_POSITION)
    config.formTop = readNumericValue("Window", "Top")
    config.formLeft = readNumericValue("Window", "Left")
  
    config.checkUpdateOnStartup = readBooleanValue("Main", "CheckUpdateOnStartup", DEFAULT_CHECK_UPDATE_ON_STARTUP)
    config.minimizeToTray = readBooleanValue("Main", "MinimizeToTray", DEFAULT_MINIMIZE_TO_TRAY)
    config.connectionTimeout = readNumericValue("Main", "ConnectionTimeout", DEFAULT_CONNECTION_TIMEOUT)
  
    config.bnetUsername = readStringValue("BNET", "Username")
    config.bnetPassword = readStringValue("BNET", "Password")
    config.bnetChannel = readStringValue("BNET", "Channel")
    config.bnetServer = readStringValue("BNET", "Server")
    config.bnlsServer = readStringValue("BNET", "BNLSServer", DEFAULT_BNLS_SERVER)
    config.bnetBroadcastPrefix = readStringValue("BNET", "BroadcastPrefix")

    config.bnetKeyCount = readNumericValue("BNET", "KeyCount")
    
    If (config.bnetKeyCount > 0) Then
        setupSockets 0, config.bnetKeyCount
    
        ReDim bnetData(config.bnetKeyCount - 1)
  
        For i = 0 To config.bnetKeyCount - 1
            With bnetData(i)
                .product = readStringValue(i, "Product")
                .CDKey = readStringValue(i, "CDKey")
            End With
        Next i
    End If
  
    config.bnetLocalHashing = readBooleanValue("BNET", "LocalHashing", DEFAULT_USE_LOCAL_HASHING)
    config.bnetW2BNVerByte = readHexValue("BNET", "W2BNVerByte", VERBYTE_W2BN)
    config.bnetD2DVVerByte = readHexValue("BNET", "D2DVVerByte", VERBYTE_D2DV)
    
    config.ircUsername = readStringValue("IRC", "Username")
    config.ircChannel = readStringValue("IRC", "Channel")
    config.ircQuitMessage = readStringValue("IRC", "QuitMessage")
    config.ircUpdateChannelOnChannelJoin = readBooleanValue("IRC", "UpdateChannelOnChannelJoin", DEFAULT_UPDATE_CHANNEL_ON_CHANNEL_JOIN)
    config.ircBroadcastPrefix = readStringValue("IRC", "BroadcastPrefix")
    config.ircServer = readStringValue("IRC", "Server")
End Sub

Public Sub saveConfig()
    writeBooleanValue "Window", "RememberWindowPosition", config.rememberWindowPosition
    writeBooleanValue "Main", "CheckUpdateOnStartup", config.checkUpdateOnStartup
    writeNumericValue "Main", "ConnectionTimeout", config.connectionTimeout
    writeBooleanValue "Main", "MinimizeToTray", config.minimizeToTray
    
    writeStringValue "BNET", "Username", config.bnetUsername
    writeStringValue "BNET", "Password", config.bnetPassword
    writeStringValue "BNET", "Channel", config.bnetChannel
    writeStringValue "BNET", "Server", config.bnetServer
    writeStringValue "BNET", "BNLSServer", config.bnlsServer
    writeNumericValue "BNET", "KeyCount", config.bnetKeyCount
    writeStringValue "BNET", "BroadcastPrefix", config.bnetBroadcastPrefix
    
    writeBooleanValue "BNET", "LocalHashing", config.bnetLocalHashing
    writeHexValue "BNET", "W2BNVerByte", config.bnetW2BNVerByte
    writeHexValue "BNET", "D2DVVerByte", config.bnetD2DVVerByte
    
    writeStringValue "IRC", "Username", config.ircUsername
    writeStringValue "IRC", "Channel", config.ircChannel
    writeStringValue "IRC", "Server", config.ircServer
    writeStringValue "IRC", "QuitMessage", config.ircQuitMessage
    writeBooleanValue "IRC", "UpdateChannelOnChannelJoin", config.ircUpdateChannelOnChannelJoin
    writeStringValue "IRC", "BroadcastPrefix", config.ircBroadcastPrefix
  
    If (config.bnetKeyCount > 0) Then
        For i = 0 To config.bnetKeyCount - 1
            With bnetData(i)
                writeStringValue i, "Product", .product
                writeStringValue i, "CDKey", .CDKey
            End With
        Next i
    End If
End Sub


