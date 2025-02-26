Attribute VB_Name = "modVars"
Public Const PROGRAM_VERSION                        As String = "1.2.0"
Public Const PROGRAM_NAME                           As String = "Battle.Net to IRC"
Public Const PROGRAM_TITLE                          As String = PROGRAM_NAME & " v" & PROGRAM_VERSION & " by Vector"
Public Const PROGRAM_SLUG                           As String = "battle-net-to-irc"
Public Const PROGRAM_UPDATE_URL                     As String = "https://distribution.codespeak.org/data_handler.php?query=check_program_version&slug=" & PROGRAM_SLUG & "&current_version=" & PROGRAM_VERSION
Public Const UPDATE_SUMMARY_URL                     As String = "https://www.codespeak.org/downloads/updatesummary/" & PROGRAM_SLUG & "/" & PROGRAM_VERSION

Public Const VERBYTE_W2BN                           As Long = &H4F
Public Const VERBYTE_D2DV                           As Long = &HE

Public Const DEFAULT_BNLS_SERVER                    As String = "jbls.davnit.net"
Public Const DEFAULT_USE_LOCAL_HASHING              As Boolean = False
Public Const DEFAULT_REMEMBER_WINDOW_POSITION       As Boolean = False
Public Const DEFAULT_CHECK_UPDATE_ON_STARTUP        As Boolean = True
Public Const DEFAULT_UPDATE_CHANNEL_ON_CHANNEL_JOIN As Boolean = False
Public Const DEFAULT_CONNECTION_TIMEOUT             As Integer = 10000
Public Const DEFAULT_MINIMIZE_TO_TRAY               As Boolean = False
Public Const DEFAULT_BROADCAST_TO_IRC               As Boolean = False
Public Const DEFAULT_BROADCAST_TO_BNET              As Boolean = False
Public Const DEFAULT_FOCUS_CHAT_TAB_ON_CHANNEL_JOIN As Boolean = False

Public Const LAST_NON_QUEUE_THRESHOLD_TIME          As Long = 7000

'// BNET SIDE
Public dicQueue As New Dictionary
Public dicQueueIndex As Integer

Public isBroadcastToIRC As Boolean
Public isBroadcastToBNET As Boolean
  
Public bnetSocketIndex As Integer
Public bnetQueueIndex As Integer

Public bnetPacketHandler() As clsPacketHandler
Public bnlsPacketHandler() As clsPacketHandler

Public updateString As String
Public manualUpdateCheck As Boolean

Public canSendQuit As Boolean

Public loadedFromCSDSClient As Boolean

Public isMinimizedToTray As Boolean

Public Enum packetType
    BNCS
    BNLS
End Enum

Public Enum ConnectionTimeoutState
    BNET_CONNECT
    BNLS_CONNECT
End Enum

Public Type DisconnectStatus
    disconnectedBNET As Boolean
    disconnectedBNLS As Boolean
End Type

Public Type ConfigStructure
    formTop As Integer
    formLeft As Integer
    rememberWindowPosition As Boolean
    checkUpdateOnStartup As Boolean
    connectionTimeout As Integer
    minimizeToTray As Boolean

    bnlsServer As String
    bnetServer As String
    bnetUsername As String
    bnetPassword As String
    bnetChannel As String
    bnetKeyCount As Integer
    bnetBroadcastPrefix As String

    bnetLocalHashing As Boolean
    bnetW2BNVerByte As Long
    bnetD2DVVerByte As Long

    ircUsername As String
    ircPassword As String
    ircServer As String
    ircChannel As String
    ircFocusChatTabOnChannelJoin As Boolean
    ircQuitMessage As String
    ircUpdateChannelOnChannelJoin As Boolean
    ircBroadcastPrefix As String
End Type
Public config As ConfigStructure

Public Type IRCDataStructure
    connectedUsername As String
    joinedChannel As String
End Type
Public IRCData As IRCDataStructure

Public Type BNETDataStructure
    accountName As String
    uniqueName As String
    currentChannel As String

    product As String
    passwordHash As String
    oldVerByte As Long
    verByte As Long
    clientToken As Long
    serverToken As Long
    dwLowDateTime As Long
    dwHighDateTime As Long
    valueString As String
    archiveFileName As String
    checksumFormula As String
    exeInfo As String
    exeVersion As String
    Checksum As String
    CDKey As String
  
    bnlsServerCode As Long
  
    nls_P As Long

    lastQueueTime As Long
    bnetConnectionState As ConnectionTimeoutState
End Type
Public bnetData() As BNETDataStructure
