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
  With bnetPacketHandler(index)
    .InsertDWORD &H0
    .InsertNonNTString "68XI" & StrReverse(bnetData(index).product)
    .InsertDWORD getVerByte(bnetData(index).product)
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
  With bnetPacketHandler(index)
    .Skip 4
    bnetData(index).clientToken = GetTickCount
    bnetData(index).serverToken = .GetDWORD
    .Skip 12
    bnetData(index).lockdownFile = .getNTString
    bnetData(index).valueString = .getNTString
  End With

  AddChat frmMain.rtbChatBNET, vbYellow, "Bot #" & index & ": [BNLS] Connecting to " & config.bnlsServer & "..."

  frmMain.sckBNLS(index).Connect config.bnlsServer, 9367
End Sub

Public Sub Send0x51(index As Integer)
  Dim cdKeyLength As Long, cdKeyHash As String * 20, cdKeyProductValue As Long, cdKeyPublicValue As Long
  Dim result As Long

  cdKeyLength = Len(bnetData(index).cdKey)
  result = decode_hash_cdkey(bnetData(index).cdKey, bnetData(index).clientToken, bnetData(index).serverToken, cdKeyPublicValue, cdKeyProductValue, cdKeyHash)

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
  Dim results As Long
  
  results = bnetPacketHandler(index).GetDWORD
  
  Select Case results
    Case &H0:    AddChat frmMain.rtbChatBNET, vbGreen, "Bot #" & index & ": [BNET] CDKey is accepted."
    Case &H100:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] Your game is out of date."
    Case &H101:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] Invalid game version."
    Case &H102:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] Downgrade game version."
    Case &H200:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] CDKey is invalid."
    Case &H201:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] CDKey in use by " & bnetPacketHandler(index).getNTString & "."
                 frmMain.sckBNET(index).Close
                 frmMain.sckBNLS(index).Close
    Case &H202:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] Key is banned."
    Case &H203:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] Key is for another product."
    Case &H210:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] Expansion key is invalid."
    Case &H211:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] Expansion key is in use by " & bnetPacketHandler(index).getNTString & "."
    Case &H212:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] Expansion key is banned."
  End Select
  
  If results = &H0 Then
    If (bnetData(index).product = "WAR3") Then
      Send0x53 index
    Else
      Send0x3A index
    End If
  End If
End Sub

Public Sub Send0x52(index As Integer)
  Dim saltHash As String: saltHash = Space(Len(config.bnetUsername) + 65)
  
  nls_account_create bnetData(index).nls_P, saltHash

  With bnetPacketHandler(index)
    .InsertNonNTString saltHash
    .sendPacket &H52
  End With
End Sub

Public Sub Recv0x52(index As Integer)
  Dim result As Long

  With bnetPacketHandler(index)
    result = .GetDWORD
  End With
  
  If result = &H0 Then
    AddChat frmMain.rtbChatBNET, vbGreen, "Bot #" & index & ": [BNET] Created the Warcraft III account " & config.bnetUsername & "!"
    Send0x53 index
  Else
    Dim reason As String
    reason = accountIdToReason(result, True)
  
    AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] Could not create the account " & config.bnetUsername & "!"
    AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] Reason: " & reason & "."
  
    disconnectAll
  End If
End Sub

Public Sub Send0x53(index As Integer)
  Dim nls_A As String

  bnetData(index).nls_P = nls_init(config.bnetUsername, config.bnetPassword)

  If bnetData(index).nls_P = 0 Then
    MsgBox "NLS made a bad call.", vbOKOnly, PROGRAM_TITLE
    
    disconnectAll
    Exit Sub
  End If

  nls_A = Space(Len(config.bnetUsername) + 33)
  
  If (nls_account_logon(bnetData(index).nls_P, nls_A) = 0) Then
    MsgBox "Unable to create NLS key.", vbOKOnly, PROGRAM_TITLE
    
    disconnectAll
    Exit Sub
  End If

  bnetPacketHandler(index).InsertNonNTString Left$(nls_A, Len(nls_A) - Len(config.bnetUsername) - 1)
  bnetPacketHandler(index).InsertNTString config.bnetUsername
  bnetPacketHandler(index).sendPacket &H53
End Sub

Public Sub Recv0x53(index As Integer)
  Select Case bnetPacketHandler(index).GetDWORD
    Case &H0: Send0x54 index   'Passed
    Case &H1: 'Account Not made
              AddChat frmMain.rtbChatBNET, vbYellow, "Bot #" & index & ": [BNET] " & config.bnetUsername & " does not exist. It will be created."
              Send0x52 index
    Case &H5                   'Upgrade...
    Case Else
      disconnectAll
  End Select
End Sub

Public Sub Send0x54(index As Integer)
  Dim proofHash As String * 20
  Dim salt      As String: salt = bnetPacketHandler(index).GetNonNTString(32)
  Dim serverKey As String: serverKey = bnetPacketHandler(index).GetNonNTString(32)

  nls_account_logon_proof bnetData(index).nls_P, proofHash, serverKey, salt

  bnetPacketHandler(index).InsertNonNTString proofHash
  bnetPacketHandler(index).sendPacket &H54
End Sub

Public Sub Recv0x54(index As Integer)
  Select Case bnetPacketHandler(index).GetDWORD
    Case &H0, &HE: AddChat frmMain.rtbChatBNET, vbGreen, "Bot #" & index & ": [BNET] Password accepted!"
              GoTo Continue
    Case &H1:
    Case &H2:
    Case &HF:
  End Select
    
  disconnectAll
    
  Exit Sub
Continue:

  nls_free (bnetData(index).nls_P)        'Unloads the NLS object to avoid overhead

  ConnectOtherBots
  Send0x14 index
  Send0x0A index
  Send0x0C index
End Sub

Public Sub Send0x14(index As Integer)
  bnetPacketHandler(index).InsertNonNTString "tenb"
  bnetPacketHandler(index).sendPacket &H14
End Sub

Public Sub Send0x3A(index As Integer)
  Dim hashCode As String * 20
  
  AddChat frmMain.rtbChatBNET, vbYellow, "Bot #" & index & ": [BNET] Logging in..."

  double_hash_password config.bnetPassword, bnetData(index).clientToken, _
                       bnetData(index).serverToken, hashCode

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
    Case &H0: AddChat frmMain.rtbChatBNET, vbGreen, "Bot #" & index & ": [BNET] Password accepted!"
              ConnectOtherBots
              Send0x0A index
              Send0x0B index
              Send0x0C index
    Case &H1: AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] Account does not exist!"
              AddChat frmMain.rtbChatBNET, vbYellow, "Bot #" & index & ": [BNET] Creating account..."
              Send0x3D index
    Case &H2: AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] Password is invalid!"
    Case &H6: AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": [BNET] Account is closed: " & bnetPacketHandler(index).getNTString
  End Select
End Sub

Public Sub Send0x3D(index As Integer)
  Dim passwordHash As String * 20
  
  hash_password config.bnetPassword, passwordHash
  
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
  Dim supressEvent As Boolean

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
                  If isBroadcastToIRC Then SendToIRC user & " has joined " & text & "."
                End If
      Case &H3:
                If Not supressEvent Then
                  AddChat frmMain.rtbChatBNET, vbWhite, user, vbYellow, " left " & text
                  If isBroadcastToIRC Then SendToIRC user & " has left " & text & "."
                End If
      Case &H5, &H17:
                If Not supressEvent Then
                  AddChat frmMain.rtbChatBNET, vbWhite, "<" & user & "> ", vbYellow, text
        
                  For i = 0 To UBound(bnetData)
                    If Left(LCase(user), Len(bnetData(i).accountName)) = LCase(bnetData(i).accountName) Then
                      Exit Sub
                    End If
                  Next i
                  
                  If isBroadcastToIRC Then
                    'SendToIRC "(" & text & " @ " & config.bnetServer & ") " & User & ": " & Text
                    SendToIRC IIf(ID = &H17, "/me ", vbNullString) & user & ": " & text
                  End If
                End If
      Case &H7: AddChat frmMain.rtbChatBNET, vbYellow, "You joined the channel ", vbWhite, text
                If isBroadcastToIRC Then SendToIRC bnetData(index).uniqueName & " has joined the Battle.Net channel " & text
    End Select
  End With
End Sub
