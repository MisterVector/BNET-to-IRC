Attribute VB_Name = "modBNET"
Public Sub Recv0x25(index As Integer)
  Send0x25 index
End Sub

Public Sub Send0x25(index As Integer)
  With bnetPacketBuffer(index)
    .InsertDWORD .GetDWORD
    .sendPacket &H25, False, index
  End With
End Sub

Public Sub Send0x50(index As Integer)
  With bnetPacketBuffer(index)
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
    .sendPacket &H50, False, index
  End With
End Sub

Public Sub Recv0x50(index As Integer)
  With bnetPacketBuffer(index)
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
  With bnetPacketBuffer(index)
    .InsertDWORD bnetData(index).clientToken
    .InsertDWORD bnetData(index).exeVersion
    .InsertDWORD bnetData(index).checksum
    .InsertDWORD &H1
    .InsertDWORD &H0
    .InsertDWORD bnetData(index).cdKeyLength
    .InsertDWORD bnetData(index).cdKeyProductValue
    .InsertDWORD bnetData(index).cdKeyPublicValue
    .InsertDWORD &H0
    .InsertNonNTString bnetData(index).cdKeyHash
    .InsertNTString bnetData(index).exeInfo
    .InsertNTString "BNET to IRC"
    .sendPacket &H51, False, index
  End With
End Sub

Public Sub Recv0x51(index As Integer)
  Dim results As Long
  
  results = bnetPacketBuffer(index).GetDWORD
  
  Select Case results
    Case &H0:    AddChat frmMain.rtbChatBNET, vbGreen, "Bot #" & index & " [BNET] CDKey is accepted."
    Case &H100:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & " [BNET] Your game is out of date."
    Case &H101:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & " [BNET] Invalid game version."
    Case &H102:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & " [BNET] Downgrade game version."
    Case &H200:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & " [BNET] CDKey is invalid."
    Case &H201:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & " [BNET] CDKey in use by " & bnetPacketBuffer(index).getNTString & "."
                 frmMain.sckBNET(index).Close
                 frmMain.sckBNLS(index).Close
    Case &H202:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & " [BNET] Key is banned."
    Case &H203:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & " [BNET] Key is for nother product."
    Case &H210:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & " [BNET] Expansion key is invalid."
    Case &H211:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & " [BNET] Expansion key is in use by" & bnetPacketBuffer(index).getNTString & "."
    Case &H212:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & " [BNET] Expansion key is banned."
  End Select
  
  If results = &H0 Then
    Send0x3A index
  End If
End Sub

Public Sub Send0x3A(index As Integer)
  AddChat frmMain.rtbChatBNET, vbYellow, "Bot #" & index & ": [BNET] Logging in..."
  
  With bnetPacketBuffer(index)
    .InsertDWORD bnetData(index).clientToken
    .InsertDWORD bnetData(index).serverToken
    .InsertNonNTString bnetData(index).passwordHash
    .InsertNTString frmMain.txtUsername.text
    .sendPacket &H3A, False, index
  End With
End Sub

Public Sub Recv0x3A(index As Integer)
  Select Case bnetPacketBuffer(index).GetDWORD
    Case &H0: AddChat frmMain.rtbChatBNET, vbGreen, "Bot #" & index & ": Password accepted!"
              ConnectOtherBots
              Send0x0A index
              Send0x0B index
              Send0x0C index
    Case &H1: AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": Account does not exist!"
              AddChat frmMain.rtbChatBNET, vbYellow, "Bot #" & index & ": Creating account..."
              newAccFlag = True
              frmMain.sckBNLS(index).Connect config.bnlsServer, 9367
              
    Case &H2: AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": Password is invalid!"
    Case &H6: AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": Account is closed: " & bnetPacketBuffer(index).getNTString
  End Select
End Sub

Public Sub Recv0x3D(index As Integer)
  Select Case bnetPacketBuffer(index).GetDWORD
    Case &H0: AddChat frmMain.rtbChatBNET, vbGreen, "Bot #" & index & ": Account created!"
              Send0x3A index
    Case &H2: AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": Account contained invalid characters"
    Case &H3: AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": Account contained a banned words."
    Case &H4: AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": Account already exists!"
    Case &H6: AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": Not enough characters!"
  End Select
End Sub

Public Sub Send0x0A(index As Integer)
  With bnetPacketBuffer(index)
    .InsertNTString vbNullString
    .InsertNTString vbNullString
    .sendPacket &HA, False, index
  End With
End Sub

Public Sub Recv0x0A(index As Integer)
  bnetData(index).uniqueName = bnetPacketBuffer(index).getNTString
  bnetPacketBuffer(index).getNTString 'skip statstring
  bnetData(index).accountName = bnetPacketBuffer(index).getNTString
  
  AddChat frmMain.rtbChatBNET, vbGreen, "Bot #" & index & ": Logged in as " & bnetData(index).uniqueName
End Sub

Public Sub Send0x0B(index As Integer)
  With bnetPacketBuffer(index)
    .InsertDWORD &H0
    .sendPacket &HB, False, index
  End With
End Sub

Public Sub Send0x0C(index As Integer)
  With bnetPacketBuffer(index)
    .InsertDWORD &H2 'IIf(bnetData(index).product = "D2DV", &H5, &H1)
    .InsertNTString config.bnetChannel
    .sendPacket &HC, False, index
  End With
End Sub

Public Sub Recv0x0F(index As Integer)
  Dim user As String, text As String, ID As Long, flags As Long
  Dim supressEvent As Boolean

  If index <> findFirstAliveBot Then supressEvent = True
  
  With bnetPacketBuffer(index)
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
                    SendToIRC IIf(ID = &H17, "/me ", "") & user & ": " & text
                  End If
                End If
      Case &H7: AddChat frmMain.rtbChatBNET, vbYellow, "You joined the channel ", vbWhite, text
                If isBroadcastToIRC Then SendToIRC bnetData(index).uniqueName & " has joined the Battle.Net channel " & text
    End Select
  End With
End Sub
