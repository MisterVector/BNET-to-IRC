Attribute VB_Name = "modBNET"
Public Sub Recv0x25(index As Integer)
  Send0x25 index
End Sub

Public Sub Send0x25(index As Integer)
  With pBNET(index)
    .InsertDWORD .GetDWORD
    .sendPacket &H25, False, index
  End With
End Sub

Public Sub Send0x50(index As Integer)
  With pBNET(index)
    .InsertDWORD &H0
    .InsertNonNTString "68XI" & StrReverse(BNET(index).prodStr)
    .InsertDWORD getVerByte(BNET(index).prodStr)
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
  With pBNET(index)
    .Skip 4
    BNET(index).ClientToken = GetTickCount
    BNET(index).ServerToken = .GetDWORD
    .Skip 12
    BNET(index).LockdownFile = .getNTString
    BNET(index).ValueString = .getNTString
  End With
  
  AddChat frmMain.rtbChatBNET, vbYellow, "Bot #" & index & ": [BNLS] Connecting to " & BNLSServer & "..."
  frmMain.sckBNLS(index).Connect BNLSServer, 9367
End Sub

Public Sub Send0x51(index As Integer)
  With pBNET(index)
    .InsertDWORD BNET(index).ClientToken
    .InsertDWORD BNET(index).EXEVersion
    .InsertDWORD BNET(index).checksum
    .InsertDWORD &H1
    .InsertDWORD &H0
    .InsertDWORD BNET(index).CDKeyLength
    .InsertDWORD BNET(index).CDKeyProductValue
    .InsertDWORD BNET(index).CDKeyPublicValue
    .InsertDWORD &H0
    .InsertNonNTString BNET(index).CDKeyHash
    .InsertNTString BNET(index).EXEInfo
    .InsertNTString "BNET to IRC"
    .sendPacket &H51, False, index
  End With
End Sub

Public Sub Recv0x51(index As Integer)
  Dim results As Long
  
  results = pBNET(index).GetDWORD
  Select Case results
    Case &H0:    AddChat frmMain.rtbChatBNET, vbGreen, "Bot #" & index & " [BNET] CDKey is accepted."
    Case &H100:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & " [BNET] Your game is out of date."
    Case &H101:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & " [BNET] Invalid game version."
    Case &H102:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & " [BNET] Downgrade game version."
    Case &H200:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & " [BNET] CDKey is invalid."
    Case &H201:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & " [BNET] CDKey in use by " & pBNET(index).getNTString & "."
                 frmMain.sckBNET(index).Close
                 frmMain.sckBNLS(index).Close
    Case &H202:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & " [BNET] Key is banned."
    Case &H203:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & " [BNET] Key is for nother product."
    Case &H210:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & " [BNET] Expansion key is invalid."
    Case &H211:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & " [BNET] Expansion key is in use by" & pBNET(index).getNTString & "."
    Case &H212:  AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & " [BNET] Expansion key is banned."
  End Select
  
  If results = &H0 Then
    Send0x3A index
  End If
End Sub

Public Sub Send0x3A(index As Integer)
  AddChat frmMain.rtbChatBNET, vbYellow, "Bot #" & index & ": [BNET] Logging in..."
  With pBNET(index)
    .InsertDWORD BNET(index).ClientToken
    .InsertDWORD BNET(index).ServerToken
    .InsertNonNTString BNET(index).PasswordHash
    .InsertNTString frmMain.txtUsername.text
    .sendPacket &H3A, False, index
  End With
End Sub

Public Sub Recv0x3A(index As Integer)
  Select Case pBNET(index).GetDWORD
    Case &H0: AddChat frmMain.rtbChatBNET, vbGreen, "Bot #" & index & ": Password accepted!"
              ConnectOtherBots
              Send0x0A index
              Send0x0B index
              Send0x0C index
    Case &H1: AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": Account does not exist!"
              AddChat frmMain.rtbChatBNET, vbYellow, "Bot #" & index & ": Creating account..."
              newAccFlag = True
              frmMain.sckBNLS(index).Connect BNLSServer, 9367
              
    Case &H2: AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": Password is invalid!"
    Case &H6: AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": Account is closed: " & pBNET(index).getNTString
  End Select
End Sub

Public Sub Recv0x3D(index As Integer)
  Select Case pBNET(index).GetDWORD
    Case &H0: AddChat frmMain.rtbChatBNET, vbGreen, "Bot #" & index & ": Account created!"
              Send0x3A index
    Case &H2: AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": Account contained invalid characters"
    Case &H3: AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": Account contained a banned words."
    Case &H4: AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": Account already exists!"
    Case &H6: AddChat frmMain.rtbChatBNET, vbRed, "Bot #" & index & ": Not enough characters!"
  End Select
End Sub

Public Sub Send0x0A(index As Integer)
  With pBNET(index)
    .InsertNTString vbNullString
    .InsertNTString vbNullString
    .sendPacket &HA, False, index
  End With
End Sub

Public Sub Recv0x0A(index As Integer)
  BNET(index).UniqueName = pBNET(index).getNTString
  pBNET(index).getNTString 'skip statstring
  BNET(index).AccountName = pBNET(index).getNTString
  
  AddChat frmMain.rtbChatBNET, vbGreen, "Bot #" & index & ": Logged in as " & BNET(index).UniqueName
End Sub

Public Sub Send0x0B(index As Integer)
  With pBNET(index)
    .InsertDWORD &H0
    .sendPacket &HB, False, index
  End With
End Sub

Public Sub Send0x0C(index As Integer)
  With pBNET(index)
    .InsertDWORD &H2 'IIf(BNET(index).prodStr = "D2DV", &H5, &H1)
    .InsertNTString channel
    .sendPacket &HC, False, index
  End With
End Sub

Public Sub Recv0x0F(index As Integer)
  Dim user As String, text As String, ID As Long, flags As Long, myChannel As String
  Dim supressEvent As Boolean

  If index <> findFirstAliveBot Then supressEvent = True
  
  With pBNET(index)
    ID = .GetDWORD
    flags = .GetDWORD
    .Skip 16
    user = .getNTString
    text = .getNTString
    Select Case ID
      Case &H2:
                If Not supressEvent Then
                  AddChat frmMain.rtbChatBNET, vbWhite, user, vbYellow, " joined " & myChannel
                  If isBroadcastToIRC Then SendToIRC user & " has joined " & myChannel & "."
                End If
      Case &H3:
                If Not supressEvent Then
                  AddChat frmMain.rtbChatBNET, vbWhite, user, vbYellow, " left " & myChannel
                  If isBroadcastToIRC Then SendToIRC user & " has left " & myChannel & "."
                End If
      Case &H5, &H17:
                If Not supressEvent Then
                  AddChat frmMain.rtbChatBNET, vbWhite, "<" & user & "> ", vbYellow, text
        
                  For i = 0 To UBound(BNET)
                    If Left(LCase(user), Len(BNET(i).AccountName)) = LCase(BNET(i).AccountName) Then
                      Exit Sub
                    End If
                  Next i
                  
                  If isBroadcastToIRC Then
                    'SendToIRC "(" & myChannel & " @ " & BNETServer & ") " & User & ": " & Text
                    SendToIRC IIf(ID = &H17, "/me ", "") & user & ": " & text
                  End If
                End If
      Case &H7: AddChat frmMain.rtbChatBNET, vbYellow, "You joined the channel ", vbWhite, text
                myChannel = text
                If isBroadcastToIRC Then SendToIRC BNET(index).UniqueName & " has joined the Battle.Net channel " & myChannel
    End Select
  End With
End Sub
