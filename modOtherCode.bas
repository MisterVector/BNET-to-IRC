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
    Case "W2BN": getVerByte = &H4F
    Case "D2DV": getVerByte = &HE
    Case "WAR3": getVerByte = &H1C
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


