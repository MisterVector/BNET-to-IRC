Attribute VB_Name = "modOtherCode"
Public Function getVerByte(ByVal product As String) As Long
  Select Case product
    Case "STAR": getVerByte = &HD3
    Case "W2BN": getVerByte = &H4F
    Case "D2DV": getVerByte = &HD
    Case "WAR3": getVerByte = &H18
  End Select
End Function

Public Function getProdID(ByVal product As String) As Long
  Select Case product
    Case "STAR": getProdID = &H1
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
End Function

Public Sub AddQ(ByVal msg As String)
  dicQueue.Add dicQueue.Count + 1, msg
  
  If Not frmMain.tmrReleaseQueue.Enabled Then
    dicIdx = 1
    frmMain.tmrReleaseQueue.Enabled = True
  End If
End Sub

Public Sub SendToBNET(ByVal msg As String)
  Dim grabPartOfMessage As String

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
    frmMain.sckIRC.SendData "PRIVMSG " & IRC.Channel & " :" & msg & vbCrLf
  End If
End Sub

Public Sub ConnectOtherBots()
  If BotCount > 1 Then
    For i = 1 To frmMain.sckBNET.Count - 1
      If frmMain.sckBNET(i).State = sckClosed Then
        frmMain.sckBNET(i).Connect frmMain.cmbServer.Text, 6112
      End If
    Next i
  End If
End Sub

