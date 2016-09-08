Attribute VB_Name = "modDNS"
Public Sub SendDNSQuery(Index As Integer)
Dim tDat As String
tDat = CStr(Index)
Do Until Len(tDat) >= 2
    tDat = "0" & tDat
Loop
tDat = tDat & Chr(1) & Chr & Chr & Chr(1) & Chr & Chr & Chr & Chr & Chr & Chr
aIP = Split(dUser.Item(Index).IP, ".")
For i = 3 To 0 Step -1
    Select Case Len(aIP)
    Case 1: tDat = tDat & Chr(1)
    Case 2: tDat = tDat & Chr(2)
    Case 3: tDat = tDat & Chr(3)
    End Select
    tDat = tDat & aIP
Next
tDat = tDat & Chr(7) & "in-addr" & Chr(4) & "arpa" & Chr & Chr & Chr(&HC) & Chr & Chr(1)
frmServer.dns.SendData tDat
End Sub
