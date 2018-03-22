Attribute VB_Name = "modBNLSChecksum"
Option Explicit

Private Const CRC32_POLYNOMIAL As Long = &HEDB88320
Private CRC32Table(0 To 255) As Long

Private Sub InitCRC32()
    Dim i As Long, j As Long, k As Long, XorVal As Long
    
    Static CRC32Initialized As Boolean
    If CRC32Initialized Then Exit Sub
    CRC32Initialized = True
    
    For i = 0 To 255
        k = i
        
        For j = 1 To 8
            If k And 1 Then XorVal = CRC32_POLYNOMIAL Else XorVal = 0
            If k < 0 Then k = ((k And &H7FFFFFFF) \ 2) Or &H40000000 Else k = k \ 2
            k = k Xor XorVal
        Next
        
        CRC32Table(i) = k
    Next
End Sub

Private Function CRC32(ByVal data As String) As Long
    Dim i As Long, j As Long
    
    Call InitCRC32
    
    CRC32 = &HFFFFFFFF
    
    For i = 1 To Len(data)
        j = CByte(Asc(Mid$(data, i, 1))) Xor (CRC32 And &HFF&)
        If CRC32 < 0 Then CRC32 = ((CRC32 And &H7FFFFFFF) \ &H100&) Or &H800000 Else CRC32 = CRC32 \ &H100&
        CRC32 = CRC32 Xor CRC32Table(j)
    Next
    
    CRC32 = Not CRC32
End Function

Public Function BNLSChecksum(ByVal Password As String, ByVal ServerCode As Long) As Long
    BNLSChecksum = CRC32(Password & Right("0000000" & Hex(ServerCode), 8))
End Function
