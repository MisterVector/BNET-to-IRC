Attribute VB_Name = "modFileHandling"
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Sub WriteINI(ByVal cSection As String, cKey As String, ByVal cValue As String, ByVal cPath As String)
  cPath = App.Path & "\" & cPath
  
  WritePrivateProfileString cSection, cKey, cValue, cPath
End Sub

Public Function ReadINI(ByVal cSection As String, cKey As String, cPath As String) As String
  Dim cBuff As String, cLen As Long, tempPath As String
  
  tempPath = App.Path & "\" & cPath
  cBuff = String(255, vbNull)
  cLen = GetPrivateProfileString(cSection, cKey, Chr(0), cBuff, 255, tempPath)
  
  If cLen > 0 Then
    ReadINI = Split(cBuff, Chr$(0))(0)
  Else
    ReadINI = vbNullString
  End If
End Function




