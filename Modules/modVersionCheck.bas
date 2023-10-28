Attribute VB_Name = "modVersionCheck"
Public Declare Function check_revision Lib "VersionCheck.dll" (ByVal ArchiveTime As String, ByVal ArchiveName As String, ByVal Seed As String, ByVal INIFile As String, ByVal INIHeader As String, ByRef Version As Long, ByRef Checksum As Long, ByVal result As String) As Long
Public Declare Function crev_error_description Lib "VersionCheck.dll" (ByVal error As Long, ByVal errorText As String, ByVal size As Long) As Long
Public Declare Function crev_max_result Lib "VersionCheck.dll" () As Long

Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Public Type FILETIME
    dwLowDateTime       As Long
    dwHighDateTime      As Long
End Type

Public Type SYSTEMTIME
    wYear               As Integer
    wMonth              As Integer
    wDayOfWeek          As Integer
    wDay                As Integer
    wHour               As Integer
    wMinute             As Integer
    wSecond             As Integer
    wMilliseconds       As Integer
End Type
Public tpLocal As SYSTEMTIME
Public tpSystem As SYSTEMTIME

Public Function GetFTTime(FT As FILETIME, Optional Shorten As Boolean = False, Optional localTime As Boolean = True) As String
    Dim LocalFT As FILETIME
    Dim SysTime As SYSTEMTIME
    Dim SetHour As String
    Dim AP      As String

    If (localTime) Then
        FileTimeToLocalFileTime FT, LocalFT
        FileTimeToSystemTime LocalFT, SysTime
    Else
        FileTimeToSystemTime FT, SysTime
    End If
  
    If (SysTime.wHour = 0) Then
        AP = "AM"
        SetHour = "12"
    ElseIf (SysTime.wHour < 12) Then
        AP = "AM"
        SetHour = Trim$(str$(SysTime.wHour))
    ElseIf (SysTime.wHour = 12) Then
        AP = "PM"
        SetHour = "12"
    Else
        AP = "PM"
        SetHour = Trim$(str$(SysTime.wHour))
    End If
  
    SysTime.wDayOfWeek = SysTime.wDayOfWeek + 1
  
    If (Shorten) Then
        GetFTTime = Format$(SysTime.wMonth, "00") & "/" & Format$(SysTime.wDay, "00") & "/" & Right$(SysTime.wYear, 2) & " " & SetHour & ":" & Format$(SysTime.wMinute, "00") & ":" & Format$(SysTime.wSecond, "00") & " " & AP
    Else
        GetFTTime = ConvertShortToLong(WeekdayName(SysTime.wDayOfWeek, True)) & ", " & MonthName(SysTime.wMonth, True) & " " & SysTime.wDay & ", " & SysTime.wYear & " at " & SetHour & ":" & Format$(SysTime.wMinute, "00") & ":" & Format$(SysTime.wSecond, "00") & " " & AP
    End If
End Function

Private Function ConvertShortToLong(Day As String)
    Select Case Day
        Case "Mon": ConvertShortToLong = "Monday"
        Case "Tue": ConvertShortToLong = "Tuesday"
        Case "Wed": ConvertShortToLong = "Wednesday"
        Case "Thu": ConvertShortToLong = "Thursday"
        Case "Fri": ConvertShortToLong = "Friday"
        Case "Sat": ConvertShortToLong = "Saturday"
        Case "Sun": ConvertShortToLong = "Sunday"
    End Select
End Function



