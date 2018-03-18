Attribute VB_Name = "modDeclares"
Public Const IDC_HAND As Long = 32649&

Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal numbytes As Long)
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
