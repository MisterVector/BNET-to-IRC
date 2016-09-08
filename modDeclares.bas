Attribute VB_Name = "modDeclares"
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal numbytes As Long)

