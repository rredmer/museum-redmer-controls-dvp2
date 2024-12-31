Attribute VB_Name = "WindowsAPI"
Option Explicit

'---- The following calls are in Kernel32.dll
Public Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpszSection$, ByVal lpszKey$, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Public Declare Sub Sleep Lib "kernel32" (ByVal Mills As Long)  'Windows SLEEP call - used for Delay
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long       '0=hidden,-1=show



