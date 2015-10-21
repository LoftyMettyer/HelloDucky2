Attribute VB_Name = "modAPI"
Option Explicit

' Generic API
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" ( _
    ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As String) As Long

Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

' Used for performance monitoring
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long

Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" ( _
    ByVal lpString1 As String, _
    ByVal lpString2 As Any) As Long

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" ( _
    ByVal hInst As Long, _
    ByVal lpsz As String, _
    ByVal uType As Long, _
    ByVal cxDesired As Long, _
    ByVal cyDesired As Long, _
    ByVal fuLoad As Long) As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
  ByVal hWnd As Long, ByVal lpoperation As String, ByVal lpFile As String, _
  ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' File system stuff
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function StringFromGUID2 Lib "OLE32.DLL" (rguid As udtGUID, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long
Public Declare Function StringFromCLSID Lib "OLE32.DLL" (ByRef pCLSID As udtGUID, ByRef pOleStr As Long) As Long
Public Declare Function CLSIDFromString Lib "OLE32.DLL" (ByVal pString As Long, ByRef pCLSID As udtGUID) As Long

Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long
Public Declare Function ClipCursorRect Lib "user32" Alias "ClipCursor" (lpRect As Rect) As Long
Public Declare Function ClipCursorClear Lib "user32" Alias "ClipCursor" (ByVal lpRect As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

