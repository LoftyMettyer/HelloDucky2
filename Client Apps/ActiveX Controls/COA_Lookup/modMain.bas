Attribute VB_Name = "modMain"
Option Explicit
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Public Sub TrySetFocus(ctl As Control)
    On Local Error Resume Next
    ctl.SetFocus
End Sub
