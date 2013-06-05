Attribute VB_Name = "modConsts"
Option Explicit

Public Const VARCHAR_MAX_Size = 2147483646 'Yup one below the actual max, needs to be otherwise things go so awfully wrong, you don't believe me, well go on then, change it, see if I care!!!)

' Window formatting
Public Const LB_SETHORIZONTALEXTENT = &H194
Public Const WS_THICKFRAME As Long = &H40000
Public Const WS_MAXIMIZE As Long = &H1000000
Public Const WS_MAXIMIZEBOX As Long = &H10000
Public Const WS_MINIMIZE As Long = &H20000000
Public Const WS_MINIMIZEBOX As Long = &H20000
Public Const WS_EX_WINDOWEDGE As Long = &H100
Public Const WS_EX_APPWINDOW As Long = &H40000
Public Const WS_EX_DLGMODALFRAME As Long = &H1
Public Const GWL_EXSTYLE As Long = (-20)
Public Const GWL_STYLE As Long = (-16)

' Advanced database settings to control the recursion levels in the database
Public Const giDefaultRecursionLevel = 8



