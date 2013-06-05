Attribute VB_Name = "modAPI"
Public Declare Function GetComputerNameAPI Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
