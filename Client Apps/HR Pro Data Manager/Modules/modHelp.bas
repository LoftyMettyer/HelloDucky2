Attribute VB_Name = "modHelp"
Option Explicit

Public Function ShowAirHelp(lngHelpContextID As Long) As Boolean

  Dim strAirHelpFile As String
  Dim strParams As String
  Dim lngHelp As Long

  On Local Error GoTo LocalErr

  strAirHelpFile = App.Path & "\HR Pro Data Manager Help\HR Pro Data Manager Help.exe"
  strParams = " -csh mapnumber " & CStr(lngHelpContextID)
  
  lngHelp = ShellExecute(0&, vbNullString, strAirHelpFile, strParams, vbNullString, vbNormalNoFocus)
  
  ShowAirHelp = (lngHelp = 42)
  
Exit Function

LocalErr:
  ShowAirHelp = False
  
End Function
