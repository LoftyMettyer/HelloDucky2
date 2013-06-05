Attribute VB_Name = "modHelp"
Option Explicit

Public Function ShowAirHelp(lngHelpContextID As Long) As Boolean

  Dim strAirHelpFile As String
  Dim strParams As String
  Dim lngHelp As Long
  
  On Local Error GoTo LocalErr

  strAirHelpFile = App.Path & "\Help\" & App.EXEName & ".exe"
  If Dir(strAirHelpFile) = vbNullString Then
    strAirHelpFile = "C:\Program Files\COA Solutions\HR Pro v" & CStr(App.Major) & "." & CStr(App.Minor) & "\Help\" & App.EXEName & ".exe"
  End If
  
  strParams = " -csh mapnumber " & CStr(lngHelpContextID)
  lngHelp = ShellExecute(0&, vbNullString, strAirHelpFile, strParams, vbNullString, vbNormalNoFocus)
  
  ShowAirHelp = (lngHelp = 42)
  
Exit Function

LocalErr:
  ShowAirHelp = False
  
End Function
