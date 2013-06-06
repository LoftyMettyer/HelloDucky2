Attribute VB_Name = "modHelp"
Option Explicit

Public Function ShowAirHelp(lngHelpContextID As Long) As Boolean

  Dim strAirHelpFile As String
  Dim strParams As String
  Dim lngHelp As Long
  
  On Local Error GoTo LocalErr

  strAirHelpFile = Environ("programfiles") & "\Advanced\HR Pro Help\" & App.EXEName & " Help\" & App.EXEName & " Help.exe"
  If Dir(strAirHelpFile) = vbNullString Then
    strAirHelpFile = "C:\Program Files (x86)\Advanced\HR Pro Help\" & App.EXEName & " Help\" & App.EXEName & " Help.exe"
    If Dir(strAirHelpFile) = vbNullString Then
      strAirHelpFile = "C:\Program Files\Advanced\HR Pro Help\" & App.EXEName & " Help\" & App.EXEName & " Help.exe"
    End If
  End If
  
  strParams = " -csh mapnumber " & CStr(lngHelpContextID)
  lngHelp = ShellExecute(0&, vbNullString, strAirHelpFile, strParams, vbNullString, vbNormalNoFocus)
  
  ShowAirHelp = (lngHelp = 42)
  
Exit Function

LocalErr:
  ShowAirHelp = False
  
End Function
