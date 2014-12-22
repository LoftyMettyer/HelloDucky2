Attribute VB_Name = "modSysProcesses"
Option Explicit

Public Function CurrentUsersPopulate(grdTemp As SSDBGrid, strUsersToLogOut As String, ByRef WebUserCount As Integer) As Boolean

  Dim rsUsers As New ADODB.Recordset
  Dim sSQL As String

  Dim sProgName As String
  Dim sHostName As String
  Dim sLoginName As String
  Dim sSPID As String

  On Local Error GoTo LocalErr

  WebUserCount = 0
  sSQL = "EXEC spASRGetCurrentUsers"
  rsUsers.Open sSQL, gADOCon, adOpenStatic, adLockReadOnly

  grdTemp.RemoveAll
  Do While Not rsUsers.EOF

    sLoginName = Trim(rsUsers!Loginame)
    'If strUsersToLogOut = vbNullString Or InStr(vbCrLf & strUsersToLogOut & vbCrLf, vbCrLf & sLoginName & vbCrLf) > 0 Then
    If strUsersToLogOut = vbNullString Or InStr(vbCrLf & LCase(strUsersToLogOut) & vbCrLf, vbCrLf & LCase(sLoginName) & vbCrLf) > 0 Then

      sProgName = Trim(rsUsers!program_name)
      sHostName = Trim(rsUsers!HostName)
      sSPID = Trim(rsUsers!Spid)
  
      'Ignore this app on this PC if this login..
      If LCase(sHostName) <> LCase(Trim(UI.GetHostName)) Or _
         LCase(sProgName) <> LCase(Trim(App.ProductName)) Or _
         LCase(sLoginName) <> LCase(Trim(gsUserName)) Then
        grdTemp.AddItem sLoginName & vbTab & sHostName & vbTab & sProgName & vbTab & sSPID
      End If
    End If

    ' Count the web users
    If rsUsers!program_name = "OpenHR Web" Then
      WebUserCount = WebUserCount + 1
    End If
      
    rsUsers.MoveNext
  Loop

  rsUsers.Close
  Set rsUsers = Nothing
  
  CurrentUsersPopulate = True

Exit Function

LocalErr:
  Screen.MousePointer = vbNormal
  MsgBox "Error checking process information" & vbCrLf & "(" & Err.Description & ")", vbExclamation + vbOKOnly, App.Title
  CurrentUsersPopulate = False

End Function


Public Function GetCurrentUsersCountOnServer(strUserName As String) As Long

  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter

  On Error GoTo LocalErr
  
  
  Set cmADO = New ADODB.Command
  With cmADO
    .CommandText = "spASRGetCurrentUsersCountOnServer"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon

    Set pmADO = .CreateParameter("Count", adInteger, adParamOutput)
    .Parameters.Append pmADO

    Set pmADO = .CreateParameter("UserName", adVarChar, adParamInput, VARCHAR_MAX_Size)
    .Parameters.Append pmADO
    pmADO.Value = strUserName

    cmADO.Execute

    GetCurrentUsersCountOnServer = IIf(IsNull(.Parameters(0).Value), 0, .Parameters(0).Value)
  End With
  Set pmADO = Nothing
  Set cmADO = Nothing

Exit Function

LocalErr:
  Screen.MousePointer = vbNormal
  MsgBox "Error checking process information" & vbCr & _
         "(GetCurrentUsersCountOnServer - " & Err.Description & ")", vbCritical

End Function


Public Function GetCurrentUsersCountInApp() As Long

  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter

  On Error GoTo LocalErr
  
  
  Set cmADO = New ADODB.Command
  With cmADO
    .CommandText = "spASRGetCurrentUsersCountInApp"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon

    Set pmADO = .CreateParameter("Count", adInteger, adParamOutput)
    .Parameters.Append pmADO

    cmADO.Execute

    GetCurrentUsersCountInApp = IIf(IsNull(.Parameters(0).Value), 0, .Parameters(0).Value)
  End With
  Set pmADO = Nothing
  Set cmADO = Nothing

Exit Function

LocalErr:
  Screen.MousePointer = vbNormal
  MsgBox "Error checking process information" & vbCr & _
         "(GetCurrentUsersCountInApp - " & Err.Description & ")", vbCritical

End Function


'ONLY USED IN SECURITY MANAGER?
Public Function GetCurrentUsersAppName(strUserName As String) As String

  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter

  On Error GoTo LocalErr
  
  
  Set cmADO = New ADODB.Command
  With cmADO
    .CommandText = "spASRGetCurrentUsersAppName"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon

    Set pmADO = .CreateParameter("AppName", adVarChar, adParamOutput, VARCHAR_MAX_Size)
    .Parameters.Append pmADO

    Set pmADO = .CreateParameter("UserName", adVarChar, adParamInput, VARCHAR_MAX_Size)
    .Parameters.Append pmADO
    pmADO.Value = strUserName

    cmADO.Execute

    GetCurrentUsersAppName = IIf(IsNull(.Parameters(0).Value), "", .Parameters(0).Value)
  End With
  Set pmADO = Nothing
  Set cmADO = Nothing

Exit Function

LocalErr:
  Screen.MousePointer = vbNormal
  MsgBox "Error checking process information" & vbCr & _
         "(GetCurrentUsersAppName - " & Err.Description & ")", vbCritical

End Function
