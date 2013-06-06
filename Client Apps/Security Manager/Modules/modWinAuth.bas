Attribute VB_Name = "modWinAuth"
Option Explicit

Private mastrWindowsGroups() As String

Public Function InitialiseWindowsDomains() As String()
  ReDim mastrWindowsGroups(1, 0)
End Function

Public Function GetDomainFromUser(ByVal strDomain As String, ByVal strUser As String) As String
  If InStr(strUser, "\") > 0 Then
    GetDomainFromUser = Left(strUser, InStr(strUser, "\") - 1)
  Else
    GetDomainFromUser = strDomain
  End If
End Function

Public Function GetDomainFromFQDN(ByVal strDomain As String) As String
  
  On Error GoTo ErrorTrap
  
  Dim rs As New ADODB.Recordset
  Dim strSQL As String

  ' Get the domain from the server
  strSQL = "SELECT [dbo].[udfASRNetGetDomainNameFromFQDN] ('" & strDomain & "')"
  
  Set rs = gADOCon.Execute(strSQL)
  
  GetDomainFromFQDN = CStr(rs.Fields(0).Value)
   
TidyUpAndExit:
  Exit Function

ErrorTrap:
  Resume TidyUpAndExit
End Function

Public Function GetWindowsDomains() As String()

  On Error GoTo ErrorTrap
  
  Dim cmdDomains As New ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim strResult As String
  Dim astrDomainList() As String

  strResult = ""

  ' Get the domain list from the server
  With cmdDomains
    .CommandText = "dbo.[spASRGetDomains]"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon
  
    Set pmADO = .CreateParameter("DomainString", adVarChar, adParamOutput, VARCHAR_MAX_Size)
    .Parameters.Append pmADO
  
    .Execute
  
    strResult = IIf(IsNull(.Parameters(0).Value), "", .Parameters(0).Value)

  End With
  Set cmdDomains = Nothing
  
  If LenB(strResult) <> 0 Then
    strResult = Mid(strResult, 1, Len(strResult) - 1)
  End If

  GetWindowsDomains = Split(strResult, ";")
  
TidyUpAndExit:
  Exit Function

ErrorTrap:
  ReDim mastrWindowsGroups(1, 0)
  GoTo TidyUpAndExit
  
End Function

' Function : GetCurrentUsersInWindowsGroups
' Parameters : pstrGroups - A comma delimited string of windows groups
Public Function GetCurrentUsersInWindowsGroups(pstrGroups As String) As String

  On Error GoTo ErrorTrap
  
  Dim CursorLoc As ADODB.CursorLocationEnum
  Dim cmdUsers As ADODB.Command
  Dim rsUsers As ADODB.Recordset
  Dim pmADO As ADODB.Parameter
  Dim strUserList As String

  strUserList = ""
  
  CursorLoc = gADOCon.CursorLocation
  gADOCon.CursorLocation = adUseClient
  
  Set cmdUsers = New ADODB.Command
  
  ' Get the domain list from the server
  With cmdUsers
    .CommandText = "dbo.[spASRGetCurrentUsersInWindowsGroups]"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon
  
    Set pmADO = .CreateParameter("@psGroupNames", adVarChar, adParamInput, 4000, pstrGroups)
    .Parameters.Append pmADO
  
    Set rsUsers = .Execute()
  End With
  
  With rsUsers
    If Not (.BOF And .EOF) Then
      .MoveFirst
      Do Until .EOF
        strUserList = strUserList & IIf(IsNull(.Fields(0).Value), "", Trim(.Fields(0).Value)) & vbCrLf
        
        .MoveNext
      Loop
    End If
  End With
  Set cmdUsers = Nothing
  
  GetCurrentUsersInWindowsGroups = strUserList
  
TidyUpAndExit:
  gADOCon.CursorLocation = CursorLoc
  Exit Function

ErrorTrap:
  Resume TidyUpAndExit
  
End Function

' Function : GetCurrentUsersInGroups
' Parameters : pstrGroups - A comma delimited string of sec groups
Public Function GetCurrentUsersInGroups(pstrGroups As String) As String

  On Error GoTo ErrorTrap
  
  Dim CursorLoc As ADODB.CursorLocationEnum
  Dim cmdUsers As ADODB.Command
  Dim rsUsers As ADODB.Recordset
  Dim pmADO As ADODB.Parameter
  Dim strUserList As String

  strUserList = ""
  
  CursorLoc = gADOCon.CursorLocation
  gADOCon.CursorLocation = adUseClient
  
  Set cmdUsers = New ADODB.Command
  
  ' Get the domain list from the server
  With cmdUsers
    .CommandText = "dbo.[spASRGetCurrentUsersInGroups]"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon
  
    Set pmADO = .CreateParameter("@psGroupNames", adVarChar, adParamInput, 4000, pstrGroups)
    .Parameters.Append pmADO
  
    Set rsUsers = .Execute()
  End With
  
  With rsUsers
    If Not (.BOF And .EOF) Then
      .MoveFirst
      Do Until .EOF
        strUserList = strUserList & IIf(IsNull(.Fields(0).Value), "", Trim(.Fields(0).Value)) & vbCrLf
        
        .MoveNext
      Loop
    End If
  End With
  Set cmdUsers = Nothing
  
  GetCurrentUsersInGroups = strUserList
  
TidyUpAndExit:
  gADOCon.CursorLocation = CursorLoc
  Exit Function

ErrorTrap:
  Resume TidyUpAndExit
  
End Function

Public Function InitialiseWindowsGroups(strDomain As String) As String()

  Dim sSQL As String
  Dim rsGroups As New ADODB.Recordset
  Dim cmdGroups As New ADODB.Command
  Dim pmADO As ADODB.Parameter
  
  Dim icount As Long

  On Local Error GoTo LocalErr

  ReDim mastrWindowsGroups(1, 0)
  
  If glngSQLVersion = 8 Then
    sSQL = "exec('master..xp_enumgroups ''" & strDomain & "''')"
    
    rsGroups.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  Else
    With cmdGroups
      .CommandText = "[spASRGetWindowsGroupsFromAssembly]"
      .CommandType = adCmdStoredProc
      .CommandTimeout = 0
      Set .ActiveConnection = gADOCon

      Set pmADO = .CreateParameter("@DomainName", adVarChar, adParamInput, 200)
      pmADO.Value = strDomain
      .Parameters.Append pmADO
      
      Set rsGroups = .Execute()

    End With
    Set cmdGroups = Nothing
  End If

  icount = 0
  While Not rsGroups.EOF
    ReDim Preserve mastrWindowsGroups(1, icount)

    mastrWindowsGroups(0, icount) = rsGroups.Fields("Group").Value
    mastrWindowsGroups(1, icount) = IIf(IsNull(rsGroups.Fields("Comment").Value), "", rsGroups.Fields("Comment").Value)

    icount = icount + 1
    rsGroups.MoveNext
  Wend
  rsGroups.Close

ExitAndTidyUp:
  Set rsGroups = Nothing
  EnsureStillConnected

  InitialiseWindowsGroups = mastrWindowsGroups

Exit Function

LocalErr:
  ReDim mastrWindowsGroups(1, 0)
  Resume ExitAndTidyUp

End Function

Public Function InitialiseWindowsUsers(strDomainName As String) As String()

  On Error GoTo ErrorTrap
  
  Dim cmdUsers As New ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim rsUsers As New ADODB.Recordset
  
  Dim strResult As String
  Dim astrUserList() As String

  strResult = ""

  ' Get the domain list from the server
  If glngSQLVersion = 8 Then
    With cmdUsers
      .CommandText = "dbo.[spASRGetWindowsUsers]"
      .CommandType = adCmdStoredProc
      .CommandTimeout = 0
      Set .ActiveConnection = gADOCon
    
      Set pmADO = .CreateParameter("DomainName", adVarChar, adParamInput, 200)
      pmADO.Value = strDomainName
      .Parameters.Append pmADO
    
      Set pmADO = .CreateParameter("DomainString", adVarChar, adParamOutput, VARCHAR_MAX_Size)
      .Parameters.Append pmADO
    
      .Execute
    
      strResult = IIf(IsNull(.Parameters(1).Value), "", .Parameters(1).Value)
  
    End With
    Set cmdUsers = Nothing
    
    If LenB(strResult) <> 0 Then
      strResult = Mid(strResult, 1, Len(strResult) - 1)
    End If
  
    InitialiseWindowsUsers = Split(strResult, ";")
  Else
    ' AE20080312 Fault #12999
    With cmdUsers
      .CommandText = "[spASRGetWindowsUsersFromAssembly]"
      .CommandType = adCmdStoredProc
      .CommandTimeout = 0
      Set .ActiveConnection = gADOCon

      Set pmADO = .CreateParameter("@DomainName", adVarChar, adParamInput, 200)
      pmADO.Value = strDomainName
      .Parameters.Append pmADO
      
      Set rsUsers = .Execute()
    End With
    Set cmdUsers = Nothing
  
    Dim icount As Integer
    icount = 0
    While Not rsUsers.EOF
      ReDim Preserve astrUserList(icount)
  
      astrUserList(icount) = IIf(IsNull(rsUsers.Fields("User").Value), "", rsUsers.Fields("User").Value)
  
      icount = icount + 1
      rsUsers.MoveNext
    Wend
    rsUsers.Close
    Set rsUsers = Nothing
    
    InitialiseWindowsUsers = astrUserList
  End If
  
TidyUpAndExit:
  Exit Function

ErrorTrap:
  MsgBox "Error reading the Domain." & vbNewLine & _
          IIf(glngSQLVersion = 8, "Ensure the OpenHR Server DLL is registered." & vbNewLine, "") & _
          "(" & Err.Number & " - " & Err.Description & ")", vbExclamation + vbOKOnly, App.Title
  ReDim InitialiseWindowsUsers(1, 0)
  GoTo TidyUpAndExit

End Function

Public Function IsWindowsGroup(strGroup As String) As Boolean

  Dim icount As Integer

  IsWindowsGroup = False
  For icount = LBound(mastrWindowsGroups, 2) To UBound(mastrWindowsGroups, 2)
    If LCase(mastrWindowsGroups(0, icount)) = LCase(strGroup) Then
      IsWindowsGroup = True
      Exit For
    End If
  Next

End Function


Public Sub EnsureStillConnected()

  Dim strConnectionString As String

  On Local Error GoTo LocalErr
  gADOCon.Execute "SELECT 'Test Connection'"

Exit Sub
  
LocalErr:
  strConnectionString = gADOCon.ConnectionString
  Set gADOCon = Nothing
  Set gADOCon = New ADODB.Connection
  
  With gADOCon
    .ConnectionString = strConnectionString
    .Provider = "SQLOLEDB"
    .CommandTimeout = 5
    .ConnectionTimeout = 5
    .CursorLocation = adUseServer
    .Mode = adModeReadWrite
    .Properties("Packet Size") = 32767
    .Open
  End With

End Sub

' Does the specified account exist?
Public Function CheckNTAccountExist(pstrUsername As String) As Boolean
  
  On Error GoTo ErrorOccurred
  
  Dim strSQL As String
  Dim rsResult As ADODB.Recordset
  Dim bResult As Boolean

  Set rsResult = New ADODB.Recordset
  strSQL = "EXEC spASRCheckNTLogin '" & Replace(pstrUsername, "'", "''") & "'"
  
  rsResult.Open strSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
  CheckNTAccountExist = rsResult.Fields(0).Value
  rsResult.Close
  Set rsResult = Nothing

  CheckNTAccountExist = True
  Exit Function
  
ErrorOccurred:

  ' Was legal but tripped error (SQL for some reason does this, don't ask me why I'm just the programmer!)
  If Err.Number = 3265 Then
    CheckNTAccountExist = True
  Else
    CheckNTAccountExist = False
  End If

End Function

