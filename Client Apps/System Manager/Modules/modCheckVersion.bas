Attribute VB_Name = "modCheckVersion"
Option Explicit

Dim mstrLastSQLServerVersion As String
Dim mstrLastDatabaseName As String
Dim mstrLastServerName As String
Dim mstrCurrentDatabaseName As String
Dim mstrCurrentServerName As String  ' Using SERVERPROPERTY('servername')
Dim mstrOldServerName As String      ' Using @@SERVERNAME

Public gfDatabaseServerChanged As Boolean
Public gfWFCredentialsChanged As Boolean

Public Function CheckVersion(sConnect As String, fReRunScript As Boolean, bIsSQLSystemAdmin As Boolean) As Boolean
  ' Check that the database version is the right one for this application's version.
  ' If everything matches then return TRUE.
  ' If not, try to update the database.
  ' If the database can be updated return TRUE, else return FALSE.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fVersionOK As Boolean
  Dim iPointer As Integer
  Dim iMajorAppVersion As Integer
  Dim iMinorAppVersion As Integer
  Dim iRevisionAppVersion As Integer
  Dim sDBVersion As String
  Dim blnNewStyleVersionNo As Boolean
  Dim sMinVersion As String
  Dim iMinimumMajor As Integer
  Dim iMinimumMinor As Integer
  
  Dim blnReRunCurrent As Boolean
  
  Dim sDependencyVersion As String
  Dim iMajorDependencyVersion As Integer
  Dim iMinorDependencyVersion As Integer
  Dim bLicenceKeyRequired As Boolean
  
  Dim idxname As Integer
  Dim rsInfo As New ADODB.Recordset
  Dim strSQLVersion As String
  
  
  fOK = True
  fVersionOK = False
  gfRefreshStoredProcedures = False
  gfDatabaseServerChanged = False
  gfWFCredentialsChanged = False
    
  blnReRunCurrent = False
  
  If fOK Then
    sDBVersion = GetDBVersion
    
    If Len(sDBVersion) = 0 Then
      fOK = False
      
      MsgBox "Error checking version compatibility." & vbNewLine & _
        "Version number not found.", _
        vbOKOnly + vbExclamation, Application.Name
    Else
      iMajorAppVersion = val(Split(sDBVersion, ".")(0))
      iMinorAppVersion = val(Split(sDBVersion, ".")(1))
      
      blnNewStyleVersionNo = (UBound(Split(sDBVersion, ".")) = 1)
      If Not blnNewStyleVersionNo Then
        iRevisionAppVersion = val(Split(sDBVersion, ".")(2))
      End If
    End If
  End If


  If fOK Then
    ' Check the System Manager version against the one for the current database.
    If (App.Major = iMajorAppVersion) And _
      (App.Minor = iMinorAppVersion) And _
      (App.Revision = iRevisionAppVersion Or blnNewStyleVersionNo) Then
      ' Application and database versions match.
      fVersionOK = True
    End If
  End If
  
  If fOK Then
    ' Check the System Manager version against the one for the current database.
    ' Application is too old for the database.
    If (App.Major < iMajorAppVersion) Or _
      ((App.Major = iMajorAppVersion) And (App.Minor < iMinorAppVersion)) Or _
      ((App.Major = iMajorAppVersion) And (App.Minor = iMinorAppVersion) And (App.Revision < iRevisionAppVersion And Not blnNewStyleVersionNo)) Then
      fOK = False
      MsgBox "The application is out of date." & vbNewLine & _
        "Contact your System Administrator for a new version of the application." & vbNewLine & vbNewLine & _
        "Database Name : " & gsDatabaseName & vbNewLine & _
        "Database Version : " & sDBVersion & vbNewLine & vbNewLine & _
        "Application Version : " & CStr(App.Major) & "." & CStr(App.Minor), _
        vbExclamation + vbOKOnly, Application.Name
    End If
    'NHRD01082011 JIRA HRPro 1501
    'Check to see if app is 4.2 and if it is make sure there are no tables starting with a reserved word prefix
    If CStr(App.Major) & "." & CStr(App.Minor) = "4.2" Then
      strSQLVersion = "SELECT name FROM dbo.sysobjects where name like 'tbstat%' or name like 'tbsys%' or name like 'tbuser%'"
      rsInfo.Open strSQLVersion, gADOCon, adOpenStatic, adLockReadOnly, adCmdText
      With rsInfo
        If Not (.BOF And .EOF) Then
          .MoveLast
          .MoveFirst
          idxname = .RecordCount
          MsgBox "There are " & idxname & " tables starting with the reserved word prefix of either 'tbstat', 'tbuser' or 'tbsys'" & _
            vbCrLf & vbCrLf & "Please rename these tables before upgrading to the latest version." & _
            vbCrLf & vbCrLf & "The System Manager will now terminate the Upgrade Script.", vbCritical, Application.Name
        End If
      End With
      
      rsInfo.Close
    End If
    
  End If
     
  ' AE20080218 Fault #12834, 12859
  If fOK Then

    Dim frmChangedPlatform As frmChangedPlatform
    Dim mavValidationMessages() As String
    
    Set frmChangedPlatform = New frmChangedPlatform
    frmChangedPlatform.ResetList
    frmChangedPlatform.Width = frmChangedPlatform.lblUsageMSG.Width

    
    ' 0 = Message
    ' 1 = Old Value
    ' 2 = New Value
    ReDim mavValidationMessages(3, 0)
    
    ' Database is too old for the application. Try to update the database.
    If (App.Major > iMajorAppVersion) Or _
      ((App.Major = iMajorAppVersion) And (App.Minor > iMinorAppVersion)) Or _
      ((App.Major = iMajorAppVersion) And (App.Minor = iMinorAppVersion) And (App.Revision > iRevisionAppVersion And Not blnNewStyleVersionNo)) Then
      
      If bIsSQLSystemAdmin Then
        ReDim Preserve mavValidationMessages(3, UBound(mavValidationMessages, 2) + 1)
        mavValidationMessages(0, UBound(mavValidationMessages, 2)) = "Database is out of date"
        mavValidationMessages(1, UBound(mavValidationMessages, 2)) = sDBVersion
        mavValidationMessages(2, UBound(mavValidationMessages, 2)) = CStr(App.Major) & "." & CStr(App.Minor)
        mavValidationMessages(3, UBound(mavValidationMessages, 2)) = "The database is out of date"
        
        ReDim Preserve mavValidationMessages(3, UBound(mavValidationMessages, 2) + 1)
        mavValidationMessages(0, UBound(mavValidationMessages, 2)) = "Licence key must be entered"
        mavValidationMessages(1, UBound(mavValidationMessages, 2)) = ""
        mavValidationMessages(2, UBound(mavValidationMessages, 2)) = ""
        mavValidationMessages(3, UBound(mavValidationMessages, 2)) = "Licence key must be entered"
        
        bLicenceKeyRequired = True
        fVersionOK = True
      Else
        fVersionOK = False
        MsgBox "The database is out of date." & vbNewLine & _
          "A System Administrator must log into the System Manager to update the database." & vbNewLine & vbNewLine & _
          "Database Name : " & gsDatabaseName & vbNewLine & _
          "Database Version : " & sDBVersion & vbNewLine & vbNewLine & _
          "Application Version : " & CStr(App.Major) & "." & CStr(App.Minor), _
          vbExclamation + vbOKOnly, Application.Name
      End If
      fOK = fVersionOK
    End If
    
    ' AE20090211 Fault #13550
    If fOK Then
      mstrCurrentServerName = GetServerName()
      mstrOldServerName = GetOldServerName()

      If mstrOldServerName <> mstrCurrentServerName Then
        fOK = False
        blnReRunCurrent = True
        
        MsgBox "The Microsoft SQL Server has been renamed but the operation is incomplete." & vbNewLine & vbNewLine & _
          "Old Server Name : " & mstrOldServerName & vbNewLine & _
          "New Server Name : " & mstrCurrentServerName & vbNewLine & vbNewLine & _
          "Please contact your System Administrator before logging in to System Manager.", _
            vbExclamation + vbOKOnly, Application.Name
      End If
    End If
    
    If fOK Then
      mstrLastSQLServerVersion = GetSystemSetting("Platform", "SQLServerVersion", 0)
      mstrLastDatabaseName = UCase$(GetSystemSetting("Platform", "DatabaseName", ""))
      mstrLastServerName = UCase$(GetSystemSetting("Platform", "ServerName", ""))
      If mstrLastServerName = "." Then mstrLastServerName = UCase$(UI.GetHostName)
      mstrCurrentDatabaseName = GetDBName()
      
      ' AE20090128 #Fault 13514
      'If Val(mstrLastSQLServerVersion) <> glngSQLVersion Then
      ' AE20090623 Fault #13661
      '    - Reversal of 13514 as now fixed in SEC
      'If mstrLastSQLServerVersion <> gstrSQLFullVersion Then
      If val(mstrLastSQLServerVersion) <> glngSQLVersion Then
        ReDim Preserve mavValidationMessages(3, UBound(mavValidationMessages, 2) + 1)
        mavValidationMessages(0, UBound(mavValidationMessages, 2)) = "Microsoft SQL Version Upgraded"
        mavValidationMessages(1, UBound(mavValidationMessages, 2)) = mstrLastSQLServerVersion
        ' AE20080311 Fault #13001
        'mavValidationMessages(2, UBound(mavValidationMessages, 2)) = glngSQLVersion
        mavValidationMessages(2, UBound(mavValidationMessages, 2)) = gstrSQLFullVersion
        mavValidationMessages(3, UBound(mavValidationMessages, 2)) = "The Microsoft SQL Version has been upgraded."
        blnReRunCurrent = True
      End If
      
      If mstrLastServerName <> mstrCurrentServerName Then
        gfDatabaseServerChanged = True
        ReDim Preserve mavValidationMessages(3, UBound(mavValidationMessages, 2) + 1)
        mavValidationMessages(0, UBound(mavValidationMessages, 2)) = "Database moved to different Microsoft SQL Server"
        mavValidationMessages(1, UBound(mavValidationMessages, 2)) = mstrLastServerName
        mavValidationMessages(2, UBound(mavValidationMessages, 2)) = mstrCurrentServerName
        mavValidationMessages(3, UBound(mavValidationMessages, 2)) = "The database has moved to a different Microsoft SQL Server."
        blnReRunCurrent = True
      End If
      
      If mstrLastDatabaseName <> mstrCurrentDatabaseName Then
        gfDatabaseServerChanged = True
        ReDim Preserve mavValidationMessages(3, UBound(mavValidationMessages, 2) + 1)
        mavValidationMessages(0, UBound(mavValidationMessages, 2)) = "Database name has changed"
        mavValidationMessages(1, UBound(mavValidationMessages, 2)) = mstrLastDatabaseName
        mavValidationMessages(2, UBound(mavValidationMessages, 2)) = mstrCurrentDatabaseName
        mavValidationMessages(3, UBound(mavValidationMessages, 2)) = "The database name has changed."
        blnReRunCurrent = True
      End If
      
    End If
    
    If fOK Then
      Dim i As Integer
      If UBound(mavValidationMessages, 2) > 0 And bIsSQLSystemAdmin Then
        For i = 1 To UBound(mavValidationMessages, 2)
          frmChangedPlatform.AddToList CStr(mavValidationMessages(0, i)), _
                              CStr(mavValidationMessages(1, i)), _
                              CStr(mavValidationMessages(2, i))
        Next i
        
        ' AE20080219 Fault #12902
        iPointer = Screen.MousePointer
        Screen.MousePointer = vbDefault
        
        frmChangedPlatform.LicenceKeyRequired = bLicenceKeyRequired
        frmChangedPlatform.ShowMessage
        
        Screen.MousePointer = iPointer
        
        If frmChangedPlatform.Choice = vbYes Then
               
          If bLicenceKeyRequired Then
            SaveSystemSetting "Licence", "Key", frmChangedPlatform.LicenceKey
            gobjLicence.ValidateCreationDate = False
            gobjLicence.LicenceKey = frmChangedPlatform.LicenceKey
          End If

          ' AE20080415 Fault #13098
          'fOK = UpdateDatabase(sConnect, False)
          'fOK = UpdateDatabase(sConnect, True, True)
          fOK = UpdateDatabase(sConnect, blnReRunCurrent, True)

        ElseIf ASRDEVELOPMENT Then
          fOK = True
        Else
          fOK = False
        End If
      
      ' AE20080415 Fault #13099
      ElseIf UBound(mavValidationMessages, 2) > 0 Then
        MsgBox mavValidationMessages(3, 1) & vbCrLf & _
          "Please ask the System Administrator to update the database in the System Manager.", _
          vbOKOnly + vbExclamation, Application.Name
        
        fOK = False
      End If
    End If
    
    UnLoad frmChangedPlatform
    Set frmChangedPlatform = Nothing
  
  End If

  If (fReRunScript Or gblnAutomaticScript) And bIsSQLSystemAdmin Then
    fVersionOK = UpdateDatabase(sConnect, fReRunScript)
  End If
  
  If fOK Then
    ' Check if a new version of the application is required due to an Intranet update
    
    sDBVersion = GetSystemSetting("Database", "Minimum Version", vbNullString)
    If Len(sDBVersion) <> 0 Then
      
      iMajorAppVersion = val(Split(sDBVersion, ".")(0))
      iMinorAppVersion = val(Split(sDBVersion, ".")(1))
      
      blnNewStyleVersionNo = (UBound(Split(sDBVersion, ".")) = 1)
      If Not blnNewStyleVersionNo Then
        iRevisionAppVersion = val(Split(sDBVersion, ".")(2))
      End If
      
      If (App.Major < iMajorAppVersion) Or _
        ((App.Major = iMajorAppVersion) And (App.Minor < iMinorAppVersion)) Or _
        ((App.Major = iMajorAppVersion) And (App.Minor = iMinorAppVersion) And (App.Revision < iRevisionAppVersion And Not blnNewStyleVersionNo)) Then

        fVersionOK = False
        MsgBox "The application is now out of date due to an update to the intranet module." & vbNewLine & _
          "Please contact your System Administrator for a new version of the application.", _
          vbOKOnly + vbExclamation, Application.Name
        fOK = fVersionOK

      End If
    End If
  End If
        
  ' Do we enable UDF functions on this installation
  gbEnableUDFFunctions = EnableUDFFunctions

  If fOK Then
    gfRefreshStoredProcedures = (GetSystemSetting("Database", "RefreshStoredProcedures", 0) = 1)
    Application.Changed = Application.Changed Or gfRefreshStoredProcedures
  End If
  
  'Check to see if the engine is the correct version
  If fOK Then
    fOK = CheckFrameworkVersion()
  End If
    
  ' Upload scripts
  If fOK Then
    fOK = UploadHotfixes
  End If
  
  ' If fOK and fVersionOK are true then the application and databases versions match.
TidyUpAndExit:
  If Not fOK Then
    fVersionOK = False
    Screen.MousePointer = vbDefault
  End If
  
  CheckVersion = fVersionOK
  Exit Function
  
ErrorTrap:
  If (Err.Number = 75) Or (Err.Number = 76) Then
    MsgBox "The database is out of date." & vbNewLine & _
      "Unable to update the database as the required update script cannot be found.", _
      vbOKOnly + vbExclamation, Application.Name
  Else
    MsgBox "Error checking database and application versions." & vbNewLine & _
      ODBC.FormatError(Err.Description), _
      vbOKOnly + vbExclamation, Application.Name
  End If
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function CheckFrameworkVersion() As Boolean

  On Error GoTo ErrorTrap
  
  Dim sActualVersion As String
  Dim sRequiredVersion As String
  Dim bOK As Boolean
  Dim rsInfo As New ADODB.Recordset
   
  bOK = True
  
  sRequiredVersion = GetSystemSetting("system framework", "version", vbNullString)
  sActualVersion = gobjHRProEngine.Version

'  If sRequiredVersion <> sActualVersion And Not ASRDEVELOPMENT Then
'    MsgBox "The System Framework is invalid." & vbNewLine & _
'      "Contact your System Administrator to install the latest System Framework" & vbNewLine & vbNewLine & _
'      "Required Version : " & sRequiredVersion & vbNewLine & _
'      "Actual Version : " & sActualVersion & vbNewLine & vbNewLine _
'      , vbExclamation + vbOKOnly, Application.Name
'    bOK = False
'  End If

TidyUpAndExit:
  CheckFrameworkVersion = bOK
  Exit Function
ErrorTrap:
  MsgBox "The System Framework is not installed." & vbNewLine & _
    "Contact your System Administrator to install the latest System Framework" & vbNewLine & vbNewLine & _
    "Required Version : " & sRequiredVersion & vbNewLine _
    , vbExclamation + vbOKOnly, Application.Name
  bOK = False
  Resume TidyUpAndExit

End Function

Private Function GetServerDLLVersion(sConnect As String) As String

  Dim cmdInfo As ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim rsExists As New ADODB.Recordset
  Dim sSQL As String

  sSQL = "SELECT COUNT(*) AS recCount FROM sysobjects " & _
         "WHERE name = 'spASRGetServerDLLVersion'"
  rsExists.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly

  If rsExists.Fields(0).value = 0 Then
    RunScript gsApplicationPath & "\Update Scripts\Lock.sql", sConnect
  End If

  rsExists.Close
  Set rsExists = Nothing

    
  Set cmdInfo = New ADODB.Command
  With cmdInfo
    .CommandText = "dbo.spASRGetServerDLLVersion"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 5
    Set .ActiveConnection = gADOCon

    Set pmADO = .CreateParameter("Version", adVarChar, adParamOutput, 255)
    .Parameters.Append pmADO

    .Execute

    GetServerDLLVersion = IIf(IsNull(.Parameters(0).value), vbNullString, .Parameters(0).value)
  End With
  Set cmdInfo = Nothing

End Function

'Private Function GetServerAssemblyVersion() As String
'
'  Dim cmdInfo As ADODB.Command
'  Dim pmADO As ADODB.Parameter
'  Dim rsAssembly As New ADODB.Recordset
'  Dim sSQL As String
'
'  sSQL = "SELECT COUNT(*) AS recCount FROM sysobjects " & _
'         "WHERE name = 'udfASRAssemblyVersion'"
'  rsAssembly.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
'
'  If rsAssembly.Fields(0).Value = 0 Then
'    GetServerAssemblyVersion = "0.0"
'    Exit Function
'  End If
'
'  rsAssembly.Close
'
'  sSQL = "SELECT dbo.udfASRAssemblyVersion()"
'  rsAssembly.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
'
'  GetServerAssemblyVersion = rsAssembly.Fields(0).Value
'
'  rsAssembly.Close
'  Set rsAssembly = Nothing
'
'End Function


Private Function GetDBVersion() As String

  Dim rsInfo As New ADODB.Recordset
  
  GetDBVersion = GetSystemSetting("Database", "Version", vbNullString)

  If GetDBVersion = vbNullString Then
    rsInfo.Open "SELECT SystemManagerVersion FROM ASRSysConfig", gADOCon, adOpenForwardOnly, adLockReadOnly
  
    If Not rsInfo.BOF And Not rsInfo.EOF Then
      GetDBVersion = rsInfo.Fields(0).value
    End If
  
    rsInfo.Close
    Set rsInfo = Nothing
  
  End If

End Function

Private Function UpdateDatabase( _
    sConnect As String, fReRunScript As Boolean, Optional fPlatform As Boolean) As Boolean

  Dim rsInfo As New ADODB.Recordset
  Dim lngDBVersion As Long
  Dim strVersion As String
  Dim intMajor As Integer
  Dim intMinor As Integer

  Dim fso As FileSystemObject
  Dim strScriptPath As String
  Dim strFileName As String
  Dim lngNewVersions As Long
  Dim blnSendMessageVisible As Boolean
  Dim bRegenerateProc As Boolean
  Dim strSQLVersion As String

  Dim strMBText As String
  Dim intMBButtons As Integer
  Dim strMBTitle As String
  Dim bOK As Boolean
  
  On Local Error GoTo LocalErr
 
  bOK = True
  
  'MH20070615 Fault 12336 - Do not upgrade SQL 7 databases to v3.5 or later.
  'TM20081126 - Do not upgrade SQL 2000 databases to v3.7 or later.
  If (glngSQLVersion < 9) Then
    Screen.MousePointer = vbDefault
    MsgBox "This version of OpenHR is only compatible with SQL Server 2005 or later. Please upgrade SQL Server before upgrading to this version of OpenHR.", vbCritical
    UpdateDatabase = False
    Exit Function
  End If

  ' AE20080218 Fault #12834
  If fReRunScript And Not fPlatform And Not gblnAutomaticScript Then
    strMBText = "Are you sure that you would like to re-run the latest update script?"
'  Else
'    strMBText = "The database is out of date.  Would you like to update it now?"
'  End If
    
    intMBButtons = vbYesNo + vbQuestion
    strMBTitle = "Update Database"
    
    If MsgBox(strMBText, intMBButtons, strMBTitle) = vbNo Then
      Screen.MousePointer = vbDefault
      UpdateDatabase = False
      Exit Function
    End If
  End If

  frmLogin.Hide

  Dim iUpdates As Integer
  ' AE20090119 Fault #13495
  'iUpdates = IIf(glngSQLVersion > 8, 10, 6)
  iUpdates = IIf(glngSQLVersion > 8, 11, 7)
  
  ' NPG20090828 Fault HRPRO-202
  gobjProgress.AVI = dbTransfer
  gobjProgress.NumberOfBars = 1
  gobjProgress.Caption = App.ProductName
  gobjProgress.Bar1Caption = "Updating Database & Software..."
  gobjProgress.Bar1MaxValue = iUpdates
  gobjProgress.Cancel = False
  gobjProgress.OpenProgress
  
  ' Need to ensure that we have the correct compatability for scripts to run successfully.
  SetDatabaseCompatability
  gobjProgress.UpdateProgress False

  'MH20010903 We don't know how old the database is so make
  'sure that the lock stuff in in there before we start...
  strScriptPath = gsApplicationPath & "\Update Scripts\"
  strFileName = "Lock.sql"
  RunScript strScriptPath & strFileName, sConnect
  gobjProgress.UpdateProgress False

  'MH20020430 From v1.32.4 Make sure that we run the script
  'for the overnight job rather than during the save process
  strFileName = "OvernightJob.sql"
  RunScript strScriptPath & strFileName, sConnect
  gobjProgress.UpdateProgress False
  
  strSQLVersion = "IF EXISTS(SELECT sysobjects.id FROM sysobjects WHERE name = 'ASRSysConfig')" & vbNewLine & _
                  "  IF EXISTS(SELECT databaseVersion FROM ASRSysConfig)" & vbNewLine & _
                  "    SELECT databaseVersion FROM ASRSysConfig" & vbNewLine & _
                  "  ELSE" & vbNewLine & _
                  "    SELECT '27'" & vbNewLine & _
                  "ELSE" & vbNewLine & _
                  "  SELECT '27'"
  rsInfo.Open strSQLVersion, gADOCon, adOpenForwardOnly, adLockReadOnly
  
  blnSendMessageVisible = (rsInfo.Fields(0).value >= 27)
  rsInfo.Close

  ' AE20080421 Fault #13112
  Call CreateUDF_SQLVersion
  gobjProgress.UpdateProgress False
  
  Call CreateSP_GetCurrentUsers
  gobjProgress.UpdateProgress False
      
  Call CreateSP_GetCurrentUsersFromMaster
  gobjProgress.UpdateProgress False
      
  If Not SaveChanges_LogoutCheck(blnSendMessageVisible) Then
    UnlockDatabase lckSaving

    DoEvents  'Prevents runtime error in EXE (beleive it or not!)

    MsgBox "Database update cancelled.", vbInformation, "Update Database"
    UpdateDatabase = False
    Exit Function
  End If
  
  Call CreateSP_LockCheck
  gobjProgress.UpdateProgress False
  
  ' Deploy the .NET assemblies
  If glngSQLVersion > 8 Then
     
    bOK = MakeServerReadyForAssembly
    gobjProgress.UpdateProgress False
    
    If (Not bOK) Then
      ' AE20090119 Fault #13498
      MsgBox "Unable to make the server ready for deploying the assembly" & vbNewLine _
        & "Please contact your System Administrator.", vbExclamation
      UpdateDatabase = False
      Exit Function
    End If
    
'    ' Generate drop script
'    If bOK Then
'      bOK = GenerateDropServerAssemblySP(True)
'      gobjProgress.UpdateProgress False
'    End If
       
    If bOK Then
      'NHRD21112011 renamed from HRProAssembly.sql
      strFileName = "Assembly.sql"
      RunScript strScriptPath & strFileName, sConnect
      gobjProgress.UpdateProgress False
    Else
      gobjProgress.CloseProgress
      MsgBox "Error deploying System Framework Assembly" & vbNewLine & "Please see your System Administrator", vbExclamation
      UpdateDatabase = False
      GoTo TidyUpAndExit
    End If
    
    ' AE20080421 Fault #13112
'    If bOK Then
'      bOK = CreateSP_GetCurrentUsers
'      gobjProgress.UpdateProgress False
'    End If
'
'    If bOK Then
'      bOK = CreateSP_GetCurrentUsersFromMaster
'      gobjProgress.UpdateProgress False
'    End If
    
    If bOK Then
      bOK = RegenerateSQLProcessAccount
      gobjProgress.UpdateProgress False
    End If
    
    ' AE20080229 Fault #12968
    If bOK Then
      bOK = TestSystemLogon
    
      If Not bOK Then
        bOK = RegenerateSQLProcessAccount(gsUserName, gsPassword)
        bRegenerateProc = True
        gobjProgress.UpdateProgress False
      End If
    End If

    If bOK Then
      bOK = RemoveWorkflowLoginCredentials()
      gobjProgress.UpdateProgress False
    End If

  End If
  gobjProgress.CloseProgress
  
  ' Calculate how many files we're going to process
  Set fso = New FileSystemObject
  lngNewVersions = fso.GetFolder(strScriptPath).Files.Count + 1 '- 14
    
  rsInfo.Open strSQLVersion, gADOCon, adOpenForwardOnly, adLockReadOnly
  lngDBVersion = CLng(rsInfo.Fields(0).value) + 1
  rsInfo.Close
  
  Dim fileCount As Integer
        
  strVersion = GetDBVersion
  
  intMajor = CInt(Split(strVersion, ".")(0))
  intMinor = CInt(Split(strVersion, ".")(1))
  If intMajor = 1 And intMinor = 1 Then
    intMinor = CInt(Split(strVersion, ".")(2))
  End If
  
  Dim nextFile As String
  nextFile = Dir$(strScriptPath & "Update-*.Sql")

  Dim iFileMajor As Integer
  Dim iFileMinor As Integer
  Dim ifileVersion As Integer
  
  fileCount = 0
  Do Until nextFile = ""
    iFileMajor = Split(nextFile, "-")(1)
    iFileMinor = Split(Split(nextFile, "-")(2), ".")(0)
    
    If iFileMajor > intMajor Then
      fileCount = fileCount + 1
    ElseIf iFileMajor = intMajor And iFileMinor >= intMinor Then
      fileCount = fileCount + 1
    End If
    
    nextFile = Dir$()
  Loop

  If lngDBVersion <= 27 Then fileCount = fileCount + 14

  gobjProgress.NumberOfBars = 1
  gobjProgress.Caption = "Updating Database..."
  gobjProgress.Bar1Caption = "Checking versions"
  gobjProgress.Bar1MaxValue = fileCount
  gobjProgress.Time = False
  gobjProgress.Cancel = False
  gobjProgress.OpenProgress
  
  'First try and get the database up to v1.1.25...
  lngDBVersion = 0
  Do
    rsInfo.Open strSQLVersion, gADOCon, adOpenForwardOnly, adLockReadOnly

    If (rsInfo.Fields(0).value) + 1 = lngDBVersion Then
      'Error has occurred as the database version has not changed!
      Exit Function
    End If

    lngDBVersion = CLng(rsInfo.Fields(0).value) + 1
    
    If lngDBVersion <= 27 Then
      strFileName = "ASRUpdate_" & CStr(lngDBVersion) & ".sql"
      gobjProgress.Bar1Caption = strFileName
      RunScript strScriptPath & strFileName, sConnect

      gobjProgress.UpdateProgress False
    End If

    rsInfo.Close
  
  Loop While lngDBVersion < 27 And Not gobjProgress.Cancelled

  ' Should now be on v1.1.25 so carry on with the new scripts way...
  If Not fReRunScript Then gobjProgress.UpdateProgress
  Do
    'strVersion = GetSystemSetting("Database", "Version", vbnullstring)
    strVersion = GetDBVersion

    intMajor = CInt(Split(strVersion, ".")(0))
    intMinor = CInt(Split(strVersion, ".")(1))
    If intMajor = 1 And intMinor = 1 Then
      intMinor = CInt(Split(strVersion, ".")(2))
    End If
    
    If intMajor < App.Major Or intMinor < App.Minor Or fReRunScript Then
    
      If fReRunScript Then
        strFileName = "Update-" & CStr(intMajor) & "-" & CStr(intMinor) & ".sql"
      Else
        strFileName = "Update-" & CStr(intMajor) & "-" & CStr(intMinor + 1) & ".sql"
      End If
      
      If Dir(strScriptPath & strFileName) = vbNullString Then
        If intMajor < App.Major Then
          intMinor = 0
          strFileName = "Update-" & CStr(intMajor + 1) & "-" & CStr(intMinor) & ".sql"
        End If
      End If
  
      ' Handle missing versions of 6 and 7
      If fso.FileExists(strScriptPath & strFileName) Then
        gobjProgress.Bar1Caption = strFileName
        RunScript strScriptPath & strFileName, sConnect
        gobjProgress.UpdateProgress False
        
        fReRunScript = False
        
        gobjHRProEngine.Options.VersionUpgraded = True
      End If
      
    End If

  Loop While intMajor < App.Major Or intMinor < App.Minor And Not gobjProgress.Cancelled

'  If IsModuleEnabled(modIntranet) Then
'    strFileName = "HRProInt-" & CStr(App.Major) & "-" & CStr(App.Minor) & ".sql"
'    If Dir(strScriptPath & strFileName) <> vbNullString Then
'      gobjProgress.Bar1Caption = strFileName
'      RunScript2 strScriptPath & strFileName, sConnect
'    End If
'  End If


  '// Recreate these incase they were modified in an Update Script
  Call CreateSP_GetCurrentUsers
  Call CreateSP_GetCurrentUsersFromMaster
 
  ' Enable the SQL service broker
  EnableServiceBroker
  
  If glngSQLVersion > 8 Then
    ' AE20080229 Fault #12968
    If bOK And bRegenerateProc Then
      bOK = RegenerateSQLProcessAccount(, , True)
      bRegenerateProc = True
    End If
  End If
    
  ClearDownCurrentSessions
        
  UnlockDatabase lckSaving

  gobjProgress.CloseProgress
  UpdateDatabase = True

TidyUpAndExit:
  Screen.MousePointer = vbDefault
  Set fso = Nothing
  Set rsInfo = Nothing
  Exit Function

LocalErr:
  gobjProgress.CloseProgress
  UpdateDatabase = False
  Screen.MousePointer = vbDefault
  MsgBox "Error running update script '" & strFileName & "'" & vbNewLine & _
         Mid(Err.Description, InStrRev(Err.Description, "]") + 1), vbExclamation
  Err.Clear
  
  GoTo TidyUpAndExit

End Function


'Private Function RunScript2(strFileName As String, sConnect As String) As Boolean
'
'  Dim sUpdateScript As SystemMgr.cStringBuilder
'  Dim sReadString As String
'
'  'sUpdateScript = vbNullString
'  Set sUpdateScript = New SystemMgr.cStringBuilder
'  sUpdateScript.TheString = vbNullString
'
'  Open strFileName For Input As #1
'  Do While Not EOF(1)
'    Line Input #1, sReadString
'    'sUpdateScript = sUpdateScript & sReadString & vbNewLine
'    sUpdateScript.Append sReadString & vbNewLine
'  Loop
'  Close #1
'
'  If sUpdateScript.Length > 0 Then
'    sReadString = sUpdateScript.ToString
'    gADOCon.Execute sReadString, , adCmdText + adExecuteNoRecords
'  End If
'
'End Function


Private Function RunScript(strFileName As String, sConnect As String) As Boolean

  Dim sUpdateScript As String
  Dim sReadString As String

  sUpdateScript = vbNullString

  Open strFileName For Input As #1
  Do While Not EOF(1)
    Line Input #1, sReadString
    sUpdateScript = sUpdateScript & sReadString & vbNewLine
  Loop
  Close #1

  If sUpdateScript <> vbNullString Then
    gADOCon.Execute sUpdateScript, , adCmdText + adExecuteNoRecords
  End If

End Function


Public Function GetOldServerName() As String

  Dim sSQL As String
  Dim rsSQLInfo As ADODB.Recordset

  sSQL = "SELECT @@SERVERNAME"
  
  Set rsSQLInfo = New ADODB.Recordset
  rsSQLInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsSQLInfo
    If Not (.BOF And .EOF) Then
      GetOldServerName = UCase$(.Fields(0).value)
    End If
    .Close
  End With
  Set rsSQLInfo = Nothing
  
End Function

Public Function GetServerName() As String

  Dim sSQL As String
  Dim rsSQLInfo As ADODB.Recordset

  ' AE20090108 Fault #13490 - @@SERVERNAME doesn't recognise network name changes
  'sSQL = "SELECT @@SERVERNAME"
  sSQL = "SELECT SERVERPROPERTY('servername')"
  
  Set rsSQLInfo = New ADODB.Recordset
  rsSQLInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsSQLInfo
    If Not (.BOF And .EOF) Then
      GetServerName = UCase$(.Fields(0).value)
    End If
    .Close
  End With
  Set rsSQLInfo = Nothing
  
End Function

Public Function GetDBName()

  Dim sSQL As String
  Dim rsSQLInfo As ADODB.Recordset
 
  sSQL = "SELECT DB_NAME()"
  
  Set rsSQLInfo = New ADODB.Recordset
  rsSQLInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsSQLInfo
    If Not (.BOF And .EOF) Then
      GetDBName = UCase$(.Fields(0).value)
    End If
    .Close
  End With
  Set rsSQLInfo = Nothing

End Function

Public Function ValidDatabaseMailDetails(psLastEmailProfile As String) As Boolean
    
On Local Error GoTo ValidDBMailErr

  Dim sSQL As String
  Dim rsProfiles As ADODB.Recordset
        
  If psLastEmailProfile = vbNullString Then
    psLastEmailProfile = "<Use Default Profile>"
  End If
      
  sSQL = "exec msdb..sysmail_help_principalprofile_sp"
  Set rsProfiles = New ADODB.Recordset
  rsProfiles.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly

  If Not (rsProfiles.BOF And rsProfiles.EOF) Then
    rsProfiles.MoveFirst
    
    Do Until rsProfiles.EOF
      If rsProfiles!profile_name = psLastEmailProfile Then
        ValidDatabaseMailDetails = True
        Exit Do
      ElseIf psLastEmailProfile = "<Use Default Profile>" _
          And rsProfiles!is_default Then
        ValidDatabaseMailDetails = True
        Exit Do
      End If
  
      rsProfiles.MoveNext
    Loop
  End If

  rsProfiles.Close

' AE20080325 Fault #12960
TidyUpAndExit:
  If Not rsProfiles Is Nothing Then
    If rsProfiles.State = adStateOpen Then
      rsProfiles.Close
    End If
    Set rsProfiles = Nothing
  End If
  
  Exit Function

ValidDBMailErr:
  ValidDatabaseMailDetails = True

  Resume TidyUpAndExit

End Function

Private Function CreateSP_GetCurrentUsers() As Boolean
  ' Create the GetCurrentUsers stored procedure.
  On Error GoTo ErrorTrap

  Dim fCreatedOK As Boolean
  Dim sProcSQL As String

  fCreatedOK = True

  ' Construct the stored procedure creation string.
  sProcSQL = _
    "/* ------------------------------------------------ */" & vbNewLine & _
    "/* Update module stored procedure.           */" & vbNewLine & _
    "/* Automatically generated by the System manager.   */" & vbNewLine & _
    "/* ------------------------------------------------ */" & vbNewLine & _
    "ALTER PROCEDURE [dbo].[spASRGetCurrentUsers]" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine & vbNewLine & _
    "   SET NOCOUNT ON;" & vbNewLine & vbNewLine
    
  sProcSQL = sProcSQL & _
    "   DECLARE @Mode         smallint" & vbNewLine & vbNewLine & _
    "   SELECT @Mode = [SettingValue] FROM ASRSysSystemSettings WHERE [Section] = 'ProcessAccount' AND [SettingKey] = 'Mode';" & vbNewLine & _
    "   IF @@ROWCOUNT = 0 SET @Mode = 0" & vbNewLine & vbNewLine & _
    "   IF (@Mode = 1 OR @Mode = 2) AND (NOT IS_SRVROLEMEMBER('sysadmin') = 1)" & vbNewLine & _
    "   BEGIN" & vbNewLine & _
    "       EXECUTE sp_executeSQL N'dbo.[spASRGetCurrentUsersFromAssembly]'" & vbNewLine & _
    "   END" & vbNewLine & _
    "   ELSE" & vbNewLine & _
    "   BEGIN" & vbNewLine & _
    "       EXECUTE sp_executeSQL N'dbo.[spASRGetCurrentUsersFromMaster]'" & vbNewLine & _
    "   END" & vbNewLine & vbNewLine & _
    "END"
    
  gADOCon.Execute sProcSQL, , adExecuteNoRecords

  ' AE20080702 Fault #13158 - Grant permissions so we can execute it!
  sProcSQL = "GRANT EXEC ON [dbo].[spASRGetCurrentUsers] TO [ASRSysGroup]"
  gADOCon.Execute sProcSQL, , adExecuteNoRecords

TidyUpAndExit:
  CreateSP_GetCurrentUsers = fCreatedOK
  Exit Function

ErrorTrap:
  fCreatedOK = False
  OutputError "Error creating GetCurrentUsers stored procedure"
  Resume TidyUpAndExit

End Function

Private Function CreateSP_GetCurrentUsersFromMaster() As Boolean
  ' Create the GetCurrentUsersFromMaster stored procedure.
  ' This should only be run on SQL Versions Post 2000
  On Error GoTo ErrorTrap

  Dim fCreatedOK As Boolean
  Dim sProcSQL As String

  fCreatedOK = True
 
  ' Construct the stored procedure creation string.
  ' As non-sa user can't KILL processes we're going to ignore failed login
  ' attempts after the creation time of Save/Manual locks
  sProcSQL = _
    "/* ------------------------------------------------ */" & vbNewLine & _
    "/* Update module stored procedure.           */" & vbNewLine & _
    "/* Automatically generated by the System manager.   */" & vbNewLine & _
    "/* ------------------------------------------------ */" & vbNewLine & _
    "ALTER PROCEDURE [dbo].[spASRGetCurrentUsersFromMaster]" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine & vbNewLine & _
    "   SET NOCOUNT ON;" & vbNewLine & vbNewLine & _
    "   DECLARE @login_time datetime;" & vbNewLine & vbNewLine & _
    "   SELECT TOP 1 @login_time = l.Lock_Time" & vbNewLine & _
    "   FROM ASRSysLock l" & vbNewLine & _
    "      INNER JOIN master..sysprocesses p ON p.spid = l.spid" & vbNewLine & _
    "         AND p.dbID = DB_ID()" & vbNewLine & _
    "         AND p.login_time = l.Login_Time" & vbNewLine & _
    "   WHERE L.Priority < 3" & vbNewLine & _
    "   ORDER BY l.Priority;" & vbNewLine & vbNewLine & _
    "   SET @login_time = ISNULL(@login_time, GETDATE());" & vbNewLine & vbNewLine
    
   sProcSQL = sProcSQL & _
    "   SELECT p.hostname, p.loginame, p.program_name, p.hostprocess" & vbNewLine & _
    "        , p.sid, p.login_time, p.spid, p.uid" & vbNewLine & _
    "   FROM     master..sysprocesses p" & vbNewLine & _
    "   WHERE    p.program_name LIKE 'OpenHR%'" & vbNewLine & _
    "     AND    p.program_name NOT LIKE 'OpenHR Web%'" & vbNewLine & _
    "     AND    p.program_name NOT LIKE 'OpenHR Workflow%'" & vbNewLine & _
    "     AND    p.program_name NOT LIKE 'OpenHR Mobile%'" & vbNewLine & _
    "     AND    p.program_name NOT LIKE 'OpenHR Outlook%'" & vbNewLine & _
    "     AND    p.program_name NOT LIKE 'System Framework Assembly%'" & vbNewLine & _
    "     AND    p.program_name NOT LIKE 'OpenHR Intranet Embedding%'" & vbNewLine & _
    "     AND    p.dbID = DB_ID()" & vbNewLine & _
    "     AND (p.login_Time < @login_time)" & vbNewLine & _
    "UNION" & vbNewLine & _
    "    SELECT DISTINCT HostName COLLATE SQL_Latin1_General_CP1_CI_AS, UserName COLLATE SQL_Latin1_General_CP1_CI_AS" & vbNewLine & _
    "     , WebArea COLLATE SQL_Latin1_General_CP1_CI_AS, 0, 0, null, 0, 0 FROM ASRSysCurrentSessions" & vbNewLine & _
    "   ORDER BY loginame;" & vbNewLine & vbNewLine & _
    "END"
  
  gADOCon.Execute sProcSQL, , adExecuteNoRecords

  ' AE20080702 Fault #13158 - Grant permissions so we can execute it!
  sProcSQL = "GRANT EXEC ON [dbo].[spASRGetCurrentUsersFromMaster] TO [ASRSysGroup]"
  gADOCon.Execute sProcSQL, , adExecuteNoRecords
  
TidyUpAndExit:
  CreateSP_GetCurrentUsersFromMaster = fCreatedOK
  Exit Function

ErrorTrap:
  fCreatedOK = False
  OutputError "Error creating GetCurrentUsersFromMaster stored procedure"
  Resume TidyUpAndExit

End Function

Private Function CreateSP_LockCheck() As Boolean

  ' Create the LockCheck stored procedure for use with the .NET Assembly
  
  On Error GoTo ErrorTrap

  Dim fCreatedOK As Boolean
  Dim sProcSQL As String

  fCreatedOK = True

  ' This should only be run on SQL Versions Post 2005

  ' Construct the stored procedure creation string.
  sProcSQL = _
    "/* ------------------------------------------------ */" & vbNewLine & _
    "/* Update module stored procedure.           */" & vbNewLine & _
    "/* Automatically generated by the System manager.   */" & vbNewLine & _
    "/* ------------------------------------------------ */" & vbNewLine & _
    "ALTER PROCEDURE [dbo].[sp_ASRLockCheck] AS " & vbNewLine & _
    "   BEGIN" & vbNewLine & _
    "     SET NOCOUNT ON;" & vbNewLine & vbNewLine & _
    "     IF APP_NAME() <> 'OpenHR Workflow Service' AND APP_NAME() <> 'OpenHR Outlook Calendar Service'" & vbNewLine & _
    "     BEGIN" & vbNewLine
    
  sProcSQL = sProcSQL & _
    "       CREATE TABLE #tmpProcesses " & vbNewLine & _
    "         (HostName varchar(100)" & vbNewLine & _
    "         ,LoginName varchar(100)" & vbNewLine & _
    "         ,Program_Name varchar(100)" & vbNewLine & _
    "         ,HostProcess int" & vbNewLine & _
    "         ,Sid binary(86)" & vbNewLine & _
    "         ,Login_Time datetime" & vbNewLine & _
    "         ,spid int" & vbNewLine & _
    "         ,uid smallint)" & vbNewLine & _
    "       INSERT #tmpProcesses EXEC dbo.[spASRGetCurrentUsers]" & vbNewLine & _
    "       SELECT l.* FROM dbo.ASRSysLock l" & vbNewLine & _
    "       LEFT OUTER JOIN #tmpProcesses syspro ON l.spid = syspro.spid AND l.login_time = syspro.login_time" & vbNewLine & _
    "       WHERE l.[priority] = 2 OR syspro.spid IS NOT NULL" & vbNewLine & _
    "       ORDER BY l.[priority]" & vbNewLine & _
    "       DROP TABLE #tmpProcesses" & vbNewLine & _
    "     END" & vbNewLine
    
  sProcSQL = sProcSQL & _
    "     ELSE" & vbNewLine & _
    "     BEGIN" & vbNewLine & _
    "       SELECT l.* FROM dbo.ASRSysLock l" & vbNewLine & _
    "       LEFT OUTER JOIN master..sysprocesses syspro ON l.spid = syspro.spid AND l.login_time = syspro.login_time" & vbNewLine & _
    "       WHERE l.[priority] = 2 OR syspro.spid IS NOT NULL" & vbNewLine & _
    "       ORDER BY l.[priority]" & vbNewLine & _
    "     END" & vbNewLine & _
    "   END"

  gADOCon.Execute sProcSQL, , adExecuteNoRecords

  sProcSQL = "GRANT EXEC ON [dbo].[sp_ASRLockCheck] TO [ASRSysGroup]"
  gADOCon.Execute sProcSQL, , adExecuteNoRecords

TidyUpAndExit:
  CreateSP_LockCheck = fCreatedOK
  Exit Function

ErrorTrap:
  fCreatedOK = False
  OutputError "Error creating LockCheck stored procedure"
  Resume TidyUpAndExit

End Function

Private Function CreateUDF_SQLVersion() As Boolean
  ' Create the udfASRSQLVersion function.
  On Error GoTo ErrorTrap

  Dim fCreatedOK As Boolean
  Dim sUDFSQL As String

  fCreatedOK = True
  
  sUDFSQL = "IF EXISTS" & _
      " (SELECT Name" & _
      "   FROM sysobjects" & _
      "   WHERE id = object_id('[dbo].[udfASRSQLVersion]')" & _
      "     AND sysstat & 0xf = 0)" & _
      " DROP FUNCTION [dbo].[udfASRSQLVersion]"
  gADOCon.Execute sUDFSQL, , adExecuteNoRecords
  
  ' Construct the function creation string.
  sUDFSQL = _
    "/* ------------------------------------------------ */" & vbNewLine & _
    "/* System module user defined function.      */" & vbNewLine & _
    "/* Automatically generated by the System manager.   */" & vbNewLine & _
    "/* ------------------------------------------------ */" & vbNewLine & _
    "CREATE FUNCTION [dbo].[udfASRSQLVersion]" & vbNewLine & _
    "(" & vbNewLine & _
    ")" & vbNewLine & _
    "RETURNS integer" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine & _
    "  RETURN convert(int,convert(float,substring(@@version,charindex('-',@@version)+2,2)))" & vbNewLine & _
    "END"

  gADOCon.Execute sUDFSQL, , adExecuteNoRecords
    
TidyUpAndExit:
  CreateUDF_SQLVersion = fCreatedOK
  Exit Function

ErrorTrap:
  fCreatedOK = False
  OutputError "Error creating SQLVersion function"
  Resume TidyUpAndExit
End Function

Public Function ApplyHotfixes(ByRef RunType As HotfixType) As Boolean

  On Error GoTo ErrorTrap

  Dim cmdHotfixes As New ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim bOK As Boolean

  bOK = True

  With cmdHotfixes
    .CommandText = "spASRApplyScripts"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon
    
    Set pmADO = .CreateParameter("runtype", adInteger, adParamInput)
    pmADO.value = RunType
    .Parameters.Append pmADO
    
    .Execute
  End With

TidyUpAndExit:
  Set cmdHotfixes = Nothing
  ApplyHotfixes = bOK
  Exit Function

ErrorTrap:
  bOK = False
  GoTo TidyUpAndExit

End Function

Public Function UploadHotfixes() As Boolean

  Dim strScriptPath As String
  Dim iLineNo As Integer
  Dim strFile As String
  Dim sReadString As String
  Dim bOK As Boolean
  
  Dim iHotfixNumber As Integer
  Dim iRunType As HotfixType
  Dim sScript As String
  Dim sRunInVersion As String
  Dim bRunOnce As Boolean
  Dim sChecksum As String
  Dim bIsValid As Boolean
  Dim sDescription As String
  Dim iSequence As Integer
  Dim DatabaseGUID As String
  Dim bUploaded As Boolean
 
  Dim lCRC32 As Long
 
  bOK = True
 
  On Error GoTo ErrorTrap

  strScriptPath = gsApplicationPath & "\Update Scripts\"
  strFile = Dir(strScriptPath & "script*.sql")
  Do While strFile <> ""

    sScript = vbNullString
    sReadString = vbNullString
    iLineNo = 1
    bIsValid = True
    
    Open strScriptPath & strFile For Input As #1
    Do While Not EOF(1)
      Line Input #1, sReadString
      
      ' Process header line
      Select Case iLineNo
        Case 1, 10
          ' Do nothing
        Case 2
          iHotfixNumber = Mid(sReadString, 15, 7)
        Case 3
          sDescription = Mid(sReadString, 18, Len(sReadString))
        Case 4
          iRunType = Mid(sReadString, 15, 1)
        Case 5
          sRunInVersion = Mid(sReadString, 21, 6)
        Case 6
          bRunOnce = IIf(Mid(sReadString, 15, 3) = "Yes", True, False)
        Case 7
          iSequence = Mid(sReadString, 15, Len(sReadString))
        Case 8
          DatabaseGUID = Mid(sReadString, 18, Len(sReadString))
        Case 9
          sChecksum = Mid(sReadString, 16, Len(sReadString))
        Case Else
          sScript = sScript & sReadString & vbNewLine
        
      End Select
      
      iLineNo = iLineNo + 1
    
    Loop
    Close #1
   
    ' Validate the file?
    'cStream.File = strScriptPath & strFile
'    lCRC32 = cCRC32.GetFileCrc32(cStream)
    bIsValid = (sChecksum = "0x" & Hex(GetChecksum(sScript)))
   
    ' Upload the file
    If bIsValid And Not bUploaded Then
      bOK = UploadScript(iRunType, sScript, bRunOnce, sRunInVersion, iSequence, sDescription)
      
      ' Mark file as uploaded
      If bOK Then
        Name strScriptPath & strFile As strScriptPath & "uploaded-" & strFile
      End If

    End If
    
    strFile = Dir()
    
  Loop

TidyUpAndExit:
  UploadHotfixes = bOK
  Exit Function

ErrorTrap:

  MsgBox "Error uploading custom scripts" & vbNewLine & Err.Description _
      , vbOKOnly + vbExclamation, Application.Name

  bOK = False
  GoTo TidyUpAndExit

End Function

Private Function UploadScript(ByVal RunType As HotfixType, ByVal Script As String _
    , RunOnce As Boolean, RunInVersion As String, Sequence As Integer, Description As String) As Boolean

  On Error GoTo ErrorTrap

  Dim cmdHotfixes As New ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim bOK As Boolean

  bOK = True

  With cmdHotfixes
    .CommandText = "spASRUploadScript"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon
    
    Set pmADO = .CreateParameter("runtype", adInteger, adParamInput)
    .Parameters.Append pmADO
    pmADO.value = RunType
    
    Set pmADO = .CreateParameter("script", adLongVarChar, adParamInput, VARCHAR_MAX_Size)
    .Parameters.Append pmADO
    pmADO.value = Script
    
    Set pmADO = .CreateParameter("runonce", adBoolean, adParamInput)
    .Parameters.Append pmADO
    pmADO.value = RunOnce
    
    Set pmADO = .CreateParameter("runinversion", adLongVarChar, adParamInput, 10)
    .Parameters.Append pmADO
    pmADO.value = RunInVersion
        
    Set pmADO = .CreateParameter("sequence", adInteger, adParamInput)
    .Parameters.Append pmADO
    pmADO.value = Sequence

    Set pmADO = .CreateParameter("description", adLongVarChar, adParamInput, VARCHAR_MAX_Size)
    .Parameters.Append pmADO
    pmADO.value = Description
        
    .Execute
  End With

TidyUpAndExit:
  Set cmdHotfixes = Nothing
  UploadScript = bOK
  Exit Function

ErrorTrap:
  bOK = False
  GoTo TidyUpAndExit

End Function

  ' Calculates the checksum for a sentence
  Public Function GetChecksum(ByVal sentence As String) As String
    
    ' Loop through all chars to get a checksum
    Dim Character As String
    Dim Checksum As Long
    Dim iCount As Integer
    
    Checksum = 0
    For iCount = 1 To Len(sentence)
      Character = Mid(sentence, iCount, 1)
      Checksum = Checksum + Asc(Character)
    Next
    
    ' Return the checksum formatted as a two-character hexadecimal
    GetChecksum = Checksum
  
  End Function

Private Function ClearDownCurrentSessions() As Boolean

  On Error GoTo ErrorTrap

  Dim sSQL As String
  Dim bOK As Boolean
    
  bOK = True
  sSQL = "DELETE FROM ASRSysCurrentSessions"
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords

TidyUpAndExit:
  ClearDownCurrentSessions = bOK
  Exit Function

ErrorTrap:
  bOK = False
  
End Function

