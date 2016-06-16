Attribute VB_Name = "modSQLPlatform"
Option Explicit

Public Function EnableServiceBroker() As Boolean
  
  Dim sSQL As String
  Dim bOK As Boolean
  
  On Error GoTo LocalErr
  
  bOK = True
  
  sSQL = "EXEC sp_ExecuteSQL spASREnableServiceBroker"
  gADOCon.Execute sSQL, -1, adExecuteNoRecords
  
TidyUpAndExit:
  EnableServiceBroker = bOK
  Exit Function
  
LocalErr:
  bOK = False
  Screen.MousePointer = vbDefault
  OutputError "Error enabling the service broker. (" & Err.Description & ")"
  Err.Clear
  GoTo TidyUpAndExit
  
End Function

Public Function SetDatabaseCompatability() As Boolean

  Dim sSQL As String
  Dim bOK As Boolean
  
  On Error GoTo LocalErr
  
  bOK = True
  Select Case Int(glngSQLVersion)
    Case 7
      sSQL = "EXEC sp_dbcmptlevel '" & gsDatabaseName & "', 70"
      ' Fault 11855 - Do not issue for SQL7 databases.
    Case 8
      sSQL = "EXEC sp_dbcmptlevel '" & gsDatabaseName & "', 80"
      gADOCon.Execute sSQL, -1, adExecuteNoRecords
    Case 9
      sSQL = "EXEC sp_dbcmptlevel '" & gsDatabaseName & "', 90"
      gADOCon.Execute sSQL, -1, adExecuteNoRecords
    Case 10
      sSQL = "ALTER DATABASE [" & gsDatabaseName & "] SET COMPATIBILITY_LEVEL=100"
      gADOCon.Execute sSQL, -1, adExecuteNoRecords
    Case 11
      sSQL = "ALTER DATABASE [" & gsDatabaseName & "] SET COMPATIBILITY_LEVEL=110"
      gADOCon.Execute sSQL, -1, adExecuteNoRecords
    Case 12
      sSQL = "ALTER DATABASE [" & gsDatabaseName & "] SET COMPATIBILITY_LEVEL=120"
      gADOCon.Execute sSQL, -1, adExecuteNoRecords
    Case 13
      sSQL = "ALTER DATABASE [" & gsDatabaseName & "] SET COMPATIBILITY_LEVEL=130"
      gADOCon.Execute sSQL, -1, adExecuteNoRecords
    Case Else
      GoTo LocalErr
  End Select
  
TidyUpAndExit:
  SetDatabaseCompatability = bOK
  Exit Function

LocalErr:
  bOK = False
  Screen.MousePointer = vbDefault
  Err.Clear
  OutputError "Error setting database compatibility."
  GoTo TidyUpAndExit

End Function

Public Function MarkDatabaseAsTrustworthy() As Boolean
  
  On Error GoTo ErrorTrap
  
  Dim bOK As Boolean
  Dim sSQL As String
  Dim rsTrusted As ADODB.Recordset
  Dim bTrusted As Boolean
  
  bTrusted = False
    
  ' If already marked as trustworthy then return true
  sSQL = "SELECT is_trustworthy_on from sys.databases where [Name] = '" & gsDatabaseName & "'"
  Set rsTrusted = New ADODB.Recordset
  rsTrusted.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  
  If Not (rsTrusted.EOF And rsTrusted.BOF) Then
    bTrusted = rsTrusted.Fields(0).value
  End If
      
  rsTrusted.Close
      
  If Not bTrusted Then
    If gbCurrentUserIsSysSecMgr Then
      sSQL = "ALTER DATABASE [" & gsDatabaseName & "] SET TRUSTWORTHY ON"
      gADOCon.Execute sSQL, , adExecuteNoRecords
      bTrusted = True
    Else
      MsgBox "The database cannot be marked as trustworthy." & vbNewLine & _
        "A system administrator must log into the System Manager to perform this operation." & vbNewLine & vbNewLine & _
        vbExclamation + vbOKOnly, Application.Name
    End If
  End If


TidyUpAndExit:
  Set rsTrusted = Nothing
  MarkDatabaseAsTrustworthy = bTrusted
  Exit Function

ErrorTrap:
  bTrusted = False
  GoTo TidyUpAndExit
 
End Function

Public Function SetAllowUpdatesOff() As Boolean

  On Error GoTo ErrorTrap

  Dim sSQL As String
  Dim bOK As Boolean

  bOK = True

  If gbCurrentUserIsSysSecMgr Then
    '// Put it all in one batch so if it fails we retain the correct db
    sSQL = "USE [master];" & vbNewLine
    sSQL = sSQL & "EXEC sp_configure 'allow updates',0;" & vbNewLine
    sSQL = sSQL & "RECONFIGURE WITH OVERRIDE;" & vbNewLine
    sSQL = sSQL & "USE [" & gsDatabaseName & "];"
    gADOCon.Execute sSQL, , adExecuteNoRecords
  End If

TidyUpAndExit:
  SetAllowUpdatesOff = bOK
  Exit Function
  
ErrorTrap:
  bOK = False
  Resume TidyUpAndExit
  
End Function

Public Function SurfaceAreaConfig_EnableCLR() As Boolean

  On Error GoTo ErrorTrap

  Dim sSQL As String
  Dim bOK As Boolean

  bOK = True

  ' NPG20081121 Fault 13429
  ' Added the WITH OVERRIDE parameter
  If gbCurrentUserIsSysSecMgr Then
    gADOCon.Execute "USE master", , adExecuteNoRecords
    gADOCon.Execute "EXEC sp_configure 'show advanced options', 1;", , adExecuteNoRecords
    gADOCon.Execute "RECONFIGURE WITH OVERRIDE", , adExecuteNoRecords
    gADOCon.Execute "EXEC sp_configure 'clr enabled', 1;", , adExecuteNoRecords
    gADOCon.Execute "RECONFIGURE WITH OVERRIDE", , adExecuteNoRecords
    gADOCon.Execute "USE [" & gsDatabaseName & "]", , adExecuteNoRecords
  End If

TidyUpAndExit:
  SurfaceAreaConfig_EnableCLR = bOK
  Exit Function
  
ErrorTrap:
  bOK = False
  Resume TidyUpAndExit

End Function

Public Function SurfaceAreaConfig_EnableOLE() As Boolean

  On Error GoTo ErrorTrap

  Dim sSQL As String
  Dim bOK As Boolean

  bOK = True

  If gbCurrentUserIsSysSecMgr Then
    gADOCon.Execute "USE master", , adExecuteNoRecords
    gADOCon.Execute "EXEC sp_configure 'show advanced options', 1;", , adExecuteNoRecords
    gADOCon.Execute "RECONFIGURE WITH OVERRIDE", , adExecuteNoRecords
    gADOCon.Execute "EXEC sp_configure 'ole automation procedures', 1;", , adExecuteNoRecords
    gADOCon.Execute "RECONFIGURE WITH OVERRIDE", , adExecuteNoRecords
    gADOCon.Execute "USE [" & gsDatabaseName & "]", , adExecuteNoRecords
  End If

TidyUpAndExit:
  SurfaceAreaConfig_EnableOLE = bOK
  Exit Function
  
ErrorTrap:
  bOK = False
  Resume TidyUpAndExit

End Function

Public Function GenerateDropServerAssemblySP(ByRef bRunIt As Boolean) As Boolean

  Dim sSQL As String
  Dim bOK As Boolean
  
  On Error GoTo ErrorTrap
 
  bOK = True
  
  ' Drop existing [spASRDropServerAssembly]
  sSQL = "IF EXISTS" & _
    " (SELECT Name" & _
    "   FROM sysobjects" & _
    "   WHERE id = object_id('spASRDropServerAssembly')" & _
    "     AND sysstat & 0xf = 4)" & _
    " DROP PROCEDURE spASRDropServerAssembly"
  gADOCon.Execute sSQL, , adExecuteNoRecords
  

  sSQL = "/* ------------------------------------------------ */" & vbNewLine & _
    "/* Drop System Framework Assembly stored procedure. */" & vbNewLine & _
    "/* Automatically generated by the System Manager.   */" & vbNewLine & _
    "/* ------------------------------------------------ */" & vbNewLine & _
    "CREATE PROCEDURE dbo.spASRDropServerAssembly" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine
      
  sSQL = sSQL & _
      "  IF EXISTS (SELECT name FROM sys.assemblies WHERE name IN (N'System Framework Assembly', N'HR Pro Server.NET'))" & vbNewLine & _
      "  BEGIN" & vbNewLine & _
      "  DECLARE @error int" & vbNewLine & _
      "  SET @error = 0" & vbNewLine & vbNewLine & _
      "  /* Drop the assembly user defined aggregates, triggers, functions and procedures */" & vbNewLine & _
      "  DECLARE @moduleId sysname" & vbNewLine & _
      "  DECLARE @moduleName sysname" & vbNewLine & _
      "  DECLARE @moduleType char(2)" & vbNewLine & _
      "  DECLARE @moduleClass tinyint" & vbNewLine & _
      "  DECLARE assemblyModules CURSOR FAST_FORWARD FOR" & vbNewLine & _
      "    SELECT t.object_id, t.name, t.type, t.parent_class as class" & vbNewLine & _
      "      FROM sys.triggers t" & vbNewLine & _
      "      INNER JOIN sys.assembly_modules m ON t.object_id = m.object_id" & vbNewLine & _
      "      INNER JOIN sys.assemblies a ON m.assembly_id = a.assembly_id" & vbNewLine & _
      "      WHERE a.Name IN (N'System Framework Assembly', N'HR Pro Server.NET')" & vbNewLine & _
      "    UNION" & vbNewLine & _
      "    SELECT o.object_id, o.name, o.type, NULL as class" & vbNewLine & _
      "      FROM sys.objects o" & vbNewLine & _
      "      INNER JOIN sys.assembly_modules m ON o.object_id = m.object_id" & vbNewLine & _
      "      INNER JOIN sys.assemblies a ON m.assembly_id = a.assembly_id" & vbNewLine & _
      "      WHERE a.Name IN (N'System Framework Assembly', N'HR Pro Server.NET')" & vbNewLine & _
      "  OPEN assemblyModules" & vbNewLine

  sSQL = sSQL & _
      "  FETCH NEXT FROM assemblyModules INTO @moduleId, @moduleName, @moduleType, @moduleClass" & vbNewLine & _
      "  WHILE (@error = 0 AND @@FETCH_STATUS = 0)" & vbNewLine & _
      "  BEGIN" & vbNewLine & _
      "    DECLARE @dropModuleString nvarchar(256)" & vbNewLine & _
      "    IF (@moduleType = 'AF') SET @dropModuleString = N'AGGREGATE'" & vbNewLine & _
      "    IF (@moduleType = 'TA') SET @dropModuleString = N'TRIGGER'" & vbNewLine & _
      "    IF (@moduleType = 'FT' OR @moduleType = 'FS') SET @dropModuleString = N'FUNCTION'" & vbNewLine & _
      "    IF (@moduleType = 'PC') SET @dropModuleString = N'PROCEDURE'" & vbNewLine & _
      "        SET @dropModuleString = N'DROP ' + @dropModuleString + ' [' + REPLACE(@moduleName, ']', ']]') + ']'" & vbNewLine & _
      "    IF (@moduleType = 'TA' AND @moduleClass = 0)" & vbNewLine & _
      "    BEGIN" & vbNewLine & _
      "      SET @dropModuleString = @dropModuleString + N' ON DATABASE'" & vbNewLine & _
      "    END" & vbNewLine & _
      "    EXEC sp_executesql @dropModuleString" & vbNewLine & _
      "    FETCH NEXT FROM assemblyModules INTO @moduleId, @moduleName, @moduleType, @moduleClass" & vbNewLine & _
      "  END" & vbNewLine & _
      "  CLOSE assemblyModules" & vbNewLine & _
      "  DEALLOCATE assemblyModules" & vbNewLine
      
  sSQL = sSQL & _
      "  /* Drop the assembly user defined types */" & vbNewLine & _
      "  DECLARE @typeId int" & vbNewLine & _
      "  DECLARE @typeName sysname" & vbNewLine & _
      "  DECLARE assemblyTypes CURSOR FAST_FORWARD" & vbNewLine & _
      "    FOR SELECT t.user_type_id, t.name" & vbNewLine & _
      "      FROM sys.assembly_types t" & vbNewLine & _
      "      INNER JOIN sys.assemblies a ON t.assembly_id = a.assembly_id" & vbNewLine & _
      "      WHERE a.Name IN (N'System Framework Assembly')" & vbNewLine & _
      "  OPEN assemblyTypes" & vbNewLine & _
      "  FETCH NEXT FROM assemblyTypes INTO @typeId, @typeName" & vbNewLine & _
      "  WHILE (@error = 0 AND @@FETCH_STATUS = 0)" & vbNewLine & _
      "  BEGIN" & vbNewLine & _
      "    DECLARE @dropTypeString nvarchar(256)" & vbNewLine & _
      "    SET @dropTypeString = N'DROP TYPE [' + REPLACE(@typeName, ']', ']]') + ']'" & vbNewLine & _
      "    IF NOT EXISTS (SELECT name FROM sys.extended_properties WHERE major_id = @typeId AND name = 'AutoDeployed')" & vbNewLine & _
      "    BEGIN" & vbNewLine & _
      "      DECLARE @quotedTypeName sysname" & vbNewLine & _
      "      SET @quotedTypeName = REPLACE(@typeName, '''', '''''')" & vbNewLine & _
      "      RAISERROR(N'The assembly user defined type ''%s'' cannot be preserved because it was not automatically deployed.', 16, 1,@quotedTypeName)" & vbNewLine & _
      "      SET @error = @@ERROR" & vbNewLine & _
      "    END" & vbNewLine & _
      "    ELSE" & vbNewLine & _
      "    BEGIN" & vbNewLine
        
  sSQL = sSQL & _
      "      EXEC sp_executesql @dropTypeString" & vbNewLine & _
      "      FETCH NEXT FROM assemblyTypes INTO @typeId, @typeName" & vbNewLine & _
      "    END" & vbNewLine & _
      "  END" & vbNewLine & _
      "  CLOSE assemblyTypes" & vbNewLine & _
      "  DEALLOCATE assemblyTypes" & vbNewLine & vbNewLine & _
      "  /* Drop the assembly */" & vbNewLine & _
      "  IF (@error = 0)" & vbNewLine & _
      "    IF EXISTS (SELECT name FROM sys.assemblies WHERE name = N'System Framework Assembly')" & vbNewLine & _
      "       DROP ASSEMBLY [System Framework Assembly] WITH NO DEPENDENTS" & vbNewLine & vbNewLine & _
      "    IF EXISTS (SELECT name FROM sys.assemblies WHERE name = N'HR Pro Server.NET')" & vbNewLine & _
      "       DROP ASSEMBLY [HR Pro Server.NET] WITH NO DEPENDENTS" & vbNewLine & vbNewLine & _
      "  END" & vbNewLine & _
      "END"
  
  ' Lets do it!
  gADOCon.Execute sSQL, , adExecuteNoRecords

  ' Since we've gone to the trouble of creating this great script we might as well run it...
  If bRunIt Then
    gADOCon.Execute "EXEC dbo.spASRDropServerAssembly", , adExecuteNoRecords
  End If

TidyUpAndExit:
  GenerateDropServerAssemblySP = bOK
  Exit Function
  
ErrorTrap:
  bOK = False
  OutputError "Error in GenerateDropServerAssemblySP"
  Resume TidyUpAndExit

End Function

' When databases are moved from server to server you can get the following error
' The database owner SID recorded in the master database differs from the database owner SID
' This function resets the owner to be the same as the master database.
Public Function RefreshDatabaseOwner() As Boolean

  Dim sSQL As String
  Dim bOK As Boolean

  On Error GoTo ErrorTrap
  bOK = True

  sSQL = "DECLARE @Owner nvarchar(1000)" & vbNewLine & _
         "SELECT @Owner = suser_sname(Owner_SID) FROM sys.databases WHERE Name = 'master'" & vbNewLine & _
         "EXEC dbo.sp_changedbowner @loginame = @Owner, @map = false"
  gADOCon.Execute sSQL, , adExecuteNoRecords
  
TidyUpAndExit:
  RefreshDatabaseOwner = bOK
  Exit Function
  
ErrorTrap:
  bOK = False
  Resume TidyUpAndExit

End Function

' Get the prerequisites for installation of an assembly
Public Function MakeServerReadyForAssembly() As Boolean

  On Error GoTo ErrorTrap:

  Dim bOK As Boolean

  bOK = True

  ' Need to have the correct database compatability settings
  bOK = SetDatabaseCompatability
  If Not bOK Then
    MakeServerReadyForAssembly = False
    Exit Function
  End If

  ' For the assembly to have external and unsafe settings the database must be marked as trustworthy.
  bOK = MarkDatabaseAsTrustworthy
  If Not bOK Then
    MakeServerReadyForAssembly = False
    Exit Function
  End If
  
  ' Make sure the database owner is attached correctly
  bOK = RefreshDatabaseOwner
  If Not bOK Then
    MakeServerReadyForAssembly = False
    Exit Function
  End If
  
  ' If the 'allow updates' option is on the switch it off
  ' see http://msdn.microsoft.com/en-us/library/ms179262(SQL.90).aspx
  ' AE20090119 Fault #13498
  bOK = SetAllowUpdatesOff
  If Not bOK Then
    MakeServerReadyForAssembly = False
    Exit Function
  End If
  
  ' For the assembly to work we must have CLR enabled on the server.
  bOK = SurfaceAreaConfig_EnableCLR
  If Not bOK Then
    MakeServerReadyForAssembly = False
    Exit Function
  End If

  ' Enable OLE Automation Procedures on the server.
  bOK = SurfaceAreaConfig_EnableOLE
  If Not bOK Then
    MakeServerReadyForAssembly = False
    Exit Function
  End If

TidyUpAndExit:
  MakeServerReadyForAssembly = bOK
  Exit Function

ErrorTrap:
  bOK = False
  OutputError "Unable to make the server ready for deploying the assembly" & vbNewLine & vbNewLine _
    & "Please contact your system administrator."
  GoTo TidyUpAndExit

End Function

Public Function RegenerateSQLProcessAccount( _
    Optional psName As String, _
    Optional psPassword As String, _
    Optional pbTrusted As Boolean) As Boolean

  Dim glngProcessMethod As ProcessAdminConfig
  Dim strEncrypted As String
  Dim bOK As Boolean
  Dim strLogon As String
  Dim sName As String
  Dim sPassword As String
  Dim sDatabase As String
  Dim sServer As String
  
  Dim sSQL As String
    
  On Error GoTo LocalErr

  If glngSQLVersion < 9 Then
    RegenerateSQLProcessAccount = True
    Exit Function
  End If

  strEncrypted = GetSystemLogon()
    
  If strEncrypted = vbNullString Then
    If psName = vbNullString And psPassword = vbNullString Then
      strEncrypted = EncryptLogonDetails("", "", gsDatabaseName, gsServerName)
      glngProcessMethod = iPROCESSADMIN_SERVICEACCOUNT
    Else
      strEncrypted = EncryptLogonDetails(psName, psPassword, gsDatabaseName, gsServerName)
      glngProcessMethod = iPROCESSADMIN_SQLACCOUNT
    End If
    
    ' AE20081003 Fault #13387
    sSQL = _
          "DELETE FROM [ASRSysModuleSetup]" & vbNewLine & _
          "WHERE  [ModuleKey] = '" & gsMODULEKEY_SQL & "'" & vbNewLine & _
          "AND    [ParameterKey] = '" & gsPARAMETERKEY_LOGINDETAILS & "'" & vbNewLine
        
    sSQL = sSQL & vbNewLine & _
          "INSERT INTO [ASRSysModuleSetup]" & vbNewLine & _
          "     ([ModuleKey]" & vbNewLine & _
          "     ,[ParameterKey]" & vbNewLine & _
          "     ,[ParameterValue]" & vbNewLine & _
          "     ,[ParameterType])" & vbNewLine & _
          "VALUES (" & vbNewLine & _
          "       '" & gsMODULEKEY_SQL & "'" & vbNewLine & _
          "      ,'" & gsPARAMETERKEY_LOGINDETAILS & "'" & vbNewLine & _
          "      ,'" & strEncrypted & "'" & vbNewLine & _
          "      ,'" & gsPARAMETERTYPE_ENCYPTED & "')"
    
    gADOCon.Execute sSQL, adExecuteNoRecords
    
    SaveSystemSetting "ProcessAccount", "Mode", glngProcessMethod
    
  Else
    DecryptLogonDetails strEncrypted, sName, sPassword, sDatabase, sServer
    
    If pbTrusted Then
      strEncrypted = EncryptLogonDetails("", "", gsDatabaseName, gsServerName)
      glngProcessMethod = iPROCESSADMIN_SERVICEACCOUNT
    ElseIf psName = vbNullString And psPassword = vbNullString Then
      strEncrypted = EncryptLogonDetails(sName, sPassword, gsDatabaseName, gsServerName)
      
      If Trim$(sName) = vbNullString Then
        glngProcessMethod = iPROCESSADMIN_SERVICEACCOUNT
      Else
        glngProcessMethod = iPROCESSADMIN_SQLACCOUNT
      End If
    Else
      strEncrypted = EncryptLogonDetails(psName, psPassword, gsDatabaseName, gsServerName)
      glngProcessMethod = iPROCESSADMIN_SQLACCOUNT
    End If
    
    sSQL = _
          "UPDATE ASRSysModuleSetup" & vbNewLine & _
          "     SET [ParameterValue] = '" & strEncrypted & "'" & vbNewLine & _
          "     WHERE [ModuleKey] = '" & gsMODULEKEY_SQL & "'" & vbNewLine & _
          "     AND   [ParameterKey] = '" & gsPARAMETERKEY_LOGINDETAILS & "'"
          
    gADOCon.Execute sSQL, adExecuteNoRecords
    
    SaveSystemSetting "ProcessAccount", "Mode", glngProcessMethod
    
  End If
  
  bOK = True

TidyUpAndExit:
  RegenerateSQLProcessAccount = bOK
  Exit Function

LocalErr:
  bOK = False
  GoTo TidyUpAndExit

End Function


Public Function RemoveWorkflowLoginCredentials() As Boolean
 
  Dim sSQL As String
  Dim strInput As String
  Dim strOutput As String
  Dim strEKey As String
  Dim strLens As String
  Dim lngCount As Long
  Dim iChar As Integer
  Dim bOK As Boolean
  Dim strUserName As String
  Dim strPassword As String
  Dim lngStart As Long
  Dim lngFinish As Long
  
  Const dummyUsername = "XX"
  Const dummyPassword = "XX"
  
  ' Read existing username/password - only reset the credentials if required.
  strInput = GetWebLogon()
    
  If strInput <> vbNullString Then

    lngStart = Len(strInput) - 12
    strEKey = Mid(strInput, lngStart + 1, 10)
    strLens = Right(strInput, 2)
    strInput = XOREncript(Left(strInput, lngStart), strEKey)
  
    lngStart = 1
    lngFinish = Asc(Mid(strLens, 1, 1)) - 127
    strUserName = Mid(strInput, lngStart, lngFinish)
  
    lngStart = lngStart + lngFinish
    lngFinish = Asc(Mid(strLens, 2, 1)) - 127
    strPassword = Mid(strInput, lngStart, lngFinish)
  
    If strUserName = dummyUsername And strPassword = dummyPassword Then
      gfWFCredentialsChanged = False
      GoTo TidyUpAndExit
    End If
  
  End If
  
  ' NPG: I've blocked out the username and password instead of removing them,
  ' because Len(sUser) is used throughout this and other modules to validate
  ' workflow credentials.
  strInput = vbNullString
  strOutput = dummyUsername & dummyPassword
  strLens = Chr(Len(dummyUsername) + 127) & Chr(Len(dummyPassword) + 127)
  
  ' AE20080229 Fault #12939, #12959
  Do While strInput = vbNullString _
    Or (CBool(InStr(strInput, Chr(0))) Or CBool(InStr(strInput, Chr(144))))
    
    strInput = vbNullString
    strEKey = vbNullString
    For lngCount = 1 To 10
      'strEKey = strEKey & Chr(Int(Rnd * 255) + 1)
      iChar = 0
      iChar = Int(Rnd * 255) + 1
      strEKey = strEKey & Chr(iChar)
    Next

    strInput = XOREncript(strOutput, strEKey) & strEKey & strLens
  Loop
  strOutput = strInput
  
  ' AE20081003 Fault #13387
  sSQL = _
        "DELETE FROM [ASRSysModuleSetup]" & vbNewLine & _
        "WHERE  [ModuleKey] = '" & gsMODULEKEY_WORKFLOW & "'" & vbNewLine & _
        "AND    [ParameterKey] = '" & gsPARAMETERKEY_WEBPARAM1 & "'" & vbNewLine
      
  sSQL = sSQL & vbNewLine & _
        "INSERT INTO [ASRSysModuleSetup]" & vbNewLine & _
        "     ([ModuleKey]" & vbNewLine & _
        "     ,[ParameterKey]" & vbNewLine & _
        "     ,[ParameterValue]" & vbNewLine & _
        "     ,[ParameterType])" & vbNewLine & _
        "VALUES (" & vbNewLine & _
        "       '" & gsMODULEKEY_WORKFLOW & "'" & vbNewLine & _
        "      ,'" & gsPARAMETERKEY_WEBPARAM1 & "'" & vbNewLine & _
        "      ,'" & strOutput & "'" & vbNewLine & _
        "      ,'" & gsPARAMETERTYPE_ENCYPTED & "')"
  
  gADOCon.Execute sSQL, adExecuteNoRecords
      
  bOK = True
  gfWFCredentialsChanged = True

TidyUpAndExit:
  RemoveWorkflowLoginCredentials = bOK
  Exit Function

LocalErr:
  bOK = False
  GoTo TidyUpAndExit

End Function



Public Function GetSystemLogon() As String

  Dim rsLogon As ADODB.Recordset
  Dim sSQL As String
  
  sSQL = "SELECT [ParameterValue] FROM ASRSysModuleSetup WHERE [ModuleKey] = 'MODULE_SQL'" & _
                      " AND [ParameterKey] = 'Param_FieldsLoginDetails'"
        
  Set rsLogon = New ADODB.Recordset
  rsLogon.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  
  If Not (rsLogon.BOF And rsLogon.EOF) Then
    GetSystemLogon = IIf(IsNull(rsLogon!parametervalue), vbNullString, rsLogon!parametervalue)
  End If

  rsLogon.Close
  Set rsLogon = Nothing
  
End Function

Public Function GetWebLogon() As String

  Dim rsLogon As ADODB.Recordset
  Dim sSQL As String
  
  sSQL = "SELECT [ParameterValue] FROM ASRSysModuleSetup WHERE [ModuleKey] =  'MODULE_WORKFLOW'" & _
                      " AND [ParameterKey] = 'Param_Web1'"
        
  Set rsLogon = New ADODB.Recordset
  rsLogon.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  
  If Not (rsLogon.BOF And rsLogon.EOF) Then
    GetWebLogon = IIf(IsNull(rsLogon!parametervalue), vbNullString, rsLogon!parametervalue)
  End If

  rsLogon.Close
  Set rsLogon = Nothing
  
End Function

Public Function TestSystemLogon() As Boolean

  Dim bOK As Boolean
  Dim strEncrypted As String
  Dim rstTest As ADODB.Recordset

  On Error GoTo LocalErr

  If glngSQLVersion < 9 Then
    TestSystemLogon = True
    Exit Function
  End If
  
  strEncrypted = GetSystemLogon
  
  ' Test the encrypted logon
  Set rstTest = New ADODB.Recordset
  rstTest.Open "SELECT dbo.udfASRNetIsProcessValid('" & Replace(strEncrypted, "'", "''") & "')", gADOCon, adOpenForwardOnly, adLockReadOnly
  
  If Not (rstTest.BOF And rstTest.EOF) Then
    bOK = (rstTest.Fields(0).value = True)
  End If
  
  rstTest.Close
  Set rstTest = Nothing
  
TidyUpAndExit:
  
  TestSystemLogon = bOK
  Exit Function

LocalErr:
  bOK = False
  GoTo TidyUpAndExit
 
End Function
