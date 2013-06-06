Attribute VB_Name = "modMobileSpecifics"
Option Explicit

Private Const msMobileCheckLogin_PROCEDURENAME = "spASRSysMobileCheckLogin"
Private Const msMobileRegistration_PROCEDURENAME = "spASRSysMobileRegistration"
Private Const msMobileCheckPendingWorkflowSteps_PROCEDURENAME = "spASRSysMobileCheckPendingWorkflowSteps"

Private mvar_fGeneralOK As Boolean
Private mvar_sGeneralMsg As String

Private mvar_sLoginColumn As String
Private mvar_sLoginTable As String
Private mvar_sUniqueEmailColumn As String
Private mvar_sLeavingDateColumn As String
Private mvar_lngUniqueEmailColumn As Long
Private mvar_lngLeavingDateColumn As Long


Public Sub DropMobileObjects()
  DropProcedure msMobileCheckLogin_PROCEDURENAME
  DropProcedure msMobileRegistration_PROCEDURENAME
  DropProcedure msMobileCheckPendingWorkflowSteps_PROCEDURENAME
End Sub



Public Function ConfigureMobileSpecifics() As Boolean
  ' Configure module specific objects (eg. stored procedures)
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sErrorMessage As String
  
  fOK = True
  
  mvar_fGeneralOK = True
  mvar_sGeneralMsg = ""
    
  If fOK Then
    ' Read the Mobile parameters.
    fOK = ReadMobileParameters
    If Not fOK Then
      mvar_fGeneralOK = False
      sErrorMessage = "Mobile specifics not correctly configured." & vbNewLine & _
        "Some functionality will be disabled if you do not change your configuration." & vbNewLine & mvar_sGeneralMsg
      
      fOK = (OutputMessage(sErrorMessage & vbNewLine & vbNewLine & "Continue saving changes ?") = vbYes)
    End If
  End If
  
  'Make sure that we drop the Mobile SPs
  DropMobileObjects
  
  
  ' Create the CheckLogin stored procedures.
  If fOK And mvar_fGeneralOK Then
    fOK = CreateSP_MobileCheckLogin
    If Not fOK Then
      DropProcedure msMobileCheckLogin_PROCEDURENAME
    End If
  End If
  
  ' Create the Mobile Registration stored procedures.
  If fOK And mvar_fGeneralOK Then
    fOK = CreateSP_MobileRegistration
    If Not fOK Then
      DropProcedure msMobileRegistration_PROCEDURENAME
    End If
  End If
  
  
  ' Create the Mobile Check Workflow Pending Steps stored procedures.
  If fOK And mvar_fGeneralOK Then
    fOK = CreateSP_MobileCheckPendingWorkflowSteps
    If Not fOK Then
      DropProcedure msMobileCheckPendingWorkflowSteps_PROCEDURENAME
    End If
  End If
    
TidyUpAndExit:
  ConfigureMobileSpecifics = fOK
  Exit Function
  
ErrorTrap:
  OutputError "Error configuring Mobile specifics"
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function ReadMobileParameters() As Boolean
  ' Read the configured Mobile parameters into member variables.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim lngLoginColumn As Long
  Dim lngLoginTable As Long
  Dim lngColumnID As Long
  Dim lngTableID As Long
  Dim sUser As String
  Dim sPassword As String
      
  fOK = True
  
  With recModuleSetup
    .Index = "idxModuleParameter"
    
    If fOK Then
      ' Get the login column
      lngLoginColumn = GetModuleSetting(gsMODULEKEY_MOBILE, gsPARAMETERKEY_LOGINNAME, 0)
   
      fOK = lngLoginColumn > 0
      If Not fOK Then
        mvar_sGeneralMsg = mvar_sGeneralMsg & vbCrLf & "  Login column not defined."
      End If
      
    End If
    
    If fOK Then
      lngLoginTable = GetTableIDFromColumnID(lngLoginColumn)
      mvar_sLoginColumn = GetColumnName(lngLoginColumn, True)
      mvar_sLoginTable = GetTableName(lngLoginTable)
   
      ' Get the Unique Email column.
      .Seek "=", gsMODULEKEY_MOBILE, gsPARAMETERKEY_UNIQUEEMAILCOLUMN
      If .NoMatch Then
        mvar_lngUniqueEmailColumn = 0
      Else
        mvar_lngUniqueEmailColumn = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
        mvar_sUniqueEmailColumn = GetColumnName(mvar_lngUniqueEmailColumn, True)
      End If
    
      ' Get the Leaving Date column.
      .Seek "=", gsMODULEKEY_MOBILE, gsPARAMETERKEY_LEAVINGDATE
      If .NoMatch Then
        mvar_lngLeavingDateColumn = 0
      Else
        mvar_lngLeavingDateColumn = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
        mvar_sLeavingDateColumn = GetColumnName(mvar_lngLeavingDateColumn, True)
      End If
        
    End If
    
  End With

TidyUpAndExit:
  ReadMobileParameters = fOK
  Exit Function
  
ErrorTrap:
  OutputError "Error reading mobile parameters"
  fOK = False
  Resume TidyUpAndExit
  
End Function



Private Function CreateSP_MobileCheckLogin() As Boolean
  ' Create the Check Login stored procedure.
  On Error GoTo ErrorTrap

  Dim fCreatedOK As Boolean
  Dim sProcSQL As String
  Dim iCount As Integer
  
  fCreatedOK = True

  ' Construct the stored procedure creation string.
  sProcSQL = "/* ------------------------------------------------ */" & vbNewLine & _
    "/* Mobile module stored procedure.         */" & vbNewLine & _
    "/* Automatically generated by the System manager.   */" & vbNewLine & _
    "/* ------------------------------------------------ */" & vbNewLine & _
    "CREATE PROCEDURE [dbo].[" & msMobileCheckLogin_PROCEDURENAME & "](" & vbNewLine & _
    "  @psKeyParameter varchar(max)," & vbNewLine & _
    "  @psPWDParameter nvarchar(max) OUTPUT," & vbNewLine & _
    "  @psMessage varchar(max) OUTPUT" & vbNewLine & _
    "  ) " & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & _
    "  DECLARE @piuserID int" & vbNewLine & _
    "  SET @piuserID = 0;" & vbNewLine & _
    "  SET @psMessage = '';" & vbNewLine & _
    "  -- get the record id for this user" & vbNewLine & _
    "  SELECT @piuserID = [ID]" & vbNewLine & _
    "    From [" & mvar_sLoginTable & "]" & vbNewLine & _
    "    WHERE ISNULL([" & mvar_sLoginColumn & "], '') = @psKeyParameter" & vbNewLine & _
    "  IF @piuserID > 0" & vbNewLine & _
    "  BEGIN" & vbNewLine & _
    "    -- return the password for the record id in its encrypted form." & vbNewLine & _
    "    SELECT TOP 1 @psPWDParameter = [password]" & vbNewLine & _
    "      From [tbsys_mobilelogins]" & vbNewLine & _
    "      WHERE (ISNULL([tbsys_mobilelogins].[userid], '') = @piuserID);" & vbNewLine & _
    "  End" & vbNewLine
  sProcSQL = sProcSQL & _
    "  IF @piuserID = 0" & vbNewLine & _
    "  BEGIN" & vbNewLine & _
    "    SET @psMessage = 'Account not found.';" & vbNewLine & _
    "  End" & vbNewLine & _
    "  IF @psMessage='' AND ISNULL(@psPWDParameter, '')  = ''" & vbNewLine & _
    "    BEGIN" & vbNewLine & _
    "      SET @psMessage = 'Account not activated.';" & vbNewLine & _
    "    END" & vbNewLine & _
    "END"

  gADOCon.Execute sProcSQL, , adExecuteNoRecords

TidyUpAndExit:
  CreateSP_MobileCheckLogin = fCreatedOK
  Exit Function

ErrorTrap:
  fCreatedOK = False
  OutputError "Error creating Check Mobile Login stored procedure (Mobile)"
  Resume TidyUpAndExit

End Function


Private Function CreateSP_MobileRegistration() As Boolean
  ' Create the Mobile Registration stored procedure.
  On Error GoTo ErrorTrap

  Dim fCreatedOK As Boolean
  Dim sProcSQL As String
  Dim iCount As Integer
  
  fCreatedOK = True

  ' Construct the stored procedure creation string.
  sProcSQL = "/* ------------------------------------------------ */" & vbNewLine & _
    "/* Mobile module stored procedure.         */" & vbNewLine & _
    "/* Automatically generated by the System manager.   */" & vbNewLine & _
    "/* ------------------------------------------------ */" & vbNewLine & _
    "create PROCEDURE [dbo].[" & msMobileRegistration_PROCEDURENAME & "](" & vbNewLine & _
    "  @psKeyParameter varchar(max)," & vbNewLine & _
    "  @psPWDParameter nvarchar(max)," & vbNewLine & _
    "  @psEmailAddress varchar(max)," & vbNewLine & _
    "  @psMessage varchar(max) OUTPUT" & vbNewLine & _
    ")" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine & vbNewLine & _
    "  DECLARE @iCount integer," & vbNewLine & _
    "  @iUserRecordID int," & vbNewLine & _
    "  @sMessage varchar(max);" & vbNewLine & _
    "  SET @iCount = 0;" & vbNewLine & _
    "  SET @psMessage = '';" & vbNewLine & _
    "  SELECT @iCount = COUNT([" & mvar_sLoginColumn & "])" & vbNewLine & _
    "    FROM " & mvar_sLoginTable & vbNewLine & _
    "    WHERE ISNULL(" & mvar_sLoginTable & "." & mvar_sUniqueEmailColumn & ", '') = @psEmailAddress;" & vbNewLine & _
    "  IF @iCount = 0" & vbNewLine & _
    "  BEGIN" & vbNewLine & _
    "    SET @psMessage = 'No records exist with the given email address.';" & vbNewLine & _
    "  END;" & vbNewLine
    
  sProcSQL = sProcSQL & "  IF @iCount > 1" & vbNewLine & _
    "  BEGIN" & vbNewLine & _
    "    SET @psMessage = 'More than 1 record exists with the given email address.';" & vbNewLine & _
    "  END;" & vbNewLine & _
    "  IF @psMessage = '' and @iCount = 1" & vbNewLine & _
    "  BEGIN" & vbNewLine & _
    "    --CHECK LOGINS TABLE NOW" & vbNewLine & _
    "    SET @iUserRecordID = 0;" & vbNewLine & _
    "    SELECT @iUserRecordID = [ID]" & vbNewLine & _
    "    FROM " & mvar_sLoginTable & vbNewLine & _
    "      WHERE ISNULL(" & mvar_sLoginTable & "." & mvar_sUniqueEmailColumn & ", '') = @psEmailAddress;" & vbNewLine & _
    "      SELECT @iCount = COUNT([userid])" & vbNewLine & _
    "        FROM tbsys_mobilelogins" & vbNewLine & _
    "        WHERE ISNULL(tbsys_mobilelogins.[userid], 0) = @iUserRecordID;" & vbNewLine & _
    "        IF @iCount <> 0" & vbNewLine & _
    "        BEGIN" & vbNewLine & _
    "          SET @psMessage = 'This user has already been registered. Use the ''forgot password'' screen to reset your password.';" & vbNewLine & _
    "        END;" & vbNewLine
        
  sProcSQL = sProcSQL & "        ELSE" & vbNewLine & _
    "        BEGIN" & vbNewLine & _
    "          UPDATE " & mvar_sLoginTable & vbNewLine & _
    "            SET " & mvar_sLoginColumn & " = @psKeyParameter" & vbNewLine & _
    "            WHERE ISNULL(" & mvar_sLoginTable & "." & mvar_sUniqueEmailColumn & ", '') = @psEmailAddress;" & vbNewLine & _
    "            -- Add the user to the logins table if required." & vbNewLine & _
    "          INSERT [dbo].[tbsys_mobilelogins](" & vbNewLine & _
    "             userid," & vbNewLine & _
    "             newpassword)" & vbNewLine & _
    "          VALUES (@iUserRecordID," & vbNewLine & _
    "             @psPWDParameter);" & vbNewLine & _
    "          SET @sMessage = 'You are now registered to use OpenHR Mobile.'    -- Your login is''' + @psKeyParameter + ''', and your password is ''' + @psPWDParameter + '''.';" & vbNewLine & _
    "          INSERT [dbo].[ASRSysEmailQueue](" & vbNewLine & _
    "             RecordDesc," & vbNewLine & _
    "             ColumnValue," & vbNewLine & _
    "             DateDue," & vbNewLine & _
    "             UserName," & vbNewLine & _
    "             [Immediate]," & vbNewLine & _
    "             RecalculateRecordDesc," & vbNewLine & _
    "             RepTo," & vbNewLine & _
    "             MsgText," & vbNewLine & _
    "             WorkflowInstanceID," & vbNewLine & _
    "             [Subject])" & vbNewLine & _
    "          VALUES (''," & vbNewLine
        
  sProcSQL = sProcSQL & "             ''," & vbNewLine & _
    "             getdate()," & vbNewLine & _
    "             'OpenHR  Mobile'," & vbNewLine & _
    "             1," & vbNewLine & _
    "             0," & vbNewLine & _
    "             @psEmailAddress," & vbNewLine & _
    "             @sMessage," & vbNewLine & _
    "             0," & vbNewLine & _
    "             'OpenHR Mobile registration details');" & vbNewLine & _
    "        END;" & vbNewLine & _
    "  END;" & vbNewLine & _
    "END;"
    

  gADOCon.Execute sProcSQL, , adExecuteNoRecords

TidyUpAndExit:
  CreateSP_MobileRegistration = fCreatedOK
  Exit Function

ErrorTrap:
  fCreatedOK = False
  OutputError "Error creating Mobile Registration stored procedure (Mobile)"
  Resume TidyUpAndExit

End Function


Private Function CreateSP_MobileCheckPendingWorkflowSteps() As Boolean

  ' Create the Check Login stored procedure.
  On Error GoTo ErrorTrap

  Dim fCreatedOK As Boolean
  Dim sProcSQL As String
  Dim iCount As Integer
  
  fCreatedOK = True

  ' Construct the stored procedure creation string.
  sProcSQL = "/* ------------------------------------------------ */" & vbNewLine & _
    "/* Mobile module stored procedure.         */" & vbNewLine & _
    "/* Automatically generated by the System manager.   */" & vbNewLine & _
    "/* ------------------------------------------------ */" & vbNewLine & _
    "CREATE PROCEDURE [dbo].[" & msMobileCheckPendingWorkflowSteps_PROCEDURENAME & "](" & vbNewLine & _
    "  @psKeyParameter varchar(max)" & vbNewLine & _
    "  ) " & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine & _
"SET NOCOUNT ON;" & vbNewLine & _
"DECLARE" & vbNewLine & _
"@sURL varchar(MAX)," & vbNewLine & _
"@sDescription varchar(MAX)," & vbNewLine & _
"@sCalcDescription varchar(MAX)," & vbNewLine & _
"@iInstanceID integer," & vbNewLine & _
"@iInstanceStepID integer," & vbNewLine & _
"@iElementID integer," & vbNewLine & _
"@hResult integer," & vbNewLine & _
"@objectToken integer," & vbNewLine & _
"@sQueryString varchar(MAX)," & vbNewLine & _
"@sParam1  varchar(MAX)," & vbNewLine & _
"@sServerName sysname," & vbNewLine & _
"@sDBName  sysname," & vbNewLine

sProcSQL = sProcSQL & "" & _
"@sSQLVersion  int;" & vbNewLine & _
"DECLARE @steps TABLE" & vbNewLine & _
"(" & vbNewLine & _
"[description] varchar(MAX)," & vbNewLine & _
"[URL] varChar(MAX)," & vbNewLine & _
"[instanceID] integer," & vbNewLine & _
"[elementID] integer," & vbNewLine & _
"[instanceStepID] integer" & vbNewLine & _
")" & vbNewLine & _
"SELECT @sURL = parameterValue" & vbNewLine & _
"FROM ASRSysModuleSetup" & vbNewLine & _
"WHERE moduleKey = 'MODULE_WORKFLOW'" & vbNewLine & _
"AND parameterKey = 'Param_URL'" & vbNewLine & _
"IF upper(right(@sURL, 5)) <> '.ASPX'" & vbNewLine & _
"AND right(@sURL, 1) <> '/'" & vbNewLine

sProcSQL = sProcSQL & "" & _
"AND len(@sURL) > 0" & vbNewLine & _
"BEGIN" & vbNewLine & _
"SET @sURL = @sURL + '/'" & vbNewLine & _
"END" & vbNewLine & _
"SELECT @sParam1 = parameterValue" & vbNewLine & _
"FROM ASRSysModuleSetup" & vbNewLine & _
"WHERE moduleKey = 'MODULE_WORKFLOW'" & vbNewLine & _
"AND parameterKey = 'Param_Web1'" & vbNewLine & _
"SET @sServerName = CONVERT(sysname,SERVERPROPERTY('servername'))" & vbNewLine & _
"SET @sDBName = db_name()" & vbNewLine & _
"SET @sSQLVersion = dbo.udfASRSQLVersion()" & vbNewLine & _
"IF @sSQLVersion <= 8" & vbNewLine & _
"EXEC @hResult = sp_OACreate 'vbpHRProServer.clsWorkflow', @objectToken OUTPUT" & vbNewLine & _
"IF (@hResult = 0 OR @sSQLVersion > 8) AND (len(@sURL) > 0)" & vbNewLine & _
"BEGIN" & vbNewLine

sProcSQL = sProcSQL & "" & _
"DECLARE @sEmailAddress_1 varchar(MAX)" & vbNewLine & _
"SELECT @sEmailAddress_1 = replace(upper(ltrim(rtrim(Personnel_Records.Email_Work))), ' ', '')" & vbNewLine & _
"FROM Personnel_Records" & vbNewLine & _
"WHERE (ISNULL(Personnel_Records.Self_Service_Login, '') = @psKeyParameter)" & vbNewLine & _
"AND len(Personnel_Records.Email_Work) > 0" & vbNewLine & _
"print @sEmailAddress_1;" & vbNewLine & _
"DECLARE steps_cursor CURSOR LOCAL FAST_FORWARD FOR" & vbNewLine & _
"SELECT ASRSysWorkflowInstanceSteps.instanceID," & vbNewLine & _
"ASRSysWorkflowInstanceSteps.elementID," & vbNewLine & _
"ASRSysWorkflowInstanceSteps.ID," & vbNewLine & _
"ASRSysWorkflows.name + ' - ' + ASRSysWorkflowElements.caption AS [description]" & vbNewLine & _
"FROM ASRSysWorkflowInstanceSteps" & vbNewLine & _
"INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID" & vbNewLine & _
"INNER JOIN ASRSysWorkflows ON ASRSysWorkflowElements.workflowID = ASRSysWorkflows.ID" & vbNewLine & _
"WHERE (ASRSysWorkflowInstanceSteps.Status = 2" & vbNewLine

sProcSQL = sProcSQL & "" & _
"OR ASRSysWorkflowInstanceSteps.Status = 7)" & vbNewLine & _
"AND (ASRSysWorkflowInstanceSteps.userName = @psKeyParameter --SUSER_SNAME()" & vbNewLine & _
"OR (';' + replace(upper(ASRSysWorkflowInstanceSteps.userEmail), ' ', '') + ';' LIKE '%;' + @sEmailAddress_1 + ';%'" & vbNewLine & _
"AND len(@sEmailAddress_1) > 0)" & vbNewLine & _
"OR ((len(@sEmailAddress_1) > 0)" & vbNewLine & _
"AND ((SELECT COUNT(*)" & vbNewLine & _
"FROM ASRSysWorkflowStepDelegation" & vbNewLine & _
"WHERE stepID = ASRSysWorkflowInstanceSteps.ID" & vbNewLine & _
"AND ';' + replace(upper(ASRSysWorkflowStepDelegation.delegateEmail), ' ', '') + ';' LIKE '%;' + @sEmailAddress_1 + ';%') > 0)))" & vbNewLine & _
"OPEN steps_cursor" & vbNewLine & _
"FETCH NEXT FROM steps_cursor INTO @iInstanceID, @iElementID, @iInstanceStepID, @sDescription" & vbNewLine & _
"WHILE (@@fetch_status = 0)" & vbNewLine & _
"BEGIN" & vbNewLine & _
"SET @sQueryString = ''" & vbNewLine & _
"IF @sSQLVersion <=8" & vbNewLine

sProcSQL = sProcSQL & "" & _
"BEGIN" & vbNewLine & _
"EXEC @hResult = sp_OAMethod @objectToken, 'GetQueryString', @sQueryString OUTPUT, @iInstanceID, @iElementID, @sParam1, @sServerName, @sDBName" & vbNewLine & _
"IF @hResult <> 0" & vbNewLine & _
"BEGIN" & vbNewLine & _
"SET @sQueryString = ''" & vbNewLine & _
"END            END" & vbNewLine & _
"ELSE" & vbNewLine & _
"SELECT @sQueryString = dbo.[udfASRNetGetWorkflowQueryString]( @iInstanceID, @iElementID, @sParam1, @sServerName, @sDBName)" & vbNewLine & _
"IF len(@sQueryString) > 0" & vbNewLine & _
"BEGIN" & vbNewLine & _
"EXEC [dbo].[spASRWorkflowStepDescription]" & vbNewLine & _
"@iInstanceStepID," & vbNewLine & _
"@sCalcDescription OUTPUT" & vbNewLine & _
"IF len(@sCalcDescription) > 0 " & vbNewLine & _
"BEGIN" & vbNewLine

sProcSQL = sProcSQL & "" & _
"SET @sDescription = @sCalcDescription" & vbNewLine & _
"END" & vbNewLine & _
"INSERT INTO @steps ([description], [url], [instanceID], [elementID], [instanceStepID])" & vbNewLine & _
"VALUES (@sDescription, @sURL + '/?' + @sQueryString, @iInstanceID, @iElementID, @iInstanceStepID)" & vbNewLine & _
"END" & vbNewLine & _
"FETCH NEXT FROM steps_cursor INTO @iInstanceID, @iElementID, @iInstanceStepID, @sDescription" & vbNewLine & _
"END" & vbNewLine & _
"CLOSE steps_cursor" & vbNewLine & _
"DEALLOCATE steps_cursor" & vbNewLine & _
"IF @sSQLVersion <= 8" & vbNewLine & _
"EXEC sp_OADestroy @objectToken" & vbNewLine & _
"END" & vbNewLine & _
"SELECT *" & vbNewLine & _
"FROM @steps" & vbNewLine & _
"ORDER BY [description]" & vbNewLine & _
"END" & vbNewLine

  gADOCon.Execute sProcSQL, , adExecuteNoRecords

TidyUpAndExit:
  CreateSP_MobileCheckPendingWorkflowSteps = fCreatedOK
  Exit Function

ErrorTrap:
  fCreatedOK = False
  OutputError "Error creating Mobile Registration stored procedure (Mobile)"
  Resume TidyUpAndExit




End Function
