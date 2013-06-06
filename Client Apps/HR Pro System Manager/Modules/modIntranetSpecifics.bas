Attribute VB_Name = "modIntranetSpecifics"
Option Explicit

Private Const msResetPassword_PROCEDURENAME = "spadmin_resetpassword"
Private Const msWorkEMailColumnNotDefined = "'Work email' column not defined."

Private mvar_fGeneralOK As Boolean
Private mvar_sGeneralMsg As String

Private mvar_sLoginColumn As String
Private mvar_sLoginTable As String
Private mvar_sWorkEmailColumn As String
Private mvar_sLeavingDateColumn As String
Private mvar_sActivatedUserColumn As String
Private mvar_lngWorkEmailColumn As Long
Private mvar_lngLeavingDateColumn As Long
Private mvar_lngActivatedUserColumn As Long

Public Sub DropIntranetObjects()
  DropProcedure msResetPassword_PROCEDURENAME
End Sub

Public Function ConfigureIntranetSpecifics() As Boolean
  ' Configure module specific objects (eg. stored procedures)
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sErrorMessage As String
  Dim sTemp As String
  
  fOK = True
  
  mvar_fGeneralOK = True
  mvar_sGeneralMsg = ""
 
  ' Read the Intranet parameters.
  fOK = ReadIntranetParameters
  If Not fOK Then
    mvar_fGeneralOK = False
    
    If mvar_sGeneralMsg Like "*" & msWorkEMailColumnNotDefined & "*" Then
      sTemp = "The Forgot Password"
    Else
      sTemp = "Some"
    End If
    
    sErrorMessage = "Intranet specifics not correctly configured in Personnel Module Setup." & vbNewLine & _
      sTemp & " functionality will be disabled if you do not change your configuration." & vbNewLine & mvar_sGeneralMsg
    
    fOK = (OutputMessage(sErrorMessage & vbNewLine & vbNewLine & "Continue saving changes ?") = vbYes)
  End If
  
  
  'Make sure that we drop the Intranet SPs
  DropIntranetObjects
  
  ' Create the ResetPassword stored procedure.
  If fOK And mvar_fGeneralOK Then
    fOK = CreateSP_ResetPassword
    If Not fOK Then
      DropProcedure msResetPassword_PROCEDURENAME
    End If
  End If
  
TidyUpAndExit:
  ConfigureIntranetSpecifics = fOK
  Exit Function
  
ErrorTrap:
  OutputError "Error configuring Intranet specifics"
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function ReadIntranetParameters() As Boolean
  ' Read the configured Intranet parameters into member variables.
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
      lngLoginColumn = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_LOGINNAME, 0)
   
      fOK = lngLoginColumn > 0
      If Not fOK Then
        mvar_sGeneralMsg = mvar_sGeneralMsg & vbCrLf & "  'Login name' column not defined in Personnel Module Setup."
      End If
      
    End If
    
    If fOK Then
      lngLoginTable = GetTableIDFromColumnID(lngLoginColumn)
      mvar_sLoginColumn = GetColumnName(lngLoginColumn, True)
      mvar_sLoginTable = GetTableName(lngLoginTable)
    End If
    
    
    If fOK Then
          ' Get the Unique Email column.
      .Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_WORKEMAIL
      If .NoMatch Then
        mvar_lngWorkEmailColumn = 0
      Else
        mvar_lngWorkEmailColumn = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
        mvar_sWorkEmailColumn = GetColumnName(mvar_lngWorkEmailColumn, True)
      End If
      
      fOK = (mvar_lngWorkEmailColumn > 0)
      If Not fOK Then mvar_sGeneralMsg = mvar_sGeneralMsg & vbCrLf & "  " & msWorkEMailColumnNotDefined
    End If
    
    If fOK Then
      ' Get the Intranet Activated User column.
      .Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_LOGINNAME
      If .NoMatch Then
        mvar_lngActivatedUserColumn = 0
        
      Else
        mvar_lngActivatedUserColumn = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
        mvar_sActivatedUserColumn = GetColumnName(mvar_lngActivatedUserColumn, True)
      End If
      
      fOK = (mvar_lngActivatedUserColumn > 0)
      If Not fOK Then mvar_sGeneralMsg = mvar_sGeneralMsg & vbCrLf & "  'Login name' column not defined."
    End If
    
  End With

TidyUpAndExit:
  ReadIntranetParameters = fOK
  Exit Function
  
ErrorTrap:
  OutputError "Error reading Intranet parameters"
  fOK = False
  Resume TidyUpAndExit
  
End Function



Private Function CreateSP_ResetPassword() As Boolean
  ' Create the Check Login stored procedure.
  On Error GoTo ErrorTrap

  Dim fCreatedOK As Boolean
  Dim sProcSQL As String
  Dim iCount As Integer
  
  fCreatedOK = True

  ' Construct the stored procedure creation string.
  ' NB SP must already exist - created by update script v5.1
  
  sProcSQL = "/* ------------------------------------------------ */" & vbNewLine & _
    "/* Intranet module stored procedure.         */" & vbNewLine & _
    "/* Automatically generated by the System manager.   */" & vbNewLine & _
    "/* ------------------------------------------------ */" & vbNewLine & _
    "CREATE PROCEDURE dbo.[" & msResetPassword_PROCEDURENAME & "](" & vbNewLine & _
    "  @psWebsiteURL VARCHAR(255)," & vbNewLine & _
    "  @psUserName VARCHAR(255)," & vbNewLine & _
    "  @psEncryptedLink VARCHAR(MAX)," & vbNewLine & _
    "  @psMessage VARCHAR(MAX) OUTPUT" & vbNewLine & _
    "  )" & vbNewLine & _
    "WITH EXECUTE AS 'dbo'" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine
    
  sProcSQL = sProcSQL & _
    "    DECLARE @iCount INTEGER," & vbNewLine & _
    "          @psEmailAddress VARCHAR(MAX)," & vbNewLine & _
    "          @dtExpiryDate DATETIME," & vbNewLine & _
    "          @sMessage VARCHAR(MAX);" & vbNewLine & vbNewLine
    
  sProcSQL = sProcSQL & _
    "  SET @iCount = 0;" & vbNewLine & _
    "  SET @psMessage = '';" & vbNewLine & _
    "" & vbNewLine & _
    "  SELECT @iCount = COUNT([" & mvar_sActivatedUserColumn & "])" & vbNewLine & _
    "    FROM " & mvar_sLoginTable & "" & vbNewLine & _
    "    WHERE ISNULL(" & mvar_sLoginTable & ".[" & mvar_sActivatedUserColumn & "], '') = @psUserName;" & vbNewLine & _
    "" & vbNewLine & _
    "  IF @iCount = 0" & vbNewLine & _
    "  BEGIN" & vbNewLine & _
    "    SET @psMessage = 'No records exist with the given user name.';" & vbNewLine & _
    "  END;" & vbNewLine & _
    "" & vbNewLine & _
    "  IF @iCount > 1" & vbNewLine & _
    "  BEGIN" & vbNewLine & _
    "    SET @psMessage = 'More than 1 record exists with the given user name.';" & vbNewLine & _
    "  END;" & vbNewLine & _
    "" & vbNewLine & _
    "  IF @iCount = 1" & vbNewLine & _
    "  BEGIN" & vbNewLine
    
  sProcSQL = sProcSQL & "    SELECT @psEmailAddress = ISNULL(" & mvar_sWorkEmailColumn & ", '')" & vbNewLine & _
    "    From " & mvar_sLoginTable & "" & vbNewLine & _
    "    WHERE ISNULL(" & mvar_sLoginTable & "." & mvar_sActivatedUserColumn & ", '') = @psUserName;" & vbNewLine & _
    "" & vbNewLine & _
    "    IF (LEN(@psEmailAddress) = 0)" & vbNewLine & _
    "    BEGIN" & vbNewLine & _
    "      SET @psMessage = 'No e-mail address exists for the given user name.';" & vbNewLine & _
    "    END" & vbNewLine & _
    "    ELSE" & vbNewLine & _
    "    BEGIN" & vbNewLine & _
    "      SET @sMessage = 'To reset your password, copy the link shown below into your browser address bar. This will take you to a web page where you can create a new password.' + CHAR(13) + CHAR(10) +" & vbNewLine & _
    "            'If you weren''t trying to reset your password, don''t worry � your account is still secure and no one has been given access to it.' + CHAR(13) + CHAR(10) + CHAR(13) + CHAR(10) +" & vbNewLine & _
    "            'Copy the link shown below into your web browser to reset your password:' + CHAR(13) + CHAR(10) +" & vbNewLine & _
    "            @psWebsiteURL + '?' + @psEncryptedLink;" & vbNewLine & vbNewLine
    
  sProcSQL = sProcSQL & "      INSERT [dbo].[ASRSysEmailQueue](" & vbNewLine & _
    "        RecordDesc," & vbNewLine & _
    "        ColumnValue," & vbNewLine & _
    "        DateDue," & vbNewLine & _
    "        UserName," & vbNewLine & _
    "        [Immediate]," & vbNewLine & _
    "        RecalculateRecordDesc," & vbNewLine & _
    "        RepTo," & vbNewLine & _
    "        MsgText," & vbNewLine & _
    "        WorkflowInstanceID," & vbNewLine & _
    "        [Subject])" & vbNewLine & _
    "      VALUES (''," & vbNewLine & _
    "        ''," & vbNewLine & _
    "        getdate()," & vbNewLine & _
    "        'OpenHR Self-service Intranet'," & vbNewLine & _
    "        1," & vbNewLine & _
    "        0," & vbNewLine & _
    "        @psEmailAddress," & vbNewLine & _
    "        @sMessage," & vbNewLine & _
    "        0," & vbNewLine & _
    "        'How to reset your self-service intranet password');" & vbNewLine & vbNewLine
    
  sProcSQL = sProcSQL & "      EXEC [dbo].[spASREmailImmediate] 'OpenHR Mobile';" & vbNewLine & _
    "    END;" & vbNewLine & _
    "  END;" & vbNewLine & _
    "END;"

  gADOCon.Execute sProcSQL, , adExecuteNoRecords
  gADOCon.Execute "GRANT EXECUTE ON dbo.spadmin_resetpassword TO [OpenHR2IIS];"

TidyUpAndExit:
  CreateSP_ResetPassword = fCreatedOK
  Exit Function

ErrorTrap:
  fCreatedOK = False
  OutputError "Error creating Reset Password stored procedure (Intranet)"
  Resume TidyUpAndExit

End Function


