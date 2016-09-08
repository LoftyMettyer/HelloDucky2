Attribute VB_Name = "modIntranetSpecifics"
Option Explicit

Private Const msResetPassword_PROCEDURENAME = "spadmin_resetpassword"
Private Const msOrgChart_PROCEDURENAME = "spASRIntOrgChart"
Private Const msOrgChart_FUNCTIONNAME = "udfASRIntOrgChartGetTopLevelID"

Private Const msWorkEMailColumnNotDefined = "'Work email' column not defined."

Private mvar_fGeneralOK As Boolean
Private mvar_sGeneralMsg As String

Private mvar_sLoginColumn As String
Private mvar_sLoginTable As String
Private mvar_sWorkEmailColumn As String
Private mvar_sStartingDateColumn As String
Private mvar_sLeavingDateColumn As String
Private mvar_sActivatedUserColumn As String
Private mvar_lngWorkEmailColumn As Long
Private mvar_lngLeavingDateColumn As Long
Private mvar_lngActivatedUserColumn As Long

Private mvar_lngEmployeeTableColumnID As Long
Private mvar_sEmployeeTable As String
Private mvar_sEmployeeNumberColumn As String
Private mvar_sEmployeeForenameColumn As String
Private mvar_sEmployeeSurnameColumn As String
Private mvar_sManagerEmployeeNumberColumn As String
Private mvar_sEmployeeJobTitleColumn As String
Private mvar_sEmployeePhotographColumn As String

Private mvar_sAbsenceTable As String
Private mvar_sAbsenceTypeColumn As String
Private mvar_sAbsenceReasonColumn As String
Private mvar_sAbsenceStartDateColumn As String
Private mvar_sAbsenceEndDateColumn As String

Private mvar_sTBTable As String
Private mvar_sTBCourseTitleColumn As String
Private mvar_sTBStartDateColumn As String
Private mvar_sTBEndDateColumn As String


Public Sub DropIntranetObjects()
  DropProcedure msResetPassword_PROCEDURENAME
  DropFunction msOrgChart_FUNCTIONNAME
  DropProcedure msOrgChart_PROCEDURENAME
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
    
    If mvar_sActivatedUserColumn = "" Or mvar_sWorkEmailColumn = "" Then
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
  If Not (mvar_sActivatedUserColumn = "" Or mvar_sWorkEmailColumn = "") Then
    fOK = CreateSP_ResetPassword
    If Not fOK Then
      DropProcedure msResetPassword_PROCEDURENAME
    End If
  End If
  
  ' Create the OrgChart scalar function.
  If fOK And mvar_fGeneralOK Then
    fOK = CreateUDFOrgChartGetTopLevelID
    If Not fOK Then
      DropFunction msOrgChart_FUNCTIONNAME
    End If
  End If
  
  ' Create the OrgChart stored procedure.
  If fOK And mvar_fGeneralOK Then
    fOK = CreateSP_OrgChart
    If Not fOK Then
      DropProcedure msOrgChart_PROCEDURENAME
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
    
    ' --------------For Organisation Charts----------------
    If fOK Then
      ' Get the Employee Table column ID.
      .Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE
      If .NoMatch Then
        mvar_lngEmployeeTableColumnID = 0
      Else
        mvar_lngEmployeeTableColumnID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
        mvar_sEmployeeTable = GetTableName(mvar_lngEmployeeTableColumnID)
      End If
      
      fOK = (mvar_lngEmployeeTableColumnID > 0)
      If Not fOK Then mvar_sGeneralMsg = mvar_sGeneralMsg & vbCrLf & "'Employee Table' not defined."
    End If
    
    If fOK Then
      ' Get the Employee Number column.
      .Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_EMPLOYEENUMBER
      If .NoMatch Then
        lngColumnID = 0
      Else
        lngColumnID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
        mvar_sEmployeeNumberColumn = GetColumnName(lngColumnID, True)
      End If
      
      fOK = (lngColumnID > 0)
      If Not fOK Then mvar_sGeneralMsg = mvar_sGeneralMsg & vbCrLf & "'Staff Number' column not defined."
    End If
    
    If fOK Then
      ' Get the Employee Forename column.
      .Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_FORENAME
      If .NoMatch Then
        lngColumnID = 0
      Else
        lngColumnID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
        mvar_sEmployeeForenameColumn = GetColumnName(lngColumnID, True)
      End If
      
      fOK = (lngColumnID > 0)
      If Not fOK Then mvar_sGeneralMsg = mvar_sGeneralMsg & vbCrLf & "'Employee Forename' column not defined."
    End If
     
    If fOK Then
      ' Get the Employee Surname column.
      .Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_SURNAME
      If .NoMatch Then
        lngColumnID = 0
      Else
        lngColumnID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
        mvar_sEmployeeSurnameColumn = GetColumnName(lngColumnID, True)
      End If
      
      fOK = (lngColumnID > 0)
      If Not fOK Then mvar_sGeneralMsg = mvar_sGeneralMsg & vbCrLf & "'Employee Forename' column not defined."
    End If
      
    If fOK Then
      ' Get the Manager staff number column.
      .Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_MANAGERSTAFFNO
      If .NoMatch Then
        lngColumnID = 0
      Else
        lngColumnID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
        mvar_sManagerEmployeeNumberColumn = GetColumnName(lngColumnID, True)
      End If
      
      fOK = (lngColumnID > 0)
      If Not fOK Then mvar_sGeneralMsg = mvar_sGeneralMsg & vbCrLf & "'Line Manager Staff Number' column not defined."
    End If
    
    If fOK Then
      ' Get the Employee Job title column.
      .Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_JOBTITLE
      If .NoMatch Then
        lngColumnID = 0
      Else
        lngColumnID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
        mvar_sEmployeeJobTitleColumn = GetColumnName(lngColumnID, True)
      End If
      
      fOK = (lngColumnID > 0)
      If Not fOK Then mvar_sGeneralMsg = mvar_sGeneralMsg & vbCrLf & "'Employee Job Title' column not defined."
    End If
    
    If fOK Then
      ' Get the Employee Photograph column.
      .Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_SSIPHOTOGRAPH
      If .NoMatch Then
        lngColumnID = 0
      Else
        lngColumnID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
        mvar_sEmployeePhotographColumn = GetColumnName(lngColumnID, True)
      End If
      
      fOK = (lngColumnID > 0)
      If Not fOK Then mvar_sGeneralMsg = mvar_sGeneralMsg & vbCrLf & "'Employee Photograph' column not defined."
    End If
       
    If fOK Then
      ' Get the Employee Start Date column.
      .Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_STARTDATE
      If .NoMatch Then
        lngColumnID = 0
      Else
        lngColumnID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
        mvar_sStartingDateColumn = GetColumnName(lngColumnID, True)
      End If
      
      fOK = (lngColumnID > 0)
      If Not fOK Then mvar_sGeneralMsg = mvar_sGeneralMsg & vbCrLf & "'Employee Start Date' column not defined."
    End If
    
    If fOK Then
      ' Get the Employee Leaving Date column.
      .Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_LEAVINGDATE
      If .NoMatch Then
        lngColumnID = 0
      Else
        lngColumnID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
        mvar_sLeavingDateColumn = GetColumnName(lngColumnID, True)
      End If
      
      fOK = (lngColumnID > 0)
      If Not fOK Then mvar_sGeneralMsg = mvar_sGeneralMsg & vbCrLf & "'Employee Leaving Date' column not defined."
    End If
    
    ' --- Absence columns for org charts
    If fOK Then
      ' Get the Absence Table Name.
      .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETABLE
      If .NoMatch Then
        lngTableID = 0
      Else
        lngTableID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
        mvar_sAbsenceTable = GetTableName(lngTableID)
      End If
      
      fOK = (lngTableID > 0)
      If Not fOK Then mvar_sGeneralMsg = mvar_sGeneralMsg & vbCrLf & "'Absence Table' not defined."
    End If
    
    If fOK Then
      ' Get the Absence Type column.
      .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPE
      If .NoMatch Then
        lngColumnID = 0
      Else
        lngColumnID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
        mvar_sAbsenceTypeColumn = GetColumnName(lngColumnID, True)
      End If
      
      fOK = (lngColumnID > 0)
      If Not fOK Then mvar_sGeneralMsg = mvar_sGeneralMsg & vbCrLf & "'Absence Reason' column not defined."
    End If
    
    If fOK Then
      ' Get the Absence Reason column.
      .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEREASON
      If .NoMatch Then
        lngColumnID = 0
      Else
        lngColumnID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
        mvar_sAbsenceReasonColumn = GetColumnName(lngColumnID, True)
      End If
      
      fOK = (lngColumnID > 0)
      If Not fOK Then mvar_sGeneralMsg = mvar_sGeneralMsg & vbCrLf & "'Absence Reason' column not defined."
    End If
    
    If fOK Then
      ' Get the Absence StartDate column.
      .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESTARTDATE
      If .NoMatch Then
        lngColumnID = 0
      Else
        lngColumnID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
        mvar_sAbsenceStartDateColumn = GetColumnName(lngColumnID, True)
      End If
      
      fOK = (lngColumnID > 0)
      If Not fOK Then mvar_sGeneralMsg = mvar_sGeneralMsg & vbCrLf & "'Absence Start Date' column not defined."
    End If
    
    If fOK Then
      ' Get the Absence EndDate column.
      .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEENDDATE
      If .NoMatch Then
        lngColumnID = 0
      Else
        lngColumnID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
        mvar_sAbsenceEndDateColumn = GetColumnName(lngColumnID, True)
      End If
      
      fOK = (lngColumnID > 0)
      If Not fOK Then mvar_sGeneralMsg = mvar_sGeneralMsg & vbCrLf & "'Absence End Date' column not defined."
    End If
    
    ' --- Training Booking columns for org charts
    
    ' *** No training booking module setup info at present, so removed. ***
    
'    If fOK Then
'      ' Get the Training Booking Table Name.
'      .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_TRAINBOOKTABLE
'      If .NoMatch Then
'        lngTableID = 0
'      Else
'        lngTableID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
'        mvar_sTBTable = GetTableName(lngTableID)
'      End If
'
'      fOK = (lngTableID > 0)
'      If Not fOK Then mvar_sGeneralMsg = mvar_sGeneralMsg & vbCrLf & "'Training Booking Table' not defined."
'    End If
'
'    If fOK Then
'      ' Get the TB Course Title column.
'      .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_TRAINBOOKCOURSETITLE
'      If .NoMatch Then
'        lngColumnID = 0
'      Else
'        If IsNull(!parametervalue) Then
'          lngColumnID = 0
'        Else
'          lngColumnID = val(!parametervalue)
'        End If
'        'lngColumnID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
'        mvar_sTBCourseTitleColumn = GetColumnName(lngColumnID, True)
'      End If
'
'      fOK = (lngColumnID > 0)
'      If Not fOK Then mvar_sGeneralMsg = mvar_sGeneralMsg & vbCrLf & "'Training Booking Course Title' column not defined."
'    End If
'
'    If fOK Then
'      ' Get the TB Course Start Date column.
'      .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSESTARTDATE ' yoinked from CR
'      If .NoMatch Then
'        lngColumnID = 0
'      Else
'        lngColumnID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
'        mvar_sTBStartDateColumn = GetColumnName(lngColumnID, True)
'      End If
'
'      fOK = (lngColumnID > 0)
'      If Not fOK Then mvar_sGeneralMsg = mvar_sGeneralMsg & vbCrLf & "'Training Booking Start Date' column not defined."
'    End If
'
'    If fOK Then
'      ' Get the TB Course End Date column.
'      .Seek "=", gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSEENDDATE ' yoinked from CR
'      If .NoMatch Then
'        lngColumnID = 0
'      Else
'        lngColumnID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
'        mvar_sTBEndDateColumn = GetColumnName(lngColumnID, True)
'      End If
'
'      fOK = (lngColumnID > 0)
'      If Not fOK Then mvar_sGeneralMsg = mvar_sGeneralMsg & vbCrLf & "'Training Booking End Date' column not defined."
'    End If
    
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
    "      SET @sMessage = 'To reset your password click the link shown below. This will take you to a web page where you can enter a new password.' + CHAR(13) + CHAR(10) +" & vbNewLine & _
    "            'If you weren''t trying to reset your password, don''t worry — your account is still secure and no one has been given access to it.' + CHAR(13) + CHAR(10) + CHAR(13) + CHAR(10) +" & vbNewLine & _
    "            '<' + @psWebsiteURL + '?' + @psEncryptedLink + '>';" & vbNewLine & vbNewLine
    
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
    "        'OpenHR Web'," & vbNewLine & _
    "        1," & vbNewLine & _
    "        0," & vbNewLine & _
    "        @psEmailAddress," & vbNewLine & _
    "        @sMessage," & vbNewLine & _
    "        0," & vbNewLine & _
    "        'How to reset your OpenHR password');" & vbNewLine & vbNewLine
    
  sProcSQL = sProcSQL & "      EXEC [dbo].[spASREmailImmediate] 'OpenHR Web';" & vbNewLine & _
    "    END;" & vbNewLine & _
    "  END;" & vbNewLine & _
    "END;"

  gADOCon.Execute sProcSQL, , adExecuteNoRecords

TidyUpAndExit:
  CreateSP_ResetPassword = fCreatedOK
  Exit Function

ErrorTrap:
  fCreatedOK = False
  OutputError "Error creating Reset Password stored procedure (Intranet)"
  Resume TidyUpAndExit

End Function

Private Function CreateSP_OrgChart() As Boolean
  ' Create the Check Login stored procedure.
  On Error GoTo ErrorTrap

  Dim fCreatedOK As Boolean
  Dim sProcSQL As String
  Dim iCount As Integer
  
  fCreatedOK = True

  ' Construct the stored procedure creation string.
  
  sProcSQL = "/* ------------------------------------------------ */" & vbNewLine & _
    "/* Intranet module stored procedure.         */" & vbNewLine & _
    "/* Automatically generated by the System manager.   */" & vbNewLine & _
    "/* ------------------------------------------------ */" & vbNewLine & _
    "CREATE PROCEDURE [dbo].[" & msOrgChart_PROCEDURENAME & "](" & vbNewLine & _
    "  @RootID INT" & vbNewLine & _
    "  )" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine
    
  sProcSQL = sProcSQL & "       SET NOCOUNT ON;" & vbNewLine & _
    "       DECLARE @staff_number VARCHAR(MAX);" & vbNewLine & _
    "       DECLARE @today DATETIME = DATEADD(dd, 0, DATEDIFF(dd, 0,  getdate()));" & vbNewLine & _
    vbNewLine & _
    vbNewLine & _
    "       -- Get top level manager" & vbNewLine & _
    "       SELECT @RootID = dbo.udfASRIntOrgChartGetTopLevelID( @RootID);" & vbNewLine & _
    vbNewLine & _
    "       SELECT @staff_number = " & mvar_sEmployeeNumberColumn & " FROM " & mvar_sEmployeeTable & " WHERE id=@RootID;" & vbNewLine & _
    vbNewLine & _
    "       WITH Emp_CTE AS (" & vbNewLine & _
    "              SELECT id, " & mvar_sEmployeeForenameColumn & ", " & mvar_sEmployeeSurnameColumn & " AS name, " & mvar_sEmployeeNumberColumn & ", " & mvar_sManagerEmployeeNumberColumn & ", " & mvar_sEmployeeJobTitleColumn & ", 1 AS HierarchyLevel, " & mvar_sEmployeePhotographColumn & vbNewLine & _
    "                     FROM " & mvar_sEmployeeTable & vbNewLine & _
    "                     WHERE " & mvar_sManagerEmployeeNumberColumn & " = @staff_number" & vbNewLine & _
    "                         AND (" & mvar_sLeavingDateColumn & " IS NULL OR " & mvar_sLeavingDateColumn & " >= @today) AND " & mvar_sStartingDateColumn & " <= @today" & vbNewLine & _
    "              UNION ALL" & vbNewLine & _
    "                     SELECT e.id, e." & mvar_sEmployeeForenameColumn & ", e." & mvar_sEmployeeSurnameColumn & ", e." & mvar_sEmployeeNumberColumn & ", e." & mvar_sManagerEmployeeNumberColumn & ", e." & mvar_sEmployeeJobTitleColumn & ", ecte.HierarchyLevel + 1 AS HierarchyLevel, e." & mvar_sEmployeePhotographColumn & vbNewLine & _
    "                     FROM " & mvar_sEmployeeTable & " e" & vbNewLine & _
    "                     INNER JOIN Emp_CTE ecte ON ecte." & mvar_sEmployeeNumberColumn & " = e." & mvar_sManagerEmployeeNumberColumn & "" & vbNewLine & _
    "                     WHERE (" & mvar_sLeavingDateColumn & " IS NULL OR " & mvar_sLeavingDateColumn & " >= @today) AND " & mvar_sStartingDateColumn & " <= @today" & vbNewLine & _
    "       )" & vbNewLine

  sProcSQL = sProcSQL & _
    "       SELECT p.*, '' AS [type], '' AS [reason], '' AS course_title FROM Emp_CTE p" & vbNewLine & _
    vbNewLine & _
    "    UNION" & vbNewLine & _
    "      SELECT id, " & mvar_sEmployeeForenameColumn & ", " & mvar_sEmployeeSurnameColumn & " AS name, " & mvar_sEmployeeNumberColumn & ", " & mvar_sManagerEmployeeNumberColumn & ", " & mvar_sEmployeeJobTitleColumn & ", 0 AS HierarchyLevel, " & mvar_sEmployeePhotographColumn & "," & vbNewLine & _
    "      NULL AS type, NULL AS reason, NULL AS course_title" & vbNewLine & _
    "        FROM " & mvar_sEmployeeTable & vbNewLine & _
    "        WHERE ID = @RootID" & vbNewLine & _
    vbNewLine & _
    "       ORDER BY hierarchylevel, " & mvar_sEmployeeJobTitleColumn & ", name" & vbNewLine & _
    "End" & vbNewLine


  gADOCon.Execute sProcSQL, , adExecuteNoRecords

TidyUpAndExit:
  CreateSP_OrgChart = fCreatedOK
  Exit Function

ErrorTrap:
  fCreatedOK = False
  OutputError "Error creating Organisation Chart stored procedure (Intranet)"
  Resume TidyUpAndExit

End Function
Private Function CreateUDFOrgChartGetTopLevelID() As Boolean
  ' Create the Check Login stored procedure.
  On Error GoTo ErrorTrap

  Dim fCreatedOK As Boolean
  Dim sProcSQL As String
  Dim iCount As Integer
  
  fCreatedOK = True

  ' Construct the stored procedure creation string.
  
  sProcSQL = "/* ------------------------------------------------ */" & vbNewLine & _
    "/* Intranet module function.         */" & vbNewLine & _
    "/* Automatically generated by the System manager.   */" & vbNewLine & _
    "/* ------------------------------------------------ */" & vbNewLine & _
    "CREATE FUNCTION [dbo].[" & msOrgChart_FUNCTIONNAME & "](" & vbNewLine & _
    "  @StaffRecordID integer" & vbNewLine & _
    "  )" & vbNewLine & _
    "RETURNS integer" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine
  sProcSQL = sProcSQL & _
    "  DECLARE @ManagerID varchar(MAX)," & vbNewLine & _
    "      @today DATETIME = DATEADD(dd, 0, DATEDIFF(dd, 0,  getdate()))," & vbNewLine & _
    "      @ManagerRecordID integer;" & vbNewLine & _
    vbNewLine & _
    "  SELECT @ManagerID = [" & mvar_sManagerEmployeeNumberColumn & "]" & vbNewLine & _
    "    FROM [" & mvar_sEmployeeTable & "]" & vbNewLine & _
    "    WHERE id = @StaffRecordID AND (" & mvar_sLeavingDateColumn & " IS NULL OR " & mvar_sLeavingDateColumn & " >= @today) AND " & mvar_sStartingDateColumn & " <= @today;" & vbNewLine & vbNewLine & _
    "  IF ISNULL(@ManagerID,'') = ''" & vbNewLine & _
    "    RETURN @StaffRecordID;" & vbNewLine & _
    "  ELSE" & vbNewLine & _
    "    BEGIN" & vbNewLine & _
    "      SELECT @ManagerRecordID = ID FROM [" & mvar_sEmployeeTable & "] WHERE [" & mvar_sEmployeeNumberColumn & "] = @ManagerID;" & vbNewLine & _
    "      IF ISNULL(@ManagerRecordID,0) = 0 RETURN @StaffRecordID" & vbNewLine & _
    "      SELECT @ManagerRecordID = dbo.udfASRIntOrgChartGetTopLevelID(@ManagerRecordID);" & vbNewLine & _
    "    END" & vbNewLine & _
    "  RETURN @ManagerRecordID;" & vbNewLine & _
    "END;"
      
    gADOCon.Execute sProcSQL, , adExecuteNoRecords

TidyUpAndExit:
  CreateUDFOrgChartGetTopLevelID = fCreatedOK
  Exit Function

ErrorTrap:
  fCreatedOK = False
  OutputError "Error creating Organisation Chart scalar function (Intranet)"
  Resume TidyUpAndExit
  
End Function
    

