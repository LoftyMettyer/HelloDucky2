Attribute VB_Name = "modAbsenceSpecifics"
Option Explicit

''26/07/2001 MH Refer modLicence
'Public Enum Module
'    Personnel = 1
'    Recruitment = 2
'    Absence = 4
'    Training = 8
'    Skills = 16
'    Web = 32
'    Afd = 64
'End Enum


' COPY THIS INTO MODSYSMGR ASAP !
Public gsDatabaseName As String
'##############################


Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hWnd As Long, ByVal lpoperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' Absence table variables.
Private mvar_lngAbsenceTableID As Long
Private mvar_sAbsenceTableName As String
Private mvar_lngAbsence_StartDateColumnID As Long
Private mvar_sAbsence_StartDateColumnName As String
Private mvar_lngAbsence_EndDateColumnID As Long
Private mvar_sAbsence_EndDateColumnName As String
Private mvar_lngAbsence_StartSessionColumnID As Long
Private mvar_sAbsence_StartSessionColumnName As String
Private mvar_lngAbsence_EndSessionColumnID As Long
Private mvar_sAbsence_EndSessionColumnName As String
Private mvar_lngAbsence_SSPAppliesColumnID As Long
Private mvar_sAbsence_SSPAppliesColumnName As String
Private mvar_lngAbsence_QualifyingDaysColumnID As Long
Private mvar_sAbsence_QualifyingDaysColumnName As String
Private mvar_lngAbsence_WaitingDaysColumnID As Long
Private mvar_sAbsence_WaitingDaysColumnName As String
Private mvar_lngAbsence_PaidDaysColumnID As Long
Private mvar_sAbsence_PaidDaysColumnName As String
Private mvar_lngAbsence_TypeColumnID As Long
Private mvar_sAbsence_TypeColumnName As String
Private mvar_iAbsenceWorkingDaysType As Integer
Private mvar_iAbsenceWorkingDaysNumericValue As Integer
Private mvar_sAbsenceWorkingDaysPatternValue As String
Private mvar_lngAbsenceWorkingDaysColumnID As Long
Private mvar_sAbsenceWorkingDaysColumnName As String
Private mvar_lngAbsenceWorkingDaysTableID As String
Private mvar_lngAbsenceContinuousColumnID As Long
Private mvar_sAbsenceContinuousColumnName As String
Private mvar_lngAbsenceDurationColumnID As Long
Private mvar_sAbsenceDurationColumnName As String

' Personnel table variables.
Private mvar_lngPersonnelTableID As Long
Private mvar_sPersonnelTableName As String
Private mvar_lngPersonnel_DateOfBirthColumnID As Long
Private mvar_sPersonnel_DateOfBirthColumnName As String

' Absence Type table variables.
Private mvar_lngAbsenceTypeTableID As Long
Private mvar_sAbsenceTypeTableName As String
Private mvar_lngAbsenceType_TypeColumnID As Long
Private mvar_sAbsenceType_TypeColumnName As String
Private mvar_lngAbsenceType_SSPAppliesColumnID As Long
Private mvar_sAbsenceType_SSPAppliesColumnName As String
Private mvar_lngAbsenceType_IncludeInBradfordColumnName As Long
Private mvar_sAbsenceType_IncludeInBradfordColumnName As String

Public Const gsSSP_PROCEDURENAME = "sp_ASR_AbsenceSSP"
Public Const gsWorkingDaysBetween2Dates_PROCEDURENAME = "sp_ASRFn_WorkingDaysBetweenTwoDates"

Private mvar_fGeneralOK As Boolean
Private mvar_sGeneralMsg As String

' NPG Fault HRPRO-735
Private mvar_fSSPGeneralOK As Boolean
Private mvar_sSSPGeneralMsg As String

Public Function ConfigureAbsenceSpecifics() As Boolean
  ' Configure module specific objects (eg. stored procedures)
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sErrorMessage As String
  
  mvar_fGeneralOK = True
  mvar_sGeneralMsg = ""
  
  mvar_fSSPGeneralOK = True
  mvar_sSSPGeneralMsg = ""
  
  ' RH 20/11/00 - Create the AbsenceBetween2Dates Stored Procedure.
  fOK = CreateAbsenceBetween2DatesStoredProcedure
  
  ' Create the AbsenceBreakdown Stored Procedure.
  If fOK Then fOK = CreateAbsenceDurationStoredProcedure(True)
  
  ' RH 04/04/01 - Create the AbsenceDuration Stored Procedure.
  If fOK Then fOK = CreateAbsenceDurationStoredProcedure(False)
  
  ' JPD20020515 Fault 3342
  If fOK Then fOK = CreateWorkingDaysBetween2DatesStoredProcedure
  

  ' Read the Absence table parameters.
  If fOK Then
    fOK = ReadAbsenceRecordParameters
  End If
  
  ' Read the Personnel table parameters.
  If fOK Then
    fOK = ReadPersonnelRecordParameters
  End If

  ' Read the Absence Type table parameters.
  If fOK Then
    fOK = ReadAbsenceTypeRecordParameters
  End If
 
  If fOK Then
    sErrorMessage = ""
    If (Not mvar_fGeneralOK) Then
      sErrorMessage = "Absence specifics not correctly configured." & vbNewLine & _
        "Some functionality will be disabled if you do not change your configuration." & mvar_sGeneralMsg
      
      fOK = (OutputMessage(sErrorMessage & vbNewLine & vbNewLine & "Continue saving changes ?") = vbYes)
    End If
  End If
  
  'MH20020308 Fault 3444
  'Make sure that we drop this SP
  'DropSSPStoredProcedure
  DropProcedure gsSSP_PROCEDURENAME
  
  ' NPG20100607 Fault HRPRO-735
  If mvar_fSSPGeneralOK Then
    ' Create the SSP stored procedure.
    If fOK And mvar_fGeneralOK Then
      fOK = CreateSSPStoredProcedure
      If Not fOK Then
        'DropSSPStoredProcedure
        DropProcedure gsSSP_PROCEDURENAME
      End If
    End If
  End If
  
TidyUpAndExit:
  ConfigureAbsenceSpecifics = fOK
  Exit Function
  
ErrorTrap:
  OutputError "Error configuring Absence specifics"
  fOK = False
  Resume TidyUpAndExit

End Function

Private Function CreateAbsenceDurationStoredProcedure(Optional pbCreateBreakdown As Boolean) As Boolean
  
  On Error GoTo ErrorTrap

  Dim fCreatedOK As Boolean
  Dim sSQL As String
  Dim sProcSQL As String
  
  Dim sUDFSQL As String
  Dim sUDFName As String
  
  Dim strUDFHeader As String
  Dim strProcHeader As String
  Dim strGeneralSQL As String
  Dim strEndProc As String
  Dim strEndUDF As String
    
  Dim iTempID As Integer
  Dim iLoop As Integer

  Dim fHistoricRegion As Boolean
  Dim sStaticRegionColumnName As String
  Dim sHistoricRegionTableName As String
  Dim sHistoricRegionColumnName As String
  Dim sHistoricRegionDateColumnName As String
  Dim fHistoricWP As Boolean
  Dim sStaticWPColumnName As String
  Dim sHistoricWPTableName As String
  Dim sHistoricWPColumnName As String
  Dim sHistoricWPDateColumnName As String
  Dim fBHolSetupOK As Boolean
  Dim iBHolRegionTableID As Integer
  Dim sBHolRegionTableName As String
  Dim sBHolRegionColumnName As String
  Dim iBHolTableID As Integer
  Dim sBHolTableName As String
  Dim sBHolDateColumnName As String
  Dim iPersonnelTableID As Integer
  Dim sPersonnelTableName As String
  Dim sStoredProcedureName As String
  Dim sExtraBreakdownParameters As String

  'Initialse different parameters
  If Not pbCreateBreakdown Then
    sStoredProcedureName = "dbo.sp_ASRFn_AbsenceDuration"
    sUDFName = "dbo.udf_ASRFn_AbsenceDuration"
    sExtraBreakdownParameters = ""
  Else
    sStoredProcedureName = "dbo.sp_ASR_AbsenceBreakdown_Calculate"
    sUDFName = "dbo.udf_ASR_AbsenceBreakdown_Calculate"
    sExtraBreakdownParameters = "  @pfMonTotal                      float OUTPUT," & vbNewLine & _
      "  @pfTueTotal                      float OUTPUT," & vbNewLine & _
      "  @pfWedTotal                      float OUTPUT," & vbNewLine & _
      "  @pfThuTotal                      float OUTPUT," & vbNewLine & _
      "  @pfFriTotal                      float OUTPUT," & vbNewLine & _
      "  @pfSatTotal                      float OUTPUT," & vbNewLine & _
      "  @pfSunTotal                      float OUTPUT," & vbNewLine
  End If

  ' Drop any existing stored procedure.
  If Not pbCreateBreakdown Then
    'fCreatedOK = DropAbsenceDurationStoredProcedure
    fCreatedOK = DropProcedure("sp_ASRFn_AbsenceDuration")
    If gbEnableUDFFunctions Then
      'fCreatedOK = DropAbsenceDurationUDF
      fCreatedOK = DropFunction("udf_ASRFn_AbsenceDuration")
    End If
  Else
    'fCreatedOK = DropAbsenceBreakdownCalcStoredProcedure
    fCreatedOK = DropProcedure("sp_ASR_AbsenceBreakdown_Calculate")
  End If

  ' Find out the things that used to be worked out in the stored procedure
  ' but need to be done here now so permission problems are solved.
  
  ' REGION STUFF
  
  ' Get the Static Region Column Name
  iTempID = 0
  recModuleSetup.Index = "idxModuleParameter"
  recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, "Param_FieldsRegion"
  If Not recModuleSetup.NoMatch Then
    iTempID = recModuleSetup!parametervalue
    recColEdit.Index = "idxColumnID"
    recColEdit.Seek "=", iTempID
    If Not recColEdit.NoMatch Then
      sStaticRegionColumnName = recColEdit!ColumnName
    Else
      sStaticRegionColumnName = vbNullString
    End If
  Else
    sStaticRegionColumnName = vbNullString
  End If
  
  ' Get the Historic Region Table Name
  iTempID = 0
  recModuleSetup.Index = "idxModuleParameter"
  recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, "Param_FieldsHRegionTable"
  If Not recModuleSetup.NoMatch Then
    iTempID = recModuleSetup!parametervalue
    recTabEdit.Index = "idxTableID"
    recTabEdit.Seek "=", iTempID
    If Not recTabEdit.NoMatch Then
      sHistoricRegionTableName = recTabEdit!TableName
    Else
      sHistoricRegionTableName = vbNullString
    End If
  Else
    sHistoricRegionTableName = vbNullString
  End If
  
  ' Get the Historic Region Column Name
  iTempID = 0
  recModuleSetup.Index = "idxModuleParameter"
  recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, "Param_FieldsHRegion"
  If Not recModuleSetup.NoMatch Then
    iTempID = recModuleSetup!parametervalue
    recColEdit.Index = "idxColumnID"
    recColEdit.Seek "=", iTempID
    If Not recColEdit.NoMatch Then
      sHistoricRegionColumnName = recColEdit!ColumnName
    Else
      sHistoricRegionColumnName = vbNullString
    End If
  Else
    sHistoricRegionColumnName = vbNullString
  End If
  
  ' Get the Historic Region Date Column Name
  iTempID = 0
  recModuleSetup.Index = "idxModuleParameter"
  recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, "Param_FieldsHRegionDate"
  If Not recModuleSetup.NoMatch Then
    iTempID = recModuleSetup!parametervalue
    recColEdit.Index = "idxColumnID"
    recColEdit.Seek "=", iTempID
    If Not recColEdit.NoMatch Then
      sHistoricRegionDateColumnName = recColEdit!ColumnName
    Else
      sHistoricRegionDateColumnName = vbNullString
    End If
  Else
    sHistoricRegionDateColumnName = vbNullString
  End If
  
  ' Define the standard heading for the stored procedure.
  strProcHeader = "/* ------------------------------------------------ */" & vbNewLine & _
              "/* HR Pro Absence module stored procedure.          */" & vbNewLine & _
              "/* Automatically generated by the System manager.   */" & vbNewLine & _
              "/* ------------------------------------------------ */" & vbNewLine & _
              "CREATE PROCEDURE " & sStoredProcedureName & "(" & vbNewLine & _
              "   @pdblResult                       float OUTPUT," & vbNewLine & _
              sExtraBreakdownParameters & _
              "   @pdtStartDate                     datetime," & vbNewLine & _
              "   @psStartSession                   varchar(255)," & vbNewLine & _
              "   @pdtEndDate                       datetime," & vbNewLine & _
              "   @psEndSession                     varchar(255)," & vbNewLine & _
              "   @iPersonnelID                     int" & vbNewLine & _
              "   ) " & vbNewLine & _
              "AS " & vbNewLine & _
              "BEGIN" & vbNewLine & vbNewLine
  strEndProc = "END"

  ' Define the standard heading for the user defined function.
  strUDFHeader = "/* ------------------------------------------------ */" & vbNewLine & _
              "/* HR Pro Absence module user defined function.     */" & vbNewLine & _
              "/* Automatically generated by the System manager.   */" & vbNewLine & _
              "/* ------------------------------------------------ */" & vbNewLine & _
              "CREATE FUNCTION " & sUDFName & "(" & vbNewLine & _
              sExtraBreakdownParameters & _
              "   @pdtStartDate                     datetime," & vbNewLine & _
              "   @psStartSession                   varchar(255)," & vbNewLine & _
              "   @pdtEndDate                       datetime," & vbNewLine & _
              "   @psEndSession                     varchar(255)," & vbNewLine & _
              "   @iPersonnelID                     int" & vbNewLine & _
              "   ) " & vbNewLine & _
              "RETURNS float" & vbNewLine & _
              "AS " & vbNewLine & _
              "BEGIN" & vbNewLine & vbNewLine & _
              "  DECLARE @pdblResult AS float" & vbNewLine
  strEndUDF = "RETURN @pdblResult" & vbNewLine & "END"

  ' Set flag now that we have got the required values from module setup
  If sStaticRegionColumnName = vbNullString Then
    If (sHistoricRegionTableName = vbNullString) Or _
       (sHistoricRegionColumnName = vbNullString) Or _
       (sHistoricRegionDateColumnName) = vbNullString Then
    
      strGeneralSQL = "/* ERROR IN MODULE SETUP...NEITHER STATIC NOR HISTORIC REGIONS ARE DEFINED */" & vbNewLine & _
                 "SET @pdblResult = 0" & vbNewLine & vbNewLine
      
      ' Create the udf and sp
      sProcSQL = strProcHeader & strGeneralSQL & strEndProc
      sUDFSQL = strUDFHeader & strGeneralSQL & strEndUDF
      
      gADOCon.Execute sProcSQL, , adExecuteNoRecords
      If Not pbCreateBreakdown And gbEnableUDFFunctions Then gADOCon.Execute sUDFSQL, , adExecuteNoRecords
      
      CreateAbsenceDurationStoredProcedure = False
      gobjProgress.Visible = False
        MsgBox "The Absence module requires that you fully complete both" & vbNewLine & _
               "the Absence and the Personnel module setup screens." & vbNewLine & _
               "You do not have a 'Region' field defined.", vbExclamation + vbOKOnly, App.Title
      gobjProgress.Visible = True
      Exit Function
    Else
      fHistoricRegion = True
    End If
  Else
    fHistoricRegion = False
  End If

  ' WORKING PATTERN STUFF
  
  ' Get the Static WP Column Name
  iTempID = 0
  recModuleSetup.Index = "idxModuleParameter"
  recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, "Param_FieldsWorkingPattern"
  If Not recModuleSetup.NoMatch Then
    iTempID = recModuleSetup!parametervalue
    recColEdit.Index = "idxColumnID"
    recColEdit.Seek "=", iTempID
    If Not recColEdit.NoMatch Then
      sStaticWPColumnName = recColEdit!ColumnName
    Else
      sStaticWPColumnName = vbNullString
    End If
  Else
    sStaticWPColumnName = vbNullString
  End If
  
  ' Get the Historic WP Table Name
  iTempID = 0
  recModuleSetup.Index = "idxModuleParameter"
  recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, "Param_FieldsHWorkingPatternTable"
  If Not recModuleSetup.NoMatch Then
    iTempID = recModuleSetup!parametervalue
    recTabEdit.Index = "idxTableID"
    recTabEdit.Seek "=", iTempID
    If Not recTabEdit.NoMatch Then
      sHistoricWPTableName = recTabEdit!TableName
    Else
      sHistoricWPTableName = vbNullString
    End If
  Else
    sHistoricWPTableName = vbNullString
  End If
  

  ' Get the Historic WP Column Name
  iTempID = 0
  recModuleSetup.Index = "idxModuleParameter"
  recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, "Param_FieldsHWorkingPattern"
  If Not recModuleSetup.NoMatch Then
    iTempID = recModuleSetup!parametervalue
    recColEdit.Index = "idxColumnID"
    recColEdit.Seek "=", iTempID
    If Not recColEdit.NoMatch Then
      sHistoricWPColumnName = recColEdit!ColumnName
    Else
      sHistoricWPColumnName = vbNullString
    End If
  Else
    sHistoricWPColumnName = vbNullString
  End If

  ' Get the Historic WP Date Column Name
  iTempID = 0
  recModuleSetup.Index = "idxModuleParameter"
  recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, "Param_FieldsHWorkingPatternDate"
  If Not recModuleSetup.NoMatch Then
    iTempID = recModuleSetup!parametervalue
    recColEdit.Index = "idxColumnID"
    recColEdit.Seek "=", iTempID
    If Not recColEdit.NoMatch Then
      sHistoricWPDateColumnName = recColEdit!ColumnName
    Else
      sHistoricWPDateColumnName = vbNullString
    End If
  Else
    sHistoricWPDateColumnName = vbNullString
  End If

  ' Set flag now that we have got the required values from module setup
  If sStaticWPColumnName = vbNullString Then
    If (sHistoricWPTableName = vbNullString) Or _
       (sHistoricWPColumnName = vbNullString) Or _
       (sHistoricWPDateColumnName) = vbNullString Then
      
      strGeneralSQL = "/* ERROR IN MODULE SETUP...NEITHER STATIC NOR HISTORIC REGIONS ARE DEFINED */" & vbNewLine & _
                 "SET @pdblResult = 0" & vbNewLine & vbNewLine
      
      ' Create the udf and sp
      sProcSQL = strProcHeader & strGeneralSQL & strEndProc
      sUDFSQL = strUDFHeader & strGeneralSQL & strEndUDF
      
      gADOCon.Execute sProcSQL, , adExecuteNoRecords
      If Not pbCreateBreakdown And gbEnableUDFFunctions Then gADOCon.Execute sUDFSQL, , adExecuteNoRecords
      
      CreateAbsenceDurationStoredProcedure = False
      gobjProgress.Visible = False
        MsgBox "The Absence module requires that you fully complete both" & vbNewLine & _
               "the Absence and the Personnel module setup screens." & vbNewLine & _
               "You do not have a 'Working Pattern' field defined.", vbExclamation + vbOKOnly, App.Title
      gobjProgress.Visible = True
      Exit Function
    Else
      fHistoricWP = True
    End If
  Else
    fHistoricWP = False
  End If

  ' BANK HOLIDAY STUFF
  
  ' Get the BHol Region Table ID and Name
  iTempID = 0
  recModuleSetup.Index = "idxModuleParameter"
  recModuleSetup.Seek "=", gsMODULEKEY_ABSENCE, "Param_TableBHolRegion"
  If Not recModuleSetup.NoMatch Then
    iTempID = recModuleSetup!parametervalue
    iBHolRegionTableID = iTempID
    recTabEdit.Index = "idxTableID"
    recTabEdit.Seek "=", iTempID
    If Not recTabEdit.NoMatch Then
      sBHolRegionTableName = recTabEdit!TableName
    Else
      iBHolRegionTableID = 0
      sBHolRegionTableName = vbNullString
    End If
  Else
    iBHolRegionTableID = 0
    sBHolRegionTableName = vbNullString
  End If
  
  ' Get the BHolRegion column in the BHolRegion Table
  iTempID = 0
  recModuleSetup.Index = "idxModuleParameter"
  recModuleSetup.Seek "=", gsMODULEKEY_ABSENCE, "Param_FieldBHolRegion"
  If Not recModuleSetup.NoMatch Then
    iTempID = recModuleSetup!parametervalue
    recColEdit.Index = "idxColumnID"
    recColEdit.Seek "=", iTempID
    If Not recColEdit.NoMatch Then
      sBHolRegionColumnName = recColEdit!ColumnName
    Else
      sBHolRegionColumnName = vbNullString
    End If
  Else
    sBHolRegionColumnName = vbNullString
  End If


  ' Get the BHol Table ID (instances of BHols)
  iTempID = 0
  recModuleSetup.Index = "idxModuleParameter"
  recModuleSetup.Seek "=", gsMODULEKEY_ABSENCE, "Param_TableBHol"
  If Not recModuleSetup.NoMatch Then
    iTempID = recModuleSetup!parametervalue
    iBHolTableID = iTempID
    recTabEdit.Index = "idxTableID"
    recTabEdit.Seek "=", iTempID
    If Not recTabEdit.NoMatch Then
      sBHolTableName = recTabEdit!TableName
    Else
      iBHolTableID = 0
      sBHolTableName = vbNullString
    End If
  Else
    iBHolTableID = 0
    sBHolTableName = vbNullString
  End If


  ' Get the BHolDate Column Name
  iTempID = 0
  recModuleSetup.Index = "idxModuleParameter"
  recModuleSetup.Seek "=", gsMODULEKEY_ABSENCE, "Param_FieldBHolDate"
  If Not recModuleSetup.NoMatch Then
    iTempID = recModuleSetup!parametervalue
    recColEdit.Index = "idxColumnID"
    recColEdit.Seek "=", iTempID
    If Not recColEdit.NoMatch Then
      sBHolDateColumnName = recColEdit!ColumnName
    Else
      sBHolDateColumnName = vbNullString
    End If
  Else
    sBHolDateColumnName = vbNullString
  End If


  ' Set flag to state whether BHols have been setup correctly or Not
  If (iBHolRegionTableID = 0) Or _
     (sBHolRegionTableName = vbNullString) Or _
     (sBHolRegionColumnName = vbNullString) Or _
     (iBHolTableID = 0) Or _
     (sBHolTableName = vbNullString) Or _
     (sBHolDateColumnName = vbNullString) Then
    fBHolSetupOK = False
  Else
    fBHolSetupOK = True
  End If
  
  
  ' PERSONNEL STUFF

  ' Get the Personnel Table ID and Name
  iTempID = 0
  recModuleSetup.Index = "idxModuleParameter"
  recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, "Param_TablePersonnel"
  If Not recModuleSetup.NoMatch Then
    iTempID = recModuleSetup!parametervalue
    iPersonnelTableID = iTempID
    recTabEdit.Index = "idxTableID"
    recTabEdit.Seek "=", iTempID
    If Not recTabEdit.NoMatch Then
      sPersonnelTableName = recTabEdit!TableName
    Else
      iPersonnelTableID = 0
      sPersonnelTableName = vbNullString
    End If
  Else
    iPersonnelTableID = 0
    sPersonnelTableName = vbNullString
  End If

' We have these flags...
'
' fHistoricRegion (True if historic False if static)
' fHistoricWP (True if historic False if static)
' fBHolSetupOK (True if we are to use BHols False if not)
'
' We have these variables...
'
' sStaticRegionColumnName
' sHistoricRegionTableName
' sHistoricRegionColumnName
' sHistoricRegionDateColumnName
' sStaticWPColumnName
' sHistoricWPTableName
' sHistoricWPColumnName
' sHistoricWPDateColumnName
' iBHolRegionTableID
' sBHolRegionTableName
' sBHolRegionColumnName
' iBHolTableID
' sBHolTableName
' sBHolDateColumnName
' iPersonnelTableID
' sPersonnelTableName


  If fCreatedOK Then
    ' Construct the stored procedure creation string (if required).
    strGeneralSQL = "  /* Date variables used when working out the next change date for historic WP/Regions - If applicable */" & vbNewLine & _
               "  DECLARE @dTempDate                 datetime" & vbNewLine & _
               "  DECLARE @dNextChange_Region        datetime" & vbNewLine & _
               "  DECLARE @dNextChange_WP            datetime" & vbNewLine & vbNewLine

    strGeneralSQL = _
    strGeneralSQL & "  /* Date variable used to cycle through dates between start date and end date */" & vbNewLine & _
               "  DECLARE @dtCurrentDate             datetime" & vbNewLine & vbNewLine

    strGeneralSQL = _
    strGeneralSQL & "  /* The current wp/region being used in the calculation */" & vbNewLine & _
               "  DECLARE @psWorkPattern             varchar(255)" & vbNewLine & _
               "  DECLARE @psPersonnelRegion         varchar(255)" & vbNewLine & _
               "  DECLARE @psNextWorkPattern             varchar(255)" & vbNewLine & _
               "  DECLARE @psNextPersonnelRegion         varchar(255)" & vbNewLine & vbNewLine

    strGeneralSQL = _
    strGeneralSQL & "  /* Flags derived from @psWorkPattern */" & vbNewLine & _
               "  DECLARE @fWorkAM                   bit" & vbNewLine & _
               "  DECLARE @fWorkPM                   bit" & vbNewLine & _
               "  DECLARE @iDayOfWeek                int" & vbNewLine & _
               "  DECLARE @sCommandString            nvarchar(MAX)" & vbNewLine & _
               "  DECLARE @iCount                    int" & vbNewLine & _
               "  DECLARE @iBHolRegionID             int" & vbNewLine & _
               "  DECLARE @sParamDefinition          nvarchar(500)" & vbNewLine & vbNewLine

    strGeneralSQL = _
    strGeneralSQL & "  /* Initialise the result to be 0 */" & vbNewLine & _
               "  SET @pdblResult = 0" & vbNewLine & vbNewLine

    strGeneralSQL = _
    strGeneralSQL & "  /* Calculate the Absence Duration if all parameters have been provided. */" & vbNewLine & _
               "  IF (NOT @pdtStartDate IS NULL) AND (NOT @psStartSession IS NULL) AND (NOT @pdtEndDate IS NULL) AND (NOT @psEndSession IS NULL)" & vbNewLine & _
               "  BEGIN" & vbNewLine & vbNewLine & _
               "    SET @pdtStartDate = convert(datetime, convert(varchar(20), @pdtStartDate, 101))" & vbNewLine & _
               "    SET @pdtEndDate = convert(datetime, convert(varchar(20), @pdtEndDate, 101))" & vbNewLine & _
               "    SET @dtCurrentDate  = @pdtStartDate" & vbNewLine & vbNewLine

    ' IF WE ARE USING STATIC REGION AND STATIC WORKING PATTERNS THEN LETS
    ' DO THIS THE QUICK WAY !
    If (fHistoricRegion = False) And (fHistoricWP = False) Then

      strGeneralSQL = _
      strGeneralSQL & "    /* STATIC REGION AND STATIC WPATTERN BEING USED */" & vbNewLine & vbNewLine
  
      strGeneralSQL = _
      strGeneralSQL & "    /* Get The Employees Working Pattern */" & vbNewLine & _
                 "    SELECT @psWorkPattern = " & sStaticWPColumnName & " FROM " & sPersonnelTableName & " WHERE ID = @iPersonnelID" & vbNewLine & vbNewLine
  
      strGeneralSQL = _
      strGeneralSQL & "    /* Get The Employees Region*/" & vbNewLine & _
                 "    SELECT @psPersonnelRegion = " & sStaticRegionColumnName & " FROM " & sPersonnelTableName & " WHERE ID = @iPersonnelID" & vbNewLine & vbNewLine
        
      If fBHolSetupOK = True Then
    
        strGeneralSQL = _
        strGeneralSQL & "    /* Get The Region ID for the persons ID*/" & vbNewLine & _
                    "    SELECT @iBholRegionID = ID FROM " & sBHolRegionTableName & " WHERE " & sBHolRegionColumnName & " = @psPersonnelRegion" & vbNewLine & vbNewLine
      End If
   
      strGeneralSQL = _
      strGeneralSQL & "    WHILE @dtCurrentDate <= @pdtEndDate" & vbNewLine & _
                 "    BEGIN" & vbNewLine & _
                 "" & vbNewLine & _
                 "      /* Check if the current date is a work day. */" & vbNewLine & _
                 "      SET @fWorkAM = 0" & vbNewLine & _
                 "      SET @fWorkPM = 0" & vbNewLine & _
                 "      SET @iDayOfWeek = DATEPART(weekday, @dtCurrentDate)" & vbNewLine & vbNewLine
                 
      For iLoop = 1 To 7
        strGeneralSQL = strGeneralSQL & _
          "      IF @iDayOfWeek = " & CStr(iLoop) & vbNewLine & _
          "      BEGIN" & vbNewLine & _
          "        IF LEN(SUBSTRING(@psWorkPattern, " & CStr((iLoop * 2) - 1) & ", 1)) > 0" & vbNewLine & _
          "        BEGIN" & vbNewLine & _
          "          SET @fWorkAM = 1" & vbNewLine & _
          "        END" & vbNewLine & _
          "        IF LEN(SUBSTRING(@psWorkPattern, " & CStr(iLoop * 2) & ", 1)) > 0" & vbNewLine & _
          "        BEGIN" & vbNewLine & _
          "          SET @fWorkPM = 1" & vbNewLine & _
          "        END" & vbNewLine & _
          "      END" & vbNewLine
      Next iLoop
  
      strGeneralSQL = _
      strGeneralSQL & vbNewLine & "      IF (@fWorkAM = 1) OR (@fWorkPM = 1)" & vbNewLine & _
                 "      BEGIN" & vbNewLine

      If fBHolSetupOK = True Then
                 
        ' BHOLS ARE BEING USED
                 
        strGeneralSQL = _
        strGeneralSQL & "        /* Check that the current date is not a company holiday. */" & vbNewLine & _
                   "        SELECT @iCount = COUNT(" & sBHolDateColumnName & ")" & vbNewLine & _
                   "        FROM " & sBHolTableName & " " & vbNewLine & _
                   "        WHERE " & sBHolDateColumnName & " = @dtCurrentDate " & vbNewLine & _
                   "        AND " & sBHolTableName & ".ID_" & iBHolRegionTableID & " = @iBHolRegionID" & vbNewLine & vbNewLine
        strGeneralSQL = _
        strGeneralSQL & "        IF @iCount = 0" & vbNewLine & _
                   "        BEGIN" & vbNewLine & _
                   "          IF (@dtCurrentDate = @pdtStartDate) AND (@dtCurrentDate = @pdtEndDate)" & vbNewLine & _
                   "          BEGIN" & vbNewLine & _
                   "            IF ((@fWorkAM = 1) AND (UPPER(@psStartSession) = 'AM')) SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "            IF ((@fWorkPM = 1) AND (UPPER(@psEndSession) = 'PM')) SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "          END" & vbNewLine & _
                   "          ELSE" & vbNewLine & _
                   "          BEGIN" & vbNewLine & _
                   "            IF @dtCurrentDate = @pdtStartDate" & vbNewLine & _
                   "            BEGIN" & vbNewLine & _
                   "              IF ((@fWorkAM = 1) AND (UPPER(@psStartSession) = 'AM')) SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "              IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "            END" & vbNewLine & _
                   "            ELSE" & vbNewLine & _
                   "            BEGIN" & vbNewLine & _
                   "              IF @dtCurrentDate = @pdtEndDate" & vbNewLine & _
                   "              BEGIN" & vbNewLine & _
                   "                IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "                IF ((@fWorkPM = 1) AND (UPPER(@psEndSession) = 'PM'))  SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "              END"
        strGeneralSQL = _
        strGeneralSQL & "              ELSE" & vbNewLine & _
                   "              BEGIN" & vbNewLine & _
                   "                IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "                IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "              END" & vbNewLine & _
                   "            END" & vbNewLine & _
                   "          END" & vbNewLine & _
                   "        END" & vbNewLine & _
                   "      END" & vbNewLine
      Else
      
        ' BHOLS ARE NOT BEING USED
      
        strGeneralSQL = _
        strGeneralSQL & "        IF (@dtCurrentDate = @pdtStartDate) AND (@dtCurrentDate = @pdtEndDate)" & vbNewLine & _
                   "        BEGIN" & vbNewLine & _
                   "          IF ((@fWorkAM = 1) AND (UPPER(@psStartSession) = 'AM')) SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "          IF ((@fWorkPM = 1) AND (UPPER(@psEndSession) = 'PM')) SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "        END" & vbNewLine & _
                   "        ELSE" & vbNewLine & _
                   "        BEGIN" & vbNewLine & _
                   "          IF @dtCurrentDate = @pdtStartDate" & vbNewLine & _
                   "          BEGIN" & vbNewLine
        strGeneralSQL = _
        strGeneralSQL & "            IF ((@fWorkAM = 1) AND (UPPER(@psStartSession) = 'AM')) SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "            IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "          END" & vbNewLine & _
                   "          ELSE" & vbNewLine & _
                   "          BEGIN" & vbNewLine & _
                   "            IF @dtCurrentDate = @pdtEndDate" & vbNewLine & _
                   "            BEGIN" & vbNewLine & _
                   "              IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "              IF ((@fWorkPM = 1) AND (UPPER(@psEndSession) = 'PM'))  SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "            END" & vbNewLine & _
                   "            ELSE" & vbNewLine & _
                   "            BEGIN" & vbNewLine & _
                   "              IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "              IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "            END" & vbNewLine & _
                   "          END" & vbNewLine & _
                   "        END" & vbNewLine & _
                   "      END" & vbNewLine '& _
                   "    END" & vbNewline
      
      End If
  
    Else
    
      ' HISTORIC REGION OR WORKING PATTERN OR BOTH
      
      strGeneralSQL = _
      strGeneralSQL & "    /* HISTORIC REGION OR WPATTERN OR BOTH ARE BEING USED */" & vbNewLine & vbNewLine
  
      If fHistoricRegion = True Then
        strGeneralSQL = strGeneralSQL & _
          "        SELECT TOP 1 @psNextPersonnelRegion = " & sHistoricRegionColumnName & vbNewLine & _
          "        FROM " & sHistoricRegionTableName & vbNewLine & _
          "        WHERE " & sHistoricRegionDateColumnName & " <= @dtCurrentDate" & vbNewLine & _
          "        AND ID_" & iPersonnelTableID & " = @iPersonnelID" & vbNewLine & _
          "        ORDER BY " & sHistoricRegionDateColumnName & " DESC" & vbNewLine & vbNewLine
      End If
      
      If fHistoricWP = True Then
        strGeneralSQL = strGeneralSQL & _
          "        SELECT TOP 1 @psNextWorkPattern = " & sHistoricWPColumnName & vbNewLine & _
          "        FROM " & sHistoricWPTableName & vbNewLine & _
          "        WHERE " & sHistoricWPDateColumnName & " <= @dtCurrentDate" & vbNewLine & _
          "        AND ID_" & iPersonnelTableID & " = @iPersonnelID" & vbNewLine & _
          "        ORDER BY " & sHistoricWPDateColumnName & " DESC" & vbNewLine & vbNewLine
      End If

      strGeneralSQL = _
      strGeneralSQL & "    WHILE @dtCurrentDate <= @pdtEndDate" & vbNewLine & _
                 "    BEGIN" & vbNewLine & vbNewLine
                 
      If fHistoricRegion = True Then
      
        strGeneralSQL = _
        strGeneralSQL & "      /* We are using a historic region, so ensure we have the right region for the @dCurrentDate */" & vbNewLine & vbNewLine & _
                   "      /* Only bother checking we have the right region if we dont know the nxt chg date or the current date is equal to nxt chg date */" & vbNewLine & _
                   "      IF (@dnextchange_region IS NULL) OR ((@dtCurrentDate >= @dNextChange_Region) And (@dtCurrentDate <> '12/31/9999'))" & vbNewLine & _
                   "      BEGIN" & vbNewLine & vbNewLine & _
                   "        /* Get The Employees Region For @dCurrentDate */" & vbNewLine & _
                   "        SET @psPersonnelRegion = @psNextPersonnelRegion" & vbNewLine & vbNewLine
                   
        If fBHolSetupOK = True Then
          strGeneralSQL = _
          strGeneralSQL & "        /* Get the Region ID for the persons Region */" & vbNewLine & _
                     "        SELECT @iBHolRegionID = ID" & vbNewLine & _
                     "        FROM " & sBHolRegionTableName & vbNewLine & _
                     "        WHERE " & sBHolRegionColumnName & " = @psPersonnelRegion" & vbNewLine & vbNewLine
        End If

        strGeneralSQL = strGeneralSQL & _
          "        /* Get the date of next change for the Region */" & vbNewLine & _
          "        SET @dTempDate = null" & vbNewLine & _
          "        SET @psNextPersonnelRegion = null" & vbNewLine & _
          "        SELECT TOP 1 @dTempDate = " & sHistoricRegionDateColumnName & "," & vbNewLine & _
          "          @psNextPersonnelRegion = " & sHistoricRegionColumnName & vbNewLine & _
          "        FROM " & sHistoricRegionTableName & vbNewLine & _
          "        WHERE " & sHistoricRegionDateColumnName & " > @dtCurrentDate" & vbNewLine & _
          "        AND ID_" & iPersonnelTableID & " = @iPersonnelID" & vbNewLine & _
          "        ORDER BY " & sHistoricRegionDateColumnName & " ASC" & vbNewLine & vbNewLine
            
        strGeneralSQL = _
        strGeneralSQL & "        IF @dTempDate IS NULL" & vbNewLine & _
                   "        BEGIN" & vbNewLine & _
                   "          SET @dNextChange_Region = '12/31/9999'" & vbNewLine & _
                   "        END" & vbNewLine & _
                   "        ELSE" & vbNewLine & _
                   "        BEGIN" & vbNewLine & _
                   "          SET @dNextChange_Region = @dTempDate" & vbNewLine & _
                   "        END" & vbNewLine & _
                   "      END" & vbNewLine & vbNewLine

      Else
        
        strGeneralSQL = _
        strGeneralSQL & "      /* We are using a static region, so get it */" & vbNewLine & vbNewLine & _
                   "      SELECT @psPersonnelRegion = " & sStaticRegionColumnName & vbNewLine & _
                   "      FROM " & sPersonnelTableName & vbNewLine & _
                   "      WHERE ID = @iPersonnelID" & vbNewLine & vbNewLine
        
        strGeneralSQL = _
        strGeneralSQL & "      /* Get the Region ID for the persons Region */" & vbNewLine & vbNewLine & _
                   "      SELECT @iBHolRegionID = ID" & vbNewLine & _
                   "      FROM " & sBHolRegionTableName & vbNewLine & _
                   "      WHERE " & sBHolRegionColumnName & " = @psPersonnelRegion" & vbNewLine & vbNewLine
        
      End If
              
      If fHistoricWP = True Then
        
        strGeneralSQL = _
        strGeneralSQL & "      /* We are using historic working pattern */" & vbNewLine & _
                   "      IF (@dnextchange_WP IS NULL) OR ((@dtCurrentDate >= @dNextChange_WP) And (@dtCurrentDate <> '12/31/9999'))" & vbNewLine & _
                   "      BEGIN" & vbNewLine & vbNewLine & _
                   "        /* Get The Employees WP For @dCurrentDate */" & vbNewLine & _
                   "        SET @psWorkPattern = @psNextWorkPattern" & vbNewLine & vbNewLine
                   
        strGeneralSQL = _
        strGeneralSQL & "        /* Get The next change date for WP */" & vbNewLine & _
          "        SET @dTempDate = null" & vbNewLine & _
          "        SET @psNextWorkPattern = null" & vbNewLine & _
          "        SELECT TOP 1 @dTempDate = " & sHistoricWPDateColumnName & "," & vbNewLine & _
          "          @psNextWorkPattern = " & sHistoricWPColumnName & vbNewLine & _
          "        FROM " & sHistoricWPTableName & vbNewLine & _
          "        WHERE " & sHistoricWPDateColumnName & " > @dtCurrentDate" & vbNewLine & _
          "        AND ID_" & iPersonnelTableID & " = @iPersonnelID" & vbNewLine & _
          "        ORDER BY " & sHistoricWPDateColumnName & " ASC" & vbNewLine & vbNewLine
          
        strGeneralSQL = _
        strGeneralSQL & "        IF @dTempDate IS NULL" & vbNewLine & _
                   "        BEGIN" & vbNewLine & _
                   "          SET @dNextChange_WP = '12/31/9999'" & vbNewLine & _
                   "        END" & vbNewLine & _
                   "        ELSE" & vbNewLine & _
                   "        BEGIN" & vbNewLine & _
                   "          SET @dNextChange_WP = @dTempDate" & vbNewLine & _
                   "        END" & vbNewLine & _
                   "      END" & vbNewLine & vbNewLine
        
      Else
      
        strGeneralSQL = _
        strGeneralSQL & "      /* We are using a static wp, so get it */" & vbNewLine & _
                   "      SELECT @psWorkPattern = " & sStaticWPColumnName & vbNewLine & _
                   "      FROM " & sPersonnelTableName & vbNewLine & _
                   "      WHERE ID = @iPersonnelID" & vbNewLine & vbNewLine
        
      End If
      
      strGeneralSQL = _
      strGeneralSQL & "      /* Check if the current date is a work day. */" & vbNewLine & _
                 "      SET @fWorkAM = 0" & vbNewLine & _
                 "      SET @fWorkPM = 0" & vbNewLine & _
                 "      SET @iDayOfWeek = DATEPART(weekday, @dtCurrentDate)" & vbNewLine
                 
      For iLoop = 1 To 7
        strGeneralSQL = strGeneralSQL & _
          "      IF @iDayOfWeek = " & CStr(iLoop) & vbNewLine & _
          "      BEGIN" & vbNewLine & _
          "        IF LEN(SUBSTRING(@psWorkPattern, " & CStr((iLoop * 2) - 1) & ", 1)) > 0" & vbNewLine & _
          "        BEGIN" & vbNewLine & _
          "          SET @fWorkAM = 1" & vbNewLine & _
          "        END" & vbNewLine & _
          "        IF LEN(SUBSTRING(@psWorkPattern, " & CStr(iLoop * 2) & ", 1)) > 0" & vbNewLine & _
          "        BEGIN" & vbNewLine & _
          "          SET @fWorkPM = 1" & vbNewLine & _
          "        END" & vbNewLine & _
          "      END" & vbNewLine
      Next iLoop

      strGeneralSQL = _
      strGeneralSQL & vbNewLine & "      IF (@fWorkAM = 1) OR (@fWorkPM = 1)" & vbNewLine & _
                 "      BEGIN" & vbNewLine & vbNewLine

      If fBHolSetupOK = True Then
      
        ' BHOLS ARE BEING USED
      
        strGeneralSQL = _
        strGeneralSQL & "        /* Check that the current date is not a company holiday. */" & vbNewLine & _
                   "        SELECT @iCount = COUNT(" & sBHolDateColumnName & ")" & vbNewLine & _
                   "        FROM " & sBHolTableName & vbNewLine & _
                   "        WHERE " & sBHolDateColumnName & " = @dtCurrentDate" & vbNewLine & _
                   "        AND " & sBHolTableName & ".ID_" & iBHolRegionTableID & " = @iBHolRegionID" & vbNewLine & vbNewLine
                   
        strGeneralSQL = _
        strGeneralSQL & "        IF @iCount = 0" & vbNewLine & _
                   "        BEGIN" & vbNewLine & _
                   "          IF (@dtCurrentDate = @pdtStartDate) AND (@dtCurrentDate = @pdtEndDate)" & vbNewLine & _
                   "          BEGIN" & vbNewLine & _
                   "            IF ((@fWorkAM = 1) AND (UPPER(@psStartSession) = 'AM')) SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "            IF ((@fWorkPM = 1) AND (UPPER(@psEndSession) = 'PM')) SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "          END" & vbNewLine & _
                   "          ELSE" & vbNewLine & _
                   "          BEGIN" & vbNewLine & _
                   "            IF @dtCurrentDate = @pdtStartDate" & vbNewLine & _
                   "            BEGIN" & vbNewLine & _
                   "              IF ((@fWorkAM = 1) AND (UPPER(@psStartSession) = 'AM')) SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "              IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "            END" & vbNewLine & _
                   "            ELSE" & vbNewLine & _
                   "            BEGIN" & vbNewLine & _
                   "              IF @dtCurrentDate = @pdtEndDate" & vbNewLine & _
                   "              BEGIN" & vbNewLine & _
                   "                IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "                IF ((@fWorkPM = 1) AND (UPPER(@psEndSession) = 'PM'))  SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "              END"
        strGeneralSQL = _
        strGeneralSQL & "              ELSE" & vbNewLine & _
                   "              BEGIN" & vbNewLine & _
                   "                IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "                IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "              END /* End for 3rd check */" & vbNewLine & _
                   "            END /* End for 2nd check */" & vbNewLine & _
                   "          END /* End for 1st check */" & vbNewLine & _
                   "        END /* End for if @iCount = 0 */" & vbNewLine & _
                   "      END /* End for if we do work either am or pm or both */" & vbNewLine & vbNewLine
      
      Else
      
        ' BHOLS ARE NOT BEING USED
      
        strGeneralSQL = _
        strGeneralSQL & "        IF (@dtCurrentDate = @pdtStartDate) AND (@dtCurrentDate = @pdtEndDate)" & vbNewLine & _
                   "        BEGIN" & vbNewLine & _
                   "          IF ((@fWorkAM = 1) AND (UPPER(@psStartSession) = 'AM')) SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "          IF ((@fWorkPM = 1) AND (UPPER(@psEndSession) = 'PM')) SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "        END" & vbNewLine & _
                   "        ELSE" & vbNewLine & _
                   "        BEGIN" & vbNewLine & _
                   "          IF @dtCurrentDate = @pdtStartDate" & vbNewLine & _
                   "          BEGIN" & vbNewLine
        strGeneralSQL = _
        strGeneralSQL & "            IF ((@fWorkAM = 1) AND (UPPER(@psStartSession) = 'AM')) SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "            IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "          END" & vbNewLine & _
                   "          ELSE" & vbNewLine & _
                   "          BEGIN" & vbNewLine & _
                   "            IF @dtCurrentDate = @pdtEndDate" & vbNewLine & _
                   "            BEGIN" & vbNewLine & _
                   "              IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "              IF ((@fWorkPM = 1) AND (UPPER(@psEndSession) = 'PM'))  SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "            END" & vbNewLine & _
                   "            ELSE" & vbNewLine & _
                   "            BEGIN" & vbNewLine & _
                   "              IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "              IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                   "            END /* End for 3rd Check */" & vbNewLine & _
                   "          END /* End for 2nd Check */" & vbNewLine & _
                   "        END /* End for 1st Check */" & vbNewLine & _
                   "      END /* End for if they work am, pm or both */" & vbNewLine '& _
                   "    END" & vbNewline
      End If
      
    End If

    ' Extra calculations if we are creating the absence breakdown calculation
    If pbCreateBreakdown = True Then
      strGeneralSQL = _
      strGeneralSQL & "     /* Absence Breakdown extras */ " & vbNewLine & _
                 "     IF @iDayOfWeek = 1 SET @pfSunTotal = @pfSunTotal + @pdblResult" & vbNewLine & _
                 "     IF @iDayOfWeek = 2 SET @pfMonTotal = @pfMonTotal + @pdblResult" & vbNewLine & _
                 "     IF @iDayOfWeek = 3 SET @pfTueTotal = @pfTueTotal + @pdblResult" & vbNewLine & _
                 "     IF @iDayOfWeek = 4 SET @pfWedTotal = @pfWedTotal + @pdblResult" & vbNewLine & _
                 "     IF @iDayOfWeek = 5 SET @pfThuTotal = @pfThuTotal + @pdblResult" & vbNewLine & _
                 "     IF @iDayOfWeek = 6 SET @pfFriTotal = @pfFriTotal + @pdblResult" & vbNewLine & _
                 "     IF @iDayOfWeek = 7 SET @pfSatTotal = @pfSatTotal + @pdblResult" & vbNewLine & vbNewLine & _
                 "     /* Reset the running duration */" & vbNewLine & _
                 "     SET @pdblResult = 0 " + vbNewLine & _
                 "     /* End absence breakdown extras */" & vbNewLine & vbNewLine

    End If
        
    strGeneralSQL = _
    strGeneralSQL & "      /* Move onto the next date. */" & vbNewLine & _
               "      SET @dtCurrentDate = @dtCurrentDate + 1" & vbNewLine & _
               "    END /* End for while Current Date <= End Date */" & vbNewLine & _
               "  END /* End for if all parameters have been passed */" & vbNewLine
  
  
    ' Create the udf and sp
    sProcSQL = strProcHeader & strGeneralSQL & strEndProc
    sUDFSQL = strUDFHeader & strGeneralSQL & strEndUDF
    
    gADOCon.Execute sProcSQL, , adExecuteNoRecords
    If Not pbCreateBreakdown And gbEnableUDFFunctions Then gADOCon.Execute sUDFSQL, , adExecuteNoRecords
  
  End If
  
TidyUpAndExit:
  
  CreateAbsenceDurationStoredProcedure = fCreatedOK
  Exit Function
  
ErrorTrap:
  fCreatedOK = False
  OutputError "Error creating Absence Duration stored procedure (Absence)"
  Resume TidyUpAndExit
  
End Function


Private Function CreateAbsenceBetween2DatesStoredProcedure() As Boolean
  On Error GoTo ErrorTrap

  Dim fCreatedOK As Boolean
  Dim sSQL As String
  Dim iLoop As Integer
  
  Dim strGenericSQL As String
  Dim sAbsenceBetweenProc As String
  Dim sAbsenceBetweenUDF As String
  Dim sBradfordUDF As String
  
  Dim lngTempID As Long
  Dim fValidConfiguration As Boolean
  Dim fBHolSetupOK  As Boolean
  Dim fHistoricRegion As Boolean
  Dim fHistoricWP As Boolean
  
  ' Absence table parameters
  Dim lngAbsTableID As Long
  Dim sAbsTableName As String
  Dim lngAbsColStartDateID As Long
  Dim sAbsColStartDateName As String
  Dim lngAbsColEndDateID As Long
  Dim sAbsColEndDateName As String
  Dim lngAbsColTypeID As Long
  Dim sAbsColTypeName As String
  Dim lngAbsColStartPeriodID As Long
  Dim sAbsColStartPeriodName As String
  Dim lngAbsColEndPeriodID As Long
  Dim sAbsColEndPeriodName As String
  Dim lngContinuousColumnID As Long
  Dim sContinuousColumnName As String

  ' Personnel table parameters
  Dim lngPersonnelTableID As Integer
  Dim sPersonnelTableName As String
  
  ' Bank Holiday Region (Primary) table parameters
  Dim lngBHolRegionTableID As Long
  Dim sBHolRegionTableName As String
  Dim lngBHolRegionColumnID As Long
  Dim sBHolRegionColumnName As String
  
  ' The Bank Holiday Instance (Child) table parameters
  Dim lngBHolTableID As Long
  Dim sBHolTableName As String
  Dim lngBHolDateColumnID As Long
  Dim sBHolDateColumnName As String
  
  Dim lngStaticRegionColumnID As Long
  Dim sStaticRegionColumnName As String
  
  Dim lngHistoricRegionTableID As Long
  Dim sHistoricRegionTableName As String
  Dim lngHistoricRegionColumnID As Long
  Dim sHistoricRegionColumnName As String
  Dim lngHistoricRegionDateColumnID As Long
  Dim sHistoricRegionDateColumnName As String
  
  Dim lngStaticWPColumnID As Long
  Dim sStaticWPColumnName As String
  
  Dim lngHistoricWPTableID As Long
  Dim sHistoricWPTableName As String
  Dim lngHistoricWPColumnID As Long
  Dim sHistoricWPColumnName As String
  Dim lngHistoricWPDateColumnID As Long
  Dim sHistoricWPDateColumnName As String
  
  ' Drop any existing stored procedure and user defined function.
  'fCreatedOK = DropAbsenceBetween2DatesStoredProcedure
  fCreatedOK = DropProcedure("sp_ASRFn_AbsenceBetweenTwoDates")
  
  If gbEnableUDFFunctions Then
    fCreatedOK = DropFunction("udf_ASRFn_AbsenceBetweenTwoDates")
    fCreatedOK = DropFunction("udf_ASRFn_BradfordFactor")
  End If

  If fCreatedOK Then
    fValidConfiguration = True
    
    ' Get the Absence Table ID and Name
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETABLE
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
      lngAbsTableID = lngTempID
      recTabEdit.Index = "idxTableID"
      recTabEdit.Seek "=", lngTempID
      If Not recTabEdit.NoMatch Then
        sAbsTableName = recTabEdit!TableName
      Else
        lngAbsTableID = 0
        sAbsTableName = vbNullString
      End If
    Else
      lngAbsTableID = 0
      sAbsTableName = vbNullString
    End If
      
    ' Set the Absence Start Date column variable
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESTARTDATE
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
      lngAbsColStartDateID = lngTempID
      recColEdit.Index = "idxColumnID"
      recColEdit.Seek "=", lngTempID
      If Not recColEdit.NoMatch Then
        sAbsColStartDateName = recColEdit!ColumnName
      Else
        lngAbsColStartDateID = 0
        sAbsColStartDateName = vbNullString
      End If
    Else
      lngAbsColStartDateID = 0
      sAbsColStartDateName = vbNullString
    End If
    
    ' Set the absence end date column variable
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEENDDATE
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
      lngAbsColEndDateID = lngTempID
      recColEdit.Index = "idxColumnID"
      recColEdit.Seek "=", lngTempID
      If Not recColEdit.NoMatch Then
        sAbsColEndDateName = recColEdit!ColumnName
      Else
        lngAbsColEndDateID = 0
        sAbsColEndDateName = vbNullString
      End If
    Else
      lngAbsColEndDateID = 0
      sAbsColEndDateName = vbNullString
    End If
    
    ' Set the absence type column variable
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPE
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
      lngAbsColTypeID = lngTempID
      recColEdit.Index = "idxColumnID"
      recColEdit.Seek "=", lngTempID
      If Not recColEdit.NoMatch Then
        sAbsColTypeName = recColEdit!ColumnName
      Else
        lngAbsColTypeID = 0
        sAbsColTypeName = vbNullString
      End If
    Else
      lngAbsColTypeID = 0
      sAbsColTypeName = vbNullString
    End If

    ' Set the absence start period column variable
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESTARTSESSION
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
      lngAbsColStartPeriodID = lngTempID
      recColEdit.Index = "idxColumnID"
      recColEdit.Seek "=", lngTempID
      If Not recColEdit.NoMatch Then
        sAbsColStartPeriodName = recColEdit!ColumnName
      Else
        lngAbsColStartPeriodID = 0
        sAbsColStartPeriodName = vbNullString
      End If
    Else
      lngAbsColStartPeriodID = 0
      sAbsColStartPeriodName = vbNullString
    End If

    ' Set the absence end period column variable
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEENDSESSION
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
      lngAbsColEndPeriodID = lngTempID
      recColEdit.Index = "idxColumnID"
      recColEdit.Seek "=", lngTempID
      If Not recColEdit.NoMatch Then
        sAbsColEndPeriodName = recColEdit!ColumnName
      Else
        lngAbsColEndPeriodID = 0
        sAbsColEndPeriodName = vbNullString
      End If
    Else
      lngAbsColEndPeriodID = 0
      sAbsColEndPeriodName = vbNullString
    End If
      
    ' Set the absence continuous column variable
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCECONTINUOUS
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
      lngContinuousColumnID = lngTempID
      recColEdit.Index = "idxColumnID"
      recColEdit.Seek "=", lngTempID
      If Not recColEdit.NoMatch Then
        sContinuousColumnName = recColEdit!ColumnName
      Else
        lngContinuousColumnID = 0
        sContinuousColumnName = vbNullString
      End If
    Else
      lngContinuousColumnID = 0
      sContinuousColumnName = vbNullString
    End If
    
    ' If any of the absence tables/columns have not been defined, then we
    ' might as well stop the SP here, returning 0.
    If (lngAbsTableID = 0) Or _
      (lngContinuousColumnID = 0) Or _
      (lngAbsColStartDateID = 0) Or _
      (lngAbsColEndDateID = 0) Or _
      (lngAbsColStartPeriodID = 0) Or _
      (lngAbsColEndPeriodID = 0) Or _
      (lngAbsColTypeID = 0) Then
      fValidConfiguration = False
    End If
    
    ' Get the bhol region table id
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLREGIONTABLE
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
      lngBHolRegionTableID = lngTempID
      recTabEdit.Index = "idxTableID"
      recTabEdit.Seek "=", lngTempID
      If Not recTabEdit.NoMatch Then
        sBHolRegionTableName = recTabEdit!TableName
      Else
        lngBHolRegionTableID = 0
        sBHolRegionTableName = vbNullString
      End If
    Else
      lngBHolRegionTableID = 0
      sBHolRegionTableName = vbNullString
    End If
    
    ' Get the BHol Region column in the bhol region table
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLREGION
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
      lngBHolRegionColumnID = lngTempID
      recColEdit.Index = "idxColumnID"
      recColEdit.Seek "=", lngTempID
      If Not recColEdit.NoMatch Then
        sBHolRegionColumnName = recColEdit!ColumnName
      Else
        lngBHolRegionColumnID = 0
        sBHolRegionColumnName = vbNullString
      End If
    Else
      lngBHolRegionColumnID = 0
      sBHolRegionColumnName = vbNullString
    End If
    
    ' Get the bhol table id
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLTABLE
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
      lngBHolTableID = lngTempID
      recTabEdit.Index = "idxTableID"
      recTabEdit.Seek "=", lngTempID
      If Not recTabEdit.NoMatch Then
        sBHolTableName = recTabEdit!TableName
      Else
        lngBHolTableID = 0
        sBHolTableName = vbNullString
      End If
    Else
      lngBHolTableID = 0
      sBHolTableName = vbNullString
    End If
    
    ' Get the name of the BHol Date column
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLDATE
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
      lngBHolDateColumnID = lngTempID
      recColEdit.Index = "idxColumnID"
      recColEdit.Seek "=", lngTempID
      If Not recColEdit.NoMatch Then
        sBHolDateColumnName = recColEdit!ColumnName
      Else
        lngBHolDateColumnID = 0
        sBHolDateColumnName = vbNullString
      End If
    Else
      lngBHolDateColumnID = 0
      sBHolDateColumnName = vbNullString
    End If
    
    ' Check whether BHols have been setup correctly or not.
    fBHolSetupOK = (lngBHolRegionTableID > 0) And _
      (lngBHolRegionColumnID > 0) And _
      (lngBHolTableID > 0) And _
      (lngBHolDateColumnID > 0)
    
    ' Set the Personnel table ID variable
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
      lngPersonnelTableID = lngTempID
      recTabEdit.Index = "idxTableID"
      recTabEdit.Seek "=", lngTempID
      If Not recTabEdit.NoMatch Then
        sPersonnelTableName = recTabEdit!TableName
      Else
        lngPersonnelTableID = 0
        sPersonnelTableName = vbNullString
      End If
    Else
      lngPersonnelTableID = 0
      sPersonnelTableName = vbNullString
    End If
    
    ' Get the region module stuff and work out static or historic
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_REGION
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
      lngStaticRegionColumnID = lngTempID
      recColEdit.Index = "idxColumnID"
      recColEdit.Seek "=", lngTempID
      If Not recColEdit.NoMatch Then
        sStaticRegionColumnName = recColEdit!ColumnName
      Else
        lngStaticRegionColumnID = 0
        sStaticRegionColumnName = vbNullString
      End If
    Else
      lngStaticRegionColumnID = 0
      sStaticRegionColumnName = vbNullString
    End If
    
    ' Get the Region Setup - Historic Region
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HREGIONTABLE
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
      lngHistoricRegionTableID = lngTempID
      recTabEdit.Index = "idxTableID"
      recTabEdit.Seek "=", lngTempID
      If Not recTabEdit.NoMatch Then
        sHistoricRegionTableName = recTabEdit!TableName
      Else
        lngHistoricRegionTableID = 0
        sHistoricRegionTableName = vbNullString
      End If
    Else
      lngHistoricRegionTableID = 0
      sHistoricRegionTableName = vbNullString
    End If
    
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HREGIONFIELD
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
      lngHistoricRegionColumnID = lngTempID
      recColEdit.Index = "idxColumnID"
      recColEdit.Seek "=", lngTempID
      If Not recColEdit.NoMatch Then
        sHistoricRegionColumnName = recColEdit!ColumnName
      Else
        lngHistoricRegionColumnID = 0
        sHistoricRegionColumnName = vbNullString
      End If
    Else
      lngHistoricRegionColumnID = 0
      sHistoricRegionColumnName = vbNullString
    End If
  
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HREGIONDATE
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
      lngHistoricRegionDateColumnID = lngTempID
      recColEdit.Index = "idxColumnID"
      recColEdit.Seek "=", lngTempID
      If Not recColEdit.NoMatch Then
        sHistoricRegionDateColumnName = recColEdit!ColumnName
      Else
        lngHistoricRegionDateColumnID = 0
        sHistoricRegionDateColumnName = vbNullString
      End If
    Else
      lngHistoricRegionDateColumnID = 0
      sHistoricRegionDateColumnName = vbNullString
    End If
  
    ' Set flag to indicate what type of regions we are to use.
    fHistoricRegion = False
    If lngStaticRegionColumnID = 0 Then
      If (lngHistoricRegionTableID = 0) Or _
        (lngHistoricRegionColumnID = 0) Or _
        (lngHistoricRegionDateColumnID = 0) Then
        fValidConfiguration = False
      Else
        fHistoricRegion = True
      End If
    End If
  
    ' Get the WP Setup - Static WP
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_WORKINGPATTERN
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
      lngStaticWPColumnID = lngTempID
      recColEdit.Index = "idxColumnID"
      recColEdit.Seek "=", lngTempID
      If Not recColEdit.NoMatch Then
        sStaticWPColumnName = recColEdit!ColumnName
      Else
        lngStaticWPColumnID = 0
        sStaticWPColumnName = vbNullString
      End If
    Else
      lngStaticWPColumnID = 0
      sStaticWPColumnName = vbNullString
    End If
  
    ' Get the Region Setup - Historic WP
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HWORKINGPATTERNTABLE
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
      lngHistoricWPTableID = lngTempID
      recTabEdit.Index = "idxTableID"
      recTabEdit.Seek "=", lngTempID
      If Not recTabEdit.NoMatch Then
        sHistoricWPTableName = recTabEdit!TableName
      Else
        lngHistoricWPTableID = 0
        sHistoricWPTableName = vbNullString
      End If
    Else
      lngHistoricWPTableID = 0
      sHistoricWPTableName = vbNullString
    End If
  
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HWORKINGPATTERNFIELD
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
      lngHistoricWPColumnID = lngTempID
      recColEdit.Index = "idxColumnID"
      recColEdit.Seek "=", lngTempID
      If Not recColEdit.NoMatch Then
        sHistoricWPColumnName = recColEdit!ColumnName
      Else
        lngHistoricWPColumnID = 0
        sHistoricWPColumnName = vbNullString
      End If
    Else
      lngHistoricWPColumnID = 0
      sHistoricWPColumnName = vbNullString
    End If
  
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HWORKINGPATTERNDATE
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
      lngHistoricWPDateColumnID = lngTempID
      recColEdit.Index = "idxColumnID"
      recColEdit.Seek "=", lngTempID
      If Not recColEdit.NoMatch Then
        sHistoricWPDateColumnName = recColEdit!ColumnName
      Else
        lngHistoricWPDateColumnID = 0
        sHistoricWPDateColumnName = vbNullString
      End If
    Else
      lngHistoricWPDateColumnID = 0
      sHistoricWPDateColumnName = vbNullString
    End If
  
    ' Check what type of wp we are to use.
    fHistoricWP = False
    If lngStaticWPColumnID = 0 Then
      If (lngHistoricWPTableID = 0) Or _
        (lngHistoricWPColumnID = 0) Or _
        (lngHistoricWPDateColumnID = 0) Then
        fValidConfiguration = False
      Else
        fHistoricWP = True
      End If
    End If
    
    ' Construct the stored procedure creation string (if required).
    sAbsenceBetweenProc = "/* ------------------------------------------------ */" & vbNewLine & _
               "/* HR Pro Absence module stored procedure.          */" & vbNewLine & _
               "/* Automatically generated by the System manager.   */" & vbNewLine & _
               "/* ------------------------------------------------ */" & vbNewLine & _
               "CREATE PROCEDURE dbo.sp_ASRFn_AbsenceBetweenTwoDates (" & vbNewLine & _
               "    @pdblResult float OUTPUT, /* Result to be passed back     */" & vbNewLine & _
               "    @pdtStartDate datetime, /* Start of the report period   */" & vbNewLine & _
               "    @pdtEndDate datetime, /* End of the report period     */" & vbNewLine & _
               "    @psAbsenceTypes varchar(255), /* Abs Type we are looking for  */" & vbNewLine & _
               "    @piPersonnelID integer /* Employees ID */" & vbNewLine & _
               ")" & vbNewLine & _
               "AS" & vbNewLine & _
               "BEGIN" & vbNewLine & _
               "   DECLARE @pdToday datetime" & vbNewLine & _
               "   SET @pdToday = GETDATE()" & vbNewLine

    sAbsenceBetweenUDF = "/* ------------------------------------------------ */" & vbNewLine & _
               "/* HR Pro Absence Between 2 Dates user defined function.     */" & vbNewLine & _
               "/* Automatically generated by the System manager.   */" & vbNewLine & _
               "/* ------------------------------------------------ */" & vbNewLine & _
               "CREATE FUNCTION dbo.udf_ASRFn_AbsenceBetweenTwoDates (" & vbNewLine & _
               "    @pdtStartDate datetime, /* Start of the report period   */" & vbNewLine & _
               "    @pdtEndDate datetime, /* End of the report period     */" & vbNewLine & _
               "    @psAbsenceTypes varchar(255), /* Abs Type we are looking for  */" & vbNewLine & _
               "    @piPersonnelID integer, /* Employees ID */" & vbNewLine & _
               "    @pdToday datetime       /* Pass in todays date because getdate() does not work in udfs */" & vbNewLine & _
               ")" & vbNewLine & _
               "RETURNS float" & vbNewLine & _
               "AS" & vbNewLine & _
               "BEGIN" & vbNewLine & _
               "    DECLARE @pdblResult float" & vbNewLine & vbNewLine

    sBradfordUDF = "/* ------------------------------------------------ */" & vbNewLine & _
               "/* HR Pro Absence Bradford Factor user defined function.     */" & vbNewLine & _
               "/* Automatically generated by the System manager.   */" & vbNewLine & _
               "/* ------------------------------------------------ */" & vbNewLine & _
               "CREATE FUNCTION dbo.udf_ASRFn_BradfordFactor (" & vbNewLine & _
               "    @pdtStartDate datetime,      /* Start of the report period   */" & vbNewLine & _
               "    @pdtEndDate datetime,        /* End of the report period     */" & vbNewLine & _
               "    @psAbsenceTypes varchar(255), /* Abs Type we are looking for  */" & vbNewLine & _
               "    @piPersonnelID integer       /* Employees ID */" & vbNewLine & _
               ")" & vbNewLine & _
               "RETURNS float" & vbNewLine & _
               "AS" & vbNewLine & _
               "BEGIN" & vbNewLine & _
               "    DECLARE @pdToday datetime" & vbNewLine & _
               "    DECLARE @pdblResult float" & vbNewLine & vbNewLine

    strGenericSQL = strGenericSQL & vbNewLine & _
              "    /* Variables to hold the adjusted absence details (if start/end outside reporting period */" & vbNewLine & _
              "    DECLARE @dtTempStartDate datetime" & vbNewLine & _
              "    DECLARE @dtTempEndDate datetime" & vbNewLine & _
              "    DECLARE @sTempStartPeriod varchar(2)" & vbNewLine & _
              "    DECLARE @sTempEndPeriod varchar(2)" & vbNewLine & _
              "    DECLARE @bContinuous bit" & vbNewLine & vbNewLine & _
              "    /* Date counter to loop thru from StartDate to EndDate */" & vbNewLine & _
              "    DECLARE @dtCurrentDate datetime" & vbNewLine & vbNewLine & _
              "    /* The current wp/region being used in the calculation */" & vbNewLine & _
              "    DECLARE @sWorkPattern varchar(255)" & vbNewLine & _
              "    DECLARE @sPersonnelRegion varchar(255)" & vbNewLine & vbNewLine & _
              "    DECLARE @sNextWorkPattern varchar(255)" & vbNewLine & _
              "    DECLARE @sNextPersonnelRegion varchar(255)" & vbNewLine & vbNewLine & _
              "    /* ID of the persons region...used to work out which dates from the BHol Instance table apply to the employee */" & vbNewLine & _
              "    DECLARE @iBHolRegionID integer" & vbNewLine & vbNewLine

    strGenericSQL = strGenericSQL & _
              "    /* Working Pattern Stuff */" & vbNewLine & _
              "    DECLARE @fWorkAM bit" & vbNewLine & _
              "    DECLARE @fWorkPM bit" & vbNewLine & _
              "    DECLARE @iDayOfWeek integer" & vbNewLine & vbNewLine

    strGenericSQL = strGenericSQL & _
              "    /* Bradford Factor Stuff */" & vbNewLine & _
              "    DECLARE @pdblBradford float" & vbNewLine & _
              "    DECLARE @bIncludeInstance bit" & vbNewLine & _
              "    DECLARE @iInstances int" & vbNewLine & vbNewLine

    If glngSQLVersion > 7 Then
      strGenericSQL = strGenericSQL & _
               "    DECLARE @AbsenceTypes table(Type nvarchar(50) COLLATE SQL_Latin1_General_CP1_CI_AS)" & vbNewLine & vbNewLine
    End If

    strGenericSQL = strGenericSQL & _
              "    DECLARE @iCount integer" & vbNewLine & vbNewLine & _
              "    /* Date variables used when working out the next change date for historic wp/regions - if applicable */" & vbNewLine & _
              "    DECLARE @dtTempDate datetime" & vbNewLine & _
              "    DECLARE @dtNextChange_Region datetime" & vbNewLine & _
              "    DECLARE @dtNextChange_WP datetime" & vbNewLine & vbNewLine

    strGenericSQL = strGenericSQL & _
              "    /* Initialise the result to be 0 */" & vbNewLine & _
              "    SET @iInstances = 0" & vbNewLine & _
              "    SET @pdblResult = 0" & vbNewLine & vbNewLine

    ' Convert entered absence types into temporary array (sql table variable)
    If fValidConfiguration Then
      
      If glngSQLVersion = 7 Then
      
        ' SQL7 will only allow one absence type to be searched upon.
        strGenericSQL = strGenericSQL & _
                  "    /* Now we need to get a recordset of all the absence records that satisfy the date and type criteria */" & vbNewLine & _
                  "    DECLARE Absence_Cursor CURSOR LOCAL FAST_FORWARD FOR" & vbNewLine & _
                  "        SELECT " & sAbsColStartDateName & ", " & sAbsColEndDateName & ", " & sAbsColStartPeriodName & ", " & sAbsColEndPeriodName & "," & sContinuousColumnName & vbNewLine & _
                  "        FROM " & sAbsTableName & vbNewLine & _
                  "        WHERE (" & sAbsTableName & ".ID_" & lngPersonnelTableID & " = @piPersonnelID)" & vbNewLine & _
                  "        AND (" & sAbsColTypeName & " = @psAbsenceTypes)" & vbNewLine & _
                  "        AND ((" & sAbsColStartDateName & " <= @pdtstartdate AND (" & sAbsColEndDateName & " >= @pdtstartdate) OR " & sAbsColEndDateName & " IS NULL)" & vbNewLine & _
                  "        OR (" & sAbsColStartDateName & " >= @pdtstartdate AND " & sAbsColStartDateName & " <= @pdtenddate))" & vbNewLine & _
                  "        ORDER BY " & sAbsColStartDateName & vbNewLine & vbNewLine & _
                  "    OPEN Absence_Cursor" & vbNewLine
      
      Else
      
        strGenericSQL = strGenericSQL & _
                  "    /* Convert the entered absence types into table variable for processing */" & vbNewLine & _
                  "    DECLARE @SplitID nvarchar(10)" & vbNewLine & _
                  "    DECLARE @Pos int" & vbNewLine & _
                  "    SET @psAbsenceTypes = LTRIM(RTRIM(@psAbsenceTypes))+ ','" & vbNewLine & _
                  "    SET @Pos = CHARINDEX(',', @psAbsenceTypes, 1)" & vbNewLine & _
                  "    IF REPLACE(@psAbsenceTypes, ',', '') <> ''" & vbNewLine & _
                  "    BEGIN" & vbNewLine & _
                  "      WHILE @Pos > 0" & vbNewLine & _
                  "      BEGIN" & vbNewLine & _
                  "        SET @SplitID = LTRIM(RTRIM(LEFT(@psAbsenceTypes, @Pos - 1)))" & vbNewLine & _
                  "        IF @SplitID <> '' INSERT INTO @AbsenceTypes (Type) VALUES (@SplitID)" & vbNewLine & _
                  "        SET @psAbsenceTypes = RIGHT(@psAbsenceTypes, LEN(@psAbsenceTypes) - @Pos)" & vbNewLine & _
                  "        SET @Pos = CHARINDEX(',', @psAbsenceTypes, 1)" & vbNewLine & _
                  "      END" & vbNewLine & _
                  "    END" & vbNewLine & vbNewLine
        
        ' Declare the cursor guff now....
        strGenericSQL = strGenericSQL & _
                  "    /* Now we need to get a recordset of all the absence records that satisfy the date and type criteria */" & vbNewLine & _
                  "    DECLARE Absence_Cursor CURSOR LOCAL FAST_FORWARD FOR" & vbNewLine & _
                  "        SELECT " & sAbsColStartDateName & ", " & sAbsColEndDateName & ", " & sAbsColStartPeriodName & ", " & sAbsColEndPeriodName & "," & sContinuousColumnName & vbNewLine & _
                  "        FROM " & sAbsTableName & vbNewLine & _
                  "        WHERE (" & sAbsTableName & ".ID_" & lngPersonnelTableID & " = @piPersonnelID)" & vbNewLine & _
                  "        AND (" & sAbsColTypeName & " IN (SELECT Type From @AbsenceTypes))" & vbNewLine & _
                  "        AND ((" & sAbsColStartDateName & " <= @pdtstartdate AND (" & sAbsColEndDateName & " >= @pdtstartdate) OR " & sAbsColEndDateName & " IS NULL)" & vbNewLine & _
                  "        OR (" & sAbsColStartDateName & " >= @pdtstartdate AND " & sAbsColStartDateName & " <= @pdtenddate))" & vbNewLine & _
                  "        ORDER BY " & sAbsColStartDateName & vbNewLine & vbNewLine & _
                  "    OPEN Absence_Cursor" & vbNewLine
                 
      End If
    
      strGenericSQL = strGenericSQL & _
                "    /* Read the first record */" & vbNewLine & _
                "    FETCH NEXT FROM Absence_Cursor INTO @dtTempStartDate, @dtTempEndDate, @sTempStartPeriod, @sTempEndPeriod, @bContinuous" & vbNewLine & _
                "    WHILE (@@fetch_status=0)" & vbNewLine & _
                "    BEGIN" & vbNewLine

      strGenericSQL = strGenericSQL & _
                "        /* If absence does not have an end date, change it to be today*/" & vbNewLine & _
                "        IF @dtTempEndDate IS NULL SET @dtTempEndDate = @pdToday" & vbNewLine & _
                "        SET @bIncludeInstance = 0" & vbNewLine & vbNewLine

      strGenericSQL = strGenericSQL & _
                "        /* If absence start is before report start, change it to be the report start*/" & vbNewLine & _
                "        IF @dtTempStartDate < @pdtStartDate" & vbNewLine & _
                "        BEGIN" & vbNewLine & _
                "            SET @dtTempStartDate = @pdtStartDate" & vbNewLine & _
                "            SET @sTempStartPeriod = 'AM'" & vbNewLine & _
                "        END" & vbNewLine & vbNewLine

      strGenericSQL = strGenericSQL & _
                "        /* If absence end is after report end, change it to be the report end*/" & vbNewLine & _
                "        IF @dtTempEndDate > @pdtEndDate" & vbNewLine & _
                "        BEGIN" & vbNewLine & _
                "            SET @dtTempEndDate = @pdtEndDate" & vbNewLine & _
                "            SET @sTempEndPeriod = 'PM'" & vbNewLine & _
                "        END" & vbNewLine & vbNewLine

      strGenericSQL = strGenericSQL & _
                "        /* Make sure the date variables are dates */" & vbNewLine & _
                "        SET @dtTempStartDate = convert(datetime, convert(varchar(20), @dtTempStartDate, 101))" & vbNewLine & _
                "        SET @dtTempEndDate = convert(datetime, convert(varchar(20), @dtTempEndDate, 101))" & vbNewLine & vbNewLine

      strGenericSQL = strGenericSQL & _
                "        /* Set temp date to the absence start date */" & vbNewLine & _
                "        SET @dtCurrentDate = @dtTempStartDate" & vbNewLine & vbNewLine

      If (fHistoricRegion = False) And (fHistoricWP = False) Then
        ' If we are using static wp and static region, do it the simple way.
        strGenericSQL = strGenericSQL & _
                "        /* Get The Employees Working Pattern */" & vbNewLine & _
                "        SELECT @sWorkPattern = " & sStaticWPColumnName & " FROM " & sPersonnelTableName & " WHERE ID = @piPersonnelID" & vbNewLine & vbNewLine
        
        If fBHolSetupOK Then
          ' If we are including bank holidays, get the region information.
          strGenericSQL = strGenericSQL & _
                "        /* Get The Employees Region */" & vbNewLine & _
                "        SELECT @sPersonnelRegion = " & sStaticRegionColumnName & " FROM " & sPersonnelTableName & " WHERE ID = @piPersonnelID" & vbNewLine & vbNewLine & _
                "        /* Get the Region ID for the persons Region */" & vbNewLine & _
                "        SELECT @iBHolRegionID = ID FROM " & sBHolRegionTableName & " WHERE " & sBHolRegionColumnName & " = @sPersonnelRegion" & vbNewLine & vbNewLine
        End If
        
        strGenericSQL = strGenericSQL & _
                "        /* Loop through absence, only counting dates btwn the rpt dates */" & vbNewLine & _
                "        WHILE @dtCurrentDate <= @dtTempEndDate" & vbNewLine & _
                "        BEGIN" & vbNewLine & _
                "            /* Check if the current date is a work day. */" & vbNewLine & _
                "            SET @fWorkAM = 0" & vbNewLine & _
                "            SET @fWorkPM = 0" & vbNewLine & _
                "            SET @iDayOfWeek = DATEPART(weekday, @dtCurrentDate)" & vbNewLine & vbNewLine
  
        For iLoop = 1 To 7
          strGenericSQL = strGenericSQL & _
            "      IF @iDayOfWeek = " & CStr(iLoop) & vbNewLine & _
            "      BEGIN" & vbNewLine & _
            "        IF LEN(SUBSTRING(@sWorkPattern, " & CStr((iLoop * 2) - 1) & ", 1)) > 0" & vbNewLine & _
            "        BEGIN" & vbNewLine & _
            "          SET @fWorkAM = 1" & vbNewLine & _
            "        END" & vbNewLine & _
            "        IF LEN(SUBSTRING(@sWorkPattern, " & CStr(iLoop * 2) & ", 1)) > 0" & vbNewLine & _
            "        BEGIN" & vbNewLine & _
            "          SET @fWorkPM = 1" & vbNewLine & _
            "        END" & vbNewLine & _
            "      END" & vbNewLine & vbNewLine
        Next iLoop
        
        strGenericSQL = strGenericSQL & _
                "            /* If its a working day */" & vbNewLine & _
                "            IF (@fWorkAM = 1) OR (@fWorkPM = 1)" & vbNewLine & _
                "            BEGIN" & vbNewLine & vbNewLine
  
        If fBHolSetupOK Then
          ' If we are including bank holidays, check for Bhols.
          strGenericSQL = strGenericSQL & _
                "                /* Check that the current date is not a company holiday. */" & vbNewLine & _
                "                SELECT @iCount = COUNT(" & sBHolDateColumnName & ") FROM " & sBHolTableName & vbNewLine & _
                "                WHERE " & sBHolDateColumnName & " = @dtCurrentDate" & vbNewLine & _
                "                AND " & sBHolTableName & ".ID_" & lngBHolRegionTableID & " = @iBHolRegionID" & vbNewLine & vbNewLine
    
          strGenericSQL = strGenericSQL & _
                "                IF @iCount = 0" & vbNewLine & _
                "                BEGIN" & vbNewLine & _
                "                    IF @dtCurrentDate = @dtTempStartDate" & vbNewLine & _
                "                    BEGIN" & vbNewLine & _
                "                        IF ((@fWorkAM = 1) AND (UPPER(@sTempStartPeriod) = 'AM')) SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                "                        IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                "                    END" & vbNewLine & _
                "                    ELSE" & vbNewLine & _
                "                    BEGIN" & vbNewLine & _
                "                        IF @dtCurrentDate = @dtTempEndDate" & vbNewLine & _
                "                        BEGIN" & vbNewLine & _
                "                            IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                "                            IF ((@fWorkPM = 1) AND (UPPER(@sTempEndPeriod) = 'PM'))  SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                "                        END" & vbNewLine & _
                "                        ELSE" & vbNewLine & _
                "                        BEGIN" & vbNewLine & _
                "                            IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                "                            IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                "                        END" & vbNewLine & _
                "                    END" & vbNewLine & _
                "                END" & vbNewLine & vbNewLine
        Else
          ' We arent using BHols, so just add to the result without checking the bhol table.
          strGenericSQL = strGenericSQL & _
                "                IF @dtCurrentDate = @dtTempStartDate" & vbNewLine & _
                "                BEGIN" & vbNewLine & _
                "                    IF ((@fWorkAM = 1) AND (UPPER(@sTempStartPeriod) = 'AM')) SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                "                    IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                "                END" & vbNewLine & _
                "                ELSE" & vbNewLine & _
                "                BEGIN" & vbNewLine & _
                "                    IF @dtCurrentDate = @dtTempEndDate" & vbNewLine & _
                "                    BEGIN" & vbNewLine & _
                "                        IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                "                        IF ((@fWorkPM = 1) AND (UPPER(@sTempEndPeriod) = 'PM'))  SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                "                    END" & vbNewLine & _
                "                    ELSE" & vbNewLine & _
                "                    BEGIN" & vbNewLine & _
                "                        IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                "                        IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                "                    END" & vbNewLine & _
                "                END" & vbNewLine & vbNewLine
        End If

        strGenericSQL = strGenericSQL & _
                "            END" & vbNewLine & vbNewLine & _
                "            -- Include in Bradford Factor?" & vbNewLine & _
                "            IF (@fWorkPM = 1) OR (@fWorkAM = 1) SET @bIncludeInstance = 1" & vbNewLine & vbNewLine & _
                "            -- Move onto the next date." & vbNewLine & _
                "            SET @dtCurrentDate = @dtCurrentDate + 1" & vbNewLine & _
                "        END" & vbNewLine & vbNewLine
      Else
        ' Historic region and/or WP.
        If fHistoricRegion Then
          strGenericSQL = strGenericSQL & _
              "                /* Get The Employees Region For @dCurrentDate */" & vbNewLine & _
              "                SELECT TOP 1 @sNextPersonnelRegion = " & sHistoricRegionColumnName & vbNewLine & _
              "                FROM " & sHistoricRegionTableName & vbNewLine & _
              "                WHERE " & sHistoricRegionDateColumnName & " <= @dtCurrentDate" & vbNewLine & _
              "                AND ID_" & lngPersonnelTableID & " = @piPersonnelID" & vbNewLine & _
              "                ORDER BY " & sHistoricRegionDateColumnName & " DESC" & vbNewLine & vbNewLine & _
              "                SET  @dtnextchange_region = null" & vbNewLine & vbNewLine
        End If
        
        If fHistoricWP Then
          ' We are using a historic wp so ensure we are getting the right wp for @dCurrentDate.
          strGenericSQL = strGenericSQL & _
              "                SELECT TOP 1 @sNextWorkPattern = " & sHistoricWPColumnName & vbNewLine & _
              "                FROM " & sHistoricWPTableName & vbNewLine & _
              "                WHERE " & sHistoricWPDateColumnName & " <= @dtCurrentDate" & vbNewLine & _
              "                AND ID_" & lngPersonnelTableID & " = @piPersonnelID" & vbNewLine & _
              "                ORDER BY " & sHistoricWPDateColumnName & " DESC" & vbNewLine & vbNewLine & _
              "                SET  @dtnextchange_WP = null" & vbNewLine & vbNewLine
        End If
        
        strGenericSQL = strGenericSQL & _
                "        /* Loop through absence, only counting dates btwn the rpt dates */" & vbNewLine & _
                "        WHILE @dtCurrentDate <= @dtTempEndDate" & vbNewLine & _
                "        BEGIN" & vbNewLine
  
        If fHistoricRegion Then
          ' We are using a historic region, so ensure we have the right region for the @dtCurrentDate.
          strGenericSQL = strGenericSQL & _
                "            /* Only bother checking we have the right region if we dont know the nxt chg date or the current date is equal to nxt chg date */" & vbNewLine & _
                "            IF (@dtnextchange_region IS NULL) OR ((@dtCurrentDate >= @dtNextChange_Region) And (@dtCurrentDate <> '12/31/9999'))" & vbNewLine & _
                "            BEGIN" & vbNewLine & vbNewLine
    
          If fBHolSetupOK Then
            strGenericSQL = strGenericSQL & _
                "                /* Get The Employees Region For @dCurrentDate */" & vbNewLine & _
                "                SET @sPersonnelRegion = @sNextPersonnelRegion" & vbNewLine & vbNewLine
    
            strGenericSQL = strGenericSQL & _
                "                /* Get the Region ID for the persons Region */" & vbNewLine & _
                "                SELECT @iBHolRegionID = ID" & vbNewLine & _
                "                FROM " & sBHolRegionTableName & vbNewLine & _
                "                WHERE " & sBHolRegionColumnName & " = @sPersonnelRegion" & vbNewLine & vbNewLine
          End If
          
          strGenericSQL = strGenericSQL & _
                "                /* Get the date of next change for the Region */" & vbNewLine & _
                "                SET @dtTempDate = null" & vbNewLine & _
                "                SET @sNextPersonnelRegion = null" & vbNewLine & _
                "                SELECT TOP 1 @dtTempDate = " & sHistoricRegionDateColumnName & "," & vbNewLine & _
                "                  @sNextPersonnelRegion = " & sHistoricRegionColumnName & vbNewLine & _
                "                FROM " & sHistoricRegionTableName & vbNewLine & _
                "                WHERE " & sHistoricRegionDateColumnName & " > @dtCurrentDate" & vbNewLine & _
                "                AND ID_" & lngPersonnelTableID & " = @piPersonnelID" & vbNewLine & _
                "                ORDER BY " & sHistoricRegionDateColumnName & " ASC" & vbNewLine & vbNewLine

          strGenericSQL = strGenericSQL & _
                "                IF @dtTempDate IS NULL" & vbNewLine & _
                "                BEGIN" & vbNewLine & _
                "                    SET @dtNextChange_Region = '12/31/9999'" & vbNewLine & _
                "                END" & vbNewLine & _
                "                ELSE" & vbNewLine & _
                "                BEGIN" & vbNewLine & _
                "                    SET @dtNextChange_Region = @dtTempDate" & vbNewLine & _
                "                END" & vbNewLine & _
                "            END" & vbNewLine & vbNewLine
        Else
          ' We are using a static region, so get it.
          If fBHolSetupOK Then
            strGenericSQL = strGenericSQL & _
                "            SELECT @sPersonnelRegion = " & sStaticRegionColumnName & " FROM " & sPersonnelTableName & " WHERE ID = @piPersonnelID" & vbNewLine & vbNewLine & _
                "            SELECT @iBHolRegionID = ID FROM " & sBHolRegionTableName & " WHERE " & sBHolRegionColumnName & " = @sPersonnelRegion " & vbNewLine & vbNewLine
          End If
        End If

        If fHistoricWP Then
          ' We are using a historic wp so ensure we are getting the right wp for @dCurrentDate.
          strGenericSQL = strGenericSQL & _
                "            IF (@dtnextchange_WP IS NULL) OR ((@dtCurrentDate >= @dtNextChange_WP) And (@dtCurrentDate <> '12/31/9999'))" & vbNewLine & _
                "            BEGIN" & vbNewLine & _
                "                /* Get The Employees WP For @dCurrentDate */" & vbNewLine & _
                "                SET @sWorkPattern = @sNextWorkPattern" & vbNewLine & vbNewLine

          strGenericSQL = strGenericSQL & _
                "                /* Get The next change date for WP */" & vbNewLine & _
                "                SET @dtTempDate = null" & vbNewLine & _
                "                SET @sNextWorkPattern = null" & vbNewLine & _
                "                SELECT TOP 1 @dtTempDate = " & sHistoricWPDateColumnName & "," & vbNewLine & _
                "                  @sNextWorkPattern = " & sHistoricWPColumnName & vbNewLine & _
                "                FROM " & sHistoricWPTableName & vbNewLine & _
                "                WHERE " & sHistoricWPDateColumnName & " > @dtCurrentDate" & vbNewLine & _
                "                AND ID_" & lngPersonnelTableID & " = @piPersonnelID" & vbNewLine & _
                "                ORDER BY " & sHistoricWPDateColumnName & " ASC" & vbNewLine & vbNewLine
    
          strGenericSQL = strGenericSQL & _
                "                IF @dtTempDate IS NULL" & vbNewLine & _
                "                BEGIN" & vbNewLine & _
                "                    SET @dtNextChange_WP = '12/31/9999'" & vbNewLine & _
                "                END" & vbNewLine & _
                "                ELSE" & vbNewLine & _
                "                BEGIN" & vbNewLine & _
                "                    SET @dtNextChange_WP = @dtTempDate" & vbNewLine & _
                "                END" & vbNewLine & _
                "            END" & vbNewLine & vbNewLine
        Else
          strGenericSQL = strGenericSQL & _
                "            /* We are using a static wp, so get it */" & vbNewLine & _
                "            SELECT @sWorkPattern = " & sStaticWPColumnName & " FROM " & sPersonnelTableName & " WHERE ID = @piPersonnelID" & vbNewLine & vbNewLine
        End If
        
        strGenericSQL = strGenericSQL & _
                "            /* Check if the current date is a work day. */" & vbNewLine & _
                "            SET @fWorkAM = 0" & vbNewLine & _
                "            SET @fWorkPM = 0" & vbNewLine & _
                "            SET @iDayOfWeek = DATEPART(weekday, @dtCurrentDate)" & vbNewLine & vbNewLine
  
        For iLoop = 1 To 7
          strGenericSQL = strGenericSQL & _
            "      IF @iDayOfWeek = " & CStr(iLoop) & vbNewLine & _
            "      BEGIN" & vbNewLine & _
            "        IF LEN(SUBSTRING(@sWorkPattern, " & CStr((iLoop * 2) - 1) & ", 1)) > 0 SET @fWorkAM = 1" & vbNewLine & _
            "        IF LEN(SUBSTRING(@sWorkPattern, " & CStr(iLoop * 2) & ", 1)) > 0 SET @fWorkPM = 1" & vbNewLine & _
            "      END" & vbNewLine & vbNewLine
        Next iLoop

        strGenericSQL = strGenericSQL & _
                "            IF (@fWorkAM = 1) OR (@fWorkPM = 1)" & vbNewLine & _
                "            BEGIN" & vbNewLine & vbNewLine
  
        If fBHolSetupOK Then
          strGenericSQL = strGenericSQL & _
                "                /* Check that the current date is not a company holiday. */" & vbNewLine & _
                "                SELECT @iCount = COUNT(" & sBHolDateColumnName & ")" & vbNewLine & _
                "                FROM " & sBHolTableName & vbNewLine & _
                "                WHERE convert(varchar(20), " & sBHolDateColumnName & ", 101) = @dtCurrentDate" & vbNewLine & _
                "                AND " & sBHolTableName & ".ID_" & lngBHolRegionTableID & " = @iBHolRegionID" & vbNewLine & vbNewLine

          strGenericSQL = strGenericSQL & _
                "                IF @iCount = 0" & vbNewLine & _
                "                BEGIN" & vbNewLine & _
                "                    IF (@fWorkAM = 1)" & vbNewLine & _
                "                    BEGIN" & vbNewLine & _
                "                        IF ((@dtCurrentDate <> @dtTempStartDate) OR (UPPER(@sTempStartPeriod) = 'AM'))" & vbNewLine & _
                "                        BEGIN" & vbNewLine & _
                "                            SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                "                            SET @bIncludeInstance = 1" & vbNewLine & _
                "                        END" & vbNewLine & _
                "                    END" & vbNewLine & _
                "                    " & vbNewLine & _
                "                    IF (@fWorkPM = 1)" & vbNewLine & _
                "                    BEGIN" & vbNewLine & _
                "                        IF ((@dtCurrentDate <> @dtTempEndDate) OR (UPPER(@sTempEndPeriod) = 'PM'))" & vbNewLine & _
                "                        BEGIN" & vbNewLine & _
                "                            SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                "                            SET @bIncludeInstance = 1" & vbNewLine & _
                "                        END" & vbNewLine & _
                "                    END" & vbNewLine & _
                "                END" & vbNewLine & vbNewLine
        Else
          ' We arent using Bholidays, so just add to the result.
          strGenericSQL = strGenericSQL & _
                "                IF (@fWorkAM = 1)" & vbNewLine & _
                "                BEGIN" & vbNewLine & _
                "                    IF ((@dtCurrentDate <> @dtTempStartDate) OR (UPPER(@sTempStartPeriod) = 'AM'))" & vbNewLine & _
                "                    BEGIN" & vbNewLine & _
                "                        SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                "                        SET @bIncludeInstance = 1" & vbNewLine & _
                "                    END" & vbNewLine & _
                "                END" & vbNewLine & _
                "                " & vbNewLine & _
                "                IF (@fWorkPM = 1)" & vbNewLine & _
                "                BEGIN" & vbNewLine & _
                "                    IF ((@dtCurrentDate <> @dtTempEndDate) OR (UPPER(@sTempEndPeriod) = 'PM'))" & vbNewLine & _
                "                    BEGIN" & vbNewLine & _
                "                        SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
                "                        SET @bIncludeInstance = 1" & vbNewLine & _
                "                    END" & vbNewLine & _
                "                END" & vbNewLine & vbNewLine
        End If
                
        strGenericSQL = strGenericSQL & _
                "            END" & vbNewLine & vbNewLine & _
                "            -- Move onto the next date." & vbNewLine & _
                "            SET @dtCurrentDate = @dtCurrentDate + 1" & vbNewLine & _
                "        END" & vbNewLine & vbNewLine
      End If
      
      strGenericSQL = strGenericSQL & _
                "        IF ((@bIncludeInstance = 1) AND (@bContinuous = 0)) SET @iInstances = @iInstances + 1" & vbNewLine & vbNewLine
      
      strGenericSQL = strGenericSQL & _
                "        FETCH NEXT FROM Absence_Cursor INTO @dtTempStartDate, @dtTempEndDate, @sTempStartPeriod, @sTempEndPeriod, @bContinuous" & vbNewLine & _
                "    END" & vbNewLine & vbNewLine & _
                "    CLOSE Absence_Cursor" & vbNewLine & _
                "    DEALLOCATE Absence_Cursor" & vbNewLine & vbNewLine
                
    End If
       
    ' Build sp and udfs
    sAbsenceBetweenProc = sAbsenceBetweenProc & strGenericSQL & "END"
    sAbsenceBetweenUDF = sAbsenceBetweenUDF & strGenericSQL & vbNewLine & "    RETURN @pdblResult" & vbNewLine & "END"
    sBradfordUDF = sBradfordUDF & strGenericSQL & vbNewLine & _
        "    IF (@pdblResult > 0 AND @iInstances = 0) SET @iInstances = 1;" & vbNewLine & vbNewLine & _
        "    RETURN ((@iInstances * @iInstances) * @pdblResult)" & vbNewLine & "END"
    
    gADOCon.Execute sAbsenceBetweenProc, , adExecuteNoRecords
    
    If gbEnableUDFFunctions Then
      gADOCon.Execute sAbsenceBetweenUDF, , adExecuteNoRecords
      gADOCon.Execute sBradfordUDF, , adExecuteNoRecords
    End If
    
  End If
  
TidyUpAndExit:
  On Error GoTo ErrorTrap
  CreateAbsenceBetween2DatesStoredProcedure = fCreatedOK
  Exit Function

ErrorTrap:
  fCreatedOK = False
  OutputError "Error creating Absence Between 2 Dates stored procedure (Absence)"
  Resume TidyUpAndExit

End Function


Private Function CreateSSPStoredProcedure() As Boolean
  ' Create the Statutory Sick Pay stored procedure.
  On Error GoTo ErrorTrap
  
  Dim fCreatedOK As Boolean
  Dim sSQL As String
  Dim sProcSQL As String
  Dim fSSPRunningTableExists As Boolean
  Dim rsInfo As New ADODB.Recordset
  
  'MH20020308 Fault 3334
  'Already dropped this SP earlier in the code...
  
  ' Drop any existing SSP stored procedure.
  'fCreatedOK = DropSSPStoredProcedure
  fCreatedOK = True
  
  'If fCreatedOK Then
    ' Construct the stored procedure creation string (if required).
    sProcSQL = "/* ------------------------------------------------ */" & vbNewLine & _
      "/* HR Pro Absence module stored procedure.          */" & vbNewLine & _
      "/* Automatically generated by the System manager.   */" & vbNewLine & _
      "/* ------------------------------------------------ */" & vbNewLine & _
      "CREATE PROCEDURE dbo." & gsSSP_PROCEDURENAME & " (" & vbNewLine & _
      "    @piAbsenceRecordID  integer" & vbNewLine & _
      ")" & vbNewLine & _
      "AS" & vbNewLine & _
      "BEGIN"
      
    sProcSQL = sProcSQL & vbNewLine & _
      "    /* Refresh the SSP fields in the Absence records for the Personnel record that is the parent of the given Absence record ID. */" & vbNewLine & vbNewLine & _
      "    /* Personnel record variables. */" & vbNewLine & _
      "    SET NOCOUNT ON;" & vbNewLine & vbNewLine & _
      "    DECLARE @iPersonnelRecordID integer," & vbNewLine & _
      "        @iWorkingDaysPerWeek integer," & vbNewLine & _
      "        @sWorkingPattern varchar(14)," & vbNewLine & _
      "        @dtDateOfBirth datetime," & vbNewLine & _
      "        @dtRetirementDate datetime," & vbNewLine & _
      "        @dtSixteenthBirthday datetime" & vbNewLine
  
    sProcSQL = sProcSQL & vbNewLine & _
      "    /* Absence record variables. */" & vbNewLine & _
      "    DECLARE @cursAbsenceRecords cursor," & vbNewLine & _
      "        @cursFollowingAbsenceRecords cursor," & vbNewLine & _
      "        @iAbsenceRecordID integer," & vbNewLine & _
      "        @dtStartDate datetime," & vbNewLine & _
      "        @dtEndDate datetime," & vbNewLine & _
      "        @sStartSession varchar(100)," & vbNewLine & _
      "        @sEndSession varchar(100)," & vbNewLine & _
      "        @dtWholeStartDate datetime," & vbNewLine & _
      "        @dtWholeEndDate datetime," & vbNewLine & _
      "        @dtFollowingStartDate datetime," & vbNewLine & _
      "        @dtFollowingEndDate datetime," & vbNewLine & _
      "        @sFollowingStartSession varchar(100)," & vbNewLine & _
      "        @sFollowingEndSession varchar(100)," & vbNewLine & _
      "        @dtFollowingWholeStartDate datetime," & vbNewLine & _
      "        @dtFollowingWholeEndDate datetime," & vbNewLine & _
      "        @fOriginalSSPApplies bit," & vbNewLine & _
      "        @dblOriginalQualifyingDays float," & vbNewLine & _
      "        @dblOriginalWaitingDays float," & vbNewLine & _
      "        @dblOriginalPaidDays float," & vbNewLine & _
      "        @iNewSSPApplies bit" & vbNewLine
  
  
    sProcSQL = sProcSQL & vbNewLine & _
      "    /* General procedure handling variables. */" & vbNewLine & _
      "    DECLARE @fOK bit," & vbNewLine & _
      "        @iLoop integer," & vbNewLine & _
      "        @iIndex integer," & vbNewLine & _
      "        @sCommandString nvarchar(MAX)," & vbNewLine & _
      "        @sParamDefinition nvarchar(500)," & vbNewLine & _
      "        @dblWaitEntitlement float," & vbNewLine & _
      "        @dblAbsenceEntitlement float," & vbNewLine & _
      "        @dblQualifyingDays float," & vbNewLine & _
      "        @dblWaitingDays float," & vbNewLine & _
      "        @dblPaidDays float," & vbNewLine & _
      "        @fSSPApplies bit," & vbNewLine & _
      "        @dtTempDate datetime," & vbNewLine & _
      "        @fAddOK bit," & vbNewLine & _
      "        @dblAddAmount float," & vbNewLine & _
      "        @fContinue bit," & vbNewLine & _
      "        @iConsecutiveRecords integer," & vbNewLine & _
      "        @dtConsecutiveStartDate datetime," & vbNewLine & _
      "        @dtConsecutiveEndDate datetime," & vbNewLine & _
      "        @dtConsecutiveWholeStartDate datetime," & vbNewLine & _
      "        @dtConsecutiveWholeEndDate datetime," & vbNewLine & _
      "        @sConsecutiveStartSession varchar(100)," & vbNewLine & _
      "        @sConsecutiveEndSession varchar(100)," & vbNewLine
      
    sProcSQL = sProcSQL & _
      "        @dtLastWholeEndDate datetime," & vbNewLine & _
      "        @dtFirstLinkedWholeStartDate datetime," & vbNewLine & _
      "        @iYearDifference integer," & vbNewLine & _
      "        @fSSPRunning bit," & vbNewLine & _
      "        @iAbsenceRecordCount integer," & vbNewLine & _
      "        @iCurrAbsRec integer" & vbNewLine & vbNewLine & _
      "    SET @fOK = 1" & vbNewLine
  
    sProcSQL = sProcSQL & vbNewLine & _
      "    /* Get the ID  of the associated record in the Personnel table. */" & vbNewLine & _
      "    SELECT @iPersonnelRecordID = id_" & Trim(Str(mvar_lngPersonnelTableID)) & vbNewLine & _
      "    FROM " & mvar_sAbsenceTableName & vbNewLine & _
      "    WHERE id = @piAbsenceRecordID" & vbNewLine & vbNewLine & _
      "    IF (@iPersonnelRecordID IS null) OR (@iPersonnelRecordID <= 0) SET @fOK = 0" & vbNewLine
  
    ' 22/03/2002 JPD Check to avoid recurrent running of the SSP stored procedure.
    rsInfo.Open "SELECT COUNT(*) FROM sysobjects WHERE name = 'ASRSysSSPRunning' AND type = 'U'", gADOCon, adOpenForwardOnly, adLockReadOnly
    If Not rsInfo.BOF And Not rsInfo.EOF Then
      fSSPRunningTableExists = (rsInfo.Fields(0).value > 0)
    Else
      fSSPRunningTableExists = False
    End If
    rsInfo.Close
    
    If fSSPRunningTableExists Then
      sProcSQL = sProcSQL & vbNewLine & _
        "    IF @fOK = 1" & vbNewLine & _
        "    BEGIN" & vbNewLine & _
        "        /* Check to avoid recurrent running of the SSP stored procedure. */" & vbNewLine & _
        "        SELECT @fSSPRunning = sspRunning" & vbNewLine & _
        "        FROM ASRSysSSPRunning" & vbNewLine & _
        "        WHERE personnelRecordID = @iPersonnelRecordID" & vbNewLine & vbNewLine & _
        "        IF @fSSPRunning IS null INSERT INTO ASRSysSSPRunning (personnelRecordID, sspRunning) VALUES(@iPersonnelRecordID, 1)" & vbNewLine & _
        "        IF @fSSPRunning = 0 UPDATE ASRSysSSPRunning SET sspRunning = 1 WHERE personnelRecordID = @iPersonnelRecordID" & vbNewLine & _
        "        IF @fSSPRunning = 1 RETURN" & vbNewLine & _
        "    END" & vbNewLine & vbNewLine
    End If
    
    If Len(mvar_sPersonnel_DateOfBirthColumnName) > 0 Then
      sProcSQL = sProcSQL & vbNewLine & _
        "    IF @fOK = 1" & vbNewLine & _
        "    BEGIN" & vbNewLine & _
        "        /* Get the retirement date, and the date of the person's sixteenth birthday. */" & vbNewLine & _
        "        SELECT @dtDateOfBirth = convert(datetime, convert(varchar(20), " & mvar_sPersonnel_DateOfBirthColumnName & ", 101))" & vbNewLine & _
        "        FROM " & mvar_sPersonnelTableName & vbNewLine & _
        "        WHERE id = @iPersonnelRecordID" & vbNewLine & vbNewLine & _
        "        IF (NOT @dtDateOfBirth IS null) SET @dtRetirementDate = dateadd(yy, 65, @dtDateOfBirth)" & vbNewLine & _
        "        IF (NOT @dtDateOfBirth IS null) SET @dtSixteenthBirthday = dateadd(yy, 16, @dtDateOfBirth)" & vbNewLine & _
        "    END" & vbNewLine
    End If
    
    sProcSQL = sProcSQL & vbNewLine & _
      "    IF @fOK = 1" & vbNewLine & _
      "    BEGIN" & vbNewLine & _
      "        /* Get the number of working days per week. */" & vbNewLine
  
    Select Case mvar_iAbsenceWorkingDaysType
      Case 0 ' The Working Days are a straight numeric value.
        sProcSQL = sProcSQL & _
          "        SET @iWorkingDaysPerWeek = " & Trim(Str(mvar_iAbsenceWorkingDaysNumericValue)) & vbNewLine & _
          "        SET @sWorkingPattern = ''" & vbNewLine
          
      Case 1 ' The Working Days are a straight working pattern value.
        sProcSQL = sProcSQL & _
          "        SET @iWorkingDaysPerWeek = " & CStr(Len(Replace(mvar_sAbsenceWorkingDaysPatternValue, " ", "")) / 2) & vbNewLine & _
          "        SET @sWorkingPattern = '" & mvar_sAbsenceWorkingDaysPatternValue & "'" & vbNewLine
          
      Case 2  ' The Working Days is a numeric field reference.
        If (mvar_lngAbsenceWorkingDaysTableID = mvar_lngPersonnelTableID) Then
          sProcSQL = sProcSQL & vbNewLine & _
            "        SET @sWorkingPattern = ''" & vbNewLine & _
            "        SELECT @iWorkingDaysPerWeek = " & mvar_sPersonnelTableName & "." & mvar_sAbsenceWorkingDaysColumnName & vbNewLine & _
            "        FROM " & mvar_sPersonnelTableName & vbNewLine & _
            "        WHERE " & mvar_sPersonnelTableName & ".id = @iPersonnelRecordID" & vbNewLine
        Else
          sProcSQL = sProcSQL & _
            "        SET @iWorkingDaysPerWeek = 0" & vbNewLine & _
            "        SET @sWorkingPattern = ''" & vbNewLine
        End If
             
      Case 3  ' The Working Days is a pattern field reference.
        If (mvar_lngAbsenceWorkingDaysTableID = mvar_lngPersonnelTableID) Then
          sProcSQL = sProcSQL & vbNewLine & _
            "        SET @iWorkingDaysPerWeek = 0" & vbNewLine & _
            "        SELECT @sWorkingPattern = " & mvar_sPersonnelTableName & "." & mvar_sAbsenceWorkingDaysColumnName & vbNewLine & _
            "        FROM " & mvar_sPersonnelTableName & vbNewLine & _
            "        WHERE " & mvar_sPersonnelTableName & ".id = @iPersonnelRecordID" & vbNewLine & _
            "        IF @sWorkingPattern IS null SET @sWorkingPattern = ''" & vbNewLine & _
            "        /* Calculate the number of qualifying days per week. */" & vbNewLine & _
            "        IF len(@sWorkingPattern) > 0" & vbNewLine & _
            "        BEGIN" & vbNewLine & _
            "            SET @iLoop = 1" & vbNewLine & vbNewLine & _
            "            WHILE (len(@sWorkingPattern) >= (@iLoop * 2)) AND (@iLoop <=14)" & vbNewLine & _
            "            BEGIN" & vbNewLine & _
            "                IF (substring(@sWorkingPattern, @iLoop, 1) <> ' ') AND (substring(@sWorkingPattern, @iLoop + 1, 1) <> ' ')" & vbNewLine & _
            "                BEGIN" & vbNewLine & _
            "                    SET @iWorkingDaysPerWeek = @iWorkingDaysPerWeek + 1" & vbNewLine & _
            "                END" & vbNewLine & vbNewLine & _
            "                SET @iLoop = @iLoop + 2" & vbNewLine & _
            "            END" & vbNewLine & _
            "        END" & vbNewLine
        Else
          sProcSQL = sProcSQL & _
            "        SET @iWorkingDaysPerWeek = 0" & vbNewLine & _
            "        SET @sWorkingPattern = ''" & vbNewLine
        End If
    End Select
  
    sProcSQL = sProcSQL & _
      "    END" & vbNewLine
    
    ' JDM-2010-06-14 - JIRA 1005 - Yup, another fix only a few days later... :-)
    ' NPG20100609 Fault HRPRO-725
    sProcSQL = sProcSQL & vbNewLine & _
      "    IF @fOK = 1" & vbNewLine & _
      "    BEGIN" & vbNewLine & _
      "        SET @iConsecutiveRecords = 0" & vbNewLine & _
      "        SET @dtLastWholeEndDate = null" & vbNewLine & vbNewLine & _
      "        /* Get count of absence records */" & vbNewLine & vbNewLine & _
      "        SELECT @iAbsenceRecordCount = COUNT(id) FROM " & mvar_sAbsenceTableName & vbNewLine & _
      "            WHERE id_" & Trim(Str(mvar_lngPersonnelTableID)) & " = @iPersonnelRecordID;" & vbNewLine & vbNewLine & _
      "        /* Create a cursor of the absence records for the current person. */" & vbNewLine & _
      "        SET @cursAbsenceRecords = CURSOR LOCAL FAST_FORWARD FOR" & vbNewLine & _
      "            SELECT " & mvar_sAbsenceTableName & ".id," & vbNewLine & _
      "                convert(datetime, convert(varchar(20), " & mvar_sAbsenceTableName & "." & mvar_sAbsence_StartDateColumnName & ", 101))," & vbNewLine & _
      "                convert(datetime, convert(varchar(20), " & mvar_sAbsenceTableName & "." & mvar_sAbsence_EndDateColumnName & ", 101))," & vbNewLine & _
      "                upper(left(" & mvar_sAbsenceTableName & "." & mvar_sAbsence_StartSessionColumnName & ", 2)), " & vbNewLine & _
      "                upper(left(" & mvar_sAbsenceTableName & "." & mvar_sAbsence_EndSessionColumnName & ", 2)), " & vbNewLine & _
      "                " & mvar_sAbsenceTableName & "." & mvar_sAbsence_SSPAppliesColumnName & ", " & vbNewLine & _
      "                " & mvar_sAbsenceTableName & "." & mvar_sAbsence_QualifyingDaysColumnName & ", " & vbNewLine & _
      "                " & mvar_sAbsenceTableName & "." & mvar_sAbsence_WaitingDaysColumnName & ", " & vbNewLine & _
      "                " & mvar_sAbsenceTableName & "." & mvar_sAbsence_PaidDaysColumnName & ", " & vbNewLine & _
      "                " & mvar_sAbsenceTypeTableName & "." & mvar_sAbsenceType_SSPAppliesColumnName
  
    If ((mvar_iAbsenceWorkingDaysType = 2) Or (mvar_iAbsenceWorkingDaysType = 3)) And _
      (mvar_lngAbsenceWorkingDaysTableID = mvar_lngAbsenceTableID) Then
      ' The Working Days is a numeric or pattern field reference.
      sProcSQL = sProcSQL & _
        ", " & vbNewLine & _
        "                " & mvar_sAbsenceTableName & "." & mvar_sAbsenceWorkingDaysColumnName
    End If
  
    sProcSQL = sProcSQL & vbNewLine & _
      "            FROM " & mvar_sAbsenceTableName & vbNewLine & _
      "            INNER JOIN " & mvar_sAbsenceTypeTableName & " ON " & mvar_sAbsenceTableName & "." & mvar_sAbsence_TypeColumnName & " = " & mvar_sAbsenceTypeTableName & "." & mvar_sAbsenceType_TypeColumnName & vbNewLine & _
      "            WHERE " & mvar_sAbsenceTableName & ".id_" & Trim(Str(mvar_lngPersonnelTableID)) & " = @iPersonnelRecordID " & vbNewLine & _
      "            ORDER BY " & mvar_sAbsenceTableName & "." & mvar_sAbsence_StartDateColumnName & ", " & mvar_sAbsenceTableName & ".id" & vbNewLine & _
      "        OPEN @cursAbsenceRecords" & vbNewLine
  
    sProcSQL = sProcSQL & vbNewLine & _
      "        /* Loop through the absence records, calculating SSP for each record." & vbNewLine & _
      "        NB. We check if any periods of absence are consecutive before checking for SSP application. */" & vbNewLine & _
      "        FETCH NEXT FROM @cursAbsenceRecords INTO @iAbsenceRecordID, @dtStartDate, @dtEndDate, @sStartSession, @sEndSession, @fOriginalSSPApplies, @dblOriginalQualifyingDays, @dblOriginalWaitingDays, @dblOriginalPaidDays, @iNewSSPApplies"
      
    If (mvar_lngAbsenceWorkingDaysTableID = mvar_lngAbsenceTableID) Then
      Select Case mvar_iAbsenceWorkingDaysType
        Case 2  ' The Working Days is a numeric field reference.
          sProcSQL = sProcSQL & ", @iWorkingDaysPerWeek"
        Case 3  ' The Working Days is a pattern field reference.
          sProcSQL = sProcSQL & ", @sWorkingPattern"
      End Select
    End If
    
    ' NPG20100609 Fault HRPRO-725
    sProcSQL = sProcSQL & vbNewLine & _
      "        WHILE (@@fetch_status = 0)" & vbNewLine & _
      "        BEGIN" & vbNewLine & _
      "            /* Increment record count */" & vbNewLine & _
      "            SET @iCurrAbsRec = @iCurrAbsRec + 1" & vbNewLine & vbNewLine & _
      "            /* Ignore incomplete absence records. */" & vbNewLine & _
      "            IF (NOT @dtStartDate IS null) AND (NOT @dtEndDate IS null)" & vbNewLine & _
      "            BEGIN" & vbNewLine & _
      "                /* Ignore absence after retirement. */" & vbNewLine & _
      "                IF NOT @dtRetirementDate IS null" & vbNewLine & _
      "                BEGIN" & vbNewLine & _
      "                    IF (@dtRetirementDate < @dtEndDate)" & vbNewLine & _
      "                    BEGIN" & vbNewLine & _
      "                        SET @dtEndDate = @dtRetirementDate" & vbNewLine & _
      "                        SET @sEndSession = 'PM'" & vbNewLine & _
      "                    END" & vbNewLine & _
      "                END" & vbNewLine
        
    sProcSQL = sProcSQL & vbNewLine & _
      "                /* Ignore absence before the sixteenth birthday. */" & vbNewLine & _
      "                IF NOT @dtSixteenthBirthday IS null" & vbNewLine & _
      "                BEGIN" & vbNewLine & _
      "                    IF (@dtSixteenthBirthday > @dtStartDate)" & vbNewLine & _
      "                    BEGIN" & vbNewLine & _
      "                        SET @dtStartDate = @dtSixteenthBirthday" & vbNewLine & _
      "                        SET @sStartSession = 'AM'" & vbNewLine & _
      "                    END" & vbNewLine & _
      "                END" & vbNewLine
  
    sProcSQL = sProcSQL & vbNewLine & _
      "                /* Get the start and end dates (whole days only) of the current absence record. */" & vbNewLine & _
      "                SET @dtWholeStartDate = @dtStartDate" & vbNewLine & _
      "                SET @dtWholeEndDate = @dtEndDate" & vbNewLine & _
      "                IF @sStartSession = 'PM' SET @dtWholeStartDate = @dtWholeStartDate + 1" & vbNewLine & _
      "                IF @sEndSession = 'AM' SET @dtWholeEndDate = @dtWholeEndDate - 1" & vbNewLine
  
    sProcSQL = sProcSQL & vbNewLine & _
      "                IF @iConsecutiveRecords = 0" & vbNewLine & _
      "                BEGIN" & vbNewLine & _
      "                    SET @dtConsecutiveStartDate = @dtStartDate" & vbNewLine & _
      "                    SET @dtConsecutiveEndDate = @dtEndDate" & vbNewLine & _
      "                    SET @sConsecutiveStartSession = @sStartSession" & vbNewLine & _
      "                    SET @sConsecutiveEndSession = @sEndSession" & vbNewLine & _
      "                    SET @dtConsecutiveWholeStartDate = @dtWholeStartDate" & vbNewLine & _
      "                    SET @dtConsecutiveWholeEndDate = @dtWholeEndDate" & vbNewLine
  
    sProcSQL = sProcSQL & vbNewLine & _
      "                    /* Create a cursor of the absence records for the current person that follow the current absence record. */" & vbNewLine & _
      "                    SET @cursFollowingAbsenceRecords = CURSOR LOCAL FAST_FORWARD FOR " & vbNewLine & _
      "                        SELECT convert(datetime, convert(varchar(20), " & mvar_sAbsenceTableName & "." & mvar_sAbsence_StartDateColumnName & ", 101)), " & vbNewLine & _
      "                            convert(datetime, convert(varchar(20), " & mvar_sAbsenceTableName & "." & mvar_sAbsence_EndDateColumnName & ", 101)), " & vbNewLine & _
      "                            upper(left(" & mvar_sAbsenceTableName & "." & mvar_sAbsence_StartSessionColumnName & ", 2)), " & vbNewLine & _
      "                            upper(left(" & mvar_sAbsenceTableName & "." & mvar_sAbsence_EndSessionColumnName & ", 2)) " & vbNewLine & _
      "                        FROM " & mvar_sAbsenceTableName & vbNewLine & _
      "                        INNER JOIN " & mvar_sAbsenceTypeTableName & " ON " & mvar_sAbsenceTableName & "." & mvar_sAbsence_TypeColumnName & " = " & mvar_sAbsenceTypeTableName & "." & mvar_sAbsenceType_TypeColumnName & vbNewLine & _
      "                        WHERE " & mvar_sAbsenceTableName & ".id_" & Trim(Str(mvar_lngPersonnelTableID)) & " = @iPersonnelRecordID" & vbNewLine & _
      "                            AND " & mvar_sAbsenceTypeTableName & "." & mvar_sAbsenceType_SSPAppliesColumnName & " = 1" & vbNewLine & _
      "                            AND (NOT " & mvar_sAbsenceTableName & "." & mvar_sAbsence_StartDateColumnName & " IS null)" & vbNewLine & _
      "                            AND (NOT " & mvar_sAbsenceTableName & "." & mvar_sAbsence_EndDateColumnName & " IS null)" & vbNewLine & _
      "                            AND ((convert(varchar(20), " & mvar_sAbsenceTableName & "." & mvar_sAbsence_StartDateColumnName + ", 112) > convert(varchar(20), @dtStartDate, 112))" & vbNewLine & _
      "                            OR ((convert(varchar(20), " & mvar_sAbsenceTableName & "." & mvar_sAbsence_StartDateColumnName + ", 112) = convert(varchar(20), @dtStartDate, 112)) AND (" & mvar_sAbsenceTableName & ".id > @iAbsenceRecordID)))" & vbNewLine & _
      "                        ORDER BY " & mvar_sAbsenceTableName & "." & mvar_sAbsence_StartDateColumnName & ", " & mvar_sAbsenceTableName & ".id" & vbNewLine & _
      "                    OPEN @cursFollowingAbsenceRecords" & vbNewLine
  
    sProcSQL = sProcSQL & vbNewLine & _
      "                    SET @fContinue = 1" & vbNewLine & _
      "                    FETCH NEXT FROM @cursFollowingAbsenceRecords INTO @dtFollowingStartDate, @dtFollowingEndDate, @sFollowingStartSession, @sFollowingEndSession" & vbNewLine & _
      "                    WHILE (@@fetch_status = 0) AND (@fContinue = 1)" & vbNewLine & _
      "                    BEGIN" & vbNewLine & _
      "                        SET @fContinue = 0" & vbNewLine & vbNewLine & _
      "                        /* Get the start and end dates (whole days only) of the current absence records. */" & vbNewLine & _
      "                        SET @dtFollowingWholeStartDate = @dtFollowingStartDate" & vbNewLine & _
      "                        SET @dtFollowingWholeEndDate = @dtFollowingEndDate" & vbNewLine & _
      "                        IF @sFollowingStartSession = 'PM' SET @dtFollowingWholeStartDate = @dtFollowingWholeStartDate + 1" & vbNewLine & _
      "                        IF @sFollowingEndSession = 'AM' SET @dtFollowingWholeEndDate = @dtFollowingWholeEndDate - 1" & vbNewLine & vbNewLine & _
      "                        IF ((@dtConsecutiveEndDate = @dtFollowingStartDate) AND (@sConsecutiveEndSession = 'AM') AND (@sFollowingStartSession = 'PM'))" & vbNewLine & _
      "                            OR (@dtConsecutiveWholeEndDate + 1 >= @dtFollowingWholeStartDate)" & vbNewLine & _
      "                        BEGIN" & vbNewLine & _
      "                            SET @iConsecutiveRecords = @iConsecutiveRecords + 1" & vbNewLine & _
      "                            SET @dtConsecutiveEndDate = @dtFollowingEndDate" & vbNewLine & _
      "                            SET @sConsecutiveEndSession = @sFollowingEndSession" & vbNewLine & _
      "                            SET @dtConsecutiveWholeEndDate = @dtFollowingWholeEndDate" & vbNewLine & _
      "                            SET @fContinue = 1" & vbNewLine & _
      "                        END" & vbNewLine & vbNewLine & _
      "                        FETCH NEXT FROM @cursFollowingAbsenceRecords INTO @dtFollowingStartDate, @dtFollowingEndDate, @sFollowingStartSession, @sFollowingEndSession" & vbNewLine & _
      "                    END" & vbNewLine & vbNewLine & _
      "                    CLOSE @cursFollowingAbsenceRecords" & vbNewLine & _
      "                    DEALLOCATE @cursFollowingAbsenceRecords" & vbNewLine
  
    sProcSQL = sProcSQL & _
      "                END" & vbNewLine & _
      "                ELSE" & vbNewLine & _
      "                BEGIN" & vbNewLine & _
      "                    SET @iConsecutiveRecords = @iConsecutiveRecords - 1" & vbNewLine & _
      "                END" & vbNewLine
  
    sProcSQL = sProcSQL & vbNewLine & _
      "                /* SSP Applies if the absence period is greater than 3 days. */" & vbNewLine & _
      "                SET @fSSPApplies = 0" & vbNewLine & _
      "                IF (datediff(dd, @dtConsecutiveWholeStartDate, @dtConsecutiveWholeEndDate) + 1) > 3 AND @iNewSSPApplies = 1 SET @fSSPApplies = 1" & vbNewLine & vbNewLine & _
      "                IF @fSSPApplies = 1" & vbNewLine & _
      "                BEGIN" & vbNewLine & _
      "                    /* Check if 56 days have passed since the previous absence period. */" & vbNewLine & _
      "                    IF @dtLastWholeEndDate IS null" & vbNewLine & _
      "                    BEGIN" & vbNewLine & _
      "                        /* First absence record so use default values. */" & vbNewLine & _
      "                        SET @dblWaitEntitlement = 3" & vbNewLine & _
      "                        SET @dblAbsenceEntitlement = @iWorkingDaysPerWeek * 28" & vbNewLine & _
      "                        SET @dtFirstLinkedWholeStartDate = @dtWholeStartDate" & vbNewLine & _
      "                    END" & vbNewLine & _
      "                    ELSE" & vbNewLine & _
      "                    BEGIN" & vbNewLine & _
      "                        IF (datediff(dd, @dtLastWholeEndDate, @dtWholeStartDate) - 1) > 56" & vbNewLine & _
      "                        BEGIN" & vbNewLine & _
      "                            /* More than 56 days since the previous absence record so use default values. */" & vbNewLine & _
      "                            SET @dblWaitEntitlement = 3" & vbNewLine & _
      "                            SET @dblAbsenceEntitlement = @iWorkingDaysPerWeek * 28" & vbNewLine & _
      "                            SET @dtFirstLinkedWholeStartDate = @dtWholeStartDate" & vbNewLine & _
      "                        END" & vbNewLine & _
      "                    END" & vbNewLine
  
    If (mvar_lngAbsenceWorkingDaysTableID = mvar_lngAbsenceTableID) Then
      Select Case mvar_iAbsenceWorkingDaysType
        Case 2  ' The Working Days is a numeric field reference.
          sProcSQL = sProcSQL & vbNewLine & _
            "                    SET @sWorkingPattern = ''" & vbNewLine & _
            "                    IF @iWorkingDaysPerWeek IS null SET @iWorkingDaysPerWeek = 0" & vbNewLine
        Case 3  ' The Working Days is a pattern field reference.
          sProcSQL = sProcSQL & vbNewLine & _
            "                    SET @iWorkingDaysPerWeek = 0" & vbNewLine & _
            "                    IF @sWorkingPattern IS null SET @sWorkingPattern = ''" & vbNewLine & _
            "                    /* Calculate the number of qualifying days per week. */" & vbNewLine & _
            "                    IF len(@sWorkingPattern) > 0" & vbNewLine & _
            "                    BEGIN" & vbNewLine & _
            "                        SET @iLoop = 1" & vbNewLine & vbNewLine & _
            "                        WHILE (len(@sWorkingPattern) >= (@iLoop * 2)) AND (@iLoop <=14)" & vbNewLine & _
            "                        BEGIN" & vbNewLine & _
            "                            IF (substring(@sWorkingPattern, @iLoop, 1) <> ' ') AND (substring(@sWorkingPattern, @iLoop + 1, 1) <> ' ')" & vbNewLine & _
            "                            BEGIN" & vbNewLine & _
            "                                SET @iWorkingDaysPerWeek = @iWorkingDaysPerWeek + 1" & vbNewLine & _
            "                            END" & vbNewLine & vbNewLine & _
            "                            SET @iLoop = @iLoop + 2" & vbNewLine & _
            "                        END" & vbNewLine & _
            "                    END" & vbNewLine
      End Select
    End If
    
    sProcSQL = sProcSQL & vbNewLine & _
      "                    /* Calculate SSP qualifying, waiting and paid days." & vbNewLine & _
      "                    NB. The start and end dates should already take into account the start and end periods (AM/PM)" & vbNewLine & _
      "                    so that only whole absence days are used. */" & vbNewLine & _
      "                    SET @dblQualifyingDays = 0" & vbNewLine & vbNewLine & _
      "                    /* Loop from the start date to the end date, incrementing the number of qualifying days for each date that qualifies. */" & vbNewLine & _
      "                    SET @dtTempDate = @dtStartDate" & vbNewLine & vbNewLine & _
      "                    WHILE (@dtTempDate <= @dtEndDate)" & vbNewLine & _
      "                    BEGIN" & vbNewLine & _
      "                        SET @fAddOK = 0" & vbNewLine & _
      "                        SET @dblAddAmount = 0" & vbNewLine & vbNewLine & _
      "                        IF len(@sWorkingPattern) = 0" & vbNewLine & _
      "                        BEGIN" & vbNewLine & _
      "                            /* No working pattern passed in, so use the 'daysPerWeek' variable. */" & vbNewLine & _
      "                            IF (@iWorkingDaysPerWeek = 7) OR" & vbNewLine & _
      "                                ((datepart(dw, @dtTempDate) >= 2) AND (datepart(dw, @dtTempDate) <= 6))" & vbNewLine & _
      "                            BEGIN" & vbNewLine & _
      "                                /* The current date qualifies if 7 days per week are worked, or if the current date is a weekday. */" & vbNewLine & _
      "                                SET @fAddOK = 1" & vbNewLine & _
      "                            END" & vbNewLine & _
      "                        END" & vbNewLine
            
    sProcSQL = sProcSQL & _
      "                        ELSE" & vbNewLine & _
      "                        BEGIN" & vbNewLine & _
      "                            /* Use the working pattern. */" & vbNewLine & _
      "                            SET @iIndex = (2 * datepart(dw, @dtTempDate)) -1" & vbNewLine & _
      "                            IF len(@sWorkingPattern) >= (@iIndex +1)" & vbNewLine & _
      "                            BEGIN" & vbNewLine & _
      "                                /* The current date qualifies if its 'day of the week' is worked in the working pattern." & vbNewLine & _
      "                                NB. Both AM and PM sessions must be worked for the day to qualify. */" & vbNewLine & _
      "                                IF (substring(@sWorkingPattern, @iIndex, 1) <> ' ') AND (substring(@sWorkingPattern, @iIndex + 1, 1) <> ' ')" & vbNewLine & _
      "                                BEGIN" & vbNewLine & _
      "                                    SET @fAddOK = 1" & vbNewLine & _
      "                                END" & vbNewLine & _
      "                            END" & vbNewLine & _
      "                        END" & vbNewLine
  
    sProcSQL = sProcSQL & vbNewLine & _
      "                        IF @fAddOK = 1" & vbNewLine & _
      "                        BEGIN" & vbNewLine & _
      "                            /* If the person is older than retirement age, then the day does not qualify. */" & vbNewLine & _
      "                            IF NOT @dtRetirementDate IS null" & vbNewLine & _
      "                            BEGIN" & vbNewLine & _
      "                                IF @dtTempDate > @dtRetirementDate SET @fAddOK = 0" & vbNewLine & _
      "                            END" & vbNewLine & _
      "                        END" & vbNewLine & vbNewLine & _
      "                        IF @fAddOK = 1" & vbNewLine & _
      "                        BEGIN" & vbNewLine & _
      "                            /* If the person is less than sixteen then the day does not qualify. */" & vbNewLine & _
      "                            IF (NOT @dtSixteenthBirthday IS null)" & vbNewLine & _
      "                            BEGIN" & vbNewLine & _
      "                                IF @dtTempDate < @dtSixteenthBirthday SET @fAddOK = 0" & vbNewLine & _
      "                            END" & vbNewLine & _
      "                        END" & vbNewLine
  
    sProcSQL = sProcSQL & vbNewLine & _
      "                        IF @fAddOK = 1" & vbNewLine & _
      "                        BEGIN" & vbNewLine & _
      "                            /* Days linked after 3 years from the start of the link do not count. */" & vbNewLine & _
      "                            exec dbo.sp_ASRFn_WholeYearsBetweenTwoDates @iYearDifference OUTPUT, @dtFirstLinkedWholeStartDate, @dtTempDate" & vbNewLine & _
      "                            IF @iYearDifference >= 3  SET @fAddOK = 0" & vbNewLine & _
      "                        END" & vbNewLine
  
    sProcSQL = sProcSQL & vbNewLine & _
      "                        /* Calculate how much to add to the Qualifying Days. */" & vbNewLine & _
      "                        IF @fAddOK = 1" & vbNewLine & _
      "                        BEGIN" & vbNewLine & _
      "                            SET @dblAddAmount = 0" & vbNewLine & vbNewLine & _
      "                            IF @dtTempDate < @dtWholeStartDate" & vbNewLine & _
      "                            BEGIN" & vbNewLine & _
      "                                /* The current date is the half day before the whole dated period starts." & vbNewLine & _
      "                                A half day qualifies only if this period of absence consecutively follows another. */" & vbNewLine & _
      "                                IF (@dtConsecutiveStartDate < @dtStartDate) OR" & vbNewLine & _
      "                                    ((@dtConsecutiveStartDate = @dtStartDate) AND (@sConsecutiveStartSession <> @sStartSession)) SET @dblAddAmount = 0.5" & vbNewLine & _
      "                            END" & vbNewLine & _
      "                            ELSE" & vbNewLine & _
      "                            BEGIN" & vbNewLine & _
      "                                IF @dtTempDate > @dtWholeEndDate" & vbNewLine & _
      "                                BEGIN" & vbNewLine & _
      "                                    /* The current date is the half day after the whole dated period end." & vbNewLine & _
      "                                    A half day qualifies only if this period of absence is consecutively followed by another. */" & vbNewLine & _
      "                                    IF (@dtConsecutiveEndDate > @dtEndDate) OR" & vbNewLine & _
      "                                        ((@dtConsecutiveEndDate = @dtEndDate) AND (@sConsecutiveEndSession <> @sStartSession) AND (@iCurrAbsRec <> @iAbsenceRecordCount)) SET @dblAddAmount = 0.5" & vbNewLine & _
      "                                END" & vbNewLine
                
    sProcSQL = sProcSQL & _
      "                                ELSE" & vbNewLine & _
      "                                BEGIN" & vbNewLine & _
      "                                    /* The current date lies within the whole dated period, so a whole day qualifies. */" & vbNewLine & _
      "                                    SET @dblAddAmount = 1" & vbNewLine & _
      "                                END" & vbNewLine & _
      "                            END" & vbNewLine & _
      "                        END" & vbNewLine
  
    sProcSQL = sProcSQL & vbNewLine & _
      "                        /* Increment the number of qualifying days. */" & vbNewLine & _
      "                        SET @dblQualifyingDays = @dblQualifyingDays + @dblAddAmount" & vbNewLine & vbNewLine & _
      "                        SET @dtTempDate = @dtTempDate + 1" & vbNewLine & _
      "                    END" & vbNewLine
  
    sProcSQL = sProcSQL & vbNewLine & _
      "                    /* Take off any waiting entitlement. */" & vbNewLine & _
      "                    IF @dblWaitEntitlement > @dblQualifyingDays" & vbNewLine & _
      "                    BEGIN" & vbNewLine & _
      "                        SET @dblWaitingDays = @dblQualifyingDays" & vbNewLine & _
      "                        SET @dblWaitEntitlement = @dblWaitEntitlement - @dblQualifyingDays" & vbNewLine & _
      "                    END" & vbNewLine & _
      "                    ELSE" & vbNewLine & _
      "                    BEGIN" & vbNewLine & _
      "                        SET @dblWaitingDays = @dblWaitEntitlement" & vbNewLine & _
      "                        SET @dblWaitEntitlement = 0" & vbNewLine & _
      "                    END" & vbNewLine
  
    sProcSQL = sProcSQL & vbNewLine & _
      "                    /* Paid days is the difference providing there is enough entitlement. */" & vbNewLine & _
      "                    SET @dblPaidDays = @dblQualifyingDays - @dblWaitingDays" & vbNewLine & vbNewLine & _
      "                    IF @dblPaidDays > @dblAbsenceEntitlement" & vbNewLine & _
      "                    BEGIN" & vbNewLine & _
      "                        SET @dblPaidDays = @dblAbsenceEntitlement" & vbNewLine & _
      "                        SET @dblAbsenceEntitlement = 0" & vbNewLine & _
      "                    END" & vbNewLine & _
      "                    ELSE" & vbNewLine & _
      "                    BEGIN" & vbNewLine & _
      "                        SET @dblAbsenceEntitlement = @dblAbsenceEntitlement - @dblPaidDays" & vbNewLine & _
      "                    END" & vbNewLine & vbNewLine & _
      "                    SET @dtLastWholeEndDate = @dtWholeEndDate" & vbNewLine
  
    sProcSQL = sProcSQL & vbNewLine & _
      "                    /* Update the SSP fields in the current absence record if required. */" & vbNewLine & _
      "                    IF (@fOriginalSSPApplies IS null) OR" & vbNewLine & _
      "                        (@fOriginalSSPApplies = 0) OR" & vbNewLine & _
      "                        (@dblOriginalQualifyingDays IS null) OR" & vbNewLine & _
      "                        (@dblOriginalQualifyingDays <> @dblQualifyingDays) OR" & vbNewLine & _
      "                        (@dblOriginalWaitingDays IS null) OR" & vbNewLine & _
      "                        (@dblOriginalWaitingDays <> @dblWaitingDays) OR" & vbNewLine & _
      "                        (@dblOriginalPaidDays IS null) OR" & vbNewLine & _
      "                        (@dblOriginalPaidDays <> @dblPaidDays)" & vbNewLine & _
      "                    BEGIN" & vbNewLine & _
      "                        UPDATE " & mvar_sAbsenceTableName & vbNewLine & _
      "                        SET " & mvar_sAbsence_SSPAppliesColumnName & " = 1," & vbNewLine & _
      "                            " & mvar_sAbsence_QualifyingDaysColumnName & " = @dblQualifyingDays," & vbNewLine & _
      "                            " & mvar_sAbsence_WaitingDaysColumnName & " = @dblWaitingDays," & vbNewLine & _
      "                            " & mvar_sAbsence_PaidDaysColumnName & " = @dblPaidDays" & vbNewLine & _
      "                        WHERE id = @iAbsenceRecordID" & vbNewLine & _
      "                    END" & vbNewLine & _
      "                END" & vbNewLine
        
    sProcSQL = sProcSQL & _
      "                ELSE" & vbNewLine & _
      "                BEGIN" & vbNewLine & _
      "                    /* Update the SSP fields in the current absence record. */" & vbNewLine & _
      "                    IF (@fOriginalSSPApplies IS null) OR" & vbNewLine & _
      "                        (@fOriginalSSPApplies = 1) OR" & vbNewLine & _
      "                        (@dblOriginalQualifyingDays IS null) OR" & vbNewLine & _
      "                        (@dblOriginalQualifyingDays <> 0) OR" & vbNewLine & _
      "                        (@dblOriginalWaitingDays IS null) OR" & vbNewLine & _
      "                        (@dblOriginalWaitingDays <> 0) OR" & vbNewLine & _
      "                        (@dblOriginalPaidDays IS null) OR" & vbNewLine & _
      "                        (@dblOriginalPaidDays <> 0)" & vbNewLine & _
      "                    BEGIN" & vbNewLine & _
      "                        UPDATE " & mvar_sAbsenceTableName & vbNewLine & _
      "                        SET " & mvar_sAbsence_SSPAppliesColumnName & " = 0," & vbNewLine & _
      "                            " & mvar_sAbsence_QualifyingDaysColumnName & " = 0," & vbNewLine & _
      "                            " & mvar_sAbsence_WaitingDaysColumnName & " = 0," & vbNewLine & _
      "                            " & mvar_sAbsence_PaidDaysColumnName & " = 0" & vbNewLine & _
      "                        WHERE id = @iAbsenceRecordID" & vbNewLine & _
      "                    END" & vbNewLine & _
      "                END" & vbNewLine & _
      "            END" & vbNewLine
      
    sProcSQL = sProcSQL & _
      "            ELSE" & vbNewLine & _
      "            BEGIN" & vbNewLine & _
      "                /* Update the SSP fields in the current absence record. */" & vbNewLine & _
      "                IF (@fOriginalSSPApplies IS null) OR" & vbNewLine & _
      "                    (@fOriginalSSPApplies = 1) OR" & vbNewLine & _
      "                    (@dblOriginalQualifyingDays IS null) OR" & vbNewLine & _
      "                    (@dblOriginalQualifyingDays <> 0) OR" & vbNewLine & _
      "                    (@dblOriginalWaitingDays IS null) OR" & vbNewLine & _
      "                    (@dblOriginalWaitingDays <> 0) OR" & vbNewLine & _
      "                    (@dblOriginalPaidDays IS null) OR" & vbNewLine & _
      "                    (@dblOriginalPaidDays <> 0)" & vbNewLine & _
      "                BEGIN" & vbNewLine & _
      "                    UPDATE " & mvar_sAbsenceTableName & vbNewLine & _
      "                    SET " & mvar_sAbsence_SSPAppliesColumnName & " = 0," & vbNewLine & _
      "                        " & mvar_sAbsence_QualifyingDaysColumnName & " = 0," & vbNewLine & _
      "                        " & mvar_sAbsence_WaitingDaysColumnName & " = 0," & vbNewLine & _
      "                        " & mvar_sAbsence_PaidDaysColumnName & " = 0" & vbNewLine & _
      "                    WHERE id = @iAbsenceRecordID" & vbNewLine & _
      "                END" & vbNewLine & _
      "            END" & vbNewLine
  
    sProcSQL = sProcSQL & vbNewLine & _
      "        FETCH NEXT FROM @cursAbsenceRecords INTO @iAbsenceRecordID, @dtStartDate, @dtEndDate, @sStartSession, @sEndSession, @fOriginalSSPApplies, @dblOriginalQualifyingDays, @dblOriginalWaitingDays, @dblOriginalPaidDays, @iNewSSPApplies"
  
    If (mvar_lngAbsenceWorkingDaysTableID = mvar_lngAbsenceTableID) Then
      Select Case mvar_iAbsenceWorkingDaysType
        Case 2  ' The Working Days is a numeric field reference.
          sProcSQL = sProcSQL & ", @iWorkingDaysPerWeek"
        Case 3  ' The Working Days is a pattern field reference.
          sProcSQL = sProcSQL & ", @sWorkingPattern"
      End Select
    End If
    
    sProcSQL = sProcSQL & vbNewLine & _
      "        END" & vbNewLine & _
      "        CLOSE @cursAbsenceRecords" & vbNewLine & _
      "        DEALLOCATE @cursAbsenceRecords" & vbNewLine & _
      "    END" & vbNewLine

    ' 22/03/2002 JPD Check to avoid recurrent running of the SSP stored procedure.
    If fSSPRunningTableExists Then
      sProcSQL = sProcSQL & vbNewLine & _
        "    IF @fOK = 1" & vbNewLine & _
        "    BEGIN" & vbNewLine & _
        "        UPDATE ASRSysSSPRunning SET sspRunning = 0 WHERE personnelRecordID = @iPersonnelRecordID" & vbNewLine & _
        "    END" & vbNewLine & vbNewLine
    End If

    sProcSQL = sProcSQL & vbNewLine & _
      "END"
  
    gADOCon.Execute sProcSQL, , adExecuteNoRecords
  'End If
  
TidyUpAndExit:
  Set rsInfo = Nothing
  CreateSSPStoredProcedure = fCreatedOK
  Exit Function
  
ErrorTrap:
  fCreatedOK = False
  OutputError "Error creating SSP stored procedure (Absence)"
  Resume TidyUpAndExit

End Function

Private Function CreateWorkingDaysBetween2DatesStoredProcedure() As Boolean
  ' JPD20020515 Fault 3342
  ' Create the WorkingDaysBetween2Dates stored procedure.
  On Error GoTo ErrorTrap
  
  Dim fCreatedOK As Boolean
  Dim sSQL As String
  Dim iLoop As Integer
  
  Dim strGenericSQL As String
  Dim strProcSQL As String
  Dim strProcStart As String
  Dim strProcEnd As String
  Dim strUDFSQL As String
  Dim strUDFStart As String
  Dim strUDFEnd As String
  
  Dim iTempID As Integer
  Dim fValidConfiguration As Boolean
  
  ' Personnel Table
  Dim iPersonnelTableID As Long
  Dim sPersonnelTable  As String
  
  ' Bank Holiday Region (Primary) Table
  Dim iBHolRegionTableID As Long
  Dim sBHolRegionTableName  As String
  Dim sBHolRegionColumnName As String
  
  ' Bank Holiday Instance (Child) Table
  Dim iBHolTableID As Long
  Dim sBHolTableName  As String
  Dim sBHolDateColumnName As String
  
  ' Flag storing if the Bank Hols are setup OK and therefore if we should use them or not
  Dim fBHolSetupOK As Boolean
  
  ' Flag stating if we are using historic region setup (True) or static (False)
  Dim fHistoricRegion As Boolean
  
  ' Variables to hold the relevant region table/column names
  Dim sStaticRegionColumnName As String
  Dim sHistoricRegionTableName As String
  Dim sHistoricRegionColumnName As String
  Dim sHistoricRegionDateColumnName As String
  
  ' Flag stating if we are using historic wp setup (True) or static (False)
  Dim fHistoricWP As Boolean
  
  ' Variables to hold the relevant wp table/column names
  Dim sStaticWPColumnName As String
  Dim sHistoricWPTableName As String
  Dim sHistoricWPColumnName As String
  Dim sHistoricWPDateColumnName As String
    
  'fCreatedOK = DropWorkingDaysBetween2DatesStoredProcedure
  fCreatedOK = DropProcedure(gsWorkingDaysBetween2Dates_PROCEDURENAME)
  
  If gbEnableUDFFunctions Then
    fCreatedOK = DropWorkingDaysBetween2DatesUDF
  End If

  If fCreatedOK Then
    fValidConfiguration = True
     
    ' Get the BHol Region Table ID and Name
    iTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_ABSENCE, "Param_TableBHolRegion"
    If Not recModuleSetup.NoMatch Then
      iTempID = recModuleSetup!parametervalue
      iBHolRegionTableID = iTempID
      recTabEdit.Index = "idxTableID"
      recTabEdit.Seek "=", iTempID
      If Not recTabEdit.NoMatch Then
        sBHolRegionTableName = recTabEdit!TableName
      Else
        iBHolRegionTableID = 0
        sBHolRegionTableName = vbNullString
      End If
    Else
      iBHolRegionTableID = 0
      sBHolRegionTableName = vbNullString
    End If
    
    ' Get the BHolRegion column in the BHolRegion Table
    iTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_ABSENCE, "Param_FieldBHolRegion"
    If Not recModuleSetup.NoMatch Then
      iTempID = recModuleSetup!parametervalue
      recColEdit.Index = "idxColumnID"
      recColEdit.Seek "=", iTempID
      If Not recColEdit.NoMatch Then
        sBHolRegionColumnName = recColEdit!ColumnName
      Else
        sBHolRegionColumnName = vbNullString
      End If
    Else
      sBHolRegionColumnName = vbNullString
    End If
    
    ' Get the BHol Table ID (instances of BHols)
    iTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_ABSENCE, "Param_TableBHol"
    If Not recModuleSetup.NoMatch Then
      iTempID = recModuleSetup!parametervalue
      iBHolTableID = iTempID
      recTabEdit.Index = "idxTableID"
      recTabEdit.Seek "=", iTempID
      If Not recTabEdit.NoMatch Then
        sBHolTableName = recTabEdit!TableName
      Else
        iBHolTableID = 0
        sBHolTableName = vbNullString
      End If
    Else
      iBHolTableID = 0
      sBHolTableName = vbNullString
    End If
    
    ' Get the BHolDate Column Name
    iTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_ABSENCE, "Param_FieldBHolDate"
    If Not recModuleSetup.NoMatch Then
      iTempID = recModuleSetup!parametervalue
      recColEdit.Index = "idxColumnID"
      recColEdit.Seek "=", iTempID
      If Not recColEdit.NoMatch Then
        sBHolDateColumnName = recColEdit!ColumnName
      Else
        sBHolDateColumnName = vbNullString
      End If
    Else
      sBHolDateColumnName = vbNullString
    End If
    
    ' Set flag to state whether BHols have been setup correctly or Not
    If (iBHolRegionTableID > 0) And _
      (sBHolRegionTableName <> vbNullString) And _
      (sBHolRegionColumnName <> vbNullString) And _
      (iBHolTableID <> 0) And _
      (sBHolTableName <> vbNullString) And _
      (sBHolDateColumnName <> vbNullString) Then
      fBHolSetupOK = True
    Else
      fBHolSetupOK = False
    End If
    
    ' Get the Personnel Table ID and Name
    iTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, "Param_TablePersonnel"
    If Not recModuleSetup.NoMatch Then
      iTempID = recModuleSetup!parametervalue
      iPersonnelTableID = iTempID
      recTabEdit.Index = "idxTableID"
      recTabEdit.Seek "=", iTempID
      If Not recTabEdit.NoMatch Then
        sPersonnelTable = recTabEdit!TableName
      Else
        iPersonnelTableID = 0
        sPersonnelTable = vbNullString
      End If
    Else
      iPersonnelTableID = 0
      sPersonnelTable = vbNullString
    End If
      
    ' Get the Static Region Column Name
    iTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, "Param_FieldsRegion"
    If Not recModuleSetup.NoMatch Then
      iTempID = recModuleSetup!parametervalue
      recColEdit.Index = "idxColumnID"
      recColEdit.Seek "=", iTempID
      If Not recColEdit.NoMatch Then
        sStaticRegionColumnName = recColEdit!ColumnName
      Else
        sStaticRegionColumnName = vbNullString
      End If
    Else
      sStaticRegionColumnName = vbNullString
    End If
    
    ' Get the Historic Region Table Name
    iTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, "Param_FieldsHRegionTable"
    If Not recModuleSetup.NoMatch Then
      iTempID = recModuleSetup!parametervalue
      recTabEdit.Index = "idxTableID"
      recTabEdit.Seek "=", iTempID
      If Not recTabEdit.NoMatch Then
        sHistoricRegionTableName = recTabEdit!TableName
      Else
        sHistoricRegionTableName = vbNullString
      End If
    Else
      sHistoricRegionTableName = vbNullString
    End If
    
    ' Get the Historic Region Column Name
    iTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, "Param_FieldsHRegion"
    If Not recModuleSetup.NoMatch Then
      iTempID = recModuleSetup!parametervalue
      recColEdit.Index = "idxColumnID"
      recColEdit.Seek "=", iTempID
      If Not recColEdit.NoMatch Then
        sHistoricRegionColumnName = recColEdit!ColumnName
      Else
        sHistoricRegionColumnName = vbNullString
      End If
    Else
      sHistoricRegionColumnName = vbNullString
    End If
    
    ' Get the Historic Region Date Column Name
    iTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, "Param_FieldsHRegionDate"
    If Not recModuleSetup.NoMatch Then
      iTempID = recModuleSetup!parametervalue
      recColEdit.Index = "idxColumnID"
      recColEdit.Seek "=", iTempID
      If Not recColEdit.NoMatch Then
        sHistoricRegionDateColumnName = recColEdit!ColumnName
      Else
        sHistoricRegionDateColumnName = vbNullString
      End If
    Else
      sHistoricRegionDateColumnName = vbNullString
    End If
    
    ' Set flag to indicate what type of regions we are to use
    If sStaticRegionColumnName = vbNullString Then
      If (sHistoricRegionTableName = vbNullString) Or _
        (sHistoricRegionColumnName = vbNullString) Or _
        (sHistoricRegionDateColumnName = vbNullString) Then
        fValidConfiguration = False
      Else
        fHistoricRegion = True
      End If
    Else
      fHistoricRegion = False
    End If
    
    ' Get the Static WP Column Name
    iTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, "Param_FieldsWorkingPattern"
    If Not recModuleSetup.NoMatch Then
      iTempID = recModuleSetup!parametervalue
      recColEdit.Index = "idxColumnID"
      recColEdit.Seek "=", iTempID
      If Not recColEdit.NoMatch Then
        sStaticWPColumnName = recColEdit!ColumnName
      Else
        sStaticWPColumnName = vbNullString
      End If
    Else
      sStaticWPColumnName = vbNullString
    End If
  
    ' Get the Historic WP Table Name
    iTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, "Param_FieldsHWorkingPatternTable"
    If Not recModuleSetup.NoMatch Then
      iTempID = recModuleSetup!parametervalue
      recTabEdit.Index = "idxTableID"
      recTabEdit.Seek "=", iTempID
      If Not recTabEdit.NoMatch Then
        sHistoricWPTableName = recTabEdit!TableName
      Else
        sHistoricWPTableName = vbNullString
      End If
    Else
      sHistoricWPTableName = vbNullString
    End If
  
    ' Get the Historic WP Column Name
    iTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, "Param_FieldsHWorkingPattern"
    If Not recModuleSetup.NoMatch Then
      iTempID = recModuleSetup!parametervalue
      recColEdit.Index = "idxColumnID"
      recColEdit.Seek "=", iTempID
      If Not recColEdit.NoMatch Then
        sHistoricWPColumnName = recColEdit!ColumnName
      Else
        sHistoricWPColumnName = vbNullString
      End If
    Else
      sHistoricWPColumnName = vbNullString
    End If
    
    ' Get the Historic WP Date Column Name
    iTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, "Param_FieldsHWorkingPatternDate"
    If Not recModuleSetup.NoMatch Then
      iTempID = recModuleSetup!parametervalue
      recColEdit.Index = "idxColumnID"
      recColEdit.Seek "=", iTempID
      If Not recColEdit.NoMatch Then
        sHistoricWPDateColumnName = recColEdit!ColumnName
      Else
        sHistoricWPDateColumnName = vbNullString
      End If
    Else
      sHistoricWPDateColumnName = vbNullString
    End If
  
    ' Set flag to indicate what type of wp we are to use
    If sStaticWPColumnName = vbNullString Then
      If (sHistoricWPTableName = vbNullString) Or _
        (sHistoricWPColumnName = vbNullString) Or _
        (sHistoricWPDateColumnName = vbNullString) Then
        fValidConfiguration = False
      Else
        fHistoricWP = True
      End If
    Else
      fHistoricWP = False
    End If
    
    ' Construct the stored procedure creation string (if required).
    strProcStart = "/* ------------------------------------------------ */" & vbNewLine & _
      "/* HR Pro Absence module stored procedure.          */" & vbNewLine & _
      "/* Automatically generated by the System manager.   */" & vbNewLine & _
      "/* ------------------------------------------------ */" & vbNewLine & _
      "CREATE PROCEDURE dbo." & gsWorkingDaysBetween2Dates_PROCEDURENAME & " (" & vbNewLine & _
      "    @pdblResult float OUTPUT," & vbNewLine & _
      "    @pdtStartDate datetime," & vbNewLine & _
      "    @pdtEndDate datetime," & vbNewLine & _
      "    @iPersonnelID integer" & vbNewLine & _
      ")" & vbNewLine & _
      "AS" & vbNewLine & _
      "BEGIN" & vbNewLine
      
    strUDFStart = "/* ------------------------------------------------ */" & vbNewLine & _
      "/* HR Pro Absence module user defined function.     */" & vbNewLine & _
      "/* Automatically generated by the System manager.   */" & vbNewLine & _
      "/* ------------------------------------------------ */" & vbNewLine & _
      "CREATE FUNCTION dbo.udf_ASRFn_WorkingDaysBetweenTwoDates (" & vbNewLine & _
      "    @pdtStartDate datetime," & vbNewLine & _
      "    @pdtEndDate datetime," & vbNewLine & _
      "    @iPersonnelID integer" & vbNewLine & _
      ")" & vbNewLine & _
      "RETURNS float" & vbNewLine & _
      "AS" & vbNewLine & _
      "BEGIN" & vbNewLine & _
      "    DECLARE @pdblResult float" & vbNewLine
      
    strGenericSQL = _
      "    DECLARE @iCount int" & vbNewLine & vbNewLine & _
      "    /* Date counter to loop thru from StartDate to EndDate */" & vbNewLine & _
      "    DECLARE @dtCurrentDate datetime" & vbNewLine & vbNewLine & _
      "    /* ID of the persons region...used to work out which dates from the BHol Instance table apply to the employee */" & vbNewLine & _
      "    DECLARE @iBHolRegionID int" & vbNewLine & vbNewLine & _
      "    /* The current wp/region being used in the calculation */" & vbNewLine & _
      "    DECLARE @psWorkPattern varchar(255)" & vbNewLine & _
      "    DECLARE @psPersonnelRegion varchar(255)" & vbNewLine & _
      "    DECLARE @psNextWorkPattern varchar(255)" & vbNewLine & _
      "    DECLARE @psNextPersonnelRegion varchar(255)" & vbNewLine & vbNewLine
  
    strGenericSQL = strGenericSQL & _
      "    /* Working Pattern Stuff */" & vbNewLine & _
      "    DECLARE @fWorkAM bit" & vbNewLine & _
      "    DECLARE @fWorkPM bit" & vbNewLine & _
      "    DECLARE @iDayOfWeek int" & vbNewLine & vbNewLine
      
    strGenericSQL = strGenericSQL & _
      "    /* Date variables used when working out the next change date for historic WP/Regions - If applicable */" & vbNewLine & _
      "    DECLARE @dtTempDate datetime" & vbNewLine & _
      "    DECLARE @dtNextChange_Region datetime" & vbNewLine & _
      "    DECLARE @dtNextChange_WP datetime" & vbNewLine & vbNewLine
  
    strGenericSQL = strGenericSQL & _
      "    /* Initialise the result to be 0 */" & vbNewLine & _
      "    SET @pdblResult = 0" & vbNewLine & vbNewLine & _
      "    /* If Calculate the Absence Duration if all parameters are valid. */" & vbNewLine & _
      "    IF (@pdtStartDate IS NULL) OR (@pdtEndDate IS NULL) OR (@iPersonnelID IS NULL)" & vbNewLine & _
      "    BEGIN" & vbNewLine & _
      "        RETURN 0" & vbNewLine & _
      "    END" & vbNewLine & vbNewLine
  
    If fValidConfiguration Then
      strGenericSQL = strGenericSQL & _
        "    /* Make sure the variables are nice sql dates */" & vbNewLine & _
        "    SET @pdtStartDate = convert(datetime, convert(varchar(20), @pdtStartDate, 101))" & vbNewLine & _
        "    SET @pdtEndDate = convert(datetime, convert(varchar(20), @pdtEndDate, 101))" & vbNewLine & vbNewLine & _
        "    SET @dtCurrentDate = @pdtStartDate" & vbNewLine & vbNewLine
    
      ' If we are using static wp and static region, do it the simple way
      If (fHistoricRegion = False) And (fHistoricWP = False) Then
        strGenericSQL = strGenericSQL & _
          "    /* Get The Employees Working Pattern */" & vbNewLine & _
          "    SELECT @psWorkPattern = " & sStaticWPColumnName & vbNewLine & _
          "        FROM " & sPersonnelTable & vbNewLine & _
          "        WHERE ID = @iPersonnelID" & vbNewLine & vbNewLine
  
        ' If we are including bank holidays, get the region information
        If fBHolSetupOK Then
          strGenericSQL = strGenericSQL & _
            "    /* Get The Employees Region */" & vbNewLine & _
            "    SELECT @psPersonnelRegion = " & sStaticRegionColumnName & vbNewLine & _
            "        FROM " & sPersonnelTable & vbNewLine & _
            "        WHERE ID = @iPersonnelID" & vbNewLine & vbNewLine & _
            "    /* Get the Region ID for the persons Region */" & vbNewLine & _
            "    SELECT @iBHolRegionID = ID" & vbNewLine & _
            "        FROM " & sBHolRegionTableName & vbNewLine & _
            "        WHERE " & sBHolRegionColumnName & " = @psPersonnelRegion" & vbNewLine & vbNewLine
        End If
  
        strGenericSQL = strGenericSQL & _
          "    /* Loop through absence, only counting dates btwn the rpt dates */" & vbNewLine & _
          "    WHILE @dtCurrentDate <= @pdtEndDate" & vbNewLine & _
          "    BEGIN" & vbNewLine & vbNewLine & _
          "        /* Check if the current date is a work day. */" & vbNewLine & _
          "        SET @fWorkAM = 0" & vbNewLine & _
          "        SET @fWorkPM = 0" & vbNewLine & _
          "        SET @iDayOfWeek = DATEPART(weekday, @dtCurrentDate)" & vbNewLine & vbNewLine
  
        For iLoop = 1 To 7
          strGenericSQL = strGenericSQL & _
            "      IF @iDayOfWeek = " & CStr(iLoop) & vbNewLine & _
            "      BEGIN" & vbNewLine & _
            "        IF LEN(SUBSTRING(@psWorkPattern, " & CStr((iLoop * 2) - 1) & ", 1)) > 0" & vbNewLine & _
            "        BEGIN" & vbNewLine & _
            "          SET @fWorkAM = 1" & vbNewLine & _
            "        END" & vbNewLine & _
            "        IF LEN(SUBSTRING(@psWorkPattern, " & CStr(iLoop * 2) & ", 1)) > 0" & vbNewLine & _
            "        BEGIN" & vbNewLine & _
            "          SET @fWorkPM = 1" & vbNewLine & _
            "        END" & vbNewLine & _
            "      END" & vbNewLine
        Next iLoop
        
        strGenericSQL = strGenericSQL & _
          "        /* If its a working day */" & vbNewLine & _
          "        IF (@fWorkAM = 1) OR (@fWorkPM = 1)" & vbNewLine & _
          "        BEGIN" & vbNewLine
          
        ' If we are including bank holidays, check for Bhols
        If fBHolSetupOK Then
          strGenericSQL = strGenericSQL & _
            "            /* Check that the current date is not a company holiday. */" & vbNewLine & _
            "            SELECT @iCount = COUNT(" & sBHolDateColumnName & ") FROM " & sBHolTableName & vbNewLine & _
            "                WHERE convert(varchar(20), " & sBHolDateColumnName & ", 101) = convert(varchar(20), @dtCurrentDate, 101)" & vbNewLine & _
            "                AND " & sBHolTableName & ".ID_" & iBHolRegionTableID & " = convert(varchar(20), @iBHolRegionID)" & vbNewLine & vbNewLine & _
            "            IF @iCount = 0" & vbNewLine & _
            "            BEGIN" & vbNewLine & _
            "                IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
            "                IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
            "            END" & vbNewLine & vbNewLine
        Else
          strGenericSQL = strGenericSQL & _
            "            /* We arent using BHols, so just add to the result without checking the bhol table */" & vbNewLine & _
            "            IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
            "            IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine & vbNewLine
        End If
  
        strGenericSQL = strGenericSQL & _
          "        END" & vbNewLine & vbNewLine & _
          "        /* Move onto the next date. */" & vbNewLine & _
          "        SET @dtCurrentDate = @dtCurrentDate + 1" & vbNewLine & _
          "    END" & vbNewLine
      Else
        If fBHolSetupOK And fHistoricRegion Then
          strGenericSQL = strGenericSQL & _
            "            /* Get The Employees Region For @dCurrentDate */" & vbNewLine & _
            "            SELECT TOP 1 @psNextPersonnelRegion = " & sHistoricRegionColumnName & vbNewLine & _
            "                FROM " & sHistoricRegionTableName & vbNewLine & _
            "                WHERE " & sHistoricRegionDateColumnName & " <= @dtCurrentDate" & vbNewLine & _
            "                AND ID_" & iPersonnelTableID & " = @iPersonnelID" & vbNewLine & _
            "                ORDER BY " & sHistoricRegionDateColumnName & " DESC" & vbNewLine & vbNewLine
        End If
        
        If fHistoricWP Then
          strGenericSQL = strGenericSQL & _
            "            /* Get The Employees WP For @dCurrentDate */" & vbNewLine & _
            "            SELECT TOP 1 @psNextWorkPattern = " & sHistoricWPColumnName & vbNewLine & _
            "                FROM " & sHistoricWPTableName & vbNewLine & _
            "                WHERE " & sHistoricWPDateColumnName & " <= @dtCurrentDate" & vbNewLine & _
            "                AND ID_" & iPersonnelTableID & " = @iPersonnelID" & vbNewLine & _
            "                ORDER BY " & sHistoricWPDateColumnName & " DESC" & vbNewLine & vbNewLine
        End If
        
        strGenericSQL = strGenericSQL & _
          "    /* Either historic wp or region, so do this...*/" & vbNewLine & _
          "    /* Loop through absence, only counting dates btwn the rpt dates */" & vbNewLine & _
          "    WHILE @dtCurrentDate <= @pdtEndDate" & vbNewLine & _
          "    BEGIN" & vbNewLine
          
        If fHistoricRegion Then
          ' We are using a historic region, so ensure we have the right region for the @dtCurrentDate
          strGenericSQL = strGenericSQL & _
            "        /* Only bother checking we have the right region if we dont know the nxt chg date or the current date is equal to nxt chg date */" & vbNewLine & _
            "        IF (@dtnextchange_region IS NULL) OR ((@dtCurrentDate >= @dtNextChange_Region) And (@dtCurrentDate <> '12/31/9999'))" & vbNewLine & _
            "        BEGIN" & vbNewLine

          If fBHolSetupOK Then
            strGenericSQL = strGenericSQL & _
              "            /* Get The Employees Region For @dCurrentDate */" & vbNewLine & _
              "            SET @psPersonnelRegion = @psNextPersonnelRegion" & vbNewLine & vbNewLine
    
            strGenericSQL = strGenericSQL & _
              "            /* Get the Region ID for the persons Region */" & vbNewLine & _
              "            SELECT @iBHolRegionID = ID" & vbNewLine & _
              "                FROM " & sBHolRegionTableName & vbNewLine & _
              "                WHERE " & sBHolRegionColumnName & " = @psPersonnelRegion" & vbNewLine & vbNewLine
          End If
          
          strGenericSQL = strGenericSQL & _
            "            /* Get the date of next change for the Region */" & vbNewLine & _
            "            SET @dtTempDate = null" & vbNewLine & _
            "            SET @psNextPersonnelRegion = null" & vbNewLine & _
            "            SELECT TOP 1 @dtTempDate = " & sHistoricRegionDateColumnName & vbNewLine
            
          If fBHolSetupOK Then
            strGenericSQL = strGenericSQL & _
              "            ,@psNextPersonnelRegion = " & sHistoricRegionColumnName & vbNewLine
          End If
          
          strGenericSQL = strGenericSQL & _
            "                FROM " & sHistoricRegionTableName & vbNewLine & _
            "                WHERE " & sHistoricRegionDateColumnName & " > @dtCurrentDate" & vbNewLine & _
            "                AND ID_" & iPersonnelTableID & " = @iPersonnelID" & vbNewLine & _
            "                ORDER BY " & sHistoricRegionDateColumnName & " ASC" & vbNewLine & vbNewLine
            
          strGenericSQL = strGenericSQL & _
            "            IF @dtTempDate IS NULL" & vbNewLine & _
            "            BEGIN" & vbNewLine & _
            "                SET @dtNextChange_Region = '12/31/9999'" & vbNewLine & _
            "            END" & vbNewLine & _
            "            ELSE" & vbNewLine & _
            "            BEGIN" & vbNewLine & _
            "                SET @dtNextChange_Region = @dtTempDate" & vbNewLine & _
            "            END" & vbNewLine & _
            "        END" & vbNewLine & vbNewLine
        Else
          ' We are using a static region, so get it
          If fBHolSetupOK Then
            strGenericSQL = strGenericSQL & _
              "        SELECT @psPersonnelRegion = " & sStaticRegionColumnName & vbNewLine & _
              "            FROM " & sPersonnelTable & vbNewLine & _
              "            WHERE ID = @iPersonnelID" & vbNewLine & vbNewLine & _
              "        /* Get the Region ID for the persons Region */" & vbNewLine & _
              "        SELECT @iBHolRegionID = ID" & vbNewLine & _
              "            FROM " & sBHolRegionTableName & vbNewLine & _
              "            WHERE " & sBHolRegionColumnName & " = @psPersonnelRegion" & vbNewLine & vbNewLine
          End If
        End If
  
        If fHistoricWP Then
          ' We are using a historic wp so ensure we are getting the right wp for @dCurrentDate
          strGenericSQL = strGenericSQL & _
            "        IF (@dtnextchange_WP IS NULL) OR ((@dtCurrentDate >= @dtNextChange_WP) And (@dtCurrentDate <> '12/31/9999'))" & vbNewLine & _
            "        BEGIN" & vbNewLine & _
            "            /* Get The Employees WP For @dCurrentDate */" & vbNewLine & _
            "            SELECT @psWorkPattern = @psNextWorkPattern" & vbNewLine & vbNewLine & _
            "            /* Get The next change date for WP */" & vbNewLine & _
            "            SET @dtTempDate = null" & vbNewLine & _
            "            SET @psNextWorkPattern = null" & vbNewLine & _
            "            SELECT TOP 1 @dtTempDate = " & sHistoricWPDateColumnName & "," & vbNewLine & _
            "              @psNextWorkPattern = " & sHistoricWPColumnName & vbNewLine & _
            "                FROM " & sHistoricWPTableName & vbNewLine & _
            "                WHERE " & sHistoricWPDateColumnName & " > @dtCurrentDate" & vbNewLine & _
            "                AND ID_" & iPersonnelTableID & " = @iPersonnelID" & vbNewLine & _
            "                ORDER BY " & sHistoricWPDateColumnName & " ASC" & vbNewLine & vbNewLine & _
            "            IF @dtTempDate IS NULL" & vbNewLine & _
            "            BEGIN" & vbNewLine & _
            "                SET @dtNextChange_WP = '12/31/9999'" & vbNewLine & _
            "            END" & vbNewLine & _
            "            ELSE" & vbNewLine & _
            "            BEGIN" & vbNewLine & _
            "                SET @dtNextChange_WP = @dtTempDate" & vbNewLine & _
            "            END" & vbNewLine & _
            "        END" & vbNewLine & vbNewLine
        Else
          ' We are using a static wp, so get it
          strGenericSQL = strGenericSQL & _
            "        SELECT @psWorkPattern = " & sStaticWPColumnName & vbNewLine & _
            "            FROM " & sPersonnelTable & vbNewLine & _
            "            WHERE ID = @iPersonnelID" & vbNewLine & vbNewLine
        End If
  
  
        strGenericSQL = strGenericSQL & _
          "        /* Check if the current date is a work day. */" & vbNewLine & _
          "        SET @fWorkAM = 0" & vbNewLine & _
          "        SET @fWorkPM = 0" & vbNewLine & _
          "        SET @iDayOfWeek = DATEPART(weekday, @dtCurrentDate)" & vbNewLine & vbNewLine
          
        For iLoop = 1 To 7
          strGenericSQL = strGenericSQL & _
            "      IF @iDayOfWeek = " & CStr(iLoop) & vbNewLine & _
            "      BEGIN" & vbNewLine & _
            "        IF LEN(SUBSTRING(@psWorkPattern, " & CStr((iLoop * 2) - 1) & ", 1)) > 0" & vbNewLine & _
            "        BEGIN" & vbNewLine & _
            "          SET @fWorkAM = 1" & vbNewLine & _
            "        END" & vbNewLine & _
            "        IF LEN(SUBSTRING(@psWorkPattern, " & CStr(iLoop * 2) & ", 1)) > 0" & vbNewLine & _
            "        BEGIN" & vbNewLine & _
            "          SET @fWorkPM = 1" & vbNewLine & _
            "        END" & vbNewLine & _
            "      END" & vbNewLine
        Next iLoop
        
        strGenericSQL = strGenericSQL & _
          "        IF (@fWorkAM = 1) OR (@fWorkPM = 1)" & vbNewLine & _
          "        BEGIN" & vbNewLine
          
        If fBHolSetupOK Then
          strGenericSQL = strGenericSQL & _
            "            /* Check that the current date is not a company holiday. */" & vbNewLine & _
            "            SELECT @iCount = COUNT(" & sBHolDateColumnName & ")" & vbNewLine & _
            "                FROM " & sBHolTableName & vbNewLine & _
            "                WHERE " & sBHolDateColumnName & " = @dtCurrentDate" & vbNewLine & _
            "                AND " & sBHolTableName & ".ID_" & iBHolRegionTableID & " = @iBHolRegionID" & vbNewLine & vbNewLine & _
            "            IF @iCount = 0" & vbNewLine & _
            "            BEGIN" & vbNewLine & _
            "                IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
            "                IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
            "            END" & vbNewLine
        Else
          strGenericSQL = strGenericSQL & _
            "            /* We arent using Bholidays, so just add to the result */" & vbNewLine & _
            "            IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine & _
            "            IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5" & vbNewLine
        End If
  
        strGenericSQL = strGenericSQL & _
          "        END" & vbNewLine & vbNewLine & _
          "        /* Move onto the next date. */" & vbNewLine & _
          "        SET @dtCurrentDate = @dtCurrentDate + 1" & vbNewLine & _
          "    END" & vbNewLine & vbNewLine
      End If ' end of the if all static else historic condition
    End If
  
    strProcEnd = "END"
    strUDFEnd = "RETURN @pdblResult" & vbNewLine & "END"
  
    strProcSQL = strProcStart & strGenericSQL & strProcEnd
    strUDFSQL = strUDFStart & strGenericSQL & strUDFEnd
  
    gADOCon.Execute strProcSQL, , adExecuteNoRecords
    If gbEnableUDFFunctions Then gADOCon.Execute strUDFSQL, , adExecuteNoRecords
    
  End If
  
TidyUpAndExit:
  CreateWorkingDaysBetween2DatesStoredProcedure = fCreatedOK
  Exit Function
  
ErrorTrap:
  fCreatedOK = False
  OutputError "Error creating Working Days Between 2 Dates stored procedure (Absence)"
  Resume TidyUpAndExit

End Function





Private Function ReadAbsenceRecordParameters() As Boolean
  ' Read the configured Absence parameters into member variables.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iSSPColsConfigured As Integer
  
  With recModuleSetup
    .Index = "idxModuleParameter"
    
    ' Get the Absence table ID and name.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETABLE
    fOK = Not .NoMatch
    If fOK Then
      'fOK = Not IsNull(!parametervalue)
      fOK = (IIf(IsNull(!parametervalue), 0, val(!parametervalue)) > 0)
    End If
    If Not fOK Then
      mvar_fGeneralOK = False
      mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  'Absence' table not defined."
    Else
      mvar_lngAbsenceTableID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))

      With recTabEdit
        .Index = "idxTableID"
        .Seek "=", mvar_lngAbsenceTableID
      
        fOK = Not .NoMatch
        If fOK Then
          fOK = Not IsNull(!TableName)
        End If
        If Not fOK Then
          mvar_fGeneralOK = False
          
          mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  'Absence' table not found."
        Else
          mvar_sAbsenceTableName = !TableName
        End If
      End With
    End If
    
    If mvar_fGeneralOK Then
      ' Get the Absence Start Date column ID.
      .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESTARTDATE
      fOK = Not .NoMatch
      If fOK Then
        'fOK = Not IsNull(!parametervalue)
        fOK = (IIf(IsNull(!parametervalue), 0, val(!parametervalue)) > 0)
      End If
      If Not fOK Then
        mvar_fGeneralOK = False
        
        mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  '" & mvar_sAbsenceTableName & "' table 'Start Date' column not defined."
      Else
        mvar_lngAbsence_StartDateColumnID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
      
        With recColEdit
          .Index = "idxColumnID"
          .Seek "=", mvar_lngAbsence_StartDateColumnID
        
          fOK = Not .NoMatch
          If fOK Then
            fOK = Not IsNull(!ColumnName)
          End If
          If Not fOK Then
            mvar_fGeneralOK = False
            
            mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  '" & mvar_sAbsenceTableName & "' table 'Start Date' column not found."
          Else
            mvar_sAbsence_StartDateColumnName = !ColumnName
          End If
        End With
      End If
    End If
  
    If mvar_fGeneralOK Then
      ' Get the Absence End Date column ID.
      .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEENDDATE
      fOK = Not .NoMatch
      If fOK Then
        fOK = Not IsNull(!parametervalue)
      End If
      If Not fOK Then
        mvar_fGeneralOK = False
        
        mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  '" & mvar_sAbsenceTableName & "' table 'End Date' column not defined."
      Else
        mvar_lngAbsence_EndDateColumnID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
      
        With recColEdit
          .Index = "idxColumnID"
          .Seek "=", mvar_lngAbsence_EndDateColumnID
        
          fOK = Not .NoMatch
          If fOK Then
            fOK = Not IsNull(!ColumnName)
          End If
          If Not fOK Then
            mvar_fGeneralOK = False
            
            mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  '" & mvar_sAbsenceTableName & "' table 'End Date' column not found."
          Else
            mvar_sAbsence_EndDateColumnName = !ColumnName
          End If
        End With
      End If
    End If
  
    If mvar_fGeneralOK Then
      ' Get the Absence Start Session column ID.
      .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESTARTSESSION
      fOK = Not .NoMatch
      If fOK Then
        fOK = Not IsNull(!parametervalue)
      End If
      If Not fOK Then
        mvar_fGeneralOK = False
        
        mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  '" & mvar_sAbsenceTableName & "' table 'Start Session' column not defined."
      Else
        mvar_lngAbsence_StartSessionColumnID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
      
        With recColEdit
          .Index = "idxColumnID"
          .Seek "=", mvar_lngAbsence_StartSessionColumnID
        
          fOK = Not .NoMatch
          If fOK Then
            fOK = Not IsNull(!ColumnName)
          End If
          If Not fOK Then
            mvar_fGeneralOK = False
            
            mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  '" & mvar_sAbsenceTableName & "' table 'Start Session' column not found."
          Else
            mvar_sAbsence_StartSessionColumnName = !ColumnName
          End If
        End With
      End If
    End If
  
    If mvar_fGeneralOK Then
      ' Get the Absence End Session column ID.
      .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEENDSESSION
      fOK = Not .NoMatch
      If fOK Then
        fOK = Not IsNull(!parametervalue)
      End If
      If Not fOK Then
        mvar_fGeneralOK = False
        
        mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  '" & mvar_sAbsenceTableName & "' table 'End Session' column not defined."
      Else
        mvar_lngAbsence_EndSessionColumnID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
      
        With recColEdit
          .Index = "idxColumnID"
          .Seek "=", mvar_lngAbsence_EndSessionColumnID
        
          fOK = Not .NoMatch
          If fOK Then
            fOK = Not IsNull(!ColumnName)
          End If
          If Not fOK Then
            mvar_fGeneralOK = False
            
            mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  '" & mvar_sAbsenceTableName & "' table 'End Session' column not found."
          Else
            mvar_sAbsence_EndSessionColumnName = !ColumnName
          End If
        End With
      End If
    End If
    
    ' NPG20100607 Fault HRPRO-735
    iSSPColsConfigured = 0
  
    If mvar_fGeneralOK Then
      ' Get the Absence SSP Applies column ID.
      .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESSPAPPLIES
      fOK = Not .NoMatch
      If fOK Then
        fOK = Not IsNull(!parametervalue)
      End If
      If Not fOK Then
        mvar_fSSPGeneralOK = False
        
        mvar_sSSPGeneralMsg = mvar_sSSPGeneralMsg & vbNewLine & "  '" & mvar_sAbsenceTableName & "' table 'SSP Applies' column not defined."
      Else
        mvar_lngAbsence_SSPAppliesColumnID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
      
        With recColEdit
          .Index = "idxColumnID"
          .Seek "=", mvar_lngAbsence_SSPAppliesColumnID
        
          fOK = Not .NoMatch
          If fOK Then
            fOK = Not IsNull(!ColumnName)
          End If
          If Not fOK Then
            mvar_fSSPGeneralOK = False
            
            mvar_sSSPGeneralMsg = mvar_sSSPGeneralMsg & vbNewLine & "  '" & mvar_sAbsenceTableName & "' table 'SSP Applies' column not found."
          Else
            mvar_sAbsence_SSPAppliesColumnName = !ColumnName
            iSSPColsConfigured = iSSPColsConfigured + 1
          End If
        End With
      End If
    End If
  
    If mvar_fGeneralOK Then
      ' Get the Absence SSP Qualifying Days column ID.
      .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESSPQUALIFYINGDAYS
      fOK = Not .NoMatch
      If fOK Then
        fOK = Not IsNull(!parametervalue)
      End If
      If Not fOK Then
        mvar_fSSPGeneralOK = False
        
        mvar_sSSPGeneralMsg = mvar_sSSPGeneralMsg & vbNewLine & "  '" & mvar_sAbsenceTableName & "' table 'SSP Qualifying Days' column not defined."
      Else
        mvar_lngAbsence_QualifyingDaysColumnID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
      
        With recColEdit
          .Index = "idxColumnID"
          .Seek "=", mvar_lngAbsence_QualifyingDaysColumnID
        
          fOK = Not .NoMatch
          If fOK Then
            fOK = Not IsNull(!ColumnName)
          End If
          If Not fOK Then
            mvar_fSSPGeneralOK = False
            
            mvar_sSSPGeneralMsg = mvar_sSSPGeneralMsg & vbNewLine & "  '" & mvar_sAbsenceTableName & "' table 'SSP Qualifying Days' column not found."
          Else
            mvar_sAbsence_QualifyingDaysColumnName = !ColumnName
            iSSPColsConfigured = iSSPColsConfigured + 1
          End If
        End With
      End If
    End If
  
    If mvar_fGeneralOK Then
      ' Get the Absence SSP Waiting Days column ID.
      .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESSPWAITINGDAYS
      fOK = Not .NoMatch
      If fOK Then
        fOK = Not IsNull(!parametervalue)
      End If
      If Not fOK Then
        mvar_fGeneralOK = False
        
        mvar_sSSPGeneralMsg = mvar_sSSPGeneralMsg & vbNewLine & "  '" & mvar_sAbsenceTableName & "' table 'SSP Waiting Days' column not defined."
      Else
        mvar_lngAbsence_WaitingDaysColumnID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
      
        With recColEdit
          .Index = "idxColumnID"
          .Seek "=", mvar_lngAbsence_WaitingDaysColumnID
        
          fOK = Not .NoMatch
          If fOK Then
            fOK = Not IsNull(!ColumnName)
          End If
          If Not fOK Then
            mvar_fSSPGeneralOK = False
            
            mvar_sSSPGeneralMsg = mvar_sSSPGeneralMsg & vbNewLine & "  '" & mvar_sAbsenceTableName & "' table 'SSP Waiting Days' column not found."
          Else
            mvar_sAbsence_WaitingDaysColumnName = !ColumnName
            iSSPColsConfigured = iSSPColsConfigured + 1
          End If
        End With
      End If
    End If
  
    If mvar_fGeneralOK Then
      ' Get the Absence SSP Paid Days column ID.
      .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESSPPAIDDAYS
      fOK = Not .NoMatch
      If fOK Then
        fOK = Not IsNull(!parametervalue)
      End If
      If Not fOK Then
        mvar_fSSPGeneralOK = False
        
        mvar_sSSPGeneralMsg = mvar_sSSPGeneralMsg & vbNewLine & "  '" & mvar_sAbsenceTableName & "' table 'SSP Paid Days' column not defined."
      Else
        mvar_lngAbsence_PaidDaysColumnID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
      
        With recColEdit
          .Index = "idxColumnID"
          .Seek "=", mvar_lngAbsence_PaidDaysColumnID
        
          fOK = Not .NoMatch
          If fOK Then
            fOK = Not IsNull(!ColumnName)
          End If
          If Not fOK Then
            mvar_fSSPGeneralOK = False
            
            mvar_sSSPGeneralMsg = mvar_sSSPGeneralMsg & vbNewLine & "  '" & mvar_sAbsenceTableName & "' table 'SSP Paid Days' column not found."
          Else
            mvar_sAbsence_PaidDaysColumnName = !ColumnName
            iSSPColsConfigured = iSSPColsConfigured + 1
          End If
        End With
      End If
    End If
  
    ' NPG20100607 Fault HRPRO-735
    ' If all SSP columns are undefined don't warn during save any more. If ANY are defined report as necessary
    If iSSPColsConfigured > 0 Then
      mvar_fGeneralOK = mvar_fSSPGeneralOK
      mvar_sGeneralMsg = mvar_sGeneralMsg & mvar_sSSPGeneralMsg
    End If
  
  
  
    If mvar_fGeneralOK Then
      ' Get the Absence Type column ID.
      .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPE
      fOK = Not .NoMatch
      If fOK Then
        fOK = Not IsNull(!parametervalue)
      End If
      If Not fOK Then
        mvar_fGeneralOK = False
        
        mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  '" & mvar_sAbsenceTableName & "' table 'Type' column not defined."
      Else
        mvar_lngAbsence_TypeColumnID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
      
        With recColEdit
          .Index = "idxColumnID"
          .Seek "=", mvar_lngAbsence_TypeColumnID
        
          fOK = Not .NoMatch
          If fOK Then
            fOK = Not IsNull(!ColumnName)
          End If
          If Not fOK Then
            mvar_fGeneralOK = False
            
            mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  '" & mvar_sAbsenceTableName & "' table 'Type' column not found."
          Else
            mvar_sAbsence_TypeColumnName = !ColumnName
          End If
        End With
      End If
    End If
  
    If mvar_fGeneralOK Then
      ' Get the Working Days type.
      .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEWORKINGDAYSTYPE
      If .NoMatch Then
        mvar_iAbsenceWorkingDaysType = 0
      Else
        mvar_iAbsenceWorkingDaysType = IIf(IsNull(!parametervalue), 0, !parametervalue)
      End If
      
      mvar_iAbsenceWorkingDaysNumericValue = 0
      mvar_sAbsenceWorkingDaysPatternValue = ""
      mvar_sAbsenceWorkingDaysColumnName = ""
      mvar_lngAbsenceWorkingDaysTableID = 0
      
      Select Case mvar_iAbsenceWorkingDaysType
        Case 0
          ' Get the Working Days numeric value.
          .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEWORKINGDAYSNUMERICVALUE
          If Not .NoMatch Then
            mvar_iAbsenceWorkingDaysNumericValue = IIf(IsNull(!parametervalue), 0, !parametervalue)
          End If
    
        Case 1
          ' Get the Working Days pattern value.
          .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEWORKINGDAYSPATTERNVALUE
          If Not .NoMatch Then
            mvar_sAbsenceWorkingDaysPatternValue = IIf(IsNull(!parametervalue), "", !parametervalue)
          End If

        Case 2, 3
          ' Get the Absence Type column ID.
          .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEWORKINGDAYSFIELD
          fOK = Not .NoMatch
          If fOK Then
            fOK = Not IsNull(!parametervalue)
          End If
          If Not fOK Then
            mvar_fGeneralOK = False
            
            mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  Absence Working Days column not defined."
          Else
            mvar_lngAbsenceWorkingDaysColumnID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
          
            With recColEdit
              .Index = "idxColumnID"
              .Seek "=", mvar_lngAbsenceWorkingDaysColumnID
            
              fOK = Not .NoMatch
              If fOK Then
                fOK = Not IsNull(!ColumnName)
              End If
              If Not fOK Then
                mvar_fGeneralOK = False
                
                mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  Absence Working Days column not found."
              Else
                mvar_sAbsenceWorkingDaysColumnName = !ColumnName
                mvar_lngAbsenceWorkingDaysTableID = !TableID
              End If
            End With
          End If
          
        Case Else
          mvar_fGeneralOK = False
          
          mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  Working Days type not found."
      End Select
    End If
    
    ' Get the Absence Continuous column.
    If mvar_fGeneralOK Then

      .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCECONTINUOUS
      fOK = Not .NoMatch
      If fOK Then
        fOK = Not IsNull(!parametervalue)
        
        mvar_lngAbsenceContinuousColumnID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
      
        With recColEdit
          .Index = "idxColumnID"
          .Seek "=", mvar_lngAbsenceContinuousColumnID
        
          fOK = Not .NoMatch
          If fOK Then
            fOK = Not IsNull(!ColumnName)
            mvar_sAbsenceContinuousColumnName = !ColumnName
          End If
        End With
      End If
    End If
    
    ' Get the Absence Duration column.
    If mvar_fGeneralOK Then

      .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEDURATION
      fOK = Not .NoMatch
      If fOK Then
        fOK = Not IsNull(!parametervalue)
        
        mvar_lngAbsenceDurationColumnID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
      
        With recColEdit
          .Index = "idxColumnID"
          .Seek "=", mvar_lngAbsenceDurationColumnID
        
          fOK = Not .NoMatch
          If fOK Then
            fOK = Not IsNull(!ColumnName)
            mvar_sAbsenceDurationColumnName = !ColumnName
          End If
        End With
      End If
    End If
    
  End With

  fOK = True
  
TidyUpAndExit:
  ReadAbsenceRecordParameters = fOK
  Exit Function
  
ErrorTrap:
  OutputError "Error reading absence record parameters (Absence)"
  fOK = False
  Resume TidyUpAndExit

End Function

Private Function ReadAbsenceTypeRecordParameters() As Boolean
  ' Read the configured Absence Type parameters into member variables.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  With recModuleSetup
    .Index = "idxModuleParameter"
    
    ' Get the Absence Type table ID and name.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPETABLE
    fOK = Not .NoMatch
    If fOK Then
      fOK = Not IsNull(!parametervalue)
    End If
    If Not fOK Then
      mvar_fGeneralOK = False
      
      mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  'Absence Type' table not defined."
    Else
      mvar_lngAbsenceTypeTableID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))

      With recTabEdit
        .Index = "idxTableID"
        .Seek "=", mvar_lngAbsenceTypeTableID
      
        fOK = Not .NoMatch
        If fOK Then
          fOK = Not IsNull(!TableName)
        End If
        If Not fOK Then
          mvar_fGeneralOK = False
          
          mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  'Absence Type' table not found."
        Else
          mvar_sAbsenceTypeTableName = !TableName
        End If
      End With
    End If
    
    If mvar_fGeneralOK Then
      ' Get the Absence Type - Type column ID.
      .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPETYPE
      fOK = Not .NoMatch
      If fOK Then
        fOK = Not IsNull(!parametervalue)
      End If
      If Not fOK Then
        mvar_fGeneralOK = False
        
        mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  '" & mvar_sAbsenceTypeTableName & "' table 'Type' column not defined."
      Else
        mvar_lngAbsenceType_TypeColumnID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
      
        With recColEdit
          .Index = "idxColumnID"
          .Seek "=", mvar_lngAbsenceType_TypeColumnID
        
          fOK = Not .NoMatch
          If fOK Then
            fOK = Not IsNull(!ColumnName)
          End If
          If Not fOK Then
            mvar_fGeneralOK = False
            
            mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  '" & mvar_sAbsenceTypeTableName & "' table 'Type' column not found."
          Else
            mvar_sAbsenceType_TypeColumnName = !ColumnName
          End If
        End With
      End If
    End If
    
    If mvar_fGeneralOK Then
      ' Get the Absence Type - SSP Applies column ID.
      .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPESSP
      fOK = Not .NoMatch
      If fOK Then
        fOK = Not IsNull(!parametervalue)
      End If
      If Not fOK Then
        ' NPG20100607 Fault HRPRO-735
        ' mvar_fGeneralOK = False
        mvar_fSSPGeneralOK = False
        
        ' NPG20100607 Fault HRPRO-735
        ' mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  '" & mvar_sAbsenceTypeTableName & "' table 'SSP Applies' column not defined."
      Else
        mvar_lngAbsenceType_SSPAppliesColumnID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
      
        With recColEdit
          .Index = "idxColumnID"
          .Seek "=", mvar_lngAbsenceType_SSPAppliesColumnID
        
          fOK = Not .NoMatch
          If fOK Then
            fOK = Not IsNull(!ColumnName)
          End If
          If Not fOK Then
            ' NPG20100607 Fault HRPRO-735
            ' mvar_fGeneralOK = False
            mvar_fSSPGeneralOK = False
            
            ' NPG20100607 Fault HRPRO-735
            ' mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  '" & mvar_sAbsenceTypeTableName & "' table 'SSP Applies' column not found."
          Else
            mvar_sAbsenceType_SSPAppliesColumnName = !ColumnName
          End If
        End With
      End If
    End If
  End With

  fOK = True
  
TidyUpAndExit:
  ReadAbsenceTypeRecordParameters = fOK
  Exit Function
  
ErrorTrap:
  OutputError "Error reading absence type record parameters (Absence)"
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function ReadPersonnelRecordParameters() As Boolean
  ' Read the configured Personnel parameters into member variables.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean

  With recModuleSetup
    .Index = "idxModuleParameter"
    
    ' Get the Personnel table ID and name.
    .Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE
    fOK = Not .NoMatch
    If fOK Then
      fOK = Not IsNull(!parametervalue)
    End If
    If Not fOK Then
      mvar_fGeneralOK = False
      
      mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  'Personnel' table not defined."
    Else
      mvar_lngPersonnelTableID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))

      With recTabEdit
        .Index = "idxTableID"
        .Seek "=", mvar_lngPersonnelTableID
      
        fOK = Not .NoMatch
        If fOK Then
          fOK = Not IsNull(!TableName)
        End If
        If Not fOK Then
          mvar_fGeneralOK = False
          
          mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  'Personnel' table not found."
        Else
          mvar_sPersonnelTableName = !TableName
        End If
      End With
    End If
    
    If mvar_fGeneralOK Then
      ' Get the Personnel Date of Birth column ID.
      mvar_lngPersonnel_DateOfBirthColumnID = 0
      mvar_sPersonnel_DateOfBirthColumnName = ""
      
      .Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_DATEOFBIRTH
      If Not .NoMatch Then
        If Not IsNull(!parametervalue) Then
          mvar_lngPersonnel_DateOfBirthColumnID = IIf(IsNull(!parametervalue), 0, val(!parametervalue))
        
          With recColEdit
            .Index = "idxColumnID"
            .Seek "=", mvar_lngPersonnel_DateOfBirthColumnID
          
            fOK = Not .NoMatch
            If fOK Then
              fOK = Not IsNull(!ColumnName)
            End If
            If Not fOK Then
              mvar_fGeneralOK = False
              
              mvar_sGeneralMsg = mvar_sGeneralMsg & vbNewLine & "  '" & mvar_sPersonnelTableName & "' table 'Date of Birth' column not found."
            Else
              mvar_sPersonnel_DateOfBirthColumnName = !ColumnName
            End If
          End With
        End If
      End If
    End If
  End With

  fOK = True
  
TidyUpAndExit:
  ReadPersonnelRecordParameters = fOK
  Exit Function
  
ErrorTrap:
  OutputError "Error reading personnel record parameters (Absence)"
  fOK = False
  Resume TidyUpAndExit
  
End Function


'End Function


'Private Function DropSSPStoredProcedure() As Boolean
'  ' Drop any existing SSP stored procedure.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim sSQL As String
'
'  fOK = True
'
'  sSQL = "IF EXISTS" & _
'    " (SELECT Name" & _
'    "   FROM sysobjects" & _
'    "   WHERE id = object_id('" & gsSSP_PROCEDURENAME & "')" & _
'    "     AND sysstat & 0xf = 4)" & _
'    " DROP PROCEDURE " & gsSSP_PROCEDURENAME
'  gADOCon.Execute sSQL, , adExecuteNoRecords
'
'TidyUpAndExit:
'  DropSSPStoredProcedure = fOK
'  Exit Function
'
'ErrorTrap:
'  fOK = False
'  OutputError "Error dropping SSP stored procedure (Absence)"
'  Resume TidyUpAndExit
'
'End Function

'Private Function DropWorkingDaysBetween2DatesStoredProcedure() As Boolean
'  ' Drop any existing WorkingDaysBetween2Dates stored procedure.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim sSQL As String
'
'  fOK = True
'
'  sSQL = "IF EXISTS" & _
'    " (SELECT Name" & _
'    "   FROM sysobjects" & _
'    "   WHERE id = object_id('" & gsWorkingDaysBetween2Dates_PROCEDURENAME & "')" & _
'    "     AND sysstat & 0xf = 4)" & _
'    " DROP PROCEDURE " & gsWorkingDaysBetween2Dates_PROCEDURENAME
'  gADOCon.Execute sSQL, , adExecuteNoRecords
'
'TidyUpAndExit:
'  DropWorkingDaysBetween2DatesStoredProcedure = fOK
'  Exit Function
'
'ErrorTrap:
'  fOK = False
'  OutputError "Error dropping Working Days Between 2 Dates stored procedure (Absence)"
'  Resume TidyUpAndExit
'
'End Function
'
'Private Function DropAbsenceBetween2DatesStoredProcedure() As Boolean
'  ' Drop any existing SSP stored procedure.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim sSQL As String
'
'  fOK = True
'
'  sSQL = "IF EXISTS" & _
'    " (SELECT name" & _
'    "   FROM sysobjects" & _
'    "   WHERE id = object_id('sp_ASRFn_AbsenceBetweenTwoDates')" & _
'    "     AND sysstat & 0xf = 4)" & _
'    " DROP PROCEDURE sp_ASRFn_AbsenceBetweenTwoDates"
'  gADOCon.Execute sSQL, , adExecuteNoRecords
'
'TidyUpAndExit:
'  DropAbsenceBetween2DatesStoredProcedure = fOK
'  Exit Function
'
'ErrorTrap:
'  fOK = False
'  OutputError "Error dropping Absence Between 2 Dates stored procedure (Absence)"
'  Resume TidyUpAndExit
'
'End Function
'Private Function DropAbsenceBreakdownCalcStoredProcedure() As Boolean
'  ' Drop any existing Absence Breakdown stored procedure.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim sSQL As String
'
'  fOK = True
'
'  sSQL = "IF EXISTS" & _
'    " (SELECT Name" & _
'    "   FROM sysobjects" & _
'    "   WHERE id = object_id('sp_ASR_AbsenceBreakdown_Calculate')" & _
'    "     AND sysstat & 0xf = 4)" & _
'    " DROP PROCEDURE sp_ASR_AbsenceBreakdown_Calculate"
'  gADOCon.Execute sSQL, , adExecuteNoRecords
'
'TidyUpAndExit:
'  DropAbsenceBreakdownCalcStoredProcedure = fOK
'  Exit Function
'
'ErrorTrap:
'  fOK = False
'  OutputError "Error dropping Absence Breakdown Calc stored procedure (Absence)"
'  Resume TidyUpAndExit
'
'End Function
'Private Function DropAbsenceDurationStoredProcedure() As Boolean
'
'  ' Drop any existing Absence Duration stored procedure.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim sSQL As String
'
'  fOK = True
'
'  sSQL = "IF EXISTS" & _
'    " (SELECT Name" & _
'    "   FROM sysobjects" & _
'    "   WHERE id = object_id('sp_ASRFn_AbsenceDuration')" & _
'    "     AND sysstat & 0xf = 4)" & _
'    " DROP PROCEDURE sp_ASRFn_AbsenceDuration"
'  gADOCon.Execute sSQL, , adExecuteNoRecords
'
'TidyUpAndExit:
'  DropAbsenceDurationStoredProcedure = fOK
'  Exit Function
'
'ErrorTrap:
'  fOK = False
'  OutputError "Error dropping Absence Duration stored procedure (Absence)"
'  Resume TidyUpAndExit
'
'End Function
'
'Private Function DropAbsenceBreakdownStoredProcedure() As Boolean
'  ' Drop any existing SSP stored procedure.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim sSQL As String
'
'  fOK = True
'
'  sSQL = "IF EXISTS" & _
'    " (SELECT *" & _
'    "   FROM sysobjects" & _
'    "   WHERE id = object_id('sp_ASRAbsenceBreakdown')" & _
'    "     AND sysstat & 0xf = 4)" & _
'    " DROP PROCEDURE sp_ASRAbsenceBreakdown"
'  rdoCon.Execute sSQL, rdExecDirect
'
'TidyUpAndExit:
'  DropAbsenceBreakdownStoredProcedure = fOK
'  Exit Function
'
'ErrorTrap:
'  fOK = False
'  OutputError "Error dropping Absence Breakdown stored procedure (Absence)"
'  Resume TidyUpAndExit
'
'End Function
'
'Private Function DropAbsenceDurationUDF() As Boolean
'
'  ' Drop any existing Absence Duration stored procedure.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim sSQL As String
'
'  fOK = True
'
'  sSQL = "IF EXISTS" & _
'    " (SELECT Name" & _
'    "   FROM sysobjects" & _
'    "   WHERE id = object_id('udf_ASRFn_AbsenceDuration')" & _
'    "     AND sysstat & 0xf = 0)" & _
'    " DROP FUNCTION udf_ASRFn_AbsenceDuration"
'  gADOCon.Execute sSQL, , adExecuteNoRecords
'
'TidyUpAndExit:
'  DropAbsenceDurationUDF = fOK
'  Exit Function
'
'ErrorTrap:
'  fOK = False
'  OutputError "Error dropping Absence Duration user defined function (Absence)"
'  Resume TidyUpAndExit
'
'End Function
'
'Private Function DropAbsenceBetween2DatesUDF() As Boolean
'  ' Drop any existing UDF.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim sSQL As String
'
'  fOK = True
'
'  sSQL = "IF EXISTS" & _
'    " (SELECT Name" & _
'    "   FROM sysobjects" & _
'    "   WHERE id = object_id('udf_ASRFn_AbsenceBetweenTwoDates')" & _
'    "     AND sysstat & 0xf = 0)" & _
'    " DROP FUNCTION udf_ASRFn_AbsenceBetweenTwoDates"
'  gADOCon.Execute sSQL, , adExecuteNoRecords
'
'TidyUpAndExit:
'  DropAbsenceBetween2DatesUDF = fOK
'  Exit Function
'
'ErrorTrap:
'  fOK = False
'  OutputError "Error dropping Absence Between 2 Dates user defined function (Absence)"
'  Resume TidyUpAndExit
'
'End Function

Private Function DropWorkingDaysBetween2DatesUDF() As Boolean
  ' Drop any existing WorkingDaysBetween2Dates user defined function.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  
  fOK = True
  
  sSQL = "IF EXISTS" & _
    " (SELECT Name" & _
    "   FROM sysobjects" & _
    "   WHERE id = object_id('udf_ASRFn_WorkingDaysBetweenTwoDates')" & _
    "     AND sysstat & 0xf = 0)" & _
    " DROP FUNCTION udf_ASRFn_WorkingDaysBetweenTwoDates"
  gADOCon.Execute sSQL, , adExecuteNoRecords

TidyUpAndExit:
  DropWorkingDaysBetween2DatesUDF = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  OutputError "Error dropping Working Days Between 2 Dates user defined function (Absence)"
  Resume TidyUpAndExit

End Function

Public Function TableIsUsedInAbsenceBetween2Dates(plngTableID As Long) As Boolean
  ' Return TRUE if the given table is used in the 'AbsenceBetween2Dates' function.
  On Error GoTo ErrorTrap
  
  Dim fIsUsed As Boolean
  Dim lngTempID As Long
  
  fIsUsed = False
  
  ' Get the Absence Table ID
  If Not fIsUsed Then
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETABLE
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
    End If
    
    If plngTableID = lngTempID Then
      fIsUsed = True
    End If
  End If
  
  ' Get the bhol region table id
  If Not fIsUsed Then
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLREGIONTABLE
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
    End If
    
    If plngTableID = lngTempID Then
      fIsUsed = True
    End If
  End If
  
  ' Get the bhol table id
  If Not fIsUsed Then
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLTABLE
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
    End If
    
    If plngTableID = lngTempID Then
      fIsUsed = True
    End If
  End If
  
  ' Set the Personnel table ID variable
  If Not fIsUsed Then
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
    End If
    
    If plngTableID = lngTempID Then
      fIsUsed = True
    End If
  End If
  
  ' Get the Region Setup - Historic Region
  If Not fIsUsed Then
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HREGIONTABLE
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
    End If
    
    If plngTableID = lngTempID Then
      fIsUsed = True
    End If
  End If
  
  ' Get the Region Setup - Historic WP
  If Not fIsUsed Then
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HWORKINGPATTERNTABLE
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
    End If
    
    If plngTableID = lngTempID Then
      fIsUsed = True
    End If
  End If

TidyUpAndExit:
  TableIsUsedInAbsenceBetween2Dates = fIsUsed
  Exit Function
  
ErrorTrap:
  fIsUsed = False
  Resume TidyUpAndExit
  
End Function


Public Function TableIsUsedInWorkingDaysBetween2Dates(plngTableID As Long) As Boolean
  ' Return TRUE if the given table is used in the 'WorkingDaysBetween2Dates' function.
  On Error GoTo ErrorTrap
  
  Dim fIsUsed As Boolean
  Dim lngTempID As Long
  
  fIsUsed = False
  
  ' Get the bhol region table id
  If Not fIsUsed Then
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLREGIONTABLE
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
    End If

    If plngTableID = lngTempID Then
      fIsUsed = True
    End If
  End If

  ' Get the bhol table id
  If Not fIsUsed Then
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLTABLE
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
    End If

    If plngTableID = lngTempID Then
      fIsUsed = True
    End If
  End If

  ' Set the Personnel table ID variable
  If Not fIsUsed Then
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
    End If

    If plngTableID = lngTempID Then
      fIsUsed = True
    End If
  End If

  ' Get the Region Setup - Historic Region
  If Not fIsUsed Then
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HREGIONTABLE
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
    End If

    If plngTableID = lngTempID Then
      fIsUsed = True
    End If
  End If

  ' Get the Region Setup - Historic WP
  If Not fIsUsed Then
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HWORKINGPATTERNTABLE
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
    End If

    If plngTableID = lngTempID Then
      fIsUsed = True
    End If
  End If

TidyUpAndExit:
  TableIsUsedInWorkingDaysBetween2Dates = fIsUsed
  Exit Function
  
ErrorTrap:
  fIsUsed = False
  Resume TidyUpAndExit
  
End Function







Public Function ReadModuleParameter(psModuleKey As String, psParameterKey As String) As String
  ' Return TRUE if the given table is used in the 'AbsenceDuration' function.
  On Error GoTo ErrorTrap
  
  Dim sResult As String
  
  recModuleSetup.Index = "idxModuleParameter"
  recModuleSetup.Seek "=", psModuleKey, psParameterKey
  If Not recModuleSetup.NoMatch Then
    sResult = recModuleSetup!parametervalue
  End If

TidyUpAndExit:
  ReadModuleParameter = sResult
  Exit Function
  
ErrorTrap:
  sResult = "0"
  Resume TidyUpAndExit
  
End Function




Public Function TableIsUsedInAbsenceDuration(plngTableID As Long) As Boolean
  ' Return TRUE if the given table is used in the 'AbsenceDuration' function.
  On Error GoTo ErrorTrap
  
  Dim fIsUsed As Boolean
  Dim lngTempID As Long
  
  fIsUsed = False
  
  ' Get the bhol region table id
  If Not fIsUsed Then
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLREGIONTABLE
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
    End If

    If plngTableID = lngTempID Then
      fIsUsed = True
    End If
  End If

  ' Get the bhol table id
  If Not fIsUsed Then
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLTABLE
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
    End If

    If plngTableID = lngTempID Then
      fIsUsed = True
    End If
  End If

  ' Set the Personnel table ID variable
  If Not fIsUsed Then
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
    End If

    If plngTableID = lngTempID Then
      fIsUsed = True
    End If
  End If

  ' Get the Region Setup - Historic Region
  If Not fIsUsed Then
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HREGIONTABLE
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
    End If

    If plngTableID = lngTempID Then
      fIsUsed = True
    End If
  End If

  ' Get the Region Setup - Historic WP
  If Not fIsUsed Then
    lngTempID = 0
    recModuleSetup.Index = "idxModuleParameter"
    recModuleSetup.Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_HWORKINGPATTERNTABLE
    If Not recModuleSetup.NoMatch Then
      lngTempID = recModuleSetup!parametervalue
    End If

    If plngTableID = lngTempID Then
      fIsUsed = True
    End If
  End If

TidyUpAndExit:
  TableIsUsedInAbsenceDuration = fIsUsed
  Exit Function
  
ErrorTrap:
  fIsUsed = False
  Resume TidyUpAndExit
  
End Function
