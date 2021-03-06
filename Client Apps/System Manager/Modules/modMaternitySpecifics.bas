Attribute VB_Name = "modMaternitySpecifics"
Option Explicit

Public Const gsPARAMETERKEY_ABSENCEPARENTALREGION = "Param_AbsenceParentalLeaveRegion"

Public Function ConfigureMaternitySpecifics() As Boolean
  
  Dim fOK As Boolean

  On Error GoTo ErrorTrap

  fOK = CreateParentalLeaveEntitlementUDF()

  If fOK Then
    fOK = CreateParentalLeaveTakenUDF()
  End If

  If fOK Then
    fOK = CreateMaternityExpectedReturnDateUDF()
  End If


TidyUpAndExit:
  ConfigureMaternitySpecifics = fOK
  Exit Function
  
ErrorTrap:
  OutputError "Error configuring Maternity specifics"
  fOK = False
  Resume TidyUpAndExit

End Function


Private Function CreateParentalLeaveEntitlementUDF() As Boolean
  
  Dim strDependantsTableName As String
  Dim strDepDateOfBirthColumn As String
  Dim strDepDateAdoptedColumn As String
  Dim strDepDisabledColumn As String
  Dim lngPersID As Long
  Dim strPersonnelTableName As String
  Dim strPersonnelRegionColumn As String

  Dim fOK As Boolean
  Dim sSQL As String
  
  On Error GoTo ErrorTrap

  fOK = True
  DropFunction "udfsys_parentalleaveentitlement"
  
  
  strDependantsTableName = GetModuleSetupValue(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_DEPENDANTSTABLE, "T")
  strDepDateOfBirthColumn = GetModuleSetupValue(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_DEPENDANTSDATEOFBIRTH, "C")
  strDepDateAdoptedColumn = GetModuleSetupValue(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_DEPENDANTSADOPTEDDATE, "C")
  strDepDisabledColumn = GetModuleSetupValue(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_DEPENDANTSDISABLED, "C")
  
  If strDependantsTableName = vbNullString Or _
     strDepDateOfBirthColumn = vbNullString Or _
     strDepDateAdoptedColumn = vbNullString Or _
     strDepDisabledColumn = vbNullString Then
              
        sSQL = "/* ------------------------------------------------ */" & vbNewLine & _
            "/* Paternity module user defined function.        */" & vbNewLine & _
            "/* Automatically generated by the System manager.   */" & vbNewLine & _
            "/* ------------------------------------------------ */" & vbNewLine & _
          "CREATE FUNCTION [dbo].[udfsys_parentalleaveentitlement] (" & vbNewLine & _
          "@iDependantID  integer)" & vbNewLine & _
          "RETURNS float" & vbNewLine & _
          "AS" & vbNewLine & _
          "BEGIN" & vbNewLine & _
          "    RETURN 0;" & vbNewLine & _
          "END"
        
  Else
  
    sSQL = "/* ------------------------------------------------ */" & vbNewLine & _
        "/* Paternity module user defined function.        */" & vbNewLine & _
        "/* Automatically generated by the System manager.   */" & vbNewLine & _
        "/* ------------------------------------------------ */" & vbNewLine & _
      "CREATE FUNCTION [dbo].[udfsys_parentalleaveentitlement] (" & vbNewLine & _
      "@iDependantID  integer)" & vbNewLine & _
      "RETURNS float" & vbNewLine & _
      "AS" & vbNewLine & _
      "BEGIN" & vbNewLine & vbNewLine
  
    sSQL = sSQL & _
      "  DECLARE  @pdblResult  float," & vbNewLine & _
      "           @DateOfBirth datetime," & vbNewLine & _
      "           @AdoptedDate datetime," & vbNewLine & _
      "           @Disabled    bit," & vbNewLine & _
      "           @Region      varchar(MAX);" & vbNewLine & vbNewLine
  
    strPersonnelRegionColumn = GetModuleSetupValue(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEPARENTALREGION, "C")
    If strPersonnelRegionColumn <> vbNullString Then
      lngPersID = val(GetModuleSetupValue(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE, ""))
      strPersonnelTableName = GetTableName(lngPersID)
  
      sSQL = sSQL & _
        "  SELECT @DateOfBirth = " & strDepDateOfBirthColumn & _
               ", @AdoptedDate = " & strDepDateAdoptedColumn & _
               ", @Disabled = " & strDepDisabledColumn & _
               ", @Region = " & strPersonnelRegionColumn & vbNewLine & _
        "  FROM [" & strDependantsTableName & "] " & _
        "  JOIN [" & strPersonnelTableName & "] ON [" & strPersonnelTableName & "].ID = [" & strDependantsTableName & "].ID_" & CStr(lngPersID) & _
        "  WHERE [" & strDependantsTableName & "].ID = @iDependantID" & vbNewLine & vbNewLine
    Else
      sSQL = sSQL & _
        "  SELECT @DateOfBirth = " & strDepDateOfBirthColumn & _
               ", @AdoptedDate = " & strDepDateAdoptedColumn & _
               ", @Disabled = " & strDepDisabledColumn & _
               ", @Region = ''" & vbNewLine & _
        "  FROM [" & strDependantsTableName & "] " & _
        "  WHERE [" & strDependantsTableName & "].ID = @iDependantID" & vbNewLine & vbNewLine
    End If
  
    sSQL = sSQL & _
      "  SELECT @pdblResult = dbo.[udfstat_ParentalLeaveEntitlement](@DateOfBirth, @AdoptedDate, @Disabled, @Region)" & vbNewLine & vbNewLine
      
    sSQL = sSQL & vbNewLine & _
      "  RETURN ISNULL(@pdblResult, 0)" & vbNewLine & _
      "END" & vbNewLine
  
  End If
  
  gADOCon.Execute sSQL, , adExecuteNoRecords

TidyUpAndExit:
  CreateParentalLeaveEntitlementUDF = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  OutputError "Error creating Parental Leave Entitlement (Maternity)"
  Resume TidyUpAndExit

End Function


Private Function CreateParentalLeaveTakenUDF() As Boolean

  Dim lngPersID As Long
  Dim strDependantsTableName As String
  Dim strDepChildNoColumn As String
  Dim strAbsenceTableName As String
  Dim strAbsChildNoColumn As String
  Dim strAbsDuration As String
  Dim strAbsTypeColumn As String
  Dim strAbsParentalType As String
  
  Dim fOK As Boolean
  Dim sSQL As String
  
  On Error GoTo ErrorTrap
  
  fOK = True
  DropFunction "udfsys_parentalleavetaken"
  
  lngPersID = val(GetModuleSetupValue(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE, ""))
  strDependantsTableName = GetModuleSetupValue(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_DEPENDANTSTABLE, "T")
  strDepChildNoColumn = GetModuleSetupValue(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_DEPENDANTSCHILDNO, "C")
  strAbsenceTableName = GetModuleSetupValue(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETABLE, "T")
  strAbsChildNoColumn = GetModuleSetupValue(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCECHILDNO, "C")
  strAbsDuration = GetModuleSetupValue(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEDURATION, "C")
  strAbsTypeColumn = GetModuleSetupValue(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPE, "C")
  strAbsParentalType = GetModuleSetupValue(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEPARENTALLEAVETYPE, "")
  
  If lngPersID = 0 Or _
     strDependantsTableName = vbNullString Or _
     strDepChildNoColumn = vbNullString Or _
     strAbsenceTableName = vbNullString Or _
     strAbsChildNoColumn = vbNullString Or _
     strAbsDuration = vbNullString Or _
     strAbsTypeColumn = vbNullString Or _
     strAbsParentalType = vbNullString Then
        
        sSQL = "/* ------------------------------------------------ */" & vbNewLine & _
            "/* Paternity module user defined function.      */" & vbNewLine & _
            "/* Automatically generated by the System manager.   */" & vbNewLine & _
            "/* ------------------------------------------------ */" & vbNewLine & _
          "CREATE FUNCTION [dbo].[udfsys_parentalleavetaken] (" & vbNewLine & _
          "@iDependantID integer)" & vbNewLine & _
          "RETURNS float" & vbNewLine & _
          "AS" & vbNewLine & _
          "BEGIN" & vbNewLine & vbNewLine & _
          "    RETURN 0;" & vbNewLine & _
          "END"
        
  Else
    
    sSQL = "/* ------------------------------------------------ */" & vbNewLine & _
        "/* Paternity module user defined function.      */" & vbNewLine & _
        "/* Automatically generated by the System manager.   */" & vbNewLine & _
        "/* ------------------------------------------------ */" & vbNewLine & _
      "CREATE FUNCTION [dbo].[udfsys_parentalleavetaken] (" & vbNewLine & _
      "@iDependantID integer)" & vbNewLine & _
      "RETURNS float" & vbNewLine & _
      "AS" & vbNewLine & _
      "BEGIN" & vbNewLine & vbNewLine
  
    sSQL = sSQL & _
      "  DECLARE @pdblResult  float," & vbNewLine & _
      "          @ChildNo     integer," & vbNewLine & _
      "          @PersID      integer;" & vbNewLine & vbNewLine
  
    sSQL = sSQL & _
      "  SELECT @PersID = ID_" & CStr(lngPersID) & _
             ", @ChildNo = " & strDepChildNoColumn & vbNewLine & _
      "  FROM " & strDependantsTableName & " WHERE ID = @iDependantID;" & vbNewLine & vbNewLine _
  
    sSQL = sSQL & _
      "  SELECT @pdblResult = SUM(" & strAbsDuration & ") FROM " & strAbsenceTableName & vbNewLine & _
      "  WHERE ID_" & CStr(lngPersID) & " = @PersID" & vbNewLine & _
      "  AND " & strAbsChildNoColumn & " = @ChildNo" & vbNewLine & _
      "  AND " & strAbsTypeColumn & " = '" & strAbsParentalType & "';" & vbNewLine & vbNewLine _
  
    sSQL = sSQL & vbNewLine & _
      "   RETURN ISNULL(@pdblResult, 0);" & vbNewLine & _
      "END" & vbNewLine

  End If

  ' Lets commit this baby...
  gADOCon.Execute sSQL, , adExecuteNoRecords


TidyUpAndExit:
  CreateParentalLeaveTakenUDF = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  OutputError "Error creating Parental Leave Taken (Maternity)"
  Resume TidyUpAndExit



End Function

Private Function CreateMaternityExpectedReturnDateUDF() As Boolean

  Dim strMaternityTableName As String
  Dim strMatEWCDateColumn As String
  Dim strMatLeaveTypeColumn As String
  Dim strMatLeaveStartColumn As String
  Dim strMatBabyBirthDateColumn As String

  Dim fOK As Boolean
  Dim sSQL As String
  
  On Error GoTo ErrorTrap
  
  
  fOK = True
  DropFunction "udfsys_maternityexpectedreturndate"
  
  
  strMaternityTableName = GetModuleSetupValue(gsMODULEKEY_MATERNITY, gsPARAMETERKEY_MATERNITYTABLE, "T")
  strMatEWCDateColumn = GetModuleSetupValue(gsMODULEKEY_MATERNITY, gsPARAMETERKEY_MATERNITYEWCDATECOLUMN, "C")
  strMatLeaveTypeColumn = GetModuleSetupValue(gsMODULEKEY_MATERNITY, gsPARAMETERKEY_MATERNITYLEAVETYPECOLUMN, "C")
  strMatLeaveStartColumn = GetModuleSetupValue(gsMODULEKEY_MATERNITY, gsPARAMETERKEY_MATERNITYLEAVESTARTCOLUMN, "C")
  strMatBabyBirthDateColumn = GetModuleSetupValue(gsMODULEKEY_MATERNITY, gsPARAMETERKEY_MATERNITYBABYBIRTHCOLUMN, "C")
  
  If strMaternityTableName = vbNullString Or _
     strMatEWCDateColumn = vbNullString Or _
     strMatLeaveTypeColumn = vbNullString Or _
     strMatLeaveStartColumn = vbNullString Or _
     strMatBabyBirthDateColumn = vbNullString Then
        
      sSQL = "/* ------------------------------------------------ */" & vbNewLine & _
          "/* Maternity module user defined function.      */" & vbNewLine & _
          "/* Automatically generated by the System manager.   */" & vbNewLine & _
          "/* ------------------------------------------------ */" & vbNewLine & _
        "CREATE FUNCTION [dbo].[udfsys_maternityexpectedreturndate] (" & vbNewLine & _
        "@iMaternityID integer)" & vbNewLine & _
        "RETURNS datetime" & vbNewLine & _
        "AS" & vbNewLine & _
        "BEGIN" & vbNewLine & _
        "    RETURN GETDATE();" & vbNewLine & _
        "END"
        
  Else
 
    sSQL = "/* ------------------------------------------------ */" & vbNewLine & _
        "/* Maternity module user defined function.      */" & vbNewLine & _
        "/* Automatically generated by the System manager.   */" & vbNewLine & _
        "/* ------------------------------------------------ */" & vbNewLine & _
      "CREATE FUNCTION [dbo].[udfsys_maternityexpectedreturndate] (" & vbNewLine & _
      "@iMaternityID integer)" & vbNewLine & _
      "RETURNS datetime" & vbNewLine & _
      "AS" & vbNewLine & _
      "BEGIN" & vbNewLine & vbNewLine
  
    sSQL = sSQL & _
      "  DECLARE  @pdblResult     datetime," & vbNewLine & _
      "           @EWCDate        datetime," & vbNewLine & _
      "           @LeaveStart     datetime," & vbNewLine & _
      "           @BabyBirthDate  datetime," & vbNewLine & _
      "           @Ordinary       varchar(MAX);" & vbNewLine & vbNewLine
  
    sSQL = sSQL & _
      "  SELECT @Ordinary = " & strMatLeaveTypeColumn & vbNewLine & _
             ", @EWCDate = " & strMatEWCDateColumn & _
             ", @LeaveStart = " & strMatLeaveStartColumn & _
             ", @BabyBirthDate = " & strMatBabyBirthDateColumn & vbNewLine & _
      "  FROM " & strMaternityTableName & _
      "  WHERE ID = @iMaternityID;" & vbNewLine & vbNewLine
  
    sSQL = sSQL & _
      "  SELECT @pdblResult = dbo.[udfstat_MaternityExpectedReturn](@EWCDate, @LeaveStart, @BabyBirthDate, @Ordinary);" & vbNewLine & vbNewLine
    
    sSQL = sSQL & vbNewLine & _
      "  RETURN @pdblResult;" & vbNewLine & _
      "END" & vbNewLine

  End If

  gADOCon.Execute sSQL, , adExecuteNoRecords

TidyUpAndExit:
  CreateMaternityExpectedReturnDateUDF = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  OutputError "Error creating Maternity Expected Return Date Procedure"
  Resume TidyUpAndExit

End Function
