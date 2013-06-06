Attribute VB_Name = "modAbsenceSpecifics"
Option Explicit

' Module parameters.
Public gfAbsenceEnabled As Boolean

' Module constants.
Public Const gsMODULEKEY_ABSENCE = "MODULE_ABSENCE"
Public Const gsPARAMETERKEY_ABSENCETABLE = "Param_TableAbsence"
Private Const gsPARAMETERKEY_ABSENCETYPETABLE = "Param_TableAbsenceType"
Private Const gsPARAMETERKEY_ABSENCESCREEN = "Param_ScreenAbsence"
Public Const gsPARAMETERKEY_ABSENCESTARTDATE = "Param_FieldStartDate"
Public Const gsPARAMETERKEY_ABSENCESTARTSESSION = "Param_FieldStartSession"
Public Const gsPARAMETERKEY_ABSENCEENDDATE = "Param_FieldEndDate"
Public Const gsPARAMETERKEY_ABSENCEENDSESSION = "Param_FieldEndSession"
Public Const gsPARAMETERKEY_ABSENCETYPE = "Param_FieldType"
Public Const gsPARAMETERKEY_ABSENCEREASON = "Param_FieldReason"
Public Const gsPARAMETERKEY_ABSENCEDURATION = "Param_FieldDuration"
Public Const gsPARAMETERKEY_ABSENCECONTINUOUS = "Param_FieldContinuous"
'Private Const gsPARAMETERKEY_ABSENCEWORKINGPATTERN = "Param_FieldWorkingPattern"
'Private Const gsPARAMETERKEY_ABSENCEREGION = "Param_FieldAbsenceRegion"
Private Const gsPARAMETERKEY_ABSENCETYPETYPE = "Param_FieldTypeType"
Private Const gsPARAMETERKEY_ABSENCETYPECODE = "Param_FieldTypeCode"
Private Const gsPARAMETERKEY_ABSENCETYPESSP = "Param_FieldTypeSSP"
Private Const gsPARAMETERKEY_ABSENCETYPECALCODE = "Param_FieldTypeCalCode"
'Public Const gsPARAMETERKEY_ABSENCETYPEINCLUDE = "Param_FieldTypeInclude"
'Public Const gsPARAMETERKEY_ABSENCETYPEBRADFORDINDEX = "Param_FieldTypeBradfordIndex"
Private Const gsPARAMETERKEY_ABSENCECALSTARTMONTH = "Param_FieldStartMonth"
Private Const gsPARAMETERKEY_ABSENCECALWEEKENDSHADING = "Param_OtherWeekendShading"
Private Const gsPARAMETERKEY_ABSENCECALBHOLSHADING = "Param_OtherBHolShading"
Private Const gsPARAMETERKEY_ABSENCECALINCLUDEWORKINGDAYSONLY = "Param_OtherIncludeWorkingsDaysOnly"
Private Const gsPARAMETERKEY_ABSENCECALBHOLINCLUDE = "Param_OtherBHolInclude"
Private Const gsPARAMETERKEY_ABSENCECALSHOWCAPTIONS = "Param_OtherShowCaptions"

' Absence Stuff
Public glngAbsenceTableID As Long
Public gsAbsenceTableName As String
Public glngAbsenceScreenID As Long
Public gsAbsenceScreenName As String

Private mvar_lngAbsenceStartDateID As Long
Public gsAbsenceStartDateColumnName As String
Private mvar_lngAbsenceStartSessionID As Long
Public gsAbsenceStartSessionColumnName As String
Private mvar_lngAbsenceEndDateID As Long
Public gsAbsenceEndDateColumnName As String
Private mvar_lngAbsenceEndSessionID As Long
Public gsAbsenceEndSessionColumnName As String
Private mvar_lngAbsenceTypeID As Long
Public gsAbsenceTypeColumnName As String
Private mvar_lngAbsenceReasonID As Long
Public gsAbsenceReasonColumnName As String
Private mvar_lngAbsenceDurationID As Long
Public gsAbsenceDurationColumnName As String

'Private mvar_lngAbsenceWorkingPatternID As Long
'Public gsAbsenceWorkingPatternColumnName As String
'Private mvar_lngAbsenceRegionID As Long
'Public gsAbsenceRegionColumnName As String

' Absence Type Stuff
Public glngAbsenceTypeTableID As Long
Public gsAbsenceTypeTableName As String

Private mvar_lngAbsenceTypeTypeID As Long
Public gsAbsenceTypeTypeColumnName As String
Private mvar_lngAbsenceTypeCodeID As Long
Public gsAbsenceTypeCodeColumnName As String
Private mvar_lngAbsenceTypeSSPID As Long
Public gsAbsenceTypeSSPColumnName As String
Private mvar_lngAbsenceTypeCalCodeID As Long
Public gsAbsenceTypeCalCodeColumnName As String
'Private mvar_lngAbsenceTypeIncludeID As Long
'Public gsAbsenceTypeIncludeColumnName As String
'Private mvar_lngAbsenceTypeBradfordIndexID As Long
'Public gsAbsenceTypeBradfordIndexColumnName As String


' Calendar Stuff
Public giAbsenceCalStartMonth As Integer
Public gfAbsenceCalWeekendShading As Boolean
Public gfAbsenceCalBHolShading As Boolean
Public gfAbsenceCalIncludeWorkingDaysOnly As Boolean
Public gfAbsenceCalBHolInclude As Boolean
Public gfAbsenceCalShowCaptions As Boolean
'Public gsAbsenceCalWorkingPattern As String

Public Sub ReadAbsenceParameters()
  
  ' Read the Absence module parameters from the database.
  glngAbsenceTableID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETABLE))
  If glngAbsenceTableID > 0 Then
    gsAbsenceTableName = datGeneral.GetTableName(glngAbsenceTableID)
  Else
    gsAbsenceTableName = ""
  End If

  glngAbsenceTypeTableID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPETABLE))
  If glngAbsenceTypeTableID > 0 Then
    gsAbsenceTypeTableName = datGeneral.GetTableName(glngAbsenceTypeTableID)
  Else
    gsAbsenceTypeTableName = ""
  End If

  glngAbsenceScreenID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESCREEN))
  If glngAbsenceScreenID > 0 Then
    gsAbsenceScreenName = datGeneral.GetScreenName(glngAbsenceScreenID)
  Else
    gsAbsenceScreenName = ""
  End If

  mvar_lngAbsenceStartDateID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESTARTDATE))
  If mvar_lngAbsenceStartDateID > 0 Then
    gsAbsenceStartDateColumnName = datGeneral.GetColumnName(mvar_lngAbsenceStartDateID)
  Else
    gsAbsenceStartDateColumnName = ""
  End If

  mvar_lngAbsenceStartSessionID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESTARTSESSION))
  If mvar_lngAbsenceStartSessionID > 0 Then
    gsAbsenceStartSessionColumnName = datGeneral.GetColumnName(mvar_lngAbsenceStartSessionID)
  Else
    gsAbsenceStartSessionColumnName = ""
  End If

  mvar_lngAbsenceEndDateID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEENDDATE))
  If mvar_lngAbsenceEndDateID > 0 Then
    gsAbsenceEndDateColumnName = datGeneral.GetColumnName(mvar_lngAbsenceEndDateID)
  Else
    gsAbsenceEndDateColumnName = ""
  End If

  mvar_lngAbsenceEndSessionID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEENDSESSION))
  If mvar_lngAbsenceEndSessionID > 0 Then
    gsAbsenceEndSessionColumnName = datGeneral.GetColumnName(mvar_lngAbsenceEndSessionID)
  Else
    gsAbsenceEndSessionColumnName = ""
  End If

  mvar_lngAbsenceTypeID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPE))
  If mvar_lngAbsenceTypeID > 0 Then
    gsAbsenceTypeColumnName = datGeneral.GetColumnName(mvar_lngAbsenceTypeID)
  Else
    gsAbsenceTypeColumnName = ""
  End If

  mvar_lngAbsenceReasonID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEREASON))
  If mvar_lngAbsenceReasonID > 0 Then
    gsAbsenceReasonColumnName = datGeneral.GetColumnName(mvar_lngAbsenceReasonID)
  Else
    gsAbsenceReasonColumnName = ""
  End If

  mvar_lngAbsenceDurationID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEDURATION))
  If mvar_lngAbsenceDurationID > 0 Then
    gsAbsenceDurationColumnName = datGeneral.GetColumnName(mvar_lngAbsenceDurationID)
  Else
    gsAbsenceDurationColumnName = ""
  End If
  
'  mvar_lngAbsenceWorkingPatternID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEWORKINGPATTERN))
'  If mvar_lngAbsenceWorkingPatternID > 0 Then
'    gsAbsenceWorkingPatternColumnName = datGeneral.GetColumnName(mvar_lngAbsenceWorkingPatternID)
'  Else
'    gsAbsenceWorkingPatternColumnName = ""
'  End If

'  mvar_lngAbsenceRegionID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEREGION))
'  If mvar_lngAbsenceRegionID > 0 Then
'    gsAbsenceRegionColumnName = datGeneral.GetColumnName(mvar_lngAbsenceRegionID)
'  Else
'    gsAbsenceRegionColumnName = ""
'  End If

  mvar_lngAbsenceTypeTypeID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPETYPE))
  If mvar_lngAbsenceTypeTypeID > 0 Then
    gsAbsenceTypeTypeColumnName = datGeneral.GetColumnName(mvar_lngAbsenceTypeTypeID)
  Else
    gsAbsenceTypeTypeColumnName = ""
  End If

  mvar_lngAbsenceTypeCodeID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPECODE))
  If mvar_lngAbsenceTypeCodeID > 0 Then
    gsAbsenceTypeCodeColumnName = datGeneral.GetColumnName(mvar_lngAbsenceTypeCodeID)
  Else
    gsAbsenceTypeCodeColumnName = ""
  End If

  mvar_lngAbsenceTypeSSPID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPESSP))
  If mvar_lngAbsenceTypeSSPID > 0 Then
    gsAbsenceTypeSSPColumnName = datGeneral.GetColumnName(mvar_lngAbsenceTypeSSPID)
  Else
    gsAbsenceTypeSSPColumnName = ""
  End If

  mvar_lngAbsenceTypeCalCodeID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPECALCODE))
  If mvar_lngAbsenceTypeCalCodeID > 0 Then
    gsAbsenceTypeCalCodeColumnName = datGeneral.GetColumnName(mvar_lngAbsenceTypeCalCodeID)
  Else
    gsAbsenceTypeCalCodeColumnName = ""
  End If

  'mvar_lngAbsenceTypeIncludeID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPEINCLUDE))
  'If mvar_lngAbsenceTypeIncludeID > 0 Then
  '  gsAbsenceTypeIncludeColumnName = datGeneral.GetColumnName(mvar_lngAbsenceTypeIncludeID)
  'Else
  '  gsAbsenceTypeIncludeColumnName = ""
  'End If

  'mvar_lngAbsenceTypeBradfordIndexID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPEBRADFORDINDEX))
  'If mvar_lngAbsenceTypeBradfordIndexID > 0 Then
  '  gsAbsenceTypeBradfordIndexColumnName = datGeneral.GetColumnName(mvar_lngAbsenceTypeBradfordIndexID)
  'Else
  '  gsAbsenceTypeBradfordIndexColumnName = ""
  'End If

  giAbsenceCalStartMonth = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCECALSTARTMONTH))

  gfAbsenceCalWeekendShading = IIf(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCECALWEEKENDSHADING) = "TRUE", True, False)
  gfAbsenceCalBHolShading = IIf(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCECALBHOLSHADING) = "TRUE", True, False)
  gfAbsenceCalIncludeWorkingDaysOnly = IIf(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCECALINCLUDEWORKINGDAYSONLY) = "TRUE", True, False)
  gfAbsenceCalBHolInclude = IIf(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCECALBHOLINCLUDE) = "TRUE", True, False)
  gfAbsenceCalShowCaptions = IIf(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCECALSHOWCAPTIONS) = "TRUE", True, False)
  
End Sub

Public Function ValidateAbsenceParameters_BreakdownReport() As Boolean
  
  ' Validate the configuration of the Absence module parameters.

  Dim fValid As Boolean
  Dim strMessage As String
  ''Dim rsType As Recordset
  
  ' Check that the Absence module is installed.
  If gfAbsenceEnabled Then
  
    ' Check the Absence Table ID is valid.
    If Not (glngAbsenceTableID > 0) Then
      strMessage = strMessage & "The Absence table is not defined." & vbCrLf
    End If

    ' Check the Absence Type Table ID is valid.
    If Not (glngAbsenceTypeTableID > 0) Then
      strMessage = strMessage & "The Absence Type table is not defined." & vbCrLf
    End If
    
    ' Check the Start Date ID is valid.
    If Not (mvar_lngAbsenceStartDateID > 0) Then
      strMessage = strMessage & "The Absence Start Date column is not defined." & vbCrLf
    End If

    ' Check the Start Session ID is valid.
    If Not (mvar_lngAbsenceStartSessionID > 0) Then
      strMessage = strMessage & "The Absence Start Session column is not defined." & vbCrLf
    End If

    ' Check the End Date ID is valid.
    If Not (mvar_lngAbsenceEndDateID > 0) Then
      strMessage = strMessage & "The Absence End Date column is not defined." & vbCrLf
    End If

    ' Check the End Session ID is valid.
    If Not (mvar_lngAbsenceEndSessionID > 0) Then
      strMessage = strMessage & "The Absence End Session column is not defined." & vbCrLf
    End If

    ' Check the Type ID is valid.
    If Not (mvar_lngAbsenceTypeID > 0) Then
      strMessage = strMessage & "The Absence Type column is not defined." & vbCrLf
    End If

    ' Check the Reason ID is valid.
    If Not (mvar_lngAbsenceReasonID > 0) Then
      strMessage = strMessage & "The Absence Reason column is not defined." & vbCrLf
    End If

    ' Check the Absence Duration column is valid.
    If Not (mvar_lngAbsenceDurationID > 0) Then
      strMessage = strMessage & "The Absence Duration column is not defined." & vbCrLf
    End If

    ' Check the TypeType ID is valid.
    If Not (mvar_lngAbsenceTypeTypeID > 0) Then
      strMessage = strMessage & "The Absence-Type Type column is not defined." & vbCrLf
    End If

    'MH20030905 Faults 6923 & 6924 - This is no longer required!
    ''' Check the TypeInclude ID is valid.
    ''If Not (mvar_lngAbsenceTypeIncludeID > 0) Then
    ''  strMessage = strMessage & "The Absence-Type Include column is not defined." & vbCrLf
    ''End If

    ''' Check that types exist
    ''If Len(strMessage) = 0 Then
    ''  Set rsType = datGeneral.GetReadOnlyRecords("SELECT *" & _
    ''                                         " FROM " & gsAbsenceTypeTableName & _
    ''                                         " ORDER BY " & gsAbsenceTypeTypeColumnName)
    ''  If rsType.BOF And rsType.EOF Then
    ''    strMessage = strMessage & "You do not have any entries in the '" & gsAbsenceTypeTableName & "' table." & vbCrLf
    ''  End If
   
    ''  Set rsType = Nothing
    ''End If
  
  End If
  
  ' If an error found, warn the user.
  If Len(strMessage) > 0 Then
    strMessage = "The Absence module is not properly configured." & vbCrLf & vbCrLf & strMessage
    COAMsgBox strMessage, vbExclamation, App.ProductName
    fValid = False
  Else
    fValid = True
  End If
  
  ' Return the validation value.
  ValidateAbsenceParameters_BreakdownReport = fValid

End Function


Public Function ValidateAbsenceParameters() As Boolean
  
  ' Validate the configuration of the Absence module parameters,
  ' and the current user's access on the configured columns.

  Dim fValid As Boolean
  Dim strMessage As String

  ' -----------------------------------------------
  If gfAbsenceEnabled Then
    
    ' Check the Absence Table ID is valid.
    If Not (glngAbsenceTableID > 0) Then
      strMessage = strMessage & "The Absence table is not defined." & vbCrLf
    End If

    ' Check the Absence Type Table ID is valid.
    If Not (glngAbsenceTypeTableID > 0) Then
      strMessage = strMessage & "The Absence Type table is not defined." & vbCrLf
    End If
    
    ' Check the Absence Screen ID is valid.
    If Not (glngAbsenceScreenID > 0) Then
      strMessage = strMessage & "The Absence Screen table is not defined." & vbCrLf
    End If
  
    ' Check the Start Date ID is valid.
    If Not (mvar_lngAbsenceStartDateID > 0) Then
      strMessage = strMessage & "The Absence Start Date column is not defined." & vbCrLf
    End If

    ' Check the Start Session ID is valid.
    If Not (mvar_lngAbsenceStartSessionID > 0) Then
      strMessage = strMessage & "The Absence Start Session column is not defined." & vbCrLf
    End If

    ' Check the End Date ID is valid.
    If Not (mvar_lngAbsenceEndDateID > 0) Then
      strMessage = strMessage & "The Absence End Date column is not defined." & vbCrLf
    End If

    ' Check the End Session ID is valid.
    If Not (mvar_lngAbsenceEndSessionID > 0) Then
      strMessage = strMessage & "The Absence End Session column is not defined." & vbCrLf
    End If

    ' Check the Type ID is valid.
    If Not (mvar_lngAbsenceTypeID > 0) Then
      strMessage = strMessage & "The Absence Type column is not defined." & vbCrLf
    End If

    ' Check the Reason ID is valid.
    If Not (mvar_lngAbsenceReasonID > 0) Then
      strMessage = strMessage & "The Absence Reason column is not defined." & vbCrLf
    End If

    ' Check the TypeType ID is valid.
    If Not (mvar_lngAbsenceTypeTypeID > 0) Then
      strMessage = strMessage & "The Absence-Type Type column is not defined." & vbCrLf
    End If

    ' Check the TypeCode ID is valid.
    If Not (mvar_lngAbsenceTypeCodeID > 0) Then
      strMessage = strMessage & "The Absence-Type Code column is not defined." & vbCrLf
    End If

    'JDM - 16/04/02 - Fault 3768 - SSP check not required for absence reports
    ' Check the TypeSSPApplicable ID is valid.
    'If Not (mvar_lngAbsenceTypeSSPID > 0) Then
    '  strMessage = strMessage & "The Absence-Type SSP Applicable column is not defined." & vbCrLf
    'End If

    ' Check the TypeCalendarCode ID is valid.
    If Not (mvar_lngAbsenceTypeCalCodeID > 0) Then
      strMessage = strMessage & "The Absence-Type Calendar Code column is not defined." & vbCrLf
    End If

    'MH20030905 Faults 6923 & 6924 - This is no longer required!
    ''' Check the TypeInclude ID is valid.
    ''If Not (mvar_lngAbsenceTypeIncludeID > 0) Then
    ''  strMessage = strMessage & "The Absence-Type Include column is not defined." & vbCrLf
    ''End If
  Else
  
    ' Absence module is not enabled (this piece of code should never fire...)
    strMessage = "The absence module is not enabled" & vbCrLf
    fValid = False
  
  End If

  ' If an error found, warn the user.
  If Len(strMessage) > 0 Then
    strMessage = "The Absence module is not properly configured." & vbCrLf & vbCrLf & strMessage
    COAMsgBox strMessage, vbExclamation, App.ProductName
    fValid = False
  Else
    fValid = True
  End If

  ' Return the validation value.
  ValidateAbsenceParameters = fValid

End Function

