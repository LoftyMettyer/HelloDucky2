Option Strict Off
Option Explicit On
Module modAbsenceSpecifics
	
	' Module parameters.
	Public gfAbsenceEnabled As Boolean
	
	' Module constants.
	Public Const gsMODULEKEY_ABSENCE As String = "MODULE_ABSENCE"
	Public Const gsPARAMETERKEY_ABSENCETABLE As String = "Param_TableAbsence"
	Private Const gsPARAMETERKEY_ABSENCETYPETABLE As String = "Param_TableAbsenceType"
	Public Const gsPARAMETERKEY_ABSENCESTARTDATE As String = "Param_FieldStartDate"
	Public Const gsPARAMETERKEY_ABSENCESTARTSESSION As String = "Param_FieldStartSession"
	Public Const gsPARAMETERKEY_ABSENCEENDDATE As String = "Param_FieldEndDate"
	Public Const gsPARAMETERKEY_ABSENCEENDSESSION As String = "Param_FieldEndSession"
	Public Const gsPARAMETERKEY_ABSENCETYPE As String = "Param_FieldType"
	Public Const gsPARAMETERKEY_ABSENCEREASON As String = "Param_FieldReason"
	Public Const gsPARAMETERKEY_ABSENCEDURATION As String = "Param_FieldDuration"
	Public Const gsPARAMETERKEY_ABSENCECONTINUOUS As String = "Param_FieldContinuous"
	'Private Const gsPARAMETERKEY_ABSENCEWORKINGPATTERN = "Param_FieldWorkingPattern"
	'Private Const gsPARAMETERKEY_ABSENCEREGION = "Param_FieldAbsenceRegion"
	Private Const gsPARAMETERKEY_ABSENCETYPETYPE As String = "Param_FieldTypeType"
	Private Const gsPARAMETERKEY_ABSENCETYPECODE As String = "Param_FieldTypeCode"
	Private Const gsPARAMETERKEY_ABSENCETYPESSP As String = "Param_FieldTypeSSP"
	Private Const gsPARAMETERKEY_ABSENCETYPECALCODE As String = "Param_FieldTypeCalCode"
	Public Const gsPARAMETERKEY_ABSENCETYPEINCLUDE As String = "Param_FieldTypeInclude"
	Public Const gsPARAMETERKEY_ABSENCETYPEBRADFORDINDEX As String = "Param_FieldTypeBradfordIndex"
	Private Const gsPARAMETERKEY_ABSENCECALSTARTMONTH As String = "Param_FieldStartMonth"
	Private Const gsPARAMETERKEY_ABSENCECALWEEKENDSHADING As String = "Param_OtherWeekendShading"
	Private Const gsPARAMETERKEY_ABSENCECALBHOLSHADING As String = "Param_OtherBHolShading"
	Private Const gsPARAMETERKEY_ABSENCECALINCLUDEWORKINGDAYSONLY As String = "Param_OtherIncludeWorkingsDaysOnly"
	Private Const gsPARAMETERKEY_ABSENCECALBHOLINCLUDE As String = "Param_OtherBHolInclude"
	Private Const gsPARAMETERKEY_ABSENCECALSHOWCAPTIONS As String = "Param_OtherShowCaptions"
	
	' Absence Stuff
	Public glngAbsenceTableID As Integer
	Public gsAbsenceTableName As String
	
	Private mvar_lngAbsenceStartDateID As Integer
	Public gsAbsenceStartDateColumnName As String
	Private mvar_lngAbsenceStartSessionID As Integer
	Public gsAbsenceStartSessionColumnName As String
	Private mvar_lngAbsenceEndDateID As Integer
	Public gsAbsenceEndDateColumnName As String
	Private mvar_lngAbsenceEndSessionID As Integer
	Public gsAbsenceEndSessionColumnName As String
	Private mvar_lngAbsenceTypeID As Integer
	Public gsAbsenceTypeColumnName As String
	Private mvar_lngAbsenceReasonID As Integer
	Public gsAbsenceReasonColumnName As String
	Private mvar_lngAbsenceDurationID As Integer
	Public gsAbsenceDurationColumnName As String
	
	'Private mvar_lngAbsenceWorkingPatternID As Long
	'Public gsAbsenceWorkingPatternColumnName As String
	'Private mvar_lngAbsenceRegionID As Long
	'Public gsAbsenceRegionColumnName As String
	
	' Absence Type Stuff
	Public glngAbsenceTypeTableID As Integer
	Public gsAbsenceTypeTableName As String
	
	Private mvar_lngAbsenceTypeTypeID As Integer
	Public gsAbsenceTypeTypeColumnName As String
	Private mvar_lngAbsenceTypeCodeID As Integer
	Public gsAbsenceTypeCodeColumnName As String
	Private mvar_lngAbsenceTypeSSPID As Integer
	Public gsAbsenceTypeSSPColumnName As String
	Private mvar_lngAbsenceTypeCalCodeID As Integer
	Public gsAbsenceTypeCalCodeColumnName As String
	Private mvar_lngAbsenceTypeIncludeID As Integer
	Public gsAbsenceTypeIncludeColumnName As String
	Private mvar_lngAbsenceTypeBradfordIndexID As Integer
	Public gsAbsenceTypeBradfordIndexColumnName As String
	
	
	' Calendar Stuff
	Public giAbsenceCalStartMonth As Short
	Public gfAbsenceCalWeekendShading As Boolean
	Public gfAbsenceCalBHolShading As Boolean
	Public gfAbsenceCalIncludeWorkingDaysOnly As Boolean
	Public gfAbsenceCalBHolInclude As Boolean
	Public gfAbsenceCalShowCaptions As Boolean
	Public gsAbsenceCalWorkingPattern As String
	
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
		
		mvar_lngAbsenceTypeIncludeID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPEINCLUDE))
		If mvar_lngAbsenceTypeIncludeID > 0 Then
			gsAbsenceTypeIncludeColumnName = datGeneral.GetColumnName(mvar_lngAbsenceTypeIncludeID)
		Else
			gsAbsenceTypeIncludeColumnName = ""
		End If
		
		mvar_lngAbsenceTypeBradfordIndexID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPEBRADFORDINDEX))
		If mvar_lngAbsenceTypeBradfordIndexID > 0 Then
			gsAbsenceTypeBradfordIndexColumnName = datGeneral.GetColumnName(mvar_lngAbsenceTypeBradfordIndexID)
		Else
			gsAbsenceTypeBradfordIndexColumnName = ""
		End If
		
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
    Dim strMessage As String = ""
		Dim rsType As ADODB.Recordset
		
		' Check that the Absence module is installed.
		If gfAbsenceEnabled Then
			
			' Check the Absence Table ID is valid.
			If Not (glngAbsenceTableID > 0) Then
				strMessage = strMessage & "The Absence table is not defined." & vbNewLine
			End If
			
			' Check the Absence Type Table ID is valid.
			If Not (glngAbsenceTypeTableID > 0) Then
				strMessage = strMessage & "The Absence Type table is not defined." & vbNewLine
			End If
			
			' Check the Start Date ID is valid.
			If Not (mvar_lngAbsenceStartDateID > 0) Then
				strMessage = strMessage & "The Absence Start Date column is not defined." & vbNewLine
			End If
			
			' Check the Start Session ID is valid.
			If Not (mvar_lngAbsenceStartSessionID > 0) Then
				strMessage = strMessage & "The Absence Start Session column is not defined." & vbNewLine
			End If
			
			' Check the End Date ID is valid.
			If Not (mvar_lngAbsenceEndDateID > 0) Then
				strMessage = strMessage & "The Absence End Date column is not defined." & vbNewLine
			End If
			
			' Check the End Session ID is valid.
			If Not (mvar_lngAbsenceEndSessionID > 0) Then
				strMessage = strMessage & "The Absence End Session column is not defined." & vbNewLine
			End If
			
			' Check the Type ID is valid.
			If Not (mvar_lngAbsenceTypeID > 0) Then
				strMessage = strMessage & "The Absence Type column is not defined." & vbNewLine
			End If
			
			' Check the Reason ID is valid.
			If Not (mvar_lngAbsenceReasonID > 0) Then
				strMessage = strMessage & "The Absence Reason column is not defined." & vbNewLine
			End If
			
			' Check the Absence Duration column is valid.
			If Not (mvar_lngAbsenceDurationID > 0) Then
				strMessage = strMessage & "The Absence Duration column is not defined." & vbNewLine
			End If
			
			' Check the TypeType ID is valid.
			If Not (mvar_lngAbsenceTypeTypeID > 0) Then
				strMessage = strMessage & "The Absence-Type Type column is not defined." & vbNewLine
			End If
			
			' Check the TypeInclude ID is valid.
			If Not (mvar_lngAbsenceTypeIncludeID > 0) Then
				strMessage = strMessage & "The Absence-Type Include column is not defined." & vbNewLine
			End If
			
			' Check that types exist
			If Len(strMessage) = 0 Then
				rsType = datGeneral.GetReadOnlyRecords("SELECT *" & " FROM " & gsAbsenceTypeTableName & " ORDER BY " & gsAbsenceTypeTypeColumnName)
				If rsType.BOF And rsType.EOF Then
					strMessage = strMessage & "You do not have any entries in the '" & gsAbsenceTypeTableName & "' table." & vbNewLine
				End If
				
				'UPGRADE_NOTE: Object rsType may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				rsType = Nothing
			End If
			
		End If
		
		' If an error found, warn the user.
		If Len(strMessage) > 0 Then
			strMessage = "The Absence module is not properly configured." & vbNewLine & vbNewLine & strMessage
			'NO MSGBOX ON THE SERVER ! - MsgBox strMessage, vbExclamation, App.ProductName
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
				strMessage = strMessage & "The Absence table is not defined." & vbNewLine
			End If
			
			' Check the Absence Type Table ID is valid.
			If Not (glngAbsenceTypeTableID > 0) Then
				strMessage = strMessage & "The Absence Type table is not defined." & vbNewLine
			End If
			
			' Check the Start Date ID is valid.
			If Not (mvar_lngAbsenceStartDateID > 0) Then
				strMessage = strMessage & "The Absence Start Date column is not defined." & vbNewLine
			End If
			
			' Check the Start Session ID is valid.
			If Not (mvar_lngAbsenceStartSessionID > 0) Then
				strMessage = strMessage & "The Absence Start Session column is not defined." & vbNewLine
			End If
			
			' Check the End Date ID is valid.
			If Not (mvar_lngAbsenceEndDateID > 0) Then
				strMessage = strMessage & "The Absence End Date column is not defined." & vbNewLine
			End If
			
			' Check the End Session ID is valid.
			If Not (mvar_lngAbsenceEndSessionID > 0) Then
				strMessage = strMessage & "The Absence End Session column is not defined." & vbNewLine
			End If
			
			' Check the Type ID is valid.
			If Not (mvar_lngAbsenceTypeID > 0) Then
				strMessage = strMessage & "The Absence Type column is not defined." & vbNewLine
			End If
			
			' Check the Reason ID is valid.
			If Not (mvar_lngAbsenceReasonID > 0) Then
				strMessage = strMessage & "The Absence Reason column is not defined." & vbNewLine
			End If
			
			' Check the TypeType ID is valid.
			If Not (mvar_lngAbsenceTypeTypeID > 0) Then
				strMessage = strMessage & "The Absence-Type Type column is not defined." & vbNewLine
			End If
			
			' Check the TypeCode ID is valid.
			If Not (mvar_lngAbsenceTypeCodeID > 0) Then
				strMessage = strMessage & "The Absence-Type Code column is not defined." & vbNewLine
			End If
			
			' JDM - Fault 3768 - 16/04/02 - SSP applicable not required for absence calendar.
			' Check the TypeSSPApplicable ID is valid.
			'If Not (mvar_lngAbsenceTypeSSPID > 0) Then
			'  strMessage = strMessage & "The Absence-Type SSP Applicable column is not defined." & vbNewLine
			'End If
			
			' Check the TypeCalendarCode ID is valid.
			If Not (mvar_lngAbsenceTypeCalCodeID > 0) Then
				strMessage = strMessage & "The Absence-Type Calendar Code column is not defined." & vbNewLine
			End If
			
			' Check the TypeInclude ID is valid.
			If Not (mvar_lngAbsenceTypeIncludeID > 0) Then
				strMessage = strMessage & "The Absence-Type Include column is not defined." & vbNewLine
			End If
		Else
			
			' Absence module is not enabled (this piece of code should never fire...)
			strMessage = "The absence module is not enabled" & vbNewLine
			fValid = False
			
		End If
		
		' If an error found, warn the user.
		If Len(strMessage) > 0 Then
			strMessage = "The Absence module is not properly configured." & vbNewLine & vbNewLine & strMessage
			'NO MSGBOX ON THE SERVER ! - MsgBox strMessage, vbExclamation, App.ProductName
			fValid = False
		Else
			fValid = True
		End If
		
		' Return the validation value.
		ValidateAbsenceParameters = fValid
		
	End Function
	
	Public Function CheckPermission_Absence() As Boolean
		
		Dim pblnOK As Boolean
		Dim pstrBadColumn As String
		Dim objTable As CTablePrivilege
		Dim objColumn As CColumnPrivileges
		Dim pblnColumnOK As Boolean
		
		pblnOK = True
		
		' Retrieve the correct asrsyschildview for the absence table
		objTable = gcoTablePrivileges.FindTableID(glngAbsenceTableID)
		
		If objTable.AllowSelect = False Then
			pblnOK = False
			pstrBadColumn = "Absence Table"
		End If
		
		gsAbsenceTableName = objTable.RealSource
		
		' Now check that read permission is available for the required columns
		objColumn = GetColumnPrivileges((objTable.TableName))
		
		' Check Absence Start Date
		If pblnOK Then
			pblnOK = objColumn.IsValid(gsAbsenceStartDateColumnName)
			If pblnOK Then
				pblnOK = objColumn.Item(gsAbsenceStartDateColumnName).AllowSelect
				If pblnOK = False Then pstrBadColumn = "Absence 'Start Date' column"
			Else
				pstrBadColumn = "Absence 'Start Date' column"
			End If
		End If
		
		' Check Absence Start Session
		If pblnOK Then
			pblnOK = objColumn.IsValid(gsAbsenceStartSessionColumnName)
			If pblnOK Then
				pblnOK = objColumn.Item(gsAbsenceStartSessionColumnName).AllowSelect
				If pblnOK = False Then pstrBadColumn = "Absence 'Start Session' column"
			Else
				pstrBadColumn = "Absence 'Start Session' column"
			End If
		End If
		
		' Check Absence End Date
		If pblnOK Then
			pblnOK = objColumn.IsValid(gsAbsenceEndDateColumnName)
			If pblnOK Then
				pblnOK = objColumn.Item(gsAbsenceEndDateColumnName).AllowSelect
				If pblnOK = False Then pstrBadColumn = "Absence 'End Date' column"
			Else
				pstrBadColumn = "Absence 'End Date' column"
			End If
		End If
		
		' Check Absence End Session
		If pblnOK Then
			pblnOK = objColumn.IsValid(gsAbsenceEndSessionColumnName)
			If pblnOK Then
				pblnOK = objColumn.Item(gsAbsenceEndSessionColumnName).AllowSelect
				If pblnOK = False Then pstrBadColumn = "Absence 'End Session' column"
			Else
				pstrBadColumn = "Absence 'End Session' column"
			End If
		End If
		
		' Check Absence Type
		If pblnOK Then
			pblnOK = objColumn.IsValid(gsAbsenceTypeColumnName)
			If pblnOK Then
				pblnOK = objColumn.Item(gsAbsenceTypeColumnName).AllowSelect
				If pblnOK = False Then pstrBadColumn = "Absence 'Type' column"
			Else
				pstrBadColumn = "Absence 'Type' column"
			End If
		End If
		
		' Check Absence Reason
		If pblnOK Then
			pblnOK = objColumn.IsValid(gsAbsenceReasonColumnName)
			If pblnOK Then
				pblnOK = objColumn.Item(gsAbsenceReasonColumnName).AllowSelect
				If pblnOK = False Then pstrBadColumn = "Absence 'Reason' column"
			Else
				pstrBadColumn = "Absence 'Reason' column"
			End If
		End If
		
		'UPGRADE_NOTE: Object objTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objTable = Nothing
		'UPGRADE_NOTE: Object objColumn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objColumn = Nothing
		
		CheckPermission_Absence = pblnOK
		
	End Function
End Module