Option Strict Off
Option Explicit On

Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Metadata
Imports System.Collections.ObjectModel

Namespace ModuleSpecifics

	Friend Class modAbsenceSpecifics
		Inherits BaseModuleSpecific

		Public Sub New(value As SessionInfo)
			MyBase.New(value)
		End Sub

		' Module parameters.
		Public gfAbsenceEnabled As Boolean

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
				gsAbsenceTableName = _tables.GetById(glngAbsenceTableID).Name
			Else
				gsAbsenceTableName = ""
			End If

			glngAbsenceTypeTableID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPETABLE))
			If glngAbsenceTypeTableID > 0 Then
				gsAbsenceTypeTableName = _tables.GetById(glngAbsenceTypeTableID).Name
			Else
				gsAbsenceTypeTableName = ""
			End If

			mvar_lngAbsenceStartDateID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESTARTDATE))
			If mvar_lngAbsenceStartDateID > 0 Then
				gsAbsenceStartDateColumnName = _columns.GetById(mvar_lngAbsenceStartDateID).Name
			Else
				gsAbsenceStartDateColumnName = ""
			End If

			mvar_lngAbsenceStartSessionID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESTARTSESSION))
			If mvar_lngAbsenceStartSessionID > 0 Then
				gsAbsenceStartSessionColumnName = _columns.GetById(mvar_lngAbsenceStartSessionID).Name
			Else
				gsAbsenceStartSessionColumnName = ""
			End If

			mvar_lngAbsenceEndDateID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEENDDATE))
			If mvar_lngAbsenceEndDateID > 0 Then
				gsAbsenceEndDateColumnName = _columns.GetById(mvar_lngAbsenceEndDateID).Name
			Else
				gsAbsenceEndDateColumnName = ""
			End If

			mvar_lngAbsenceEndSessionID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEENDSESSION))
			If mvar_lngAbsenceEndSessionID > 0 Then
				gsAbsenceEndSessionColumnName = _columns.GetById(mvar_lngAbsenceEndSessionID).Name
			Else
				gsAbsenceEndSessionColumnName = ""
			End If

			mvar_lngAbsenceTypeID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPE))
			If mvar_lngAbsenceTypeID > 0 Then
				gsAbsenceTypeColumnName = _columns.GetById(mvar_lngAbsenceTypeID).Name
			Else
				gsAbsenceTypeColumnName = ""
			End If

			mvar_lngAbsenceReasonID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEREASON))
			If mvar_lngAbsenceReasonID > 0 Then
				gsAbsenceReasonColumnName = _columns.GetById(mvar_lngAbsenceReasonID).Name
			Else
				gsAbsenceReasonColumnName = ""
			End If

			mvar_lngAbsenceDurationID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEDURATION))
			If mvar_lngAbsenceDurationID > 0 Then
				gsAbsenceDurationColumnName = _columns.GetById(mvar_lngAbsenceDurationID).Name
			Else
				gsAbsenceDurationColumnName = ""
			End If

			mvar_lngAbsenceTypeTypeID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPETYPE))
			If mvar_lngAbsenceTypeTypeID > 0 Then
				gsAbsenceTypeTypeColumnName = _columns.GetById(mvar_lngAbsenceTypeTypeID).Name
			Else
				gsAbsenceTypeTypeColumnName = ""
			End If

			mvar_lngAbsenceTypeCodeID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPECODE))
			If mvar_lngAbsenceTypeCodeID > 0 Then
				gsAbsenceTypeCodeColumnName = _columns.GetById(mvar_lngAbsenceTypeCodeID).Name
			Else
				gsAbsenceTypeCodeColumnName = ""
			End If

			mvar_lngAbsenceTypeSSPID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPESSP))
			If mvar_lngAbsenceTypeSSPID > 0 Then
				gsAbsenceTypeSSPColumnName = _columns.GetById(mvar_lngAbsenceTypeSSPID).Name
			Else
				gsAbsenceTypeSSPColumnName = ""
			End If

			mvar_lngAbsenceTypeCalCodeID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPECALCODE))
			If mvar_lngAbsenceTypeCalCodeID > 0 Then
				gsAbsenceTypeCalCodeColumnName = _columns.GetById(mvar_lngAbsenceTypeCalCodeID).Name
			Else
				gsAbsenceTypeCalCodeColumnName = ""
			End If

			mvar_lngAbsenceTypeIncludeID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPEINCLUDE))
			If mvar_lngAbsenceTypeIncludeID > 0 Then
				gsAbsenceTypeIncludeColumnName = _columns.GetById(mvar_lngAbsenceTypeIncludeID).Name
			Else
				gsAbsenceTypeIncludeColumnName = ""
			End If

			mvar_lngAbsenceTypeBradfordIndexID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPEBRADFORDINDEX))
			If mvar_lngAbsenceTypeBradfordIndexID > 0 Then
				gsAbsenceTypeBradfordIndexColumnName = _columns.GetById(mvar_lngAbsenceTypeBradfordIndexID).Name
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
			Dim rsType As DataTable

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
					Dim objDataAccess = New clsDataAccess(_objLogin)
					rsType = objDataAccess.GetDataTable("SELECT * FROM " & gsAbsenceTypeTableName & " ORDER BY " & gsAbsenceTypeTypeColumnName)
					If rsType.Rows.Count = 0 Then
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

	End Class

End Namespace