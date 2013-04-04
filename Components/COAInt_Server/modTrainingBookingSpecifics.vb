Option Strict Off
Option Explicit On
Module modTrainingBookingSpecifics
	
	' Module parameters.
	Public gfTrainingBookingEnabled As Boolean
	
	' Module constants.
	Public Const gsMODULEKEY_TRAININGBOOKING As String = "MODULE_TRAININGBOOKING"
	Public Const gsPARAMETERKEY_TRAINBOOKTABLE As String = "Param_TrainBookTable"
	Public Const gsPARAMETERKEY_COURSESTARTDATE As String = "Param_CourseStartDate"
	Public Const gsPARAMETERKEY_COURSEENDDATE As String = "Param_CourseEndDate"
	
	' Training Booking Stuff
	Public glngTrainingBookingTableID As Integer
	Public gsTrainingBookingTableName As String
	
	Private mvar_lngTrainingBookingStartDateID As Integer
	Public gsTrainingBookingStartDateColumnName As String
	Private mvar_lngTrainingBookingEndDateID As Integer
	Public gsTrainingBookingEndDateColumnName As String
	
	
	Public Sub ReadTrainingBookingParameters()
		
		' Read the Training Booking module parameters from the database.
		glngTrainingBookingTableID = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_TRAINBOOKTABLE))
		If glngTrainingBookingTableID > 0 Then
			gsTrainingBookingTableName = datGeneral.GetTableName(glngTrainingBookingTableID)
		Else
			gsTrainingBookingTableName = ""
		End If
		
		mvar_lngTrainingBookingStartDateID = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSESTARTDATE))
		If mvar_lngTrainingBookingStartDateID > 0 Then
			gsTrainingBookingStartDateColumnName = datGeneral.GetColumnName(mvar_lngTrainingBookingStartDateID)
		Else
			gsTrainingBookingStartDateColumnName = ""
		End If
		
		mvar_lngTrainingBookingEndDateID = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSEENDDATE))
		If mvar_lngTrainingBookingEndDateID > 0 Then
			gsTrainingBookingEndDateColumnName = datGeneral.GetColumnName(mvar_lngTrainingBookingEndDateID)
		Else
			gsTrainingBookingEndDateColumnName = ""
		End If
		
	End Sub
	
	Public Function ValidateTrainingBookingParameters() As Boolean
		
		' Validate the configuration of the Training Booking module parameters,
		' and the current user's access on the configured columns.
		
		Dim fValid As Boolean
		Dim strMessage As String
		
		' -----------------------------------------------
		If gfTrainingBookingEnabled Then
			
			' Check the Training Booking Table ID is valid.
			If Not (glngTrainingBookingTableID > 0) Then
				strMessage = strMessage & "The Training Bookings table is not defined." & vbNewLine
			End If
			
			' Check the Start Date ID is valid.
			If Not (mvar_lngTrainingBookingStartDateID > 0) Then
				strMessage = strMessage & "The Course Start Date column is not defined." & vbNewLine
			End If
			
			' Check the End Date ID is valid.
			If Not (mvar_lngTrainingBookingEndDateID > 0) Then
				strMessage = strMessage & "The Course End Date column is not defined." & vbNewLine
			End If
			
		Else
			
			' Training Booking module is not enabled
			strMessage = "The Training Booking module is not enabled" & vbNewLine
			fValid = False
			
		End If
		
		' If an error found, warn the user.
		If Len(strMessage) > 0 Then
			strMessage = "The Training Booking module is not properly configured." & vbNewLine & vbNewLine & strMessage
			'NO MSGBOX ON THE SERVER ! - MsgBox strMessage, vbExclamation, App.ProductName
			fValid = False
		Else
			fValid = True
		End If
		
		' Return the validation value.
		ValidateTrainingBookingParameters = fValid
		
	End Function
	
	Public Function CheckPermission_TrainingBooking() As Boolean
		
		Dim pblnOK As Boolean
		Dim pstrBadColumn As String
		Dim objTable As CTablePrivilege
		Dim objColumn As CColumnPrivileges
		Dim pblnColumnOK As Boolean
		
		pblnOK = True
		
		' Retrieve the correct asrsyschildview for the Training Booking table
		objTable = gcoTablePrivileges.FindTableID(glngTrainingBookingTableID)
		
		If objTable.AllowSelect = False Then
			pblnOK = False
			pstrBadColumn = "Training Booking Table"
		End If
		
		gsTrainingBookingTableName = objTable.RealSource
		
		' Now check that read permission is available for the required columns
		objColumn = GetColumnPrivileges((objTable.TableName))
		
		' Check Course Start Date
		If pblnOK Then
			pblnOK = objColumn.IsValid(gsTrainingBookingStartDateColumnName)
			If pblnOK Then
				pblnOK = objColumn.Item(gsTrainingBookingStartDateColumnName).AllowSelect
				If pblnOK = False Then pstrBadColumn = "Course 'Start Date' column"
			Else
				pstrBadColumn = "Course 'Start Date' column"
			End If
		End If
		
		' Check Course End Date
		If pblnOK Then
			pblnOK = objColumn.IsValid(gsTrainingBookingEndDateColumnName)
			If pblnOK Then
				pblnOK = objColumn.Item(gsTrainingBookingEndDateColumnName).AllowSelect
				If pblnOK = False Then pstrBadColumn = "Course 'End Date' column"
			Else
				pstrBadColumn = "Course 'End Date' column"
			End If
		End If
		
		'UPGRADE_NOTE: Object objTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objTable = Nothing
		'UPGRADE_NOTE: Object objColumn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objColumn = Nothing
		
		CheckPermission_TrainingBooking = pblnOK
		
	End Function
End Module