Option Strict Off
Option Explicit On

Imports System.Collections.Generic
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.Structures
Imports HR.Intranet.Server.Metadata

Namespace BaseClasses

	Public MustInherit Class BaseForDMI

		Protected RegionalSettings As RegionalSettings
		Friend AbsenceModule As modAbsenceSpecifics
		Friend BankHolidayModule As modBankHolidaySpecifics
		Friend PersonnelModule As modPersonnelSpecifics

		Friend Tables As ICollection(Of Table)
		Friend Columns As ICollection(Of Column)
		Friend Relations As List(Of Relation)

		Private _moduleSettings As ICollection(Of ModuleSetting)
		Private _userSettings As ICollection(Of UserSetting)
		Public SystemSettings As IList(Of UserSetting)

		Public DB As clsDataAccess
		Protected General As clsGeneral
		Protected Logs As clsEventLog
		Protected AccessLog As AccessLog

		Private _sessionInfo As SessionInfo
		Protected _login As LoginInfo

		Public Property SessionInfo() As SessionInfo
			Set(value As SessionInfo)
				_sessionInfo = value
				_login = _sessionInfo.LoginInfo

				DB = New clsDataAccess(_sessionInfo.LoginInfo)
				General = New clsGeneral(_sessionInfo.LoginInfo)
				Logs = New clsEventLog(_sessionInfo.LoginInfo)
				AccessLog = New AccessLog(_sessionInfo.LoginInfo)

				RegionalSettings = value.RegionalSettings
				AbsenceModule = value.AbsenceModule
				BankHolidayModule = value.BankHolidayModule
				PersonnelModule = value.PersonnelModule

				Tables = value.Tables
				Columns = value.Columns
				Relations = value.Relations

				_moduleSettings = value.ModuleSettings
				_userSettings = value.UserSettings
				SystemSettings = value.SystemSettings

			End Set
			Get
				Return _sessionInfo
			End Get
		End Property

		Friend Function NewExpression() As clsExprExpression
			Return New clsExprExpression(_sessionInfo)
		End Function

#Region "FROM Declarations"

		Public ReadOnly Property gcoTablePrivileges As ICollection(Of TablePrivilege)
			<DebuggerStepThrough()> _
			Get
				Return _sessionInfo.gcoTablePrivileges
			End Get
		End Property

		Public ReadOnly Property gcolColumnPrivilegesCollection As Collection
			<DebuggerStepThrough()> _
				 Get
				Return _sessionInfo.gcolColumnPrivilegesCollection
			End Get
		End Property

#End Region

#Region "FROM modExpression"

		Protected Function IsFilterValid(varID As Integer) As String

			' Since validation occurs whenn saving expression this function should be unnecessary
			Return ""

			'		Dim objExpr As clsExprExpression
			'		Dim strRuntimeCode As String
			'		Dim strFilterName As String
			'		Dim avDummyPrompts(,) As Object

			'		On Error GoTo LocalErr

			'		ReDim avDummyPrompts(1, 0)

			'		strFilterName = vbNullString
			'		IsFilterValid = IsSelectionValid(varID, "filter")

			'		If IsFilterValid = vbNullString Then
			'			objExpr = New clsExprExpression()
			'			With objExpr
			'				'JPD 20030324 Fault 5161
			'				'UPGRADE_WARNING: Couldn't resolve default property of object varID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'				.ExpressionID = CInt(varID)
			'				.ConstructExpression()
			'				If (.ValidateExpression(True) <> ExprValidationCodes.giEXPRVALIDATION_NOERRORS) Then
			'					IsFilterValid = "The filter '" & strFilterName & "' used in this definition is invalid."
			'				End If

			'			End With
			'			'UPGRADE_NOTE: Object objExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			'			objExpr = Nothing
			'		End If

			'		Exit Function

			'LocalErr:
			'		If strFilterName <> vbNullString Then
			'			IsFilterValid = "'" & strFilterName & "' "
			'		End If
			'		IsFilterValid = "The filter " & IsFilterValid & "used in this definition is invalid"

		End Function

		Protected Function IsPicklistValid(varID As Integer) As String
			Return IsSelectionValid(varID, "picklist")
		End Function

		Protected Function IsCalcValid(varID As Integer) As String
			Return IsSelectionValid(varID, "calculation")
		End Function

		Protected Function IsSelectionValid(varID As Integer, strType As String) As String

			Dim rsTemp As DataTable

			IsSelectionValid = vbNullString
			'UPGRADE_WARNING: Couldn't resolve default property of object varID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If varID = 0 Then Exit Function

			rsTemp = GetSelectionAccess(varID, strType)

			If rsTemp.Rows.Count = 0 Then
				IsSelectionValid = "The " & strType & " used in this definition has been deleted by another user"

			ElseIf LCase(Trim(rsTemp.Rows(0)("Username").ToString())) <> LCase(Trim(_login.Username)) And rsTemp.Rows(0)("Access").ToString() = "HD" Then
				'JPD 20040706 Fault 8781
				If Not CurrentUserIsSysSecMgr() Then
					IsSelectionValid = "The " & strType & " used in this definition has been " & "hidden by another user"
				End If
			End If
			'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsTemp = Nothing

		End Function

		Private Function GetSelectionAccess(ID As Integer, strType As String) As DataTable

			Dim strSQL As String

			If strType = "picklist" Then
				strSQL = String.Format("SELECT Access, UserName FROM AsrSysPicklistName WHERE PickListID = {0}", ID)
			Else
				strSQL = String.Format("SELECT Access, UserName FROM AsrSysExpressions WHERE ExprID = {0}", ID)
			End If
			Return DB.GetDataTable(strSQL)

		End Function


#End Region

#Region "From modUtilityAccess"

		Protected Function CurrentUserIsSysSecMgr() As Boolean
			Return _login.IsSystemOrSecurityAdmin
		End Function



		Private Function ValidatePicklist(plngID As Integer) As RecordSelectionValidityCodes
			' Return an integer code representing the validity of the picklist.
			' Return 0 if the picklist is OK.
			' Return 1 if the picklist has been deleted by another user.
			' Return 2 if the picklist is hidden, and is owned by the current user.
			' Return 3 if the picklist is hidden, and is NOT owned by the current user.
			' Return 4 if the picklist is no longer valid.
			Dim iResult As RecordSelectionValidityCodes
			Dim rstemp As DataTable
			Dim sSQL As String

			Try

				iResult = RecordSelectionValidityCodes.REC_SEL_VALID_OK

				If plngID > 0 Then

					sSQL = "SELECT access, userName FROM ASRSysPickListName WHERE picklistID = " & CStr(plngID)
					rstemp = DB.GetDataTable(sSQL)

					If rstemp.Rows.Count = 0 Then
						' Picklist no longer exists
						iResult = RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
					Else
						If (rstemp.Rows(0)("Access").ToString() = ACCESS_HIDDEN) Then
							If (LCase(Trim(rstemp.Rows(0)("Username").ToString())) = LCase(Trim(_login.Username))) Then
								' Picklist is hidden by the current user.
								iResult = RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYUSER
							Else
								' Picklist is hidden by another user.
								iResult = RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
							End If
						End If
					End If

				End If

				Return iResult

			Catch ex As Exception
				Return RecordSelectionValidityCodes.REC_SEL_VALID_INVALID

			End Try

		End Function


		Private Function ValidateFilter(plngID As Integer) As RecordSelectionValidityCodes
			' Return an integer code representing the validity of the filter.
			' Return 0 if the filter is OK.
			' Return 1 if the filter has been deleted by another user.
			' Return 2 if the filter is hidden, and is owned by the current user.
			' Return 3 if the filter is hidden, and is NOT owned by the current user.
			' Return 4 if the filter is no longer valid.
			On Error GoTo ErrorTrap

			Dim iResult As RecordSelectionValidityCodes
			Dim rstemp As DataTable
			Dim sSQL As String

			iResult = RecordSelectionValidityCodes.REC_SEL_VALID_OK

			If plngID > 0 Then

				sSQL = "SELECT access, userName FROM ASRSysExpressions WHERE exprID = " & CStr(plngID)
				rstemp = DB.GetDataTable(sSQL)

				If rstemp.Rows.Count = 0 Then
					' Filter no longer exists
					iResult = RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
				Else
					If (rstemp.Rows(0)("Access").ToString() = ACCESS_HIDDEN) Or General.HasHiddenComponents(CInt(plngID)) Then
						If (LCase(Trim(rstemp.Rows(0)("Username").ToString())) = LCase(Trim(_login.Username))) Then
							' Filter is hidden by the current user.
							iResult = RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYUSER
						Else
							' Filter is hidden by another user.
							iResult = RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
						End If
					End If
				End If

			End If

TidyUpAndExit:
			ValidateFilter = iResult
			Exit Function

ErrorTrap:
			iResult = RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
			Resume TidyUpAndExit

		End Function

		Protected Function ValidateRecordSelection(piType As RecordSelectionType, plngID As Integer) As RecordSelectionValidityCodes
			' Return an integer code representing the validity of the record selection (picklist or filter).
			' Return 0 if the record selection is OK.
			' Return 1 if the record selection has been deleted by another user.
			' Return 2 if the record selection is hidden, and is owned by the current user.
			' Return 3 if the record selection is hidden, and is NOT owned by the current user.
			' Return 4 if the record selection is no longer valid.
			On Error GoTo ErrorTrap

			Dim iResult As RecordSelectionValidityCodes

			iResult = RecordSelectionValidityCodes.REC_SEL_VALID_OK

			Select Case piType

				Case RecordSelectionType.Picklist
					iResult = ValidatePicklist(plngID)

				Case RecordSelectionType.Filter
					iResult = ValidateFilter(plngID)
			End Select

TidyUpAndExit:
			ValidateRecordSelection = iResult
			Exit Function

ErrorTrap:
			iResult = RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
			Resume TidyUpAndExit

		End Function

		Protected Function ValidateCalculation(ByVal plngID As Integer) As RecordSelectionValidityCodes
			' Return an integer code representing the validity of the Calculation.
			' Return 0 if the Calculation is OK.
			' Return 1 if the Calculation has been deleted by another user.
			' Return 2 if the Calculation is hidden, and is owned by the current user.
			' Return 3 if the Calculation is hidden, and is NOT owned by the current user.
			' Return 4 if the Calculation is no longer valid.
			On Error GoTo ErrorTrap

			Dim iResult As RecordSelectionValidityCodes
			Dim rstemp As DataTable
			Dim sSQL As String

			iResult = RecordSelectionValidityCodes.REC_SEL_VALID_OK

			If plngID > 0 Then

				sSQL = "SELECT access, userName FROM ASRSysExpressions WHERE exprID = " & CStr(plngID)
				rstemp = DB.GetDataTable(sSQL)

				If rstemp.Rows.Count = 0 Then
					' Filter no longer exists
					iResult = RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
				Else
					If (rstemp.Rows(0)("Access").ToString() = ACCESS_HIDDEN) Or General.HasHiddenComponents(CInt(plngID)) Then
						If (LCase(Trim(rstemp.Rows(0)("Username").ToString())) = LCase(Trim(_login.Username))) Then
							' Calculation is hidden by the current user.
							iResult = RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYUSER
						Else
							' Calculation is hidden by another user.
							iResult = RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
						End If
					End If
				End If


				'UPGRADE_NOTE: Object rstemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				rstemp = Nothing

			End If

TidyUpAndExit:
			Return iResult

ErrorTrap:
			iResult = RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
			Resume TidyUpAndExit

		End Function



#End Region

#Region "From clsGeneral"

		Public Function GetModuleParameter(psModuleKey As String, psParameterKey As String) As String
			Return _moduleSettings.GetSetting(psModuleKey, psParameterKey).ParameterValue
		End Function

		Protected Function FilteredIDs(plngExprID As Integer, ByRef psIDSQL As String, ByRef psUDFs() As String, Optional paPrompts As Object = Nothing) As Boolean
			' Return a string describing the record IDs from the given table
			' that satisfy the given criteria.
			Dim fOK As Boolean
			Dim objExpr As clsExprExpression = New clsExprExpression(_sessionInfo)

			With objExpr
				' Initialise the filter expression object.
				fOK = .Initialise(0, plngExprID, ExpressionTypes.giEXPR_RUNTIMEFILTER, ExpressionValueTypes.giEXPRVALUE_LOGIC)

				If fOK Then
					fOK = objExpr.RuntimeFilterCode(psIDSQL, True, psUDFs, False, paPrompts)
				End If

			End With

			Return fOK

		End Function

		Public Function GetDataType(lTableID As Integer, lngColumnID As Integer) As ColumnDataType
			Return Columns.GetById(lngColumnID).DataType
		End Function

		Public Function GetColumnTable(plngColumnID As Integer) As Integer
			Return Columns.GetById(plngColumnID).TableID
		End Function

		Public Function GetColumnName(plngColumnID As Integer) As String
			If plngColumnID = 0 Then
				Return ""
			Else
				Return Columns.GetById(plngColumnID).Name
			End If
		End Function

		Friend Function GetColumnDataType(plngColumnID As Integer) As ColumnDataType
			Return Columns.GetById(plngColumnID).DataType
		End Function

		Public Function IsPhotoDataType(lngColumnID As Integer) As Boolean
			Return Columns.GetById(lngColumnID).DataType = ColumnDataType.sqlVarBinary
		End Function

		Public Function GetColumnTableName(plngColumnID As Integer) As String
			Return Columns.GetById(plngColumnID).TableName
		End Function

		Public Function IsAChildOf(lTestTableID As Integer, lBaseTableID As Integer) As Boolean
			Return Relations.IsRelation(lBaseTableID, lTestTableID)
		End Function

		Public Function IsAParentOf(lTestTableID As Integer, lBaseTableID As Integer) As Boolean
			Return Relations.IsRelation(lTestTableID, lBaseTableID)
		End Function

#End Region

#Region "Accessible from dmi.net - may need to move to a metadata class at a future date"

		Public Function GetTableName(TableID As Integer) As String
			Return Tables.GetById(TableID).Name
		End Function

		Public Function GetTableFromColumnID(ColumnID As Integer) As Table
			Dim objColumn = Columns.GetById(ColumnID)
			Return Tables.GetById(objColumn.TableID)
		End Function

#End Region

#Region "FROM modPermissions"

		Friend Function GetColumnPrivileges(psTableViewName As String) As CColumnPrivileges

			Dim iLoop As Integer
			Dim objColumnPrivileges As CColumnPrivileges

			Try
				' If the given table/view's column privilege collection has already been read then simply return it.
				For iLoop = 1 To gcolColumnPrivilegesCollection.Count()
					If UCase(gcolColumnPrivilegesCollection.Item(iLoop).Tag) = UCase(psTableViewName) Then
						Return gcolColumnPrivilegesCollection.Item(iLoop)
					End If
				Next iLoop

				Return objColumnPrivileges

			Catch ex As Exception
				Return Nothing
			End Try

		End Function

#End Region

#Region "FROM clsSettings"

		Public Function GetUserSetting(strSection As String, strKey As String, varDefault As Object) As Object

			Dim objSetting = _userSettings.GetUserSetting(strSection, strKey)
			If objSetting Is Nothing Then
				Return varDefault
			Else
				Return objSetting.Value
			End If

		End Function


#End Region

#Region "From modUtilityAccess"

		Protected Function CurrentUserAccess(piUtilityType As UtilityType, plngID As Integer) As String

			' Return the access code (RW/RO/HD) of the current user's access
			' on the given utility.
			Dim sAccessCode As String
			Dim sSQL As String
			Dim sDefaultAccess As String
			Dim rsAccess As DataTable
			Dim sTableName As String
			Dim sAccessTableName As String
			Dim sIDColumnName As String

			sTableName = ""
			sAccessTableName = ""

			If plngID > 0 Then
				sDefaultAccess = ACCESS_HIDDEN
			Else
				sDefaultAccess = ACCESS_HIDDEN
			End If

			' Construct the SQL code to get the current user's access settings for the given utility.
			' NB. System and Security Manager users automatically have Read/Write access.
			Select Case piUtilityType
				Case UtilityType.utlBatchJob
					sTableName = "ASRSysBatchJobName"
					sAccessTableName = "ASRSysBatchJobAccess"
					sIDColumnName = "ID"

				Case UtilityType.utlCalendarReport
					sTableName = "ASRSysCalendarReports"
					sAccessTableName = "ASRSysCalendarReportAccess"
					sIDColumnName = "ID"

				Case UtilityType.utlCrossTab
					sTableName = "ASRSysCrossTab"
					sAccessTableName = "ASRSysCrossTabAccess"
					sIDColumnName = "CrossTabID"

				Case UtilityType.utlCustomReport
					sTableName = "ASRSysCustomReportsName"
					sAccessTableName = "ASRSysCustomReportAccess"
					sIDColumnName = "ID"

				Case UtilityType.utlDataTransfer
					sTableName = "ASRSysDataTransferName"
					sAccessTableName = "ASRSysDataTransferAccess"
					sIDColumnName = "DataTransferID"

				Case UtilityType.utlExport
					sTableName = "ASRSysExportName"
					sAccessTableName = "ASRSysExportAccess"
					sIDColumnName = "ID"

				Case UtilityType.UtlGlobalAdd, UtilityType.utlGlobalDelete, UtilityType.utlGlobalUpdate
					sTableName = "ASRSysGlobalFunctions"
					sAccessTableName = "ASRSysGlobalAccess"
					sIDColumnName = "functionID"

				Case UtilityType.utlImport
					sTableName = "ASRSysImportName"
					sAccessTableName = "ASRSysImportAccess"
					sIDColumnName = "ID"

				Case UtilityType.utlLabel, UtilityType.utlMailMerge
					sTableName = "ASRSysMailMergeName"
					sAccessTableName = "ASRSysMailMergeAccess"
					sIDColumnName = "mailMergeID"

				Case UtilityType.utlRecordProfile
					sTableName = "ASRSysRecordProfileName"
					sAccessTableName = "ASRSysRecordProfileAccess"
					sIDColumnName = "recordProfileID"

				Case UtilityType.utlMatchReport, UtilityType.utlSuccession, UtilityType.utlCareer
					sTableName = "ASRSysMatchReportName"
					sAccessTableName = "ASRSysMatchReportAccess"
					sIDColumnName = "matchReportID"

			End Select

			Try

				If Len(sAccessTableName) > 0 Then
					sSQL = "SELECT  CASE WHEN (SELECT count(*) FROM ASRSysGroupPermissions INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & "        AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & "        OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & "      INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & "        AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & "      WHERE b.Name = ASRSysGroupPermissions.groupname" & "        AND ASRSysGroupPermissions.permitted = 1) > 0 THEN '" & ACCESS_READWRITE & "'" & "    WHEN " & sTableName & ".userName = system_user THEN '" & ACCESS_READWRITE & "'" & "    ELSE" & "      CASE" & "        WHEN " & sAccessTableName & ".access IS null THEN '" & sDefaultAccess & "'" & "        ELSE " & sAccessTableName & ".access" & "      END" & "  END AS Access" & " FROM sysusers b" & " INNER JOIN sysusers a ON b.uid = a.gid" & " LEFT OUTER JOIN " & sAccessTableName & " ON (b.name = " & sAccessTableName & ".groupName" & "   AND " & sAccessTableName & ".id = " & CStr(plngID) & ")" & " INNER JOIN " & sTableName & " ON " & sAccessTableName & ".ID = " & sTableName & "." & sIDColumnName & " WHERE b.name = '" & _login.UserGroup & "'"


					rsAccess = DB.GetDataTable(sSQL)
					With rsAccess
						If .Rows.Count = 0 Then
							sAccessCode = sDefaultAccess
						Else
							sAccessCode = .Rows(0)("Access").ToString()
						End If

					End With
				Else
					sAccessCode = ACCESS_UNKNOWN
				End If

				Return sAccessCode

			Catch ex As Exception
				Return sDefaultAccess

			End Try

		End Function


#End Region

		Public ReadOnly Property LocaleDateFormat As String
			Get
				If Globalization.CultureInfo.CurrentUICulture.ToString() = "en-US" Or Globalization.CultureInfo.CurrentUICulture.ToString() = "en" Then
					' Force 2-digit days and months
					Return "MM/dd/yyyy"
				Else
					Return RegionalSettings.DateFormat.ShortDatePattern
				End If


			End Get
		End Property

		Protected ReadOnly Property SQLDateFormat As String
			Get
				Return "MM/dd/yyyy"
			End Get
		End Property


	End Class
End Namespace