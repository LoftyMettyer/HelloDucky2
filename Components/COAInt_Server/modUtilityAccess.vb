Option Strict Off
Option Explicit On
Module modUtilityAccess
	
	Public Enum RecordSelectionTypes
		REC_SEL_ALL = 0
		REC_SEL_PICKLIST = 1
		REC_SEL_FILTER = 2
	End Enum
	
	Public Const ACCESS_READWRITE As String = "RW"
	Public Const ACCESS_READONLY As String = "RO"
	Public Const ACCESS_HIDDEN As String = "HD"
	Public Const ACCESS_UNKNOWN As String = ""
	
	Public Const ACCESSDESC_READWRITE As String = "Read / Write"
	Public Const ACCESSDESC_READONLY As String = "Read Only"
	Public Const ACCESSDESC_HIDDEN As String = "Hidden"
	Public Const ACCESSDESC_UNKNOWN As String = "Unknown"
	
	Public Enum RecordSelectionValidityCodes
		REC_SEL_VALID_OK = 0
		REC_SEL_VALID_DELETED = 1
		REC_SEL_VALID_HIDDENBYUSER = 2
		REC_SEL_VALID_HIDDENBYOTHER = 3
		REC_SEL_VALID_INVALID = 4
	End Enum
	
	Public Function ValidateRecordSelection(ByRef piType As RecordSelectionTypes, ByRef plngID As Integer) As RecordSelectionValidityCodes
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
			Case RecordSelectionTypes.REC_SEL_PICKLIST
				iResult = ValidatePicklist(plngID)
				
			Case RecordSelectionTypes.REC_SEL_FILTER
				iResult = ValidateFilter(plngID)
		End Select
		
TidyUpAndExit: 
		ValidateRecordSelection = iResult
		Exit Function
		
ErrorTrap: 
		iResult = RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
		Resume TidyUpAndExit
		
	End Function
	
	
	Public Function ValidatePicklist(ByRef plngID As Integer) As RecordSelectionValidityCodes
		' Return an integer code representing the validity of the picklist.
		' Return 0 if the picklist is OK.
		' Return 1 if the picklist has been deleted by another user.
		' Return 2 if the picklist is hidden, and is owned by the current user.
		' Return 3 if the picklist is hidden, and is NOT owned by the current user.
		' Return 4 if the picklist is no longer valid.
		On Error GoTo ErrorTrap
		
		Dim iResult As RecordSelectionValidityCodes
		Dim rstemp As ADODB.Recordset
		Dim sSQL As String
		Dim datData As clsDataAccess
		
		sSQL = ""
		iResult = RecordSelectionValidityCodes.REC_SEL_VALID_OK
		
		If plngID > 0 Then
			datData = New clsDataAccess
			
			sSQL = "SELECT access, userName" & " FROM ASRSysPickListName" & " WHERE picklistID = " & CStr(plngID)
			
			rstemp = datData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
			
			If rstemp.BOF And rstemp.EOF Then
				' Picklist no longer exists
				iResult = RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
			Else
				If (rstemp.Fields("Access").Value = ACCESS_HIDDEN) Then
					If (LCase(Trim(rstemp.Fields("Username").Value)) = LCase(Trim(gsUsername))) Then
						' Picklist is hidden by the current user.
						iResult = RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYUSER
					Else
						' Picklist is hidden by another user.
						iResult = RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
					End If
				End If
			End If
			
			rstemp.Close()
			'UPGRADE_NOTE: Object rstemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rstemp = Nothing
			
			'UPGRADE_NOTE: Object datData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			datData = Nothing
		End If
		
TidyUpAndExit: 
		ValidatePicklist = iResult
		Exit Function
		
ErrorTrap: 
		iResult = RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
		Resume TidyUpAndExit
		
	End Function
	
	
	
	
	Public Function ValidateFilter(ByRef plngID As Integer) As RecordSelectionValidityCodes
		' Return an integer code representing the validity of the filter.
		' Return 0 if the filter is OK.
		' Return 1 if the filter has been deleted by another user.
		' Return 2 if the filter is hidden, and is owned by the current user.
		' Return 3 if the filter is hidden, and is NOT owned by the current user.
		' Return 4 if the filter is no longer valid.
		On Error GoTo ErrorTrap
		
		Dim iResult As RecordSelectionValidityCodes
		Dim rstemp As ADODB.Recordset
		Dim sSQL As String
		Dim objExpr As clsExprExpression
		Dim datData As clsDataAccess
		
		sSQL = ""
		iResult = RecordSelectionValidityCodes.REC_SEL_VALID_OK
		
		If plngID > 0 Then
			datData = New clsDataAccess
			
			sSQL = "SELECT access, userName" & " FROM ASRSysExpressions" & " WHERE exprID = " & CStr(plngID)
			
			rstemp = datData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
			
			If rstemp.BOF And rstemp.EOF Then
				' Filter no longer exists
				iResult = RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
			Else
				If (rstemp.Fields("Access").Value = ACCESS_HIDDEN) Or HasHiddenComponents(CInt(plngID)) Then
					If (LCase(Trim(rstemp.Fields("Username").Value)) = LCase(Trim(gsUsername))) Then
						' Filter is hidden by the current user.
						iResult = RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYUSER
					Else
						' Filter is hidden by another user.
						iResult = RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
					End If
				Else
					'JPD 20040804 This function is only called when validating filters
					' used in reports. It takes a long time to do all of its validation checks,
					' and in truth the filter should not have become invalid (if somebody has changed it to be invalid
					' the full check will be made when they try to save the filter). So I don't want
					' to do this full check now.
					'        Set objExpr = New clsExprExpression
					'        With objExpr
					'          .ExpressionID = CLng(plngID)
					'          .ConstructExpression
					'          If (.ValidateExpression(True) <> giEXPRVALIDATION_NOERRORS) Then
					'            iResult = REC_SEL_VALID_INVALID
					'          End If
					'        End With
					'        Set objExpr = Nothing
				End If
			End If
			
			rstemp.Close()
			'UPGRADE_NOTE: Object rstemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rstemp = Nothing
			
			'UPGRADE_NOTE: Object datData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			datData = Nothing
		End If
		
TidyUpAndExit: 
		ValidateFilter = iResult
		Exit Function
		
ErrorTrap: 
		iResult = RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
		Resume TidyUpAndExit
		
	End Function
	
	
	Public Function CurrentUserIsSysSecMgr() As Boolean
		Dim sSQL As String
		Dim rsAccess As ADODB.Recordset
		Dim datData As clsDataAccess
		Dim fIsSysSecUser As Boolean
		
		'sSQL = "SELECT count(*) AS [result]" & _
		'" FROM ASRSysGroupPermissions" & _
		'" INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
		'"   AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
		'"   OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
		'" INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
		'"   AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
		'" INNER JOIN sysusers b ON b.name = ASRSysGroupPermissions.groupname" & _
		'" INNER JOIN sysusers a ON b.uid = a.gid" & _
		'"   AND a.Name = current_user" & _
		'" WHERE ASRSysGroupPermissions.permitted = 1"
		sSQL = "SELECT count(*) AS [result]" & " FROM ASRSysGroupPermissions" & " INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & "   AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & "   OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & " INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & "   AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & " WHERE ASRSysGroupPermissions.permitted = 1" & "   AND ASRSysGroupPermissions.groupname = '" & gsUserGroup & "'"
		
		datData = New clsDataAccess
		rsAccess = datData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		With rsAccess
			fIsSysSecUser = (.Fields("Result").Value > 0)
			
			.Close()
		End With
		'UPGRADE_NOTE: Object rsAccess may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsAccess = Nothing
		
		'UPGRADE_NOTE: Object datData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		datData = Nothing
		
		CurrentUserIsSysSecMgr = fIsSysSecUser
		
	End Function
	
	
	
	Public Function ValidateCalculation(ByRef plngID As Integer) As RecordSelectionValidityCodes
		' Return an integer code representing the validity of the Calculation.
		' Return 0 if the Calculation is OK.
		' Return 1 if the Calculation has been deleted by another user.
		' Return 2 if the Calculation is hidden, and is owned by the current user.
		' Return 3 if the Calculation is hidden, and is NOT owned by the current user.
		' Return 4 if the Calculation is no longer valid.
		On Error GoTo ErrorTrap
		
		Dim iResult As RecordSelectionValidityCodes
		Dim rstemp As ADODB.Recordset
		Dim sSQL As String
		Dim objExpr As clsExprExpression
		Dim datData As clsDataAccess
		
		sSQL = ""
		iResult = RecordSelectionValidityCodes.REC_SEL_VALID_OK
		
		If plngID > 0 Then
			datData = New clsDataAccess
			
			sSQL = "SELECT access, userName" & " FROM ASRSysExpressions" & " WHERE exprID = " & CStr(plngID)
			
			rstemp = datData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
			
			If rstemp.BOF And rstemp.EOF Then
				' Filter no longer exists
				iResult = RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
			Else
				If (rstemp.Fields("Access").Value = ACCESS_HIDDEN) Or HasHiddenComponents(CInt(plngID)) Then
					If (LCase(Trim(rstemp.Fields("Username").Value)) = LCase(Trim(gsUsername))) Then
						' Calculation is hidden by the current user.
						iResult = RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYUSER
					Else
						' Calculation is hidden by another user.
						iResult = RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
					End If
				Else
					'JPD 20040804 This function is only called when validating calculations
					' used in reports. It takes a long time to do all of its validation checks,
					' and in truth the calc should not have become invalid (if somebody has changed it to be invalid
					' the full check will be made when they try to save the calc.). So I don't want
					' to do this full check now.
					'        Set objExpr = New clsExprExpression
					'        With objExpr
					'          .ExpressionID = CLng(plngID)
					'          .ConstructExpression
					'          If (.ValidateExpression(True) <> giEXPRVALIDATION_NOERRORS) Then
					'            iResult = REC_SEL_VALID_INVALID
					'          End If
					'        End With
					'        Set objExpr = Nothing
				End If
			End If
			
			rstemp.Close()
			'UPGRADE_NOTE: Object rstemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rstemp = Nothing
			
			'UPGRADE_NOTE: Object datData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			datData = Nothing
		End If
		
TidyUpAndExit: 
		ValidateCalculation = iResult
		Exit Function
		
ErrorTrap: 
		iResult = RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
		Resume TidyUpAndExit
		
	End Function
	
	
	
	Public Function AccessCode(ByRef psDescription As String) As String
		' Return the descriptive string associated with the given Access code.
		Select Case psDescription
			Case ACCESSDESC_READWRITE
				AccessCode = ACCESS_READWRITE
			Case ACCESSDESC_READONLY
				AccessCode = ACCESS_READONLY
			Case ACCESSDESC_HIDDEN
				AccessCode = ACCESS_HIDDEN
			Case Else
				AccessCode = ACCESS_UNKNOWN
		End Select
		
	End Function
	
	Public Function AccessDescription(ByRef psCode As String) As String
		' Return the descriptive string associated with the given Access code.
		Select Case psCode
			Case ACCESS_READWRITE
				AccessDescription = ACCESSDESC_READWRITE
			Case ACCESS_READONLY
				AccessDescription = ACCESSDESC_READONLY
			Case ACCESS_HIDDEN
				AccessDescription = ACCESSDESC_HIDDEN
			Case Else
				AccessDescription = ACCESSDESC_UNKNOWN
		End Select
		
	End Function
	
	
	Public Function CurrentUserAccess(ByRef piUtilityType As modUtilAccessLog.UtilityType, ByRef plngID As Integer) As String
		' Return the access code (RW/RO/HD) of the current user's access
		' on the given utility.
		On Error GoTo ErrorTrap
		
		Dim sAccessCode As String
		Dim sSQL As String
		Dim sDefaultAccess As String
		Dim rsAccess As ADODB.Recordset
		Dim datData As clsDataAccess
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
			Case modUtilAccessLog.UtilityType.utlBatchJob
				sTableName = "ASRSysBatchJobName"
				sAccessTableName = "ASRSysBatchJobAccess"
				sIDColumnName = "ID"
				
			Case modUtilAccessLog.UtilityType.utlCalendarReport
				sTableName = "ASRSysCalendarReports"
				sAccessTableName = "ASRSysCalendarReportAccess"
				sIDColumnName = "ID"
				
			Case modUtilAccessLog.UtilityType.utlCrossTab
				sTableName = "ASRSysCrossTab"
				sAccessTableName = "ASRSysCrossTabAccess"
				sIDColumnName = "CrossTabID"
				
			Case modUtilAccessLog.UtilityType.utlCustomReport
				sTableName = "ASRSysCustomReportsName"
				sAccessTableName = "ASRSysCustomReportAccess"
				sIDColumnName = "ID"
				
			Case modUtilAccessLog.UtilityType.utlDataTransfer
				sTableName = "ASRSysDataTransferName"
				sAccessTableName = "ASRSysDataTransferAccess"
				sIDColumnName = "DataTransferID"
				
			Case modUtilAccessLog.UtilityType.utlExport
				sTableName = "ASRSysExportName"
				sAccessTableName = "ASRSysExportAccess"
				sIDColumnName = "ID"
				
			Case modUtilAccessLog.UtilityType.UtlGlobalAdd, modUtilAccessLog.UtilityType.utlGlobalDelete, modUtilAccessLog.UtilityType.utlGlobalUpdate
				sTableName = "ASRSysGlobalFunctions"
				sAccessTableName = "ASRSysGlobalAccess"
				sIDColumnName = "functionID"
				
			Case modUtilAccessLog.UtilityType.utlImport
				sTableName = "ASRSysImportName"
				sAccessTableName = "ASRSysImportAccess"
				sIDColumnName = "ID"
				
			Case modUtilAccessLog.UtilityType.utlLabel, modUtilAccessLog.UtilityType.utlMailMerge
				sTableName = "ASRSysMailMergeName"
				sAccessTableName = "ASRSysMailMergeAccess"
				sIDColumnName = "mailMergeID"
				
			Case modUtilAccessLog.UtilityType.utlRecordProfile
				sTableName = "ASRSysRecordProfileName"
				sAccessTableName = "ASRSysRecordProfileAccess"
				sIDColumnName = "recordProfileID"
				
			Case modUtilAccessLog.UtilityType.utlMatchReport, modUtilAccessLog.UtilityType.utlSuccession, modUtilAccessLog.UtilityType.utlCareer
				sTableName = "ASRSysMatchReportName"
				sAccessTableName = "ASRSysMatchReportAccess"
				sIDColumnName = "matchReportID"
				
		End Select
		
		If Len(sAccessTableName) > 0 Then
			sSQL = "SELECT" & "  CASE" & "    WHEN (SELECT count(*)" & "      FROM ASRSysGroupPermissions" & "      INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & "        AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & "        OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & "      INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & "        AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & "      WHERE b.Name = ASRSysGroupPermissions.groupname" & "        AND ASRSysGroupPermissions.permitted = 1) > 0 THEN '" & ACCESS_READWRITE & "'" & "    WHEN " & sTableName & ".userName = system_user THEN '" & ACCESS_READWRITE & "'" & "    ELSE" & "      CASE" & "        WHEN " & sAccessTableName & ".access IS null THEN '" & sDefaultAccess & "'" & "        ELSE " & sAccessTableName & ".access" & "      END" & "  END AS Access" & " FROM sysusers b" & " INNER JOIN sysusers a ON b.uid = a.gid" & " LEFT OUTER JOIN " & sAccessTableName & " ON (b.name = " & sAccessTableName & ".groupName" & "   AND " & sAccessTableName & ".id = " & CStr(plngID) & ")" & " INNER JOIN " & sTableName & " ON " & sAccessTableName & ".ID = " & sTableName & "." & sIDColumnName & " WHERE b.name = '" & gsUserGroup & "'"
			
			'      " WHERE a.Name = current_user"
			
			datData = New clsDataAccess
			
			rsAccess = datData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
			With rsAccess
				If .BOF And .EOF Then
					sAccessCode = sDefaultAccess
				Else
					sAccessCode = .Fields("Access").Value
				End If
				
				.Close()
			End With
			'UPGRADE_NOTE: Object rsAccess may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsAccess = Nothing
			
			'UPGRADE_NOTE: Object datData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			datData = Nothing
		Else
			sAccessCode = ACCESS_UNKNOWN
		End If
		
TidyUpAndExit: 
		CurrentUserAccess = sAccessCode
		Exit Function
		
ErrorTrap: 
		Resume TidyUpAndExit
		
	End Function
End Module