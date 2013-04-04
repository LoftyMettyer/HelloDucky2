Option Strict Off
Option Explicit On
Module modUtilAccessLog
	
	Public Enum UtilityType
		utlBatchJob = 0
		utlCrossTab = 1
		utlCustomReport = 2
		utlDataTransfer = 3
		utlExport = 4
		UtlGlobalAdd = 5
		utlGlobalDelete = 6
		utlGlobalUpdate = 7
		utlImport = 8
		utlMailMerge = 9
		utlPicklist = 10
		utlFilter = 11
		utlCalculation = 12
		utlOrder = 13
		utlMatchReport = 14
		utlAbsenceBreakdown = 15
		utlBradfordFactor = 16
		utlCalendarReport = 17
		utlLabel = 18
		utlLabelType = 19
		utlRecordProfile = 20
		utlEmailAddress = 21
		utlEmailGroup = 22
		utlSuccession = 23
		utlCareer = 24
		utlWorkflow = 25
	End Enum
	
	Public Sub UtilCreated(ByRef utlType As UtilityType, ByRef lngID As Integer)
		
		Dim strSQL As String
		
		strSQL = "INSERT ASRSysUtilAccessLog " & "(Type, UtilID, CreatedBy, CreatedDate, CreatedHost, SavedBy, SavedDate, SavedHost) " & "VALUES (" & "'" & utlType & "', " & CStr(lngID) & ", " & " system_user, getdate(), host_name(), system_user, getdate(), host_name())"
		
		gADOCon.Execute(strSQL)
		
	End Sub
	
	
	Public Sub UtilUpdateLastSaved(ByRef utlType As UtilityType, ByRef lngID As Integer)
		Call UpdateUserAndDate("Saved", utlType, lngID)
	End Sub
	
	Public Sub UtilUpdateLastSavedMultiple(ByRef utlType As UtilityType, ByRef sIDs As String)
		
		Dim lngIDs As Object
		Dim intCount As Short
		
		If InStr(sIDs, ",") > 0 Then
			
			'UPGRADE_WARNING: Couldn't resolve default property of object lngIDs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			lngIDs = Split(sIDs, ",")
			
			For intCount = LBound(lngIDs) To UBound(lngIDs)
				
				'UPGRADE_WARNING: Couldn't resolve default property of object lngIDs(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Trim(lngIDs(intCount)) <> vbNullString Then
					
					'UPGRADE_WARNING: Couldn't resolve default property of object lngIDs(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Call UpdateUserAndDate("Saved", utlType, CInt(lngIDs(intCount)))
					
				End If
				
			Next 
			
		Else
			
			Call UpdateUserAndDate("Saved", utlType, CInt(sIDs))
			
		End If
		
	End Sub
	
	
	Public Sub UtilUpdateLastRun(ByRef utlType As UtilityType, ByRef lngID As Integer)
		Call UpdateUserAndDate("Run", utlType, lngID)
	End Sub
	
	
	Private Sub UpdateUserAndDate(ByRef strMode As String, ByRef utlType As UtilityType, ByRef lngID As Integer)
		
		Dim datData As clsDataAccess
		Dim rsTemp As ADODB.Recordset
		Dim strSQL As String
		Dim strHostName As String
		
		datData = New clsDataAccess
		
		strSQL = "SELECT * FROM ASRSysUtilAccessLog " & "WHERE UtilID = " & CStr(lngID) & " AND Type = " & CStr(utlType)
		rsTemp = datData.OpenRecordset(strSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		'Have to do this to catch existing utilities !
		If rsTemp.BOF And rsTemp.EOF Then
			strSQL = "INSERT ASRSysUtilAccessLog " & "(Type, UtilID, " & strMode & "By, " & strMode & "Date, " & strMode & "Host) " & "VALUES (" & "'" & utlType & "', " & CStr(lngID) & ", " & "system_user, getdate(), host_name() )"
		Else
			strSQL = "UPDATE ASRSysUtilAccessLog SET " & strMode & "By = system_user, " & strMode & "Date = getdate(), " & strMode & "Host = host_name() " & "WHERE UtilID = " & CStr(lngID) & " AND Type = " & CStr(utlType)
		End If
		gADOCon.Execute(strSQL)
		
		rsTemp.Close()
		'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsTemp = Nothing
		'UPGRADE_NOTE: Object datData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		datData = Nothing
		
	End Sub
	
	
	Public Sub DeleteUtilAccessLog(ByRef utlType As UtilityType, ByRef lngID As Integer)
		
		Dim strSQL As String
		
		strSQL = "DELETE FROM ASRSYSUtilAccessLog " & "WHERE UtilID = " & CStr(lngID) & " AND Type = " & CStr(utlType)
		
		gADOCon.Execute(strSQL)
		
	End Sub
End Module