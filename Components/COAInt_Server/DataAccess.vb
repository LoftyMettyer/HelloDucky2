Option Strict Off
Option Explicit On
Friend Class clsDataAccess
	
	Public Function OpenRecordset(ByRef sSQL As String, ByRef CursorType As ADODB.CursorTypeEnum, ByRef LockType As ADODB.LockTypeEnum, Optional ByRef varCursorLocation As Object = Nothing) As ADODB.Recordset
		
		On Error GoTo ErrorTrap
		
		' Open a recordset from the given SQL query, with the given recordset properties.
		Dim rsTemp As ADODB.Recordset
		Dim fDoneOK As Boolean
		Dim iOldCursorLocation As Short
		
		iOldCursorLocation = gADOCon.CursorLocation
		fDoneOK = True
		
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(varCursorLocation) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object varCursorLocation. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			varCursorLocation = ADODB.CursorLocationEnum.adUseClient
		End If
		
		rsTemp = New ADODB.Recordset
		
		'UPGRADE_WARNING: Couldn't resolve default property of object varCursorLocation. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		gADOCon.CursorLocation = varCursorLocation
		
		rsTemp.Open(sSQL, gADOCon, CursorType, LockType, ADODB.CommandTypeEnum.adCmdText)
		
		gADOCon.CursorLocation = iOldCursorLocation
		
		If fDoneOK Then
			OpenRecordset = rsTemp
		End If
		
TidyUpAndExit: 
		If (iOldCursorLocation = ADODB.CursorLocationEnum.adUseClient) Or (iOldCursorLocation = ADODB.CursorLocationEnum.adUseServer) Then
			gADOCon.CursorLocation = iOldCursorLocation
		Else
			gADOCon.CursorLocation = ADODB.CursorLocationEnum.adUseServer
		End If
		
		Exit Function
		
ErrorTrap: 
		fDoneOK = False
		GoTo TidyUpAndExit
		
	End Function
	
	Public Function OpenPersistentRecordset(ByRef sSQL As String, ByRef CursorType As ADODB.CursorTypeEnum, ByRef LockType As ADODB.LockTypeEnum) As ADODB.Recordset
		' Open a recordset from the given SQL query, with the given recordset properties.
		Dim rsTemp As ADODB.Recordset
		
		rsTemp = New ADODB.Recordset
		
		rsTemp.let_ActiveConnection(gADOCon)
		rsTemp.Properties("Preserve On Commit").Value = True
		rsTemp.Properties("Preserve On Abort").Value = True
		rsTemp.Open(sSQL,  , CursorType, LockType, ADODB.CommandTypeEnum.adCmdText)
		
		OpenPersistentRecordset = rsTemp
		
	End Function
	
	
	Public Sub ExecuteSql(ByRef sSQL As String)
		' Execute the given SQL statement.
		gADOCon.Execute(sSQL,  , ADODB.CommandTypeEnum.adCmdText)
		
	End Sub
	
	Public Function ExecuteSqlReturnAffected(ByRef sSQL As String) As Integer
		' Execute the given SQL statement, and return the number of rows affected.
		Dim lngAffected As Integer
		
		gADOCon.Execute(sSQL, lngAffected, ADODB.CommandTypeEnum.adCmdText)
		ExecuteSqlReturnAffected = lngAffected
		
	End Function
End Class