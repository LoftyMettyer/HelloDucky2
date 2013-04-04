Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
<System.Runtime.InteropServices.ProgId("clsSettings_NET.clsSettings")> Public Class clsSettings
	
	Private Declare Function RegOpenKeyEx Lib "advapi32"  Alias "RegOpenKeyExA"(ByVal hKey As Integer, ByVal lpSubKey As String, ByVal ulOptions As Integer, ByVal samDesired As Integer, ByRef phkResult As Integer) As Integer
	Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Integer) As Integer
	
	Private Const HKEY_LOCAL_MACHINE As Integer = &H80000002
	Private Const KEY_READ As Integer = &H20019
	
	Public Function GetUserSetting(ByRef strSection As String, ByRef strKey As String, ByRef varDefault As Object) As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object modSettings.GetUserSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetUserSetting. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetUserSetting = modSettings.GetUserSetting(strSection, strKey, varDefault)
	End Function
	
	
	Public Function GetWordColourIndex(ByRef lngColourValue As Integer) As Integer
		
		Dim rsTemp As ADODB.Recordset
		Dim strSQL As String
		
		On Error GoTo LocalErr
		
		strSQL = "SELECT WordColourIndex FROM ASRSysColours " & " WHERE ColValue = " & CStr(lngColourValue)
		rsTemp = datGeneral.GetReadOnlyRecords(strSQL)
		
		With rsTemp
			If Not .BOF And Not .EOF Then
				GetWordColourIndex = rsTemp.Fields("WordColourIndex").Value
			Else
				GetWordColourIndex = 0
			End If
		End With
		
		rsTemp.Close()
		'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsTemp = Nothing
		
		Exit Function
		
LocalErr: 
		GetWordColourIndex = 0
		
	End Function
	
	
	Public Function GetEmailGroupAddresses(ByRef lngGroupID As Integer) As String
		
		Dim datData As clsDataAccess
		Dim rsTemp As ADODB.Recordset
		Dim strSQL As String
		Dim strOutput As String
		
		On Error GoTo LocalErr
		
		strOutput = vbNullString
		
		datData = New clsDataAccess
		
		strSQL = "SELECT ASRSysEmailAddress.Fixed From ASRSysEmailGroupName " & "JOIN ASRSysEmailGroupItems ON ASRSysEmailGroupName.EmailGroupID = ASRSysEmailGroupItems.EmailGroupID " & "JOIN ASRSysEmailAddress ON ASRSysEmailGroupItems.EmailDefID = ASRSysEmailAddress.EmailID " & "WHERE ASRSysEmailGroupName.EmailGroupID = " & CStr(lngGroupID)
		rsTemp = datData.OpenRecordset(strSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		With rsTemp
			If Not .BOF And Not .EOF Then
				
				Do While Not rsTemp.EOF
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If Not IsDbNull(rsTemp.Fields(0).Value) Then
						strOutput = strOutput & IIf(strOutput <> vbNullString, ";", "") & rsTemp.Fields(0).Value
					End If
					rsTemp.MoveNext()
				Loop 
				
			End If
		End With
		
		rsTemp.Close()
		
		'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsTemp = Nothing
		'UPGRADE_NOTE: Object datData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		datData = Nothing
		
		
		GetEmailGroupAddresses = strOutput
		
		Exit Function
		
LocalErr: 
		
	End Function
	
	Public Function GetSystemSetting(ByRef strSection As Object, ByRef strKey As Object, ByRef varDefault As Object) As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object strSection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		strSection = Replace(strSection, "'", "''")
		'UPGRADE_WARNING: Couldn't resolve default property of object strKey. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		strKey = Replace(strKey, "'", "''")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object strKey. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object strSection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object modSettings.GetSystemSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetSystemSetting. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetSystemSetting = modSettings.GetSystemSetting(CStr(strSection), CStr(strKey), varDefault)
	End Function
	
	Public Function GetModuleParameter(ByRef psModuleKey As String, ByRef psParameterKey As String) As String
		' Return the value of the given parameter.
		GetModuleParameter = datGeneral.GetModuleParameter(psModuleKey, psParameterKey)
		
	End Function
	Public WriteOnly Property Connection() As Object
		Set(ByVal Value As Object)
			
			' Connection object passed in from the asp page
			
			' JDM - Create connection object differently if we are in development mode (i.e. debug mode)
			If ASRDEVELOPMENT Then
				gADOCon = New ADODB.Connection
				'UPGRADE_WARNING: Couldn't resolve default property of object vConnection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				gADOCon.Open(Value)
			Else
				gADOCon = Value
			End If
			
		End Set
	End Property
	
	Public Function GetFieldNameFromModuleSetup(ByRef psModuleKey As Object, ByRef psParameterKey As Object) As Object
		
		Dim sColumnID As String
		'UPGRADE_WARNING: Couldn't resolve default property of object psParameterKey. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object psModuleKey. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sColumnID = GetModuleParameter(CStr(psModuleKey), CStr(psParameterKey))
		
		If IsNumeric(sColumnID) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object GetFieldNameFromModuleSetup. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetFieldNameFromModuleSetup = datGeneral.GetColumnName(CInt(sColumnID))
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object GetFieldNameFromModuleSetup. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetFieldNameFromModuleSetup = ""
		End If
		
	End Function
	
	' Return date of report in SQL (American date format)
	Public Function GetStandardReportDate(ByRef psReportType As Object, ByRef psReportDateType As Object) As Object
		
		Dim blnCustom As Boolean
		Dim strRecSelStatus As String
		Dim lngID As Integer
		Dim lngCount As Integer
		Dim lngDateExprID As Integer
		Dim dStartDate As Date
		Dim dEndDate As Date
		
		'UPGRADE_WARNING: Couldn't resolve default property of object psReportType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetSystemSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		blnCustom = (GetSystemSetting(CStr(psReportType), "Custom Dates", "0") = "1")
		
		If blnCustom Then
			'UPGRADE_WARNING: Couldn't resolve default property of object psReportDateType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object psReportType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object GetSystemSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			lngDateExprID = GetSystemSetting(CStr(psReportType), CStr(psReportDateType), 0)
			strRecSelStatus = IsCalcValid(lngDateExprID)
			If strRecSelStatus <> vbNullString Then
				dEndDate = DateAdd(Microsoft.VisualBasic.DateInterval.Day, VB.Day(Today) * -1, Today)
				'UPGRADE_WARNING: Couldn't resolve default property of object psReportDateType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If psReportDateType = "End Date" Then
					GetStandardReportDate = VB6.Format(dEndDate, "mm/dd/yyyy")
				Else
					GetStandardReportDate = VB6.Format(DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, DateAdd(Microsoft.VisualBasic.DateInterval.Year, -1, dEndDate)), "mm/dd/yyyy")
				End If
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object datGeneral.GetValueForRecordIndependantCalc(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				GetStandardReportDate = VB6.Format(datGeneral.GetValueForRecordIndependantCalc(lngDateExprID), "mm/dd/yyyy")
			End If
			
		Else
			dEndDate = DateAdd(Microsoft.VisualBasic.DateInterval.Day, VB.Day(Today) * -1, Today)
			'UPGRADE_WARNING: Couldn't resolve default property of object psReportDateType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If psReportDateType = "End Date" Then
				GetStandardReportDate = VB6.Format(dEndDate, "mm/dd/yyyy")
			Else
				GetStandardReportDate = VB6.Format(DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, DateAdd(Microsoft.VisualBasic.DateInterval.Year, -1, dEndDate)), "mm/dd/yyyy")
			End If
		End If
		
	End Function
	
	Public Function GetPicklistFilterName(ByRef psReportType As Object, ByRef pstrType As Object) As Object
		
		Dim strRecSelStatus As String
		Dim plngID As Integer
		Dim strName As String
		
		SetupTablesCollection()
		
		'UPGRADE_WARNING: Couldn't resolve default property of object GetSystemSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		plngID = GetSystemSetting(psReportType, "ID", 0)
		
		Select Case pstrType
			Case "A"
				
			Case "F"
				strRecSelStatus = IsFilterValid(plngID)
				If strRecSelStatus <> vbNullString Then
					strName = "<None>"
				Else
					strName = datGeneral.GetFilterName(plngID)
				End If
			Case "P"
				strRecSelStatus = IsPicklistValid(plngID)
				If strRecSelStatus <> vbNullString Then
					strName = "<None>"
				Else
					strName = datGeneral.GetPicklistName(plngID)
				End If
		End Select
		
		GetPicklistFilterName = Replace(strName, """", "")
		
		
	End Function
	
	Public Function GetTableNameFromModuleSetup(ByRef psModuleKey As Object, ByRef psParameterKey As Object) As Object
		
		Dim sTableID As String
		'UPGRADE_WARNING: Couldn't resolve default property of object psParameterKey. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object psModuleKey. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sTableID = GetModuleParameter(CStr(psModuleKey), CStr(psParameterKey))
		
		If IsNumeric(sTableID) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object GetTableNameFromModuleSetup. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetTableNameFromModuleSetup = datGeneral.GetTableName(CInt(sTableID))
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object GetTableNameFromModuleSetup. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetTableNameFromModuleSetup = ""
		End If
		
	End Function
	
	Public Function GetEmailGroupName(ByRef lngGroupID As Integer) As String
		
		Dim datData As clsDataAccess
		Dim rsTemp As ADODB.Recordset
		Dim strSQL As String
		Dim strOutput As String
		
		strOutput = vbNullString
		
		datData = New clsDataAccess
		
		strSQL = "SELECT Name From ASRSysEmailGroupName " & "WHERE EmailGroupID = " & CStr(lngGroupID)
		rsTemp = datData.OpenRecordset(strSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		With rsTemp
			If Not .BOF And Not .EOF Then
				
				Do While Not rsTemp.EOF
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If Not IsDbNull(rsTemp.Fields(0).Value) Then
						strOutput = rsTemp.Fields(0).Value
					End If
					rsTemp.MoveNext()
				Loop 
				
			End If
		End With
		
		rsTemp.Close()
		
		'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsTemp = Nothing
		'UPGRADE_NOTE: Object datData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		datData = Nothing
		
		GetEmailGroupName = strOutput
		
	End Function
	
	Public Function GetSQLNCLIVersion() As Short
		On Error GoTo SQLNCLI_Err
		
		Dim Rc As Integer ' Return Code
		Dim hKey As Integer ' Handle To An Open Registry Key
		Dim tmpKey As Short
		tmpKey = 0
		
		' Paths to the SQL Native Client registry keys
		Const sREGKEYSQLNCLI As String = "SOFTWARE\Microsoft\Microsoft SQL Native Client\CurrentVersion"
		Const sREGKEYSQLNCLI10 As String = "SOFTWARE\Microsoft\Microsoft SQL Server Native Client 10.0\CurrentVersion"
		Const sREGKEYSQLNCLI11 As String = "SOFTWARE\Microsoft\Microsoft SQL Server Native Client 11.0\CurrentVersion"
		
		Rc = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sREGKEYSQLNCLI, 0, KEY_READ, hKey) ' Open Registry Key
		If (Rc = 0) Then
			tmpKey = 9
			Rc = RegCloseKey(hKey) ' Close Registry Key
		End If
		
		Rc = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sREGKEYSQLNCLI10, 0, KEY_READ, hKey) ' Open Registry Key
		If (Rc = 0) Then
			tmpKey = 10
			Rc = RegCloseKey(hKey) ' Close Registry Key
		End If
		
		Rc = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sREGKEYSQLNCLI11, 0, KEY_READ, hKey) ' Open Registry Key
		If (Rc = 0) Then
			tmpKey = 11
			Rc = RegCloseKey(hKey) ' Close Registry Key
		End If
		
		
SQLNCLI_Err_Handler: 
		GetSQLNCLIVersion = tmpKey
		Exit Function
		
SQLNCLI_Err: 
		Rc = RegCloseKey(hKey) ' Close Registry Key
		tmpKey = 0
		Resume SQLNCLI_Err_Handler
	End Function
	
	' Get native provider string
	Public Function GetSQLProviderString() As String
		
		Dim iVersion As Short
		
		iVersion = GetSQLNCLIVersion
		
		If iVersion = 9 Then
			GetSQLProviderString = "Provider=SQLNCLI;"
		ElseIf iVersion = 10 Then 
			GetSQLProviderString = "Provider=SQLNCLI10;"
		ElseIf iVersion = 11 Then 
			GetSQLProviderString = "Provider=SQLNCLI11;"
		Else
			GetSQLProviderString = vbNullString
		End If
		
	End Function
End Class