Option Strict On
Option Explicit On

Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Metadata
Imports VB = Microsoft.VisualBasic

Public Class clsSettings
	Inherits BaseForDMI

	Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Integer, ByVal lpSubKey As String, ByVal ulOptions As Integer, ByVal samDesired As Integer, ByRef phkResult As Integer) As Integer
	Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Integer) As Integer

	Private Const HKEY_LOCAL_MACHINE As Integer = &H80000002
	Private Const KEY_READ As Integer = &H20019

	Public Function GetUserSetting(ByRef strSection As String, ByRef strKey As String, ByRef varDefault As Object) As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object modSettings.GetUserSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetUserSetting. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetUserSetting = General.GetUserSetting(strSection, strKey, varDefault)
	End Function

	Public Function GetWordColourIndex(ByRef lngColourValue As Long) As Integer

		Dim sSQL As String = String.Format("SELECT WordColourIndex FROM ASRSysColours WHERE ColValue = {0}", lngColourValue)
		With DB.GetDataTable(sSQL)
			If .Rows.Count > 0 Then
				Return CInt(.Rows(0)(0))
			Else
				Return 0
			End If
		End With

	End Function

	Public Function GetSystemSetting(ByRef strSection As String, ByRef strKey As String, ByRef varDefault As Object) As Object

		''UPGRADE_WARNING: Couldn't resolve default property of object strSection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'strSection = Replace(strSection, "'", "''")
		''UPGRADE_WARNING: Couldn't resolve default property of object strKey. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'strKey = Replace(strKey, "'", "''")

		Dim objSetting As UserSetting

		objSetting = SystemSettings.GetUserSetting(strSection, strKey)
		If objSetting Is Nothing Then
			Return varDefault
		Else
			Return objSetting.Value
		End If

	End Function

	' Return date of report in SQL (American date format)
	Public Function GetStandardReportDate(ByRef psReportType As String, ByRef psReportDateType As String) As String

		Dim blnCustom As Boolean
		Dim strRecSelStatus As String
		Dim lngDateExprID As Integer
		Dim dEndDate As Date

		'UPGRADE_WARNING: Couldn't resolve default property of object psReportType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetSystemSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		blnCustom = (GetSystemSetting(psReportType, "Custom Dates", "0").ToString() = "1")

		If blnCustom Then
			'UPGRADE_WARNING: Couldn't resolve default property of object psReportDateType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object psReportType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object GetSystemSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			lngDateExprID = CInt(GetSystemSetting(CStr(psReportType), CStr(psReportDateType), 0))
			strRecSelStatus = IsCalcValid(lngDateExprID)

			If strRecSelStatus <> vbNullString Then
				dEndDate = DateAdd(Microsoft.VisualBasic.DateInterval.Day, VB.Day(Today) * -1, Today)
				'UPGRADE_WARNING: Couldn't resolve default property of object psReportDateType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If psReportDateType = "End Date" Then
					GetStandardReportDate = VB6.Format(dEndDate, "MM/dd/yyyy")
				Else
					GetStandardReportDate = VB6.Format(DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, DateAdd(Microsoft.VisualBasic.DateInterval.Year, -1, dEndDate)), "MM/dd/yyyy")
				End If
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object datGeneral.GetValueForRecordIndependantCalc(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				GetStandardReportDate = VB6.Format(General.GetValueForRecordIndependantCalc(lngDateExprID), "MM/dd/yyyy")

				dEndDate = CDate(General.GetValueForRecordIndependantCalc(lngDateExprID))
				GetStandardReportDate = dEndDate.ToString("MM/dd/yyyy")


			End If

		Else
			dEndDate = DateAdd(Microsoft.VisualBasic.DateInterval.Day, VB.Day(Today) * -1, Today)
			'UPGRADE_WARNING: Couldn't resolve default property of object psReportDateType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If psReportDateType = "End Date" Then
				GetStandardReportDate = VB6.Format(dEndDate, "MM/dd/yyyy")
			Else
				GetStandardReportDate = VB6.Format(DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, DateAdd(Microsoft.VisualBasic.DateInterval.Year, -1, dEndDate)), "MM/dd/yyyy")
			End If
		End If

	End Function

	Public Function GetPicklistFilterName(ByRef psReportType As String, ByRef pstrType As String) As String

		Dim strRecSelStatus As String
		Dim plngID As Integer
		Dim strName As String

		'UPGRADE_WARNING: Couldn't resolve default property of object GetSystemSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		plngID = CInt(GetSystemSetting(psReportType, "ID", 0))

		Select Case pstrType
			Case "A"

			Case "F"
				strRecSelStatus = IsFilterValid(plngID)
				If strRecSelStatus <> vbNullString Then
					strName = "<None>"
				Else
					strName = General.GetFilterName(plngID)
				End If
			Case "P"
				strRecSelStatus = IsPicklistValid(plngID)
				If strRecSelStatus <> vbNullString Then
					strName = "<None>"
				Else
					strName = General.GetPicklistName(plngID)
				End If
		End Select

		Return Replace(strName, """", "")


	End Function

	Public Function GetTableNameFromModuleSetup(ByRef psModuleKey As Object, ByRef psParameterKey As Object) As Object

		Dim sTableID As String
		'UPGRADE_WARNING: Couldn't resolve default property of object psParameterKey. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object psModuleKey. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sTableID = GetModuleParameter(CStr(psModuleKey), CStr(psParameterKey))

		If IsNumeric(sTableID) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object GetTableNameFromModuleSetup. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetTableNameFromModuleSetup = GetTableName(CInt(sTableID))
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object GetTableNameFromModuleSetup. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetTableNameFromModuleSetup = ""
		End If

	End Function

	Public Function GetEmailGroupName(lngGroupID As Integer) As String

		Dim sSQL As String = String.Format("SELECT Name From ASRSysEmailGroupName WHERE EmailGroupID = {0}", lngGroupID)
		With DB.GetDataTable(sSQL)
			If .Rows.Count > 0 Then
				Return Trim(.Rows(0)(0).ToString())
			Else
				Return vbNullString
			End If
		End With

	End Function

	Public Function GetSQLNCLIVersion() As Short
		On Error GoTo SQLNCLI_Err

		Dim Rc As Integer	' Return Code
		Dim hKey As Integer	' Handle To An Open Registry Key
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

		iVersion = GetSQLNCLIVersion()

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