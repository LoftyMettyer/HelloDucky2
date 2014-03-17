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

		objSetting = SystemSettings.GetUserSetting(strSection.ToLower(), strKey.ToLower())
		If objSetting Is Nothing Then
			Return varDefault
		Else
			Return objSetting.Value
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

End Class