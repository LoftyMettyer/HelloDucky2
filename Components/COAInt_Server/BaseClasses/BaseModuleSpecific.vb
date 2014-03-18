Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports HR.Intranet.Server.Metadata
Imports HR.Intranet.Server.Structures

Namespace BaseClasses
	Public Class BaseModuleSpecific

		Protected _objLogin As LoginInfo
		Protected _tables As ICollection(Of Table)
		Protected _columns As ICollection(Of Column)
		Protected _moduleSettings As ICollection(Of ModuleSetting)
		Protected _systemSettings As IList(Of UserSetting)
		Protected _tablePrivileges As ICollection(Of TablePrivilege)

		Public Sub New(value As SessionInfo)
			_objLogin = value.LoginInfo
			_tables = value.Tables
			_columns = value.Columns
			_moduleSettings = value.ModuleSettings
			_systemSettings = value.SystemSettings
			_tablePrivileges = value.gcoTablePrivileges
		End Sub

		Friend Function GetModuleParameter(psModuleKey As String, psParameterKey As String) As String
			Return _moduleSettings.GetSetting(psModuleKey, psParameterKey).ParameterValue
		End Function

		Friend Function GetSystemSetting(strSection As String, strKey As String, varDefault As Object) As Object
			Return _systemSettings.GetSetting(strSection, strKey, varDefault).Value
		End Function

	End Class
End Namespace
