Imports System.Collections.Generic
Imports HR.Intranet.Server.Metadata
Imports System.Collections.ObjectModel

Public Class SessionInfo
	'Public Shared gADOCon As ADODB.Connection

	'Public Shared datGeneral As New clsGeneral
	'Public Shared dataAccess As New clsDataAccess

	'Public Shared gsUsername As String
	'Public Shared gsActualLogin As String
	'Public Shared gsUserGroup As String

	'Public gcoTablePrivileges As ICollection(Of CTablePrivilege)

	'Friend gcolColumnPrivilegesCollection As Collection
	'Friend gcolLinks As Collection
	'Friend gcolNavigationLinks As Collection

	'Public Tables As ICollection(Of Metadata.Table)
	'Public Columns As ICollection(Of Metadata.Column)
	'Public Relations As ICollection(Of Metadata.Relation)
	'Public ModuleSettings As ICollection(Of Metadata.ModuleSetting)
	'Public UserSettings As ICollection(Of Metadata.UserSetting)
	'Public Functions As ICollection(Of Metadata.Function)
	'Public Operators As ICollection(Of Metadata.Operator)
	'Public ReadOnly Property Permissions() As ICollection(Of Permission)
	'	Get
	'		Return Declarations.Permissions
	'	End Get
	'End Property

	Public Function IsPermissionGranted(ByVal sKey As String) As Boolean
		Return Permissions.GetByKey(sKey)
	End Function


Public Sub Initialise()
	Tables = Nothing
	gcoTablePrivileges = Nothing
End Sub

End Class
