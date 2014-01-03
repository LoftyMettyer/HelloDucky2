Imports System.Collections.Generic
Imports ADODB
Imports HR.Intranet.Server.Metadata
Imports HR.Intranet.Server.Structures

Public Class SessionInfo
	'Public Shared gADOCon As ADODB.Connection

	'Public Shared datGeneral As New clsGeneral
	'Public Shared dataAccess As New clsDataAccess

	Public ActiveConnections As Integer = 0

	Public Property Username() As String
		Get
			Return gsUsername
		End Get
		Set(value As String)
			gsUsername = value
		End Set
	End Property

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
	Public ReadOnly Property Permissions() As ICollection(Of Permission)
		Get
			Return Declarations.Permissions
		End Get
	End Property

	Public Function IsPermissionGranted(ByVal sKey As String) As Boolean
		Return Declarations.Permissions.GetByKey(sKey)
	End Function

	Public Function GetUserSetting(ByVal Section As String, ByVal Key As String, ByVal DefaultValue As Object) As Object

		Dim objSetting As UserSetting = UserSettings.GetUserSetting(Section, Key)

		If objSetting Is Nothing Then Return DefaultValue
		Return objSetting.Value

	End Function

	Public Sub Initialise()
		Tables = Nothing
		gcoTablePrivileges = Nothing
		gcolColumnPrivilegesCollection = Nothing

		PopulateMetadata()		
		SetupTablesCollection()
		ActiveConnections = 1
	End Sub

	Public WriteOnly Property Connection() As Connection
		Set(ByVal Value As Connection)
			gADOCon = Value
		End Set
	End Property

	Public Sub LoginInfo(value As LoginInfo)
		Login = value
	End Sub

End Class
