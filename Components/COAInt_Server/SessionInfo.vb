Imports System.Collections.Generic
Imports ADODB
Imports HR.Intranet.Server.Metadata
Imports HR.Intranet.Server.Structures

Public Class SessionInfo

	Private _objLogin As LoginInfo

	Public ActiveConnections As Integer = 0

	Public Property Username() As String
		Get
			Return gsUsername
		End Get
		Set(value As String)
			gsUsername = value
		End Set
	End Property

	Public ReadOnly Property Permissions() As ICollection(Of Permission)
		Get
			Return Declarations.Permissions
		End Get
	End Property

	Public Function IsPermissionGranted(ByVal sKey As String) As Boolean
		Return Declarations.Permissions.GetByKey(sKey)
	End Function

	Public Function IsModuleEnabled(ByVal name As String) As Boolean
		Return Modules.GetByKey(name).Enabled
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

		PopulateMetadata(_objLogin)
		SetupTablesCollection()

		ReadPersonnelParameters()

		ActiveConnections = 1
	End Sub

	Public Property Connection() As Connection
		Get
			Return gADOCon
		End Get
		Set(value As Connection)
			gADOCon = value
		End Set
	End Property

	Public Property LoginInfo() As LoginInfo
		Get
			Return _objLogin
		End Get
		Set(value As LoginInfo)

			_objLogin = value
			dataAccess = New clsDataAccess(value)

		End Set
	End Property


End Class
