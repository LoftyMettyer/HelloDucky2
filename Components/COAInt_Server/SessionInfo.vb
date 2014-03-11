Imports System.Collections.Generic
Imports HR.Intranet.Server.Metadata
Imports HR.Intranet.Server.Structures
Imports System.Collections.ObjectModel
Imports System.Data.SqlClient
Imports System.Web

Public Class SessionInfo

	Private _objLogin As LoginInfo
	Private _licenseKey As String

	Public ActiveConnections As Integer = 0
	Public DatabaseStatus As New DatabaseStatus
	Public Permissions As ICollection(Of Permission)

	Public ReadOnly Property LoginInfo As LoginInfo
		Get
			Return _objLogin
		End Get
	End Property

	Public Function IsPermissionGranted(Category As String, Key As String) As Boolean
		Return Permissions.IsPermitted(Category, Key)
	End Function

	Public Function IsModuleEnabled(name As String) As Boolean
		Return Modules.GetByKey(name).Enabled
	End Function

	Public Function GetUserSetting(ByVal Section As String, ByVal Key As String, ByVal DefaultValue As Object) As Object

		Dim objSetting As UserSetting = UserSettings.GetUserSetting(Section, Key)

		If objSetting Is Nothing Then Return DefaultValue
		Return objSetting.Value

	End Function

	Public Function SessionLogin(UserName As String, sPassword As String, sDatabaseName As String, sServerName As String, bWindowsAuthentication As Boolean) As LoginInfo

		Dim objRow As DataRow

		_objLogin = New LoginInfo With {
			.Username = UserName,
			.Password = sPassword,
			.Database = sDatabaseName,
			.Server = sServerName,
			.TrustedConnection = bWindowsAuthentication}

		Try

			Dim objDataAccess As New clsDataAccess(_objLogin)
			Dim dsLoginData As DataSet = objDataAccess.GetDataSet("spASRIntGetLoginDetails")

			Dim rowDBInfo = dsLoginData.Tables(1).Rows(0)
			_licenseKey = rowDBInfo("LicenseKey").ToString()

			DatabaseStatus.SysMgrVersion = Version.Parse(rowDBInfo("SysMgrDBVersion").ToString())
			DatabaseStatus.IntranetVersion = Version.Parse(rowDBInfo("IntDBVersion").ToString())
			DatabaseStatus.IsUpdateInProgress = CBool(rowDBInfo("UpdateInProgress"))
			DatabaseStatus.IsLocked = CBool(rowDBInfo("IsLocked"))
			DatabaseStatus.LockMessage = rowDBInfo("lockmessage").ToString()

			' Populate our system settings
			Permissions = New Collection(Of Permission)
			For Each objRow In dsLoginData.Tables(2).Rows
				Dim objPermissionItem = New Permission
				objPermissionItem.CategoryKey = objRow("categorykey").ToString()
				objPermissionItem.Key = objRow("itemkey").ToString()
				objPermissionItem.IsPermitted = CBool(objRow("permitted"))
				Permissions.Add(objPermissionItem)
			Next

			_objLogin.UserGroup = dsLoginData.Tables(0).Rows(0)(1)

			_objLogin.IsDMIUser = Permissions.IsPermitted("MODULEACCESS", "INTRANET")
			_objLogin.IsDMISingle = Permissions.IsPermitted("MODULEACCESS", "INTRANET_SELFSERVICE")
			_objLogin.IsSSIUser = Permissions.IsPermitted("MODULEACCESS", "SSINTRANET")
			_objLogin.IsSystemOrSecurityAdmin = Permissions.IsPermitted("MODULEACCESS", "SYSTEMMANAGER")

			objRow = dsLoginData.Tables(3).Rows(0)
			_objLogin.IsServerRole = CBool(objRow("IsServeradmin")) Or CBool(objRow("IsSecurityadmin")) Or CBool(objRow("IsSysadmin"))


		Catch ex As SqlException

			Select Case ex.Number

				' This procedure not found - likely an out of date database
				Case 2812
					DatabaseStatus.SysMgrVersion = New Version(0, 0, 0, 0)
					DatabaseStatus.IntranetVersion = New Version(0, 0, 0, 0)

					' Force Password change
				Case 18487, 18488
					_objLogin.MustChangePassword = True

					' Anything else
				Case Else
					_objLogin.LoginFailReason = ex.Message

			End Select

		Catch ex As Exception
			Throw

		End Try

		Return _objLogin

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

	Public Sub TrackUser(IsLogin As Boolean)

		Dim objDataAccess As New clsDataAccess(_objLogin)
		Dim sMachineName As String

		Try
			Dim objUserMachine = Net.Dns.GetHostEntry(HttpContext.Current.Request.UserHostName)
			sMachineName = objUserMachine.HostName

		Catch ex As Exception
			sMachineName = "Unknown"

		End Try

		Try
			Dim prmLoginTime = New SqlParameter("LoginTime", SqlDbType.DateTime) With {.Direction = ParameterDirection.Output}

			objDataAccess.ExecuteSP("spASRTrackSession" _
					, New SqlParameter("LoggingIn", SqlDbType.Bit) With {.Value = IsLogin} _
					, New SqlParameter("Application", SqlDbType.VarChar, 255) With {.Value = "OpenHR Web"} _
					, New SqlParameter("ClientMachine", SqlDbType.VarChar, 255) With {.Value = sMachineName} _
					, prmLoginTime)

			_objLogin.LoginTime = prmLoginTime.Value


		Catch ex As Exception
			Throw

		End Try

	End Sub

End Class
