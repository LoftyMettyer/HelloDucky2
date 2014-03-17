Imports System.Collections.Generic
Imports HR.Intranet.Server.Enums
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

	Public RegionalSettings As RegionalSettings
	Friend AbsenceModule As modAbsenceSpecifics
	Friend BankHolidayModule As modBankHolidaySpecifics
	Friend PersonnelModule As modPersonnelSpecifics

	Friend Tables As ICollection(Of Table)
	Friend Columns As ICollection(Of Column)
	Friend Relations As List(Of Relation)

	Friend ModuleSettings As ICollection(Of ModuleSetting)
	Friend UserSettings As ICollection(Of UserSetting)
	Friend SystemSettings As IList(Of UserSetting)

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

		ActiveConnections = 1
	End Sub

	Public Sub ReadModuleParameters()

		AbsenceModule = New modAbsenceSpecifics(Me)
		BankHolidayModule = New modBankHolidaySpecifics(Me)
		PersonnelModule = New modPersonnelSpecifics(Me)

		AbsenceModule.ReadAbsenceParameters()
		BankHolidayModule.ReadBankHolidayParameters()
		PersonnelModule.ReadPersonnelParameters()

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


#Region "FROM modPermissions"

	Friend Sub PopulateMetadata(Login As LoginInfo)

		Dim objDataAccess As New clsDataAccess(_objLogin)

		Tables = New Collection(Of Table)
		Columns = New Collection(Of Column)
		Relations = New List(Of Relation)
		ModuleSettings = New Collection(Of ModuleSetting)
		UserSettings = New Collection(Of UserSetting)
		SystemSettings = New List(Of UserSetting)
		Functions = New Collection(Of Metadata.Function)
		Operators = New Collection(Of Metadata.Operator)
		Modules = New List(Of ModuleSetting)

		Try

			Dim objData As DataSet = objDataAccess.GetDataSet("spASRGetMetadata", New SqlParameter("username", Login.Username))

			For Each objRow As DataRow In objData.Tables(0).Rows
				Dim table As New Table
				table.ID = CInt(objRow("TableID"))
				table.TableType = objRow("TableType")
				table.Name = objRow("TableName").ToString()
				table.DefaultOrderID = CInt(objRow("DefaultOrderID"))
				table.RecordDescExprID = CInt(objRow("RecordDescExprID"))
				Tables.Add(table)
			Next

			For Each objRow As DataRow In objData.Tables(1).Rows
				Dim column As New Column
				column.ID = CInt(objRow("columnid"))
				column.TableID = CInt(objRow("tableid"))
				column.TableName = Tables.GetById(column.TableID).Name
				column.Name = objRow("columnname").ToString()
				column.DataType = objRow("datatype")
				column.Use1000Separator = CBool(objRow("use1000separator"))
				column.Size = CLng(objRow("size"))
				column.Decimals = CShort(objRow("decimals"))
				Columns.Add(column)
			Next


			For Each objRow As DataRow In objData.Tables(2).Rows
				Dim relation As New Relation
				relation.ParentID = CInt(objRow("parentid"))
				relation.ChildID = CInt(objRow("childid"))
				Relations.Add(relation)
			Next


			For Each objRow As DataRow In objData.Tables(3).Rows
				Dim moduleSetting As New ModuleSetting
				moduleSetting.ModuleKey = objRow("ModuleKey").ToString()
				moduleSetting.ParameterKey = objRow("ParameterKey").ToString()
				moduleSetting.ParameterValue = objRow("ParameterValue").ToString()
				moduleSetting.ParameterType = objRow("ParameterType").ToString()
				ModuleSettings.Add(moduleSetting)
			Next


			For Each objRow As DataRow In objData.Tables(4).Rows
				Dim userSetting As New UserSetting
				userSetting.Section = objRow("Section").ToString()
				userSetting.Key = objRow("SettingKey").ToString()
				userSetting.Value = objRow("SettingValue")
				UserSettings.Add(userSetting)
			Next


			For Each objRow As DataRow In objData.Tables(5).Rows
				Dim objFunction = New [Function]
				objFunction.ID = CInt(objRow("functionID"))
				objFunction.Name = objRow("functionName").ToString()
				objFunction.ReturnType = CInt(objRow("returnType"))
				objFunction.Parameters = New Collection(Of FunctionParameter)()
				Functions.Add(objFunction)
			Next


			For Each objRow As DataRow In objData.Tables(6).Rows
				Dim objParameter = New [FunctionParameter]
				objParameter.ParameterIndex = CInt(objRow("ParameterIndex"))
				objParameter.ParameterType = objRow("ParameterType").ToString()
				objParameter.Name = objRow("ParameterName").ToString()
				Dim objFunction = Functions.GetById(CInt(objRow("functionID")))
				objFunction.Parameters.Add(objParameter)
			Next


			For Each objRow As DataRow In objData.Tables(7).Rows
				Dim objOperator = New [Operator]

				objOperator.ID = CInt(objRow("OperatorID"))
				objOperator.Name = objRow("Name").ToString()
				objOperator.ReturnType = objRow("returnType")
				objOperator.Precedence = CInt(objRow("Precedence"))
				objOperator.OperandCount = CInt(objRow("OperandCount"))
				objOperator.SPName = objRow("SPName").ToString()
				objOperator.SQLCode = objRow("SQLCode").ToString()
				objOperator.SQLType = objRow("SQLType").ToString()
				objOperator.CheckDivideByZero = CBool(objRow("CheckDivideByZero"))
				objOperator.SQLFixedParam1 = objRow("SQLFixedParam1").ToString()
				objOperator.CastAsFloat = CBool(objRow("CastAsFloat"))
				objOperator.Parameters = New Collection(Of OperatorParameter)()
				Operators.Add(objOperator)
			Next


			For Each objRow As DataRow In objData.Tables(8).Rows
				Dim objParameter = New OperatorParameter
				objParameter.ParameterType = objRow("ParameterType").ToString()
				Operators.GetById(CInt(objRow("operatorID"))).Parameters.Add(objParameter)
			Next


			For Each objRow As DataRow In objData.Tables(9).Rows
				Dim objModule = New ModuleSetting
				objModule.ModuleKey = objRow("Name").ToString()
				objModule.Enabled = CBool(objRow("Enabled"))
				Modules.Add(objModule)
			Next

			For Each objRow As DataRow In objData.Tables(10).Rows
				Dim systemSetting As New UserSetting
				systemSetting.Section = objRow("Section").ToString()
				systemSetting.Key = objRow("SettingKey").ToString()
				systemSetting.Value = objRow("SettingValue")
				SystemSettings.Add(systemSetting)
			Next


		Catch ex As Exception
			Throw

		End Try

	End Sub

	Friend Sub SetupTablesCollection()

		Const SecurityTable = 0
		Const SecurityPermissions = 1
		Const ViewTable = 2

		' Read the list of tables the current user has permission to see.
		Dim fSysSecManager As Boolean
		Dim sSQL As String

		Dim aryRealSource As DataTable

		Dim sTableViewName As String

		Dim objDataAccess As New clsDataAccess(_objLogin)
		Dim dsPermissions As DataSet

		Dim dtInfo As DataTable
		Dim objTableView As TablePrivilege
		Dim objColumnPrivileges As CColumnPrivileges
		Dim colTablePermissions As IList(Of TablePermission)
		Dim objTablePermission As TablePermission

		Dim sLastTableView As String
		Dim sColumnName As String
		Dim objItem As TablePrivilege

		' Don't need to recreate the tables & columns collections if they already exist.
		If Not gcoTablePrivileges Is Nothing Then
			Exit Sub
		End If

		' Instantiate a new collection of table privileges.
		gcoTablePrivileges = New Collection(Of TablePrivilege)()

		dsPermissions = objDataAccess.GetDataSet("spASRIntSetupTablesCollection")

		Dim objSecurityRow = dsPermissions.Tables(SecurityTable).Rows(0)
		gsUsername = objSecurityRow("UserName").ToString()
		gsActualLogin = objSecurityRow("ActualLogin").ToString()
		gsUserGroup = objSecurityRow("UserGroup").ToString()
		fSysSecManager = CBool(objSecurityRow("IsSysSecMgr"))


		' Initialise the collection with items for each TABLE in the system.
		For Each objTable In Tables
			objItem = New TablePrivilege()
			objItem.TableName = UCase(objTable.Name)
			objItem.TableID = objTable.ID
			objItem.TableType = objTable.TableType
			objItem.DefaultOrderID = objTable.DefaultOrderID
			objItem.RecordDescriptionID = objTable.RecordDescExprID
			objItem.IsTable = True
			objItem.ViewID = 0
			objItem.ViewName = ""
			objItem.RealSource = UCase(objTable.Name)
			gcoTablePrivileges.Add(objItem)
		Next

		' Initialise the collection with items for each VIEW in the system.
		For Each objRow As DataRow In dsPermissions.Tables(ViewTable).Rows
			objItem = New TablePrivilege()
			objItem.TableName = objRow("TableName").ToString()
			objItem.TableID = CInt(objRow("TableID"))
			objItem.TableType = objRow("TableType")
			objItem.DefaultOrderID = CInt(objRow("DefaultOrderID"))
			objItem.RecordDescriptionID = CInt(objRow("RecordDescExprID"))
			objItem.IsTable = False
			objItem.ViewID = CInt(objRow("ViewID"))
			objItem.ViewName = objRow("ViewName").ToString()
			objItem.RealSource = objRow("ViewName").ToString()
			gcoTablePrivileges.Add(objItem)
		Next


		Dim lngTableId As Long

		' Get the 'realSource' and permissions for each table or view.
		If fSysSecManager Then

			For Each objTableView In gcoTablePrivileges
				objTableView.AllowSelect = True
				objTableView.AllowUpdate = True
				objTableView.AllowDelete = True
				objTableView.AllowInsert = True
			Next

			sSQL = "SELECT tableid, childViewID FROM ASRSysChildViews2 WHERE role = '" & Replace(gsUserGroup, "'", "''") & "'"
			dtInfo = objDataAccess.GetDataTable(sSQL, CommandType.Text)

			For Each objRow As DataRow In dtInfo.Rows
				lngTableId = CInt(objRow("tableid"))
				objTableView = gcoTablePrivileges.GetItemByTableId(lngTableId)

				If objTableView.TableType = TableTypes.tabChild Then
					objTableView.RealSource = Left("ASRSysCV" & Trim(objRow("childViewID").ToString) & "#" & Replace(objTableView.TableName, " ", "_") & "#" & Replace(gsUserGroup, " ", "_"), 255)
				Else
					objTableView.RealSource = CStr(IIf(objTableView.IsTable, objTableView.TableName, objTableView.ViewName))

				End If

			Next

		Else
			' If the user is NOT a 'system manager' or 'security manager'
			' read the table permissions from the server.
			sSQL = "exec spASRIntAllTablePermissions '" & Replace(gsActualLogin, "'", "''") & "'"

			dtInfo = objDataAccess.GetDataTable(sSQL, CommandType.Text)

			colTablePermissions = New List(Of TablePermission)
			For Each objRow As DataRow In dtInfo.Rows
				objTablePermission = New TablePermission()
				objTablePermission.Name = objRow("Name").ToString()
				objTablePermission.Action = CInt(objRow("Action"))
				objTablePermission.TableID = CInt(objRow("TableID"))
				colTablePermissions.Add(objTablePermission)
			Next

			For Each objTablePermission In colTablePermissions

				objTableView = Nothing

				If Left(objTablePermission.Name, 8) = "ASRSYSCV" Then
					' Determine which table the child view is for.
					objTableView = gcoTablePrivileges.FindTableID(objTablePermission.TableID)

				Else
					If Left(objTablePermission.Name, 6) <> "ASRSYS" Then
						objTableView = gcoTablePrivileges.Item(objTablePermission.Name)
					End If
				End If

				If Not objTableView Is Nothing Then
					objTableView.RealSource = objTablePermission.Name

					Select Case objTablePermission.Action
						Case 193 ' Select permission.
							objTableView.AllowSelect = True
						Case 195 ' Insert permission.
							objTableView.AllowInsert = True
						Case 196 ' Delete permission.
							objTableView.AllowDelete = True
						Case 197 ' Update permission.
							objTableView.AllowUpdate = True
					End Select
				End If
			Next
		End If

		' Get the column permissions for each table/view.
		aryRealSource = New DataTable()
		'	Dim objblah = gcoTablePrivileges.WithRealSource()

		'	aryRealSource = gcoTablePrivileges.WithRealSource()

		aryRealSource.Columns.Add("tablename", Type.GetType("System.String"))
		For Each objTableView In gcoTablePrivileges.Collection

			If Len(objTableView.RealSource) > 0 Then

				Dim objRealSourceRow = aryRealSource.NewRow()
				objRealSourceRow("tablename") = objTableView.RealSource.ToUpper()

				aryRealSource.Rows.Add(objRealSourceRow)

			End If
		Next objTableView
		'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objTableView = Nothing

		' JPD20030313 Don't need to recreate the columns collection if it already exists.
		If Not gcolColumnPrivilegesCollection Is Nothing Then
			Exit Sub
		End If

		If gcoTablePrivileges.Count > 0 Then
			' Instantiate  the Column Privileges collection if it does not already exist.
			If gcolColumnPrivilegesCollection Is Nothing Then
				gcolColumnPrivilegesCollection = New Collection
			End If

			' Get the list of all columns in all tables/views.
			dtInfo = objDataAccess.GetDataTable("spASRIntGetColumnsFromTablesAndViews", CommandType.Text)

			For Each objRow As DataRow In dtInfo.Rows

				' If the current column's collection is NOT already instantiated, instantiate it.
				If sLastTableView <> objRow("tableviewname").ToString() Then
					sLastTableView = objRow("tableviewname").ToString()
					objColumnPrivileges = New CColumnPrivileges
					objColumnPrivileges.Tag = objRow("tableviewname").ToString()
					gcolColumnPrivilegesCollection.Add(objColumnPrivileges, objRow("tableviewname").ToString())
				End If

				sColumnName = objRow("ColumnName").ToString()
				If Not objColumnPrivileges.IsValid(sColumnName) Then
					' Add the column object to the collection.
					' If the current user is a system/security maneger then set column privileges to TRUE,
					' else set them to FALSE.
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					objColumnPrivileges.Add(fSysSecManager, fSysSecManager, sColumnName, objRow("ColumnType"), objRow("DataType"), objRow("ColumnID"), IIf(IsDBNull(objRow("UniqueCheckType")), False, objRow("UniqueCheckType") <> 0))

				End If

			Next

			' If the current user is not a system/security manager then read the column permissions from SQL.
			If Not fSysSecManager Then

				sLastTableView = ""

				dtInfo = objDataAccess.GetDataTable("spASRIntGetColumnPermissions", "SourceList", aryRealSource)
				For Each objRow As DataRow In dtInfo.Rows

					If sLastTableView <> objRow("tableviewname").ToString() Then
						sLastTableView = objRow("tableviewname").ToString()

						objTableView = gcoTablePrivileges.FindRealSource(objRow("tableviewname").ToString())
						If objTableView.IsTable Then
							sTableViewName = objTableView.TableName
						Else
							sTableViewName = objRow("tableviewname").ToString()
						End If

						objColumnPrivileges = gcolColumnPrivilegesCollection.Item(sTableViewName)
					End If

					If CInt(objRow("Action")) = 193 Then
						objColumnPrivileges.Item(objRow("ColumnName")).AllowSelect = CBool(objRow("Permission"))
					Else
						objColumnPrivileges.Item(objRow("ColumnName")).AllowUpdate = CBool(objRow("Permission"))
					End If

				Next
			End If
		End If

	End Sub


#End Region

End Class
