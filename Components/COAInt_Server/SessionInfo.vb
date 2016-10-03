Imports System.Collections.Generic
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.Metadata
Imports HR.Intranet.Server.Structures
Imports System.Collections.ObjectModel
Imports System.Data.SqlClient
Imports System.Linq
Imports System.Web
Imports System.Security

Public Class SessionInfo

	Private _objLogin As LoginInfo
	Private _licenseKey As String

	Public DatabaseStatus As New DatabaseStatus
	Public Permissions As ICollection(Of Permission)

	Public RegionalSettings As RegionalSettings
	Friend AbsenceModule As modAbsenceSpecifics
	Friend BankHolidayModule As modBankHolidaySpecifics
	Friend PersonnelModule As modPersonnelSpecifics

   Public Tables As ICollection(Of Table)
   Public Views As ICollection(Of View)
   Public Columns As ICollection(Of Column)
	Public Relations As List(Of Relation)

	Friend ModuleSettings As ICollection(Of ModuleSetting)
	Friend UserSettings As ICollection(Of UserSetting)
	Friend SystemSettings As IList(Of UserSetting)

	Friend Functions As ICollection(Of Metadata.Function)
	Friend Operators As ICollection(Of Metadata.Operator)

	Friend gcoTablePrivileges As ICollection(Of TablePrivilege)
	Friend gcolColumnPrivilegesCollection As Collection

	Public ReadOnly Property LoginInfo As LoginInfo
		Get
			Return _objLogin
		End Get
	End Property

	Public Function IsCategoryGranted(Category As UtilityType) As Boolean

		For Each objPermission In Permissions.Where(Function(m) m.CategoryKey = Category.ToSecurityPrefix())
			If objPermission.IsPermitted Then Return True
		Next

		Return False
	End Function


	Public Function IsPermissionGranted(Category As String, Key As String) As Boolean
		Return Permissions.IsPermitted(Category, Key.ToUpper())
	End Function

	Public Function IsPhotoDataType(lngColumnID As Integer) As Boolean
		Return Columns.GetById(lngColumnID).DataType = ColumnDataType.sqlVarBinary
	End Function

	Public Function GetUserSetting(ByVal Section As String, ByVal Key As String, ByVal DefaultValue As Object) As Object

		Dim objSetting As UserSetting = UserSettings.GetUserSetting(Section, Key)

		If objSetting Is Nothing Then Return DefaultValue
		Return objSetting.Value

	End Function

	Public Function GetColumn(ColumnID As Integer) As Column

		Try
			If ColumnID > 0 Then
				Return Columns.GetById(ColumnID)
			Else
				Return New Column
			End If

		Catch ex As Exception
			Return New Column

		End Try

	End Function

	Public Function GetTable(TableID As Integer) As Table

		Try
			If TableID > 0 Then
				Return Tables.GetById(TableID)
			Else
				Return New Table
			End If

		Catch ex As Exception
			Return New Table

		End Try

	End Function

	Public Function SessionLogin(UserName As String, sPassword As String, sDatabaseName As String, sServerName As String, bWindowsAuthentication As Boolean, verifyOnly As Boolean) As LoginInfo

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

			If Not verifyOnly Then

				' Populate our system settings
				Permissions = New Collection(Of Permission)
				For Each objRow In dsLoginData.Tables(2).Rows
					Dim objPermissionItem = New Permission
					objPermissionItem.CategoryKey = objRow("categorykey").ToString()
					objPermissionItem.Key = objRow("itemkey").ToString()
					objPermissionItem.IsPermitted = CBool(objRow("permitted"))
					Permissions.Add(objPermissionItem)
				Next

				_objLogin.UserGroup = dsLoginData.Tables(0).Rows(0)(1).ToString()

				_objLogin.IsDMIUser = Permissions.IsPermitted("MODULEACCESS", "INTRANET")
				_objLogin.IsSSIUser = Permissions.IsPermitted("MODULEACCESS", "SSINTRANET")
				_objLogin.IsSystemOrSecurityAdmin = Permissions.IsPermitted("MODULEACCESS", "SYSTEMMANAGER")

				objRow = dsLoginData.Tables(3).Rows(0)
				_objLogin.IsServerRole = CBool(objRow("IsServeradmin")) OrElse CBool(objRow("IsSecurityadmin")) OrElse CBool(objRow("IsSysadmin"))

				If _objLogin.IsDMIUser Then
					_objLogin.DefaultWebArea = WebArea.DMI
				End If

			End If

		Catch ex As SqlException

			Select Case ex.Number

				' This procedure not found - likely an out of date database
				Case 2812
					DatabaseStatus.SysMgrVersion = New Version(0, 0, 0, 0)
					DatabaseStatus.IntranetVersion = New Version(0, 0, 0, 0)

					' Force Password change
				Case 18487, 18488
					_objLogin.MustChangePassword = True

				Case 18456, 4060
					' invalid login credentials
					_objLogin.LoginFailReason = "Login Failed."

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

		Try
			Tables = Nothing
			gcoTablePrivileges = Nothing
			gcolColumnPrivilegesCollection = Nothing

			PopulateMetadata(_objLogin)
			SetupTablesCollection()

		Catch ex As Exception
			Throw

		End Try

	End Sub

	Public Sub ReadModuleParameters()

		AbsenceModule = New modAbsenceSpecifics(Me)
		BankHolidayModule = New modBankHolidaySpecifics(Me)
		PersonnelModule = New modPersonnelSpecifics(Me)

		AbsenceModule.ReadAbsenceParameters()
		BankHolidayModule.ReadBankHolidayParameters()
		PersonnelModule.ReadPersonnelParameters()

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
				column.DataType = CType(objRow("datatype"), ColumnDataType)
				column.ColumnType = CType(objRow("columnType"), ColumnType)
				column.Use1000Separator = CBool(objRow("use1000separator"))
				column.Size = CLng(objRow("size"))
				column.ColumnSize = CLng(objRow("columnSize"))
				column.Decimals = CShort(objRow("decimals"))
				column.LookupTableID = CInt(objRow("LookupTableID"))
				column.LookupColumnID = CInt(objRow("LookupColumnID"))
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
         objItem.OriginalViewName = objRow("OriginalViewName").ToString()
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

         sSQL = "SELECT tableid, childViewID FROM ASRSysChildViews2 WHERE role = '" & Replace(LoginInfo.UserGroup, "'", "''") & "'"
         dtInfo = objDataAccess.GetDataTable(sSQL, CommandType.Text)

         For Each objRow As DataRow In dtInfo.Rows
            lngTableId = CInt(objRow("tableid"))
            objTableView = gcoTablePrivileges.GetItemByTableId(lngTableId)

            If objTableView.TableType = TableTypes.tabChild Then
               objTableView.RealSource = Left("ASRSysCV" & Trim(objRow("childViewID").ToString) & "#" & Replace(objTableView.TableName, " ", "_") & "#" & Replace(LoginInfo.UserGroup, " ", "_"), 255)
            Else
               objTableView.RealSource = CStr(IIf(objTableView.IsTable, objTableView.TableName, objTableView.ViewName))

            End If

         Next

      Else
         ' If the user is NOT a 'system manager' or 'security manager'
         ' read the table permissions from the server.
         sSQL = "exec spASRIntAllTablePermissions '" & Replace(LoginInfo.Username, "'", "''") & "'"
         dtInfo = objDataAccess.GetDataTable(sSQL, CommandType.Text)

         colTablePermissions = New List(Of TablePermission)
         For Each objRow As DataRow In dtInfo.Rows
            objTablePermission = New TablePermission()
            objTablePermission.Name = objRow("Name").ToString()
            objTablePermission.Action = CInt(objRow("Action"))
            objTablePermission.TableID = CInt(objRow("TableID"))

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

            colTablePermissions.Add(objTablePermission)

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

   Public Function GetTableAssociatedViews(tableId As Integer) As List(Of View)     ' Instantiate a new collection of table privileges.

      Dim views As New List(Of View)

      For Each objTableView As TablePrivilege In gcoTablePrivileges
         If (Not objTableView.IsTable) AndAlso (objTableView.TableID = tableId) AndAlso (objTableView.AllowSelect) Then
            Dim view As New View
            view.TableId = tableId
            view.TableName = objTableView.TableName
            view.ViewId = objTableView.ViewID
            view.ViewName = objTableView.OriginalViewName
            views.Add(view)
         End If
      Next
      Return views

   End Function

   Public Function ValidateColumnPermissions(psTableViewName As String, columnName As String) As Boolean     ' Instantiate a new collection of table privileges.

      Dim iLoop As Integer
      Dim objColumnPrivileges As New CColumnPrivileges
      Dim isValidColumn As Boolean

      Try
         ' If the given table/view's column privilege collection has already been read then simply return it.
         For iLoop = 1 To gcolColumnPrivilegesCollection.Count()
            If UCase(gcolColumnPrivilegesCollection.Item(iLoop).Tag) = UCase(psTableViewName) Then
               objColumnPrivileges = gcolColumnPrivilegesCollection.Item(iLoop)
               Exit For
            End If
         Next iLoop

         If objColumnPrivileges.IsValid(columnName) Then
            If objColumnPrivileges.Item(columnName).AllowSelect Then
               isValidColumn = True
            End If
         End If

         Return isValidColumn

      Catch ex As Exception
         Return isValidColumn
      End Try

   End Function

   Public Function GetSystemViewName(piTableID As String) As String

      Dim SystemViewName As String = ""

      Try
         For Each objTableView In gcoTablePrivileges.Collection
            If (objTableView.TableID = piTableID) And (objTableView.AllowSelect) Then
               SystemViewName = objTableView.RealSource
               Exit For
            End If
         Next objTableView

         Return SystemViewName.ToString()

      Catch ex As Exception
         Return SystemViewName.ToString()
      End Try

   End Function

#End Region

End Class

