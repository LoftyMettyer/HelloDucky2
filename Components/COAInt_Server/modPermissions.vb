Option Strict Off
Option Explicit On

Imports System.Collections.ObjectModel
Imports System.Collections.Generic
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.Metadata
Imports HR.Intranet.Server.Structures
Imports System.Data.SqlClient

Module modPermissions

	Friend Sub PopulateMetadata(ByVal Login As LoginInfo)

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

			Dim objData As DataSet = dataAccess.GetDataSet("spASRGetMetadata", New SqlParameter("username", Login.Username))

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

		dsPermissions = dataAccess.GetDataSet("spASRIntSetupTablesCollection")

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
			dtInfo = dataAccess.GetDataTable(sSQL, CommandType.Text)

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

			dtInfo = dataAccess.GetDataTable(sSQL, CommandType.Text)

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
			dtInfo = dataAccess.GetDataTable("spASRIntGetColumnsFromTablesAndViews", CommandType.Text)

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

				dtInfo = dataAccess.GetDataTable("spASRIntGetColumnPermissions", "SourceList", aryRealSource)
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

	Public Function GetColumnPrivileges(ByRef psTableViewName As String) As CColumnPrivileges
		' Return the column privileges collection for the given table.
		On Error GoTo ErrorTrap

		Dim fOK As Boolean
		Dim iLoop As Integer
		Dim objColumnPrivileges As CColumnPrivileges

		fOK = True

		' Instantiate  the Column Privileges collection if it does not already exist.
		If gcolColumnPrivilegesCollection Is Nothing Then
			gcolColumnPrivilegesCollection = New Collection
		End If

		' If the given table/view's column privilege collection has already been
		' read then simply return it.
		For iLoop = 1 To gcolColumnPrivilegesCollection.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object gcolColumnPrivilegesCollection().Tag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If UCase(gcolColumnPrivilegesCollection.Item(iLoop).Tag) = UCase(psTableViewName) Then
				Return gcolColumnPrivilegesCollection.Item(iLoop)
			End If
		Next iLoop

TidyUpAndExit:
		If fOK Then
			GetColumnPrivileges = objColumnPrivileges
		Else
			'UPGRADE_NOTE: Object GetColumnPrivileges may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			GetColumnPrivileges = Nothing
		End If
		Exit Function

ErrorTrap:
		'NO MSGBOX ON THE SERVER ! - MsgBox Err.Description & " - GetColumnPrivileges"
		fOK = False
		Resume TidyUpAndExit

	End Function

End Module