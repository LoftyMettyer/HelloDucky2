Option Strict Off
Option Explicit On

Imports System.Collections.ObjectModel
Imports System.Collections.Generic
Imports ADODB
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.Metadata
Imports HR.Intranet.Server.Structures

Module modPermissions

	Public Sub SetupTablesCollection()
		' Read the list of tables the current user has permission to see.
		Dim fSysSecManager As Boolean
		Dim sSQL As String

		Dim aryRealSource As DataTable

		Dim sTableViewName As String
		Dim rsInfo As Recordset
		Dim rsViews As DataTable
		Dim rsPermissions As Recordset
		Dim objTableView As TablePrivilege
		Dim objColumnPrivileges As CColumnPrivileges
		Dim colTablePermissions As IList(Of TablePermission)
		Dim objTablePermission As TablePermission

		Dim sLastTableView As String
		Dim sColumnName As String
		Dim iOriginalCursorLocation As Short
		Dim objItem As TablePrivilege

		If Tables Is Nothing Then
			datGeneral.PopulateMetadata()
		End If

		' Don't need to recreate the tables & columns collections if they already exist.
		If Not gcoTablePrivileges Is Nothing Then
			Exit Sub
		End If

		' Switch to client cursor for performance reasons.
		iOriginalCursorLocation = gADOCon.CursorLocation

		' Instantiate a new collection of table privileges.
		gcoTablePrivileges = New Collection(Of TablePrivilege)()

		sSQL = "SELECT system_user AS [name]"
		rsInfo = New Recordset
		rsInfo.Open(sSQL, gADOCon, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly, CommandTypeEnum.adCmdText)
		datGeneral.Username = rsInfo.Fields("Name").Value
		rsInfo.Close()
		'UPGRADE_NOTE: Object rsInfo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsInfo = Nothing

		' Check if the user is a 'system manager' or 'security manager'.
		' If so then we can save time by applying all table permissions, instead of having to read them first.
		sSQL = "SELECT count(*) AS recCount FROM ASRSysGroupPermissions INNER JOIN ASRSysPermissionItems ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID" & " INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & " INNER JOIN sysusers a ON ASRSysGroupPermissions.groupName = a.name" & "   AND a.name = '" & gsUserGroup & "'" & " WHERE (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & " OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER')" & " AND ASRSysGroupPermissions.permitted = 1" & " AND ASRSysPermissionCategories.categorykey = 'MODULEACCESS'"

		rsInfo = New Recordset
		rsInfo.Open(sSQL, gADOCon, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly, CommandTypeEnum.adCmdText)
		fSysSecManager = (rsInfo.Fields("recCount").Value > 0)
		rsInfo.Close()
		'UPGRADE_NOTE: Object rsInfo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsInfo = Nothing

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
		sSQL = "SELECT v.viewID, UPPER(v.viewName) AS [viewname], t.tableID, UPPER(t.tableName) AS [tablename], t.tableType, t.defaultOrderID, t.recordDescExprID FROM ASRSysViews v INNER JOIN ASRSysTables t ON v.viewTableID = t.tableID"
		rsViews = clsDataAccess.GetDataTable(sSQL, CommandType.Text)

		For Each objRow In rsViews.Rows
			objItem = New TablePrivilege()
			objItem.TableName = objRow("TableName")
			objItem.TableID = objRow("TableID")
			objItem.TableType = objRow("TableType")
			objItem.DefaultOrderID = objRow("DefaultOrderID")
			objItem.RecordDescriptionID = objRow("RecordDescExprID")
			objItem.IsTable = False
			objItem.ViewID = objRow("ViewID")
			objItem.ViewName = objRow("ViewName")
			objItem.RealSource = objRow("ViewName")
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
			rsInfo = New Recordset
			rsInfo.Open(sSQL, gADOCon, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly, CommandTypeEnum.adCmdText)

			With rsInfo
				.MoveFirst()
				Do While Not .EOF
					lngTableId = Trim(rsInfo.Fields("tableid").Value)
					objTableView = gcoTablePrivileges.GetItemByTableId(lngTableId)

					If objTableView.TableType = TableTypes.tabChild Then
						objTableView.RealSource = Left("ASRSysCV" & Trim(Str(.Fields("childViewID").Value)) & "#" & Replace(objTableView.TableName, " ", "_") & "#" & Replace(gsUserGroup, " ", "_"), 255)
					Else
						objTableView.RealSource = IIf(objTableView.IsTable, objTableView.TableName, objTableView.ViewName)
					End If

					.MoveNext()
				Loop
				.Close()
			End With

			objTableView = Nothing
		Else
			' If the user is NOT a 'system manager' or 'security manager'
			' read the table permissions from the server.
			sSQL = "exec spASRIntAllTablePermissions '" & Replace(gsActualLogin, "'", "''") & "'"
			rsPermissions = New Recordset
			rsPermissions.Open(sSQL, gADOCon, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly, CommandTypeEnum.adCmdText)

			colTablePermissions = New List(Of TablePermission)
			Do While Not rsPermissions.EOF
				objTablePermission = New TablePermission()
				objTablePermission.Name = rsPermissions.Fields("Name").Value
				objTablePermission.Action = rsPermissions.Fields("Action").Value
				objTablePermission.TableID = rsPermissions.Fields("TableID").Value
				colTablePermissions.Add(objTablePermission)
				rsPermissions.MoveNext()
			Loop
			rsPermissions.Close()
			'UPGRADE_NOTE: Object rsPermissions may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsPermissions = Nothing

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
			rsInfo = New Recordset
			gADOCon.CursorLocation = CursorLocationEnum.adUseClient
			rsInfo.Open("spASRIntGetColumnsFromTablesAndViews", gADOCon, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly, CommandTypeEnum.adCmdStoredProc)



			Do While Not rsInfo.EOF
				' If the current column's collection is NOT already instantiated, instantiate it.
				If sLastTableView <> rsInfo.Fields("tableviewname").Value Then
					sLastTableView = rsInfo.Fields("tableviewname").Value
					objColumnPrivileges = New CColumnPrivileges
					objColumnPrivileges.Tag = rsInfo.Fields("tableviewname").Value
					gcolColumnPrivilegesCollection.Add(objColumnPrivileges, rsInfo.Fields("tableviewname").Value)
				End If

				' JPD20020926 Fault 3980
				sColumnName = rsInfo.Fields("ColumnName").Value
				If Not objColumnPrivileges.IsValid(sColumnName) Then
					' Add the column object to the collection.
					' If the current user is a system/security maneger then set column privileges to TRUE,
					' else set them to FALSE.
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					objColumnPrivileges.Add(fSysSecManager, fSysSecManager, sColumnName, rsInfo.Fields("ColumnType").Value, rsInfo.Fields("DataType").Value, rsInfo.Fields("ColumnID").Value, IIf(IsDBNull(rsInfo.Fields("UniqueCheckType").Value), False, rsInfo.Fields("UniqueCheckType").Value <> 0))

				End If

				rsInfo.MoveNext()
			Loop
			rsInfo.Close()
			'UPGRADE_NOTE: Object rsInfo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsInfo = Nothing

			' If the current user is not a system/security manager then read the column permissions from SQL.
			If Not fSysSecManager Then

				sLastTableView = ""

				Dim rsInfo2 = clsDataAccess.GetDataTable("spASRIntGetColumnPermissions", "SourceList", aryRealSource)
				For Each objRow In rsInfo2.Rows

					If sLastTableView <> objRow("tableviewname") Then
						sLastTableView = objRow("tableviewname")

						objTableView = gcoTablePrivileges.FindRealSource(objRow("tableviewname"))
						If objTableView.IsTable Then
							sTableViewName = objTableView.TableName
						Else
							sTableViewName = objRow("tableviewname")
						End If

						objColumnPrivileges = gcolColumnPrivilegesCollection.Item(sTableViewName)
					End If

					If objRow("Action") = 193 Then
						objColumnPrivileges.Item(objRow("ColumnName")).AllowSelect = objRow("Permission")
					Else
						objColumnPrivileges.Item(objRow("ColumnName")).AllowUpdate = objRow("Permission")
					End If

				Next
			End If
		End If

		' Restore original cursor location
		gADOCon.CursorLocation = iOriginalCursorLocation

	End Sub


	Public Function GetColumnPrivileges(ByRef psTableViewName As String) As CColumnPrivileges
		' Return the column privileges collection for the given table.
		On Error GoTo ErrorTrap

		Dim fOK As Boolean
		Dim iLoop As Short
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
				GetColumnPrivileges = gcolColumnPrivilegesCollection.Item(iLoop)
				Exit Function
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