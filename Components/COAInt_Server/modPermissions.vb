Option Strict Off
Option Explicit On

Imports System.Collections.ObjectModel
Imports System.Collections.Generic
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.Metadata
Imports HR.Intranet.Server.Structures

Module modPermissions

	Public Sub SetupTablesCollection()
		' Read the list of tables the current user has permission to see.
		Dim fSysSecManager As Boolean
		Dim sSQL As String
		Dim sRealSourceList As String
		Dim sTableViewName As String
		Dim rsInfo As ADODB.Recordset
		Dim rsViews As ADODB.Recordset
		Dim rsPermissions As ADODB.Recordset
		Dim objTableView As TablePrivilege
		Dim objColumnPrivileges As CColumnPrivileges
		Dim colTablePermissions As IList(Of TablePermission)
		Dim objTablePermissions As TablePermission

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
		rsInfo = New ADODB.Recordset
		rsInfo.Open(sSQL, gADOCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		datGeneral.Username = rsInfo.Fields("Name").Value
		rsInfo.Close()
		'UPGRADE_NOTE: Object rsInfo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsInfo = Nothing

		' Check if the user is a 'system manager' or 'security manager'.
		' If so then we can save time by applying all table permissions, instead of having to read them first.
		sSQL = "SELECT count(*) AS recCount FROM ASRSysGroupPermissions INNER JOIN ASRSysPermissionItems ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID" & " INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & " INNER JOIN sysusers a ON ASRSysGroupPermissions.groupName = a.name" & "   AND a.name = '" & gsUserGroup & "'" & " WHERE (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & " OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER')" & " AND ASRSysGroupPermissions.permitted = 1" & " AND ASRSysPermissionCategories.categorykey = 'MODULEACCESS'"

		rsInfo = New ADODB.Recordset
		rsInfo.Open(sSQL, gADOCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		fSysSecManager = (rsInfo.Fields("recCount").Value > 0)
		rsInfo.Close()
		'UPGRADE_NOTE: Object rsInfo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsInfo = Nothing

		' Initialise the collection with items for each TABLE in the system.
		For Each objTable In Tables
			objItem = New TablePrivilege()
			objItem.TableName = objTable.Name
			objItem.TableID = objTable.ID
			objItem.TableType = objTable.TableType
			objItem.DefaultOrderID = objTable.DefaultOrderID
			objItem.RecordDescriptionID = objTable.RecordDescExprID
			objItem.IsTable = True
			objItem.ViewID = 0
			objItem.ViewName = ""
			objItem.RealSource = objTable.Name
			gcoTablePrivileges.Add(objItem)
		Next


		' Initialise the collection with items for each VIEW in the system.
		rsViews = datGeneral.GetAllViews
		With rsViews
			Do While Not .EOF
				'	objTableView = gcoTablePrivileges.Add(.Fields("TableName").Value, .Fields("TableID").Value, .Fields("TableType").Value, .Fields("DefaultOrderID").Value, .Fields("RecordDescExprID").Value, False, .Fields("ViewID").Value, .Fields("ViewName").Value)

				objItem = New TablePrivilege()
				objItem.TableName = .Fields("TableName").Value
				objItem.TableID = .Fields("TableID").Value
				objItem.TableType = .Fields("TableType").Value
				objItem.DefaultOrderID = .Fields("DefaultOrderID").Value
				objItem.RecordDescriptionID = .Fields("RecordDescExprID").Value
				objItem.IsTable = False
				objItem.ViewID = .Fields("ViewID").Value
				objItem.ViewName = .Fields("ViewName").Value
				objItem.RealSource = .Fields("ViewName").Value

				gcoTablePrivileges.Add(objItem)


				.MoveNext()
			Loop
			.Close()
		End With
		'UPGRADE_NOTE: Object rsViews may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsViews = Nothing

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
			rsInfo = New ADODB.Recordset
			rsInfo.Open(sSQL, gADOCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)

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
			rsPermissions = New ADODB.Recordset
			rsPermissions.Open(sSQL, gADOCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)

			colTablePermissions = New List(Of TablePermission)
			Do While Not rsPermissions.EOF
				objTablePermissions = New TablePermission()
				objTablePermissions.Name = rsPermissions.Fields("Name").Value
				objTablePermissions.Action = rsPermissions.Fields("Action").Value
				objTablePermissions.TableID = rsPermissions.Fields("TableID").Value
				colTablePermissions.Add(objTablePermissions)
				rsPermissions.MoveNext()
			Loop
			rsPermissions.Close()
			'UPGRADE_NOTE: Object rsPermissions may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsPermissions = Nothing

			For Each objTablePermissions In colTablePermissions

				objTableView = Nothing

				If UCase(Left(objTablePermissions.Name, 8)) = "ASRSYSCV" Then
					' Determine which table the child view is for.
					objTableView = gcoTablePrivileges.FindTableID(objTablePermissions.TableID)

				Else
					If UCase(Left(objTablePermissions.Name, 6)) <> "ASRSYS" Then
						objTableView = gcoTablePrivileges.Item(objTablePermissions.Name)
					End If
				End If

				If Not objTableView Is Nothing Then
					objTableView.RealSource = objTablePermissions.Name

					Select Case objTablePermissions.Action
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
		sRealSourceList = ""
		For Each objTableView In gcoTablePrivileges.Collection
			If Len(objTableView.RealSource) > 0 Then
				sRealSourceList = sRealSourceList & IIf(Len(sRealSourceList) > 0, ", '", "'") & objTableView.RealSource & "'"
			End If
		Next objTableView
		'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objTableView = Nothing

		' JPD20030313 Don't need to recreate the columns collection if it already exists.
		If Not gcolColumnPrivilegesCollection Is Nothing Then
			Exit Sub
		End If

		If Len(sRealSourceList) > 0 Then
			' Instantiate  the Column Privileges collection if it does not already exist.
			If gcolColumnPrivilegesCollection Is Nothing Then
				gcolColumnPrivilegesCollection = New Collection
			End If

			' Get the list of all columns in all tables/views.
			rsInfo = New ADODB.Recordset
			gADOCon.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			rsInfo.Open("spASRIntGetColumnsFromTablesAndViews", gADOCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdStoredProc)

			Do While Not rsInfo.EOF
				' If the current column's collection is NOT already instantiated, instantiate it.
				If sLastTableView <> UCase(rsInfo.Fields("tableviewname").Value) Then
					sLastTableView = UCase(rsInfo.Fields("tableviewname").Value)
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

				sSQL = "SELECT sysobjects.name AS tableViewName, syscolumns.name AS columnName, p.action, CASE p.protectType WHEN 205 THEN 1 WHEN 204 THEN 1 ELSE 0 END AS permission" _
					& " FROM #SysProtects p INNER JOIN sysobjects ON p.id = sysobjects.id INNER JOIN syscolumns ON p.id = syscolumns.id WHERE (p.action = 193 or p.action = 197)" _
					& " AND syscolumns.name <> 'timestamp' AND sysobjects.name in (" & sRealSourceList & ") AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0" _
					& " AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0) OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0" _
					& " AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0)) ORDER BY tableViewName"
				rsInfo = New ADODB.Recordset
				rsInfo.Open(sSQL, gADOCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)

				sLastTableView = ""

				Do While Not rsInfo.EOF
					' Get the current column's table/view name.
					If sLastTableView <> UCase(rsInfo.Fields("tableviewname").Value) Then
						sLastTableView = UCase(rsInfo.Fields("tableviewname").Value)

						objTableView = gcoTablePrivileges.FindRealSource(rsInfo.Fields("tableviewname").Value)
						If objTableView.IsTable Then
							sTableViewName = objTableView.TableName
						Else
							sTableViewName = rsInfo.Fields("tableviewname").Value
						End If

						objColumnPrivileges = gcolColumnPrivilegesCollection.Item(sTableViewName)
					End If

					If rsInfo.Fields("Action").Value = 193 Then
						objColumnPrivileges.Item(rsInfo.Fields("ColumnName").Value).AllowSelect = rsInfo.Fields("Permission").Value
					Else
						objColumnPrivileges.Item(rsInfo.Fields("ColumnName").Value).AllowUpdate = rsInfo.Fields("Permission").Value
					End If

					rsInfo.MoveNext()
				Loop
				rsInfo.Close()
				'UPGRADE_NOTE: Object rsInfo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				rsInfo = Nothing
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