Option Strict Off
Option Explicit On

Imports System.Collections.ObjectModel
Imports HR.Intranet.Server.Enums

Module modPermissions

  Public Sub SetupTablesCollection()
        ' Read the list of tables the current user has permission to see.
        Dim fSysSecManager As Boolean
				Dim lngRoleID As Integer
        Dim lngChildViewID As Integer
        Dim sSQL As String
        Dim sRealSourceList As String
        Dim sTableViewName As String
        Dim rsInfo As ADODB.Recordset
        Dim rsTables As ADODB.Recordset
        Dim rsViews As ADODB.Recordset
        Dim rsPermissions As ADODB.Recordset
        Dim objTableView As CTablePrivilege
        Dim objColumnPrivileges As CColumnPrivileges
        Dim avChildViews(,) As Object
        Dim lngNextIndex As Integer
        'Dim sRoleName As String
        Dim iTemp As Short
        Dim avTablePermissions(,) As Object
        Dim iLoop2 As Short
        Dim sTableName As String
        Dim sLastTableView As String
        Dim sColumnName As String
        Dim iAction As Short
        Dim iOriginalCursorLocation As Short
        Dim objItem As CTablePrivilege
        Dim lngBaseTableId As Long

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
				gcoTablePrivileges = New Collection(Of CTablePrivilege)()

				sSQL = "SELECT system_user AS [name]"
				rsInfo = New ADODB.Recordset
				rsInfo.Open(sSQL, gADOCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
				datGeneral.Username = rsInfo.Fields("Name").Value
				rsInfo.Close()
				'UPGRADE_NOTE: Object rsInfo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				rsInfo = Nothing

				' Create an array of child view IDs and their associated table names.
				' Column 1 - child view ID
				' Column 2 - associated table name
				' Column 3 - 0=OR, 1=AND
				ReDim avChildViews(3, 0)

				'sSQL = "SELECT ASRSysChildViews2.childViewID, ASRSysTables.tableName, ASRSysChildViews2.type" & _
				'" FROM ASRSysChildViews2" & _
				'" INNER JOIN ASRSysTables ON ASRSysChildViews2.tableID = ASRSysTables.tableID" & _
				'" WHERE ASRSysChildViews2.role = '" & Replace(sRoleName, "'", "''") & "'"
				sSQL = "SELECT ASRSysChildViews2.childViewID, ASRSysTables.tableName, ASRSysChildViews2.type" & " FROM ASRSysChildViews2" & " INNER JOIN ASRSysTables ON ASRSysChildViews2.tableID = ASRSysTables.tableID" & " WHERE ASRSysChildViews2.role = '" & Replace(gsUserGroup, "'", "''") & "'"

				rsInfo = New ADODB.Recordset
				rsInfo.Open(sSQL, gADOCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)

				Do While Not rsInfo.EOF
						lngNextIndex = UBound(avChildViews, 2) + 1
						ReDim Preserve avChildViews(3, lngNextIndex)
						'UPGRADE_WARNING: Couldn't resolve default property of object avChildViews(1, lngNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						avChildViews(1, lngNextIndex) = rsInfo.Fields("childViewID").Value
						'UPGRADE_WARNING: Couldn't resolve default property of object avChildViews(2, lngNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						avChildViews(2, lngNextIndex) = rsInfo.Fields("TableName").Value
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						'UPGRADE_WARNING: Couldn't resolve default property of object avChildViews(3, lngNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						avChildViews(3, lngNextIndex) = IIf(IsDBNull(rsInfo.Fields("Type").Value), 0, rsInfo.Fields("Type").Value)

						rsInfo.MoveNext()
				Loop
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
					objItem = New CTablePrivilege()
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

								objItem = New CTablePrivilege()
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

						sSQL = "SELECT tableid, childViewID" & " FROM ASRSysChildViews2" & " WHERE role = '" & Replace(gsUserGroup, "'", "''") & "'"
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

						' JPD20020926 Fault 3980 - Suffered the 'Connection is busy with results for another hstmt'
						' fault when running ADO queries whilst looping through the permissions recordset.
						' Not sure why, but reading the permissions into an array, closing the permissions recordset,
						' and then running ADO queries whilst looping through the permissions array solved the problem.
						ReDim avTablePermissions(3, 0)
						Do While Not rsPermissions.EOF
								ReDim Preserve avTablePermissions(3, UBound(avTablePermissions, 2) + 1)
								'UPGRADE_WARNING: Couldn't resolve default property of object avTablePermissions(1, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								avTablePermissions(1, UBound(avTablePermissions, 2)) = rsPermissions.Fields("Name").Value
								'UPGRADE_WARNING: Couldn't resolve default property of object avTablePermissions(2, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								avTablePermissions(2, UBound(avTablePermissions, 2)) = rsPermissions.Fields("Action").Value

								avTablePermissions(3, UBound(avTablePermissions, 2)) = rsPermissions.Fields("TableID").Value

								rsPermissions.MoveNext()
						Loop
						rsPermissions.Close()
						'UPGRADE_NOTE: Object rsPermissions may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
						rsPermissions = Nothing


						For iLoop2 = 1 To UBound(avTablePermissions, 2)
								'UPGRADE_WARNING: Couldn't resolve default property of object avTablePermissions(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sTableName = CStr(avTablePermissions(1, iLoop2))
								'UPGRADE_WARNING: Couldn't resolve default property of object avTablePermissions(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								iAction = CShort(avTablePermissions(2, iLoop2))
								lngBaseTableId = CLng(avTablePermissions(3, iLoop2))

								'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
								objTableView = Nothing

								If UCase(Left(sTableName, 8)) = "ASRSYSCV" Then
										' Determine which table the child view is for.
										iTemp = InStr(sTableName, "#")
										lngChildViewID = Val(Mid(sTableName, 9, iTemp - 9))

										objTableView = gcoTablePrivileges.FindTableID(lngBaseTableId)

								Else
										If UCase(Left(sTableName, 6)) <> "ASRSYS" Then
												objTableView = gcoTablePrivileges.Item(sTableName)
										End If
								End If

								If Not objTableView Is Nothing Then
										objTableView.RealSource = sTableName

										Select Case iAction
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
						Next iLoop2
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
								' Get the SQL group id of the current user.
								' JPD20020809 Fault 3901
								sSQL = "SELECT gid" & " FROM sysusers" & " WHERE name = '" & Replace(gsUserGroup, "'", "''") & "'"
								'" WHERE name = current_user"
								rsInfo = New ADODB.Recordset
								rsInfo.Open(sSQL, gADOCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
								If Not (rsInfo.BOF And rsInfo.EOF) Then
										lngRoleID = rsInfo.Fields("gid").Value
								End If

								rsInfo.Close()
								'UPGRADE_NOTE: Object rsInfo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
								rsInfo = Nothing

								sSQL = "SELECT sysobjects.name AS tableViewName," & " syscolumns.name AS columnName," & " p.action," & " CASE p.protectType" & "   WHEN 205 THEN 1" & "   WHEN 204 THEN 1" & "   ELSE 0" & " END AS permission" & " FROM #SysProtects p" & " INNER JOIN sysobjects ON p.id = sysobjects.id" & " INNER JOIN syscolumns ON p.id = syscolumns.id" & " WHERE (p.action = 193 or p.action = 197)" & " AND syscolumns.name <> 'timestamp'" & " AND sysobjects.name in (" & sRealSourceList & ")" & " AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0" & " AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)" & " OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0" & " AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))" & " ORDER BY tableViewName"
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