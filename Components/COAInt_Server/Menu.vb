Option Strict Off
Option Explicit On

Imports System.Collections.ObjectModel
Imports System.Collections.Generic
Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.Metadata
Imports HR.Intranet.Server.Structures

Public Class Menu
	Inherits BaseForDMI

	Public Function GetHistoryScreens() As List(Of HistoryScreen)
		' Return an array of information that can be used to format the History tables menu for the current user.
		' The recordset contains a row for each primary table in the HR Pro database.

		Dim sSQL As String
		Dim rsTableScreens As DataTable
		Dim HistoryScreensList As New List(Of HistoryScreen)
		Dim objHistory As HistoryScreen

		sSQL = "SELECT ASRSysTables.tableName AS [childTableName], childScreens.tableID AS [childTableID], childScreens.screenID AS [childScreenID], childScreens.name AS [childScreenName], parentScreen.screenid AS [parentScreenID] FROM ASRSysScreens parentScreen INNER JOIN ASRSysHistoryScreens ON parentScreen.screenID = ASRSysHistoryScreens.parentScreenID INNER JOIN ASRSysScreens childScreens ON ASRSysHistoryScreens.historyScreenID = childScreens.screenID INNER JOIN ASRSysTables ON childScreens.tableID = ASRSysTables.tableID WHERE childScreens.quickEntry = 0 ORDER BY parentScreen.screenid, ASRSysTables.tableName DESC, childScreens.Name DESC"
		rsTableScreens = DB.GetDataTable(sSQL, CommandType.Text)

		For Each objRow As DataRow In rsTableScreens.Rows
			If gcoTablePrivileges.Item(objRow("childTableName").ToString()).AllowSelect Then
				objHistory = New HistoryScreen
				objHistory.parentScreenID = CInt(objRow("parentScreenID"))
				objHistory.childTableID = CInt(objRow("childTableID"))
				objHistory.childTableName = objRow("childTableName").ToString()
				objHistory.childScreenID = CInt(objRow("childScreenID"))
				objHistory.childScreenName = Replace(objRow("childScreenName").ToString(), "&", "&&")
				HistoryScreensList.Add(objHistory)
			End If

		Next

		'Sort the menu items; note that I'm sorting in descending order because the code that actually creates the menu adds the item in inverse order (don't ask me why),
		'so the net effect is that the menu is sorted
		HistoryScreensList.Sort(Function(item1 As HistoryScreen, item2 As HistoryScreen)
																Return item1.childScreenName > item2.childScreenName
															End Function)

		Return HistoryScreensList

	End Function

	Public Function GetPrimaryTableMenu() As List(Of MenuInfo)

		' Return an array of information that can be used to format the main Database menu for the current user.
		' The recordset contains a row for each primary table in the HR Pro database.
		' For each primary table the following information is given :
		'  tableID     ID of the primary table
		'  tableName   Name of the primary table
		'  tableScreenCount  Number of screens associated with the primary table
		'  tableScreenID   If the current user has SELECT permission on the primary table, and the primary table has only one screen associated with it,
		'        and (the current user does not have SELECT permission on any views on the primary table, or there are no screens associated with the view)
		'        then the ID of the one screen associated with the primary table is returned
		'  tableReadable   True if the current user has SELECT permission on the primary table
		'  tableViewCount    Number of views on the primary table that the current user has SELECT permission on
		'  viewID      If (the current user does not have SELECT permission on the primary table, or the primary table has no screens associated with it),
		'        and (the current user has SELECT permission on only one view on the primary table, and there are is only one screen associated with the view)
		'        then the ID of the view is returned
		'  viewName    If (the current user does not have SELECT permission on the primary table, or the primary table has no screens associated with it),
		'        and (the current user has SELECT permission on only one view on the primary table, and there are is only one screen associated with the view)
		'        then the name of the view is returned
		'  viewScreenCount Number of screens associated with the views
		'  viewScreenID    If (the current user does not have SELECT permission on the primary table, or the primary table has no screens associated with it),
		'        and (the current user has SELECT permission on only one view on the primary table, and there are is only one screen associated with the view)
		'        then the name of the screen associated with the view is returned
		'  tableScreenPictureID - ID of the table screen's icon
		'  viewScreenPictureID - ID of the view screen's icon
		'
		' If the recordset for a primary table has a non-zero value in the tableScreenID field then the table just requires a tool on the Database menu that calls up the given screen.
		' Else, if the recordset for a primary table has a non-zero value in the viewID field then the table just requires a tool on the Database menu that calls up the given screen for the permitted view.
		' Else, if the recordset for a primary table has a non-zero value in the viewScreenCount field OR (the tableReadable value is 1 AND the tableScreenCount is greater than 1) then
		'  the table requires a tool on the Database menu that calls up a sub-band of the collection of views/screens available for the primary table.
		' Else, the primary table should not appear on the menu.
		Dim iTotalViewScreenCount As Integer
		Dim sSQL As String
		Dim sViewList As String
		Dim rsViews As DataTable
		Dim rsViewScreen As DataTable
		Dim rsTables As DataTable
		Dim rsTableScreen As DataTable
		Dim objTableView As TablePrivilege

		Dim avTableInfo As New List(Of MenuInfo)
		Dim objMenuInfo As MenuInfo

		Try

			' Get a recordset of the primary tables in the database.
			sSQL = "SELECT ASRSysTables.tableID, ASRSysTables.tableName AS [tablename], COUNT(DISTINCT ASRSysScreens.name) AS tableScreenCount FROM ASRSysTables INNER JOIN ASRSysScreens ON ASRSysTables.tableID = ASRSysScreens.tableID AND ((ASRSysScreens.ssIntranet IS null) OR (ASRSysScreens.ssIntranet = 0)) AND ((ASRSysScreens.quickEntry IS null) OR (ASRSysScreens.quickEntry = 0)) GROUP BY ASRSysTables.tableID, ASRSysTables.tableName, ASRSysTables.tableType HAVING ASRSysTables.tableType = 1 ORDER BY ASRSysTables.tableName DESC"
			rsTables = DB.GetDataTable(sSQL, CommandType.Text)

			For Each objRow As DataRow In rsTables.Rows
				objMenuInfo = New MenuInfo()
				objMenuInfo.TableID = CInt(objRow("TableID"))
				objMenuInfo.TableName = objRow("TableName").ToString()
				objMenuInfo.TableScreenCount = CInt(objRow("tableScreenCount"))
				objMenuInfo.TableScreenID = 0
				objMenuInfo.TableReadable = gcoTablePrivileges.Item(objRow("TableName").ToString()).AllowSelect
				objMenuInfo.TableViewCount = 0
				objMenuInfo.ViewID = 0
				objMenuInfo.ViewName = ""
				objMenuInfo.ViewScreenCount = 0
				objMenuInfo.ViewScreenID = 0
				objMenuInfo.TableScreenPictureID = 0
				objMenuInfo.ViewScreenPictureID = 0

				iTotalViewScreenCount = 0

				' Create a list of the current user's permitted views on the current table.
				sViewList = "0"
				For Each objTableView In gcoTablePrivileges.Collection
					If Not (objTableView.IsTable) And (objTableView.TableID = CInt(objRow("TableID"))) And (objTableView.AllowSelect) Then

						sViewList = sViewList & ", " & Trim(Str(objTableView.ViewID))
					End If
				Next objTableView
				'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				objTableView = Nothing

				' Get view information for the current table.
				sSQL = "SELECT ASRSysViews.viewName, COUNT (ASRSysViewScreens.ScreenID) AS viewScreenCount FROM ASRSysViews LEFT OUTER JOIN ASRSysViewScreens ON ASRSysViews.viewID = ASRSysViewScreens.viewID GROUP BY ASRSysViews.viewName, ASRSysViews.viewID HAVING ASRSysViews.viewID IN (" & sViewList & ")"
				rsViews = DB.GetDataTable(sSQL, CommandType.Text)
				For Each objViewRow As DataRow In rsViews.Rows
					If CInt(objViewRow("viewScreenCount")) > 0 Then
						objMenuInfo.TableViewCount += 1
						iTotalViewScreenCount = iTotalViewScreenCount + CInt(objViewRow("viewScreenCount"))
					End If

				Next

				' Get view screen info if required.
				If objMenuInfo.TableViewCount > 0 Then
					objMenuInfo.ViewScreenCount = iTotalViewScreenCount

					If (iTotalViewScreenCount = 1) And (objMenuInfo.TableReadable = False Or objMenuInfo.TableScreenCount = 0) Then

						sSQL = "SELECT ASRSysViews.viewID, ASRSysViews.viewName, ASRSysViewScreens.screenID, ASRSysScreens.pictureID FROM ASRSysViews INNER JOIN ASRSysViewScreens ON ASRSysViews.viewID = ASRSysViewScreens.viewID INNER JOIN ASRSysScreens ON ASRSysViewScreens.screenID = ASRSysScreens.screenID WHERE ASRSysViews.viewid IN (" & sViewList & ")"
						rsViewScreen = DB.GetDataTable(sSQL, CommandType.Text)

						If rsViewScreen.Rows.Count > 0 Then

							Dim objViewScreenRow = rsViewScreen.Rows(0)
							objMenuInfo.ViewID = CInt(objViewScreenRow("ViewID"))
							objMenuInfo.ViewName = objViewScreenRow("ViewName").ToString()
							objMenuInfo.ViewScreenID = CInt(objViewScreenRow("ScreenID"))
							objMenuInfo.ViewScreenPictureID = CInt(objViewScreenRow("pictureID"))
						End If
					End If
				End If

				If (objMenuInfo.TableScreenCount = 1) And (iTotalViewScreenCount = 0) And objMenuInfo.TableReadable = True Then
					sSQL = "SELECT ASRSysScreens.screenID, ASRSysScreens.pictureID FROM ASRSysScreens WHERE ASRSysScreens.tableID = " & Trim(objRow("TableID").ToString()) & "   AND ((ASRSysScreens.ssIntranet IS null) OR (ASRSysScreens.ssIntranet = 0))"

					rsTableScreen = DB.GetDataTable(sSQL, CommandType.Text)
					If rsTableScreen.Rows.Count > 0 Then
						Dim objTableScreenRow = rsTableScreen.Rows(0)
						objMenuInfo.TableScreenID = CInt(objTableScreenRow("ScreenID"))
						objMenuInfo.TableScreenPictureID = CInt(objTableScreenRow("pictureID"))
					End If
				End If

				' Add sub items
				objMenuInfo.SubItems = GetPrimaryTableSubMenu(objMenuInfo.TableID)

				avTableInfo.Add(objMenuInfo)

			Next

		Catch ex As Exception
			Throw

		End Try

		Return avTableInfo

	End Function

	Private Function GetPrimaryTableSubMenu(ByVal plngTableID As Integer) As Collection(Of MenuInfo)
		' Return an array of information that can be used to format the given table's sub-menu
		' on the Database menu for the current user.
		' The array contains a row for each screen and view screen.

		Dim sSQL As String
		Dim sViewList As String
		Dim rsScreens As DataTable
		Dim avScreenInfo As New Collection(Of MenuInfo)
		Dim objSubMenu As MenuInfo
		Dim objTableView As TablePrivilege
		Dim sTableName As String

		' Create a list of the current user's permitted views on the current table.
		sViewList = "0"
		For Each objTableView In gcoTablePrivileges.Collection
			If Not (objTableView.IsTable) And (objTableView.TableID = plngTableID) And (objTableView.AllowSelect) Then
				sViewList = sViewList & ", " & Trim(Str(objTableView.ViewID))
			End If

			If (objTableView.IsTable) And (objTableView.TableID = plngTableID) Then
				sTableName = objTableView.TableName
			End If
		Next objTableView
		'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objTableView = Nothing

		sSQL = "SELECT ASRSysViewScreens.screenID, ASRSysScreens.name, ASRSysViews.viewID, ASRSysViews.ViewName, ASRSysScreens.PictureID FROM ASRSysViews INNER JOIN ASRSysViewScreens ON ASRSysViews.viewID = ASRSysViewScreens.viewID INNER JOIN ASRSysScreens ON ASRSysViewScreens.screenID = ASRSysScreens.screenID WHERE ASRSysViews.viewID IN (" & sViewList & ")"

		If gcoTablePrivileges.Item(sTableName).AllowSelect Then
			' The current user does have SELECT permission on the given table, so populate the array
			' table with records for each screen associated with the primary table.
			sSQL = sSQL & " UNION SELECT ASRSysScreens.screenID, ASRSysScreens.Name, 0 AS viewID, '' AS viewName, ASRSysScreens.pictureID FROM ASRSysScreens WHERE (ASRSysScreens.tableID = " & Trim(Str(plngTableID)) & ")" & " AND ((ASRSysScreens.ssIntranet IS null) OR (ASRSysScreens.ssIntranet = 0)) AND ((ASRSysScreens.quickEntry IS null) OR (ASRSysScreens.quickEntry = 0))"
		End If

		sSQL = sSQL & " ORDER BY ASRSysScreens.name DESC, viewName DESC"
		rsScreens = DB.GetDataTable(sSQL, CommandType.Text)

		For Each objRow As DataRow In rsScreens.Rows
			objSubMenu = New MenuInfo

			objSubMenu.TableScreenID = CInt(objRow("ScreenID"))
			objSubMenu.ScreenName = Replace(objRow("Name").ToString, "&", "&&")
			objSubMenu.ViewID = CInt(objRow("ViewID"))
			objSubMenu.ViewName = objRow("ViewName").ToString()
			objSubMenu.TableScreenPictureID = CInt(objRow("pictureID"))

			avScreenInfo.Add(objSubMenu)
		Next

		Return avScreenInfo

	End Function

	Public Function GetTableScreens() As List(Of TableScreen)
		' Return an array of information that can be used to format the Lookup tables menu for the current user.
		' The recordset contains a row for each primary table in the HR Pro database.

		Dim sSQL As String
		Dim rsTableScreens As DataTable
		Dim TableInfo As New List(Of TableScreen)
		Dim objTableScreen As TableScreen

		sSQL = String.Format("SELECT ASRSysTables.tableID, ASRSysTables.tableName, ASRSysScreens.screenID" & _
			" FROM ASRSysTables" & _
			" INNER JOIN ASRSysScreens ON ASRSysTables.tableID = ASRSysScreens.tableID" & _
			" LEFT JOIN ASRSysViewMenuPermissions v ON v.TableName = ASRSysTables.tablename AND v.groupName = '{0}'" & _
			" WHERE ASRSysTables.tableType = {1}" & _
			" AND ISNULL(v.HideFromMenu, 0) = 0" & _
			" AND ((ASRSysScreens.ssIntranet IS null) OR (ASRSysScreens.ssIntranet = 0)) AND ((ASRSysScreens.quickEntry IS null)" & _
			" OR (ASRSysScreens.quickEntry = 0))" & _
			" ORDER BY ASRSysTables.tableName DESC", _login.UserGroup, Trim(Str(TableTypes.tabLookup)))

		rsTableScreens = DB.GetDataTable(sSQL, CommandType.Text)
		For Each objRow As DataRow In rsTableScreens.Rows
			objTableScreen = New TableScreen With {
				.TableID = CInt(objRow("TableID")),
				.TableName = objRow("TableName").ToString(),
				.ScreenID = CInt(objRow("ScreenID"))}
			TableInfo.Add(objTableScreen)
		Next

		Return TableInfo

	End Function

	Public Function GetQuickEntryScreens() As List(Of MenuInfo)
		Dim sSQL As String
		Dim avTableInfo As New List(Of MenuInfo)
		Dim objMenuItem As MenuInfo

		sSQL = "SELECT ASRSysScreens.screenID, ASRSysScreens.name, UPPER(ASRSysTables.tableName) AS [tablename], ASRSysTables.tableID FROM ASRSysScreens INNER JOIN ASRSysTables ON ASRSysScreens.tableID = ASRSysTables.tableID WHERE ASRSysScreens.quickEntry = 1"
		Using rsScreens = DB.GetDataTable(sSQL, CommandType.Text)

			For Each objRow As DataRow In rsScreens.Rows
				'First see if we have privileges to see this table
				If gcoTablePrivileges.Item(objRow("TableName").ToString()).AllowSelect Then

					' Check that the current user has 'select' permission on at least one parent table,
					' or at least one view of one parent table referenced by the quick entry screen.
					If ViewQuickEntry(CInt(objRow("ScreenID"))) Then
						objMenuItem = New MenuInfo()

						objMenuItem.TableID = CInt(objRow("TableID"))
						objMenuItem.TableName = Replace(objRow("Name").ToString(), "&", "&&")
						objMenuItem.TableScreenID = CInt(objRow("ScreenID"))

						avTableInfo.Add(objMenuItem)
					End If
				End If
			Next
		End Using

		Return avTableInfo

	End Function

	Private Function ViewQuickEntry(plngScreenID As Integer) As Boolean
		' Return TRUE if the current user can see at least one parent table (or view of a parent table)
		' of given quick view screen.
		Dim rsTables As DataTable
		Dim rsViews As DataTable

		Try

			' Get the list of parent tables used in the quick entry screen.
			rsTables = GetQuickEntryTables(plngScreenID)

			With rsTables
				If .Rows.Count > 0 Then
					Return True
				End If

				' Loop through parent tables, seeing if we have select permissions on these tables.
				For Each objTableRow As DataRow In rsTables.Rows

					' Check if the current user has 'select' permission on the given table.
					If gcoTablePrivileges.Item(UCase(objTableRow("TableName").ToString())).AllowSelect Then
						Return True
					Else
						' No select permissions, can we use a view instead ???
						rsViews = GetQuickEntryViews(CInt(objTableRow("TableID")))

						Dim objTableView As TablePrivilege

						'Loop through the views, and see if we have permission on these
						For Each objViewRow As DataRow In rsViews.Rows

							objTableView = gcoTablePrivileges.Item(UCase(objViewRow("ViewName").ToString()))

							If objTableView.AllowSelect Then
								'We have a view we can use, let's get outta here
								Return True
							End If

						Next
					End If

				Next
			End With

		Catch ex As Exception
			Return False

		End Try

		Return False

	End Function

	Private Function GetQuickEntryTables(plngScreenID As Integer) As DataTable
		' Return a recordset of all the table id's of the controls which aren't related to the base table
		' in the given screen.
		Dim sSQL As String = "exec sp_ASRGetQuickEntryTables " & plngScreenID
		Return DB.GetDataTable(sSQL, CommandType.Text)

	End Function

	Private Function GetQuickEntryViews(plngTableID As Integer) As DataTable
		' Return a recordset of the user defined views on the given table.
		Dim sSQL As String = "SELECT viewID, viewName FROM ASRSysViews WHERE viewTableID = " & Trim(Str(plngTableID))
		Return DB.GetDataTable(sSQL, CommandType.Text)

	End Function

End Class