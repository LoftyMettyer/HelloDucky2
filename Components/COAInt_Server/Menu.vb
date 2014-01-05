Option Strict Off
Option Explicit On

Imports System.Collections.ObjectModel
Imports System.Collections.Generic
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.Metadata

Public Class Menu

	Public Function GetHistoryScreens() As Collection(Of Structures.HistoryScreen)
		' Return an array of information that can be used to format the History tables menu for the current user.
		' The recordset contains a row for each primary table in the HR Pro database.

		Dim sSQL As String
		Dim rsTableScreens As DataTable
		Dim HistoryScreensCollection As New Collection(Of Structures.HistoryScreen)
		Dim HistoryScreensList As New List(Of Structures.HistoryScreen)	'Used for sorting the menu items
		Dim objHistory As Structures.HistoryScreen

		sSQL = "SELECT ASRSysTables.tableName AS [childTableName], childScreens.tableID AS [childTableID], childScreens.screenID AS [childScreenID], childScreens.name AS [childScreenName], parentScreen.screenid AS [parentScreenID] FROM ASRSysScreens parentScreen INNER JOIN ASRSysHistoryScreens ON parentScreen.screenID = ASRSysHistoryScreens.parentScreenID INNER JOIN ASRSysScreens childScreens ON ASRSysHistoryScreens.historyScreenID = childScreens.screenID INNER JOIN ASRSysTables ON childScreens.tableID = ASRSysTables.tableID WHERE childScreens.quickEntry = 0 ORDER BY parentScreen.screenid, ASRSysTables.tableName DESC, childScreens.Name DESC"
		rsTableScreens = clsDataAccess.GetDataTable(sSQL, CommandType.Text)

		For Each objRow In rsTableScreens.Rows
			If gcoTablePrivileges.Item(objRow("childTableName")).AllowSelect Then
				objHistory = New Structures.HistoryScreen
				objHistory.parentScreenID = objRow("parentScreenID")
				objHistory.childTableID = objRow("childTableID")
				objHistory.childTableName = objRow("childTableName")
				objHistory.childScreenID = objRow("childScreenID")
				objHistory.childScreenName = Replace(objRow("childScreenName"), "&", "&&")
				HistoryScreensList.Add(objHistory)
			End If

		Next

		'Sort the menu items; note that I'm sorting in descending order because the code that actually creates the menu adds the item in inverse order (don't ask me why),
		'so the net effect is that the menu is sorted
		HistoryScreensList.Sort(Function(item1 As Structures.HistoryScreen, item2 As Structures.HistoryScreen)
																Return item1.childScreenName > item2.childScreenName
															End Function)

		For Each objHistory In HistoryScreensList
			HistoryScreensCollection.Add(objHistory)
		Next

		Return HistoryScreensCollection

	End Function

	Public Function GetPrimaryTableMenu() As Object

		On Error GoTo ErrorTrap

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
		Dim iNextIndex As Short
		Dim iTotalViewScreenCount As Short
		Dim sSQL As String
		Dim sViewList As String
		Dim rsViews As DataTable
		Dim rsViewScreen As DataTable
		Dim rsTables As DataTable
		Dim rsTableScreen As DataTable
		Dim avTableInfo(,) As Object
		Dim objTableView As TablePrivilege

		ReDim avTableInfo(12, 0)
		' Index 1 = table ID
		' Index 2 = table name
		' Index 3 = table screen count
		' Index 4 = table screen ID
		' Index 5 = table readable
		' Index 6 = table view count
		' Index 7 = view ID
		' Index 8 = view name
		' Index 9 = view screen count
		' Index 10 = view screen ID
		' Index 11 = table screen picture ID
		' Index 12 = view screen picture ID

		' Get a recordset of the primary tables in the database.
		sSQL = "SELECT ASRSysTables.tableID, ASRSysTables.tableName AS [tablename], COUNT(DISTINCT ASRSysScreens.name) AS tableScreenCount FROM ASRSysTables INNER JOIN ASRSysScreens ON ASRSysTables.tableID = ASRSysScreens.tableID AND ((ASRSysScreens.ssIntranet IS null) OR (ASRSysScreens.ssIntranet = 0)) AND ((ASRSysScreens.quickEntry IS null) OR (ASRSysScreens.quickEntry = 0)) GROUP BY ASRSysTables.tableID, ASRSysTables.tableName, ASRSysTables.tableType HAVING ASRSysTables.tableType = 1 ORDER BY ASRSysTables.tableName DESC"
		rsTables = clsDataAccess.GetDataTable(sSQL, CommandType.Text)

		For Each objRow In rsTables.Rows
	
			' Initialise an entry into our array of table info for each primary table.
			iNextIndex = UBound(avTableInfo, 2) + 1
			ReDim Preserve avTableInfo(12, iNextIndex)

			'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(1, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			avTableInfo(1, iNextIndex) = objRow("TableID")
			'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(2, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			avTableInfo(2, iNextIndex) = objRow("TableName")
			'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(3, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			avTableInfo(3, iNextIndex) = objRow("tableScreenCount")
			'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(4, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			avTableInfo(4, iNextIndex) = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(5, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			avTableInfo(5, iNextIndex) = gcoTablePrivileges.Item(objRow("TableName")).AllowSelect
			'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(6, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			avTableInfo(6, iNextIndex) = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(7, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			avTableInfo(7, iNextIndex) = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(8, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			avTableInfo(8, iNextIndex) = ""
			'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(9, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			avTableInfo(9, iNextIndex) = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(10, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			avTableInfo(10, iNextIndex) = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(11, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			avTableInfo(11, iNextIndex) = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(12, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			avTableInfo(12, iNextIndex) = 0

			iTotalViewScreenCount = 0

			' Create a list of the current user's permitted views on the current table.
			sViewList = "0"
			For Each objTableView In gcoTablePrivileges.Collection
				If Not (objTableView.IsTable) And (objTableView.TableID = objRow("TableID")) And (objTableView.AllowSelect) Then

					sViewList = sViewList & ", " & Trim(Str(objTableView.ViewID))
				End If
			Next objTableView
			'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objTableView = Nothing

			' Get view information for the current table.
			sSQL = "SELECT ASRSysViews.viewName, COUNT (ASRSysViewScreens.ScreenID) AS viewScreenCount FROM ASRSysViews LEFT OUTER JOIN ASRSysViewScreens ON ASRSysViews.viewID = ASRSysViewScreens.viewID GROUP BY ASRSysViews.viewName, ASRSysViews.viewID HAVING ASRSysViews.viewID IN (" & sViewList & ")"
			rsViews = clsDataAccess.GetDataTable(sSQL, CommandType.Text)
			For Each objViewRow In rsViews.Rows
				If objViewRow("viewScreenCount") > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(6, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					avTableInfo(6, iNextIndex) = avTableInfo(6, iNextIndex) + 1
					iTotalViewScreenCount = iTotalViewScreenCount + objViewRow("viewScreenCount")
				End If

			Next

			' Get view screen info if required.
			'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(6, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If avTableInfo(6, iNextIndex) > 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(9, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				avTableInfo(9, iNextIndex) = iTotalViewScreenCount

				'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(3, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(5, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (iTotalViewScreenCount = 1) And ((avTableInfo(5, iNextIndex) = False) Or (avTableInfo(3, iNextIndex) = 0)) Then

					sSQL = "SELECT ASRSysViews.viewID, ASRSysViews.viewName, ASRSysViewScreens.screenID, ASRSysScreens.pictureID FROM ASRSysViews INNER JOIN ASRSysViewScreens ON ASRSysViews.viewID = ASRSysViewScreens.viewID INNER JOIN ASRSysScreens ON ASRSysViewScreens.screenID = ASRSysScreens.screenID WHERE ASRSysViews.viewid IN (" & sViewList & ")"
					rsViewScreen = clsDataAccess.GetDataTable(sSQL, CommandType.Text)

					If rsViewScreen.Rows.Count > 0 Then

						Dim objViewScreenRow = rsViewScreen.Rows(0)
						'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(7, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						avTableInfo(7, iNextIndex) = objViewScreenRow("ViewID")
						'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(8, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						avTableInfo(8, iNextIndex) = objViewScreenRow("ViewName")
						'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(10, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						avTableInfo(10, iNextIndex) = objViewScreenRow("ScreenID")
						'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(12, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						avTableInfo(12, iNextIndex) = objViewScreenRow("pictureID")
					End If
				End If
			End If

			'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(5, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(3, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (avTableInfo(3, iNextIndex) = 1) And (iTotalViewScreenCount = 0) And (avTableInfo(5, iNextIndex) = True) Then
				sSQL = "SELECT ASRSysScreens.screenID, ASRSysScreens.pictureID FROM ASRSysScreens WHERE ASRSysScreens.tableID = " & Trim(Str(objRow("TableID"))) & "   AND ((ASRSysScreens.ssIntranet IS null) OR (ASRSysScreens.ssIntranet = 0))"

				rsTableScreen = clsDataAccess.GetDataTable(sSQL, CommandType.Text)
				If rsTableScreen.Rows.Count > 0 Then
					Dim objTableScreenRow = rsTableScreen.Rows(0)
					'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(4, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					avTableInfo(4, iNextIndex) = objTableScreenRow("ScreenID")
					'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(11, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					avTableInfo(11, iNextIndex) = objTableScreenRow("pictureID")
				End If
			End If

		Next

		'UPGRADE_WARNING: Couldn't resolve default property of object GetPrimaryTableMenu. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Return VB6.CopyArray(avTableInfo)

TidyUpAndExit:
		Exit Function
ErrorTrap:

	End Function

	Public Function GetPrimaryTableSubMenu(ByVal plngTableID As Integer) As Object
		' Return an array of information that can be used to format the given table's sub-menu
		' on the Database menu for the current user.
		' The array contains a row for each screen and view screen.
		'
		' For each primary table the following information is given :
		' screenID    ID of the screen
		' screenName  Name of the screen
		' viewID      ID of the view
		' viewName    Name of the view
		' screenPictureID ID of the screen's icon
		Dim iNextIndex As Short
		Dim sSQL As String
		Dim sViewList As String
		Dim rsScreens As DataTable
		Dim avScreenInfo(,) As Object
		Dim objTableView As TablePrivilege
		Dim sTableName As String

		' Create an array with records for each screen for each permitted view on the primary table.
		ReDim avScreenInfo(5, 0)

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
		rsScreens = clsDataAccess.GetDataTable(sSQL, CommandType.Text)

		For Each objRow In rsScreens.Rows
			iNextIndex = UBound(avScreenInfo, 2) + 1
			ReDim Preserve avScreenInfo(5, iNextIndex)
			'UPGRADE_WARNING: Couldn't resolve default property of object avScreenInfo(1, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			avScreenInfo(1, iNextIndex) = objRow("ScreenID")
			'UPGRADE_WARNING: Couldn't resolve default property of object avScreenInfo(2, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			avScreenInfo(2, iNextIndex) = Replace(objRow("Name"), "&", "&&")
			'UPGRADE_WARNING: Couldn't resolve default property of object avScreenInfo(3, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			avScreenInfo(3, iNextIndex) = objRow("ViewID")
			'UPGRADE_WARNING: Couldn't resolve default property of object avScreenInfo(4, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			avScreenInfo(4, iNextIndex) = objRow("ViewName")
			'UPGRADE_WARNING: Couldn't resolve default property of object avScreenInfo(5, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			avScreenInfo(5, iNextIndex) = objRow("pictureID")

		Next

		'UPGRADE_WARNING: Couldn't resolve default property of object GetPrimaryTableSubMenu. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Return VB6.CopyArray(avScreenInfo)

	End Function

	Public Function GetTableScreens() As Collection(Of Structures.TableScreen)
		' Return an array of information that can be used to format the Lookup tables menu for the current user.
		' The recordset contains a row for each primary table in the HR Pro database.

		Dim sSQL As String
		Dim rsTableScreens As DataTable
		Dim TableInfo As New Collection(Of Structures.TableScreen)
		Dim objTableScreen As Structures.TableScreen

		sSQL = String.Format("SELECT ASRSysTables.tableID, ASRSysTables.tableName, ASRSysScreens.screenID, v.HideFromMenu AS [result]" & _
			" FROM ASRSysTables" & _
			" INNER JOIN ASRSysScreens ON ASRSysTables.tableID = ASRSysScreens.tableID" & _
			" INNER JOIN ASRSysViewMenuPermissions v ON v.TableName = ASRSysTables.tablename AND v.groupName = '{0}'" & _
			" WHERE ASRSysTables.tableType = {1}" & _
			" AND ((ASRSysScreens.ssIntranet IS null) OR (ASRSysScreens.ssIntranet = 0)) AND ((ASRSysScreens.quickEntry IS null)" & _
			" OR (ASRSysScreens.quickEntry = 0))" & _
			" ORDER BY ASRSysTables.tableName DESC", gsUserGroup, Trim(Str(TableTypes.tabLookup)))

		rsTableScreens = clsDataAccess.GetDataTable(sSQL, CommandType.Text)
		For Each objRow In rsTableScreens.Rows
			If objRow("Result") = 0 Then
				objTableScreen = New Structures.TableScreen
				objTableScreen.TableID = objRow("TableID")
				objTableScreen.TableName = objRow("TableName")
				objTableScreen.ScreenID = objRow("ScreenID")
				TableInfo.Add(objTableScreen)
			End If

		Next

		Return TableInfo

	End Function

	Public Function GetQuickEntryScreens() As Object
		Dim sSQL As String
		Dim rsScreens As DataTable
		Dim avTableInfo(,) As Object
		Dim iNextIndex As Short

		ReDim avTableInfo(3, 0)
		' Index 1 = table ID
		' Index 2 = screen name
		' Index 3 = table screen ID

		sSQL = "SELECT ASRSysScreens.screenID, ASRSysScreens.name, UPPER(ASRSysTables.tableName) AS [tablename], ASRSysTables.tableID FROM ASRSysScreens INNER JOIN ASRSysTables ON ASRSysScreens.tableID = ASRSysTables.tableID WHERE ASRSysScreens.quickEntry = 1"
		rsScreens = clsDataAccess.GetDataTable(sSQL, CommandType.Text)

		For Each objRow In rsScreens.Rows
			'First see if we have privileges to see this table
			If gcoTablePrivileges.Item(objRow("TableName")).AllowSelect Then

				' Check that the current user has 'select' permission on at least one parent table,
				' or at least one view of one parent table referenced by the quick entry screen.
				If ViewQuickEntry(objRow("ScreenID")) Then
					iNextIndex = UBound(avTableInfo, 2) + 1
					ReDim Preserve avTableInfo(3, iNextIndex)

					'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(1, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					avTableInfo(1, iNextIndex) = objRow("TableID")
					'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(2, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					avTableInfo(2, iNextIndex) = Replace(objRow("Name"), "&", "&&")
					'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(3, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					avTableInfo(3, iNextIndex) = objRow("ScreenID")
				End If
			End If
		Next

		'UPGRADE_NOTE: Object rsScreens may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsScreens = Nothing

		'UPGRADE_WARNING: Couldn't resolve default property of object GetQuickEntryScreens. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Return VB6.CopyArray(avTableInfo)

	End Function

	Private Function ViewQuickEntry(ByVal plngScreenID As Integer) As Boolean
		' Return TRUE if the current user can see at least one parent table (or view of a parent table)
		' of given quick view screen.
		On Error GoTo ErrorTrap

		Dim rsTables As DataTable
		Dim rsViews As DataTable

		' Get the list of parent tables used in the quick entry screen.
		rsTables = GetQuickEntryTables(plngScreenID)

		With rsTables
			If .Rows.Count > 0 Then
				Return True
			End If

			' Loop through parent tables, seeing if we have select permissions on these tables.
			For Each objTableRow In rsTables.Rows

				' Check if the current user has 'select' permission on the given table.
				If gcoTablePrivileges.Item(UCase(objTableRow("TableName"))).AllowSelect Then
					Return True
				Else
					' No select permissions, can we use a view instead ???
					rsViews = GetQuickEntryViews(objTableRow("TableID"))

					Dim objTableView As TablePrivilege

					'Loop through the views, and see if we have permission on these
					For Each objViewRow In rsViews.Rows

						objTableView = gcoTablePrivileges.Item(UCase(objViewRow("ViewName")))

						If objTableView.AllowSelect Then
							'We have a view we can use, let's get outta here
							Return True
						End If

					Next
				End If

			Next
		End With

		Return False

ErrorTrap:
		If Err.Number = 457 Then
			Resume Next
		End If

	End Function

	Private Function GetQuickEntryTables(ByVal plngScreenID As Integer) As DataTable
		' Return a recordset of all the table id's of the controls which aren't related to the base table
		' in the given screen.
		Dim sSQL As String = "exec sp_ASRGetQuickEntryTables " & plngScreenID
		Return clsDataAccess.GetDataTable(sSQL, CommandType.Text)

	End Function

	Private Function GetQuickEntryViews(ByVal plngTableID As Integer) As DataTable
		' Return a recordset of the user defined views on the given table.
		Dim sSQL As String = "SELECT viewID, viewName FROM ASRSysViews WHERE viewTableID = " & Trim(Str(plngTableID))
		Return clsDataAccess.GetDataTable(sSQL, CommandType.Text)

	End Function

End Class