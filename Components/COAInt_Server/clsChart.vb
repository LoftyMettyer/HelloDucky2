Option Strict On
Option Explicit On

Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.Metadata
Imports HR.Intranet.Server.Interfaces

Public Class clsChart
	Inherits BaseForDMI
	Implements IChart

	Private mastrUDFsRequired() As String
	Private mstrRealSource As String
	Private mstrBaseTableRealSource As String
	Private mlngTableViews(,) As Integer
	Private mstrViews() As String
	Private mobjTableView As TablePrivilege
	Private mobjColumnPrivileges As CColumnPrivileges

	' Strings to hold the SQL statement
	Private mstrSQLSelect As String
	Private mstrSQLSelectColour As String
	Private mstrSQLFrom As String
	Private mstrSQLJoin As String
	Private mstrSQLWhere As String
	Private mstrSQLOrderBy As String
	Private mstrSQL As String
	Private mstrErrorString As String

	Public Function GetChartData(plngTableID As Integer, plngColumnID As Integer, plngFilterID As Integer,
																piAggregateType As Integer, piElementType As ElementType,
																plngTableID_2 As Integer, plngColumnID_2 As Integer, plngTableID_3 As Integer, plngColumnID_3 As Integer,
																plngSortOrderID As Integer, piSortDirection As Integer, plngChart_ColourID As Integer) As DataTable Implements IChart.GetChartData

		Dim fOK As Boolean
		Dim strTableName As String
		Dim strColumnName As String
		Dim lngTableID As Integer
		Dim lngColumnID As Integer
		Dim lngFilterID As Integer
		Dim iAggregateType As Short
		Dim iElementType As Short
		Dim lngSortOrderID As Integer
		Dim iSortDirection As Integer
		Dim lngColourID As Integer
		Dim strColourColumnName As String

		fOK = True

		lngTableID = plngTableID
		lngColumnID = plngColumnID
		lngFilterID = plngFilterID
		iAggregateType = CShort(piAggregateType)
		iElementType = CShort(piElementType)
		lngSortOrderID = plngSortOrderID
		iSortDirection = CShort(piSortDirection)
		lngColourID = plngChart_ColourID

		strTableName = GetTableName(lngTableID)
		strColumnName = GetColumnName(lngColumnID)
		strColourColumnName = GetColumnName(lngColourID)

		If fOK Then fOK = GenerateSQLSelect(lngTableID, strTableName, lngColumnID, strColumnName, False)
		If fOK And piElementType = 2 And lngColourID > 0 Then fOK = GenerateSQLSelect(lngTableID, strTableName, lngColourID, strColourColumnName, True)
		If fOK Then fOK = GenerateSQLFrom(strTableName)
		If fOK Then fOK = GenerateSQLJoin(lngTableID)
		If fOK Then fOK = GenerateSQLWhere(lngTableID, lngFilterID)
		If fOK Then fOK = GenerateSQLOrderBy(lngSortOrderID, iSortDirection)
		If fOK Then fOK = MergeSQLStrings(iAggregateType, iElementType)

		If Not fOK Then	' Probably got a select permission denied - no column access, so default the data...
			If piElementType = 4 Then
				mstrSQL = "SELECT 'No Data' AS [Aggregate]"
			Else
				mstrSQL = "SELECT 'No Access' AS [COLUMN], 'No Access' AS [Aggregate], '" & Str(ColorTranslator.ToOle(Color.White)) & "' AS [Colour]"
			End If
		End If

		' Execute the SQL and store in recordset
		Return DB.GetDataTable(mstrSQL, CommandType.Text)

	End Function


	Public Function MergeSQLStrings(iAggregateType As Short, iElementType As Short) As Boolean

		Dim pstrAggregate As String

		Select Case iAggregateType
			Case 1 ' sum
				pstrAggregate = "ISNULL(SUM(" & mstrSQLSelect & "),0) AS [Aggregate]"
			Case 2 '  Average
				pstrAggregate = "ISNULL(AVG(" & mstrSQLSelect & "),0) AS [Aggregate]"
			Case 3 '  Minimum
				pstrAggregate = "ISNULL(MIN(" & mstrSQLSelect & "),0) AS [Aggregate]"
			Case 4 '  Maximum
				pstrAggregate = "ISNULL(MAX(" & mstrSQLSelect & "),0) AS [Aggregate]"
			Case Else	' unknown or not set - default to count
				pstrAggregate = "ISNULL(COUNT(" & mstrSQLSelect & "),0) AS [Aggregate]"
		End Select

		If iElementType = 4 Then
			mstrSQL = "SELECT " & pstrAggregate & " FROM " & mstrSQLFrom & IIf(Len(mstrSQLJoin) = 0, "", " " & mstrSQLJoin).ToString() & IIf(Len(mstrSQLWhere) = 0, "", " " & mstrSQLWhere).ToString()
		Else
			mstrSQL = "SELECT " & mstrSQLSelect & " AS [COLUMN], " & pstrAggregate & IIf(mstrSQLSelectColour <> vbNullString, mstrSQLSelectColour & " AS [COLOUR] ", ", " & Str(ColorTranslator.ToOle(Color.White)).ToString() & " AS [COLOUR] ").ToString() & " FROM " & mstrSQLFrom & IIf(Len(mstrSQLJoin) = 0, "", " " & mstrSQLJoin).ToString() & IIf(Len(mstrSQLWhere) = 0, "", " " & mstrSQLWhere).ToString() & " GROUP BY " & mstrSQLSelect & mstrSQLSelectColour & mstrSQLOrderBy
		End If

		Return True

	End Function


	Public Function GenerateSQLSelect(lngTableID As Integer, strTableName As String, lngColumnID As Integer, strColumnName As String, pfColourFlag As Boolean) As Boolean

		Dim plngTempTableID As Integer
		Dim pstrTempTableName As String
		Dim pstrTempColumnName As String

		Dim pblnOK As Boolean
		Dim pblnColumnOK As Boolean
		Dim pblnNoSelect As Boolean
		Dim pblnFound As Boolean

		Dim pstrColumnList As String
		Dim pstrColumnCode As String
		Dim pstrSource As String
		Dim pintNextIndex As Integer

		Dim objTableView As TablePrivilege

		' Set flags with their starting values
		pblnOK = True
		pblnNoSelect = False

		Try

			ReDim mastrUDFsRequired(0)

			' JPD20030219 Fault 5068
			' Check the user has permission to read the base table.
			pblnOK = False
			For Each objTableView In gcoTablePrivileges.Collection
				If (objTableView.TableID = lngTableID) And (objTableView.AllowSelect) Then
					pblnOK = True
					Exit For
				End If
			Next objTableView
			'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objTableView = Nothing

			If Not pblnOK Then
				GenerateSQLSelect = False
				mstrErrorString = "You do not have permission to read the base table either directly or through any views."
				Exit Function
			End If

			' Start off the select statement
			If pfColourFlag = True Then
				mstrSQLSelectColour = ", "
			Else
				mstrSQLSelect = ""
			End If

			' Dimension an array of tables/views joined to the base table/view
			' Column 1 = 0 if this row is for a table, 1 if it is for a view
			' Column 2 = table/view ID
			' (should contain everything which needs to be joined to the base tbl/view)
			ReDim mlngTableViews(2, 0)

			' Load the temp variables
			plngTempTableID = lngTableID
			pstrTempTableName = strTableName
			pstrTempColumnName = strColumnName

			' Check permission on that column
			mobjColumnPrivileges = GetColumnPrivileges(pstrTempTableName)
			mstrRealSource = gcoTablePrivileges.Item(pstrTempTableName).RealSource
			pblnColumnOK = mobjColumnPrivileges.IsValid(pstrTempColumnName)

			If pblnColumnOK Then
				pblnColumnOK = mobjColumnPrivileges.Item(pstrTempColumnName).AllowSelect
			End If

			If pblnColumnOK Then

				' this column can be read direct from the tbl/view or from a parent table
				pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ",", "").ToString() & mstrRealSource & "." & Trim(pstrTempColumnName)

				' If the table isnt the base table (or its realsource) then
				' Check if it has already been added to the array. If not, add it.
				If plngTempTableID <> lngTableID Then
					pblnFound = False
					For pintNextIndex = 1 To UBound(mlngTableViews, 2)
						If mlngTableViews(1, pintNextIndex) = 0 And mlngTableViews(2, pintNextIndex) = plngTempTableID Then
							pblnFound = True
							Exit For
						End If
					Next pintNextIndex

					If Not pblnFound Then
						pintNextIndex = UBound(mlngTableViews, 2) + 1
						ReDim Preserve mlngTableViews(2, pintNextIndex)
						mlngTableViews(1, pintNextIndex) = 0
						mlngTableViews(2, pintNextIndex) = plngTempTableID
					End If
				End If
			Else

				' this column cannot be read direct. If its from a parent, try parent views
				' Loop thru the views on the table, seeing if any have read permis for the column

				ReDim mstrViews(0)
				For Each mobjTableView In gcoTablePrivileges.Collection
					If (Not mobjTableView.IsTable) And (mobjTableView.TableID = lngTableID) And (mobjTableView.AllowSelect) Then

						pstrSource = mobjTableView.ViewName
						mstrRealSource = gcoTablePrivileges.Item(pstrSource).RealSource

						' Get the column permission for the view
						mobjColumnPrivileges = GetColumnPrivileges(pstrSource)

						' If we can see the column from this view
						If mobjColumnPrivileges.IsValid(pstrTempColumnName) Then
							If mobjColumnPrivileges.Item(pstrTempColumnName).AllowSelect Then

								ReDim Preserve mstrViews(UBound(mstrViews) + 1)
								mstrViews(UBound(mstrViews)) = mobjTableView.ViewName

								' Check if view has already been added to the array
								pblnFound = False
								For pintNextIndex = 1 To UBound(mlngTableViews, 2)
									If mlngTableViews(1, pintNextIndex) = 1 And mlngTableViews(2, pintNextIndex) = mobjTableView.ViewID Then
										pblnFound = True
										Exit For
									End If
								Next pintNextIndex

								If Not pblnFound Then

									' View hasnt yet been added, so add it !
									pintNextIndex = UBound(mlngTableViews, 2) + 1
									ReDim Preserve mlngTableViews(2, pintNextIndex)
									mlngTableViews(1, pintNextIndex) = 1
									mlngTableViews(2, pintNextIndex) = mobjTableView.ViewID

								End If
							End If
						End If
					End If

				Next mobjTableView

				'UPGRADE_NOTE: Object mobjTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				mobjTableView = Nothing

				' Does the user have select permission thru ANY views ?
				If UBound(mstrViews) = 0 Then
					pblnNoSelect = True
				Else

					' Add the column to the column list
					pstrColumnCode = ""
					For pintNextIndex = 1 To UBound(mstrViews)
						'CHANGE TO COALESCE
						pstrColumnCode = pstrColumnCode & " WHEN NOT " & mstrViews(pintNextIndex) & "." & pstrTempColumnName & " IS NULL THEN " & mstrViews(pintNextIndex) & "." & pstrTempColumnName

					Next pintNextIndex

					If Len(pstrColumnCode) > 0 Then
						pstrColumnCode = "CASE " & pstrColumnCode & " ELSE NULL" & " END "
						pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ",", "").ToString() & pstrColumnCode
					End If

				End If

				' If we cant see a column, then get outta here
				If pblnNoSelect Then
					GenerateSQLSelect = False
					mstrErrorString = "You do not have permission to see the column '" & strColumnName & "' either directly or through any views."
					Exit Function
				End If


				If Not pblnOK Then
					GenerateSQLSelect = False
					Exit Function
				End If

			End If

			If pfColourFlag = True Then
				mstrSQLSelectColour = mstrSQLSelectColour & pstrColumnList
			Else
				mstrSQLSelect = mstrSQLSelect & pstrColumnList
			End If

		Catch ex As Exception
			mstrErrorString = "Error generating SQL Select statement." & vbNewLine & ex.Message
			Return False

		End Try

		Return True

	End Function

	Private Function GenerateSQLFrom(strTableName As String) As Boolean

		mstrSQLFrom = gcoTablePrivileges.Item(strTableName).RealSource
		Return True

	End Function

	Private Function GenerateSQLJoin(lngTableID As Integer) As Boolean

		' Purpose : Add the join strings for parent/child/views.
		'           Also adds filter clauses to the joins if used

		Dim pobjTableView As TablePrivilege
		Dim pintLoop As Integer

		Try

			' Get the base table real source
			mstrBaseTableRealSource = mstrSQLFrom

			' First, do the join for all the views etc...
			For pintLoop = 1 To UBound(mlngTableViews, 2)

				' Get the table/view object from the id stored in the array
				If mlngTableViews(1, pintLoop) = 0 Then
					pobjTableView = gcoTablePrivileges.FindTableID(mlngTableViews(2, pintLoop))
				Else
					pobjTableView = gcoTablePrivileges.FindViewID(mlngTableViews(2, pintLoop))
				End If

				If (pobjTableView.TableID = lngTableID) Then
					If (pobjTableView.ViewName <> mstrBaseTableRealSource) Then
						mstrSQLJoin = mstrSQLJoin & " LEFT OUTER JOIN " & pobjTableView.RealSource & " ON " & mstrBaseTableRealSource & ".ID = " & pobjTableView.RealSource & ".ID"
					End If
				Else
					If (pobjTableView.ViewName <> mstrBaseTableRealSource) Then
						mstrSQLJoin = mstrSQLJoin & " LEFT OUTER JOIN " & pobjTableView.RealSource & " ON " & mstrBaseTableRealSource & ".ID_" & CStr(pobjTableView.TableID) & " = " & pobjTableView.RealSource & ".ID"
					End If

				End If
			Next pintLoop

		Catch ex As Exception
			mstrErrorString = "Error in GenerateSQLJoin." & vbNewLine & ex.Message
			Return False

		End Try

		Return True

	End Function

	Private Function GenerateSQLOrderBy(plngSortOrderID As Integer, piSortDirection As Integer) As Boolean

		' Purpose : Returns order by string from the sort order array

		' get the sort order - this is stored as decimal, but represents a 4 digit binary value
		' e.g. 5 = 0101 in binary, which represents sort orders 1 & 3=desc, 2 = asc.
		' Only first digit (leftmost) used in single axis charting.
		' Digit 1 = Horizontal Data Sort order
		' Digit 2 = Vertical Data Sort order
		' Digit 3 = 'Sort by Aggregate' tickbox
		' Digit 4 = 'Sort by Aggregate' sort order

		Dim pstrBinaryString As String
		pstrBinaryString = DecToBin(plngSortOrderID, 4)

		If Mid(pstrBinaryString, 3, 1) = "1" Then	' The third switch is for 'Sort by Aggregate'
			piSortDirection = CInt(Right(pstrBinaryString, 1))
			mstrSQLOrderBy = "[AGGREGATE] " & IIf(piSortDirection = 0, "ASC", "DESC").ToString()
		Else
			piSortDirection = CInt(Mid(pstrBinaryString, 1, 1))
			mstrSQLOrderBy = "[COLUMN] " & IIf(piSortDirection = 0, "ASC", "DESC").ToString()
		End If

		If Len(mstrSQLOrderBy) > 0 Then mstrSQLOrderBy = " ORDER BY " & mstrSQLOrderBy

		Return True

	End Function


	Private Function GenerateSQLWhere(lngTableID As Integer, lngFilterID As Integer) As Boolean

		' Purpose : Generate the where clauses that cope with the joins
		'           NB Need to add the where clauses for filters/picklists etc

		Dim pintLoop As Integer
		Dim pobjTableView As TablePrivilege
		Dim blnOK As Boolean
		Dim strFilterIDs As String

		Try

			pobjTableView = gcoTablePrivileges.FindTableID(lngTableID)
			If pobjTableView.AllowSelect = False Then

				' First put the where clauses in for the joins...only if base table is a top level table
				If UCase(Left(mstrBaseTableRealSource, 6)) <> "ASRSYS" Then

					For pintLoop = 1 To UBound(mlngTableViews, 2)
						' Get the table/view object from the id stored in the array
						If mlngTableViews(1, pintLoop) = 0 Then
							pobjTableView = gcoTablePrivileges.FindTableID(mlngTableViews(2, pintLoop))
						Else
							pobjTableView = gcoTablePrivileges.FindViewID(mlngTableViews(2, pintLoop))
						End If

						' dont add where clause for the base/chil/p1/p2 TABLES...only add views here
						' JPD20030207 Fault 5034
						If (mlngTableViews(1, pintLoop) = 1) Then
							mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " OR ", " WHERE (").ToString() & mstrBaseTableRealSource & ".ID IN (SELECT ID FROM " & pobjTableView.RealSource & ")"
						End If

					Next pintLoop

					If Len(mstrSQLWhere) > 0 Then mstrSQLWhere = mstrSQLWhere & ")"

				End If

			End If


			If lngFilterID > 0 Then

				blnOK = FilteredIDs(lngFilterID, strFilterIDs, mastrUDFsRequired)

				If blnOK Then
					mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ").ToString() & mstrSQLFrom & ".ID IN (" & strFilterIDs & ")"
				Else
					' Permission denied on something in the filter.
					mstrErrorString = "You do not have permission to use the '" & General.GetFilterName(lngFilterID) & "' filter."
					Return False
				End If
			End If

		Catch ex As Exception
			mstrErrorString = "Error in GenerateSQLWhere." & vbNewLine & ex.Message
			Return False

		End Try

		Return True

	End Function

	Public Shadows Property SessionInfo As SessionInfo Implements IChart.SessionInfo
		Set(value As SessionInfo)
			MyBase.SessionInfo = value
		End Set
		Get
			Return MyBase.SessionInfo
		End Get
	End Property

End Class