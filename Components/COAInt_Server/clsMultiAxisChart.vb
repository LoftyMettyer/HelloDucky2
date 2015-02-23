Option Strict On
Option Explicit On

Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.Metadata
Imports HR.Intranet.Server.Interfaces

Public Class clsMultiAxisChart
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

	Private mstrSQLGroupBy As String
	Private mstrSQLSelectVerticalID As String
	Private mstrSQLSelectHorizontalID As String
	Private mstrSQLFrom As String
	Private mstrSQLJoin As String
	Private mstrSQLWhere As String
	Private mstrSQLOrderBy As String
	Private mstrSQL As String
	Private mstrErrorString As String


	Private mlngBaseTableID As Integer
	Private mstrBaseTableName As String

	' Recordset to store legend data from sQL
	Private mrstChartLegendData As New DataTable

	Public Function GetChartData(plngTableID As Integer, plngColumnID As Integer, plngFilterID As Integer,
																piAggregateType As Integer, piElementType As ElementType,
																plngTableID_2 As Integer, plngColumnID_2 As Integer, plngTableID_3 As Integer, plngColumnID_3 As Integer,
																plngSortOrderID As Integer, piSortDirection As Integer, plngChart_ColourID As Integer) As DataTable Implements IChart.GetChartData


		Dim fOK As Boolean
		Dim strTableName As String = ""
		Dim strColumnName As String = ""
		Dim strTableName2 As String = ""
		Dim strColumnName2 As String = ""
		Dim strTableName3 As String = ""
		Dim strColumnName3 As String = ""
		Dim strColourColumnName As String = ""
		Dim lngTableID As Integer
		Dim lngColumnID As Integer
		Dim lngTableID2 As Integer
		Dim lngColumnID2 As Integer
		Dim lngTableID3 As Integer
		Dim lngColumnID3 As Integer
		Dim lngFilterID As Integer
		Dim iAggregateType As Short
		Dim iElementType As Short
		Dim lngColourColumnID As Integer

		Dim lngSortOrderID As Integer
		Dim iSortDirection As Integer

		fOK = True

		lngTableID = plngTableID
		lngColumnID = plngColumnID
		lngTableID2 = plngTableID_2
		lngColumnID2 = plngColumnID_2
		lngTableID3 = plngTableID_3
		lngColumnID3 = plngColumnID_3
		lngFilterID = plngFilterID
		iAggregateType = CShort(piAggregateType)
		iElementType = CShort(piElementType)
		lngColourColumnID = plngChart_ColourID

		lngSortOrderID = plngSortOrderID
		iSortDirection = CShort(piSortDirection)

		If lngTableID > 0 Then
			strTableName = GetTableName(lngTableID)
			strColumnName = GetColumnName(lngColumnID)
		End If

		If lngTableID2 > 0 Then
			strTableName2 = GetTableName(lngTableID2)
			strColumnName2 = GetColumnName(lngColumnID2)
		End If

		If lngTableID3 > 0 Then
			strTableName3 = GetTableName(lngTableID3)
			strColumnName3 = GetColumnName(lngColumnID3)
		End If

		strColourColumnName = GetColumnName(lngColourColumnID)

		If IsAChildOf(lngTableID, lngTableID2) = True Then
			If IsAChildOf(lngTableID, lngTableID3) = True Then
				' 1 is base
				mlngBaseTableID = lngTableID
			Else
				' 3 is base
				mlngBaseTableID = lngTableID3
			End If
		Else
			If IsAChildOf(lngTableID2, lngTableID3) = True Then
				' 2 is base
				mlngBaseTableID = lngTableID2
			Else
				' 3 is base
				mlngBaseTableID = lngTableID3
			End If
		End If

		mstrBaseTableName = GetTableName(CInt(mlngBaseTableID))

		' Fault HRPRO 1354 - Default column 3 name to 'ID' if no column is
		' set in the database and aggregate is count. This is for tables
		' that have no numeric columns - unable to specify the column in sysmgr
		' mod setup...
		If iAggregateType = 0 And strTableName3 <> vbNullString And strColumnName3 = vbNullString Then
			strColumnName3 = "ID"
		End If

		If fOK Then fOK = GenerateSQLSelect(lngTableID, strTableName, lngColumnID, strColumnName, lngTableID2, lngColumnID2, strTableName2, strColumnName2, lngTableID3, lngColumnID3, strTableName3, strColumnName3, iAggregateType, lngColourColumnID, strColourColumnName)
		If fOK Then GenerateSQLFrom(mstrBaseTableName)
		If fOK Then fOK = GenerateSQLJoin(mlngBaseTableID)
		If fOK Then fOK = GenerateSQLWhere(mlngBaseTableID, lngFilterID)
		If fOK Then fOK = GenerateSQLOrderBy(lngSortOrderID, iSortDirection)
		If fOK Then MergeSQLStrings()
		If fOK Then fOK = SQLSelectVerticalID(lngColumnID2)
		If fOK Then fOK = SQLSelectHorizontalID(iSortDirection, lngColumnID)

		If Not fOK Then	' Probably got a select permission denied - no column access, so default the data...
			If mstrErrorString = "No Data" Or mstrErrorString = "No Access" Then
				mstrSQL = "SELECT '" & mstrErrorString & "' AS [HORIZONTAL], '" & mstrErrorString & "' AS [HORIZONTAL_ID], '" & mstrErrorString & "' AS [VERTICAL], '" & mstrErrorString & "' AS [VERTICAL_ID], '" & mstrErrorString & "' AS [Aggregate], '" & mstrErrorString & "' AS [COLOUR]"
			Else
				mstrSQL = "SELECT 'No Access' AS [HORIZONTAL], 'No Access' AS [HORIZONTAL_ID], 'No Access' AS [VERTICAL], 'No Access' AS [VERTICAL_ID], 'No Access' AS [Aggregate], 'No Access' AS [COLOUR]"
			End If
			fOK = True
		End If

		' Execute the SQL and store in recordset
		Return DB.GetDataTable(mstrSQL, CommandType.Text)

	End Function

	Private Function SQLSelectVerticalID(plngColumnID As Integer) As Boolean
		Dim pstrSQL As String
		Dim pstrCaseStatements As String
		Dim piCount As Integer
		Dim pstrVerticalIDColumn As String
		Dim pfNullFlag As Boolean
		Dim piNull_ID As Integer

		Try

			If Len(mstrSQLSelectVerticalID) = 0 Or plngColumnID = 0 Then
				' No vertical axis (2-D table)

			Else

				pstrSQL = "SELECT DISTINCT(" & mstrSQLSelectVerticalID & ") AS [VERTICAL_ID] FROM " & mstrSQLFrom & IIf(Len(mstrSQLJoin) = 0, "", " " & mstrSQLJoin).ToString() & IIf(Len(mstrSQLWhere) = 0, "", " " & mstrSQLWhere).ToString() & " ORDER BY 1 "

				' Execute the SQL and store in recordset
				mrstChartLegendData = DB.GetDataTable(pstrSQL, CommandType.Text)
				pstrCaseStatements = ""

				' Now we've a recordset of unique values to add to the case when statement. Replacing the <$> placeholder.
				If mrstChartLegendData.Rows.Count = 0 Then
					mstrErrorString = "No Data"
					Return False
				End If

				piCount = 1
				pfNullFlag = False

				For Each objRow As DataRow In mrstChartLegendData.Rows

					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If IsDBNull(objRow("VERTICAL_ID")) Then
						' set the flag and store the value
						pfNullFlag = True
						piNull_ID = piCount
					Else
						pstrVerticalIDColumn = Trim(objRow("VERTICAL_ID").ToString())
						pstrVerticalIDColumn = Replace(pstrVerticalIDColumn, "'", "''")

						If GetColumnDataType(plngColumnID) = ColumnDataType.sqlDate Then
							pstrVerticalIDColumn = ReverseDateTextField(pstrVerticalIDColumn)
						End If
						pstrCaseStatements = pstrCaseStatements & " WHEN " & IIf(pstrVerticalIDColumn = "NULL", "NULL", "'" & pstrVerticalIDColumn & "'").ToString() & " THEN " & CStr(piCount)
					End If
					piCount = piCount + 1
				Next

				' append the 'end' statement (and 'ELSE' statement if required)
				If pfNullFlag = True And piNull_ID > 0 Then
					pstrCaseStatements = pstrCaseStatements & " ELSE " & CStr(piNull_ID)
				End If

				pstrCaseStatements = pstrCaseStatements & " END"

				' Replace the marker (<$>) in 'mstrSQL' with the case when statements...
				mstrSQL = Replace(mstrSQL, "<$>", pstrCaseStatements)
			End If

		Catch ex As Exception
			mstrErrorString = "Error selecting SQL vertical IDs." & vbNewLine & ex.Message.RemoveSensitive()
			Return False

		End Try

		Return True

	End Function

	Private Function SQLSelectHorizontalID(ByRef piSortDirection As Integer, ByRef lngColumnID As Integer) As Boolean
		Dim pstrSQL As String
		Dim pstrCaseStatements As String
		Dim piCount As Integer
		Dim pstrSQLOrderBy As String
		Dim pstrHorizontalIDColumn As String
		Dim pfNullFlag As Boolean
		Dim piNull_ID As Integer

		Try

			pstrSQLOrderBy = " ORDER BY 1 " & IIf(piSortDirection = 0, " ASC", " DESC").ToString()

			pstrSQL = "SELECT DISTINCT(" & mstrSQLSelectHorizontalID & ") AS [HORIZONTAL_ID] FROM " & mstrSQLFrom & IIf(Len(mstrSQLJoin) = 0, "", " " & mstrSQLJoin).ToString() & IIf(Len(mstrSQLWhere) = 0, "", " " & mstrSQLWhere).ToString() & pstrSQLOrderBy

			' Execute the SQL and store in recordset
			mrstChartLegendData = DB.GetDataTable(pstrSQL, CommandType.Text)
			pstrCaseStatements = ""

			' Now we've a recordset of unique values to add to the case when statement. Replacing the <$> placeholder.
			If mrstChartLegendData.Rows.Count = 0 Then
				mstrErrorString = "No Data"
				Exit Function
			End If

			piCount = 1
			pfNullFlag = False

			' loop through
			For Each objRow As DataRow In mrstChartLegendData.Rows

				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If IsDBNull(objRow("HORIZONTAL_ID")) Then
					' set the flag and store the value
					pfNullFlag = True
					piNull_ID = piCount
				Else

					Dim dc As DataColumn = objRow.Table.Columns("HORIZONTAL_ID")

					If dc.DataType Is GetType(Decimal) Then
						pstrHorizontalIDColumn = ConvertNumberForSQL(objRow("HORIZONTAL_ID").ToString())
						pstrHorizontalIDColumn = Trim(pstrHorizontalIDColumn)
					Else
						pstrHorizontalIDColumn = Trim(objRow("HORIZONTAL_ID").ToString())
					End If

					pstrHorizontalIDColumn = Replace(pstrHorizontalIDColumn, "'", "''")

					If GetColumnDataType(lngColumnID) = ColumnDataType.sqlDate Then
						pstrHorizontalIDColumn = ReverseDateTextField(pstrHorizontalIDColumn)
					End If
					pstrCaseStatements = pstrCaseStatements & " WHEN " & IIf(pstrHorizontalIDColumn = "NULL", "NULL", "'" & pstrHorizontalIDColumn & "'").ToString() & " THEN " & CStr(piCount)
				End If
				piCount += 1
			Next

			' append the 'end' statement (and 'ELSE' statement if required)
			If pfNullFlag = True And piNull_ID > 0 Then
				pstrCaseStatements = pstrCaseStatements & " ELSE " & CStr(piNull_ID)
			End If

			pstrCaseStatements = pstrCaseStatements & " END"

			' Replace the marker (<^>) in 'mstrSQL' with the case when statements...
			mstrSQL = Replace(mstrSQL, "<^>", pstrCaseStatements)

		Catch ex As Exception
			mstrErrorString = "Error selecting SQL Horizontal IDs." & vbNewLine & ex.Message.RemoveSensitive()
			Return False

		End Try

		Return True

	End Function

	Private Function GenerateSQLOrderBy(plngSortOrderID As Integer, ByRef piSortDirection As Integer) As Boolean
		' Purpose : Returns order by string from the sort order array

		Try

			' get the sort order - this is stored as decimal, but represents a 4 digit binary value
			' e.g. 5 = 0101 in binary, which represents sort orders 1 & 3=desc, 2 = asc.
			' Only first digit (leftmost) used in single axis charting.
			' Digit 1 = Horizontal Data Sort order
			' Digit 2 = Vertical Data Sort order
			' Digit 3 = 'Sort by Aggregate' tickbox
			' Digit 4 = 'Sort by Aggregate' sort order

			Dim pstrBinaryString As String
			pstrBinaryString = DecToBin(plngSortOrderID, 4)

			If Len(mstrSQLOrderBy) > 0 Then mstrSQLOrderBy = " ORDER BY " & mstrSQLOrderBy

			If Mid(pstrBinaryString, 3, 1) = "1" Then	' The third switch is for 'Sort by Aggregate'
				piSortDirection = CInt(Right(pstrBinaryString, 1))
				mstrSQLOrderBy = "[AGGREGATE] " & IIf(piSortDirection = 0, "ASC", "DESC").ToString()
				mstrSQLOrderBy = mstrSQLOrderBy & ", " & IIf(Left(pstrBinaryString, 1) = "0", " [HORIZONTAL] ASC ", " [HORIZONTAL] DESC ").ToString()
				If Len(mstrSQLSelectVerticalID) > 0 Then ' may be 2 axis chart
					mstrSQLOrderBy = mstrSQLOrderBy & ", " & IIf(Mid(pstrBinaryString, 2, 1) = "0", " [VERTICAL] ASC ", " [VERTICAL] DESC ").ToString()
				End If

			Else
				mstrSQLOrderBy = IIf(Left(pstrBinaryString, 1) = "0", " [HORIZONTAL] ASC ", " [HORIZONTAL] DESC ").ToString()
				If Len(mstrSQLSelectVerticalID) > 0 Then ' may be 2 axis chart
					mstrSQLOrderBy = mstrSQLOrderBy & ", " & IIf(Mid(pstrBinaryString, 2, 1) = "0", " [VERTICAL] ASC ", " [VERTICAL] DESC ").ToString()
				End If
				mstrSQLOrderBy = mstrSQLOrderBy & ", " & IIf(Mid(pstrBinaryString, 4, 1) = "0", " [AGGREGATE] ASC ", " [AGGREGATE] DESC ").ToString()
			End If


			If Len(mstrSQLOrderBy) > 0 Then mstrSQLOrderBy = " ORDER BY " & mstrSQLOrderBy

		Catch ex As Exception
			Return False

		End Try

		Return True

	End Function

	Private Sub MergeSQLStrings()
		mstrSQL = "SELECT " & mstrSQLSelect & " FROM " & mstrSQLFrom & IIf(Len(mstrSQLJoin) = 0, "", " " & mstrSQLJoin).ToString() & IIf(Len(mstrSQLWhere) = 0, "", " " & mstrSQLWhere).ToString() & " GROUP BY " & mstrSQLGroupBy & mstrSQLOrderBy
	End Sub

	Private Function GenerateSQLSelect(ByRef lngTableID As Integer, ByRef strTableName As String, ByRef lngColumnID As Integer, ByRef strColumnName As String, ByRef lngTableID2 As Integer, ByRef lngColumnID2 As Integer, ByRef strTableName2 As String, ByRef strColumnName2 As String, ByRef lngTableID3 As Integer, ByRef lngColumnID3 As Integer, ByRef strTableName3 As String, ByRef strColumnName3 As String, ByRef iAggregateType As Short, ByRef lngColourColumnName As Integer, ByRef strColourColumnName As String) As Boolean

		Dim plngTempTableID As Integer
		Dim pstrTempTableName As String
		Dim pstrTempColumnName As String

		Dim pblnOK As Boolean
		Dim pblnColumnOK As Boolean
		Dim pblnNoSelect As Boolean
		Dim pblnFound As Boolean

		Dim pintLoop As Short
		Dim pstrColumnList As String
		Dim pstrColumnListClean As String
		Dim pstrColumnListforVerticalID As String
		Dim pstrColumnListforHorizontalID As String
		Dim pstrColumnCode As String
		Dim pstrSource As String
		Dim pintNextIndex As Integer

		Dim objTableView As TablePrivilege

		Dim pstrAggregatePrefix As String

		Try

			' Set flags with their starting values
			pblnOK = True
			pblnNoSelect = False

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
				mstrErrorString = "No Access"
				Exit Function
			End If

			' Start off the select statement
			mstrSQLSelect = ""
			mstrSQLGroupBy = ""
			mstrSQLSelectVerticalID = ""
			mstrSQLSelectHorizontalID = ""

			' Dimension an array of tables/views joined to the base table/view
			' Column 1 = 0 if this row is for a table, 1 if it is for a view
			' Column 2 = table/view ID
			' (should contain everything which needs to be joined to the base tbl/view)
			ReDim mlngTableViews(2, 0)

			' Array to hold the columns used in the chart
			Dim mvarColDetails(3, 3) As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(0, 0) = lngTableID	' X-AXIS
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(0, 1) = strTableName
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(0, 2) = strColumnName

			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(1, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(1, 0) = lngTableID2 ' Z-AXIS - optional
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(1, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(1, 1) = strTableName2
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(1, 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(1, 2) = strColumnName2

			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(2, 0) = lngTableID3 ' Y-AXIS - need to apply aggregate.
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(2, 1) = strTableName3
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(2, 2) = strColumnName3

			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(3, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(3, 0) = lngTableID3 ' Colour Column Name
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(3, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(3, 1) = strTableName3
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(3, 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(3, 2) = strColourColumnName

			For pintLoop = 0 To 3

				pstrAggregatePrefix = ""

				If pintLoop = 2 Then 'This is the intersection column, prefix the column as required
					Select Case iAggregateType
						Case 0 ' Count
							pstrAggregatePrefix = "COUNT("
						Case 1 ' sum
							pstrAggregatePrefix = "SUM("
						Case 2 ' Average
							pstrAggregatePrefix = "AVG("
						Case 3 ' Minimum
							pstrAggregatePrefix = "MIN("
						Case 4 ' Maximum
							pstrAggregatePrefix = "MAX("
						Case Else	' unknown or not set
							pstrAggregatePrefix = ""
					End Select
				End If

				' Load the temp variables
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(pintLoop, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				plngTempTableID = CInt(mvarColDetails(pintLoop, 0))
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(pintLoop, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pstrTempTableName = mvarColDetails(pintLoop, 1).ToString()
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(pintLoop, 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pstrTempColumnName = mvarColDetails(pintLoop, 2).ToString()

				If plngTempTableID <> 0 Then ' should only happen for the z-axis (2 dimensional chart)

					If pintLoop = 3 And pstrTempColumnName = vbNullString Then
						pstrColumnList = pstrColumnList & ", '" & System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White) & "' AS [COLOUR]"
						Exit For ' No Colour Column.
					End If

					' Check permission on that column
					mobjColumnPrivileges = GetColumnPrivileges(pstrTempTableName)
					mstrRealSource = gcoTablePrivileges.Item(pstrTempTableName).RealSource
					pblnColumnOK = mobjColumnPrivileges.IsValid(pstrTempColumnName)

					If pblnColumnOK Then
						pblnColumnOK = mobjColumnPrivileges.Item(pstrTempColumnName).AllowSelect
					End If

					If pblnColumnOK Then

						' this column can be read direct from the tbl/view or from a parent table

						pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ",", "").ToString() & pstrAggregatePrefix & IIf(pintLoop = 1 Or pintLoop = 0, "CASE ", "").ToString() & mstrRealSource & "." & Trim(pstrTempColumnName)

						' once again, for the groupby, without aggregate prefix.
						' Excluding any aggregates
						If pintLoop <> 2 Then
							pstrColumnListClean = pstrColumnListClean & IIf(Len(pstrColumnListClean) > 0, ",", "").ToString() & mstrRealSource & "." & Trim(pstrTempColumnName)
						End If

						' If this is the 'Horizontal' column, create the select statement for the chart legend
						If pintLoop = 0 Then
							pstrColumnListforHorizontalID = mstrRealSource & "." & Trim(pstrTempColumnName)
						End If

						' If this is the 'vertical' column, create the select statement for the chart legend
						If pintLoop = 1 Then
							pstrColumnListforVerticalID = mstrRealSource & "." & Trim(pstrTempColumnName)
						End If

						' If the table isnt the base table (or its realsource) then
						' Check if it has already been added to the array. If not, add it.
						If plngTempTableID <> mlngBaseTableID Then
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
							If (Not mobjTableView.IsTable) And (mobjTableView.TableID = plngTempTableID) And (mobjTableView.AllowSelect) Then

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
							pstrColumnCode = " COALESCE("
							For pintNextIndex = 1 To UBound(mstrViews)
								'CHANGE TO COALESCE
								' pstrColumnCode = pstrColumnCode & _
								'' " WHEN NOT " & mstrViews(pintNextIndex) & "." & pstrTempColumnName & " IS NULL THEN " & mstrViews(pintNextIndex) & "." & pstrTempColumnName
								' CHanged to COALESCE
								pstrColumnCode = pstrColumnCode & IIf(pintNextIndex > 1, ", ", "").ToString() & mstrViews(pintNextIndex) & "." & pstrTempColumnName

							Next pintNextIndex

							If Len(pstrColumnCode) > 0 Then
								'            pstrColumnCode = pstrColumnCode & _
								'" ELSE NULL" & _
								'" END AS '" & mvarColDetails(0, pintLoop) & "'"
								' NPG - change to coalesce
								'          pstrColumnCode = "CASE " & pstrColumnCode & _
								'" ELSE NULL" & _
								'" END "

								pstrColumnCode = pstrColumnCode & ", NULL)"

								pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ",", "").ToString() & pstrAggregatePrefix & IIf(pintLoop = 1 Or pintLoop = 0, "CASE ", "").ToString() & pstrColumnCode

								' Repeat for groupby...BUT EXCLUDE the intersection column.
								If pintLoop <> 2 Then
									pstrColumnListClean = pstrColumnListClean & IIf(Len(pstrColumnListClean) > 0, ",", "").ToString() & pstrColumnCode
								End If

								' If this is the 'Horizontal' column, create the select statement for the chart legend
								If pintLoop = 0 Then
									pstrColumnListforHorizontalID = pstrColumnCode
								End If

								' If this is the 'vertical' column, create the select statement for the chart legend
								If pintLoop = 1 Then
									pstrColumnListforVerticalID = pstrColumnCode
								End If

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

					Select Case pintLoop
						Case 0 ' Horizontal
							pstrColumnList = pstrColumnList & " <^> AS [HORIZONTAL_ID]"
							' Insert the 'Horizontal' column value
							pstrColumnList = pstrColumnList & IIf(Len(pstrColumnListforHorizontalID) > 0, ", " & pstrColumnListforHorizontalID & " AS [HORIZONTAL]", "").ToString()
						Case 1 ' vertical_ID
							pstrColumnList = pstrColumnList & " <$> AS [VERTICAL_ID]"
							' Insert the 'Vertical' column value
							pstrColumnList = pstrColumnList & IIf(Len(pstrColumnListforVerticalID) > 0, ", " & pstrColumnListforVerticalID & " AS [VERTICAL]", "").ToString()
						Case 2 ' intersection
							If Len(pstrAggregatePrefix) > 0 And pintLoop = 2 Then
								pstrColumnList = pstrColumnList & ") AS [Aggregate]"
							End If
						Case 3 ' colour column
							pstrColumnList = pstrColumnList & " AS [COLOUR]"
					End Select
				Else
					' 2-D chart, fix Z-Axis at 0...
					pstrColumnList = pstrColumnList & ", 1 AS [VERTICAL_ID], '' as [VERTICAL]"
				End If

			Next pintLoop

			mstrSQLSelect = mstrSQLSelect & pstrColumnList
			mstrSQLGroupBy = mstrSQLGroupBy & pstrColumnListClean
			mstrSQLSelectVerticalID = pstrColumnListforVerticalID
			mstrSQLSelectHorizontalID = pstrColumnListforHorizontalID

		Catch ex As Exception
			mstrErrorString = "Error generating SQL Select statement." & vbNewLine & ex.Message.RemoveSensitive()
			Return False

		End Try

		Return True

	End Function

	Private Sub GenerateSQLFrom(strTableName As String)
		mstrSQLFrom = gcoTablePrivileges.Item(strTableName).RealSource
	End Sub

	Private Function GenerateSQLJoin(lngTableID As Integer) As Boolean

		' Purpose : Add the join strings for parent/child/views.
		'           Also adds filter clauses to the joins if used
		Dim pobjTableView As TablePrivilege
		Dim pintLoop As Integer

		Try

			' Get the base table real source
			mstrBaseTableRealSource = mstrSQLFrom

			'sOtherParentJoinCode = ""

			' First, do the join for all the views etc...

			For pintLoop = 1 To UBound(mlngTableViews, 2)

				' Get the table/view object from the id stored in the array
				If mlngTableViews(1, pintLoop) = 0 Then
					pobjTableView = gcoTablePrivileges.FindTableID(mlngTableViews(2, pintLoop))
				Else
					pobjTableView = gcoTablePrivileges.FindViewID(mlngTableViews(2, pintLoop))
				End If

				' Dont add a join here if its the child table...do that later
				If (pobjTableView.TableID = lngTableID) Then
					If (pobjTableView.RealSource <> mstrBaseTableRealSource) Then
						mstrSQLJoin = mstrSQLJoin & " LEFT OUTER JOIN " & pobjTableView.RealSource & " ON " & mstrBaseTableRealSource & ".ID = " & pobjTableView.RealSource & ".ID"
					End If
				Else
					If (pobjTableView.ViewName <> mstrBaseTableRealSource) Then
						If IsAChildOf((pobjTableView.TableID), lngTableID) = True Then
							mstrSQLJoin = mstrSQLJoin & " LEFT OUTER JOIN " & pobjTableView.RealSource & " ON " & pobjTableView.RealSource & ".ID_" & lngTableID & " = " & mstrBaseTableRealSource & ".ID"
						Else
							'
							mstrSQLJoin = mstrSQLJoin & " LEFT OUTER JOIN " & pobjTableView.RealSource & " ON " & mstrBaseTableRealSource & ".ID_" & CStr(pobjTableView.TableID) & " = " & pobjTableView.RealSource & ".ID"
						End If
					End If

				End If
			Next pintLoop


		Catch ex As Exception
			mstrErrorString = "Error in GenerateSQLJoin." & vbNewLine & ex.Message.RemoveSensitive()
			Return False

		End Try

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
			mstrErrorString = "Error in GenerateSQLWhere." & vbNewLine & ex.Message.RemoveSensitive()
			Return False

		End Try

		Return True

	End Function

	Private Function ReverseDateTextField(pDateValue As String) As String

		' '30/12/1998' becomes '1998/12/30'

		If InStr(pDateValue, "/") = 0 Or Len(pDateValue) <> 10 Then
			ReverseDateTextField = pDateValue
			Exit Function
		End If

		Return Mid(pDateValue, InStrRev(pDateValue, "/") + 1, 4) & "/" & Mid(pDateValue, InStr(pDateValue, "/") + 1, 2) & "/" & Left(pDateValue, 2)

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