Option Strict Off
Option Explicit On

Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Metadata

Public Class clsTodaysAbsence
	Inherits BaseForDMI

	Private mobjTableView As TablePrivilege
	Private mobjColumnPrivileges As CColumnPrivileges
	Private mstrRealSource As String
	Private mstrSQLString As String
	Private mstrSQLFrom As String
	Private mstrSQLJoin As String
	Private mstrSQLWhere As String
	Private mstrBaseTableRealSource As String
	Private mstrErrorString As String
	Private mlngTableViews(,) As Integer
	Private mstrSQL As String
	Private mstrAbsenceRealSource As String

	Public Function GetTodaysAbsences(ByRef RecordID As Integer, Optional ByRef dtStartDate As Date = #12:00:00 AM#, Optional ByRef dtEndDate As Date = #12:00:00 AM#) As DataTable

		Dim plngEmployeeID As Integer
		Dim pblnOK As Boolean

		Try

			' Check the user has permission to read the necessary absence columns.
			Dim permissions = GetColumnPrivileges(AbsenceModule.gsAbsenceTableName)

			If permissions(AbsenceModule.gsAbsenceStartDateColumnName).AllowSelect And _
				permissions(AbsenceModule.gsAbsenceStartSessionColumnName).AllowSelect And _
				permissions(AbsenceModule.gsAbsenceEndDateColumnName).AllowSelect And _
				permissions(AbsenceModule.gsAbsenceEndSessionColumnName).AllowSelect Then

				' Store the absence view name
				mstrAbsenceRealSource = gcoTablePrivileges.Item(AbsenceModule.gsAbsenceTableName).RealSource

				' Build the Personnel Select string
				pblnOK = GenerateSQLSelect()
				If pblnOK Then pblnOK = GenerateSQLFrom(PersonnelModule.gsPersonnelTableName)
				If pblnOK Then pblnOK = GenerateSQLJoin(PersonnelModule.glngPersonnelTableID)
				If pblnOK Then pblnOK = GenerateSQLWhere(PersonnelModule.glngPersonnelTableID, plngEmployeeID, RecordID)
				If pblnOK Then pblnOK = MergeSQLStrings()

				Return DB.GetDataTable(mstrSQL, CommandType.Text)

			End If

		Catch ex As Exception
			Throw

		End Try

		Return Nothing

	End Function

	Private Function GenerateSQLSelect() As Boolean

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
		Dim pintNextColLoop As Short

		Try

			' Set flags with their starting values
			pblnNoSelect = False

			' JPD20030219 Fault 5068
			' Check the user has permission to read the base table.
			pblnOK = False
			For Each objTableView In gcoTablePrivileges.Collection
				If (objTableView.TableID = PersonnelModule.glngPersonnelTableID) And (objTableView.AllowSelect) Then
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

			mstrSQLString = ""

			' Dimension an array of tables/views joined to the base table/view
			' Column 1 = 0 if this row is for a table, 1 if it is for a view
			' Column 2 = table/view ID
			' (should contain everything which needs to be joined to the base tbl/view)
			ReDim mlngTableViews(2, 0)

			' Load the temp variables
			plngTempTableID = PersonnelModule.glngPersonnelTableID
			pstrTempTableName = PersonnelModule.gsPersonnelTableName

			' Fault HRPRO-1362 - changed "forename surname" to "surname, forename"
			For pintNextColLoop = 1 To 2
				If pintNextColLoop = 1 Then
					pstrTempColumnName = PersonnelModule.gsPersonnelSurnameColumnName
				ElseIf pintNextColLoop = 2 Then
					pstrTempColumnName = PersonnelModule.gsPersonnelForenameColumnName
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
					pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, " + ' ' + ", "") & mstrRealSource & "." & Trim(pstrTempColumnName)

					' If the table isnt the base table (or its realsource) then
					' Check if it has already been added to the array. If not, add it.
					If plngTempTableID <> PersonnelModule.glngPersonnelTableID Then
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

					Dim mstrViews(0) As Object
					For Each mobjTableView In gcoTablePrivileges.Collection
						If (Not mobjTableView.IsTable) And (mobjTableView.TableID = PersonnelModule.glngPersonnelTableID) And (mobjTableView.AllowSelect) Then

							pstrSource = mobjTableView.ViewName
							mstrRealSource = gcoTablePrivileges.Item(pstrSource).RealSource

							' Get the column permission for the view
							mobjColumnPrivileges = GetColumnPrivileges(pstrSource)

							' If we can see the column from this view
							If mobjColumnPrivileges.IsValid(pstrTempColumnName) Then
								If mobjColumnPrivileges.Item(pstrTempColumnName).AllowSelect Then

									ReDim Preserve mstrViews(UBound(mstrViews) + 1)
									'UPGRADE_WARNING: Couldn't resolve default property of object mstrViews(UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
						pstrColumnCode = "COALESCE("
						For pintNextIndex = 1 To UBound(mstrViews)
							'UPGRADE_WARNING: Couldn't resolve default property of object mstrViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							pstrColumnCode = pstrColumnCode & IIf(pintNextIndex = 1, "", ", ") & mstrViews(pintNextIndex) & "." & pstrTempColumnName
						Next pintNextIndex

						pstrColumnCode = pstrColumnCode & ", NULL)"

						pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, " + ', ' + ", "") & pstrColumnCode

					End If

					' If we cant see a column, then get outta here
					If pblnNoSelect Then
						GenerateSQLSelect = False
						mstrErrorString = "You do not have permission to see the column '" & pstrTempColumnName & "' either directly or through any views."
						Exit Function
					End If

					If Not pblnOK Then
						GenerateSQLSelect = False
						Exit Function
					End If

				End If

				mstrSQLString = pstrColumnList
			Next pintNextColLoop

		Catch ex As Exception
			mstrErrorString = "Error generating SQL Select statement." & vbNewLine & ex.Message.RemoveSensitive()
			Return False

		End Try

		Return True

	End Function

	Private Function GenerateSQLFrom(ByRef strTableName As String) As Boolean

		mstrSQLFrom = gcoTablePrivileges.Item(strTableName).RealSource
		Return True

	End Function

	Private Function GenerateSQLJoin(ByRef lngTableID As Integer) As Boolean

		' Purpose : Add the join strings for parent/child/views.
		'           Also adds filter clauses to the joins if used
		Dim pobjTableView As TablePrivilege
		Dim pintLoop As Integer

		Try

			' Get the base table real source
			mstrBaseTableRealSource = mstrSQLFrom

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

			' Append the absence table
			mstrSQLJoin = mstrSQLJoin & " JOIN " & mstrAbsenceRealSource & " ON " & mstrAbsenceRealSource & ".ID_" & CStr(PersonnelModule.glngPersonnelTableID) & " = " & mstrBaseTableRealSource & ".ID"

		Catch ex As Exception
			mstrErrorString = "Error in GenerateSQLJoin." & vbNewLine & ex.Message.RemoveSensitive()
			Return False

		End Try

		Return True

	End Function


	Private Function GenerateSQLWhere(ByRef lngTableID As Integer, ByRef lngEmployeeID As Integer, ByRef lngRecordID As Integer) As Boolean
		Dim pstrSQL As String
		Dim strAM_End As String
		Dim strPM_Start As String
		Dim strCurrentSession As String

		mstrSQLWhere = "WHERE "


		' Get the start and end session variables and compare to local time.
		strAM_End = SystemSettings.GetSetting("outlook", "amendtime", "12:30").Value
		strPM_Start = SystemSettings.GetSetting("outlook", "pmstarttime", "13:30").Value


		If TimeOfDay < CDate(strPM_Start) And TimeOfDay > CDate(strAM_End) Then
			strCurrentSession = ""
		ElseIf TimeOfDay < CDate(strPM_Start) Then
			strCurrentSession = "AM"
		ElseIf TimeOfDay > CDate(strAM_End) Then
			strCurrentSession = "PM"
		End If

		' Get today's absences...
		If strCurrentSession = "PM" Then
			pstrSQL = " (DATEDIFF(d, " & mstrAbsenceRealSource & "." & AbsenceModule.gsAbsenceStartDateColumnName & ", GETDATE()) > 0" & " OR (DATEDIFF(d, " & mstrAbsenceRealSource & "." & AbsenceModule.gsAbsenceStartDateColumnName & ", GETDATE()) = 0))" & " AND ((DATEDIFF(d," & mstrAbsenceRealSource & "." & AbsenceModule.gsAbsenceEndDateColumnName & ", GETDATE()) < 0 OR " & mstrAbsenceRealSource & "." & AbsenceModule.gsAbsenceEndDateColumnName & " IS NULL)" & " OR (DATEDIFF(d," & mstrAbsenceRealSource & "." & AbsenceModule.gsAbsenceEndDateColumnName & ", GETDATE()) = 0 AND (" & mstrAbsenceRealSource & "." & AbsenceModule.gsAbsenceEndSessionColumnName & " = 'PM')))"
		ElseIf strCurrentSession = "AM" Then
			pstrSQL = " (DATEDIFF(d," & mstrAbsenceRealSource & "." & AbsenceModule.gsAbsenceStartDateColumnName & ", GETDATE()) > 0" & " OR (DATEDIFF(d, " & mstrAbsenceRealSource & "." & AbsenceModule.gsAbsenceStartDateColumnName & ", GETDATE()) = 0 AND (" & mstrAbsenceRealSource & "." & AbsenceModule.gsAbsenceStartSessionColumnName & "='AM')))" & " AND ((DATEDIFF(d, " & mstrAbsenceRealSource & "." & AbsenceModule.gsAbsenceEndDateColumnName & ", GETDATE()) < 0 OR " & mstrAbsenceRealSource & "." & AbsenceModule.gsAbsenceEndDateColumnName & " IS NULL)" & " OR (DATEDIFF(d, " & mstrAbsenceRealSource & "." & AbsenceModule.gsAbsenceEndDateColumnName & ", GETDATE()) = 0))"
		Else
			' Lunch! Any absence that spans today.
			pstrSQL = " DATEDIFF(d," & mstrAbsenceRealSource & "." & AbsenceModule.gsAbsenceStartDateColumnName & ", GETDATE()) >= 0" & " AND (DATEDIFF(d, " & mstrAbsenceRealSource & "." & AbsenceModule.gsAbsenceEndDateColumnName & ", GETDATE()) <= 0 OR " & mstrAbsenceRealSource & "." & AbsenceModule.gsAbsenceEndDateColumnName & " IS NULL)"
		End If

		mstrSQLWhere = mstrSQLWhere & pstrSQL

		Return True

	End Function

	Private Function MergeSQLStrings() As Boolean

		mstrSQL = "SELECT DISTINCT " & mstrSQLString & " FROM " & mstrSQLFrom & IIf(Len(mstrSQLJoin) = 0, "", " " & mstrSQLJoin) & IIf(Len(mstrSQLWhere) = 0, "", " " & mstrSQLWhere) & " ORDER BY 1"

		Return True


	End Function
End Class