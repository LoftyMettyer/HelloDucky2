Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsTodaysAbsence_NET.clsTodaysAbsence")> Public Class clsTodaysAbsence
	Private mclsData As New clsDataAccess
	Private AbsentList As Collection
	
	Private Const mstrDATESQL As String = "mm/dd/yyyy"
	Private mobjTableView As CTablePrivilege
	Private mobjColumnPrivileges As CColumnPrivileges
	Private mstrRealSource As String
	Private mstrSQLString As String
	Private mstrSQLFrom As String
	Private mstrSQLJoin As String
	Private mstrSQLWhere As String
	Private mstrBaseTableRealSource As String
	Private mstrErrorString As String
	Private mlngTableViews() As Integer
	Private mstrSQL As String
	Private mstrAbsenceRealSource As String
	
	Public Function GetTodaysAbsences(ByRef RecordID As Object, Optional ByRef dtStartDate As Date = #12:00:00 AM#, Optional ByRef dtEndDate As Date = #12:00:00 AM#) As Object
		
		Dim blnWhere As Boolean
		Dim strWhere As String
		Dim plngEmployeeID As Integer
		Dim rsEmployeeName As ADODB.Recordset
		Dim objTableView As CTablePrivilege
		Dim lngTableID As Integer
		Dim pblnOK As Boolean
		Dim pblnColumnOK As Boolean
		
		SetupTablesCollection()
		
		' Get the absence table name and Personnel Records ID and name from Module Setup
		ReadAbsenceParameters()
		ReadPersonnelParameters()
		
		' Check the user has permission to read the absence table.
		For	Each objTableView In gcoTablePrivileges.Collection
			If (objTableView.TableID = glngAbsenceTableID) And (objTableView.AllowSelect) Then
				pblnOK = True
				Exit For
			End If
		Next objTableView
		'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objTableView = Nothing
		
		If Not pblnOK Then
			'UPGRADE_WARNING: Couldn't resolve default property of object GetTodaysAbsences. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetTodaysAbsences = ""
			mstrErrorString = "You do not have permission to read the base table either directly or through any views."
			Exit Function
		End If
		
		' Store the absence view name
		mstrAbsenceRealSource = gcoTablePrivileges.Item("Absence").RealSource
		
		' Build the Personnel Select string
		If pblnOK Then pblnOK = GenerateSQLSelect
		If pblnOK Then pblnOK = GenerateSQLFrom(gsPersonnelTableName)
		If pblnOK Then pblnOK = GenerateSQLJoin(glngPersonnelTableID)
		'UPGRADE_WARNING: Couldn't resolve default property of object RecordID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If pblnOK Then pblnOK = GenerateSQLWhere(glngPersonnelTableID, plngEmployeeID, CInt(RecordID))
		'If pblnOK Then pblnOK = GenerateSQLOrderBy(lngSortOrderID, iSortDirection)
		If pblnOK Then pblnOK = MergeSQLStrings()
		
		GetTodaysAbsences = mclsData.OpenPersistentRecordset(mstrSQL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
		
LocalErr: 
		
	End Function
	
	Public WriteOnly Property Connection() As Object
		Set(ByVal Value As Object)
			
			' JDM - Create connection object differently if we are in development mode (i.e. debug mode)
			If ASRDEVELOPMENT Then
				gADOCon = New ADODB.Connection
				'UPGRADE_WARNING: Couldn't resolve default property of object vConnection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				gADOCon.Open(Value)
			Else
				gADOCon = Value
			End If
			
		End Set
	End Property
	
	Public WriteOnly Property Username() As String
		Set(ByVal Value As String)
			' Username passed in from the asp page
			gsUsername = Value
		End Set
	End Property
	
	
	Private Function GenerateSQLSelect() As Boolean
		
		On Error GoTo GenerateSQLSelect_ERROR
		
		Dim plngTempTableID As Integer
		Dim pstrTempTableName As String
		Dim pstrTempColumnName As String
		
		Dim pblnOK As Boolean
		Dim pblnColumnOK As Boolean
		Dim iLoop1 As Short
		Dim pblnNoSelect As Boolean
		Dim pblnFound As Boolean
		
		Dim pintLoop As Short
		Dim pstrColumnList As String
		Dim pstrColumnCode As String
		Dim pstrSource As String
		Dim pintNextIndex As Short
		
		Dim blnOK As Boolean
		Dim sCalcCode As String
		Dim alngSourceTables() As Integer
		Dim objCalcExpr As clsExprExpression
		Dim objTableView As CTablePrivilege
		Dim pintNextColLoop As Short
		
		' Set flags with their starting values
		pblnOK = True
		pblnNoSelect = False
		
		Dim mastrUDFsRequired(0) As Object
		
		' JPD20030219 Fault 5068
		' Check the user has permission to read the base table.
		pblnOK = False
		For	Each objTableView In gcoTablePrivileges.Collection
			If (objTableView.TableID = glngPersonnelTableID) And (objTableView.AllowSelect) Then
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
		plngTempTableID = glngPersonnelTableID
		pstrTempTableName = gsPersonnelTableName
		
		' Fault HRPRO-1362 - changed "forename surname" to "surname, forename"
		For pintNextColLoop = 1 To 2
			If pintNextColLoop = 1 Then
				pstrTempColumnName = gsPersonnelSurnameColumnName
			ElseIf pintNextColLoop = 2 Then 
				pstrTempColumnName = gsPersonnelForenameColumnName
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
				If plngTempTableID <> glngPersonnelTableID Then
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
				For	Each mobjTableView In gcoTablePrivileges.Collection
					If (Not mobjTableView.IsTable) And (mobjTableView.TableID = glngPersonnelTableID) And (mobjTableView.AllowSelect) Then
						
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
		
		GenerateSQLSelect = True
		
		Exit Function
		
GenerateSQLSelect_ERROR: 
		
		GenerateSQLSelect = False
		mstrErrorString = "Error generating SQL Select statement." & vbNewLine & Err.Description
		
	End Function
	
	Private Function GenerateSQLFrom(ByRef strTableName As String) As Boolean
		
		Dim iLoop As Short
		Dim pobjTableView As CTablePrivilege
		
		pobjTableView = New CTablePrivilege
		
		mstrSQLFrom = gcoTablePrivileges.Item(strTableName).RealSource
		
		'UPGRADE_NOTE: Object pobjTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pobjTableView = Nothing
		
		GenerateSQLFrom = True
		Exit Function
		
GenerateSQLFrom_ERROR: 
		
		GenerateSQLFrom = False
		mstrErrorString = "Error in GenerateSQLFrom." & vbNewLine & Err.Description
		
	End Function
	
	Private Function GenerateSQLJoin(ByRef lngTableID As Integer) As Boolean
		
		' Purpose : Add the join strings for parent/child/views.
		'           Also adds filter clauses to the joins if used
		
		On Error GoTo GenerateSQLJoin_ERROR
		
		Dim pobjTableView As CTablePrivilege
		Dim objChildTable As CTablePrivilege
		Dim pintLoop As Short
		Dim sChildJoinCode As String
		Dim sReuseJoinCode As String
		Dim sChildOrderString As String
		Dim rsTemp As ADODB.Recordset
		Dim strFilterIDs As String
		Dim blnOK As Boolean
		Dim pblnChildUsed As Boolean
		Dim sChildJoin As String
		Dim lngTempChildID As Integer
		Dim lngTempMaxRecords As Integer
		Dim lngTempFilterID As Integer
		Dim lngTempOrderID As Integer
		Dim i As Short
		Dim sOtherParentJoinCode As String
		Dim iLoop2 As Short
		
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
		
		' NPG20110126 Fault HRPRO-
		
		'  If pobjTableView Is Nothing Then
		' Full table access
		' Append the absence table
		mstrSQLJoin = mstrSQLJoin & " JOIN " & mstrAbsenceRealSource & " ON " & mstrAbsenceRealSource & ".ID_" & CStr(glngPersonnelTableID) & " = " & mstrBaseTableRealSource & ".ID"
		'  Else
		'    ' Append the absence table
		'    mstrSQLJoin = mstrSQLJoin & _
		''    " JOIN " & mstrAbsenceRealSource & _
		''    " ON " & mstrAbsenceRealSource & ".ID_" & CStr(glngPersonnelTableID) & " = " & pobjTableView.RealSource & ".ID"
		'  End If
		
		GenerateSQLJoin = True
		Exit Function
		
GenerateSQLJoin_ERROR: 
		
		GenerateSQLJoin = False
		mstrErrorString = "Error in GenerateSQLJoin." & vbNewLine & Err.Description
		
	End Function
	
	
	Private Function GenerateSQLWhere(ByRef lngTableID As Integer, ByRef lngEmployeeID As Integer, ByRef lngRecordID As Integer) As Boolean
		Dim pstrSQL As String
		Dim strAM_End As String
		Dim strPM_Start As String
		Dim strCurrentSession As String
		
		' NPG20110126 Fault HRPRO-1343
		'  mstrSQLWhere = "WHERE ((" & mstrAbsenceRealSource & ".ID_" & CStr(glngPersonnelTableID) & " = " & lngRecordID & " OR " & mstrAbsenceRealSource & ".ID_" & CStr(glngPersonnelTableID) & " IN (" & _
		''      "SELECT ID FROM dbo.udf_ASRFn_ByID_IsPersonnelSubordinateOfUser(" & lngRecordID & ")))"
		
		mstrSQLWhere = "WHERE "
		
		
		' Get the start and end session variables and compare to local time.
		'UPGRADE_WARNING: Couldn't resolve default property of object GetSystemSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		strAM_End = GetSystemSetting("outlook", "amendtime", "12:30")
		'UPGRADE_WARNING: Couldn't resolve default property of object GetSystemSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		strPM_Start = GetSystemSetting("outlook", "pmstarttime", "13:30")
		
		If TimeOfDay < CDate(strPM_Start) And TimeOfDay > CDate(strAM_End) Then
			strCurrentSession = ""
		ElseIf TimeOfDay < CDate(strPM_Start) Then 
			strCurrentSession = "AM"
		ElseIf TimeOfDay > CDate(strAM_End) Then 
			strCurrentSession = "PM"
		End If
		
		' Get today's absences...
		If strCurrentSession = "PM" Then
			pstrSQL = " (DATEDIFF(d, " & mstrAbsenceRealSource & "." & gsAbsenceStartDateColumnName & ", GETDATE()) > 0" & " OR (DATEDIFF(d, " & mstrAbsenceRealSource & "." & gsAbsenceStartDateColumnName & ", GETDATE()) = 0))" & " AND ((DATEDIFF(d," & mstrAbsenceRealSource & "." & gsAbsenceEndDateColumnName & ", GETDATE()) < 0 OR " & mstrAbsenceRealSource & "." & gsAbsenceEndDateColumnName & " IS NULL)" & " OR (DATEDIFF(d," & mstrAbsenceRealSource & "." & gsAbsenceEndDateColumnName & ", GETDATE()) = 0 AND (" & mstrAbsenceRealSource & "." & gsAbsenceEndSessionColumnName & " = 'PM')))"
		ElseIf strCurrentSession = "AM" Then 
			pstrSQL = " (DATEDIFF(d," & mstrAbsenceRealSource & "." & gsAbsenceStartDateColumnName & ", GETDATE()) > 0" & " OR (DATEDIFF(d, " & mstrAbsenceRealSource & "." & gsAbsenceStartDateColumnName & ", GETDATE()) = 0 AND (" & mstrAbsenceRealSource & "." & gsAbsenceStartSessionColumnName & "='AM')))" & " AND ((DATEDIFF(d, " & mstrAbsenceRealSource & "." & gsAbsenceEndDateColumnName & ", GETDATE()) < 0 OR " & mstrAbsenceRealSource & "." & gsAbsenceEndDateColumnName & " IS NULL)" & " OR (DATEDIFF(d, " & mstrAbsenceRealSource & "." & gsAbsenceEndDateColumnName & ", GETDATE()) = 0))"
		Else
			' Lunch! Any absence that spans today.
			pstrSQL = " DATEDIFF(d," & mstrAbsenceRealSource & "." & gsAbsenceStartDateColumnName & ", GETDATE()) >= 0" & " AND (DATEDIFF(d, " & mstrAbsenceRealSource & "." & gsAbsenceEndDateColumnName & ", GETDATE()) <= 0 OR " & mstrAbsenceRealSource & "." & gsAbsenceEndDateColumnName & " IS NULL)"
		End If
		
		mstrSQLWhere = mstrSQLWhere & pstrSQL
		
		GenerateSQLWhere = True
		Exit Function
		
GenerateSQLWhere_ERROR: 
		
		GenerateSQLWhere = False
		mstrErrorString = "Error in GenerateSQLWhere." & vbNewLine & Err.Description
		
	End Function
	
	Private Function MergeSQLStrings() As Boolean
		Dim pstrAggregate As String
		
		On Error GoTo MergeSQLStrings_ERROR
		
		mstrSQL = "SELECT DISTINCT " & mstrSQLString & " FROM " & mstrSQLFrom & IIf(Len(mstrSQLJoin) = 0, "", " " & mstrSQLJoin) & IIf(Len(mstrSQLWhere) = 0, "", " " & mstrSQLWhere) & " ORDER BY 1"
		
		MergeSQLStrings = True
		Exit Function
		
MergeSQLStrings_ERROR: 
		MergeSQLStrings = False
		
	End Function
End Class