Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsChart_NET.clsChart")> Public Class clsChart
	
	Private mastrUDFsRequired() As String
	Private mvarPrompts() As Object
	Private mstrRealSource As String
	Private mstrBaseTableRealSource As String
  Private mlngTableViews(,) As Integer
	Private mstrViews() As String
	Private mobjTableView As CTablePrivilege
	Private mobjColumnPrivileges As CColumnPrivileges
	
	' Classes
	Private mclsGeneral As clsGeneral
	Private mclsData As clsDataAccess
	
	' Strings to hold the SQL statement
	Private mstrSQLSelect As String
	Private mstrSQLSelectColour As String
	Private mstrSQLFrom As String
	Private mstrSQLJoin As String
	Private mstrSQLWhere As String
	Private mstrSQLOrderBy As String
	Private mstrSQL As String
	Private mstrErrorString As String
	
	' Recordset to store the final data from SQL
	'Private mrstChartDataOutput As New adodb.Recordset
	
	' Array holding the columns to sort the report by
	Private mvarSortOrder() As Object
	
	
	Public Function GetChartData(ByRef plngTableID As Object, ByRef plngColumnID As Object, ByRef plngFilterID As Object, ByRef piAggregateType As Object, ByRef piElementType As Object, ByRef plngSortOrderID As Object, ByRef piSortDirection As Object, ByRef plngChart_ColourID As Object) As Object
		
		Dim fOK As Boolean
		Dim strTableName As String
		Dim strColumnName As String
		Dim lngTableID As Integer
		Dim lngColumnID As Integer
		Dim lngFilterID As Integer
		Dim iAggregateType As Short
		Dim iElementType As Short
		Dim lngSortOrderID As Integer
		Dim iSortDirection As Short
		Dim lngColourID As Integer
		Dim strColourColumnName As String
		
		Dim sSQL As String
		
		fOK = True
		
		'UPGRADE_WARNING: Couldn't resolve default property of object plngTableID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		lngTableID = CInt(plngTableID)
		'UPGRADE_WARNING: Couldn't resolve default property of object plngColumnID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		lngColumnID = CInt(plngColumnID)
		'UPGRADE_WARNING: Couldn't resolve default property of object plngFilterID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		lngFilterID = CInt(plngFilterID)
		'UPGRADE_WARNING: Couldn't resolve default property of object piAggregateType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		iAggregateType = CShort(piAggregateType)
		'UPGRADE_WARNING: Couldn't resolve default property of object piElementType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		iElementType = CShort(piElementType)
		'UPGRADE_WARNING: Couldn't resolve default property of object plngSortOrderID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		lngSortOrderID = CInt(plngSortOrderID)
		'UPGRADE_WARNING: Couldn't resolve default property of object piSortDirection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		iSortDirection = CShort(piSortDirection)
		'UPGRADE_WARNING: Couldn't resolve default property of object plngChart_ColourID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		lngColourID = CInt(plngChart_ColourID)
		
		strTableName = datGeneral.GetTableName(CInt(lngTableID))
		strColumnName = mclsGeneral.GetColumnName(CInt(lngColumnID))
		strColourColumnName = mclsGeneral.GetColumnName(CInt(lngColourID))
		
		If fOK Then fOK = GenerateSQLSelect(lngTableID, strTableName, lngColumnID, strColumnName, False)
		'UPGRADE_WARNING: Couldn't resolve default property of object piElementType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If fOK And piElementType = 2 And lngColourID > 0 Then fOK = GenerateSQLSelect(lngTableID, strTableName, lngColourID, strColourColumnName, True)
		If fOK Then fOK = GenerateSQLFrom(strTableName)
		If fOK Then fOK = GenerateSQLJoin(lngTableID)
		If fOK Then fOK = GenerateSQLWhere(lngTableID, lngFilterID)
		If fOK Then fOK = GenerateSQLOrderBy(lngSortOrderID, iSortDirection)
		If fOK Then fOK = MergeSQLStrings(iAggregateType, iElementType)
		
		If Not fOK Then ' Probably got a select permission denied - no column access, so default the data...
			'UPGRADE_WARNING: Couldn't resolve default property of object piElementType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If piElementType = 4 Then
				mstrSQL = "SELECT 'No Data' AS [Aggregate]"
			Else
				mstrSQL = "SELECT 'No Access' AS [COLUMN], 'No Access' AS [Aggregate], '" & Str(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)) & "' AS [Colour]"
				fOK = True
			End If
		End If
		
		' Execute the SQL and store in recordset
		' Set mrstChartDataOutput = mclsData.OpenRecordset(sSQL, adOpenStatic, adLockReadOnly)
		
		GetChartData = mclsData.OpenRecordset(mstrSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
		
	End Function
	
	
	Public Function MergeSQLStrings(ByRef iAggregateType As Short, ByRef iElementType As Short) As Boolean
		
		Dim pstrAggregate As String
		
		On Error GoTo MergeSQLStrings_ERROR
		
		Select Case iAggregateType
			Case 1 ' sum
				pstrAggregate = "SUM(" & mstrSQLSelect & ") AS [Aggregate]"
			Case 2 '  Average
				pstrAggregate = "AVG(" & mstrSQLSelect & ") AS [Aggregate]"
			Case 3 '  Minimum
				pstrAggregate = "MIN(" & mstrSQLSelect & ") AS [Aggregate]"
			Case 4 '  Maximum
				pstrAggregate = "MAX(" & mstrSQLSelect & ") AS [Aggregate]"
			Case Else ' unknown or not set - default to count
				pstrAggregate = "COUNT(" & mstrSQLSelect & ") AS [Aggregate]"
		End Select
		
		If iElementType = 4 Then
			mstrSQL = "SELECT " & pstrAggregate & " FROM " & mstrSQLFrom & IIf(Len(mstrSQLJoin) = 0, "", " " & mstrSQLJoin) & IIf(Len(mstrSQLWhere) = 0, "", " " & mstrSQLWhere)
		Else
			mstrSQL = "SELECT " & mstrSQLSelect & " AS [COLUMN], " & pstrAggregate & IIf(mstrSQLSelectColour <> vbNullString, mstrSQLSelectColour & " AS [COLOUR] ", ", " & Str(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)) & " AS [COLOUR] ") & " FROM " & mstrSQLFrom & IIf(Len(mstrSQLJoin) = 0, "", " " & mstrSQLJoin) & IIf(Len(mstrSQLWhere) = 0, "", " " & mstrSQLWhere) & " GROUP BY " & mstrSQLSelect & mstrSQLSelectColour & mstrSQLOrderBy
		End If
		
		MergeSQLStrings = True
		
		Exit Function
		
MergeSQLStrings_ERROR: 
		MergeSQLStrings = False
		mstrErrorString = "Error merging SQL string components." & vbNewLine & Err.Description
		
	End Function
	
	
	Public Function GenerateSQLSelect(ByRef lngTableID As Integer, ByRef strTableName As String, ByRef lngColumnID As Integer, ByRef strColumnName As String, ByRef pfColourFlag As Boolean) As Boolean
		
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
		
		
		SetupTablesCollection()
		
		' Set flags with their starting values
		pblnOK = True
		pblnNoSelect = False
		
		ReDim mastrUDFsRequired(0)
		
		' JPD20030219 Fault 5068
		' Check the user has permission to read the base table.
		pblnOK = False
		For	Each objTableView In gcoTablePrivileges.Collection
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
			pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ",", "") & mstrRealSource & "." & Trim(pstrTempColumnName)
			
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
			For	Each mobjTableView In gcoTablePrivileges.Collection
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
					'            pstrColumnCode = pstrColumnCode & _
					'" ELSE NULL" & _
					'" END AS '" & mvarColDetails(0, pintLoop) & "'"
					pstrColumnCode = "CASE " & pstrColumnCode & " ELSE NULL" & " END "
					
					pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ",", "") & pstrColumnCode
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
		
		'sOtherParentJoinCode = ""
		
		' First, do the join for all the views etc...
		
		For pintLoop = 1 To UBound(mlngTableViews, 2)
			
			' Get the table/view object from the id stored in the array
			If mlngTableViews(1, pintLoop) = 0 Then
				pobjTableView = gcoTablePrivileges.FindTableID(mlngTableViews(2, pintLoop))
			Else
				pobjTableView = gcoTablePrivileges.FindViewID(mlngTableViews(2, pintLoop))
			End If
			
			'    ' Dont add a join here if its the child table...do that later
			'    'If pobjTableView.TableID <> mlngCustomReportsChildTable Then
			'    If Not IsReportChildTable(pobjTableView.TableID) Then
			'      If pobjTableView.TableID <> mlngCustomReportsParent1Table Then
			'        If pobjTableView.TableID <> mlngCustomReportsParent2Table Then
			
			If (pobjTableView.TableID = lngTableID) Then
				If (pobjTableView.ViewName <> mstrBaseTableRealSource) Then
					mstrSQLJoin = mstrSQLJoin & " LEFT OUTER JOIN " & pobjTableView.RealSource & " ON " & mstrBaseTableRealSource & ".ID = " & pobjTableView.RealSource & ".ID"
				End If
			Else
				If (pobjTableView.ViewName <> mstrBaseTableRealSource) Then
					mstrSQLJoin = mstrSQLJoin & " LEFT OUTER JOIN " & pobjTableView.RealSource & " ON " & mstrBaseTableRealSource & ".ID_" & CStr(pobjTableView.TableID) & " = " & pobjTableView.RealSource & ".ID"
				End If
				
				
				'            'JPD 20031119 Fault 7660
				'            ' This is a parent of a child of the report base table, not explicitly
				'            ' included in the report, but referred to by a child table calculation.
				'            For iLoop2 = 1 To UBound(mlngTableViews, 2)
				'              If mlngTableViews(1, iLoop2) = 0 Then
				'                If mclsGeneral.IsAChildOf(mlngTableViews(2, iLoop2), pobjTableView.TableID) Then
				'                  Set objChildTable = gcoTablePrivileges.FindTableID(mlngTableViews(2, iLoop2))
				'
				'                  sOtherParentJoinCode = sOtherParentJoinCode & _
				''                    " LEFT OUTER JOIN " & pobjTableView.RealSource & _
				''                    " ON " & objChildTable.RealSource & ".ID_" & CStr(pobjTableView.TableID) & " = " & pobjTableView.RealSource & ".ID"
				'                  Exit For
				'                End If
				'              End If
				'            Next iLoop2
			End If
			'        End If
			'      End If
			'    End If
			
			'    'If (pobjTableView.TableID = mlngCustomReportsParent1Table) Or _
			''    '(pobjTableView.TableID = mlngCustomReportsParent2Table) Then
			'      mstrSQLJoin = mstrSQLJoin & _
			''           " LEFT OUTER JOIN " & pobjTableView.RealSource & _
			''           " ON " & mstrBaseTableRealSource & ".ID_" & pobjTableView.TableID & " = " & pobjTableView.RealSource & ".ID"
			'    'End If
		Next pintLoop
		
		''  'Now do the childview(s) bit, if required
		''
		''  lngTempChildID = 0
		''  lngTempMaxRecords = 0
		''  lngTempFilterID = 0
		''
		'''  If mlngCustomReportsChildTable > 0 Then
		''  If miChildTablesCount > 0 Then
		''    For i = 0 To UBound(mvarChildTables, 2) Step 1
		''      lngTempChildID = mvarChildTables(0, i)
		''      lngTempFilterID = mvarChildTables(1, i)
		''      lngTempOrderID = mvarChildTables(5, i)
		''      lngTempMaxRecords = mvarChildTables(2, i)
		''
		''      pblnChildUsed = False
		''
		'''      ' are any child fields in the report ? # 12/06/00 RH - FAULT 419
		'''      For pintLoop = 1 To UBound(mvarColDetails, 2)
		'''        If GetTableIDFromColumn(CLng(mvarColDetails(12, pintLoop))) = lngTempChildID Then
		'''          pblnChildUsed = True
		'''          Exit For
		'''        End If
		'''      Next pintLoop
		''
		''      'TM20020409 Fault 3745 - Only do the join if columns from the table are used.
		''      pblnChildUsed = IsChildTableUsed(lngTempChildID)
		''
		''      mvarChildTables(4, i) = pblnChildUsed
		''      If pblnChildUsed Then miUsedChildCount = miUsedChildCount + 1
		''
		''      If pblnChildUsed = True Then
		''
		'''        Set objChildTable = gcoTablePrivileges.FindTableID(mlngCustomReportsChildTable)
		''        Set objChildTable = gcoTablePrivileges.FindTableID(lngTempChildID)
		''
		''        If objChildTable.AllowSelect Then
		''          sChildJoinCode = sChildJoinCode & " LEFT OUTER JOIN " & objChildTable.RealSource & _
		'''                           " ON " & mstrBaseTableRealSource & ".ID = " & _
		'''                           objChildTable.RealSource & ".ID_" & mlngCustomReportsBaseTable
		''
		''          sChildJoinCode = sChildJoinCode & " AND " & objChildTable.RealSource & ".ID IN"
		''
		'''          sChildJoinCode = sChildJoinCode & _
		''''          " (SELECT TOP" & IIf(mlngCustomReportsChildMaxRecords = 0, " 100 PERCENT", " " & mlngCustomReportsChildMaxRecords) & _
		''''          " " & objChildTable.RealSource & ".ID FROM " & objChildTable.RealSource
		''
		''          'TM20020328 Fault 3714 - ensure the maxrecords is >= zero.
		''          sChildJoinCode = sChildJoinCode & _
		'''          " (SELECT TOP" & IIf(lngTempMaxRecords < 1, " 100 PERCENT", " " & lngTempMaxRecords) & _
		'''          " " & objChildTable.RealSource & ".ID FROM " & objChildTable.RealSource
		''
		''          ' Now the child order by bit - done here in case tables need to be joined.
		'''          Set rsTemp = datGeneral.GetOrderDefinition(datGeneral.GetDefaultOrder(mlngCustomReportsChildTable))
		''          If lngTempOrderID > 0 Then
		''            Set rsTemp = datGeneral.GetOrderDefinition(lngTempOrderID)
		''          Else
		''            Set rsTemp = datGeneral.GetOrderDefinition(datGeneral.GetDefaultOrder(lngTempChildID))
		''          End If
		''
		''          sChildOrderString = DoChildOrderString(rsTemp, sChildJoin, lngTempChildID)
		''          Set rsTemp = Nothing
		''
		''          sChildJoinCode = sChildJoinCode & sChildJoin
		''
		''          sChildJoinCode = sChildJoinCode & _
		'''            " WHERE (" & objChildTable.RealSource & ".ID_" & mlngCustomReportsBaseTable & _
		'''            " = " & mstrBaseTableRealSource & ".ID)"
		''
		''          ' is the child filtered ?
		''
		''  '        If mlngCustomReportsChildFilterID > 0 Then
		''          If lngTempFilterID > 0 Then
		'''            blnOK = datGeneral.FilteredIDs(mlngCustomReportsChildFilterID, strFilterIDs, mvarPrompts)
		''            blnOK = datGeneral.FilteredIDs(lngTempFilterID, strFilterIDs, mvarPrompts)
		''
		''            ' Generate any UDFs that are used in this filter
		''            If blnOK Then
		''              datGeneral.FilterUDFs lngTempFilterID, mastrUDFsRequired()
		''            End If
		''
		''            If blnOK Then
		''              sChildJoinCode = sChildJoinCode & " AND " & _
		'''                objChildTable.RealSource & ".ID IN (" & strFilterIDs & ")"
		''            Else
		''              ' Permission denied on something in the filter.
		'''              mstrErrorString = "You do not have permission to use the '" & datGeneral.GetFilterName(mlngCustomReportsChildFilterID) & "' filter."
		''              mstrErrorString = "You do not have permission to use the '" & datGeneral.GetFilterName(lngTempFilterID) & "' filter."
		''              GenerateSQLJoin = False
		''              Exit Function
		''            End If
		''          End If
		''
		''        End If
		''
		''        sChildJoinCode = sChildJoinCode & IIf(Len(sChildOrderString) > 0, " ORDER BY " & sChildOrderString & ")", "")
		''
		''      End If
		''    Next i
		''  End If
		''
		'  mstrSQLJoin = mstrSQLJoin & sChildJoinCode & IIf(Len(sChildOrderString) > 0, " ORDER BY " & sChildOrderString & ")", "")
		'  mstrSQLJoin = mstrSQLJoin & sChildJoinCode
		'  mstrSQLJoin = mstrSQLJoin & sOtherParentJoinCode
		
		GenerateSQLJoin = True
		Exit Function
		
GenerateSQLJoin_ERROR: 
		
		GenerateSQLJoin = False
		mstrErrorString = "Error in GenerateSQLJoin." & vbNewLine & Err.Description
		
	End Function
	
	Private Function GenerateSQLOrderBy(ByRef plngSortOrderID As Integer, ByRef piSortDirection As Short) As Boolean
		
		' Purpose : Returns order by string from the sort order array
		
		On Error GoTo GenerateSQLOrderBy_ERROR
		
		' get the sort order - this is stored as decimal, but represents a 4 digit binary value
		' e.g. 5 = 0101 in binary, which represents sort orders 1 & 3=desc, 2 = asc.
		' Only first digit (leftmost) used in single axis charting.
		' Digit 1 = Horizontal Data Sort order
		' Digit 2 = Vertical Data Sort order
		' Digit 3 = 'Sort by Aggregate' tickbox
		' Digit 4 = 'Sort by Aggregate' sort order
		
		Dim pstrBinaryString As String
		pstrBinaryString = DecToBin(plngSortOrderID, 4)
		
		If Mid(pstrBinaryString, 3, 1) = "1" Then ' The third switch is for 'Sort by Aggregate'
			piSortDirection = Val(Right(pstrBinaryString, 1))
			mstrSQLOrderBy = "[AGGREGATE] " & IIf(piSortDirection = 0, "ASC", "DESC")
		Else
			piSortDirection = Val(Mid(pstrBinaryString, 1, 1))
			mstrSQLOrderBy = "[COLUMN] " & IIf(piSortDirection = 0, "ASC", "DESC")
		End If
		
		If Len(mstrSQLOrderBy) > 0 Then mstrSQLOrderBy = " ORDER BY " & mstrSQLOrderBy
		
		GenerateSQLOrderBy = True
		Exit Function
		
GenerateSQLOrderBy_ERROR: 
		
		GenerateSQLOrderBy = False
		
	End Function
	
	
	Private Function GenerateSQLWhere(ByRef lngTableID As Integer, ByRef lngFilterID As Integer) As Boolean
		
		' Purpose : Generate the where clauses that cope with the joins
		'           NB Need to add the where clauses for filters/picklists etc
		
		On Error GoTo GenerateSQLWhere_ERROR
		
		Dim pintLoop As Short
		Dim pobjTableView As CTablePrivilege
		Dim prstTemp As New ADODB.Recordset
		Dim pstrPickListIDs As String
		Dim blnOK As Boolean
		Dim strFilterIDs As String
		Dim objExpr As clsExprExpression
		Dim pstrParent1PickListIDs As String
		Dim pstrParent2PickListIDs As String
		
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
						mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " OR ", " WHERE (") & mstrBaseTableRealSource & ".ID IN (SELECT ID FROM " & pobjTableView.RealSource & ")"
					End If
					
				Next pintLoop
				
				If Len(mstrSQLWhere) > 0 Then mstrSQLWhere = mstrSQLWhere & ")"
				
			End If
			
		End If
		
		
		If lngFilterID > 0 Then
			
			blnOK = datGeneral.FilteredIDs(lngFilterID, strFilterIDs, mvarPrompts)
			
			' Generate any UDFs that are used in this filter
			If blnOK Then
				datGeneral.FilterUDFs(lngFilterID, mastrUDFsRequired)
			End If
			
			If blnOK Then
				mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & mstrSQLFrom & ".ID IN (" & strFilterIDs & ")"
			Else
				' Permission denied on something in the filter.
				mstrErrorString = "You do not have permission to use the '" & datGeneral.GetFilterName(lngFilterID) & "' filter."
				GenerateSQLWhere = False
				Exit Function
			End If
		End If
		
		'UPGRADE_NOTE: Object prstTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		prstTemp = Nothing
		
		GenerateSQLWhere = True
		Exit Function
		
GenerateSQLWhere_ERROR: 
		
		GenerateSQLWhere = False
		mstrErrorString = "Error in GenerateSQLWhere." & vbNewLine & Err.Description
		
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
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		' Initialise the the classes/arrays to be used
		mclsData = New clsDataAccess
		mclsGeneral = New clsGeneral
		' ReDim mvarSortOrder(2, 0)
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		' Clear references to classes and clear collection objects
		'UPGRADE_NOTE: Object mclsData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mclsData = Nothing
		'UPGRADE_NOTE: Object mclsGeneral may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mclsGeneral = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Private Function DecToBin(ByRef DeciValue As Integer, Optional ByRef NoOfBits As Short = 8) As String
		
		Dim i As Short 'make sure there are enough bits to contain the number
		Do While DeciValue > (2 ^ NoOfBits) - 1
			NoOfBits = NoOfBits + 8
		Loop 
		DecToBin = vbNullString
		'build the string
		For i = 0 To (NoOfBits - 1)
			DecToBin = CStr(CShort(DeciValue And 2 ^ i) / 2 ^ i) & DecToBin
		Next i
	End Function
	
	Private Function GetTableIDFromColumn(ByRef lngColumnID As Integer) As Integer
		
		' Purpose : To return the table id for which the given column belongs
		
		Dim rsInfo As ADODB.Recordset
		Dim strSQL As String
		
		strSQL = "SELECT ASRSysTables.TableID " & "FROM ASRSysColumns JOIN ASRSysTables " & "ON (ASRSysTables.TableID = ASRSysColumns.TableID) " & "WHERE ColumnID = " & CStr(lngColumnID)
		
		rsInfo = mclsData.OpenRecordset(strSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		If rsInfo.BOF And rsInfo.EOF Then
			GetTableIDFromColumn = 0
		Else
			GetTableIDFromColumn = rsInfo.Fields("TableID").Value
		End If
		
		'UPGRADE_NOTE: Object rsInfo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsInfo = Nothing
		
	End Function
	
	
	Public Sub resetGlobals()
		
		' reset the Global variables as they interfere with other users.
		'UPGRADE_NOTE: Object gcoTablePrivileges may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		gcoTablePrivileges = Nothing
		'UPGRADE_NOTE: Object gcolColumnPrivilegesCollection may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		gcolColumnPrivilegesCollection = Nothing
		
	End Sub
End Class