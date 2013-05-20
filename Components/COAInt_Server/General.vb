Option Strict Off
Option Explicit On
Friend Class clsGeneral
	
	Private datData As clsDataAccess
	Private UI As clsUI
	Private asViewTables() As String
	
	Public Function ConvertNumberForSQL(ByVal strInput As String) As String
		'Get a number in the correct format for a SQL string
		'(e.g. on french systems replace decimal comma for a decimal point)
		ConvertNumberForSQL = Replace(strInput, UI.GetSystemDecimalSeparator, ".")
	End Function
	
	Public Function ConvertNumberForDisplay(ByVal strInput As String) As String
		'Get a number in the correct format for display
		'(e.g. on french systems replace decimal point for a decimal comma)
		ConvertNumberForDisplay = Replace(strInput, ".", UI.GetSystemDecimalSeparator)
	End Function
	
	Public Function ConvertSQLDateToSystemFormat(ByRef pstrDateString As String) As Date
		
		Dim dtTemp As Date
		Dim strDateFormat As String
		Dim lngDay_CR As Integer
		Dim lngMonth_CR As Integer
		Dim lngYear_CR As Integer
		
		Dim blnDateComplete As Boolean
		Dim blnMonthDone As Boolean
		Dim blnDayDone As Boolean
		Dim blnYearDone As Boolean
		
		Dim strShortDate As String
		
		Dim strDateSeparator As String
		
		Dim i As Short
		
		' eg. DateFormat = "mm/dd/yyyy"
		'     Calendar   = "dd/mm/yyyy"
		'     DateString = "06/02/2000"
		'     Compare to = 02/06/2000
		
		strDateFormat = UI.GetSystemDateFormat
		
		strDateSeparator = UI.GetSystemDateSeparator
		
		blnDateComplete = False
		blnMonthDone = False
		blnDayDone = False
		blnYearDone = False
		
		'Assume American Date format mm/dd/yyyy
		lngMonth_CR = CInt(Mid(pstrDateString, 1, 2))
		lngDay_CR = CInt(Mid(pstrDateString, 4, 2))
		lngYear_CR = CInt(Mid(pstrDateString, 7, 4))
		
		strShortDate = vbNullString
		
		For i = 1 To Len(strDateFormat) Step 1
			
			If (LCase(Mid(strDateFormat, i, 1)) = "d") And (Not blnDayDone) Then
				strShortDate = strShortDate & LCase(Mid(strDateFormat, i, 1))
				blnDayDone = True
			End If
			
			If (LCase(Mid(strDateFormat, i, 1)) = "m") And (Not blnMonthDone) Then
				strShortDate = strShortDate & LCase(Mid(strDateFormat, i, 1))
				blnMonthDone = True
			End If
			
			If (LCase(Mid(strDateFormat, i, 1)) = "y") And (Not blnYearDone) Then
				strShortDate = strShortDate & LCase(Mid(strDateFormat, i, 1))
				blnYearDone = True
			End If
			
			If blnDayDone And blnMonthDone And blnYearDone Then
				blnDateComplete = True
				Exit For
			End If
			
		Next i
		
		Select Case strShortDate
			Case "dmy" : dtTemp = CDate(lngDay_CR & strDateSeparator & lngMonth_CR & strDateSeparator & lngYear_CR)
			Case "mdy" : dtTemp = CDate(lngMonth_CR & strDateSeparator & lngDay_CR & strDateSeparator & lngYear_CR)
			Case "ydm" : dtTemp = CDate(lngYear_CR & strDateSeparator & lngDay_CR & strDateSeparator & lngMonth_CR)
			Case "myd" : dtTemp = CDate(lngMonth_CR & strDateSeparator & lngYear_CR & strDateSeparator & lngDay_CR)
			Case "ymd" : dtTemp = CDate(lngYear_CR & strDateSeparator & lngMonth_CR & strDateSeparator & lngDay_CR)
		End Select
		
		ConvertSQLDateToSystemFormat = dtTemp
		
	End Function
	
	Public Function EnableUDFFunctions() As Boolean
		
		Dim sSQL As String
		Dim rsUser As ADODB.Recordset
		
		sSQL = "exec master..xp_msver"
		
		rsUser = datData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		rsUser.MoveNext()
		
		Select Case Val(rsUser.Fields(3).Value)
			Case Is >= 8
				EnableUDFFunctions = True
			Case Else
				EnableUDFFunctions = False
		End Select
		
		rsUser.Close()
		'UPGRADE_NOTE: Object rsUser may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsUser = Nothing
		
	End Function
	
	Public Function FilterUDFs(ByRef plngFilterID As Integer, ByRef pastrUDFs() As String) As Boolean
		
		On Error GoTo ErrorTrap
		
		' Return a string describing the record IDs from the given table
		' that satisfy the given criteria.
		Dim fOK As Boolean
		Dim objExpr As clsExprExpression
		
		fOK = True
		
		objExpr = New clsExprExpression
		With objExpr
			.Initialise(0, plngFilterID, modExpression.ExpressionTypes.giEXPR_RUNTIMEFILTER, modExpression.ExpressionValueTypes.giEXPRVALUE_LOGIC)
			.UDFFilterCode(pastrUDFs, True, True)
		End With
		'UPGRADE_NOTE: Object objExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objExpr = Nothing
		
		FilterUDFs = fOK
		
TidyUpAndExit: 
		Exit Function
ErrorTrap: 
		
	End Function
	
	
	Public Function GetRecordsInTransaction(ByRef sSQL As String) As ADODB.Recordset
		' Return the required STATIC/read-only recordset.
		' This is useful when getting a recordset in the middle of a transaction.
		' An error occurs when getting more than one forward-only, read-only recordset in the middle
		' of a transaction.
		GetRecordsInTransaction = datData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
		
	End Function
	
	
	Public Function FilteredIDs(ByRef plngExprID As Integer, ByRef psIDSQL As String, Optional ByRef paPrompts As Object = Nothing) As Boolean
		' Return a string describing the record IDs from the given table
		' that satisfy the given criteria.
		Dim fOK As Boolean
		Dim objExpr As clsExprExpression
		
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		
		objExpr = New clsExprExpression
		With objExpr
			' Initialise the filter expression object.
			fOK = .Initialise(0, plngExprID, modExpression.ExpressionTypes.giEXPR_RUNTIMEFILTER, modExpression.ExpressionValueTypes.giEXPRVALUE_LOGIC)
			
			If fOK Then
				fOK = objExpr.RuntimeFilterCode(psIDSQL, True, False, paPrompts)
			End If
			
		End With
		'UPGRADE_NOTE: Object objExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objExpr = Nothing
		
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		
		FilteredIDs = fOK
		
	End Function
	
	Public Function GetOrderDefinition(ByRef plngOrderID As Integer) As ADODB.Recordset
		' Return a recordset of the order items (both Find Window and Sort Order columns)
		' for the given order.
		Dim sSQL As String
		Dim rsInfo As ADODB.Recordset
		
		sSQL = "EXEC sp_ASRGetOrderDefinition " & plngOrderID
		rsInfo = datData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		GetOrderDefinition = rsInfo
		
	End Function
	
	
	Public Property Username() As String
		Get
			Username = gsUsername
		End Get
		Set(ByVal Value As String)
			
			gsUsername = Value
			GetActualUserDetails()
			
		End Set
	End Property
	
	Public Function GetAllTables() As ADODB.Recordset
		' Return a recordset of the user-defined tables.
		Dim sSQL As String
		Dim rsTables As ADODB.Recordset
		
		sSQL = "SELECT tableID, tableName, tableType, defaultOrderID, recordDescExprID" & " FROM ASRSysTables"
		
		rsTables = datData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		GetAllTables = rsTables
		
	End Function
	
	Public Sub GetColumnPermissions(ByRef pobjColumnPrivileges As CColumnPrivileges)
	End Sub
	
	Public Function GetAllViews() As ADODB.Recordset
		Dim sSQL As String
		Dim rsViews As ADODB.Recordset
		
		sSQL = "SELECT ASRSysViews.viewID, ASRSysViews.viewName, ASRSysTables.tableID," & " ASRSysTables.tableName, ASRSysTables.tableType, ASRSysTables.defaultOrderID, ASRSysTables.recordDescExprID" & " FROM ASRSysViews" & " INNER JOIN ASRSysTables ON ASRSysViews.viewTableID = ASRSysTables.tableID"
		rsViews = datData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		GetAllViews = rsViews
		
	End Function
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
		datData = New clsDataAccess
		UI = New clsUI
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
  Public Function GetValueForRecordIndependantCalc(ByRef lngExprID As Integer, Optional ByRef pvarPrompts As Object = Nothing) As Object

    Dim objExpr As clsExprExpression
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim fOK As Boolean
    Dim lngViews() As Integer

    ReDim lngViews(0)

    On Error GoTo LocalErr

    'UPGRADE_WARNING: Couldn't resolve default property of object GetValueForRecordIndependantCalc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    GetValueForRecordIndependantCalc = vbNullString

    objExpr = New clsExprExpression
    With objExpr
      ' Initialise the filter expression object.
      fOK = .Initialise(0, lngExprID, modExpression.ExpressionTypes.giEXPR_RECORDINDEPENDANTCALC, modExpression.ExpressionValueTypes.giEXPRVALUE_UNDEFINED)

      If fOK Then
        fOK = objExpr.RuntimeCalculationCode(lngViews, strSQL, True, False, pvarPrompts)
      End If

      If fOK Then
        rsTemp = GetReadOnlyRecords("SELECT " & strSQL)
        If Not rsTemp.BOF And Not rsTemp.EOF Then
          'UPGRADE_WARNING: Couldn't resolve default property of object GetValueForRecordIndependantCalc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          GetValueForRecordIndependantCalc = rsTemp.Fields(0).Value
        End If
      End If

    End With
    'UPGRADE_NOTE: Object objExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    objExpr = Nothing


    Exit Function

LocalErr:

  End Function
	
	Public Function GetActualLogin() As String
		
		Dim sSQL As String
		Dim cmdLoginInfo As New ADODB.Command
		Dim prmActutalLogin As New ADODB.Parameter
		
		cmdLoginInfo.CommandText = "spASRIntGetActualLogin"
		cmdLoginInfo.CommandType = 4
		cmdLoginInfo.let_ActiveConnection(gADOCon)
		
		prmActutalLogin = cmdLoginInfo.CreateParameter("ActualLogin", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 250)
		cmdLoginInfo.Parameters.Append(prmActutalLogin)
		
		cmdLoginInfo.ActiveConnection.Errors.Clear()
		cmdLoginInfo.Execute()
		
		GetActualLogin = cmdLoginInfo.Parameters("ActualLogin").Value
		
		'UPGRADE_NOTE: Object prmActutalLogin may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		prmActutalLogin = Nothing
		'UPGRADE_NOTE: Object cmdLoginInfo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		cmdLoginInfo = Nothing
		
	End Function
	
	Public Function GetActualUserDetails() As String
		
		Dim sSQL As String
		Dim cmdUserInfo As New ADODB.Command
		Dim prmActualUser As New ADODB.Parameter
		Dim prmActualUserGroup As New ADODB.Parameter
		Dim prmActualUserGroupID As New ADODB.Parameter
		
		cmdUserInfo.CommandText = "spASRIntGetActualUserDetails"
		cmdUserInfo.CommandType = 4
		cmdUserInfo.let_ActiveConnection(gADOCon)
		
    prmActualUser = cmdUserInfo.CreateParameter("psUsername", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 250)
		cmdUserInfo.Parameters.Append(prmActualUser)
		
    prmActualUserGroup = cmdUserInfo.CreateParameter("psUserGroup", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 250)
		cmdUserInfo.Parameters.Append(prmActualUserGroup)
		
    prmActualUserGroupID = cmdUserInfo.CreateParameter("piUserGroupID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamOutput)
		cmdUserInfo.Parameters.Append(prmActualUserGroupID)
		
		cmdUserInfo.ActiveConnection.Errors.Clear()
		cmdUserInfo.Execute()
		
    gsActualLogin = cmdUserInfo.Parameters("psUsername").Value
    gsUserGroup = cmdUserInfo.Parameters("psUserGroup").Value
		
		'UPGRADE_NOTE: Object prmActualUser may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		prmActualUser = Nothing
		'UPGRADE_NOTE: Object prmActualUserGroup may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		prmActualUserGroup = Nothing
		'UPGRADE_NOTE: Object prmActualUserGroupID may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		prmActualUserGroupID = Nothing
		'UPGRADE_NOTE: Object cmdUserInfo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		cmdUserInfo = Nothing
		
	End Function
	
	
	Public Function GetUserDetails() As String
		
		Dim sSQL As String
		Dim rsUser As ADODB.Recordset
		
		sSQL = "exec sp_helpuser '" & Replace(gsActualLogin, "'", "''") & "'"
		
		rsUser = datData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		'MH20031107 Fault 5627
		'If rsUser!GroupName = "db_owner" Then
		'  rsUser.MoveNext
		'End If
		Do While rsUser.Fields("GroupName").Value = "db_owner" Or LCase(Left(rsUser.Fields("GroupName").Value, 6)) = "asrsys"
			rsUser.MoveNext()
		Loop 
		GetUserDetails = rsUser.Fields("GroupName").Value
		
		rsUser.Close()
		'UPGRADE_NOTE: Object rsUser may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsUser = Nothing
		
	End Function
	
	Public Function GetRecords(ByRef sSQL As String) As ADODB.Recordset
		' Return the required forward-only/read-only recordset.
		'  Set GetRecords = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
		
		' JPD - Changed cursor-type to static.
		' The SQL Server ODBC driver uses a special cursor when the cursor is
		' forward-only, read-only, and the ODBC rowset size is one.
		' The cursor is called a "firehose" cursor because it is the fastest way to retrieve the data.
		' Unfortunately, a side affect of the cursor is that it only permits one active recordset per connection.
		' To get around this we'll try using a STATIC cursor.
		GetRecords = datData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
		
	End Function
	
	
	Public Function GetReadOnlyRecords(ByRef sSQL As String) As ADODB.Recordset
		' Return the required dynamic/read-only recordset.
		'  Set GetReadOnlyRecords = datData.OpenRecordset(sSQL, adOpenDynamic, adLockReadOnly)
		
		' JPD 7/6/00 Changed the cursor type from dynamic to static.
		' A dynamic/read-only recordset opened with a sql select statement that includes the 'distinct' parameter,
		' and returns a single record will automatically be set to be a forwardonly/read-only (firehose) recordet.
		GetReadOnlyRecords = datData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
		
	End Function
	
	Public Function GetTableName(ByVal plngTableID As Integer) As String
		Dim rsTable As ADODB.Recordset
		Dim sSQL As String
		
		sSQL = "SELECT tableName " & " FROM ASRSysTables " & " WHERE tableID=" & plngTableID
		
		rsTable = datData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic)
		With rsTable
			If Not (.BOF And .EOF) Then
				GetTableName = Trim(.Fields(0).Value)
			Else
				GetTableName = vbNullString
			End If
			
			.Close()
		End With
		
		'UPGRADE_NOTE: Object rsTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsTable = Nothing
		
	End Function
	
	Public Function GetFilterName(ByVal lFilterID As Integer) As String
		Dim rsFilter As ADODB.Recordset
		Dim sSQL As String
		
		sSQL = "SELECT name " & "FROM ASRSysExpressions " & "WHERE ExprID=" & lFilterID
		
		rsFilter = datData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic)
		With rsFilter
			If Not (.BOF And .EOF) Then
				GetFilterName = Trim(.Fields(0).Value)
			Else
				GetFilterName = vbNullString
			End If
			.Close()
		End With
		
		'UPGRADE_NOTE: Object rsFilter may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsFilter = Nothing
		
	End Function
	Public Function GetPicklistName(ByVal lPicklistID As Integer) As String
		Dim rsPicklist As ADODB.Recordset
		Dim sSQL As String
		
		sSQL = "SELECT Name " & "FROM ASRSysPicklistName " & "WHERE PicklistID=" & lPicklistID
		
		rsPicklist = datData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic)
		With rsPicklist
			If Not (.BOF And .EOF) Then
				GetPicklistName = Trim(.Fields(0).Value)
			Else
				GetPicklistName = vbNullString
			End If
			.Close()
		End With
		
		'UPGRADE_NOTE: Object rsPicklist may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsPicklist = Nothing
		
	End Function
	
	Public Function GetRecDescExprID(ByVal TableID As Integer) As String
		' JPD - Return the Record Description Expression ID for the given table.
		Dim rsTable As ADODB.Recordset
		Dim sSQL As String
		
		sSQL = "SELECT recordDescExprID " & "FROM ASRSysTables " & "WHERE TableID=" & TableID
		
		rsTable = datData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic)
		With rsTable
			If Not (.BOF And .EOF) Then
				GetRecDescExprID = .Fields(0).Value
			Else
				GetRecDescExprID = CStr(0)
			End If
			.Close()
		End With
		
		'UPGRADE_NOTE: Object rsTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsTable = Nothing
		
	End Function
	
	Public Function GetDataType(ByRef lTableID As Integer, ByRef lColumnID As Integer) As Integer
		
		Dim sSQL As String
		Dim rsData As ADODB.Recordset
		
		sSQL = "Select datatype From ASRSysColumns Where columnid= " & lColumnID & " And tableID = " & lTableID
		rsData = datData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		If Not rsData.BOF And Not rsData.EOF Then
			GetDataType = rsData.Fields(0).Value
		End If
		
		rsData.Close()
		'UPGRADE_NOTE: Object rsData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsData = Nothing
		
	End Function
	
	Public Function GetColumnTable(ByRef plngColumnID As Integer) As Integer
		' Return the table id of the given column.
		Dim sSQL As String
		Dim rsData As ADODB.Recordset
		
		sSQL = "SELECT tableID" & " FROM ASRSysColumns" & " WHERE columnID = " & Trim(Str(plngColumnID))
		rsData = datData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		If Not rsData.BOF And Not rsData.EOF Then
			GetColumnTable = rsData.Fields("TableID").Value
		Else
			GetColumnTable = 0
		End If
		
		rsData.Close()
		'UPGRADE_NOTE: Object rsData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsData = Nothing
		
	End Function
	
	Public Function GetDefaultOrder(ByRef plngTableID As Integer) As Integer
		' Return the default order ID for the given table.
		Dim sSQL As String
		Dim rsInfo As ADODB.Recordset
		
		sSQL = "SELECT defaultOrderID" & " FROM ASRSysTables" & " WHERE tableID = " & Trim(Str(plngTableID))
		rsInfo = datData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		If Not (rsInfo.BOF And rsInfo.EOF) Then
			GetDefaultOrder = rsInfo.Fields(0).Value
		Else
			GetDefaultOrder = 0
		End If
		rsInfo.Close()
		'UPGRADE_NOTE: Object rsInfo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsInfo = Nothing
		
	End Function
	
	Public Function GetOrder(ByRef lOrderID As Integer) As ADODB.Recordset
		
		Dim sSQL As String
		Dim rsOrder As ADODB.Recordset
		
		sSQL = "SELECT * " & "FROM ASRSysOrders " & "WHERE OrderID=" & lOrderID
		rsOrder = datData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		GetOrder = rsOrder
		
	End Function
	
	Public Function GetColumnName(ByRef plngColumnID As Integer) As String
		' Return the name of the given column.
		Dim sSQL As String
		Dim rsTemp As ADODB.Recordset
		
		If plngColumnID = 0 Then
			GetColumnName = ""
			'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsTemp = Nothing
			Exit Function
		End If
		
		sSQL = "SELECT columnName FROM ASRSysColumns WHERE columnID = " & plngColumnID
		rsTemp = datData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		GetColumnName = rsTemp.Fields(0).Value
		rsTemp.Close()
		'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsTemp = Nothing
		
	End Function
	
	Public Function GetModuleParameter(ByRef psModuleKey As String, ByRef psParameterKey As String) As String
		' Return the value of the given module parameter.
		Dim sSQL As String
		Dim rsModule As ADODB.Recordset
		
		sSQL = "SELECT parameterValue" & " FROM ASRSysModuleSetup" & " WHERE moduleKey = '" & psModuleKey & "'" & " AND parameterKey = '" & psParameterKey & "'"
		rsModule = datData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		If Not (rsModule.BOF And rsModule.EOF) Then
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If IsDbNull(rsModule.Fields("parameterValue").Value) Then
				GetModuleParameter = vbNullString
			Else
				GetModuleParameter = rsModule.Fields("parameterValue").Value
			End If
		Else
			GetModuleParameter = vbNullString
		End If
		
		'UPGRADE_NOTE: Object rsModule may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsModule = Nothing
		
	End Function
	
	Public Function UniqueSQLObjectName(ByRef strPrefix As String, ByRef intType As Short) As String
		
		'TM20020530 Fault 3756 - function altered as the sp needs to insert a record into a table
		'before returning a value, so collect the returned parameter rather than a recordset.
		
		Dim cmdUniqObj As New ADODB.Command
		Dim pmADO As ADODB.Parameter
		
		With cmdUniqObj
			.CommandText = "sp_ASRUniqueObjectName"
			.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
			.CommandTimeout = 0
			.ActiveConnection = gADOCon
			
			pmADO = .CreateParameter("UniqueObjectName", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 255)
      .Parameters.Append(pmADO)

			pmADO = .CreateParameter("Prefix", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 255)
			.Parameters.Append(pmADO)
			pmADO.Value = strPrefix
			
			pmADO = .CreateParameter("Type", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
			.Parameters.Append(pmADO)
			pmADO.Value = intType
			
			'UPGRADE_NOTE: Object pmADO may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			pmADO = Nothing
			
			.Execute()
			
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			UniqueSQLObjectName = IIf(IsDbNull(.Parameters(0).Value), vbNullString, .Parameters(0).Value)
			
		End With
		
		'UPGRADE_NOTE: Object cmdUniqObj may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		cmdUniqObj = Nothing
		
	End Function
	
  Public Function DropUniqueSQLObject(ByVal sSQLObjectName As String, ByRef iType As Short) As Boolean

    On Error GoTo ErrorTrap

    Dim cmdUniqObj As New ADODB.Command
    Dim pmADO As ADODB.Parameter

    If Len(sSQLObjectName) > 0 Then
      With cmdUniqObj
        .CommandText = "sp_ASRDropUniqueObject"
        .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        .CommandTimeout = 0
        .ActiveConnection = gADOCon

        pmADO = .CreateParameter("UniqueObjectName", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 255)
        .Parameters.Append(pmADO)
        pmADO.Value = sSQLObjectName

        pmADO = .CreateParameter("Type", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
        .Parameters.Append(pmADO)
        pmADO.Value = iType

        'UPGRADE_NOTE: Object pmADO may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        pmADO = Nothing

        .Execute()
      End With
    End If

    DropUniqueSQLObject = True

TidyUpAndExit:
    'UPGRADE_NOTE: Object cmdUniqObj may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    cmdUniqObj = Nothing
    Exit Function

ErrorTrap:
    DropUniqueSQLObject = False
    GoTo TidyUpAndExit

  End Function
	
  Public Function DoesColumnUseSeparators(ByVal plngColumnID As Integer) As Boolean

    ' Returns whether the column uses 1000 separators...
    Dim sSQL As String
    Dim rsData As ADODB.Recordset

    sSQL = "SELECT Use1000Separator FROM ASRSysColumns" & " WHERE columnID = " & Trim(Str(plngColumnID))
    rsData = datData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

    If Not rsData.BOF And Not rsData.EOF Then
      'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
      DoesColumnUseSeparators = IIf(Not IsDBNull(rsData.Fields("Use1000Separator").Value), rsData.Fields("Use1000Separator").Value, False)
    Else
      DoesColumnUseSeparators = False
    End If

    rsData.Close()
    'UPGRADE_NOTE: Object rsData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsData = Nothing

  End Function
	
	' Returns the amount of decimals that are specificed for a column
  Public Function GetDecimalsSize(ByVal plngColumnID As Integer) As Short

    Dim sSQL As String
    Dim rsData As ADODB.Recordset

    sSQL = "SELECT Decimals FROM ASRSysColumns" & " WHERE columnID = " & Trim(Str(plngColumnID))
    rsData = datData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

    If Not rsData.BOF And Not rsData.EOF Then
      'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
      GetDecimalsSize = IIf(Not IsDBNull(rsData.Fields("Decimals").Value), rsData.Fields("Decimals").Value, 0)
    Else
      GetDecimalsSize = 0
    End If

    rsData.Close()
    'UPGRADE_NOTE: Object rsData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsData = Nothing

  End Function
	
  Public Function IsAChildOf(ByVal lTestTableID As Integer, ByVal lBaseTableID As Integer) As Boolean

    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String

    strSQL = "SELECT * FROM ASRSysRelations WHERE ParentID = " & lBaseTableID & " AND ChildID = " & lTestTableID

    rsTemp = datData.OpenRecordset(strSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

    If rsTemp.BOF And rsTemp.EOF Then
      IsAChildOf = False
    Else
      IsAChildOf = True
    End If

    'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsTemp = Nothing

  End Function
	
  Public Function IsAParentOf(ByVal lTestTableID As Integer, ByVal lBaseTableID As Integer) As Boolean

    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String

    strSQL = "SELECT * FROM ASRSysRelations WHERE ChildID = " & lBaseTableID & " AND ParentID = " & lTestTableID

    rsTemp = datData.OpenRecordset(strSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

    If rsTemp.BOF And rsTemp.EOF Then
      IsAParentOf = False
    Else
      IsAParentOf = True
    End If

    'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsTemp = Nothing

  End Function
	
  Public Function DateColumn(ByVal strType As String, ByVal lngTableID As Integer, ByVal lngColumnID As Integer) As Boolean

    'MH20000705
    Dim objCalcExpr As New clsExprExpression

    DateColumn = False


    Select Case strType
      Case "C" 'Column
        DateColumn = (GetDataType(lngTableID, lngColumnID) = Declarations.SQLDataType.sqlDate)

      Case Else 'Calculation

        ' RH 29/05/01 - This was commented out. Removed the comments to hopefully fix 2214.

        objCalcExpr = New clsExprExpression
        objCalcExpr.Initialise(lngTableID, lngColumnID, modExpression.ExpressionTypes.giEXPR_RUNTIMECALCULATION, modExpression.ExpressionValueTypes.giEXPRVALUE_UNDEFINED)
        objCalcExpr.ConstructExpression()
        objCalcExpr.ValidateExpression(True)

        DateColumn = (objCalcExpr.ReturnType = modExpression.ExpressionValueTypes.giEXPRVALUE_DATE)
        'UPGRADE_NOTE: Object objCalcExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objCalcExpr = Nothing

    End Select

  End Function
	
  Public Function GetColumnDataType(ByVal lColumnID As Integer) As Integer

    Dim sSQL As String
    Dim rsData As ADODB.Recordset

    sSQL = "Select datatype From ASRSysColumns Where columnid = " & lColumnID
    rsData = datData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

    If Not rsData.BOF And Not rsData.EOF Then
      GetColumnDataType = rsData.Fields(0).Value
    End If

    rsData.Close()
    'UPGRADE_NOTE: Object rsData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsData = Nothing

  End Function
	
  Public Function BitColumn(ByVal strType As String, ByVal lngTableID As Integer, ByVal lngColumnID As Integer) As Boolean

    'RH20000713
    Dim objCalcExpr As New clsExprExpression

    BitColumn = False

    Select Case strType
      Case "C" 'Column
        BitColumn = (GetDataType(lngTableID, lngColumnID) = Declarations.SQLDataType.sqlBoolean)

      Case Else 'Calculation
        objCalcExpr = New clsExprExpression
        objCalcExpr.Initialise(lngTableID, lngColumnID, modExpression.ExpressionTypes.giEXPR_RUNTIMECALCULATION, modExpression.ExpressionValueTypes.giEXPRVALUE_UNDEFINED)
        objCalcExpr.ConstructExpression()
        objCalcExpr.ValidateExpression(True)

        BitColumn = (objCalcExpr.ReturnType = modExpression.ExpressionValueTypes.giEXPRVALUE_LOGIC)
        'UPGRADE_NOTE: Object objCalcExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objCalcExpr = Nothing

    End Select

  End Function
	
  Public Function GetColumnTableName(ByVal plngColumnID As Integer) As String

    ' Return the table id of the given column.
    Dim sSQL As String
    Dim rsData As ADODB.Recordset

    sSQL = "SELECT tableName" & " FROM ASRSysColumns " & " JOIN ASRSysTables ON ASRSysColumns.TableID = ASRSysTables.TableID" & " WHERE columnID = " & Trim(Str(plngColumnID))
    rsData = datData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

    If Not rsData.BOF And Not rsData.EOF Then
      GetColumnTableName = rsData.Fields("TableName").Value
    Else
      GetColumnTableName = ""
    End If

    rsData.Close()
    'UPGRADE_NOTE: Object rsData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsData = Nothing

  End Function
	
  Public Function IsPhotoDataType(ByVal lColumnID As Integer) As Boolean

    Dim sSQL As String
    Dim rsData As ADODB.Recordset

    sSQL = "Select datatype From ASRSysColumns Where columnid= " & lColumnID
    rsData = datData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

    If Not rsData.BOF And Not rsData.EOF Then
      IsPhotoDataType = IIf(rsData.Fields(0).Value = -4, False, True)
    End If

    rsData.Close()
    'UPGRADE_NOTE: Object rsData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsData = Nothing

  End Function
End Class