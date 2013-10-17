Option Strict Off
Option Explicit On

Imports System.Globalization
Imports ADODB
Imports HR.Intranet.Server.Enums
Imports System.Collections.ObjectModel
Imports HR.Intranet.Server.Metadata

Friend Class clsGeneral

	Private datData As New clsDataAccess

	Const FUNCTIONPREFIX As String = "udf_ASRSys_"

  Public Function ConvertNumberForSQL(ByVal strInput As String) As String
    'Get a number in the correct format for a SQL string
    '(e.g. on french systems replace decimal comma for a decimal point)
    ConvertNumberForSQL = Replace(strInput, CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator, ".")
  End Function

  Public Function ConvertNumberForDisplay(ByVal strInput As String) As String
    'Get a number in the correct format for display
    '(e.g. on french systems replace decimal point for a decimal comma)
    ConvertNumberForDisplay = Replace(strInput, ".", CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)
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

    ' eg. DateFormat = "MM/dd/yyyy"
    '     Calendar   = "dd/mm/yyyy"
    '     DateString = "06/02/2000"
    '     Compare to = 02/06/2000

    strDateFormat = CultureInfo.CurrentCulture.DateTimeFormat.ToString
    strDateSeparator = CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator

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
		Dim rsUser As Recordset

    sSQL = "exec master..xp_msver"

		rsUser = datData.OpenRecordset(sSQL, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly)
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
		Dim objExpr As clsExprExpression = New clsExprExpression

		fOK = True

		With objExpr
			.Initialise(0, plngFilterID, ExpressionTypes.giEXPR_RUNTIMEFILTER, ExpressionValueTypes.giEXPRVALUE_LOGIC)
			.UDFFilterCode(pastrUDFs, True, True)
		End With
		'UPGRADE_NOTE: Object objExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objExpr = Nothing

		FilterUDFs = fOK

TidyUpAndExit:
		Exit Function
ErrorTrap:

	End Function

	Public Function GetRecordsInTransaction(ByRef sSQL As String) As Recordset
		' Return the required STATIC/read-only recordset.
		' This is useful when getting a recordset in the middle of a transaction.
		' An error occurs when getting more than one forward-only, read-only recordset in the middle
		' of a transaction.
		GetRecordsInTransaction = datData.OpenRecordset(sSQL, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly)

	End Function

  Public Function FilteredIDs(ByRef plngExprID As Integer, ByRef psIDSQL As String, Optional ByRef paPrompts As Object = Nothing) As Boolean
    ' Return a string describing the record IDs from the given table
    ' that satisfy the given criteria.
    Dim fOK As Boolean
		Dim objExpr As clsExprExpression = New clsExprExpression

		With objExpr
			' Initialise the filter expression object.
			fOK = .Initialise(0, plngExprID, ExpressionTypes.giEXPR_RUNTIMEFILTER, ExpressionValueTypes.giEXPRVALUE_LOGIC)

			If fOK Then
				fOK = objExpr.RuntimeFilterCode(psIDSQL, True, False, paPrompts)
			End If

		End With
    'UPGRADE_NOTE: Object objExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    objExpr = Nothing

    FilteredIDs = fOK

  End Function

	Public Function GetOrderDefinition(ByRef plngOrderID As Integer) As Recordset
		' Return a recordset of the order items (both Find Window and Sort Order columns)
		' for the given order.
		Dim sSQL As String
		Dim rsInfo As Recordset

		sSQL = "EXEC sp_ASRGetOrderDefinition " & plngOrderID
		rsInfo = datData.OpenRecordset(sSQL, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly)
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

	Public Function GetAllViews() As Recordset
		Dim sSQL As String
		Dim rsViews As Recordset

		sSQL = "SELECT ASRSysViews.viewID, ASRSysViews.viewName, ASRSysTables.tableID, ASRSysTables.tableName, ASRSysTables.tableType, ASRSysTables.defaultOrderID, ASRSysTables.recordDescExprID FROM ASRSysViews INNER JOIN ASRSysTables ON ASRSysViews.viewTableID = ASRSysTables.tableID"
		rsViews = datData.OpenRecordset(sSQL, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly)
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
		Dim rsTemp As Recordset
    Dim strSQL As String
    Dim fOK As Boolean
		Dim lngViews(,) As Integer

		On Error GoTo LocalErr

    'UPGRADE_WARNING: Couldn't resolve default property of object GetValueForRecordIndependantCalc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    GetValueForRecordIndependantCalc = vbNullString

    objExpr = New clsExprExpression
    With objExpr
      ' Initialise the filter expression object.
			fOK = .Initialise(0, lngExprID, ExpressionTypes.giEXPR_RECORDINDEPENDANTCALC, ExpressionValueTypes.giEXPRVALUE_UNDEFINED)

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

		Dim cmdLoginInfo As New Command
		Dim prmActutalLogin As Parameter

    cmdLoginInfo.CommandText = "spASRIntGetActualLogin"
    cmdLoginInfo.CommandType = 4
    cmdLoginInfo.let_ActiveConnection(gADOCon)

		prmActutalLogin = cmdLoginInfo.CreateParameter("ActualLogin", DataTypeEnum.adVarChar, ParameterDirectionEnum.adParamOutput, 250)
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

		Dim cmdUserInfo As New Command
		Dim prmActualUser As Parameter
		Dim prmActualUserGroup As Parameter
		Dim prmActualUserGroupID As Parameter

    cmdUserInfo.CommandText = "spASRIntGetActualUserDetails"
    cmdUserInfo.CommandType = 4
    cmdUserInfo.let_ActiveConnection(gADOCon)

		prmActualUser = cmdUserInfo.CreateParameter("psUsername", DataTypeEnum.adVarChar, ParameterDirectionEnum.adParamOutput, 250)
    cmdUserInfo.Parameters.Append(prmActualUser)

		prmActualUserGroup = cmdUserInfo.CreateParameter("psUserGroup", DataTypeEnum.adVarChar, ParameterDirectionEnum.adParamOutput, 250)
    cmdUserInfo.Parameters.Append(prmActualUserGroup)

		prmActualUserGroupID = cmdUserInfo.CreateParameter("piUserGroupID", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamOutput)
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
		Dim rsUser As Recordset

    sSQL = "exec sp_helpuser '" & Replace(gsActualLogin, "'", "''") & "'"

		rsUser = datData.OpenRecordset(sSQL, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly)

    Do While rsUser.Fields("GroupName").Value = "db_owner" Or LCase(Left(rsUser.Fields("GroupName").Value, 6)) = "asrsys"
      rsUser.MoveNext()
    Loop
    GetUserDetails = rsUser.Fields("GroupName").Value

    rsUser.Close()
    'UPGRADE_NOTE: Object rsUser may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsUser = Nothing

  End Function

	Public Function GetRecords(ByRef sSQL As String) As Recordset
		' Return the required forward-only/read-only recordset.
		'  Set GetRecords = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

		' JPD - Changed cursor-type to static.
		' The SQL Server ODBC driver uses a special cursor when the cursor is
		' forward-only, read-only, and the ODBC rowset size is one.
		' The cursor is called a "firehose" cursor because it is the fastest way to retrieve the data.
		' Unfortunately, a side affect of the cursor is that it only permits one active recordset per connection.
		' To get around this we'll try using a STATIC cursor.
		GetRecords = datData.OpenRecordset(sSQL, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly)

	End Function

	Public Function GetReadOnlyRecords(ByRef sSQL As String) As Recordset
		' Return the required dynamic/read-only recordset.
		'  Set GetReadOnlyRecords = datData.OpenRecordset(sSQL, adOpenDynamic, adLockReadOnly)

		' JPD 7/6/00 Changed the cursor type from dynamic to static.
		' A dynamic/read-only recordset opened with a sql select statement that includes the 'distinct' parameter,
		' and returns a single record will automatically be set to be a forwardonly/read-only (firehose) recordet.
		GetReadOnlyRecords = datData.OpenRecordset(sSQL, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly)

	End Function

  Public Function GetTableName(ByVal plngTableID As Integer) As String
		Return Tables.GetById(plngTableID).Name
	End Function

  Public Function GetFilterName(ByVal lFilterID As Integer) As String
		Dim rsFilter As Recordset
    Dim sSQL As String

		sSQL = "SELECT name FROM ASRSysExpressions WHERE ExprID=" & lFilterID

		rsFilter = datData.OpenRecordset(sSQL, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly)
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
		Dim rsPicklist As Recordset
    Dim sSQL As String

		sSQL = "SELECT Name FROM ASRSysPicklistName WHERE PicklistID=" & lPicklistID

		rsPicklist = datData.OpenRecordset(sSQL, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly)
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
		Dim rsTable As Recordset
    Dim sSQL As String

		sSQL = "SELECT recordDescExprID FROM ASRSysTables WHERE TableID=" & TableID

		rsTable = datData.OpenRecordset(sSQL, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockOptimistic)
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

	Public Function GetDataType(ByRef lTableID As Integer, ByRef lngColumnID As Integer) As SQLDataType
		Return Columns.GetById(lngColumnID).DataType
	End Function

	Public Function GetColumnTable(ByRef plngColumnID As Integer) As Integer
		Return Columns.GetById(plngColumnID).TableID
	End Function

	Public Function GetDefaultOrder(ByRef plngTableID As Integer) As Integer
		Return Tables.GetById(plngTableID).DefaultOrderID
	End Function

	Public Function GetOrder(ByVal lOrderID As Integer) As Recordset

		Dim sSQL As String
		Dim rsOrder As Recordset

		sSQL = "SELECT * FROM ASRSysOrders WHERE OrderID=" & lOrderID
		rsOrder = datData.OpenRecordset(sSQL, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly)

		GetOrder = rsOrder

	End Function

	Public Function GetColumnName(ByVal plngColumnID As Integer) As String
		If plngColumnID = 0 Then
			Return ""
		Else
			Return Columns.GetById(plngColumnID).Name
		End If
	End Function

	Public Function GetModuleParameter(ByRef psModuleKey As String, ByRef psParameterKey As String) As String
		Return ModuleSettings.GetSetting(psModuleKey, psParameterKey).ParameterValue
	End Function
	
	Public Function GetUserSetting(ByRef strSection As String, ByRef strKey As String, ByRef varDefault As Object) As Object
		Dim objData = UserSettings.GetUserSetting(strSection, strKey)

		If objData Is Nothing Then
			Return varDefault
		Else
			Return objData
		End If

	End Function

  Public Function UniqueSQLObjectName(ByRef strPrefix As String, ByRef intType As Short) As String

    'TM20020530 Fault 3756 - function altered as the sp needs to insert a record into a table
    'before returning a value, so collect the returned parameter rather than a recordset.

		Dim cmdUniqObj As New Command
		Dim pmADO As Parameter

    With cmdUniqObj
      .CommandText = "sp_ASRUniqueObjectName"
			.CommandType = CommandTypeEnum.adCmdStoredProc
      .CommandTimeout = 0
      .ActiveConnection = gADOCon

			pmADO = .CreateParameter("UniqueObjectName", DataTypeEnum.adVarChar, ParameterDirectionEnum.adParamOutput, 255)
      .Parameters.Append(pmADO)

			pmADO = .CreateParameter("Prefix", DataTypeEnum.adVarChar, ParameterDirectionEnum.adParamInput, 255)
      .Parameters.Append(pmADO)
      pmADO.Value = strPrefix

			pmADO = .CreateParameter("Type", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput)
      .Parameters.Append(pmADO)
      pmADO.Value = intType

      'UPGRADE_NOTE: Object pmADO may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
      pmADO = Nothing

      .Execute()

      'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
      UniqueSQLObjectName = IIf(IsDBNull(.Parameters(0).Value), vbNullString, .Parameters(0).Value)

    End With

    'UPGRADE_NOTE: Object cmdUniqObj may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    cmdUniqObj = Nothing

  End Function

  Public Function DropUniqueSQLObject(ByVal sSQLObjectName As String, ByRef iType As Short) As Boolean

    On Error GoTo ErrorTrap

		Dim cmdUniqObj As New Command
		Dim pmADO As Parameter

    If Len(sSQLObjectName) > 0 Then
      With cmdUniqObj
        .CommandText = "sp_ASRDropUniqueObject"
				.CommandType = CommandTypeEnum.adCmdStoredProc
        .CommandTimeout = 0
        .ActiveConnection = gADOCon

				pmADO = .CreateParameter("UniqueObjectName", DataTypeEnum.adVarChar, ParameterDirectionEnum.adParamInput, 255)
        .Parameters.Append(pmADO)
        pmADO.Value = sSQLObjectName

				pmADO = .CreateParameter("Type", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput)
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
		Return Columns.GetById(plngColumnID).Use1000Separator
	End Function

  ' Returns the amount of decimals that are specificed for a column
  Public Function GetDecimalsSize(ByVal plngColumnID As Integer) As Short
		Return Columns.GetById(plngColumnID).Decimals
	End Function

  Public Function IsAChildOf(ByVal lTestTableID As Integer, ByVal lBaseTableID As Integer) As Boolean
		Return Relations.IsRelation(lBaseTableID, lTestTableID)
	End Function

  Public Function IsAParentOf(ByVal lTestTableID As Integer, ByVal lBaseTableID As Integer) As Boolean
		Return Relations.IsRelation(lTestTableID, lBaseTableID)
	End Function

  Public Function DateColumn(ByVal strType As String, ByVal lngTableID As Integer, ByVal lngColumnID As Integer) As Boolean

    Select Case strType
      Case "C" 'Column
				DateColumn = (Columns.GetById(lngColumnID).DataType = SQLDataType.sqlDate)

      Case Else 'Calculation

				Dim objCalcExpr As clsExprExpression
        objCalcExpr = New clsExprExpression
				objCalcExpr.Initialise(lngTableID, lngColumnID, ExpressionTypes.giEXPR_RUNTIMECALCULATION, ExpressionValueTypes.giEXPRVALUE_UNDEFINED)
        objCalcExpr.ConstructExpression()
        objCalcExpr.ValidateExpression(True)

				DateColumn = (objCalcExpr.ReturnType = ExpressionValueTypes.giEXPRVALUE_DATE)
        'UPGRADE_NOTE: Object objCalcExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objCalcExpr = Nothing

    End Select

  End Function

	Public Function GetColumnDataType(ByVal plngColumnID As Integer) As SQLDataType
		Return Columns.GetById(plngColumnID).DataType
	End Function

  Public Function BitColumn(ByVal strType As String, ByVal lngTableID As Integer, ByVal lngColumnID As Integer) As Boolean

    'RH20000713
		Dim objCalcExpr As clsExprExpression

		Select Case strType
			Case "C" 'Column
				BitColumn = (Columns.GetById(lngColumnID).DataType = SQLDataType.sqlBoolean)

			Case Else	'Calculation
				objCalcExpr = New clsExprExpression
				objCalcExpr.Initialise(lngTableID, lngColumnID, ExpressionTypes.giEXPR_RUNTIMECALCULATION, ExpressionValueTypes.giEXPRVALUE_UNDEFINED)
				objCalcExpr.ConstructExpression()
				objCalcExpr.ValidateExpression(True)

				BitColumn = (objCalcExpr.ReturnType = ExpressionValueTypes.giEXPRVALUE_LOGIC)

		End Select

  End Function

  Public Function GetColumnTableName(ByVal plngColumnID As Integer) As String
		Return Columns.GetById(plngColumnID).TableName
	End Function

	Public Function IsPhotoDataType(ByVal lngColumnID As Integer) As Boolean
		Return Columns.GetById(lngColumnID).DataType = SQLDataType.sqlBoolean
	End Function

	Friend Function UDFFunctions(ByRef paFunctions As String(), ByRef pbCreate As Boolean) As Boolean

		Dim iCount As Integer
		Dim strDropCode As String
		Dim strFunctionName As String
		Dim sCode As String
		Dim iStart As Short
		Dim iEnd As Short
		Dim strFunctionNumber As String

		Try

			If Not paFunctions Is Nothing Then
				For iCount = 0 To paFunctions.Length - 1

					If Not paFunctions(iCount) Is Nothing Then
						iStart = InStr(paFunctions(iCount), FUNCTIONPREFIX) + Len(FUNCTIONPREFIX)
						iEnd = InStr(1, Mid(paFunctions(iCount), 1, 1000), "(@Per")
						strFunctionNumber = Mid(paFunctions(iCount), iStart, iEnd - iStart)
						strFunctionName = FUNCTIONPREFIX & strFunctionNumber

						'Drop existing function (could exist if the expression is used more than once in a report)
						strDropCode = "IF EXISTS" & " (SELECT *" & "   FROM sysobjects" & "   WHERE id = object_id('[" & Replace(gsUsername, "'", "''") & "]." & strFunctionName & "')" & "     AND sysstat & 0xf = 0)" & " DROP FUNCTION [" & gsUsername & "]." & strFunctionName
						datData.ExecuteSql(strDropCode)

						' Create the new function
						If pbCreate Then
							sCode = paFunctions(iCount)
							datData.ExecuteSql(sCode)
						End If
					End If

				Next iCount
			End If

		Catch ex As Exception
			Return False

		End Try

		Return True

	End Function

Public Sub PopulateMetadata()

	Dim rstData As Recordset
	Dim sSQL As String

	Tables = New Collection(Of Table)
	Columns = New Collection(Of Column)
	Relations = New Collection(Of Relation)
	ModuleSettings = New Collection(Of ModuleSetting)
	UserSettings = New Collection(Of UserSetting)
	Functions = New Collection(Of Metadata.Function)

	Try

		sSQL = "SELECT TableID, TableName, TableType, DefaultOrderID, RecordDescExprID FROM ASRSysTables"
		rstData = datData.OpenRecordset(sSQL, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly)
		Do While Not rstData.EOF
			Dim table As New Table
			table.ID = CInt(rstData.Fields("TableID").Value.ToString)
			table.TableType = rstData.Fields("TableType").Value.ToString
			table.Name = rstData.Fields("TableName").Value.ToString
			table.DefaultOrderID = rstData.Fields("DefaultOrderID").Value.ToString
			table.RecordDescExprID = rstData.Fields("RecordDescExprID").Value.ToString
			Tables.Add(table)
			rstData.MoveNext()
		Loop


		sSQL = "SELECT ColumnID, TableID, ColumnName, DataType, Use1000Separator, Size, Decimals FROM ASRSysColumns"
		rstData = datData.OpenRecordset(sSQL, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly)
		Do While Not rstData.EOF
			Dim column As New Column
			column.ID = rstData.Fields("columnid").Value.ToString
			column.TableID = rstData.Fields("tableid").Value.ToString
			column.TableName = Tables.GetById(column.TableID).Name
			column.Name = rstData.Fields("columnname").Value.ToString
			column.DataType = rstData.Fields("datatype").Value.ToString
			column.Use1000Separator = rstData.Fields("use1000separator").Value.ToString
			column.Size = rstData.Fields("size").Value.ToString
			column.Decimals = rstData.Fields("decimals").Value.ToString
			Columns.Add(column)
			rstData.MoveNext()
		Loop


		sSQL = "SELECT ParentID, ChildID FROM ASRSysRelations"
		rstData = datData.OpenRecordset(sSQL, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly)
		Do While Not rstData.EOF
			Dim relation As New Relation
			relation.ParentID = rstData.Fields("parentid").Value.ToString
			relation.ChildID = rstData.Fields("childid").Value.ToString
			Relations.Add(relation)
			rstData.MoveNext()
		Loop

		sSQL = "SELECT * FROM ASRSysModuleSetup"
		rstData = datData.OpenRecordset(sSQL, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly)
		Do While Not rstData.EOF
			Dim moduleSetting As New ModuleSetting
			moduleSetting.ModuleKey = rstData.Fields("ModuleKey").Value.ToString
			moduleSetting.ParameterKey = rstData.Fields("ParameterKey").Value.ToString
			moduleSetting.ParameterValue = rstData.Fields("ParameterValue").Value.ToString
			moduleSetting.ParameterType = rstData.Fields("ParameterType").Value.ToString
			ModuleSettings.Add(moduleSetting)
			rstData.MoveNext()
		Loop


		sSQL = String.Format("SELECT * FROM ASRSysUserSettings WHERE Username = '{0}'", gsUsername)
		rstData = datData.OpenRecordset(sSQL, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly)
		Do While Not rstData.EOF
			Dim userSetting As New UserSetting
			userSetting.Section = rstData.Fields("Section").Value.ToString
			userSetting.Key = rstData.Fields("SettingKey").Value.ToString
			userSetting.Value = rstData.Fields("SettingValue").Value.ToString
			UserSettings.Add(userSetting)
			rstData.MoveNext()
		Loop


		sSQL = "SELECT functionID, functionName, returnType FROM ASRSysFunctions"
		rstData = datData.OpenRecordset(sSQL, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly)
		Do While Not rstData.EOF
			Dim objFunction = New [Function]
			objFunction.ID = rstData.Fields("functionID").Value.ToString
			objFunction.FunctionName = rstData.Fields("functionName").Value.ToString
			objFunction.ReturnType = rstData.Fields("returnType").Value.ToString
			Functions.Add(objFunction)
			rstData.MoveNext()
		Loop


	Catch ex As Exception
			Throw

	End Try

End Sub


End Class