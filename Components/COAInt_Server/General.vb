Option Strict Off
Option Explicit On

Imports System.Globalization
Imports ADODB
Imports System.Collections.Generic
Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Enums
Imports System.Collections.ObjectModel
Imports HR.Intranet.Server.Metadata
Imports System.Data.SqlClient
Imports HR.Intranet.Server.Structures

Public Class clsGeneral

	Private datData As New clsDataAccess
	Private ReadOnly _login As LoginInfo

	Public Sub New(ByVal LoginInfo As LoginInfo)
		_login = LoginInfo
		datData = New clsDataAccess(LoginInfo)
	End Sub

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

	Public Function GetRecordsInTransaction(ByRef sSQL As String) As Recordset
		' Return the required STATIC/read-only recordset.
		' This is useful when getting a recordset in the middle of a transaction.
		' An error occurs when getting more than one forward-only, read-only recordset in the middle
		' of a transaction.
		GetRecordsInTransaction = datData.OpenRecordset(sSQL, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly)

	End Function

	Public Function FilteredIDs(ByRef plngExprID As Integer, ByRef psIDSQL As String, ByRef psUDFs() As String, Optional ByRef paPrompts As Object = Nothing) As Boolean
		' Return a string describing the record IDs from the given table
		' that satisfy the given criteria.
		Dim fOK As Boolean
		Dim objExpr As clsExprExpression = New clsExprExpression(_login)

		With objExpr
			' Initialise the filter expression object.
			fOK = .Initialise(0, plngExprID, ExpressionTypes.giEXPR_RUNTIMEFILTER, ExpressionValueTypes.giEXPRVALUE_LOGIC)

			If fOK Then
				fOK = objExpr.RuntimeFilterCode(psIDSQL, True, psUDFs, False, paPrompts)
			End If

		End With
		'UPGRADE_NOTE: Object objExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objExpr = Nothing

		Return fOK

	End Function

	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()

		datData = New clsDataAccess

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

		objExpr = New clsExprExpression(_login)
		With objExpr
			' Initialise the filter expression object.
			fOK = .Initialise(0, lngExprID, ExpressionTypes.giEXPR_RECORDINDEPENDANTCALC, ExpressionValueTypes.giEXPRVALUE_UNDEFINED)

			If fOK Then
				fOK = objExpr.RuntimeCalculationCode(lngViews, strSQL, Nothing, True, False, pvarPrompts)
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

	Friend Function GetRecords(ByRef sSQL As String) As Recordset
		' Return the required forward-only/read-only recordset.
		'  Set GetRecords = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

		' JPD - Changed cursor-type to static.
		' The SQL Server ODBC driver uses a special cursor when the cursor is
		' forward-only, read-only, and the ODBC rowset size is one.
		' The cursor is called a "firehose" cursor because it is the fastest way to retrieve the data.
		' Unfortunately, a side affect of the cursor is that it only permits one active recordset per connection.
		' To get around this we'll try using a STATIC cursor.
		GetRecords = datData.OpenRecordset(sSQL, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly)

	End Function

	Friend Function GetReadOnlyRecords(ByRef sSQL As String) As Recordset

		' JDM - 13/12/2003 - Converted to a firehose cursor. This will not reurn a .recordcount so use with caution. 
		'   If a record count is needed then use a function which uses a static cursor type - be warned of permformance though)
		GetReadOnlyRecords = datData.OpenRecordset(sSQL, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly)

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

	Public Function GetColumnName(ByVal plngColumnID As Integer) As String
		If plngColumnID = 0 Then
			Return ""
		Else
			Return Columns.GetById(plngColumnID).Name
		End If
	End Function

	Friend Function GetModuleParameter(ByRef psModuleKey As String, ByRef psParameterKey As String) As String
		Return ModuleSettings.GetSetting(psModuleKey, psParameterKey).ParameterValue
	End Function

	Public Function GetUserSetting(ByRef strSection As String, ByRef strKey As String, ByRef varDefault As Object) As Object
		Dim objData = UserSettings.GetUserSetting(strSection, strKey)

		If objData Is Nothing Then
			Return varDefault
		Else
			Return objData.Value
		End If

	End Function

	Friend Function UniqueSQLObjectName(ByRef strPrefix As String, ByRef intType As Integer) As String

		Try

			Dim prmName As New SqlParameter("psUniqueObjectName", SqlDbType.NVarChar, 128)
			prmName.Direction = ParameterDirection.Output

			Dim prmPrefix As New SqlParameter("Prefix", SqlDbType.NVarChar, 128)
			prmPrefix.Value = strPrefix

			Dim prmType As New SqlParameter("Type", SqlDbType.Int)
			prmType.Value = intType

			datData.ExecuteSP("sp_ASRUniqueObjectName", prmName, prmPrefix, prmType)

			Return prmName.Value

		Catch ex As Exception
			Return ""

		End Try

	End Function

	Public Function DropUniqueSQLObject(ByVal sSQLObjectName As String, ByRef iType As Short) As Boolean

		Try

			Dim prmName As New SqlParameter("psUniqueObjectName", SqlDbType.NVarChar, 128)
			prmName.Value = sSQLObjectName

			Dim prmType As New SqlParameter("piType", SqlDbType.Int)
			prmType.Value = iType

			datData.ExecuteSP("sp_ASRDropUniqueObject", prmName, prmType)

		Catch ex As Exception
			Throw

		End Try

		Return True

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

			Case Else	'Calculation

				Dim objCalcExpr As clsExprExpression
				objCalcExpr = New clsExprExpression(_login)
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
				objCalcExpr = New clsExprExpression(_login)
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
		Return Columns.GetById(lngColumnID).DataType = SQLDataType.sqlVarBinary
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

End Class