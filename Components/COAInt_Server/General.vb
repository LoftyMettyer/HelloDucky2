Option Strict On
Option Explicit On

Imports System.Globalization
Imports HR.Intranet.Server.Enums
Imports System.Data.SqlClient
Imports HR.Intranet.Server.Structures

Public Class clsGeneral

	Private DB As New clsDataAccess
	Private ReadOnly _login As LoginInfo

	Public Sub New(ByVal LoginInfo As LoginInfo)
		_login = LoginInfo
		DB = New clsDataAccess(LoginInfo)
	End Sub

	Const FUNCTIONPREFIX As String = "udf_ASRSys_"


	Public Function ConvertSQLDateToSystemFormat(pstrDateString As String) As Date

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

		Dim i As Integer

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

		Return dtTemp

	End Function

	Public Sub New()
		MyBase.New()
		DB = New clsDataAccess
	End Sub


	Friend Function GetReadOnlyRecords(sSQL As String) As DataTable
		Return DB.GetDataTable(sSQL)
	End Function

	'Public Function GetTableName(plngTableID As Integer) As String
	'	Return Tables.GetById(plngTableID).Name
	'End Function

	Public Function GetFilterName(lFilterID As Integer) As String

		Dim sSQL As String = String.Format("SELECT name FROM ASRSysExpressions WHERE ExprID={0}", lFilterID)
		With DB.GetDataTable(sSQL)
			If .Rows.Count > 0 Then
				Return Trim(.Rows(0)(0).ToString())
			Else
				Return vbNullString
			End If
		End With

	End Function

	Public Function GetPicklistName(lPicklistID As Integer) As String

		Dim sSQL As String = String.Format("SELECT Name FROM ASRSysPicklistName WHERE PicklistID={0}", lPicklistID)
		With DB.GetDataTable(sSQL)
			If .Rows.Count > 0 Then
				Return Trim(.Rows(0)(0).ToString())
			Else
				Return vbNullString
			End If
		End With

	End Function

	Public Function GetRecDescExprID(TableID As Integer) As Integer

		Dim sSQL As String = String.Format("SELECT recordDescExprID FROM ASRSysTables WHERE TableID={0}", TableID)
		With DB.GetDataTable(sSQL)
			If .Rows.Count > 0 Then
				Return CInt(.Rows(0)(0))
			Else
				Return 0
			End If
		End With

	End Function


	Friend Function UniqueSQLObjectName(strPrefix As String, intType As Integer) As String

		Try

			Dim prmName As New SqlParameter("psUniqueObjectName", SqlDbType.NVarChar, 128)
			prmName.Direction = ParameterDirection.Output

			Dim prmPrefix As New SqlParameter("Prefix", SqlDbType.NVarChar, 128)
			prmPrefix.Value = strPrefix

			Dim prmType As New SqlParameter("Type", SqlDbType.Int)
			prmType.Value = intType

			DB.ExecuteSP("sp_ASRUniqueObjectName", prmName, prmPrefix, prmType)

			Return prmName.Value.ToString()

		Catch ex As Exception
			Return ""

		End Try

	End Function

	Public Function DropUniqueSQLObject(sSQLObjectName As String, iType As Short) As Boolean

		Try

			Dim prmName As New SqlParameter("psUniqueObjectName", SqlDbType.NVarChar, 128)
			prmName.Value = sSQLObjectName

			Dim prmType As New SqlParameter("piType", SqlDbType.Int)
			prmType.Value = iType

			DB.ExecuteSP("sp_ASRDropUniqueObject", prmName, prmType)

		Catch ex As Exception
			Throw

		End Try

		Return True

	End Function


	Friend Function UDFFunctions(ByRef paFunctions As String(), ByRef pbCreate As Boolean) As Boolean

		Dim iCount As Integer
		Dim strDropCode As String
		Dim strFunctionName As String
		Dim sCode As String
		Dim iStart As Integer
		Dim iEnd As Integer
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
						DB.ExecuteSql(strDropCode)

						' Create the new function
						If pbCreate Then
							sCode = paFunctions(iCount)
							DB.ExecuteSql(sCode)
						End If
					End If

				Next iCount
			End If

		Catch ex As Exception
			Return False

		End Try

		Return True

	End Function

#Region "From modUtilityAccess"

	Friend Function CurrentUserAccess(piUtilityType As UtilityType, plngID As Integer) As String

		' Return the access code (RW/RO/HD) of the current user's access
		' on the given utility.
		Dim sAccessCode As String
		Dim sSQL As String
		Dim sDefaultAccess As String
		Dim rsAccess As DataTable
		Dim sTableName As String
		Dim sAccessTableName As String
		Dim sIDColumnName As String

		sTableName = ""
		sAccessTableName = ""

		If plngID > 0 Then
			sDefaultAccess = ACCESS_HIDDEN
		Else
			sDefaultAccess = ACCESS_HIDDEN
		End If

		' Construct the SQL code to get the current user's access settings for the given utility.
		' NB. System and Security Manager users automatically have Read/Write access.
		Select Case piUtilityType
			Case UtilityType.utlBatchJob
				sTableName = "ASRSysBatchJobName"
				sAccessTableName = "ASRSysBatchJobAccess"
				sIDColumnName = "ID"

			Case UtilityType.utlCalendarReport
				sTableName = "ASRSysCalendarReports"
				sAccessTableName = "ASRSysCalendarReportAccess"
				sIDColumnName = "ID"

			Case UtilityType.utlCrossTab
				sTableName = "ASRSysCrossTab"
				sAccessTableName = "ASRSysCrossTabAccess"
				sIDColumnName = "CrossTabID"

			Case UtilityType.utlCustomReport
				sTableName = "ASRSysCustomReportsName"
				sAccessTableName = "ASRSysCustomReportAccess"
				sIDColumnName = "ID"

			Case UtilityType.utlDataTransfer
				sTableName = "ASRSysDataTransferName"
				sAccessTableName = "ASRSysDataTransferAccess"
				sIDColumnName = "DataTransferID"

			Case UtilityType.utlExport
				sTableName = "ASRSysExportName"
				sAccessTableName = "ASRSysExportAccess"
				sIDColumnName = "ID"

			Case UtilityType.UtlGlobalAdd, UtilityType.utlGlobalDelete, UtilityType.utlGlobalUpdate
				sTableName = "ASRSysGlobalFunctions"
				sAccessTableName = "ASRSysGlobalAccess"
				sIDColumnName = "functionID"

			Case UtilityType.utlImport
				sTableName = "ASRSysImportName"
				sAccessTableName = "ASRSysImportAccess"
				sIDColumnName = "ID"

			Case UtilityType.utlLabel, UtilityType.utlMailMerge
				sTableName = "ASRSysMailMergeName"
				sAccessTableName = "ASRSysMailMergeAccess"
				sIDColumnName = "mailMergeID"

			Case UtilityType.utlRecordProfile
				sTableName = "ASRSysRecordProfileName"
				sAccessTableName = "ASRSysRecordProfileAccess"
				sIDColumnName = "recordProfileID"

			Case UtilityType.utlMatchReport, UtilityType.utlSuccession, UtilityType.utlCareer
				sTableName = "ASRSysMatchReportName"
				sAccessTableName = "ASRSysMatchReportAccess"
				sIDColumnName = "matchReportID"

		End Select

		Try

			If Len(sAccessTableName) > 0 Then
				sSQL = "SELECT" & "  CASE" & "    WHEN (SELECT count(*)" & "      FROM ASRSysGroupPermissions" & "      INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & "        AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & "        OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & "      INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & "        AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & "      WHERE b.Name = ASRSysGroupPermissions.groupname" & "        AND ASRSysGroupPermissions.permitted = 1) > 0 THEN '" & ACCESS_READWRITE & "'" & "    WHEN " & sTableName & ".userName = system_user THEN '" & ACCESS_READWRITE & "'" & "    ELSE" & "      CASE" & "        WHEN " & sAccessTableName & ".access IS null THEN '" & sDefaultAccess & "'" & "        ELSE " & sAccessTableName & ".access" & "      END" & "  END AS Access" & " FROM sysusers b" & " INNER JOIN sysusers a ON b.uid = a.gid" & " LEFT OUTER JOIN " & sAccessTableName & " ON (b.name = " & sAccessTableName & ".groupName" & "   AND " & sAccessTableName & ".id = " & CStr(plngID) & ")" & " INNER JOIN " & sTableName & " ON " & sAccessTableName & ".ID = " & sTableName & "." & sIDColumnName & " WHERE b.name = '" & gsUserGroup & "'"


				rsAccess = DB.GetDataTable(sSQL)
				With rsAccess
					If .Rows.Count = 0 Then
						sAccessCode = sDefaultAccess
					Else
						sAccessCode = .Rows(0)("Access").ToString()
					End If

				End With
			Else
				sAccessCode = ACCESS_UNKNOWN
			End If

			Return sAccessCode

		Catch ex As Exception
			Return sDefaultAccess

		End Try

	End Function


#End Region


#Region "From modExpression"

	Public Function HasHiddenComponents(ByVal lngExprID As Integer) As Boolean

		'********************************************************************************
		' HasHiddenComponents - Loops through the passed expression searching for       *
		'                       hidden expressions (calcs/filters).                     *
		'                       Note: This function calls itself and drills down the    *
		'                       expression checking for hidden calcs & filters, then    *
		'                       works its way up the expressions/components.            *
		'                                                                               *
		' 'TM20010802 Fault 2617                                                        *
		'********************************************************************************

		Dim rsExpr As DataTable
		Dim rsExprComp As DataTable
		Dim lngCalcFilterID As Integer
		Dim bHasHiddenComp As Boolean
		Dim sSQL As String

		On Error GoTo ErrorTrap

		sSQL = String.Format("SELECT *, ISNULL(e.Access,'') AS [Access] FROM ASRSysExprComponents c INNER JOIN ASRSysExpressions e ON c.ExprID = e.ExprID WHERE c.ExprID = {0}", lngExprID)
		rsExprComp = DB.GetDataTable(sSQL)
		bHasHiddenComp = False

		With rsExprComp

			For Each objRow As DataRow In .Rows

				Select Case CType(objRow("Type"), ExpressionComponentTypes)
					Case ExpressionComponentTypes.giCOMPONENT_CALCULATION
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						lngCalcFilterID = CInt(IIf(IsDBNull(objRow("CalculationID")), 0, objRow("CalculationID")))

						If lngCalcFilterID > 0 Then
							If HasHiddenComponents(lngCalcFilterID) Or objRow("Access").ToString() = ACCESS_HIDDEN Then
								bHasHiddenComp = True
							End If
						End If

					Case ExpressionComponentTypes.giCOMPONENT_FILTER
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						lngCalcFilterID = CInt(IIf(IsDBNull(objRow("FilterID")), 0, objRow("FilterID")))


						If lngCalcFilterID > 0 Then
							If HasHiddenComponents(lngCalcFilterID) Or objRow("Access").ToString() = ACCESS_HIDDEN Then
								bHasHiddenComp = True
								'TM20011003
								'Need this function to just find out if there are any hidden components,
								'it was also setting the access of the functions and therefore changing
								'time stamp.
								'SetExprAccess lngCalcFilterID, "HD"
							End If
						End If

					Case ExpressionComponentTypes.giCOMPONENT_FIELD
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						lngCalcFilterID = CInt(IIf(IsDBNull(objRow("FieldSelectionFilter")), 0, objRow("FieldSelectionFilter")))

						If lngCalcFilterID > 0 Then
							If HasHiddenComponents(lngCalcFilterID) Or objRow("Access").ToString() = ACCESS_HIDDEN Then
								bHasHiddenComp = True
							End If
						End If

					Case ExpressionComponentTypes.giCOMPONENT_FUNCTION
						sSQL = "SELECT exprID, Access FROM ASRSysExpressions WHERE parentComponentID = " & CStr(objRow("ComponentID"))
						rsExpr = DB.GetDataTable(sSQL)
						For Each objFunctionRow As DataRow In rsExpr.Rows

							If HasHiddenComponents(CInt(objFunctionRow("ExprID"))) Or objFunctionRow("Access").ToString() = ACCESS_HIDDEN Then
								bHasHiddenComp = True
								Exit For
							End If

						Next
						'UPGRADE_NOTE: Object rsExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
						rsExpr = Nothing
				End Select

				If bHasHiddenComp Then
					Exit For
				End If

			Next
		End With

		Return bHasHiddenComp

ErrorTrap:
		Return False

	End Function



#End Region


End Class