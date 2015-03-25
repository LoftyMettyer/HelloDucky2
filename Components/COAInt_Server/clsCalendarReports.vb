Option Strict Off
Option Explicit On

Imports System.Globalization
Imports System.Collections.Generic
Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Enums
Imports System.Text
Imports HR.Intranet.Server.Metadata
Imports HR.Intranet.Server.Structures
Imports HR.Intranet.Server.Expressions

Public Class CalendarReport
	Inherits BaseReport

	Public rsPersonnelBHols As DataTable
	Public rsTempPersonnelBHols As DataTable

	Public Legend As List(Of CalendarLegend)
	Public LegendColors As List(Of LegendColor)

	Public rsCareerChange As DataTable

	Private mstrSQLSelect_RegInfoRegion As String
	Private mstrSQLSelect_BankHolDate As String
	Private mstrSQLSelect_BankHolDesc As String

	Private mstrSQLSelect_PersonnelStaticRegion As String

	Private mvarTableViews(,) As Object

	'TableViews
	Private mstrRealSource As String
	Private mstrBaseTableRealSource As String
	Private mlngTableViews(,) As Integer
	Private mstrViews() As String
	Private mobjTableView As TablePrivilege
	Private mobjColumnPrivileges As CColumnPrivileges

	Private mvarEventColumnViews(,) As Object
	Private mlngEventViewColumn As Integer

	'************************************************************

	Private mlngCalendarReportID As Integer
	Private mstrErrorString As String
	Private mstrTempTableName As String
	Private mblnUDFsCreated As Boolean
	Private mblnTempTableCreated As Boolean
	Private mblnOrderByCreated As Boolean
	Private mlngBaseDescriptionType As Integer

	Private mblnStaticReg As Boolean
	Private mblnStaticWP As Boolean

	'Variables to store definition (report level variables)
	Private mlngCalendarReportsBaseTable As Integer
	Private mstrCalendarReportsBaseTableName As String
	Private mlngCalendarReportsPickListID As Integer
	Private mlngCalendarReportsFilterID As Integer
	Private mlngDescription1 As Integer
	Private mstrDescription1 As String
	Private mblnDesc1IsDate As Boolean
	Private mlngDescription2 As Integer
	Private mstrDescription2 As String
	Private mblnDesc2IsDate As Boolean
	Private mlngDescriptionExpr As Integer
	Private mblnDescExprIsDate As Boolean

	Private mstrDescriptionSeparator As String

	Private mstrBaseIDColumn As String
	Private mstrEventIDColumn As String

	Private mblnDescCalcCode As Boolean
	Private mstrDescCalcCode As String

	Private mlngRegion As Integer
	Private mstrRegion As String
	Private mblnGroupByDescription As Boolean

	Private mlngStartDateExpr As Integer
	Private mdtStartDate As Date
	Private mlngEndDateExpr As Integer
	Private mdtEndDate As Date

	Private mblnShowBankHolidays As Boolean
	Private mblnShowCaptions As Boolean
	Private mblnShowWeekends As Boolean
	Private mbStartOnCurrentMonth As Boolean
	Private mblnIncludeWorkingDaysOnly As Boolean
	Private mblnIncludeBankHolidays As Boolean
	Private mblnCustomReportsPrintFilterHeader As Boolean
	Private mstrFilteredIDs As String

	'New Default Output Variables

	Private mblnOutputPrinter As Boolean
	Private mstrOutputPrinterName As String
	Private mblnOutputSave As Boolean
	Private mlngOutputSaveExisting As Integer
	Private mblnOutputEmail As Boolean
	Private mlngOutputEmailID As Integer
	Private mstrOutputEmailName As String
	Private mstrOutputEmailSubject As String
	Private mstrOutputEmailAttachAs As String

	'Recordset to store the final data from the temp table
	Private mrsCalendarReportsOutput As DataTable
	Private mrsCalendarBaseInfo As DataTable

	Private mstrClientDateFormat As String
	Private mstrLocalDecimalSeparator As String

	'Strings to hold the SQL statement
	Private mstrSQLEvent As String
	Private mstrSQLSelect As String
	Private mstrSQLFrom As String
	Private mstrSQLJoin As String
	Private mstrSQLWhere As String
	Private mstrSQLOrderBy As String
	Private mstrSQL As String
	Private mstrSQLBaseData As String
	Private mstrSQLBaseDateClause As String
	Private mstrSQLOrderList As String
	Private mstrSQLIDs As String
	Private mstrSQLDynamicLegendWhere As String
	Private mintDynamicEventCount As Short
	Private mstrSQLCreateTable As String

	Private mblnHasEventFilterIDs As Boolean
	Private mstrEventFilterIDs As String

	'used to temporarily store the Base table Start & End date table.columnname for the
	'current event. Then used when creating the mstrSQLBaseDateClause.
	Private mstrSQLBaseStartDateColumn As String
	Private mstrSQLBaseStartSessionColumn As String
	Private mstrSQLBaseEndDateColumn As String
	Private mstrSQLBaseEndSessionColumn As String
	Private mstrSQLBaseDurationColumn As String

	'Array holding the columns to sort the report by
	Private mvarSortOrder(,) As Object
	Private mvarPrompts(,) As Object

	Private mcolEvents As clsCalendarEvents

	'Instance of the previewform
	'Private mfrmOutput As frmCalendarReportPreview

	'Does the report generate no records ?
	Private mblnNoRecords As Boolean


	'Runnning report for single record only!
	Private mlngSingleRecordID As Integer

	' Array holding the User Defined functions that are needed for this report
	Private mastrUDFsRequired() As String

	Private mcolStaticBankHolidays As Collection
	Private mcolHistoricBankHolidays As Collection
	Private mcolStaticWorkingPatterns As Collection
	Private mcolHistoricWorkingPatterns As Collection

	Private mblnPersonnelBase As Boolean

	Private mstrRegionFormString As String
	Private mstrBHolFormString As StringBuilder
	Private mstrWPFormString As String

	'****************************************************
	'variables for outputting
	Private mavOutputDateIndex(,) As Object

	Private mdtVisibleStartDate_Output As Date
	Private mdtVisibleEndDate_Output As Date
	Private mstrEventLegend_Output As String

	'****************************************************

	Private mvarOutputArray_Definition() As Object
	Private mvarOutputArray_Columns() As Object
	Private mvarOutputArray_Data() As Object
	Private mvarOutputArray_Styles() As Object
	Private mvarOutputArray_Merges() As Object

	Private mavLegend(,) As Object
	'****************************************************
	'variables for checking for multiple events

	Private mblnHasMultipleEvents As Boolean
	'****************************************************

	Private mblnDisableRegions As Boolean
	Private mblnDisableWPs As Boolean

	Private mstrCurrentEventKey As String

	Private Const CALREP_DATEFORMAT As String = "dd/MM/yyyy"

	Private mavCareerRanges(,) As Object

	Private mintType_BaseDesc1 As Short
	Private mintType_BaseDesc2 As Short
	Private mintType_BaseDescExpr As Short
	Private mstrFormat_BaseDesc1 As String
	Private mstrFormat_BaseDesc2 As String

	Private mblnCheckingDescColumn As Boolean

	Private Function SQLDateConvertToLocale(ByRef pstrTableColumn As String) As String

		'Takes the Column value and Returns a string with the SQL Code to format the
		'SQL date value into the known locale.

		Dim strDateFormat As String

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

		strDateFormat = mstrClientDateFormat

		strDateSeparator = mstrLocalDecimalSeparator

		blnDateComplete = False
		blnMonthDone = False
		blnDayDone = False
		blnYearDone = False

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
			Case "dmy"
				SQLDateConvertToLocale = "CONVERT(varchar, DATEPART(d," & pstrTableColumn & ")) + '" & strDateSeparator & "' + " & "CONVERT(varchar, DATEPART(m," & pstrTableColumn & ")) + '" & strDateSeparator & "' + " & "CONVERT(varchar, DATEPART(yyyy," & pstrTableColumn & ")) "
			Case "mdy"
				SQLDateConvertToLocale = "CONVERT(varchar, DATEPART(m," & pstrTableColumn & ")) + '" & strDateSeparator & "' + " & "CONVERT(varchar, DATEPART(d," & pstrTableColumn & ")) + '" & strDateSeparator & "' + " & "CONVERT(varchar, DATEPART(yyyy," & pstrTableColumn & ")) "
			Case "ydm"
				SQLDateConvertToLocale = "CONVERT(varchar, DATEPART(yyyy," & pstrTableColumn & ")) + '" & strDateSeparator & "' + " & "CONVERT(varchar, DATEPART(d," & pstrTableColumn & ")) + '" & strDateSeparator & "' + " & "CONVERT(varchar, DATEPART(m," & pstrTableColumn & ")) "
			Case "myd"
				SQLDateConvertToLocale = "CONVERT(varchar, DATEPART(m," & pstrTableColumn & ")) + '" & strDateSeparator & "' + " & "CONVERT(varchar, DATEPART(yyyy," & pstrTableColumn & ")) + '" & strDateSeparator & "' + " & "CONVERT(varchar, DATEPART(d," & pstrTableColumn & ")) "
			Case "ymd"
				SQLDateConvertToLocale = "CONVERT(varchar, DATEPART(yyyy," & pstrTableColumn & ")) + '" & strDateSeparator & "' + " & "CONVERT(varchar, DATEPART(m," & pstrTableColumn & ")) + '" & strDateSeparator & "' + " & "CONVERT(varchar, DATEPART(d," & pstrTableColumn & ")) "
		End Select


	End Function

	Public Property EventLogID() As Integer
		Get
			EventLogID = Logs.EventLogID
		End Get
		Set(ByVal Value As Integer)
			Logs.EventLogID = Value
		End Set
	End Property

	Public WriteOnly Property Cancelled() As Boolean
		Set(ByVal Value As Boolean)

			' Connection object passed in from the asp page
			If Value = True Then
				Logs.ChangeHeaderStatus(EventLog_Status.elsCancelled)
			Else
				Logs.ChangeHeaderStatus(EventLog_Status.elsSuccessful)
			End If

		End Set
	End Property

	Public WriteOnly Property Failed() As Boolean
		Set(ByVal Value As Boolean)

			' Connection object passed in from the asp page
			If Value = True Then
				Logs.ChangeHeaderStatus(EventLog_Status.elsFailed)
			End If

		End Set
	End Property

	Public WriteOnly Property FailedMessage() As String
		Set(ByVal Value As String)
			Logs.AddDetailEntry(Value)
		End Set
	End Property

	Public ReadOnly Property NoRecords() As Boolean
		Get
			' Does the report have any records ?
			NoRecords = mblnNoRecords
		End Get
	End Property

	Public ReadOnly Property OutputPrinter() As Boolean
		Get
			OutputPrinter = mblnOutputPrinter
		End Get
	End Property

	Public ReadOnly Property OutputPrinterName() As String
		Get
			OutputPrinterName = mstrOutputPrinterName
		End Get
	End Property

	Public ReadOnly Property OutputSave() As Boolean
		Get
			OutputSave = mblnOutputSave
		End Get
	End Property

	Public ReadOnly Property OutputSaveExisting() As Integer
		Get
			OutputSaveExisting = mlngOutputSaveExisting
		End Get
	End Property

	Public ReadOnly Property OutputEmail() As Boolean
		Get
			OutputEmail = mblnOutputEmail
		End Get
	End Property

	Public ReadOnly Property OutputEmailID() As Integer
		Get
			OutputEmailID = mlngOutputEmailID
		End Get
	End Property

	Public ReadOnly Property OutputEmailGroupName() As String
		Get
			OutputEmailGroupName = mstrOutputEmailName
		End Get
	End Property

	Public ReadOnly Property OutputEmailSubject() As String
		Get
			OutputEmailSubject = mstrOutputEmailSubject
		End Get
	End Property

	Public ReadOnly Property OutputEmailAttachAs() As String
		Get
			OutputEmailAttachAs = mstrOutputEmailAttachAs
		End Get
	End Property

	Public ReadOnly Property HasMultipleEvents() As Boolean
		Get
			HasMultipleEvents = mblnHasMultipleEvents
		End Get
	End Property

	Public ReadOnly Property IncludeBankHolidays_Enabled() As Boolean
		Get
			If (Not mblnGroupByDescription) And (Not mblnDisableRegions) And ((PersonnelBase And (Len(Trim(PersonnelModule.gsPersonnelRegionColumnName)) > 0) And (BankHolidayModule.glngBHolRegionID > 0)) Or (PersonnelBase And (Len(Trim(PersonnelModule.gsPersonnelHRegionColumnName)) > 0) And (BankHolidayModule.glngBHolRegionID > 0)) Or (mlngRegion > 0)) Then

				IncludeBankHolidays_Enabled = True
			Else
				IncludeBankHolidays_Enabled = False
			End If
		End Get
	End Property

	Public ReadOnly Property IncludeWorkingDaysOnly_Enabled() As Boolean
		Get
			If (Not mblnGroupByDescription) And (Not mblnDisableWPs) And ((PersonnelBase And (Len(Trim(PersonnelModule.gsPersonnelWorkingPatternColumnName)) > 0)) Or (PersonnelBase And (Len(Trim(PersonnelModule.gsPersonnelHWorkingPatternColumnName)) > 0))) Then

				IncludeWorkingDaysOnly_Enabled = True
			Else
				IncludeWorkingDaysOnly_Enabled = False
			End If
		End Get
	End Property

	Public WriteOnly Property LocalDecimalSeparator() As String
		Set(ByVal Value As String)

			' Clients date format passed in from the asp page
			mstrLocalDecimalSeparator = Value

		End Set
	End Property

	Public WriteOnly Property ClientDateFormat() As String
		Set(ByVal Value As String)

			' Clients date format passed in from the asp page
			mstrClientDateFormat = Value

		End Set
	End Property

	Public ReadOnly Property ShowBankHolidays_Enabled() As Boolean
		Get
			If (Not mblnGroupByDescription) And (Not mblnDisableRegions) And ((PersonnelBase And (Len(Trim(PersonnelModule.gsPersonnelRegionColumnName)) > 0) And (BankHolidayModule.glngBHolRegionID > 0)) Or (PersonnelBase And (Len(Trim(PersonnelModule.gsPersonnelHRegionColumnName)) > 0) And (BankHolidayModule.glngBHolRegionID > 0)) Or (mlngRegion > 0)) Then
				ShowBankHolidays_Enabled = True
			Else
				ShowBankHolidays_Enabled = False
			End If
		End Get
	End Property

	Public ReadOnly Property SQLCalendarBaseInfo() As String
		Get
			SQLCalendarBaseInfo = mstrSQLBaseData
		End Get
	End Property

	Public ReadOnly Property SQLOutput() As String
		Get
			SQLOutput = mstrSQL
		End Get
	End Property

	Public ReadOnly Property StaticWP() As Boolean
		Get
			StaticWP = mblnStaticWP
		End Get
	End Property

	Public ReadOnly Property StaticReg() As Boolean
		Get
			StaticReg = mblnStaticReg
		End Get
	End Property

	Public Property CalendarReportID() As Integer
		Get
			CalendarReportID = mlngCalendarReportID
		End Get
		Set(ByVal Value As Integer)
			mlngCalendarReportID = Value
		End Set
	End Property

	Public WriteOnly Property SingleRecordID() As Integer
		Set(ByVal Value As Integer)
			mlngSingleRecordID = Value
		End Set
	End Property

	Public Property ErrorString() As String
		Get
			ErrorString = mstrErrorString
		End Get
		Set(ByVal Value As String)
			mstrErrorString = Value
		End Set
	End Property

	Public ReadOnly Property BaseIDColumn() As String
		Get
			BaseIDColumn = "?ID_" & mstrCalendarReportsBaseTableName
		End Get
	End Property

	Public ReadOnly Property EventIDColumn() As String
		Get
			EventIDColumn = "?ID_EventID"
		End Get
	End Property

	Public ReadOnly Property PersonnelBase() As Boolean
		Get
			PersonnelBase = (mlngCalendarReportsBaseTable = PersonnelModule.glngPersonnelTableID)
		End Get
	End Property

	Public ReadOnly Property BaseTableRealSource() As String
		Get
			BaseTableRealSource = mstrBaseTableRealSource
		End Get
	End Property

	Public ReadOnly Property BaseTableID() As String
		Get
			BaseTableID = CStr(mlngCalendarReportsBaseTable)
		End Get
	End Property

	Public ReadOnly Property BaseDesc1IsDate() As Boolean
		Get
			BaseDesc1IsDate = mblnDesc1IsDate
		End Get
	End Property

	Public ReadOnly Property BaseDesc2IsDate() As Boolean
		Get
			BaseDesc2IsDate = mblnDesc2IsDate
		End Get
	End Property

	Public ReadOnly Property BaseDescExprIsDate() As Boolean
		Get
			BaseDescExprIsDate = mblnDescExprIsDate
		End Get
	End Property

	Public ReadOnly Property SQLIDs() As String
		Get
			SQLIDs = mstrSQLIDs
		End Get
	End Property

	Public ReadOnly Property StaticRegionColumn() As String
		Get
			StaticRegionColumn = mstrRegion
		End Get
	End Property

	Public ReadOnly Property StaticRegionColumnID() As Integer
		Get
			StaticRegionColumnID = mlngRegion
		End Get
	End Property

	Public ReadOnly Property ReportStartDate() As Date
		Get
			Return mdtStartDate
		End Get
	End Property

	Public ReadOnly Property ReportEndDate() As Date
		Get
			Return mdtEndDate
		End Get
	End Property

	Public ReadOnly Property CalendarReportTitle() As String
		Get
			If mblnCustomReportsPrintFilterHeader Then
				If (mlngCalendarReportsFilterID > 0) Then
					CalendarReportTitle = Name & " (Base Table filter : " & General.GetFilterName(mlngCalendarReportsFilterID) & ")"
				ElseIf (mlngCalendarReportsPickListID > 0) Then
					CalendarReportTitle = Name & " (Base Table picklist : " & General.GetPicklistName(mlngCalendarReportsPickListID) & ")"
				End If
			Else
				CalendarReportTitle = Name
			End If
		End Get
	End Property

	Public Events As DataTable

	Public ReadOnly Property EventsRecordset() As DataTable
		Get
			Return mrsCalendarReportsOutput
		End Get
	End Property

	Public ReadOnly Property BaseRecordset() As DataTable
		Get
			Return mrsCalendarBaseInfo
		End Get
	End Property

	Public ReadOnly Property GroupByDescription() As Boolean
		Get
			GroupByDescription = mblnGroupByDescription
		End Get
	End Property

	Public Property ShowBankHolidays() As Boolean
		Get
			ShowBankHolidays = mblnShowBankHolidays
		End Get
		Set(ByVal Value As Boolean)
			mblnShowBankHolidays = Value
		End Set
	End Property

	Public Property ShowCaptions() As Boolean
		Get
			ShowCaptions = mblnShowCaptions
		End Get
		Set(ByVal Value As Boolean)
			mblnShowCaptions = Value
		End Set
	End Property

	Public Property ShowWeekends() As Boolean
		Get
			ShowWeekends = mblnShowWeekends
		End Get
		Set(ByVal Value As Boolean)
			mblnShowWeekends = Value
		End Set
	End Property

	Public Property StartOnCurrentMonth() As Boolean
		Get
			Return mbStartOnCurrentMonth
		End Get
		Set(ByVal Value As Boolean)
			mbStartOnCurrentMonth = Value
		End Set
	End Property

	Public Property IncludeWorkingDaysOnly() As Boolean
		Get
			IncludeWorkingDaysOnly = mblnIncludeWorkingDaysOnly
		End Get
		Set(ByVal Value As Boolean)
			mblnIncludeWorkingDaysOnly = Value
		End Set
	End Property

	Public Property IncludeBankHolidays() As Boolean
		Get
			IncludeBankHolidays = mblnIncludeBankHolidays
		End Get
		Set(ByVal Value As Boolean)
			mblnIncludeBankHolidays = Value
		End Set
	End Property

	Public ReadOnly Property OutputArray_Definition() As Object
		Get

			' Holds the HTML for the grid definition (object tag etc)
			Return VB6.CopyArray(mvarOutputArray_Definition)

		End Get
	End Property

	Public ReadOnly Property OutputArray_Columns() As Object
		Get

			' Holds the HTML for the columns in the grid (2 + No. fields on report)
			Return VB6.CopyArray(mvarOutputArray_Columns)

		End Get
	End Property

	Public ReadOnly Property OutputArray_Merges() As Object
		Get
			Return VB6.CopyArray(mvarOutputArray_Merges)
		End Get
	End Property

	Public ReadOnly Property OutputArray_Styles() As Object
		Get

			OutputArray_Styles = VB6.CopyArray(mvarOutputArray_Styles)

		End Get
	End Property

	Public ReadOnly Property OutputArray_Data() As Object
		Get

			' Holds the HTML for the actual data (and closes object tag)
			OutputArray_Data = VB6.CopyArray(mvarOutputArray_Data)

		End Get
	End Property

	Public Sub OutputArray_Clear()
		ReDim mvarOutputArray_Definition(0)
		ReDim mvarOutputArray_Columns(0)
		ReDim mvarOutputArray_Data(0)
		ReDim mvarOutputArray_Styles(0)
		ReDim mvarOutputArray_Merges(0)
	End Sub

	Private Function CreateTempTable() As Boolean

		Dim strSQL As String

		Try
			mstrTempTableName = General.UniqueSQLObjectName("ASRSysTempCalendarReport", 3)
			strSQL = String.Format("CREATE TABLE [{0}] ({1})", mstrTempTableName, mstrSQLCreateTable)
			DB.ExecuteSql(strSQL)
			mblnTempTableCreated = True

		Catch ex As Exception
			Throw

		End Try

		Return True
	End Function

	Public Function ConvertEventDescription(ByVal plngColumnID As Integer, ByVal pvarValue As String) As String

		Dim strTempEventDesc As String
		Dim iDecimals As Short
		Dim strFormat As String

		'get the datatype/properties for the desc1 column
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If (plngColumnID > 0) And (Not IsDBNull(pvarValue)) Then
			If DoesColumnUseSeparators(plngColumnID) Then
				iDecimals = GetDecimalsSize(plngColumnID)
				strFormat = "#,0" & IIf(iDecimals > 0, "." & New String("#", iDecimals), "")
				strTempEventDesc = Format(pvarValue, strFormat)

			ElseIf GetColumnDataType(plngColumnID) = ColumnDataType.sqlBoolean Then
				strTempEventDesc = pvarValue

			ElseIf GetColumnDataType(plngColumnID) = ColumnDataType.sqlDate Then
				strTempEventDesc = VB6.Format(pvarValue, mstrClientDateFormat)

			Else
				strTempEventDesc = pvarValue

			End If
		Else
			strTempEventDesc = vbNullString
		End If

		ConvertEventDescription = strTempEventDesc

	End Function

	Private Function GetDescriptionDataTypes() As Boolean

		Dim iDecimals As Short

		mintType_BaseDesc1 = -1
		mintType_BaseDesc2 = -1
		mintType_BaseDescExpr = -1
		mstrFormat_BaseDesc1 = vbNullString
		mstrFormat_BaseDesc2 = vbNullString

		'get the datatype/properties for the desc1 column
		If (mlngDescription1 > 0) Then
			If DoesColumnUseSeparators(mlngDescription1) Then
				mintType_BaseDesc1 = 3
				iDecimals = GetDecimalsSize(mlngDescription1)
				mstrFormat_BaseDesc1 = "#,0" & IIf(iDecimals > 0, "." & New String("#", iDecimals), "")
			ElseIf IsBitColumn("C", mlngCalendarReportsBaseTable, mlngDescription1) Then
				mintType_BaseDesc1 = 2
			ElseIf IsDateColumn("C", mlngCalendarReportsBaseTable, mlngDescription1) Then
				mintType_BaseDesc1 = 1
			Else
				mintType_BaseDesc1 = 0
			End If
		End If
		'get the datatype/properties for the desc2 column
		If (mlngDescription2 > 0) Then
			If DoesColumnUseSeparators(mlngDescription2) Then
				mintType_BaseDesc2 = 3
				iDecimals = GetDecimalsSize(mlngDescription2)
				mstrFormat_BaseDesc2 = "#,0" & IIf(iDecimals > 0, "." & New String("#", iDecimals), "")
			ElseIf IsBitColumn("C", mlngCalendarReportsBaseTable, mlngDescription2) Then
				mintType_BaseDesc2 = 2
			ElseIf IsDateColumn("C", mlngCalendarReportsBaseTable, mlngDescription2) Then
				mintType_BaseDesc2 = 1
			Else
				mintType_BaseDesc2 = 0
			End If
		End If
		'get the datatype/properties for the descexpr column
		If (mlngDescriptionExpr > 0) Then
			If IsBitColumn("X", mlngCalendarReportsBaseTable, mlngDescriptionExpr) Then
				mintType_BaseDescExpr = 2
			ElseIf IsDateColumn("X", mlngCalendarReportsBaseTable, mlngDescriptionExpr) Then
				mintType_BaseDescExpr = 1
			Else
				mintType_BaseDescExpr = 0
			End If
		End If

	End Function

	Private Function InsertIntoTempTable(ByRef pstrSelectString As String) As Boolean

		Dim strSQL As String
		Dim fOK As Boolean

		fOK = True

		If Not mblnTempTableCreated Then
			If Not CreateTempTable() Then
				InsertIntoTempTable = False
				mstrErrorString = "Error creating the temporary table"
				GoTo TidyUpAndExit
			End If
		End If

		If (Not mblnUDFsCreated) Then
			fOK = General.UDFFunctions(mastrUDFsRequired, True)
			mblnUDFsCreated = fOK
			If Not fOK Then
				InsertIntoTempTable = False
				mstrErrorString = "Error creating SQL User Defined Functions"
				GoTo TidyUpAndExit
			End If
		End If

		strSQL = "INSERT INTO [" & mstrTempTableName & "] " & pstrSelectString
		DB.ExecuteSql(strSQL)

		Return True

TidyUpAndExit:
		Exit Function

	End Function

	Public Function IsBankHoliday(pdtDate As Date, plngBaseID As Integer, pstrRegion As String) As Boolean

		Dim colBankHolidays As clsBankHolidays
		Dim objBankHoliday As clsBankHoliday

		Try

			If mblnPersonnelBase And (PersonnelModule.grtRegionType = RegionType.rtHistoricRegion) And (Not mblnGroupByDescription) And (mlngRegion < 1) Then

				'Need to get the current region from the previously populated.
				'NB. cant get the region from the collection as the current region is required even
				'when the date is NOT a bank holiday
				pstrRegion = GetCurrentRegion(plngBaseID, pdtDate)

				'Historic Region Bank Holidays
				colBankHolidays = mcolHistoricBankHolidays.Item(CStr(plngBaseID))

				For Each objBankHoliday In colBankHolidays.Collection
					With objBankHoliday
						If pdtDate = .HolidayDate Then
							Return True
						End If
					End With
				Next objBankHoliday

			ElseIf ((mlngRegion > 0) Or (mblnPersonnelBase And (PersonnelModule.grtRegionType = RegionType.rtStaticRegion))) And (Not mblnGroupByDescription) Then

				'Static Region Bank Holidays
				colBankHolidays = mcolStaticBankHolidays.Item(CStr(plngBaseID))

				For Each objBankHoliday In colBankHolidays.Collection
					With objBankHoliday
						If pdtDate = .HolidayDate Then
							Return True
						End If
					End With
				Next objBankHoliday

			End If

		Catch ex As Exception
			Return False

		End Try

		Return False

	End Function

	Private Function GetCurrentRegion(plngBaseRecordID As Integer, pdtDate As Date) As String

		Dim intCount As Integer

		Try

			For intCount = 1 To UBound(mavCareerRanges, 2) Step 1
				'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If plngBaseRecordID = CInt(mavCareerRanges(0, intCount)) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(2, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If mavCareerRanges(2, intCount) <> "" Then
						'has a career change in the past
						'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If (pdtDate >= CDate(mavCareerRanges(1, intCount))) And (pdtDate < CDate(mavCareerRanges(2, intCount))) Then
							'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(3, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Return mavCareerRanges(3, intCount)
						End If
					Else
						'has a effective start date but has no end date. (most recent career change)
						'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If (pdtDate >= CDate(mavCareerRanges(1, intCount))) Then
							'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(3, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Return mavCareerRanges(3, intCount)
						End If
					End If
				End If
			Next intCount

		Catch ex As Exception
			Return ""

		End Try

		Return ""

	End Function

	Public Function Output_GetCalArrayIndex(ByRef pintBaseRecordIndex As Short, ByRef pdtDate As Date, ByRef pblnSession As Boolean) As Short

		' This function returns the index value for the specified date and session.

		Dim dtFirstDate As Date
		Dim dtLastDate As Date

		Dim iCount As Short
		Dim varTempArray As Object

		'  dtFirstDate = CDate("01/" & mlngMonth_Output & "/" & mlngYear_Output)
		'  dtLastDate = CDate(DaysInMonth(dtFirstDate) & "/" & mlngMonth_Output & "/" & mlngYear_Output)
		dtFirstDate = mdtVisibleStartDate_Output
		dtLastDate = mdtVisibleEndDate_Output

		If (pdtDate < dtFirstDate) Or (pdtDate > dtLastDate) Then
			Output_GetCalArrayIndex = -1
			Exit Function
		End If

		'UPGRADE_WARNING: Couldn't resolve default property of object mavOutputDateIndex(2, pintBaseRecordIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		varTempArray = mavOutputDateIndex(2, pintBaseRecordIndex)

		For iCount = 1 To 74 Step 2

			'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(1, iCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If varTempArray(1, iCount) <> "  /  /    " Then
				'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(1, iCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (varTempArray(1, iCount) = VB6.Format(pdtDate, CALREP_DATEFORMAT)) Then
					Output_GetCalArrayIndex = IIf(pblnSession, iCount + 1, iCount)
					Exit Function
				End If
			End If

		Next iCount

		Output_GetCalArrayIndex = -1

	End Function

	Private Function HexValue(plngColour As Integer) As String

		Dim strHEX As String

		strHEX = Hex(plngColour)

		If Len(strHEX) < 6 Then
			strHEX = New String("0", 6 - Len(strHEX)) & strHEX
		End If

		HexValue = "&H" & strHEX

	End Function

	Private Function GetLegendColour(ByRef pstrEventKey As String) As String

		Dim i As Integer
		Dim lngTemp As Integer

		For i = 0 To UBound(mavLegend, 2) Step 1
			'UPGRADE_WARNING: Couldn't resolve default property of object mavLegend(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If UCase(RTrim(mavLegend(0, i))) = UCase(RTrim(pstrEventKey)) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object mavLegend(3, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngTemp = mavLegend(3, i)
				GetLegendColour = HexValue(lngTemp)
				Exit Function
			End If
		Next i

		Return ColorTranslator.ToOle(Color.Black).ToString

	End Function

	Private Function CheckColumnPermissions(plngTableID As Integer, pstrTableName As String, pstrColumnName As String, ByRef strSQLRef As String) As Boolean

		'This function checks if the current user has read(select) permissions
		'on this column. If the user only has access through views then the
		'relevent views are added to the mlngTableViews() array which in turn
		'are used to create the join part of the query.

		Dim lngTempTableID As Integer
		Dim strTempTableName As String
		Dim strTempColumnName As String
		Dim blnColumnOK As Boolean
		Dim blnFound As Boolean
		Dim blnNoSelect As Boolean
		Dim strSource As String
		Dim intNextIndex As Short
		Dim blnOK As Boolean
		Dim strTable As String = vbNullString
		Dim strColumn As String = vbNullString

		Dim pintNextIndex As Short

		Dim bDateColumn As Boolean

		' Set flags with their starting values
		blnOK = True
		blnNoSelect = False
		bDateColumn = False

		' Load the temp variables
		lngTempTableID = plngTableID
		strTempTableName = pstrTableName
		strTempColumnName = pstrColumnName

		' Check permission on that column
		mobjColumnPrivileges = GetColumnPrivileges(strTempTableName)
		mstrRealSource = gcoTablePrivileges.Item(strTempTableName).RealSource

		blnColumnOK = mobjColumnPrivileges.IsValid(strTempColumnName)

		If blnColumnOK Then
			blnColumnOK = mobjColumnPrivileges.Item(strTempColumnName).AllowSelect
		End If

		If blnColumnOK Then

			If mobjColumnPrivileges.Item(strTempColumnName).DataType = ColumnDataType.sqlDate Then
				bDateColumn = True
			End If

			' this column can be read direct from the tbl/view or from a parent table
			strTable = mstrRealSource
			strColumn = strTempColumnName

			' If the table isnt the base table (or its realsource) then
			' Check if it has already been added to the array. If not, add it.
			If lngTempTableID <> mlngCalendarReportsBaseTable Then
				blnFound = False
				For intNextIndex = 1 To UBound(mlngTableViews, 2)
					If mlngTableViews(1, intNextIndex) = 0 And mlngTableViews(2, intNextIndex) = lngTempTableID Then
						blnFound = True
						Exit For
					End If
				Next intNextIndex

				If Not blnFound Then
					intNextIndex = UBound(mlngTableViews, 2) + 1
					ReDim Preserve mlngTableViews(2, intNextIndex)
					mlngTableViews(1, intNextIndex) = 0
					mlngTableViews(2, intNextIndex) = lngTempTableID
				End If
			End If

			If bDateColumn And mblnCheckingDescColumn Then
				strSQLRef = SQLDateConvertToLocale(strTable & "." & strColumn)
			Else
				strSQLRef = strTable & "." & strColumn
			End If

		Else

			' this column cannot be read direct. If its from a parent, try parent views
			' Loop thru the views on the table, seeing if any have read permis for the column

			ReDim mstrViews(0)
			For Each mobjTableView In gcoTablePrivileges.Collection
				If (Not mobjTableView.IsTable) And (mobjTableView.TableID = lngTempTableID) And (mobjTableView.AllowSelect) Then

					strSource = mobjTableView.ViewName
					mstrRealSource = gcoTablePrivileges.Item(strSource).RealSource

					' Get the column permission for the view
					mobjColumnPrivileges = GetColumnPrivileges(strSource)

					' If we can see the column from this view
					If mobjColumnPrivileges.IsValid(strTempColumnName) Then
						If mobjColumnPrivileges.Item(strTempColumnName).AllowSelect Then

							ReDim Preserve mstrViews(UBound(mstrViews) + 1)
							mstrViews(UBound(mstrViews)) = mobjTableView.ViewName

							If mlngEventViewColumn > 0 Then
								ReDim Preserve mvarEventColumnViews(1, UBound(mvarEventColumnViews, 2) + 1)
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarEventColumnViews(0, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								mvarEventColumnViews(0, UBound(mvarEventColumnViews, 2)) = mobjTableView.ViewID
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarEventColumnViews(1, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								mvarEventColumnViews(1, UBound(mvarEventColumnViews, 2)) = mlngEventViewColumn
							End If

							' Check if view has already been added to the array
							blnFound = False
							For intNextIndex = 0 To UBound(mlngTableViews, 2)
								If mlngTableViews(1, intNextIndex) = 1 And mlngTableViews(2, intNextIndex) = mobjTableView.ViewID Then
									blnFound = True
									Exit For
								End If
							Next intNextIndex

							If Not blnFound Then
								' View hasnt yet been added, so add it !
								intNextIndex = UBound(mlngTableViews, 2) + 1
								ReDim Preserve mlngTableViews(2, intNextIndex)
								mlngTableViews(1, intNextIndex) = 1
								mlngTableViews(2, intNextIndex) = mobjTableView.ViewID
							End If

						End If
					End If
				End If

			Next mobjTableView

			'UPGRADE_NOTE: Object mobjTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			mobjTableView = Nothing

			' Does the user have select permission thru ANY views ?
			If UBound(mstrViews) = 0 Then
				blnNoSelect = True
			Else
				strSQLRef = ""
				For pintNextIndex = 1 To UBound(mstrViews)
					If pintNextIndex = 1 Then
						strSQLRef = "CASE"
					End If

					strSQLRef = strSQLRef & " WHEN NOT " & mstrViews(pintNextIndex) & "." & strTempColumnName & " IS NULL THEN "

					If bDateColumn And mblnCheckingDescColumn Then
						strSQLRef = strSQLRef & SQLDateConvertToLocale(mstrViews(pintNextIndex) & "." & strTempColumnName)
					Else
						strSQLRef = strSQLRef & mstrViews(pintNextIndex) & "." & strTempColumnName
					End If

				Next pintNextIndex

				If Len(strSQLRef) > 0 Then
					strSQLRef = strSQLRef & " ELSE NULL" & " END "
				End If

			End If

			' If we cant see a column, then get outta here
			If blnNoSelect Then
				strSQLRef = vbNullString
				CheckColumnPermissions = False
				mstrErrorString = vbNewLine & vbNewLine & "You do not have permission to see the column '" & pstrColumnName & "'" & vbNewLine & "either directly or through any views."
				Exit Function
			End If

			If Not blnOK Then
				strSQLRef = vbNullString
				CheckColumnPermissions = False
				Exit Function
			End If

		End If

		mlngEventViewColumn = 0
		Return True

	End Function

	Private Function GenerateSQLEvent(pstrEventKey As String, pstrDynamicKey As String, pstrDynamicName As String) As Boolean

		Dim fOK As Boolean = True

		If fOK Then fOK = GenerateSQLSelect(pstrEventKey, pstrDynamicKey, pstrDynamicName)
		If fOK Then fOK = GenerateSQLFrom()
		If fOK Then fOK = GenerateSQLJoin(pstrEventKey)
		If fOK Then fOK = GenerateSQLWhere(pstrEventKey)

		If fOK Then
			mstrSQLEvent = mstrSQLSelect & vbNewLine & mstrSQLFrom & vbNewLine & mstrSQLJoin & vbNewLine & mstrSQLWhere & vbNewLine
		End If

		' reset strings to hold the SQL statement
		mstrSQLSelect = vbNullString
		mstrSQLFrom = vbNullString
		mstrSQLJoin = vbNullString
		mstrSQLWhere = vbNullString

		Return fOK

	End Function

	' Get columns defined as a SortOrder and load into array
	Public Function GetOrderArray() As Boolean

		Dim rsTemp As DataTable
		Dim sSQL As String
		Dim intTemp As Integer

		Try
			sSQL = String.Format("SELECT o.ColumnID, o.OrderType, c.ColumnName FROM ASRSysCalendarReportOrder o" _
					& " INNER JOIN ASRSysTables t ON t.tableID = o.TableID" _
					& " INNER JOIN ASRSysColumns c ON c.ColumnID = o.ColumnID AND c.tableid = t.tableid" _
					& " WHERE CalendarReportID = {0}" _
					& " ORDER BY [OrderSequence]", mlngCalendarReportID)
			rsTemp = DB.GetDataTable(sSQL)

			With rsTemp
				If .Rows.Count = 0 Then
					mstrErrorString = "No columns have been defined as a sort order for the specified Calendar Report definition." & vbNewLine & "Please remove this definition and create a new one."
					Return False
				End If

				For Each objRow As DataRow In rsTemp.Rows

					intTemp = UBound(mvarSortOrder, 2) + 1
					ReDim Preserve mvarSortOrder(2, intTemp)

					'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(0, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarSortOrder(0, intTemp) = CInt(objRow("ColumnID"))
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(1, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarSortOrder(1, intTemp) = objRow("ColumnName").ToString()
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(2, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarSortOrder(2, intTemp) = objRow("OrderType")

				Next
			End With

		Catch ex As Exception
			mstrErrorString = "Error whilst retrieving the event details recordsets'." & vbNewLine & ex.Message
			Return False

		End Try

		Return True

	End Function

	Public Function SetPromptedValues(ByRef pavPromptedValues As Object) As Boolean

		' Purpose : This function calls the individual functions that
		'           generate the components of the main SQL string.

		Dim fOK As Boolean = True
		Dim iLoop As Short
		Dim iDataType As Short
		Dim lngComponentID As Integer

		Try


			ReDim mvarPrompts(1, 0)

			If IsArray(pavPromptedValues) Then
				ReDim mvarPrompts(1, UBound(pavPromptedValues, 2))

				For iLoop = 0 To UBound(pavPromptedValues, 2)
					' Get the prompt data type.
					'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If Len(Trim(Mid(pavPromptedValues(0, iLoop), 10))) > 0 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						lngComponentID = CInt(Mid(pavPromptedValues(0, iLoop), 10))
						'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						iDataType = CShort(Mid(pavPromptedValues(0, iLoop), 8, 1))

						'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(0, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarPrompts(0, iLoop) = lngComponentID

						' NB. Locale to server conversions are done on the client.
						Select Case iDataType
							Case 2
								' Numeric.
								'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								mvarPrompts(1, iLoop) = CDbl(pavPromptedValues(1, iLoop))
							Case 3
								' Logic.
								'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								mvarPrompts(1, iLoop) = (UCase(CStr(pavPromptedValues(1, iLoop))) = "TRUE")
							Case 4
								' Date.
								' JPD 20040212 Fault 8082 - DO NOT CONVERT DATE PROMPTED VALUES
								' THEY ARE PASSED IN FROM THE ASPs AS STRING VALUES IN THE CORRECT
								' FORMAT (mm/dd/yyyy) AND DOING ANY KIND OF CONVERSION JUST SCREWS
								' THINGS UP.
								'mvarPrompts(1, iLoop) = CDate(Format(pavPromptedValues(1, iLoop), "MM/dd/yyyy"))
								'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								mvarPrompts(1, iLoop) = pavPromptedValues(1, iLoop)
							Case Else
								'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								mvarPrompts(1, iLoop) = CStr(pavPromptedValues(1, iLoop))
						End Select
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(0, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarPrompts(0, iLoop) = 0
					End If
				Next iLoop
			End If

		Catch ex As Exception
			mstrErrorString = "Error setting prompted values." & vbNewLine & ex.Message
			Logs.AddDetailEntry(mstrErrorString)
			Logs.ChangeHeaderStatus(EventLog_Status.elsFailed)
			Return False

		End Try

		Return fOK

	End Function

	Public Function WorkingPatternTitle() As String

		Return "<tr align=middle>" & vbNewLine _
			& "   <td ALIGN=center VALIGN=middle></TD>" & vbNewLine _
			& "   <td ALIGN=center VALIGN=middle>" & Left(WeekdayName(1, True, FirstDayOfWeek.Sunday), 1) & "</td>" & vbNewLine _
			& "   <td ALIGN=center VALIGN=middle>" & Left(WeekdayName(2, True, FirstDayOfWeek.Sunday), 1) & "</td>" & vbNewLine _
			& "   <td ALIGN=center VALIGN=middle>" & Left(WeekdayName(3, True, FirstDayOfWeek.Sunday), 1) & "</td>" & vbNewLine _
			& "   <td ALIGN=center VALIGN=middle>" & Left(WeekdayName(4, True, FirstDayOfWeek.Sunday), 1) & "</td>" & vbNewLine _
			& "   <td ALIGN=center VALIGN=middle>" & Left(WeekdayName(5, True, FirstDayOfWeek.Sunday), 1) & "</td>" & vbNewLine _
			& "   <td ALIGN=center VALIGN=middle>" & Left(WeekdayName(6, True, FirstDayOfWeek.Sunday), 1) & "</td>" & vbNewLine _
			& "   <td ALIGN=center VALIGN=middle>" & Left(WeekdayName(7, True, FirstDayOfWeek.Sunday), 1) & "</td>" & vbNewLine _
			& "</tr>" & vbNewLine

	End Function

	Public Function Write_Static_Historic_Forms() As String

		Write_Static_Historic_Forms = mstrWPFormString & vbNewLine & vbNewLine

		Write_Static_Historic_Forms = Write_Static_Historic_Forms & mstrBHolFormString.ToString() & vbNewLine & vbNewLine

		Write_Static_Historic_Forms = Write_Static_Historic_Forms & mstrRegionFormString & vbNewLine & vbNewLine

	End Function

	Public Sub Initialise()

		' Purpose : Sets references to other classes and redimensions arrays
		'           used for table usage information
		Dim rstData As DataTable

		Legend = New List(Of CalendarLegend)()

		ReDim mvarSortOrder(2, 0)
		ReDim mlngTableViews(2, 0)
		ReDim mstrViews(0)
		ReDim mastrUDFsRequired(0)
		ReDim mvarTableViews(3, 0)

		ReDim mvarOutputArray_Definition(0)
		ReDim mvarOutputArray_Columns(0)
		ReDim mvarOutputArray_Data(0)
		ReDim mvarOutputArray_Styles(0)
		ReDim mvarOutputArray_Merges(0)

		ReDim mvarEventColumnViews(1, 0)

		LegendColors = New List(Of LegendColor)()

		rstData = DB.GetDataTable("spASRIntGetCalendarColours", CommandType.StoredProcedure)
		For Each objRow As DataRow In rstData.Rows
			Dim objItem = New LegendColor
			objItem.ColOrder = objRow("colorder").ToString()
			objItem.ColValue = objRow("ColValue").ToString
			objItem.ColDesc = objRow("ColDesc").ToString
			objItem.WordColorIndex = objRow("WordColourIndex").ToString
			objItem.IsCalendarLegendColor = objRow("CalendarLegendColour")
			LegendColors.Add(objItem)
		Next

		' Add bank holiday to the legend
		Dim objLegendEvent As New CalendarLegend
		objLegendEvent.LegendKey = "Bank Holiday"
		objLegendEvent.LegendDescription = "Bank Holiday"
		objLegendEvent.HexColor = "#74B8FD"
		Legend.Add(objLegendEvent)

	End Sub

	Private Function IsColumnInView(plngViewID As Integer, plngColumnID As Integer) As Boolean

		Dim lngCount As Integer

		For lngCount = 1 To UBound(mvarEventColumnViews, 2) Step 1
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarEventColumnViews(1, lngCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarEventColumnViews(0, lngCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (mvarEventColumnViews(0, lngCount) = plngViewID) And (mvarEventColumnViews(1, lngCount) = plngColumnID) Then
				Return True
			End If
		Next lngCount

		Return False

	End Function

	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()

		' Purpose : Clears references to other classes.

		'Set mfrmOutput = Nothing
		'UPGRADE_NOTE: Object mcolEvents may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mcolEvents = Nothing
		'UPGRADE_NOTE: Object mobjTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mobjTableView = Nothing
		'UPGRADE_NOTE: Object mobjColumnPrivileges may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mobjColumnPrivileges = Nothing

		'UPGRADE_NOTE: Object mcolHistoricBankHolidays may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mcolHistoricBankHolidays = Nothing
		'UPGRADE_NOTE: Object mcolStaticBankHolidays may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mcolStaticBankHolidays = Nothing
		'UPGRADE_NOTE: Object mcolHistoricWorkingPatterns may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mcolHistoricWorkingPatterns = Nothing
		'UPGRADE_NOTE: Object mcolStaticWorkingPatterns may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mcolStaticWorkingPatterns = Nothing

	End Sub

	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub

	Public Function ExecuteSql() As Boolean

		' Purpose : This function executes the SQL string 'into' a recordset.

		Try


			'  'get all the base & event data into a recordset
			mstrSQL = String.Format("SELECT * FROM [{0}] ", mstrTempTableName)

			'get the ORDER BY statement which applies to the entire UNIONed query.
			GenerateSQLOrderBy()
			mstrSQL = mstrSQL & mstrSQLOrderBy

			mrsCalendarReportsOutput = DB.GetDataTable(mstrSQL)

			If mrsCalendarReportsOutput.Rows.Count = 0 Then
				ExecuteSql = False
				mstrErrorString = "No records meet the selection criteria."
				mblnNoRecords = True
				Logs.ChangeHeaderStatus(EventLog_Status.elsSuccessful)
				Logs.AddDetailEntry(mstrErrorString)
				Exit Function
			End If

			MultipleCheck()

			'get only the base table info into a recordset
			mrsCalendarBaseInfo = DB.GetDataTable(mstrSQLBaseData)

			If mrsCalendarBaseInfo.Rows.Count = 0 Then
				ExecuteSql = False
				mstrErrorString = "No records meet the selection criteria."
				mblnNoRecords = True
				Logs.ChangeHeaderStatus(EventLog_Status.elsSuccessful)
				Logs.AddDetailEntry(mstrErrorString)
				Exit Function
			End If

			GetDescriptionDataTypes()

			'TM08102003
			General.UDFFunctions(mastrUDFsRequired, False)

		Catch ex As Exception
			mstrErrorString = "Error whilst executing SQL statement." & vbNewLine & ex.Message.RemoveSensitive()
			Return False

		End Try

		Return True

	End Function

	Private Function MultipleCheck() As Boolean

		Dim rsMultiple As DataTable
		Dim sSQL As String
		Dim dtSD As Date
		Dim dtED As Date
		Dim strStartSession As String
		Dim strEndSession As String
		Dim lngBaseID As Integer
		Dim strDescription1 As String
		Dim strDescription2 As String
		Dim strDescriptionExpr As String
		Dim lngCurrentBaseID As Integer
		Dim avDateRanges(,) As Object
		Dim i As Short
		Dim intNewIndex As Short
		Dim strFullDesc As String
		Dim strCurrentDesc As String
		Dim blnFirstCalendarRecord As Boolean = True

		ReDim avDateRanges(6, 0)

		sSQL = "SELECT [BaseID], [Description1], [Description2], [DescriptionExpr], [StartDate], [StartSession], [EndDate], [EndSession] " _
			& "FROM [" & _login.Username & "].[" & mstrTempTableName & "] " & mstrSQLOrderBy & vbNewLine
		rsMultiple = DB.GetDataTable(sSQL)

		If Not rsMultiple Is Nothing Then
			With rsMultiple
				For Each objRow As DataRow In .Rows
					dtSD = objRow("StartDate")
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					dtED = IIf(IsDBNull(objRow("EndDate")), objRow("StartDate"), objRow("EndDate"))
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					strStartSession = IIf(IsDBNull(objRow("StartSession")), "AM", objRow("StartSession"))
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If ((IsDBNull(objRow("EndDate"))) And (IsDBNull(objRow("EndSession"))) And (IsDBNull(objRow("StartSession")))) Then
						strEndSession = "PM"
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					ElseIf ((IsDBNull(objRow("EndDate"))) And (IsDBNull(objRow("EndSession")))) Then
						strEndSession = strStartSession
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					ElseIf IsDBNull(objRow("EndSession")) Then
						strEndSession = "PM"
					Else
						strEndSession = objRow("EndSession")
					End If

					If mblnGroupByDescription Then
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						strDescription1 = IIf(IsDBNull(objRow("Description1")), "", objRow("Description1"))
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						strDescription2 = IIf(IsDBNull(objRow("Description2")), "", objRow("Description2"))
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						strDescriptionExpr = IIf(IsDBNull(objRow("DescriptionExpr")), "", objRow("DescriptionExpr"))
						strFullDesc = strDescription1 & mstrDescriptionSeparator & strDescription2 & mstrDescriptionSeparator & strDescriptionExpr

						If (strFullDesc <> strCurrentDesc) Or blnFirstCalendarRecord Then
							strCurrentDesc = strFullDesc
							blnFirstCalendarRecord = False

							ReDim avDateRanges(6, 0)
							'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(0, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							avDateRanges(0, 0) = strFullDesc
							'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(1, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							avDateRanges(1, 0) = dtSD
							'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(2, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							avDateRanges(2, 0) = dtED
							'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(3, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							avDateRanges(3, 0) = strStartSession
							'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(4, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							avDateRanges(4, 0) = strEndSession

						Else
							'Loop through the array for the current calendar row, checking if any dates overlap.
							For i = 0 To UBound(avDateRanges, 2) Step 1

								'if the start or end dates 'equal' any other start orend dates then check if the sessions are also equal.
								'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(4, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(2, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(3, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(1, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If ((dtSD = avDateRanges(1, i)) And (strStartSession = avDateRanges(3, i))) Or ((dtSD = avDateRanges(2, i)) And (strStartSession = avDateRanges(4, i))) Or ((dtED = avDateRanges(1, i)) And (strEndSession = avDateRanges(3, i))) Or ((dtED = avDateRanges(2, i)) And (strEndSession = avDateRanges(4, i))) Then
									mblnHasMultipleEvents = True
									MultipleCheck = True
									GoTo TidyUpAndExit
								End If

								'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(2, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(1, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If ((dtSD > avDateRanges(1, i)) And (dtED < avDateRanges(2, i))) Or ((dtSD > avDateRanges(1, i)) And (dtSD < avDateRanges(2, i)) And (dtED > avDateRanges(2, i))) Or ((dtED > avDateRanges(1, i)) And (dtED < avDateRanges(2, i)) And (dtSD < avDateRanges(1, i))) Or ((dtSD < avDateRanges(1, i)) And (dtED > avDateRanges(2, i))) Then
									mblnHasMultipleEvents = True
									MultipleCheck = True
									GoTo TidyUpAndExit
								End If
							Next i

							intNewIndex = UBound(avDateRanges, 2) + 1
							ReDim Preserve avDateRanges(6, intNewIndex)
							'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(0, intNewIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							avDateRanges(0, intNewIndex) = lngBaseID
							'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(1, intNewIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							avDateRanges(1, intNewIndex) = dtSD
							'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(2, intNewIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							avDateRanges(2, intNewIndex) = dtED
							'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(3, intNewIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							avDateRanges(3, intNewIndex) = strStartSession
							'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(4, intNewIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							avDateRanges(4, intNewIndex) = strEndSession

						End If
					Else
						lngBaseID = CInt(objRow("BaseID"))

						If (lngBaseID <> lngCurrentBaseID) Or blnFirstCalendarRecord Then
							lngCurrentBaseID = lngBaseID
							blnFirstCalendarRecord = False

							ReDim avDateRanges(6, 0)
							'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(0, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							avDateRanges(0, 0) = lngBaseID
							'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(1, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							avDateRanges(1, 0) = dtSD
							'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(2, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							avDateRanges(2, 0) = dtED
							'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(3, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							avDateRanges(3, 0) = strStartSession
							'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(4, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							avDateRanges(4, 0) = strEndSession

						Else
							'Loop through the array for the current calendar row, checking if any dates overlap.
							For i = 0 To UBound(avDateRanges, 2) Step 1

								'if the start or end dates 'equal' any other start orend dates then check if the sessions are also equal.
								'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(4, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(2, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(3, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(1, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If ((dtSD = avDateRanges(1, i)) And (strStartSession = avDateRanges(3, i))) Or ((dtSD = avDateRanges(2, i)) And (strStartSession = avDateRanges(4, i))) Or ((dtED = avDateRanges(1, i)) And (strEndSession = avDateRanges(3, i))) Or ((dtED = avDateRanges(2, i)) And (strEndSession = avDateRanges(4, i))) Then
									mblnHasMultipleEvents = True
									MultipleCheck = True
									GoTo TidyUpAndExit
								End If

								'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(2, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(1, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If ((dtSD > avDateRanges(1, i)) And (dtED < avDateRanges(2, i))) Or ((dtSD > avDateRanges(1, i)) And (dtSD < avDateRanges(2, i)) And (dtED > avDateRanges(2, i))) Or ((dtED > avDateRanges(1, i)) And (dtED < avDateRanges(2, i)) And (dtSD < avDateRanges(1, i))) Or ((dtSD < avDateRanges(1, i)) And (dtED > avDateRanges(2, i))) Then
									mblnHasMultipleEvents = True
									MultipleCheck = True
									GoTo TidyUpAndExit
								End If
							Next i

							intNewIndex = UBound(avDateRanges, 2) + 1
							ReDim Preserve avDateRanges(6, intNewIndex)
							'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(0, intNewIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							avDateRanges(0, intNewIndex) = lngBaseID
							'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(1, intNewIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							avDateRanges(1, intNewIndex) = dtSD
							'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(2, intNewIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							avDateRanges(2, intNewIndex) = dtED
							'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(3, intNewIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							avDateRanges(3, intNewIndex) = strStartSession
							'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(4, intNewIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							avDateRanges(4, intNewIndex) = strEndSession

						End If
					End If

				Next

			End With
		End If

		mblnHasMultipleEvents = False

		Return True

TidyUpAndExit:
		'UPGRADE_NOTE: Object rsMultiple may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsMultiple = Nothing
		Exit Function

	End Function

	Public Function GetCalendarReportDefinition() As Boolean

		' Purpose : This function retrieves the basic definition details
		'           and stores it in module level variables

		
		Dim rsTemp As DataTable

		Dim sSQL As String
		Dim sDateInterval As String

		Dim rsIDs As DataTable
		Dim blnOK As Boolean
		Dim iStartDateType As CalendarDataType
		Dim iEndDateType As CalendarDataType

		Try

			mstrSQLIDs = vbNullString

			sSQL = String.Format("SELECT * FROM ASRSYSCalendarReports WHERE ID = {0}", mlngCalendarReportID)
			rsTemp = DB.GetDataTable(sSQL)

			Dim pblnOK As Object
			Dim objTableView As TablePrivilege
			Dim objExpression As clsExprExpression
			With rsTemp

				If .Rows.Count = 0 Then
					mstrErrorString = "Could not find specified Calendar Report definition."
					Return False
				End If

				Dim rowDefinition = .Rows(0)

				'JPD 20040729 Fault 8972 & Fault 8990
				If LCase(rowDefinition("Username").ToString()) <> LCase(_login.Username) And CurrentUserAccess(UtilityType.utlCalendarReport, mlngCalendarReportID) = ACCESS_HIDDEN Then
					mstrErrorString = "Report has been made hidden by another user."
					mblnNoRecords = True
					Return False
				End If

				Name = rowDefinition("Name").ToString
				Logs.AddHeader(EventLog_Type.eltCalandarReport, Name)
				mlngCalendarReportsBaseTable = CInt(rowDefinition("BaseTable"))
				mstrCalendarReportsBaseTableName = GetTableName(mlngCalendarReportsBaseTable)

				' Check the user has permission to read the base table.
				'UPGRADE_WARNING: Couldn't resolve default property of object pblnOK. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pblnOK = False
				For Each objTableView In gcoTablePrivileges.Collection
					If (objTableView.TableID = mlngCalendarReportsBaseTable) And (objTableView.AllowSelect) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object pblnOK. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						pblnOK = True
						Exit For
					End If
				Next objTableView
				'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				objTableView = Nothing
				If Not pblnOK Then
					mstrErrorString = "You do not have permission to read the base table" & vbNewLine & "either directly or through any views."
					mblnNoRecords = True
					Return False
				End If

				mlngCalendarReportsPickListID = CInt(rowDefinition("picklist"))
				mlngCalendarReportsFilterID = CInt(rowDefinition("Filter"))

				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				mlngDescription1 = IIf(IsDBNull(rowDefinition("Description1")), 0, rowDefinition("Description1"))
				If mlngDescription1 > 0 Then
					mstrDescription1 = GetColumnName(rowDefinition("Description1"))
					mblnDesc1IsDate = (GetDataType(mlngCalendarReportsBaseTable, mlngDescription1) = ColumnDataType.sqlDate)
				Else
					mstrDescription1 = vbNullString
					mblnDesc1IsDate = False
				End If

				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				mlngDescription2 = IIf(IsDBNull(rowDefinition("Description2")), 0, rowDefinition("Description2"))
				If mlngDescription2 > 0 Then
					mstrDescription2 = GetColumnName(rowDefinition("Description2"))
					mblnDesc2IsDate = (GetDataType(mlngCalendarReportsBaseTable, mlngDescription2) = ColumnDataType.sqlDate)
				Else
					mstrDescription2 = vbNullString
					mblnDesc2IsDate = False
				End If

				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				mlngDescriptionExpr = IIf(IsDBNull(rowDefinition("DescriptionExpr")), 0, rowDefinition("DescriptionExpr"))
				If mlngDescriptionExpr > 0 Then

					objExpression = NewExpression()
					objExpression.ExpressionID = mlngDescriptionExpr
					objExpression.ConstructExpression()
					objExpression.ValidateExpression(True)
					If objExpression.ReturnType = 4 Then ' its date
						mblnDescExprIsDate = True
					Else
						mblnDescExprIsDate = False
					End If
					mlngBaseDescriptionType = objExpression.ReturnType
				Else
					mlngBaseDescriptionType = -1
					mblnDescExprIsDate = False
				End If
				'UPGRADE_NOTE: Object objExpression may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				objExpression = Nothing

				mlngRegion = rowDefinition("Region")
				If mlngRegion > 0 Then
					mstrRegion = GetColumnName(rowDefinition("Region"))

				ElseIf (mlngCalendarReportsBaseTable = PersonnelModule.glngPersonnelTableID) And (PersonnelModule.grtRegionType = RegionType.rtStaticRegion) Then

					mlngRegion = BankHolidayModule.glngBHolRegionID
					mstrRegion = BankHolidayModule.gsBHolRegionColumnName

				Else
					mstrRegion = vbNullString

				End If

				mblnGroupByDescription = IIf(rowDefinition("GroupByDesc"), True, False)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				mstrDescriptionSeparator = IIf(IsDBNull(rowDefinition("DescriptionSeparator")), " ", rowDefinition("DescriptionSeparator"))

				'create the events collection here so that the event filters can bee checked
				If Not GetEventsCollection() Then
					Return False
				End If

				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				mlngStartDateExpr = IIf(IsDBNull(rowDefinition("StartDateExpr")), 0, rowDefinition("StartDateExpr"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				mlngEndDateExpr = IIf(IsDBNull(rowDefinition("EndDateExpr")), 0, rowDefinition("EndDateExpr"))

				If Not IsRecordSelectionValid() Then
					Return False
				End If

				'************** Must do the dates stuff here *****************
				'calculate and store the start and end dates
				iStartDateType = rowDefinition("StartType")
				iEndDateType = rowDefinition("EndType")

				'START DATE
				Select Case iStartDateType
					Case CalendarDataType.Fixed
						mdtStartDate = rowDefinition("FixedStart")
					Case CalendarDataType.CurrentDate
						mdtStartDate = Today
					Case CalendarDataType.Custom
						mdtStartDate = GetValueForRecordIndependantCalc(mlngStartDateExpr, mvarPrompts)
				End Select

				'END DATE
				Select Case iEndDateType
					Case CalendarDataType.Fixed
						mdtEndDate = rowDefinition("FixedEnd")
					Case CalendarDataType.CurrentDate
						mdtEndDate = Today
					Case CalendarDataType.Custom
						mdtEndDate = CDate(GetValueForRecordIndependantCalc(mlngEndDateExpr, mvarPrompts))
				End Select

				If iStartDateType = CalendarDataType.Offset And iEndDateType = CalendarDataType.Offset Then
					'START DATE
					Select Case rowDefinition("StartPeriod")
						Case DatePeriod.Days : sDateInterval = "d"
						Case DatePeriod.Weeks : sDateInterval = "ww"
						Case DatePeriod.Months : sDateInterval = "m"
						Case DatePeriod.Years : sDateInterval = "yyyy"
					End Select
					mdtStartDate = DateAdd(sDateInterval, CDbl(rowDefinition("StartFrequency")), Today)

					'END DATE
					Select Case rowDefinition("EndPeriod")
						Case DatePeriod.Days : sDateInterval = "d"
						Case DatePeriod.Weeks : sDateInterval = "ww"
						Case DatePeriod.Months : sDateInterval = "m"
						Case DatePeriod.Years : sDateInterval = "yyyy"
					End Select
					mdtEndDate = DateAdd(sDateInterval, CDbl(rowDefinition("EndFrequency")), Today)

				ElseIf iStartDateType = CalendarDataType.Offset And Not iEndDateType = CalendarDataType.Offset Then
					'START DATE
					Select Case rowDefinition("StartPeriod")
						Case DatePeriod.Days : sDateInterval = "d"
						Case DatePeriod.Weeks : sDateInterval = "ww"
						Case DatePeriod.Months : sDateInterval = "m"
						Case DatePeriod.Years : sDateInterval = "yyyy"
					End Select
					mdtStartDate = DateAdd(sDateInterval, CDbl(rowDefinition("StartFrequency")), mdtEndDate)

				ElseIf iEndDateType = CalendarDataType.Offset And Not iStartDateType = CalendarDataType.Offset Then
					'END DATE
					Select Case rowDefinition("EndPeriod")
						Case DatePeriod.Days : sDateInterval = "d"
						Case DatePeriod.Weeks : sDateInterval = "ww"
						Case DatePeriod.Months : sDateInterval = "m"
						Case DatePeriod.Years : sDateInterval = "yyyy"
					End Select
					mdtEndDate = CStr(DateAdd(sDateInterval, CDbl(rowDefinition("EndFrequency")), mdtStartDate))

				End If

				If mdtStartDate > mdtEndDate Then
					mstrErrorString = "The report end date is before the report start date."
					mblnNoRecords = True
					Return False
				End If

				'************************************************

				mblnShowBankHolidays = rowDefinition("ShowBankHolidays")
				mblnShowCaptions = rowDefinition("ShowCaptions")
				mblnShowWeekends = rowDefinition("ShowWeekends")
				mbStartOnCurrentMonth = rowDefinition("StartOnCurrentMonth")
				mblnIncludeWorkingDaysOnly = rowDefinition("IncludeWorkingDaysOnly")
				mblnIncludeBankHolidays = rowDefinition("IncludeBankHolidays")
				mblnCustomReportsPrintFilterHeader = rowDefinition("PrintFilterHeader")

				OutputPreview = rowDefinition("OutputPreview")
				OutputFormat = rowDefinition("OutputFormat")
				OutputScreen = rowDefinition("OutputScreen")
				mblnOutputPrinter = rowDefinition("OutputPrinter")
				mstrOutputPrinterName = rowDefinition("OutputPrinterName")
				mblnOutputSave = rowDefinition("OutputSave")
				mlngOutputSaveExisting = rowDefinition("OutputSaveExisting")
				mblnOutputEmail = rowDefinition("OutputEmail")
				mlngOutputEmailID = rowDefinition("OutputEmailAddr")
				mstrOutputEmailName = GetEmailGroupName(rowDefinition("OutputEmailAddr"))
				mstrOutputEmailSubject = rowDefinition("OutputEmailSubject")
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				mstrOutputEmailAttachAs = IIf(IsDBNull(rowDefinition("OutputEmailAttachAs")), vbNullString, rowDefinition("OutputEmailAttachAs"))
				OutputFilename = rowDefinition("OutputFilename")

				mblnPersonnelBase = (mlngCalendarReportsBaseTable = PersonnelModule.glngPersonnelTableID)

				If mblnCustomReportsPrintFilterHeader And (mlngSingleRecordID < 1) Then
					If (mlngCalendarReportsFilterID > 0) Then
						Name = Name & " (Base Table filter : " & General.GetFilterName(mlngCalendarReportsFilterID) & ")"
					ElseIf (mlngCalendarReportsPickListID > 0) Then
						Name = Name & " (Base Table picklist : " & General.GetPicklistName(mlngCalendarReportsPickListID) & ")"
					End If
				End If

				If mlngSingleRecordID > 0 Then
					'DebugMSG "Single Record ID = " & CStr(mlngSingleRecordID), True
					mstrSQLIDs = CStr(mlngSingleRecordID)

				ElseIf mlngCalendarReportsPickListID > 0 Then
					rsIDs = DB.GetDataTable("EXEC sp_ASRGetPickListRecords " & mlngCalendarReportsPickListID)

					If rsIDs.Rows.Count = 0 Then
						Logs.ChangeHeaderStatus(EventLog_Status.elsSuccessful)
						Logs.AddDetailEntry(mstrErrorString)
						mstrErrorString = "The selected picklist contains no records."
						Return False
					End If

					For Each objRow As DataRow In rsIDs.Rows
						mstrSQLIDs = mstrSQLIDs & IIf(Len(mstrSQLIDs) > 0, ", ", "") & objRow(0)
					Next


				ElseIf mlngCalendarReportsFilterID > 0 Then
					blnOK = FilteredIDs(mlngCalendarReportsFilterID, mstrFilteredIDs, mastrUDFsRequired, mvarPrompts)

					If blnOK Then
						blnOK = General.UDFFunctions(mastrUDFsRequired, True)
						If blnOK Then
							rsIDs = DB.GetDataTable(mstrFilteredIDs)
						End If

						If rsIDs.Rows.Count = 0 Then
							mstrErrorString = "The base table filter returned no records."
							Logs.ChangeHeaderStatus(EventLog_Status.elsSuccessful)
							Logs.AddDetailEntry(mstrErrorString)
							mblnNoRecords = True
							Return False
						End If

						For Each objRow As DataRow In rsIDs.Rows
							mstrSQLIDs = mstrSQLIDs & IIf(Len(mstrSQLIDs) > 0, ", ", "") & objRow(0)
						Next

						blnOK = General.UDFFunctions(mastrUDFsRequired, False)

					Else
						' Permission denied on something in the filter.
						mstrErrorString = "You do not have permission to use the '" & General.GetFilterName(mlngCalendarReportsFilterID) & "' filter."
						Logs.ChangeHeaderStatus(EventLog_Status.elsSuccessful)
						Logs.AddDetailEntry(mstrErrorString)
						mblnNoRecords = True
						Return False
					End If

				End If

			End With

			mstrBaseIDColumn = "?ID_" & mstrCalendarReportsBaseTableName
			mstrEventIDColumn = "?ID_EventID"

			mstrBaseTableRealSource = gcoTablePrivileges.Item(mstrCalendarReportsBaseTableName).RealSource

		Catch ex As Exception
			mstrErrorString = "Error whilst reading the Calendar Report definition." & vbNewLine & ex.Message.RemoveSensitive()
			Return False

		End Try

		Return True

	End Function

	Public Function Initialise_WP_Region() As Boolean

		Dim fOK As Boolean
		Dim blnRegionEnabled As Boolean
		Dim blnWorkingPatternEnabled As Boolean

		mcolHistoricBankHolidays = New Collection
		mcolStaticBankHolidays = New Collection
		mcolHistoricWorkingPatterns = New Collection
		mcolStaticWorkingPatterns = New Collection

		fOK = True

		blnRegionEnabled = False
		blnWorkingPatternEnabled = False

		If (fOK And mblnPersonnelBase And (PersonnelModule.grtRegionType = RegionType.rtHistoricRegion) And (Not mblnGroupByDescription) And (mlngRegion < 1)) Or (fOK And ((mlngRegion > 0) Or (mblnPersonnelBase And (PersonnelModule.grtRegionType = RegionType.rtStaticRegion))) And (Not mblnGroupByDescription)) Then

			blnRegionEnabled = CheckPermission_RegionInfo()
		End If

		If blnRegionEnabled Then
			If fOK And mblnPersonnelBase And (PersonnelModule.grtRegionType = RegionType.rtHistoricRegion) And (Not mblnGroupByDescription) And (mlngRegion < 1) Then

				'get historical bank holidays
				fOK = Get_HistoricBankHolidays()

			ElseIf fOK And ((mlngRegion > 0) Or (mblnPersonnelBase And (PersonnelModule.grtRegionType = RegionType.rtStaticRegion))) And (Not mblnGroupByDescription) Then

				'get static bank holidays collection
				fOK = Get_StaticBankHolidays()

				If fOK Then
					mblnStaticReg = True
				End If

			Else
				mblnDisableRegions = True

			End If
		End If


		If (fOK And mblnPersonnelBase And (PersonnelModule.gwptWorkingPatternType = WorkingPatternType.wptHistoricWPattern) And (Not mblnGroupByDescription)) Or (fOK And (mblnPersonnelBase And (PersonnelModule.gwptWorkingPatternType = WorkingPatternType.wptStaticWPattern) And (Not mblnGroupByDescription))) Then
			blnWorkingPatternEnabled = CheckPermission_WPInfo()
		End If

		If blnWorkingPatternEnabled Then
			If fOK And mblnPersonnelBase And (PersonnelModule.gwptWorkingPatternType = WorkingPatternType.wptHistoricWPattern) And (Not mblnGroupByDescription) Then

				'get historical working patterns
				fOK = Get_HistoricWorkingPatterns()

			ElseIf fOK And (mblnPersonnelBase And (PersonnelModule.gwptWorkingPatternType = WorkingPatternType.wptStaticWPattern) And (Not mblnGroupByDescription)) Then

				'get static working patterns
				fOK = Get_StaticWorkingPatterns()

				If fOK Then
					mblnStaticWP = True
				End If

			Else
				mblnDisableWPs = True

			End If
		End If

		Initialise_WP_Region = True

	End Function

	Public Function Get_HistoricWorkingPatterns() As Boolean

		Dim strSQLCC As String 'sql for retieving career change data

		Try

			If mblnDisableWPs Then
				Return False
			End If

			strSQLCC = "SELECT " & vbNewLine _
					& "     " & PersonnelModule.gsPersonnelHWorkingPatternTableRealSource & ".ID_" & mlngCalendarReportsBaseTable & " AS [BaseID]," & vbNewLine _
					& "     " & PersonnelModule.gsPersonnelHWorkingPatternTableRealSource & "." & PersonnelModule.gsPersonnelHWorkingPatternDateColumnName & " AS [WP_Date], " & vbNewLine _
					& "     " & PersonnelModule.gsPersonnelHWorkingPatternTableRealSource & "." & PersonnelModule.gsPersonnelHWorkingPatternColumnName & "	AS [WP_Pattern], " & vbNewLine _
					& "     (SELECT COUNT(B.ID) FROM " & PersonnelModule.gsPersonnelHWorkingPatternTableRealSource & " B WHERE B.ID_" & mlngCalendarReportsBaseTable & " = " & PersonnelModule.gsPersonnelHWorkingPatternTableRealSource & ".ID_" & mlngCalendarReportsBaseTable & " AND B." & PersonnelModule.gsPersonnelHWorkingPatternDateColumnName & " IS NOT NULL) AS 'CareerChanges' " & vbNewLine _
					& "FROM " & PersonnelModule.gsPersonnelHWorkingPatternTableRealSource & " " & vbNewLine
			If Len(Trim(mstrSQLIDs)) > 0 Then
				strSQLCC = strSQLCC & "WHERE " & vbNewLine _
						& "     " & PersonnelModule.gsPersonnelHWorkingPatternTableRealSource & ".ID_" & mlngCalendarReportsBaseTable & " IN (" & mstrSQLIDs & ") " & vbNewLine _
						& " AND " & PersonnelModule.gsPersonnelHWorkingPatternTableRealSource & "." & PersonnelModule.gsPersonnelHWorkingPatternDateColumnName & " IS NOT NULL " & vbNewLine
			Else
				strSQLCC = strSQLCC & "WHERE " & vbNewLine _
						& "      " & PersonnelModule.gsPersonnelHWorkingPatternTableRealSource & "." & PersonnelModule.gsPersonnelHWorkingPatternDateColumnName & " IS NOT NULL " & vbNewLine
			End If
			strSQLCC = strSQLCC & "ORDER BY " _
					& "     " & PersonnelModule.gsPersonnelHWorkingPatternTableRealSource & ".ID_" & mlngCalendarReportsBaseTable & ", " _
					& "     " & PersonnelModule.gsPersonnelHWorkingPatternTableRealSource & "." & PersonnelModule.gsPersonnelHWorkingPatternDateColumnName & " "
			rsCareerChange = DB.GetDataTable(strSQLCC)

		Catch ex As Exception
			Return False

		End Try

		Return True

	End Function

	Public Function Get_HistoricBankHolidays() As Boolean


		If mblnDisableRegions Then Return False

		Dim rsCC As DataTable	'career change data for base records
		Dim colBankHolidays As clsBankHolidays

		Dim strSQLCC As String 'sql for retieving career change data
		Dim strSQLAllBHols As String
		Dim strSQLSelect As String
		Dim strSQLWhere As String
		Dim strSQLDateRegion As String
		Dim strSQLOrder As String

		Dim dtStartDate As Date
		Dim dtEndDate As Date

		Dim intNextIndex As Short
		Dim blnNewBaseRecord As Boolean
		Dim lngBaseRecordID As Integer
		Dim lng100Counter As Integer
		Dim lngBaseRowCount As Integer
		Dim lngMainBaseCounter As Integer
		Dim lngTotalCareerChanges As Integer

		Dim intCount As Short
		Dim intBHolCount As Short
		Dim lngCount As Integer
		Dim fFinalCareerChange As Boolean

		ReDim mavCareerRanges(4, 0)

		Dim INPUT_STRING As String
		Dim intRecordBHol As Short

		Try

			rsPersonnelBHols = New DataTable("rsPersonnelBHols")

			intRecordBHol = 0
			mstrBHolFormString = New StringBuilder()
			mstrBHolFormString.Append("<FORM id=frmBHol name=frmBHol style=""visibility:hidden;display:none"">" & vbNewLine)

			strSQLCC = "SELECT " & PersonnelModule.gsPersonnelHRegionTableRealSource & ".ID_" & mlngCalendarReportsBaseTable & "," _
					& "     " & PersonnelModule.gsPersonnelHRegionTableRealSource & "." & PersonnelModule.gsPersonnelHRegionDateColumnName & ", " _
					& "     " & PersonnelModule.gsPersonnelHRegionTableRealSource & "." & PersonnelModule.gsPersonnelHRegionColumnName & ", " _
					& "     (SELECT COUNT(B.ID) FROM " & PersonnelModule.gsPersonnelHRegionTableRealSource & " B WHERE B.ID_" & mlngCalendarReportsBaseTable & " = " & PersonnelModule.gsPersonnelHRegionTableRealSource & ".ID_" & mlngCalendarReportsBaseTable & " AND B." & PersonnelModule.gsPersonnelHRegionDateColumnName & " IS NOT NULL) AS 'CareerChanges' " _
					& " FROM " & PersonnelModule.gsPersonnelHRegionTableRealSource & " " & vbNewLine

			If Len(Trim(mstrSQLIDs)) > 0 Then
				strSQLCC = strSQLCC & "WHERE " & vbNewLine _
						& "     " & PersonnelModule.gsPersonnelHRegionTableRealSource & ".ID_" & mlngCalendarReportsBaseTable & " IN (" & mstrSQLIDs & ") " & vbNewLine _
						& " AND " & PersonnelModule.gsPersonnelHRegionTableRealSource & "." & PersonnelModule.gsPersonnelHRegionDateColumnName & " IS NOT NULL " & vbNewLine
			Else
				strSQLCC = strSQLCC & "WHERE " & PersonnelModule.gsPersonnelHRegionTableRealSource & "." & PersonnelModule.gsPersonnelHRegionDateColumnName & " IS NOT NULL " & vbNewLine
			End If

			strSQLCC = strSQLCC & "ORDER BY " & PersonnelModule.gsPersonnelHRegionTableRealSource & ".ID_" & mlngCalendarReportsBaseTable & ", " _
					& "     " & PersonnelModule.gsPersonnelHRegionTableRealSource & "." & PersonnelModule.gsPersonnelHRegionDateColumnName & " "
			rsCC = DB.GetDataTable(strSQLCC)

			lngBaseRecordID = -1
			blnNewBaseRecord = False
			lng100Counter = 0
			lngMainBaseCounter = 0

			'******************************************************************************
			'Create an array containing the ranges of career change period

			With rsCC

				If Not (.Rows.Count = 0) Then

					For Each objRow As DataRow In .Rows

						intNextIndex = UBound(mavCareerRanges, 2) + 1
						ReDim Preserve mavCareerRanges(4, intNextIndex)

						If lngBaseRecordID <> objRow("ID_" & CStr(mlngCalendarReportsBaseTable)) Then
							lngBaseRecordID = objRow("ID_" & CStr(mlngCalendarReportsBaseTable))
							blnNewBaseRecord = True
							lngBaseRowCount = lngBaseRowCount + 1
							dtStartDate = objRow(PersonnelModule.gsPersonnelHRegionDateColumnName)

							'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(0, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mavCareerRanges(0, intNextIndex) = lngBaseRecordID 'BaseRecordID
							'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(1, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mavCareerRanges(1, intNextIndex) = dtStartDate 'Start Date
							'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(2, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mavCareerRanges(2, intNextIndex) = ""	'End Date
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(3, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mavCareerRanges(3, intNextIndex) = IIf(IsDBNull(objRow(PersonnelModule.gsPersonnelHRegionColumnName)), "", objRow(PersonnelModule.gsPersonnelHRegionColumnName))	'Region
							'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(4, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mavCareerRanges(4, intNextIndex) = objRow("CareerChanges")	'Career Change Count

						Else
							dtStartDate = objRow(PersonnelModule.gsPersonnelHRegionDateColumnName)

							dtEndDate = dtStartDate
							'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(2, intNextIndex - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mavCareerRanges(2, intNextIndex - 1) = dtEndDate 'End Date

							'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(0, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mavCareerRanges(0, intNextIndex) = lngBaseRecordID 'BaseRecordID
							'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(1, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mavCareerRanges(1, intNextIndex) = dtStartDate 'Start Date
							'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(2, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mavCareerRanges(2, intNextIndex) = ""	'End Date
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(3, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mavCareerRanges(3, intNextIndex) = IIf(IsDBNull(objRow(PersonnelModule.gsPersonnelHRegionColumnName)), "", objRow(PersonnelModule.gsPersonnelHRegionColumnName)) 'Region
							'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(4, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mavCareerRanges(4, intNextIndex) = objRow("CareerChanges")	'Career Change Count

						End If

						blnNewBaseRecord = False
					Next
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object Get_HistoricBankHolidays. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Return True

				End If

			End With

			lngTotalCareerChanges = UBound(mavCareerRanges, 2)

			'******************************************************************************

			lngBaseRecordID = -1
			blnNewBaseRecord = False

			INPUT_STRING = vbNullString
			intRecordBHol = 0

			mstrRegionFormString = "<form id=frmRegion name=frmRegion style=""visibility:hidden;display:none"">" & vbNewLine

			For intCount = 1 To UBound(mavCareerRanges, 2) Step 1

				'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If lngBaseRecordID <> CInt(mavCareerRanges(0, intCount)) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(4, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(0, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mstrRegionFormString = mstrRegionFormString & vbNewLine & vbTab & "<input NAME=txtRegionCOUNT_" & mavCareerRanges(0, intCount) & " ID=txtRegionCOUNT_" & mavCareerRanges(0, intCount) & " VALUE=""" & mavCareerRanges(4, intCount) & """>" & vbNewLine
					'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(0, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					lngBaseRecordID = mavCareerRanges(0, intCount)
					intRecordBHol = 0
				End If

				intRecordBHol = intRecordBHol + 1

				INPUT_STRING = vbNullString
				'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				INPUT_STRING = INPUT_STRING & mavCareerRanges(1, intCount) & "_"
				'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				INPUT_STRING = INPUT_STRING & mavCareerRanges(2, intCount) & "_"
				'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				INPUT_STRING = INPUT_STRING & mavCareerRanges(3, intCount)

				'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(0, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mstrRegionFormString = mstrRegionFormString & vbTab & "<input NAME=txtRegion_" & mavCareerRanges(0, intCount) & "_" & intRecordBHol & " ID=txtRegion_" & mavCareerRanges(0, intCount) & "_" & intRecordBHol & " VALUE=""" & INPUT_STRING & """>" & vbNewLine

			Next intCount

			mstrRegionFormString = mstrRegionFormString & "</form>" & vbNewLine


			lngBaseRecordID = -1
			blnNewBaseRecord = False

			'------------------------------------------------------------------------------
			'Create and execute a 'single' sql string which returns all the bank holidays
			'for all the selcted base table records.

			strSQLAllBHols = vbNullString
			strSQLSelect = vbNullString
			strSQLWhere = vbNullString
			strSQLDateRegion = vbNullString

			For intCount = 1 To UBound(mavCareerRanges, 2) Step 1

				'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(0, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If lngBaseRecordID <> mavCareerRanges(0, intCount) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(0, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					lngBaseRecordID = mavCareerRanges(0, intCount)
					blnNewBaseRecord = True
					lng100Counter = lng100Counter + 1
					lngMainBaseCounter = lngMainBaseCounter + 1
					strSQLSelect = vbNullString
					strSQLDateRegion = strSQLDateRegion & "         ( " & vbNewLine

					strSQLWhere = "WHERE " & vbNewLine

					intBHolCount = 0
				End If

				intBHolCount = intBHolCount + 1

				'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strSQLSelect = vbNewLine & "SELECT  '" & mavCareerRanges(0, intCount) & "' AS 'ID' , " _
				& "       " & mstrSQLSelect_RegInfoRegion & " AS 'Region', " _
				& "       " & mstrSQLSelect_BankHolDate & " , " _
				& "       " & mstrSQLSelect_BankHolDesc & " FROM " & BankHolidayModule.gsBHolRegionTableName & " " & vbNewLine

				For lngCount = 0 To UBound(mvarTableViews, 2) Step 1
					'<REGIONAL CODE>
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(0, lngCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If mvarTableViews(0, lngCount) = BankHolidayModule.glngBHolRegionTableID Then
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						strSQLSelect = strSQLSelect & "           LEFT OUTER JOIN " & mvarTableViews(3, lngCount) & vbNewLine _
								& "           ON  " & BankHolidayModule.gsBHolRegionTableName & ".ID = " & mvarTableViews(3, lngCount) & ".ID" & vbNewLine
					End If
				Next lngCount

				strSQLSelect = strSQLSelect & "           INNER JOIN " & BankHolidayModule.gsBHolTableRealSource & vbNewLine _
						& "           ON  " & BankHolidayModule.gsBHolRegionTableName & ".ID = " & BankHolidayModule.gsBHolTableRealSource & ".ID_" & BankHolidayModule.glngBHolRegionTableID & vbNewLine

				If intBHolCount > 1 Then
					strSQLDateRegion = strSQLDateRegion & " OR " & vbNewLine
				End If

				'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				fFinalCareerChange = (intBHolCount = CShort(mavCareerRanges(4, intCount)))

				If fFinalCareerChange Then
					strSQLDateRegion = strSQLDateRegion & "((" & BankHolidayModule.gsBHolTableRealSource & "." & BankHolidayModule.gsBHolDateColumnName & " >= CONVERT(datetime, '" & VB6.Format(mavCareerRanges(1, intCount), "MM/dd/yyyy") & "')) " & vbNewLine _
							& " AND (" & BankHolidayModule.gsBHolTableRealSource & "." & BankHolidayModule.gsBHolDateColumnName & " >= '" & VB6.Format(mdtStartDate, "MM/dd/yyyy") & "') " & vbNewLine _
							& " AND (" & BankHolidayModule.gsBHolTableRealSource & "." & BankHolidayModule.gsBHolDateColumnName & " <= '" & VB6.Format(mdtEndDate, "MM/dd/yyyy") & "') " & vbNewLine _
							& " AND " & vbNewLine _
							& "(" & mstrSQLSelect_RegInfoRegion & " = '" & mavCareerRanges(3, intCount) & "') " & vbNewLine _
							& ")) " & vbNewLine
				Else
					strSQLDateRegion = strSQLDateRegion & "((" & BankHolidayModule.gsBHolTableRealSource & "." & BankHolidayModule.gsBHolDateColumnName & " >= CONVERT(datetime, '" & VB6.Format(mavCareerRanges(1, intCount), "MM/dd/yyyy") & "') " & vbNewLine & " AND (" & BankHolidayModule.gsBHolTableRealSource & "." & BankHolidayModule.gsBHolDateColumnName & " < CONVERT(datetime, '" & VB6.Format(mavCareerRanges(1, intCount + 1), "MM/dd/yyyy") & "'))) " & vbNewLine _
							& " AND (" & BankHolidayModule.gsBHolTableRealSource & "." & BankHolidayModule.gsBHolDateColumnName & " >= '" & VB6.Format(mdtStartDate, "MM/dd/yyyy") & "') " & vbNewLine _
							& " AND (" & BankHolidayModule.gsBHolTableRealSource & "." & BankHolidayModule.gsBHolDateColumnName & " <= '" & VB6.Format(mdtEndDate, "MM/dd/yyyy") & "') " & vbNewLine _
							& " AND (" & mstrSQLSelect_RegInfoRegion & " = '" & mavCareerRanges(3, intCount) & "')) " & vbNewLine
				End If

				If fFinalCareerChange Then
					strSQLAllBHols = strSQLAllBHols & strSQLSelect & vbNewLine _
							& strSQLWhere & vbNewLine _
							& strSQLDateRegion & vbNewLine _
							& " UNION ALL "
					strSQLWhere = vbNullString
					strSQLDateRegion = vbNullString
				End If

				'Send the query to SQL Server in batches of approximately 100, to avoid 256(260) Table/Views limit.
				'Do not split base records in to more than one batch!
				If ((lng100Counter = lngBaseRowCount) And fFinalCareerChange) Or ((lng100Counter > 100) And fFinalCareerChange) Or ((lngMainBaseCounter = lngBaseRowCount) And fFinalCareerChange) Then

					strSQLAllBHols = Left(strSQLAllBHols, Len(strSQLAllBHols) - 11)
					strSQLOrder = " ORDER BY 'ID', 'Region' " & vbNewLine
					strSQLAllBHols = strSQLAllBHols & strSQLOrder

					rsTempPersonnelBHols = DB.GetDataTable(strSQLAllBHols)
					rsPersonnelBHols.Merge(rsTempPersonnelBHols)
					'Accept changes.
					rsPersonnelBHols.AcceptChanges()

					lngBaseRecordID = -1
					blnNewBaseRecord = False
					intRecordBHol = 0

					'##############################################################################
					'populate collections with new data
					With rsTempPersonnelBHols

						INPUT_STRING = vbNullString

						If .Rows.Count > 0 Then

							For Each objRow As DataRow In .Rows

								' Append total bank holidays for this base record
								If Not lngBaseRecordID = -1 And lngBaseRecordID <> CInt(objRow("ID")) Then
									mstrBHolFormString.Append("<input NAME=txtBHolCOUNT_" & lngBaseRecordID & " ID=txtBHolCOUNT_" & lngBaseRecordID & " VALUE=""" & intRecordBHol & """>")
									intRecordBHol = 0
								End If

								If lngBaseRecordID <> CInt(objRow("ID")) Then

									If Not (colBankHolidays Is Nothing) Then
										mcolHistoricBankHolidays.Add(colBankHolidays, CStr(lngBaseRecordID))
										'UPGRADE_NOTE: Object colBankHolidays may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
										colBankHolidays = Nothing
									End If
									colBankHolidays = New clsBankHolidays

									lngBaseRecordID = CInt(objRow("ID"))

								End If

								'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
								colBankHolidays.Add(IIf(IsDBNull(objRow("Region")), "", objRow("Region")), IIf(IsDBNull(objRow(BankHolidayModule.gsBHolDescriptionColumnName)), "", objRow(BankHolidayModule.gsBHolDescriptionColumnName)), IIf(IsDBNull(objRow(BankHolidayModule.gsBHolDateColumnName)), "", objRow(BankHolidayModule.gsBHolDateColumnName)))

								intRecordBHol = intRecordBHol + 1

								'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
								INPUT_STRING = IIf(IsDBNull(objRow(BankHolidayModule.gsBHolDateColumnName)), "", VB6.Format(objRow(BankHolidayModule.gsBHolDateColumnName), mstrClientDateFormat)) & "_"
								'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
								INPUT_STRING = INPUT_STRING & IIf(IsDBNull(objRow("Region")), "", objRow("Region"))

								mstrBHolFormString.Append(String.Format("<input name=txtBHol_{0}_{1} id=txtBHol_{0}_{1} value=""{2}"">", lngBaseRecordID, intRecordBHol, Replace(INPUT_STRING, """", "&quot;")))

							Next

							mstrBHolFormString.Append("<input NAME=txtBHolCOUNT_" & lngBaseRecordID & " ID=txtBHolCOUNT_" & lngBaseRecordID & " VALUE=""" & intRecordBHol & """>")
							intRecordBHol = 0

							If Not colBankHolidays Is Nothing Then
								mcolHistoricBankHolidays.Add(colBankHolidays, CStr(lngBaseRecordID))
								'UPGRADE_NOTE: Object colBankHolidays may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
								colBankHolidays = Nothing
							End If


						Else

							Return True

						End If

					End With
					'##############################################################################

					'Reset SQL string variables ready for next batch to be created.
					strSQLAllBHols = vbNullString
					strSQLSelect = vbNullString
					strSQLWhere = vbNullString
					strSQLDateRegion = vbNullString
					lng100Counter = 0
					lngBaseRecordID = -1
				End If

				blnNewBaseRecord = False

			Next intCount

			mstrBHolFormString.Append("</FORM>")

		Catch ex As Exception
			Return False

		End Try

		Return True

	End Function

	Public Function Get_StaticBankHolidays() As Boolean

		If mblnDisableRegions Then
			Return False
		End If


		Dim colBankHolidays As clsBankHolidays
		Dim strSQLAllBHols As String
		Dim lngBaseRecordID As Integer
		Dim intBHolCount As Short
		Dim lngCount As Integer
		Dim INPUT_STRING As String

		Try

			strSQLAllBHols = "SELECT DISTINCT [Base].ID, [RegionInfo].Region, [RegionInfo].Holiday_Date, [RegionInfo].Description " & "FROM (SELECT DISTINCT " & mstrCalendarReportsBaseTableName & ".ID AS 'ID', " & mstrSQLSelect_PersonnelStaticRegion & " AS 'Region' " & vbNewLine

			If mlngRegion > 0 Then
				strSQLAllBHols = strSQLAllBHols & "      FROM " & mstrCalendarReportsBaseTableName & vbNewLine
			Else
				strSQLAllBHols = strSQLAllBHols & "      FROM " & PersonnelModule.gsPersonnelTableName & vbNewLine
			End If

			For lngCount = 0 To UBound(mvarTableViews, 2) Step 1
				'<PERSONNEL CODE>
				If mlngRegion > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(0, lngCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If mvarTableViews(0, lngCount) = mlngCalendarReportsBaseTable Then
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						strSQLAllBHols = strSQLAllBHols & "           LEFT OUTER JOIN " & mvarTableViews(3, lngCount) & vbNewLine _
								& "           ON  " & mstrCalendarReportsBaseTableName & ".ID = " & mvarTableViews(3, lngCount) & ".ID" & vbNewLine
					End If
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(0, lngCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If mvarTableViews(0, lngCount) = mlngCalendarReportsBaseTable Then
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						strSQLAllBHols = strSQLAllBHols & "           LEFT OUTER JOIN " & mvarTableViews(3, lngCount) & vbNewLine _
								& "           ON  " & PersonnelModule.gsPersonnelTableName & ".ID = " & mvarTableViews(3, lngCount) & ".ID" & vbNewLine
					End If
				End If
			Next lngCount

			If Len(Trim(mstrSQLIDs)) > 0 Then
				strSQLAllBHols = strSQLAllBHols & "      WHERE " & mstrCalendarReportsBaseTableName & ".ID IN (" & mstrSQLIDs & ") " & vbNewLine
			End If

			strSQLAllBHols = strSQLAllBHols & String.Format(" ) AS [Base] INNER JOIN (SELECT DISTINCT {0}.ID AS [ID], {1} AS [Region], {2}, {3} FROM {0}" _
				, BankHolidayModule.gsBHolRegionTableName, mstrSQLSelect_RegInfoRegion, mstrSQLSelect_BankHolDate, mstrSQLSelect_BankHolDesc)

			For lngCount = 0 To UBound(mvarTableViews, 2) Step 1
				'<REGIONAL CODE>
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(0, lngCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If mvarTableViews(0, lngCount) = BankHolidayModule.glngBHolRegionTableID Then
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strSQLAllBHols = strSQLAllBHols & "           LEFT OUTER JOIN " & mvarTableViews(3, lngCount) & " ON " & BankHolidayModule.gsBHolRegionTableName & ".ID = " & mvarTableViews(3, lngCount) & ".ID" & vbNewLine
				End If
			Next lngCount

			strSQLAllBHols = strSQLAllBHols & "           INNER JOIN " & BankHolidayModule.gsBHolTableRealSource & vbNewLine _
				& "           ON  " & BankHolidayModule.gsBHolRegionTableName & ".ID = " & BankHolidayModule.gsBHolTableRealSource & ".ID_" & BankHolidayModule.glngBHolRegionTableID & vbNewLine

			If Len(Trim(mstrSQLIDs)) > 0 Then
				strSQLAllBHols = strSQLAllBHols & "     WHERE (" & BankHolidayModule.gsBHolTableRealSource & "." & BankHolidayModule.gsBHolDateColumnName & " >= '" & VB6.Format(mdtStartDate, "MM/dd/yyyy") & "') " & vbNewLine _
						& "         AND (" & BankHolidayModule.gsBHolTableRealSource & "." & BankHolidayModule.gsBHolDateColumnName & " <= '" & VB6.Format(mdtEndDate, "MM/dd/yyyy") & "') " & vbNewLine
			Else
				strSQLAllBHols = strSQLAllBHols & "     WHERE (" & BankHolidayModule.gsBHolTableRealSource & "." & BankHolidayModule.gsBHolDateColumnName & " >= '" & VB6.Format(mdtStartDate, "MM/dd/yyyy") & "') " & vbNewLine _
						& "         AND (" & BankHolidayModule.gsBHolTableRealSource & "." & BankHolidayModule.gsBHolDateColumnName & " <= '" & VB6.Format(mdtEndDate, "MM/dd/yyyy") & "') " & vbNewLine
			End If

			strSQLAllBHols = strSQLAllBHols & "    ) AS [RegionInfo] ON [Base].Region = [RegionInfo].Region ORDER BY [Base].ID "
			rsPersonnelBHols = DB.GetDataTable(strSQLAllBHols)

			lngBaseRecordID = -1

			'##############################################################################
			'populate collections with new data
			With rsPersonnelBHols

				INPUT_STRING = vbNullString
				intBHolCount = 0

				mstrBHolFormString = New StringBuilder
				mstrBHolFormString.Append("<form id=frmBHol name=frmBHol style=""visibility:hidden;display:none"">")

				If .Rows.Count > 0 Then

					For Each objRow As DataRow In .Rows

						' Append total bank holidays for this base record
						If Not lngBaseRecordID = -1 And lngBaseRecordID <> CInt(objRow("ID")) Then
							mstrBHolFormString.Append("<input NAME=txtBHolCOUNT_" & lngBaseRecordID & " ID=txtBHolCOUNT_" & lngBaseRecordID & " VALUE=""" & intBHolCount & """>")
							intBHolCount = 0
						End If

						If lngBaseRecordID <> objRow("ID") Then
							intBHolCount = 0
							If Not (colBankHolidays Is Nothing) Then
								mcolStaticBankHolidays.Add(colBankHolidays, CStr(lngBaseRecordID))
								'UPGRADE_NOTE: Object colBankHolidays may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
								colBankHolidays = Nothing
							End If
							colBankHolidays = New clsBankHolidays
							lngBaseRecordID = objRow("ID")

						End If

						intBHolCount = intBHolCount + 1

						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						colBankHolidays.Add(IIf(IsDBNull(objRow("Region")), "", objRow("Region")), IIf(IsDBNull(objRow(BankHolidayModule.gsBHolDescriptionColumnName)), "", objRow(BankHolidayModule.gsBHolDescriptionColumnName)), IIf(IsDBNull(objRow(BankHolidayModule.gsBHolDateColumnName)), "", objRow(BankHolidayModule.gsBHolDateColumnName)))

						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						INPUT_STRING = IIf(IsDBNull(objRow(BankHolidayModule.gsBHolDateColumnName)), "", VB6.Format(objRow(BankHolidayModule.gsBHolDateColumnName), mstrClientDateFormat)) & "_"
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						INPUT_STRING = INPUT_STRING & IIf(IsDBNull(objRow("Region")), "", objRow("Region"))
						mstrBHolFormString.Append(String.Format("<input name=txtBHol_{0}_{1} id=txtBHol_{0}_{1} value=""{2}"">", objRow("ID"), intBHolCount, Replace(INPUT_STRING, """", "&quot;")))

					Next

					mstrBHolFormString.Append("<input NAME=txtBHolCOUNT_" & lngBaseRecordID & " ID=txtBHolCOUNT_" & lngBaseRecordID & " VALUE=""" & intBHolCount & """>")
					intBHolCount = 0

					If Not (colBankHolidays Is Nothing) Then
						mcolStaticBankHolidays.Add(colBankHolidays, CStr(lngBaseRecordID))
						'UPGRADE_NOTE: Object colBankHolidays may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
						colBankHolidays = Nothing
					End If

				Else
					mstrBHolFormString.Append("<input NAME=txtBHolCOUNT_" & lngBaseRecordID & " ID=txtBHolCOUNT_" & lngBaseRecordID & " VALUE=""" & intBHolCount & """>")

				End If

				mstrBHolFormString.Append("</form>")

			End With
			'##############################################################################

		Catch ex As Exception
			Return False

		End Try

		Return True

	End Function

	Public Function Get_StaticWorkingPatterns() As Boolean


		If mblnDisableWPs Then
			Return False
		End If

		Dim colWorkingPatterns As clsCalendarEvents

		Dim blnNewBaseRecord As Boolean
		Dim lngBaseRecordID As Integer

		Dim INPUT_STRING As String

		Try

			INPUT_STRING = vbNullString

			mstrWPFormString = "<FORM id=frmWP name=frmWP style=""visibility:hidden;display:none"">" & vbNewLine

			lngBaseRecordID = -1
			blnNewBaseRecord = False

			'##############################################################################
			'populate collections with new data
			With mrsCalendarBaseInfo

				If Not (.Rows.Count = 0) Then

					For Each objRow As DataRow In .Rows

						If lngBaseRecordID <> objRow(mstrBaseIDColumn) Then

							If Not (colWorkingPatterns Is Nothing) Then
								mstrWPFormString = mstrWPFormString & vbNewLine & vbTab & "<input NAME=txtWPCOUNT_" & lngBaseRecordID & " ID=txtWPCOUNT_" & lngBaseRecordID & " VALUE=1>" & vbNewLine

								mcolStaticWorkingPatterns.Add(colWorkingPatterns, CStr(lngBaseRecordID))
								'UPGRADE_NOTE: Object colWorkingPatterns may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
								colWorkingPatterns = Nothing
							End If
							colWorkingPatterns = New clsCalendarEvents
							colWorkingPatterns.SessionInfo = SessionInfo
							lngBaseRecordID = objRow(mstrBaseIDColumn)
							blnNewBaseRecord = True

						End If

						'lngBaseRecordID = .Fields(mstrBaseIDColumn).Value

						INPUT_STRING = vbNullString
						INPUT_STRING = INPUT_STRING & objRow(PersonnelModule.gsPersonnelWorkingPatternColumnName)

						mstrWPFormString = mstrWPFormString & vbTab & "<INPUT NAME=txtWP_" & lngBaseRecordID & " ID=txtBHol_" & lngBaseRecordID & " VALUE=""" & Replace(INPUT_STRING, """", "&quot;") & """>" & vbNewLine

						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						colWorkingPatterns.Add(CStr(colWorkingPatterns.Count), CStr(lngBaseRecordID), , , , , , , , , , , , , , , , , , , , , , , , , , , , , , IIf(IsDBNull(objRow(PersonnelModule.gsPersonnelWorkingPatternColumnName)), "              ", objRow(PersonnelModule.gsPersonnelWorkingPatternColumnName)))

						blnNewBaseRecord = False

					Next

					If Not colWorkingPatterns Is Nothing Then
						mstrWPFormString = mstrWPFormString & vbNewLine & vbTab & "<input NAME=txtWPCOUNT_" & lngBaseRecordID & " ID=txtWPCOUNT_" & lngBaseRecordID & " VALUE=1>" & vbNewLine
						mcolStaticWorkingPatterns.Add(colWorkingPatterns, CStr(lngBaseRecordID))
						'UPGRADE_NOTE: Object colWorkingPatterns may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
						colWorkingPatterns = Nothing
					End If

				Else
					Return True

				End If

				mstrWPFormString = mstrWPFormString & "</FORM>" & vbNewLine

			End With
			'##############################################################################

		Catch ex As Exception
			Return False

		End Try

		Return True

	End Function

	Private Function CheckPermission_RegionInfo() As Boolean

		Dim strTableColumn As String


		'Check the  Bank Holiday Region Table - Region Table
		'           Bank Holiday Region Table - Region Column
		'           Bank Holidays Table - Bank Holiday Table
		'           Bank Holidays Table - Date Column
		'           Bank Holidays Table - Descripiton Column
		'...Bank Holiday module setup information.
		'If any are blank then we need to allow the report to run, but disable the Bank Holiday Display Options.
		If BankHolidayModule.gsBHolRegionTableName = "" Or BankHolidayModule.gsBHolRegionColumnName = "" Or BankHolidayModule.gsBHolTableName = "" Or BankHolidayModule.gsBHolDateColumnName = "" Or BankHolidayModule.gsBHolDescriptionColumnName = "" Then

			GoTo DisableRegions
		End If

		'Check the  Career Change Region - Static Region Column
		'           Career Change Region - Historic Region Table
		'           Career Change Region - Historic Region Column
		'           Career Change Region - Historic Region Effective Date Column
		'...Personnel - Career Change module setup information.
		'If any are blank then we need to allow the report to run, but disable the Bank Holiday Display Options.
		If PersonnelModule.gsPersonnelRegionColumnName = "" Then
			If PersonnelModule.gsPersonnelHRegionTableName = "" Or PersonnelModule.gsPersonnelHRegionColumnName = "" Or PersonnelModule.gsPersonnelHRegionDateColumnName = "" Then

				GoTo DisableRegions
			End If
		End If




		'*******************************************************************
		' All Region module information is setup                           *
		' Now check the permissions on the Region module setup information *
		'*******************************************************************
		'Bank Holiday Region Table - Region Table (Regional Information)
		'Bank Holiday Region Table - Region Column
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
		If CheckPermission_Columns(BankHolidayModule.glngBHolRegionTableID, BankHolidayModule.gsBHolRegionTableName, BankHolidayModule.gsBHolRegionColumnName, strTableColumn) Then
			mstrSQLSelect_RegInfoRegion = strTableColumn
			strTableColumn = vbNullString
		Else
			GoTo DisableRegions
		End If
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\


		'Bank Holidays Table - Bank Holiday Table (Region History)
		'Bank Holidays Table - Date Column
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
		If CheckPermission_Columns(BankHolidayModule.glngBHolTableID, BankHolidayModule.gsBHolTableName, BankHolidayModule.gsBHolDateColumnName, strTableColumn) Then
			mstrSQLSelect_BankHolDate = strTableColumn
			strTableColumn = vbNullString
		Else
			GoTo DisableRegions
		End If
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\


		'Bank Holidays Table - Bank Holiday Table (Region History)
		'Bank Holidays Table - Descripiton Column
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
		If CheckPermission_Columns(BankHolidayModule.glngBHolTableID, BankHolidayModule.gsBHolTableName, BankHolidayModule.gsBHolDescriptionColumnName, strTableColumn) Then
			mstrSQLSelect_BankHolDesc = strTableColumn
			strTableColumn = vbNullString
		Else
			GoTo DisableRegions
		End If
		'///////////////////////////////////////////////
		'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\




		'*******************************************************************
		' Permission granted on all Region module information.             *
		' Now check the permissions on the                                 *
		' Personnel Career Change Region module setup information          *
		'*******************************************************************
		If mlngRegion > 0 Then
			'///////////////////////////////////////////////
			'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
			If CheckPermission_Columns(mlngCalendarReportsBaseTable, mstrCalendarReportsBaseTableName, mstrRegion, strTableColumn) Then
				mstrSQLSelect_PersonnelStaticRegion = strTableColumn
				strTableColumn = vbNullString
			Else
				GoTo DisableRegions
			End If
			'///////////////////////////////////////////////
			'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

		Else
			'Check Career Change Region access
			If PersonnelModule.gsPersonnelRegionColumnName <> "" Then
				'Personnel Table
				'Career Change Region - Static Region Column
				'///////////////////////////////////////////////
				'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
				If CheckPermission_Columns(PersonnelModule.glngPersonnelTableID, PersonnelModule.gsPersonnelTableName, PersonnelModule.gsPersonnelRegionColumnName, strTableColumn) Then
					mstrSQLSelect_PersonnelStaticRegion = strTableColumn
					strTableColumn = vbNullString
				Else
					GoTo DisableRegions
				End If
				'///////////////////////////////////////////////
				'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

			Else
				'Career Change Region - Historic Region Table
				'Career Change Region - Historic Region Column
				'///////////////////////////////////////////////
				'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
				If CheckPermission_Columns(PersonnelModule.glngPersonnelHRegionTableID, PersonnelModule.gsPersonnelHRegionTableName, PersonnelModule.gsPersonnelHRegionColumnName, strTableColumn) Then
					strTableColumn = vbNullString
				Else
					GoTo DisableRegions
				End If
				'///////////////////////////////////////////////
				'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

				'Career Change Region - Historic Region Table
				'Career Change Region - Historic Region Effective Date Column
				'///////////////////////////////////////////////
				'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
				If CheckPermission_Columns(PersonnelModule.glngPersonnelHRegionTableID, PersonnelModule.gsPersonnelHRegionTableName, PersonnelModule.gsPersonnelHRegionDateColumnName, strTableColumn) Then
					strTableColumn = vbNullString
				Else
					GoTo DisableRegions
				End If
				'///////////////////////////////////////////////
				'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

			End If
		End If

		CheckPermission_RegionInfo = True

TidyUpAndExit:
		Exit Function

DisableRegions:
		mblnDisableRegions = True
		CheckPermission_RegionInfo = False
		GoTo TidyUpAndExit

	End Function

	Private Function CheckPermission_Columns(plngTableID As Integer, pstrTableName As String, pstrColumnName As String, ByRef strSQLRef As String) As Boolean

		'This function checks if the current user has read(select) permissions
		'on this column. If the user only has access through views then the
		'relevent views are added to the mvarTableViews() array which in turn
		'are used to create the join part of the query.

		Dim lngTempTableID As Integer
		Dim strTempTableName As String
		Dim strTempColumnName As String
		Dim blnColumnOK As Boolean
		Dim blnFound As Boolean
		Dim blnNoSelect As Boolean
		Dim strSource As String
		Dim intNextIndex As Short
		Dim blnOK As Boolean
		Dim strTable As String
		Dim strColumn As String

		Dim pintNextIndex As Short

		' Set flags with their starting values
		blnOK = True
		blnNoSelect = False

		strTable = vbNullString
		strColumn = vbNullString

		' Load the temp variables
		lngTempTableID = plngTableID
		strTempTableName = pstrTableName
		strTempColumnName = pstrColumnName

		' Check permission on that column
		mobjColumnPrivileges = GetColumnPrivileges(strTempTableName)
		mstrRealSource = gcoTablePrivileges.Item(strTempTableName).RealSource

		blnColumnOK = mobjColumnPrivileges.IsValid(strTempColumnName)

		If blnColumnOK Then
			blnColumnOK = mobjColumnPrivileges.Item(strTempColumnName).AllowSelect
		End If

		If blnColumnOK Then
			' this column can be read direct from the tbl/view or from a parent table
			strTable = mstrRealSource
			strColumn = strTempColumnName

			'    ' If the table isnt the base table (or its realsource) then
			'    ' Check if it has already been added to the array. If not, add it.
			'    If lngTempTableID <> mlngCalendarReportsBaseTable Then
			blnFound = False
			For intNextIndex = 1 To UBound(mvarTableViews, 2)
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(2, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(1, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If mvarTableViews(1, intNextIndex) = 0 And mvarTableViews(2, intNextIndex) = lngTempTableID Then
					blnFound = True
					Exit For
				End If
			Next intNextIndex

			If Not blnFound Then
				intNextIndex = UBound(mvarTableViews, 2) + 1
				ReDim Preserve mvarTableViews(3, intNextIndex)
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(1, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarTableViews(1, intNextIndex) = 0
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(2, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarTableViews(2, intNextIndex) = lngTempTableID
			End If
			'    End If

			strSQLRef = strTable & "." & strColumn
		Else

			' this column cannot be read direct. If its from a parent, try parent views
			' Loop thru the views on the table, seeing if any have read permis for the column

			ReDim mstrViews(0)
			For Each mobjTableView In gcoTablePrivileges.Collection
				If (Not mobjTableView.IsTable) And (mobjTableView.TableID = lngTempTableID) And (mobjTableView.AllowSelect) Then

					strSource = mobjTableView.ViewName
					mstrRealSource = gcoTablePrivileges.Item(strSource).RealSource

					' Get the column permission for the view
					mobjColumnPrivileges = GetColumnPrivileges(strSource)

					' If we can see the column from this view
					If mobjColumnPrivileges.IsValid(strTempColumnName) Then
						If mobjColumnPrivileges.Item(strTempColumnName).AllowSelect Then

							ReDim Preserve mstrViews(UBound(mstrViews) + 1)
							mstrViews(UBound(mstrViews)) = mobjTableView.ViewName

							' Check if view has already been added to the array
							blnFound = False
							For intNextIndex = 0 To UBound(mvarTableViews, 2)
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(2, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(1, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If mvarTableViews(1, intNextIndex) = 1 And mvarTableViews(2, intNextIndex) = mobjTableView.ViewID Then
									blnFound = True
									Exit For
								End If
							Next intNextIndex

							If Not blnFound Then
								' View hasnt yet been added, so add it !
								intNextIndex = UBound(mvarTableViews, 2) + 1
								ReDim Preserve mvarTableViews(3, intNextIndex)
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(0, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								mvarTableViews(0, intNextIndex) = mobjTableView.TableID
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(1, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								mvarTableViews(1, intNextIndex) = 1
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(2, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								mvarTableViews(2, intNextIndex) = mobjTableView.ViewID
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(3, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								mvarTableViews(3, intNextIndex) = mobjTableView.ViewName
							End If

						End If
					End If
				End If

			Next mobjTableView
			'UPGRADE_NOTE: Object mobjTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			mobjTableView = Nothing

			' Does the user have select permission thru ANY views ?
			If UBound(mstrViews) = 0 Then
				blnNoSelect = True
			Else
				strSQLRef = ""
				For pintNextIndex = 1 To UBound(mstrViews)
					If pintNextIndex = 1 Then
						strSQLRef = "CASE"
					End If

					strSQLRef = strSQLRef & " WHEN NOT " & mstrViews(pintNextIndex) & "." & strTempColumnName & " IS NULL THEN " & mstrViews(pintNextIndex) & "." & strTempColumnName
				Next pintNextIndex

				If Len(strSQLRef) > 0 Then
					strSQLRef = strSQLRef & " ELSE NULL" & " END "
				End If

			End If

			' If we cant see a column, then get outta here
			If blnNoSelect Then
				strSQLRef = vbNullString
				CheckPermission_Columns = False
				Exit Function
			End If

			If Not blnOK Then
				strSQLRef = vbNullString
				CheckPermission_Columns = False
				Exit Function
			End If

		End If

		Return True

	End Function

	Private Function CheckPermission_WPInfo() As Boolean

		Dim objColumn As CColumnPrivileges
		Dim pblnColumnOK As Boolean
		Dim strTableColumn As String

		'Check the  Career Change Working Pattern - Static Working Pattern Column
		'           Career Change Working Pattern - Historic Working Pattern Table
		'           Career Change Working Pattern - Historic Working Pattern Column
		'           Career Change Working Pattern - Historic Working Pattern Effective Date Column
		'...Personnel - Career Change module setup information.
		'If any are blank then we need to allow the report to run, but disable the Working Dys Display Option.
		If PersonnelModule.gsPersonnelWorkingPatternColumnName = "" Then
			If PersonnelModule.gsPersonnelHWorkingPatternTableName = "" Or PersonnelModule.gsPersonnelHWorkingPatternColumnName = "" Or PersonnelModule.gsPersonnelHWorkingPatternDateColumnName = "" Then

				GoTo DisableWPs
			End If
		End If

		'****************************************************************************
		' All Working Pattern module information is setup                           *
		' Now check the permissions on the Working Pattern module setup information *
		'****************************************************************************

		'Check Career Change Working Pattern access
		If PersonnelModule.gsPersonnelWorkingPatternColumnName <> "" Then
			'Career Change Working Pattern - Static Working Pattern Column
			'///////////////////////////////////////////////
			'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
			If CheckPermission_Columns(PersonnelModule.glngPersonnelTableID, PersonnelModule.gsPersonnelTableName, PersonnelModule.gsPersonnelWorkingPatternColumnName, strTableColumn) Then
				strTableColumn = vbNullString
			Else
				GoTo DisableWPs
			End If
			'///////////////////////////////////////////////
			'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

		Else
			'Career Change Working Pattern - Historic Working Pattern Table
			objColumn = GetColumnPrivileges(PersonnelModule.gsPersonnelHWorkingPatternTableName)

			'Career Change Working Pattern - Historic Working Pattern Column
			pblnColumnOK = objColumn.IsValid(PersonnelModule.gsPersonnelHWorkingPatternColumnName)
			If pblnColumnOK Then
				pblnColumnOK = objColumn.Item(PersonnelModule.gsPersonnelHWorkingPatternColumnName).AllowSelect
			End If
			If pblnColumnOK = False Then
				GoTo DisableWPs
			End If

			'Career Change Working Pattern - Historic Working Pattern Effective Date Column
			pblnColumnOK = objColumn.IsValid(PersonnelModule.gsPersonnelHWorkingPatternDateColumnName)
			If pblnColumnOK Then
				pblnColumnOK = objColumn.Item(PersonnelModule.gsPersonnelHWorkingPatternDateColumnName).AllowSelect
			End If
			If pblnColumnOK = False Then
				GoTo DisableWPs
			End If

		End If

		CheckPermission_WPInfo = True

TidyUpAndExit:
		'UPGRADE_NOTE: Object objColumn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objColumn = Nothing
		Exit Function

DisableWPs:
		mblnDisableWPs = True
		IncludeWorkingDaysOnly = False
		CheckPermission_WPInfo = False
		GoTo TidyUpAndExit

	End Function

	Public Function GetEventsCollection() As Boolean

		Dim sSQL As String
		Dim rsTemp As DataTable

		Dim sTempTableName As String
		Dim sTempStartDateName As String
		Dim sTempStartSessionName As String
		Dim sTempEndDateName As String
		Dim sTempEndSessionName As String
		Dim sTempDurationName As String
		Dim sTempLegendTableName As String
		Dim sTempLegendColumnName As String
		Dim sTempLegendCodeName As String
		Dim sTempLegendEventTypeName As String
		Dim sTempDesc1Name As String
		Dim sTempDesc2Name As String

		Try

			' Get the column information from the Details table, in order
			sSQL = String.Format("SELECT e.*, t.TableName FROM ASRSysCalendarReportEvents e INNER JOIN ASRSysTables t ON t.tableID = e.TableID" &
					" WHERE CalendarReportID = {0} ORDER BY Name ASC", mlngCalendarReportID)
			rsTemp = DB.GetDataTable(sSQL)
			With rsTemp
				If .Rows.Count = 0 Then
					mstrErrorString = "No events found in the specified Calendar Report definition." & vbNewLine & "Please remove this definition and create a new one."
					Return False
				End If

				mcolEvents = New clsCalendarEvents
				mcolEvents.SessionInfo = SessionInfo

				For Each objRow As DataRow In rsTemp.Rows

					sTempTableName = objRow("TableName")

					If objRow("EventStartDateID") > 0 Then
						sTempStartDateName = GetColumnName(objRow("EventStartDateID"))
					Else
						Return False
					End If

					If objRow("EventStartSessionID") > 0 Then
						sTempStartSessionName = GetColumnName(objRow("EventStartSessionID"))
					Else
						sTempStartSessionName = vbNullString
					End If

					If objRow("EventEndDateID") > 0 Then
						sTempEndDateName = GetColumnName(objRow("EventEndDateID"))
					Else
						sTempEndDateName = vbNullString
					End If

					If objRow("EventEndSessionID") > 0 Then
						sTempEndSessionName = GetColumnName(objRow("EventEndSessionID"))
					Else
						sTempEndSessionName = vbNullString
					End If

					If objRow("EventDurationID") > 0 Then
						sTempDurationName = GetColumnName(objRow("EventDurationID"))
					Else
						sTempDurationName = vbNullString
					End If

					If objRow("LegendLookupTableID") > 0 Then
						sTempLegendTableName = GetTableName(objRow("LegendLookupTableID"))
					Else
						sTempLegendTableName = vbNullString
					End If

					If objRow("LegendLookupColumnID") > 0 Then
						sTempLegendColumnName = GetColumnName(objRow("LegendLookupColumnID"))
					Else
						sTempLegendColumnName = vbNullString
					End If

					If objRow("LegendLookupCodeID") > 0 Then
						sTempLegendCodeName = GetColumnName(objRow("LegendLookupCodeID"))
					Else
						sTempLegendCodeName = vbNullString
					End If

					If objRow("LegendEventColumnID") > 0 Then
						sTempLegendEventTypeName = GetColumnName(objRow("LegendEventColumnID"))
					Else
						sTempLegendEventTypeName = vbNullString
					End If

					If objRow("EventDesc1ColumnID") > 0 Then
						sTempDesc1Name = GetColumnName(objRow("EventDesc1ColumnID"))
					Else
						sTempDesc1Name = vbNullString
					End If

					If objRow("EventDesc2ColumnID") > 0 Then
						sTempDesc2Name = GetColumnName(objRow("EventDesc2ColumnID"))
					Else
						sTempDesc2Name = vbNullString
					End If

					mcolEvents.Add(objRow("EventKey"), objRow("Name"), objRow("TableID"), sTempTableName, objRow("FilterID"), objRow("EventStartDateID") _
												 , sTempStartDateName, objRow("EventStartSessionID"), sTempStartSessionName, objRow("EventEndDateID"), sTempEndDateName, objRow("EventEndSessionID") _
												 , sTempEndSessionName, objRow("EventDurationID"), sTempDurationName, objRow("LegendType"), objRow("LegendCharacter") _
												 , objRow("LegendLookupTableID"), sTempLegendTableName, objRow("LegendLookupColumnID"), sTempLegendColumnName, objRow("LegendLookupCodeID") _
												 , sTempLegendCodeName, objRow("LegendEventColumnID"), sTempLegendEventTypeName, objRow("EventDesc1ColumnID"), sTempDesc1Name, objRow("EventDesc2ColumnID") _
												 , sTempDesc2Name)

				Next
			End With

		Catch ex As Exception
			mstrErrorString = "Error whilst retrieving the event details recordsets'." & vbNewLine & ex.Message.RemoveSensitive()
			Return False

		End Try

		Return True

	End Function

	Public Function GenerateSQL() As Boolean

		Dim fOK As Boolean = True
		Dim objEvent As clsCalendarEvent
		Dim rsLegendBreakdown As DataTable
		Dim objLegendEvent As CalendarLegend

		Dim strSQL As String
		Dim strDynamicKey As String
		Dim strDynamicName As String

		mintDynamicEventCount = 0

		'loop through the events col and UNION the Event queries together
		For Each objEvent In mcolEvents.Collection

			mblnHasEventFilterIDs = False
			mstrEventFilterIDs = vbNullString

			With objEvent
				If (.LegendType = 1) And (.LegendTableID > 0) Then
					'Event is using a lookup table to find the calendar code for the event.
					'Therefore use the unique types from the legend information.

					strSQL = String.Format("SELECT DISTINCT {0}, {1} FROM {2}", .LegendColumnName, .LegendCodeName, .LegendTableName)
					rsLegendBreakdown = DB.GetDataTable(strSQL)

					If rsLegendBreakdown.Rows.Count = 0 Then
						mstrErrorString = "The '" & .LegendTableName & "' event lookup table contains no records."
						Return False
					End If

					For Each objRow As DataRow In rsLegendBreakdown.Rows

						mintDynamicEventCount = mintDynamicEventCount + 1

						strDynamicKey = "DYNAMICEVENT" & CStr(mintDynamicEventCount)
						strDynamicName = Replace(IIf(IsDBNull(objRow(0)), "", objRow(0)), "'", "''")
						mstrSQLDynamicLegendWhere = vbNullString

						objLegendEvent = New CalendarLegend
						objLegendEvent.LegendKey = strDynamicKey
						objLegendEvent.LegendDescription = IIf(IsDBNull(objRow(1)), "", objRow(1).ToString())

						Legend.Add(objLegendEvent)

						If fOK Then fOK = GenerateSQLEvent(objEvent.Key, strDynamicKey, strDynamicName)

						If Not fOK Then
							mblnNoRecords = True
							Return False
						End If

						fOK = InsertIntoTempTable(mstrSQLEvent)
						mstrSQLEvent = vbNullString

					Next

				Else

					objLegendEvent = New CalendarLegend
					objLegendEvent.LegendKey = objEvent.Key
					objLegendEvent.LegendDescription = objEvent.Name
					objLegendEvent.HexColor = "#f3f3f3"
					Legend.Add(objLegendEvent)

					If fOK Then fOK = GenerateSQLEvent(objEvent.Key, objEvent.Key, objEvent.Name)

					If Not fOK Then
						mblnNoRecords = True
						Return False
					End If

					If fOK Then fOK = InsertIntoTempTable(mstrSQLEvent)
					mstrSQLEvent = vbNullString

				End If
			End With

		Next objEvent

		Return fOK

	End Function

	Private Function GenerateSQLSelect(pstrEventKey As String, pstrDynamicKey As String, pstrDynamicName As String) As Boolean

		' Purpose : This function compiles the SQLSelect string looping
		'           thru the column details recordset.

		Dim objEvent As clsCalendarEvent

		Dim strColList As String
		Dim strBaseColList As String

		Dim strLegendSQL As String
		Dim strTableColumn As String
		Dim lngTempTableID As Integer
		Dim strTempTableName As String
		Dim strTempColumnName As String

		Dim strTempStartSession As String
		Dim strTempEndSession As String

		Try

			'Get the Base ID column values so that these can be used in the group by clause when checking
			'for multiple events in MultipleCheck().
			mstrSQLCreateTable = "[LegendName] nvarchar(255), [BaseID] [Integer] NOT NULL, "
			strColList = "'" & Replace(pstrDynamicName, "'", "''") & "' AS [LegendName], [" & mstrBaseTableRealSource & "].[ID] AS 'BaseID', " & vbNewLine

			If mlngDescription1 > 0 Then
				If CheckColumnPermissions(mlngCalendarReportsBaseTable, mstrCalendarReportsBaseTableName, mstrDescription1, strTableColumn) Then
					strColList = strColList & " CONVERT(varchar," & strTableColumn & ") AS 'Description1', " & vbNewLine
					strBaseColList = strBaseColList & " CONVERT(varchar," & strTableColumn & ") AS 'Description1', " & vbNewLine
					strTableColumn = vbNullString
				Else
					Return False
				End If
			Else
				strBaseColList = strBaseColList & "NULL AS 'Description1', " & vbNewLine
				strColList = strColList & "NULL AS 'Description1', " & vbNewLine
			End If
			mstrSQLCreateTable = mstrSQLCreateTable & "[Description1] [varchar] (MAX) NULL, "

			If mlngDescription2 > 0 Then
				If CheckColumnPermissions(mlngCalendarReportsBaseTable, mstrCalendarReportsBaseTableName, mstrDescription2, strTableColumn) Then
					strColList = strColList & "CONVERT(varchar," & strTableColumn & ") AS 'Description2', " & vbNewLine
					strBaseColList = strBaseColList & "CONVERT(varchar," & strTableColumn & ") AS 'Description2', " & vbNewLine
					strTableColumn = vbNullString
				Else
					Return False
				End If
			Else
				strBaseColList = strBaseColList & "NULL AS 'Description2', " & vbNewLine
				strColList = strColList & "NULL AS 'Description2', " & vbNewLine
			End If
			mstrSQLCreateTable = mstrSQLCreateTable & "[Description2] [varchar] (MAX) NULL, "

			If mlngDescriptionExpr > 0 Then
				If mblnDescCalcCode Then
					strColList = strColList & " " & mstrDescCalcCode & " AS 'DescriptionExpr', " & vbNewLine
					strBaseColList = strBaseColList & " " & mstrDescCalcCode & " AS 'DescriptionExpr', " & vbNewLine

				Else
					If CheckCalculationPermissions(mlngCalendarReportsBaseTable, mlngDescriptionExpr, strTableColumn) Then
						mstrDescCalcCode = strTableColumn
						mblnDescCalcCode = True
						strColList = strColList & " " & strTableColumn & " AS 'DescriptionExpr', " & vbNewLine
						strBaseColList = strBaseColList & " " & strTableColumn & " AS 'DescriptionExpr', " & vbNewLine
						strTableColumn = vbNullString
					Else
						Return False
					End If

				End If
			Else
				strBaseColList = strBaseColList & "NULL AS 'DescriptionExpr', " & vbNewLine
				strColList = strColList & "NULL AS 'DescriptionExpr', " & vbNewLine
			End If

			'need to set the type of the expression column for the CREAT TABLE...statement.
			Select Case mlngBaseDescriptionType
				Case ExpressionValueTypes.giEXPRVALUE_NUMERIC, ExpressionValueTypes.giEXPRVALUE_BYREF_NUMERIC
					mstrSQLCreateTable = mstrSQLCreateTable & "[DescriptionExpr] [float] NULL, "
				Case Else
					mstrSQLCreateTable = mstrSQLCreateTable & "[DescriptionExpr] [varchar] (MAX) NULL, "
			End Select


			objEvent = mcolEvents.Item(pstrEventKey)
			With objEvent
				If pstrDynamicKey <> vbNullString Then
					strColList = strColList & "'" & Replace(pstrDynamicKey, "'", "''") & "' AS '?ID_EventID', " & vbNewLine
				Else
					strColList = strColList & "'" & Replace(.Key, "'", "''") & "' AS '?ID_EventID', " & vbNewLine
				End If
				mstrSQLCreateTable = mstrSQLCreateTable & "[?ID_EventID] [varchar] (255) NULL, "

				If pstrDynamicName <> vbNullString Then
					strColList = strColList & "'" & Replace(pstrDynamicName, "'", "''") & "' AS 'Name', " & vbNewLine
				Else
					strColList = strColList & "'" & Replace(.Name, "'", "''") & "' AS 'Name', " & vbNewLine
				End If
				mstrSQLCreateTable = mstrSQLCreateTable & "[Name] [varchar] (255) NULL, "


				'****************************************************************************
				strColList = strColList & gcoTablePrivileges.Item(.TableName).RealSource & ".[ID] AS [ID], " & vbNewLine
				mstrSQLCreateTable = mstrSQLCreateTable & "[ID] [integer] NULL, "


				mlngEventViewColumn = .StartDateID
				If CheckColumnPermissions(.TableID, .TableName, .StartDateName, strTableColumn) Then
					strColList = strColList & strTableColumn & " AS 'StartDate', " & vbNewLine

					mstrSQLBaseStartDateColumn = strTableColumn
					strTableColumn = vbNullString
				Else
					GenerateSQLSelect = False
					Exit Function
				End If

				mstrSQLCreateTable = mstrSQLCreateTable & "[StartDate] [datetime] NULL, "

				If .StartSessionID > 0 Then
					If CheckColumnPermissions(.TableID, .TableName, .StartSessionName, strTableColumn) Then
						mstrSQLBaseStartSessionColumn = strTableColumn
						strTableColumn = vbNullString
					Else
						GenerateSQLSelect = False
						Exit Function
					End If
				End If

				If .EndDateID > 0 Then
					mlngEventViewColumn = .EndDateID
					If CheckColumnPermissions(.TableID, .TableName, .EndDateName, strTableColumn) Then
						mstrSQLBaseEndDateColumn = strTableColumn
						strTableColumn = vbNullString
					Else
						GenerateSQLSelect = False
						Exit Function
					End If
				End If

				If .EndSessionID > 0 Then
					If CheckColumnPermissions(.TableID, .TableName, .EndSessionName, strTableColumn) Then
						mstrSQLBaseEndSessionColumn = strTableColumn
						strTableColumn = vbNullString
					Else
						GenerateSQLSelect = False
						Exit Function
					End If
				End If

				If .DurationID > 0 Then
					mlngEventViewColumn = .DurationID
					If CheckColumnPermissions(.TableID, .TableName, .DurationName, strTableColumn) Then
						mstrSQLBaseDurationColumn = strTableColumn
						strTableColumn = vbNullString
					Else
						GenerateSQLSelect = False
						Exit Function
					End If

					If .StartSessionID > 0 Then
						strColList = strColList & mstrSQLBaseStartSessionColumn & " AS 'StartSession', " & vbNewLine
					Else
						strColList = strColList & "'AM' AS 'StartSession', " & vbNewLine
						mstrSQLBaseStartSessionColumn = "'AM'"
					End If

					strColList = strColList & "CASE " & vbNewLine _
							& "      WHEN  (RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '1' " & vbNewLine _
							& "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '2' " & vbNewLine _
							& "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '3' " & vbNewLine _
							& "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '4' " & vbNewLine _
							& "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '5') THEN " & vbNewLine _
							& "         DATEADD(dd " & vbNewLine _
							& "                 , CONVERT(integer,LEFT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),10,1))) " & vbNewLine _
							& "                           , CHARINDEX('.',LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),10,1))))- 1 )) " & vbNewLine _
							& "                 , " & mstrSQLBaseStartDateColumn & ") " & vbNewLine _
							& "      WHEN  (RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '6' " & vbNewLine _
							& "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '7' " & vbNewLine _
							& "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '8' " & vbNewLine _
							& "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '9') " & vbNewLine _
							& "          AND (" & mstrSQLBaseStartSessionColumn & " = 'AM') THEN " & vbNewLine

					strColList = strColList & "         DATEADD(dd " & vbNewLine _
							& "                 , CONVERT(integer,LEFT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),10,1))) " & vbNewLine _
							& "                         , CHARINDEX('.',LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),10,1))))- 1 )) " & vbNewLine _
							& "                 , " & mstrSQLBaseStartDateColumn & ") " & vbNewLine _
							& "      WHEN  (RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '6' " & vbNewLine _
							& "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '7' " & vbNewLine _
							& "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '8' " & vbNewLine _
							& "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '9') " & vbNewLine _
							& "          AND (" & mstrSQLBaseStartSessionColumn & " = 'PM') THEN " & vbNewLine _
							& "           DATEADD(dd " & vbNewLine _
							& "                 , CONVERT(integer,LEFT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),10,1))) " & vbNewLine _
							& "                         , CHARINDEX('.',LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),10,1))))- 1 ) + 1) " & vbNewLine _
							& "                 , " & mstrSQLBaseStartDateColumn & ") " & vbNewLine _
							& "      WHEN  (RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '0') " & vbNewLine _
							& "          AND (" & mstrSQLBaseStartSessionColumn & " = 'AM') THEN " & vbNewLine

					strColList = strColList & "           DATEADD(dd " & vbNewLine _
							& "                   , CONVERT(integer,LEFT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),10,1))) " & vbNewLine _
							& "                         , CHARINDEX('.',LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),10,1))))- 1 ) - 1) " & vbNewLine _
							& "                   , " & mstrSQLBaseStartDateColumn & ") " & vbNewLine _
							& "      WHEN  (RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '0') " & vbNewLine _
							& "          AND (" & mstrSQLBaseStartSessionColumn & " = 'PM') THEN " & vbNewLine _
							& "           DATEADD(dd " & vbNewLine _
							& "                 , CONVERT(integer,LEFT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),10,1))) " & vbNewLine _
							& "                               , CHARINDEX('.',LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),10,1))))- 1 )) " & vbNewLine _
							& "                 , " & mstrSQLBaseStartDateColumn & ") " & vbNewLine _
							& "END AS 'EndDate', " & vbNewLine

					If .EndSessionID > 0 Then
						strColList = strColList & mstrSQLBaseEndSessionColumn & " AS 'EndSession', " & vbNewLine
					Else

						strColList = strColList & "CASE" _
								& "      WHEN  (RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '1' " & vbNewLine _
								& "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '2' " & vbNewLine _
								& "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '3' " & vbNewLine _
								& "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '4' " & vbNewLine _
								& "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '5') THEN " & vbNewLine _
								& "           " & mstrSQLBaseStartSessionColumn & " " & vbNewLine _
								& "      WHEN  (RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '6' " & vbNewLine _
								& "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '7' " & vbNewLine _
								& "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '8' " & vbNewLine _
								& "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '9' " & vbNewLine _
								& "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '0') " & vbNewLine _
								& "          AND (" & mstrSQLBaseStartSessionColumn & " = 'AM') THEN " & vbNewLine _
								& "           'PM'  " & vbNewLine

						strColList = strColList & "      WHEN  (RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '6' " & vbNewLine _
								& "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '7' " & vbNewLine _
								& "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '8' " & vbNewLine _
								& "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '9' " & vbNewLine _
								& "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '0') " & vbNewLine _
								& "          AND (" & mstrSQLBaseStartSessionColumn & " = 'PM') THEN " & vbNewLine _
								& "           'AM' " & vbNewLine _
								& "END AS 'EndSession', " & vbNewLine
					End If

					strColList = strColList & mstrSQLBaseDurationColumn & " AS 'Duration', " & vbNewLine

				ElseIf .EndDateID > 0 Then

					If .StartSessionID > 0 Then
						strColList = strColList & mstrSQLBaseStartSessionColumn & " AS 'StartSession', " & vbNewLine
						strTempStartSession = mstrSQLBaseStartSessionColumn
					Else
						strColList = strColList & "'AM' AS 'StartSession'," & vbNewLine
						strTempStartSession = "'AM'"
					End If

					strColList = strColList & mstrSQLBaseEndDateColumn & " AS 'EndDate', " & vbNewLine

					If .EndSessionID > 0 Then
						strColList = strColList & mstrSQLBaseEndSessionColumn & " AS 'EndSession', " & vbNewLine
						strTempEndSession = mstrSQLBaseEndSessionColumn
					Else
						strColList = strColList & "'PM' AS 'EndSession'," & vbNewLine
						strTempEndSession = "'PM'"
					End If

					strColList = strColList & " CASE " & vbNewLine _
							& " WHEN " & strTempStartSession & " = " & strTempEndSession & vbNewLine _
							& "   THEN CONVERT(float,(DATEDIFF(dd, " & mstrSQLBaseStartDateColumn & ", " & mstrSQLBaseEndDateColumn & ") + 0.5)) " & vbNewLine _
							& " ELSE " & vbNewLine _
							& "   CONVERT(float,(DATEDIFF(dd, " & mstrSQLBaseStartDateColumn & ", " & mstrSQLBaseEndDateColumn & ") + 1)) " & vbNewLine _
							& " END AS 'Duration'," & vbNewLine
				Else

					If .StartSessionID > 0 Then
						strColList = strColList & mstrSQLBaseStartSessionColumn & " AS 'StartSession', " & vbNewLine
					Else
						strColList = strColList & "'AM' AS 'StartSession'," & vbNewLine
					End If

					strColList = strColList & mstrSQLBaseStartDateColumn & " AS 'EndDate', " & vbNewLine

					If .StartSessionID > 0 Then
						strColList = strColList & mstrSQLBaseStartSessionColumn & " AS 'EndSession'," & vbNewLine _
								& " 0.5 AS 'Duration', " & vbNewLine
					Else
						strColList = strColList & "'PM' AS 'EndSession'," & vbNewLine _
								& " 1 AS 'Duration', " & vbNewLine
					End If

				End If
				mstrSQLCreateTable = mstrSQLCreateTable & "[StartSession] [varchar] (255) NULL, " _
						& "[EndDate] [datetime] NULL, " _
						& "[EndSession] [varchar] (255) NULL, " _
						& "[Duration] [float] NULL, "
				'****************************************************************************

				If .Description1ID > 0 Then
					lngTempTableID = .Description1_TableID
					strTempTableName = .Description1_TableName
					strTempColumnName = .Description1_ColumnName

					If CheckColumnPermissions(lngTempTableID, strTempTableName, strTempColumnName, strTableColumn) Then
						strColList = strColList & .Description1ID & " AS 'EventDescription1ColumnID', " & vbNewLine _
								& "'" & .Description1Name & "' AS 'EventDescription1Column', " & vbNewLine

						'TM20030407 Fault 5259 - if logic field...convert to 'Y' or 'N' accordingly.
						If GetDataType(lngTempTableID, .Description1ID) = ColumnDataType.sqlBoolean Then
							strColList = strColList & "CASE " & strTableColumn & " WHEN 1 THEN 'Y' ELSE 'N' END AS 'EventDescription1', " & vbNewLine
						Else
							strColList = strColList & "CONVERT(varchar(MAX)," & strTableColumn & ") AS 'EventDescription1', " & vbNewLine
						End If

						strTableColumn = vbNullString
					Else
						GenerateSQLSelect = False
						Exit Function
					End If
				Else
					strColList = strColList & "NULL AS 'EventDescription1ColumnID', " & vbNewLine _
							& "NULL AS 'EventDescription1Column', " & vbNewLine _
							& "NULL AS 'EventDescription1', " & vbNewLine
				End If
				mstrSQLCreateTable = mstrSQLCreateTable & "[EventDescription1ColumnID] [int] NULL, " _
						& "[EventDescription1Column] [varchar] (MAX) NULL, " _
						& "[EventDescription1] [varchar] (MAX) NULL, "

				If .Description2ID > 0 Then
					lngTempTableID = .Description2_TableID
					strTempTableName = .Description2_TableName
					strTempColumnName = .Description2_ColumnName

					If CheckColumnPermissions(lngTempTableID, strTempTableName, strTempColumnName, strTableColumn) Then
						strColList = strColList & .Description2ID & " AS 'EventDescription2ColumnID', " & vbNewLine _
								& "'" & .Description2Name & "' AS 'EventDescription2Column', " & vbNewLine

						'TM20030407 Fault 5259 - if logic field...convert to 'Y' or 'N' accordingly.
						If GetDataType(lngTempTableID, .Description2ID) = ColumnDataType.sqlBoolean Then
							strColList = strColList & "CASE " & strTableColumn & " WHEN 1 THEN 'Y' ELSE 'N' END AS 'EventDescription2', " & vbNewLine
						Else
							strColList = strColList & "CONVERT(varchar(MAX)," & strTableColumn & ") AS 'EventDescription2', " & vbNewLine
						End If

						strTableColumn = vbNullString
					Else
						GenerateSQLSelect = False
						Exit Function
					End If
				Else
					strColList = strColList & "NULL AS 'EventDescription2ColumnID', " & vbNewLine _
							& "NULL AS 'EventDescription2Column', " & vbNewLine _
							& "NULL AS 'EventDescription2', " & vbNewLine
				End If
				mstrSQLCreateTable = mstrSQLCreateTable & "[EventDescription2ColumnID] [int] NULL, " _
						 & "[EventDescription2Column] [varchar] (MAX) NULL, " _
						 & "[EventDescription2] [varchar] (MAX) NULL, "

				If .LegendType = 1 Then
					If CheckColumnPermissions(.LegendTableID, .LegendTableName, .LegendCodeName, strTableColumn) Then
						strLegendSQL = "LEFT((SELECT TOP 1 " & strTableColumn
						strTableColumn = vbNullString
					Else
						GenerateSQLSelect = False
						Exit Function
					End If

					strLegendSQL = strLegendSQL & " FROM " & .LegendTableName

					If CheckColumnPermissions(.LegendTableID, .LegendTableName, .LegendColumnName, strTableColumn) Then
						strLegendSQL = strLegendSQL & " WHERE " & strTableColumn
						strTableColumn = vbNullString
					Else
						GenerateSQLSelect = False
						Exit Function
					End If

					If CheckColumnPermissions(.TableID, .TableName, .LegendEventTypeName, strTableColumn) Then
						strLegendSQL = strLegendSQL & " = " & strTableColumn & "),2) AS 'Legend', "
						mstrSQLDynamicLegendWhere = strTableColumn & " = '" & pstrDynamicName & "' "
						strTableColumn = vbNullString
					Else
						GenerateSQLSelect = False
						Exit Function
					End If

				Else
					strLegendSQL = "'" & Replace(Left(.LegendCharacter, 2), "'", "''") & "' AS 'Legend', "

				End If

				strColList = strColList & strLegendSQL & vbNewLine
				mstrSQLCreateTable = mstrSQLCreateTable & "[Legend] [varchar] (MAX) NULL, "
			End With

			'Add the static region column if required.
			If mlngRegion > 0 Then

				If CheckColumnPermissions(mlngCalendarReportsBaseTable, mstrCalendarReportsBaseTableName, mstrRegion, strTableColumn) Then
					strColList = strColList & "CONVERT(varchar," & strTableColumn & ") AS 'Region', " & vbNewLine
					strBaseColList = strBaseColList & "CONVERT(varchar," & strTableColumn & ") AS 'Region', " & vbNewLine
					strTableColumn = vbNullString
				Else
					strColList = strColList & "NULL AS 'Region', "
					strBaseColList = strBaseColList & "NULL AS 'Region', "
				End If

			Else
				strColList = strColList & "NULL AS 'Region', "
				strBaseColList = strBaseColList & "NULL AS 'Region', "
			End If
			mstrSQLCreateTable = mstrSQLCreateTable & "[Region] [varchar] (MAX) NULL, "

			If CheckColumnPermissions(mlngCalendarReportsBaseTable, mstrCalendarReportsBaseTableName, "ID", strTableColumn) Then
				strColList = strColList & strTableColumn & " AS '?ID_" & mstrCalendarReportsBaseTableName & "', " & vbNewLine
				strBaseColList = strBaseColList & strTableColumn & " AS '?ID_" & mstrCalendarReportsBaseTableName & "', " & vbNewLine
				strTableColumn = vbNullString
			Else
				Return False
			End If
			mstrSQLCreateTable = mstrSQLCreateTable & "[?ID_" & mstrCalendarReportsBaseTableName & "] [varchar] (255) NULL, "

			'Add the static Working Pattern column if required.
			If (mlngCalendarReportsBaseTable = PersonnelModule.glngPersonnelTableID) And (PersonnelModule.gwptWorkingPatternType = WorkingPatternType.wptStaticWPattern) And (Not mblnGroupByDescription) Then
				If CheckColumnPermissions(mlngCalendarReportsBaseTable, mstrCalendarReportsBaseTableName, PersonnelModule.gsPersonnelWorkingPatternColumnName, strTableColumn) Then
					strColList = strColList & "CONVERT(varchar," & strTableColumn & ") AS 'Working_Pattern', " & vbNewLine
					strBaseColList = strBaseColList & "CONVERT(varchar," & strTableColumn & ") AS 'Working_Pattern', " & vbNewLine
					strTableColumn = vbNullString
				Else
					strColList = strColList & "NULL AS 'Working_Pattern', "
					strBaseColList = strBaseColList & "NULL AS 'Working_Pattern', "
				End If
			Else
				strColList = strColList & "NULL AS 'Working_Pattern', "
				strBaseColList = strBaseColList & "NULL AS 'Working_Pattern', "
			End If
			mstrSQLCreateTable = mstrSQLCreateTable & "[Working_Pattern] [varchar] (255) NULL, "

			Dim intOrderCount As Short
			Dim strOrderColumn As String

			mstrSQLOrderList = vbNullString

			For intOrderCount = 1 To UBound(mvarSortOrder, 2) Step 1
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(1, intOrderCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strOrderColumn = mvarSortOrder(1, intOrderCount)
				strTableColumn = vbNullString
				If CheckColumnPermissions(mlngCalendarReportsBaseTable, mstrCalendarReportsBaseTableName, strOrderColumn, strTableColumn) Then
					strColList = strColList & strTableColumn & " AS 'ORDER_" & CStr(intOrderCount) & "',"
					strBaseColList = strBaseColList & strTableColumn & " AS 'ORDER_" & CStr(intOrderCount) & "',"
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mstrSQLOrderList = mstrSQLOrderList & strTableColumn & " " & mvarSortOrder(2, intOrderCount) & ","
					strTableColumn = vbNullString
					If Not mblnOrderByCreated Then
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(2, intOrderCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mstrSQLOrderBy = mstrSQLOrderBy & " [ORDER_" & CStr(intOrderCount) & "] " & mvarSortOrder(2, intOrderCount) & ","
					End If
				Else
					Return False
				End If
				mstrSQLCreateTable = mstrSQLCreateTable & "[ORDER_" & CStr(intOrderCount) & "] [varchar] (MAX) NULL,"
			Next intOrderCount

			If Not mblnOrderByCreated Then
				mstrSQLOrderBy = Left(mstrSQLOrderBy, Len(mstrSQLOrderBy) - 1)
				mblnOrderByCreated = True
			End If
			mstrSQLCreateTable = Left(mstrSQLCreateTable, Len(mstrSQLCreateTable) - 1)

			strColList = Left(strColList, Len(strColList) - 1)
			strBaseColList = Left(strBaseColList, Len(strBaseColList) - 1)
			mstrSQLOrderList = Left(mstrSQLOrderList, Len(mstrSQLOrderList) - 1)

			' Start off the select statement
			mstrSQLSelect = "SELECT " & strColList
			mstrSQLBaseData = "SELECT " & strBaseColList

		Catch ex As Exception
			mstrErrorString = "Error whilst generating SQL Select statement." & vbNewLine & ex.Message.RemoveSensitive()
			Return False

		End Try

		Return True

	End Function

	Private Function CheckCalculationPermissions(plngBaseTableID As Integer, plngExprID As Integer, ByRef strSQLRef As String) As Boolean

		'This function checks if the current user has read(select) permissions
		'on this calculation. If the user only has access through views then the
		'relevent views are added to the mlngTableViews() array which in turn
		'are used to create the join part of the query.

		Dim blnFound As Boolean
		Dim iLoop1 As Short
		Dim intNextIndex As Short
		Dim blnOK As Boolean
		Dim sCalcCode As String
		Dim alngSourceTables(,) As Integer
		Dim objCalcExpr As clsExprExpression

		' Set flags with their starting values
		blnOK = True

		' Get the calculation SQL, and the array of tables/views that are used to create it.
		' Column 1 = 0 if this row is for a table, 1 if it is for a view.
		' Column 2 = table/view ID.
		ReDim alngSourceTables(2, 0)
		objCalcExpr = NewExpression()
		blnOK = objCalcExpr.Initialise(plngBaseTableID, plngExprID, ExpressionTypes.giEXPR_RUNTIMECALCULATION, ExpressionValueTypes.giEXPRVALUE_UNDEFINED)
		If blnOK Then
			blnOK = objCalcExpr.RuntimeCalculationCode(alngSourceTables, sCalcCode, mastrUDFsRequired, True, False, mvarPrompts)
		End If

		'The "SELECT ... INTO..." statement errors when it trys to create a column for
		'and empty string. Therefore wrap this empty sting in a CONVERT(varchar... clause if an sql empty string
		'is returned.
		'TM20030521 Fault 5702 - Compare the empty string with the calc code value converted to varchar
		sCalcCode = "CASE WHEN CONVERT(varchar," & sCalcCode & ") = '' " & "THEN CONVERT(varchar," & sCalcCode & ") " & "ELSE " & sCalcCode & " END"

		If blnOK Then
			strSQLRef = sCalcCode

			' Add the required views to the JOIN code.
			For iLoop1 = 1 To UBound(alngSourceTables, 2)
				If alngSourceTables(1, iLoop1) = 1 Then
					' Check if view has already been added to the array
					blnFound = False
					For intNextIndex = 1 To UBound(mlngTableViews, 2)
						If mlngTableViews(1, intNextIndex) = 1 And mlngTableViews(2, intNextIndex) = alngSourceTables(2, iLoop1) Then
							blnFound = True
							Exit For
						End If
					Next intNextIndex

					If Not blnFound Then

						' View hasnt yet been added, so add it !
						intNextIndex = UBound(mlngTableViews, 2) + 1
						ReDim Preserve mlngTableViews(2, intNextIndex)
						mlngTableViews(1, intNextIndex) = 1
						mlngTableViews(2, intNextIndex) = alngSourceTables(2, iLoop1)

					End If
					'********************************************************************************
				ElseIf alngSourceTables(1, iLoop1) = 0 Then
					' Check if table has already been added to the array
					blnFound = False
					For intNextIndex = 1 To UBound(mlngTableViews, 2)
						If mlngTableViews(1, intNextIndex) = 0 And mlngTableViews(2, intNextIndex) = alngSourceTables(2, iLoop1) Then
							blnFound = True
							Exit For
						End If
					Next intNextIndex

					'Only want to check if the source table is the base table
					' if we have NOT just found the source table in the array of joined tables.
					If Not blnFound Then
						blnFound = (alngSourceTables(2, iLoop1) = mlngCalendarReportsBaseTable)
					End If

					If Not blnFound Then
						' table hasnt yet been added, so add it !
						intNextIndex = UBound(mlngTableViews, 2) + 1
						ReDim Preserve mlngTableViews(2, intNextIndex)
						mlngTableViews(1, intNextIndex) = 0
						mlngTableViews(2, intNextIndex) = alngSourceTables(2, iLoop1)
					End If
					'********************************************************************************
				End If
			Next iLoop1
		Else
			' Permission denied on something in the calculation.
			mstrErrorString = "You do not have permission to use the '" & objCalcExpr.Name & "' calculation."
			CheckCalculationPermissions = False
			Exit Function
		End If

		'UPGRADE_NOTE: Object objCalcExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objCalcExpr = Nothing

		CheckCalculationPermissions = True

	End Function

	Private Function GenerateSQLFrom() As Boolean

		mstrSQLFrom = "FROM " & mstrBaseTableRealSource & vbNewLine
		mstrSQLBaseData = mstrSQLBaseData & " FROM " & mstrBaseTableRealSource & vbNewLine

		Return True

	End Function

	Private Function GenerateSQLJoin(pstrEventKey As String) As Boolean

		Dim objTableView As TablePrivilege
		Dim objChildTable As TablePrivilege
		Dim objEvent As clsCalendarEvent

		Dim sChildJoinCode As String
		Dim strFilterIDs As String

		Dim blnOK As Boolean

		Dim intLoop As Short

		Dim bViewContains_StartColumn As Boolean
		Dim bViewContains_EndColumn As Boolean
		Dim bViewContains_DurationColumn As Boolean

		' First, do the join for all the views etc...

		Try

			objEvent = mcolEvents.Item(pstrEventKey)

			For intLoop = 1 To UBound(mlngTableViews, 2)

				If mlngTableViews(1, intLoop) = 0 Then
					objTableView = gcoTablePrivileges.FindTableID(mlngTableViews(2, intLoop))
					bViewContains_StartColumn = True
					bViewContains_EndColumn = True
					bViewContains_DurationColumn = True
				Else
					objTableView = gcoTablePrivileges.FindViewID(mlngTableViews(2, intLoop))
					bViewContains_StartColumn = IsColumnInView((objTableView.ViewID), (objEvent.StartDateID))
					bViewContains_EndColumn = IsColumnInView((objTableView.ViewID), (objEvent.EndDateID))
					bViewContains_DurationColumn = IsColumnInView((objTableView.ViewID), (objEvent.DurationID))
				End If

				If (objTableView.TableID = mlngCalendarReportsBaseTable) And (objTableView.ViewID > 0) Or (IsAParentOf((objTableView.TableID), mlngCalendarReportsBaseTable)) Then
					' Get the table/view object from the id stored in the array

					If (IsAParentOf((objTableView.TableID), mlngCalendarReportsBaseTable)) Then
						mstrSQLJoin = mstrSQLJoin & " LEFT OUTER JOIN " & objTableView.RealSource & " ON " & mstrBaseTableRealSource & ".ID_" & objTableView.TableID & " = " & objTableView.RealSource & ".ID"

						mstrSQLBaseData = mstrSQLBaseData & " LEFT OUTER JOIN " & objTableView.RealSource & " ON " & mstrBaseTableRealSource & ".ID_" & objTableView.TableID & " = " & objTableView.RealSource & ".ID"

					Else
						mstrSQLJoin = mstrSQLJoin & " LEFT OUTER JOIN " & objTableView.RealSource & " ON " & mstrBaseTableRealSource & ".ID = " & objTableView.RealSource & ".ID"

						mstrSQLBaseData = mstrSQLBaseData & " LEFT OUTER JOIN " & objTableView.RealSource & " ON " & mstrBaseTableRealSource & ".ID = " & objTableView.RealSource & ".ID"
					End If

					If (objTableView.TableID = objEvent.TableID) Then
						'add clause to SQL, so that only dates within the specified range are retrieved.
						If (objEvent.StartDateID > 0) And (objEvent.EndDateID > 0) And bViewContains_StartColumn And bViewContains_EndColumn Then
							'event is defined by start date and end date
							mstrSQLJoin = mstrSQLJoin & " AND ((" & objTableView.RealSource & "." & objEvent.StartDateName & " <= convert(datetime, '" & VB6.Format(mdtStartDate, "MM/dd/yyyy") & "') AND " & objTableView.RealSource & "." & objEvent.EndDateName & " >= convert(datetime, '" & VB6.Format(mdtStartDate, "MM/dd/yyyy") & "'))" & vbNewLine & " OR (" & objTableView.RealSource & "." & objEvent.StartDateName & " >= convert(datetime, '" & VB6.Format(mdtStartDate, "MM/dd/yyyy") & "') AND " & objTableView.RealSource & "." & objEvent.EndDateName & " <= convert(datetime, '" & VB6.Format(mdtEndDate, "MM/dd/yyyy") & "'))" & vbNewLine & " OR (((" & objTableView.RealSource & "." & objEvent.StartDateName & " >= convert(datetime, '" & VB6.Format(mdtStartDate, "MM/dd/yyyy") & "')) AND (" & objTableView.RealSource & "." & objEvent.StartDateName & " <= convert(datetime, '" & VB6.Format(mdtEndDate, "MM/dd/yyyy") & "'))) AND " & objTableView.RealSource & "." & objEvent.EndDateName & " >= convert(datetime, '" & VB6.Format(mdtEndDate, "MM/dd/yyyy") & "'))" & vbNewLine & " OR (" & objTableView.RealSource & "." & objEvent.StartDateName & " <= convert(datetime, '" & VB6.Format(mdtStartDate, "MM/dd/yyyy") & "') AND " & objTableView.RealSource & "." & objEvent.EndDateName & " >= convert(datetime, '" & VB6.Format(mdtStartDate, "MM/dd/yyyy") & "')))" & vbNewLine _
									& " AND (" & objTableView.RealSource & "." & objEvent.EndDateName & " >= " & objTableView.RealSource & "." & objEvent.StartDateName & ")" & vbNewLine

						ElseIf (objEvent.StartDateID > 0) And (objEvent.DurationID > 0) And bViewContains_StartColumn And bViewContains_DurationColumn Then
							'event is defined by start date and duration
							mstrSQLJoin = mstrSQLJoin & " OR (" & objTableView.RealSource & "." & objEvent.StartDateName & " IS NOT NULL AND " & objTableView.RealSource & "." & objEvent.DurationName & " > 0)" & vbNewLine

						ElseIf (objEvent.StartDateID) > 0 And (objEvent.EndDateID < 1) And (objEvent.DurationID < 1) And bViewContains_StartColumn Then
							'event is defined by just the start date - one off event with a range of one
							mstrSQLJoin = mstrSQLJoin & " AND ((" & objTableView.RealSource & "." & objEvent.StartDateName & " >= convert(datetime, '" & VB6.Format(mdtStartDate, "MM/dd/yyyy") & "') AND " & objTableView.RealSource & "." & objEvent.StartDateName & " <= convert(datetime, '" & VB6.Format(mdtEndDate, "MM/dd/yyyy") & "'))" & vbNewLine
							mstrSQLJoin = mstrSQLJoin & ")" & vbNewLine

						End If
					End If

				ElseIf (IsAChildOf(mlngTableViews(2, intLoop), mlngCalendarReportsBaseTable)) And (objEvent.TableID = objTableView.TableID) Then
					objChildTable = gcoTablePrivileges.FindTableID(mlngTableViews(2, intLoop))

					If objChildTable.AllowSelect Then
						sChildJoinCode = sChildJoinCode & " INNER JOIN " & objChildTable.RealSource & " ON " & mstrBaseTableRealSource & ".ID = " & objChildTable.RealSource & ".ID_" & mlngCalendarReportsBaseTable

						If (objEvent.FilterID > 0) Then

							'TM20030407 Fault 5257 - only get the filter string once for each event to avoid being prompted
							'more tahn once for the save event if the event is split into dynamic events.
							If mblnHasEventFilterIDs Then
								blnOK = True
							Else
								blnOK = FilteredIDs((objEvent.FilterID), strFilterIDs, mastrUDFsRequired, mvarPrompts)
								mblnHasEventFilterIDs = blnOK
								mstrEventFilterIDs = strFilterIDs
							End If

							If blnOK Then
								sChildJoinCode = sChildJoinCode & " AND " & objChildTable.RealSource & ".ID IN (" & mstrEventFilterIDs & ")"
							Else
								' Permission denied on something in the filter.
								mstrErrorString = "You do not have permission to use the '" & General.GetFilterName(objEvent.FilterID) & "' filter."
								Return False
							End If
						End If

						'add clause to SQL, so that only dates within the specified range are retrieved.
						If (objEvent.StartDateID > 0 And objEvent.EndDateID > 0) Then
							'event is defined by start date and end date
							sChildJoinCode = sChildJoinCode & " AND ((" & objChildTable.RealSource & "." & objEvent.StartDateName & " <= convert(datetime, '" & VB6.Format(mdtStartDate, "MM/dd/yyyy") & "') AND " & objChildTable.RealSource & "." & objEvent.EndDateName & " >= convert(datetime, '" & VB6.Format(mdtStartDate, "MM/dd/yyyy") & "'))" & vbNewLine & " OR (" & objChildTable.RealSource & "." & objEvent.StartDateName & " >= convert(datetime, '" & VB6.Format(mdtStartDate, "MM/dd/yyyy") & "') AND " & objChildTable.RealSource & "." & objEvent.EndDateName & " <= convert(datetime, '" & VB6.Format(mdtEndDate, "MM/dd/yyyy") & "'))" & vbNewLine & " OR (((" & objChildTable.RealSource & "." & objEvent.StartDateName & " >= convert(datetime, '" & VB6.Format(mdtStartDate, "MM/dd/yyyy") & "')) AND (" & objChildTable.RealSource & "." & objEvent.StartDateName & " <= convert(datetime, '" & VB6.Format(mdtEndDate, "MM/dd/yyyy") & "'))) AND " & objChildTable.RealSource & "." & objEvent.EndDateName & " >= convert(datetime, '" & VB6.Format(mdtEndDate, "MM/dd/yyyy") & "'))" & vbNewLine & " OR (" & objChildTable.RealSource & "." & objEvent.StartDateName & " <= convert(datetime, '" & VB6.Format(mdtStartDate, "MM/dd/yyyy") & "') AND " & objChildTable.RealSource & "." & objEvent.EndDateName & " >= convert(datetime, '" & VB6.Format(mdtStartDate, "MM/dd/yyyy") & "')))" & vbNewLine _
									& " AND (" & objChildTable.RealSource & "." & objEvent.EndDateName & " >= " & objChildTable.RealSource & "." & objEvent.StartDateName & ") "

						ElseIf (objEvent.StartDateID > 0) And (objEvent.DurationID > 0) Then
							'event is defined by start date and duration
							sChildJoinCode = sChildJoinCode & " AND (" & objChildTable.RealSource & "." & objEvent.StartDateName & " IS NOT NULL AND " & objChildTable.RealSource & "." & objEvent.DurationName & " > 0)" & vbNewLine

						ElseIf objEvent.StartDateID > 0 And (objEvent.EndDateID < 1) And (objEvent.DurationID < 1) Then
							'event is defined by just the start date - one off event with a range of one
							sChildJoinCode = sChildJoinCode & " AND ((" & objChildTable.RealSource & "." & objEvent.StartDateName & " >= convert(datetime, '" & VB6.Format(mdtStartDate, "MM/dd/yyyy") & "') AND " & objChildTable.RealSource & "." & objEvent.StartDateName & " <= convert(datetime, '" & VB6.Format(mdtEndDate, "MM/dd/yyyy") & "')))" & vbNewLine

						End If
					End If
				End If

			Next intLoop

			mstrSQLJoin = mstrSQLJoin & sChildJoinCode

		Catch ex As Exception
			mstrErrorString = "Error in GenerateSQLJoin." & vbNewLine & ex.Message.RemoveSensitive()
			Return False

		End Try

		Return True

	End Function

	Private Function GenerateSQLWhere(pstrEventKey As String) As Boolean

		' Purpose : Generate the where clauses that cope with the joins
		'           NB Need to add the where clauses for filters/picklists etc

		Dim objEvent As clsCalendarEvent
		Dim strFilterIDs As String

		Dim blnOK As Boolean

		Try

			objEvent = mcolEvents.Item(pstrEventKey)

			'*******************************************************************************
			Dim pintLoop As Short
			Dim pobjTableView As TablePrivilege

			pobjTableView = gcoTablePrivileges.FindTableID(mlngCalendarReportsBaseTable)
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
						If (mlngTableViews(1, pintLoop) = 1) Then
							mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " OR ", " WHERE (") & mstrBaseTableRealSource & ".ID IN (SELECT ID FROM " & pobjTableView.RealSource & ")"
						End If

					Next pintLoop

					If Len(mstrSQLWhere) > 0 Then mstrSQLWhere = mstrSQLWhere & ")"
				End If

			End If
			'*******************************************************************************
			If mlngSingleRecordID > 0 Then
				mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & mstrBaseTableRealSource & ".ID IN (" & mstrSQLIDs & ") "

				' Now if we are using a picklist, add a where clause for that
				'Get List of IDs from Picklist
			ElseIf mlngCalendarReportsPickListID > 0 Then
				mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & mstrBaseTableRealSource & ".ID IN (" & mstrSQLIDs & ")"

			ElseIf mlngCalendarReportsFilterID > 0 Then
				mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & mstrBaseTableRealSource & ".ID IN (" & mstrSQLIDs & ")"

			End If

			'TM03032004 Fault 8184
			mstrSQLBaseData = mstrSQLBaseData & mstrSQLWhere

			'add clause to SQL, so that only dates within the specified range are retieved.
			If objEvent.TableID = mlngCalendarReportsBaseTable Then
				mstrSQLBaseDateClause = vbNullString

				'add clause to SQL, so that only dates within the specified range are retrieved.
				If objEvent.StartDateID > 0 And objEvent.EndDateID > 0 Then
					'event is defined by start date and end date
					mstrSQLBaseDateClause = mstrSQLBaseDateClause & "((" & mstrSQLBaseStartDateColumn & " <= convert(datetime, '" & VB6.Format(mdtStartDate, "MM/dd/yyyy") & "') AND " & mstrSQLBaseEndDateColumn & " >= convert(datetime, '" & VB6.Format(mdtStartDate, "MM/dd/yyyy") & "'))" & vbNewLine
					mstrSQLBaseDateClause = mstrSQLBaseDateClause & " OR (" & mstrSQLBaseStartDateColumn & " >= convert(datetime, '" & VB6.Format(mdtStartDate, "MM/dd/yyyy") & "') AND " & mstrSQLBaseEndDateColumn & " <= convert(datetime, '" & VB6.Format(mdtEndDate, "MM/dd/yyyy") & "'))" & vbNewLine
					mstrSQLBaseDateClause = mstrSQLBaseDateClause & " OR (((" & mstrSQLBaseStartDateColumn & " >= convert(datetime, '" & VB6.Format(mdtStartDate, "MM/dd/yyyy") & "')) AND (" & mstrSQLBaseStartDateColumn & " <= convert(datetime, '" & VB6.Format(mdtEndDate, "MM/dd/yyyy") & "'))) AND " & mstrSQLBaseEndDateColumn & " >= convert(datetime, '" & VB6.Format(mdtEndDate, "MM/dd/yyyy") & "'))" & vbNewLine
					mstrSQLBaseDateClause = mstrSQLBaseDateClause & " OR (" & mstrSQLBaseStartDateColumn & " <= convert(datetime, '" & VB6.Format(mdtStartDate, "MM/dd/yyyy") & "') AND " & mstrSQLBaseEndDateColumn & " >= convert(datetime, '" & VB6.Format(mdtStartDate, "MM/dd/yyyy") & "')))" & vbNewLine
					mstrSQLBaseDateClause = mstrSQLBaseDateClause & " AND (" & mstrSQLBaseEndDateColumn & ">=" & mstrSQLBaseStartDateColumn & ")"

				ElseIf (objEvent.StartDateID > 0) And (objEvent.DurationID > 0) Then
					'TM 25/04/2005 - Faults 10039 & 10040 - Check if the Start Date + Duration puts event in the report range.
					'event is defined by start date and duration
					mstrSQLBaseDateClause = mstrSQLBaseDateClause & "    (" & mstrSQLBaseDurationColumn & " > 0)" & vbNewLine
					mstrSQLBaseDateClause = mstrSQLBaseDateClause & "    AND (" & vbNewLine & vbNewLine

					' 1 Event Start Date before Report Start Date, Duration carrys event into, but not beyond the Report Range.
					mstrSQLBaseDateClause = mstrSQLBaseDateClause & "        (" & mstrSQLBaseStartDateColumn & " < convert(datetime, '" & Replace(VB6.Format(mdtStartDate, "MM/dd/yyyy"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & "')" & vbNewLine
					mstrSQLBaseDateClause = mstrSQLBaseDateClause & "      AND (DATEADD(day, " & mstrSQLBaseDurationColumn & ", " & mstrSQLBaseStartDateColumn & ") >= convert(datetime, '" & Replace(VB6.Format(mdtStartDate, "MM/dd/yyyy"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & "'))" & vbNewLine
					mstrSQLBaseDateClause = mstrSQLBaseDateClause & "        AND (DATEADD(day, " & mstrSQLBaseDurationColumn & ", " & mstrSQLBaseStartDateColumn & ") <= convert(datetime, '" & Replace(VB6.Format(mdtEndDate, "MM/dd/yyyy"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & "')))" & vbNewLine & vbNewLine

					' 2 Event Start Date within Report Range, Duration carrys event beyond Report End Date.
					mstrSQLBaseDateClause = mstrSQLBaseDateClause & "     OR ((" & mstrSQLBaseStartDateColumn & " >= convert(datetime, '" & Replace(VB6.Format(mdtStartDate, "MM/dd/yyyy"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & "'))" & vbNewLine
					mstrSQLBaseDateClause = mstrSQLBaseDateClause & "        AND (" & mstrSQLBaseStartDateColumn & " <= convert(datetime, '" & Replace(VB6.Format(mdtEndDate, "MM/dd/yyyy"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & "'))" & vbNewLine
					mstrSQLBaseDateClause = mstrSQLBaseDateClause & "      AND (DATEADD(day, " & mstrSQLBaseDurationColumn & ", " & mstrSQLBaseStartDateColumn & ") > convert(datetime, '" & Replace(VB6.Format(mdtEndDate, "MM/dd/yyyy"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & "')))" & vbNewLine & vbNewLine

					' 3 Event Start Date within Report Range and Duration keeps event within Report Range.
					mstrSQLBaseDateClause = mstrSQLBaseDateClause & "     OR ((" & mstrSQLBaseStartDateColumn & " >= convert(datetime, '" & Replace(VB6.Format(mdtStartDate, "MM/dd/yyyy"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & "'))" & vbNewLine
					mstrSQLBaseDateClause = mstrSQLBaseDateClause & "        AND (" & mstrSQLBaseStartDateColumn & " <= convert(datetime, '" & Replace(VB6.Format(mdtEndDate, "MM/dd/yyyy"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & "'))" & vbNewLine
					mstrSQLBaseDateClause = mstrSQLBaseDateClause & "      AND (DATEADD(day, " & mstrSQLBaseDurationColumn & ", " & mstrSQLBaseStartDateColumn & ") <= convert(datetime, '" & Replace(VB6.Format(mdtEndDate, "MM/dd/yyyy"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & "')))" & vbNewLine & vbNewLine

					' 4 Event Start Date before Report Start Date and Duration carrys event beyond Report End Date.
					mstrSQLBaseDateClause = mstrSQLBaseDateClause & "     OR ((" & mstrSQLBaseStartDateColumn & " < convert(datetime, '" & Replace(VB6.Format(mdtStartDate, "MM/dd/yyyy"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & "'))" & vbNewLine
					mstrSQLBaseDateClause = mstrSQLBaseDateClause & "      AND (DATEADD(day, " & mstrSQLBaseDurationColumn & ", " & mstrSQLBaseStartDateColumn & ") > convert(datetime, '" & Replace(VB6.Format(mdtEndDate, "MM/dd/yyyy"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & "')))" & vbNewLine & vbNewLine

					mstrSQLBaseDateClause = mstrSQLBaseDateClause & "        )" & vbNewLine

				ElseIf objEvent.StartDateID > 0 And (objEvent.EndDateID < 1) And (objEvent.DurationID < 1) Then
					'event is defined by just the start date - one off event with a range of one
					mstrSQLBaseDateClause = mstrSQLBaseDateClause & "(" & mstrSQLBaseStartDateColumn & " >= convert(datetime, '" & VB6.Format(mdtStartDate, "MM/dd/yyyy") & "') AND " & mstrSQLBaseStartDateColumn & " <= convert(datetime, '" & VB6.Format(mdtEndDate, "MM/dd/yyyy") & "')) "

				End If

				mstrSQLBaseDateClause = mstrSQLBaseDateClause & " AND (" & mstrSQLBaseStartDateColumn & " IS NOT NULL)"

				If objEvent.FilterID > 0 Then
					blnOK = FilteredIDs(objEvent.FilterID, strFilterIDs, mastrUDFsRequired, mvarPrompts)

					If blnOK Then
						mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & mstrBaseTableRealSource & ".ID IN (" & strFilterIDs & ")"
					Else
						' Permission denied on something in the filter.
						mstrErrorString = "You do not have permission to use the '" & General.GetFilterName(objEvent.FilterID) & "' filter."
						Return False
					End If
				End If

				If Len(mstrSQLWhere) > 0 Then
					mstrSQLWhere = mstrSQLWhere & " AND (" & mstrSQLBaseDateClause & ") "
				Else
					mstrSQLWhere = mstrSQLWhere & " WHERE " & mstrSQLBaseDateClause
				End If
			End If

			If (Len(mstrSQLDynamicLegendWhere) > 0) Then
				If (Len(mstrSQLWhere) > 0) Then
					mstrSQLWhere = mstrSQLWhere & " AND (" & mstrSQLDynamicLegendWhere & ") "
				Else
					mstrSQLWhere = mstrSQLWhere & " WHERE " & mstrSQLDynamicLegendWhere
				End If
			End If

			mstrSQLDynamicLegendWhere = vbNullString
			mstrSQLBaseStartDateColumn = vbNullString
			mstrSQLBaseEndDateColumn = vbNullString
			mstrSQLBaseDurationColumn = vbNullString

		Catch ex As Exception
			mstrErrorString = "Error in GenerateSQLWhere." & vbNewLine & ex.Message.RemoveSensitive()
			Return False

		End Try

		Return True

	End Function

	Private Function GenerateSQLOrderBy() As Boolean

		' Purpose : Returns order by string from the sort order array
		mstrSQLOrderBy = " ORDER BY " & mstrSQLOrderBy
		mstrSQLBaseData = mstrSQLBaseData & mstrSQLOrderBy
		Return True

	End Function

	Public Function ClearUp() As Boolean

		Try
			AccessLog.UtilUpdateLastRun(UtilityType.utlCalendarReport, mlngCalendarReportID)

			' Delete the temptable if exists
			General.DropUniqueSQLObject(mstrTempTableName, 3)

			Return True

		Catch ex As Exception
			Throw

		End Try

	End Function

	Public Function IsRecordSelectionValid() As Boolean

		Dim lngFilterID As Integer
		Dim objEvent As clsCalendarEvent
		Dim iResult As RecordSelectionValidityCodes
		Dim fCurrentUserIsSysSecMgr As Boolean

		fCurrentUserIsSysSecMgr = CurrentUserIsSysSecMgr()

		' Base Table First
		If mlngCalendarReportsFilterID > 0 Then
			iResult = ValidateRecordSelection(RecordSelectionType.Filter, mlngCalendarReportsFilterID)
			Select Case iResult
				Case RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
					mstrErrorString = "The base table filter used in this definition has been deleted by another user."
				Case RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
					mstrErrorString = "The base table filter used in this definition is invalid."
				Case RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
					If Not fCurrentUserIsSysSecMgr Then
						mstrErrorString = "The base table filter used in this definition has been made hidden by another user."
					End If
			End Select
		ElseIf mlngCalendarReportsPickListID > 0 Then
			iResult = ValidateRecordSelection(RecordSelectionType.Picklist, mlngCalendarReportsPickListID)
			Select Case iResult
				Case RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
					mstrErrorString = "The base table picklist used in this definition has been deleted by another user."
				Case RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
					mstrErrorString = "The base table picklist used in this definition is invalid."
				Case RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
					If Not fCurrentUserIsSysSecMgr Then
						mstrErrorString = "The base table picklist used in this definition has been made hidden by another user."
					End If
			End Select
		End If

		'Description Calculation
		If Len(mstrErrorString) = 0 Then
			If mlngDescriptionExpr > 0 Then
				iResult = ValidateCalculation(mlngDescriptionExpr)
				Select Case iResult
					Case RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
						mstrErrorString = "The base description calculation used in this definition has been deleted by another user."
					Case RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
						mstrErrorString = "The base description calculation used in this definition is invalid."
					Case RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
						If Not fCurrentUserIsSysSecMgr Then
							mstrErrorString = "The base description calculation used in this definition has been made hidden by another user."
						End If
				End Select
			End If
		End If

		'Events Filters
		For Each objEvent In mcolEvents.Collection
			lngFilterID = objEvent.FilterID
			If lngFilterID > 0 Then
				iResult = ValidateRecordSelection(RecordSelectionType.Filter, lngFilterID)
				Select Case iResult
					Case RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
						mstrErrorString = "An event table filter used in this definition has been deleted by another user."
					Case RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
						mstrErrorString = "An event table filter used in this definition is invalid."
					Case RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
						If Not fCurrentUserIsSysSecMgr Then
							mstrErrorString = "An event table filter used in this definition has been made hidden by another user."
						End If
				End Select
			End If
		Next objEvent
		'UPGRADE_NOTE: Object objEvent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objEvent = Nothing

		'Report Start Date Calculation
		If Len(mstrErrorString) = 0 Then
			If mlngStartDateExpr > 0 Then
				iResult = ValidateCalculation(mlngStartDateExpr)
				Select Case iResult
					Case RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
						mstrErrorString = "The report start date calculation used in this definition has been deleted by another user."
					Case RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
						mstrErrorString = "The report start date calculation used in this definition is invalid."
					Case RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
						If Not fCurrentUserIsSysSecMgr Then
							mstrErrorString = "The report start date calculation used in this definition has been made hidden by another user."
						End If
				End Select
			End If
		End If

		'Report End Date Calculation
		If Len(mstrErrorString) = 0 Then
			If mlngEndDateExpr > 0 Then
				iResult = ValidateCalculation(mlngEndDateExpr)
				Select Case iResult
					Case RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
						mstrErrorString = "The report end date calculation used in this definition has been deleted by another user."
					Case RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
						mstrErrorString = "The report end date calculation used in this definition is invalid."
					Case RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
						If Not fCurrentUserIsSysSecMgr Then
							mstrErrorString = "The report end date calculation used in this definition has been made hidden by another user."
						End If
				End Select
			End If
		End If

		Return (Len(mstrErrorString) = 0)

	End Function

	Public Function ConvertDescription(pvarDesc1 As String, pvarDesc2 As String, pvarDesc3 As String) As String

		Dim strBaseDescription1, strBaseDescription2 As String
		Dim strBaseDescriptionExpr As String
		Dim strTempRecordDesc As String

		'Get base description 1
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If Not IsDBNull(pvarDesc1) Then
			Select Case mintType_BaseDesc1
				Case 3
					strBaseDescription1 = Format(pvarDesc1, mstrFormat_BaseDesc1)
				Case 2
					strBaseDescription1 = IIf(pvarDesc1, "Y", "N")
				Case 1
					strBaseDescription1 = VB6.Format(pvarDesc1, mstrClientDateFormat)
				Case 0
					strBaseDescription1 = pvarDesc1
			End Select
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object strBaseDescription1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strBaseDescription1 = vbNullString
		End If
		' Get base description 2
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If Not IsDBNull(pvarDesc2) Then
			Select Case mintType_BaseDesc2
				Case 3
					strBaseDescription2 = VB6.Format(pvarDesc2, mstrFormat_BaseDesc2)
				Case 2
					strBaseDescription2 = IIf(pvarDesc2, "Y", "N")
				Case 1
					strBaseDescription2 = VB6.Format(pvarDesc2, mstrClientDateFormat)
				Case 0
					strBaseDescription2 = pvarDesc2
			End Select
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object strBaseDescription2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strBaseDescription2 = vbNullString
		End If
		' Get base description expression
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If Not IsDBNull(pvarDesc3) Then
			Select Case mintType_BaseDescExpr
				Case 2
					strBaseDescriptionExpr = IIf(pvarDesc3, "Y", "N")
				Case 1
					strBaseDescriptionExpr = VB6.Format(pvarDesc3, mstrClientDateFormat)
				Case 0
					strBaseDescriptionExpr = pvarDesc3
			End Select
		Else
			strBaseDescriptionExpr = vbNullString
		End If

		strTempRecordDesc = strBaseDescription1
		strTempRecordDesc = strTempRecordDesc & IIf((Len(strTempRecordDesc) > 0) And (Len(strBaseDescription2) > 0), mstrDescriptionSeparator, "") & strBaseDescription2
		strTempRecordDesc = strTempRecordDesc & IIf((Len(strTempRecordDesc) > 0) And (Len(strBaseDescriptionExpr) > 0), mstrDescriptionSeparator, "") & strBaseDescriptionExpr

		Return strTempRecordDesc

	End Function

	Private Function GetValueForRecordIndependantCalc(lngExprID As Integer, Optional ByRef pvarPrompts As Object = Nothing) As Date

		Dim objExpr As clsExprExpression
		Dim rsTemp As DataTable
		Dim strSQL As String
		Dim fOK As Boolean
		Dim lngViews(,) As Integer

		Try

			objExpr = New clsExprExpression(SessionInfo)
			With objExpr
				' Initialise the filter expression object.
				fOK = .Initialise(0, lngExprID, ExpressionTypes.giEXPR_RECORDINDEPENDANTCALC, ExpressionValueTypes.giEXPRVALUE_UNDEFINED)

				If fOK Then
					fOK = objExpr.RuntimeCalculationCode(lngViews, strSQL, Nothing, True, False, pvarPrompts)
				End If

				If fOK Then
					rsTemp = DB.GetDataTable("SELECT " & strSQL)
					If rsTemp.Rows.Count > 0 Then
						Return CDate(rsTemp.Rows(0)(0))
					End If
				End If

			End With

		Catch ex As Exception
			Return Date.MinValue

		End Try

		Return Date.MinValue

	End Function

	Public Function GetDefaultRegion(plngBaseRecordID As Integer, pdtDate As Date) As String

		Dim intCount As Integer

		Try

			For intCount = 1 To UBound(mavCareerRanges, 2) Step 1
				'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If plngBaseRecordID = CInt(mavCareerRanges(0, intCount)) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(2, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If (Not String.IsNullOrEmpty(mavCareerRanges(2, intCount))) Then
						'has a career change in the past
						'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If (pdtDate >= CDate(mavCareerRanges(1, intCount))) And (pdtDate < CDate(mavCareerRanges(2, intCount))) Then
							'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(3, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Return mavCareerRanges(3, intCount)
						End If
					Else
						'has a effective start date but has no end date. (most recent career change)
						'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If (pdtDate >= CDate(mavCareerRanges(1, intCount))) Then
							'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(3, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Return mavCareerRanges(3, intCount)
						End If
					End If
				End If
			Next intCount

		Catch ex As Exception
			Return ""

		End Try

		Return ""

	End Function

End Class