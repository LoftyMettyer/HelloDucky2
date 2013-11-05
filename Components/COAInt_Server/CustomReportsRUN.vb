Option Strict Off
Option Explicit On

Imports ADODB
Imports HR.Intranet.Server.Enums

Public Class Report

	' To hold Properties
	Private mlngCustomReportID As Integer
	Private mstrErrorString As String

	' Variables to store definition
	Private mstrCustomReportsName As String
	Private mlngCustomReportsBaseTable As Integer
	Private mstrCustomReportsBaseTableName As String
	Private mlngCustomReportsPickListID As Integer
	Private mlngCustomReportsFilterID As Integer
	Private mlngCustomReportsParent1Table As Integer
	Private mlngCustomReportsParent1FilterID As Integer
	Private mlngCustomReportsParent2Table As Integer
	Private mlngCustomReportsParent2FilterID As Integer
	Private mblnCustomReportsSummaryReport As Boolean
	Private mblnIgnoreZerosInAggregates As Boolean
	Private mblnCustomReportsPrintFilterHeader As Boolean

	'New Default Output Variables
	Private mlngOutputFormat As Integer
	Private mblnOutputScreen As Boolean
	Private mblnOutputPrinter As Boolean
	Private mstrOutputPrinterName As String
	Private mblnOutputSave As Boolean
	Private mlngOutputSaveExisting As Integer
	Private mblnOutputEmail As Boolean
	Private mlngOutputEmailID As Integer
	Private mstrOutputEmailName As String
	Private mstrOutputEmailSubject As String
	Private mstrOutputEmailAttachAs As String
	Private mstrOutputFilename As String
	Private mblnOutputPreview As Boolean

	Private mvarChildTables(,) As Object
	Private miChildTablesCount As Short
	Private miUsedChildCount As Short

	Private mlngCustomReportsParent1PickListID As Integer
	Private mlngCustomReportsParent2PickListID As Integer

	' Recordsets to store the definition and column information
	Private mrstCustomReportsDetails As New Recordset

	' Classes
	Private mclsData As clsDataAccess
	Private mclsGeneral As clsGeneral
	Private mobjEventLog As clsEventLog

	' TableViewsGuff
	Private mstrRealSource As String
	Private mstrBaseTableRealSource As String
	Private mlngTableViews(,) As Integer
	Private mstrViews() As String
	Private mobjTableView As CTablePrivilege
	Private mobjColumnPrivileges As CColumnPrivileges

	' Strings to hold the SQL statement
	Private mstrSQLSelect As String
	Private mstrSQLFrom As String
	Private mstrSQLJoin As String
	Private mstrSQLWhere As String
	Private mstrSQLOrderBy As String
	Private mstrSQL As String

	' Array holding the columns to sort the report by
	Private mvarSortOrder(,) As Object

	' Array to hold the columns used in the report
	Dim mvarColDetails(,) As Object
	Dim mstrExcelFormats() As String
	Dim mvarVisibleColumns(,) As Object

	'Array used to store the 'GroupWithNextColumn' option strings.
	Private mvarGroupWith(,) As Object

	'Array used to store the 'POC' values when outputting.
	Private mvarPageBreak() As Object
	Private mblnPageBreak As Boolean
	Private mintPageBreakRowIndex As Integer

	' String to hold the temp table name
	Private mstrTempTableName As String

	' Recordset to store the final data from the temp table
	Private mrstCustomReportsOutput As New Recordset

	'Does the report generate no records ?
	Private mblnNoRecords As Boolean

	' Is this a Bradford Index Report
	Private mbIsBradfordIndexReport As Boolean

	Private mvarOutputArray_Columns() As String
	Private mvarOutputArray_Data As ArrayList
	Private mvarPrompts(,) As Object

	' Flags used when populating the grid
	Private mblnReportHasSummaryInfo As Boolean
	Private mblnReportHasPageBreak As Boolean
	Private mblnDoesHaveGrandSummary As Boolean

	Private mstrClientDateFormat As String
	Private mstrLocalDecimalSeparator As String
	Private mlngColumnLimit As Integer

	Private Const lng_SEQUENCECOLUMNNAME As String = "?ID_SEQUENCE_COLUMN"

	Private mbUseSequence As Boolean

	Private mstrBradfordStartDate As String
	Private mstrBradfordEndDate As String

	'Variables to hold Bradford Factor display/include options
	Private mbBradfordSRV As Boolean
	Private mbBradfordTotals As Boolean
	Private mbBradfordCount As Boolean
	Private mbBradfordWorkings As Boolean
	Private mstrOrderByColumn As String
	Private mlngOrderByColumnID As Integer
	Private mstrGroupByColumn As String
	Private mlngGroupByColumnID As Integer
	Private mbOrderBy1Asc As Boolean
	Private mbOrderBy2Asc As Boolean
	Private mbOmitBeforeStart As Boolean
	Private mbOmitAfterEnd As Boolean

	Private mstrAbsenceRealSource As String

	Private mbMinBradford As Boolean
	Private mlngMinBradfordAmount As Integer
	Private mbDisplayBradfordDetail As Boolean

	Private mlngPersonnelID As Integer

	' Array holding the User Defined functions that are needed for this report
	Private mastrUDFsRequired() As String

	'Runnning report for single record only!
	Private mlngSingleRecordID As Integer

Private Enum enum_RowType
	Data = 0
	Count = 1
	Average = 2
	Total = 3
End Enum

	Public ReadOnly Property HasSummaryColumns() As Boolean
		Get
			Return mblnReportHasSummaryInfo Or mbIsBradfordIndexReport
		End Get
	End Property

	Public WriteOnly Property SingleRecordID() As Integer
		Set(ByVal Value As Integer)
			mlngSingleRecordID = Value

		End Set
	End Property
	Public ReadOnly Property BaseTableName() As String
		Get
			BaseTableName = mstrCustomReportsBaseTableName
		End Get
	End Property


	Public ReadOnly Property ChildCount() As Short
		Get
			ChildCount = miChildTablesCount
		End Get
	End Property



	Public ReadOnly Property UsedChildCount() As Short
		Get
			UsedChildCount = miUsedChildCount
		End Get
	End Property


	'-----------------------------------------
	' Variables used for intranet are above this line

	'Batch Job Mode ?
	'Private mblnBatchMode As Boolean

	'Has the user cancelled the report ?
	'Private mblnUserCancelled As Boolean

	Public WriteOnly Property ClientDateFormat() As String
		Set(ByVal Value As String)

			' Clients date format passed in from the asp page
			mstrClientDateFormat = Value

		End Set
	End Property

	Public WriteOnly Property ColumnLimit() As Integer
		Set(ByVal Value As Integer)

			' Clients date format passed in from the asp page
			mlngColumnLimit = Value

		End Set
	End Property

	Public WriteOnly Property LocalDecimalSeparator() As String
		Set(ByVal Value As String)

			' Clients date format passed in from the asp page
			mstrLocalDecimalSeparator = Value

		End Set
	End Property

	Public WriteOnly Property Failed() As Boolean
		Set(ByVal Value As Boolean)

			' Connection object passed in from the asp page
			If Value = True Then
				mobjEventLog.ChangeHeaderStatus(EventLog_Status.elsFailed)
			End If

		End Set
	End Property

	Public WriteOnly Property FailedMessage() As String
		Set(ByVal Value As String)
			mobjEventLog.AddDetailEntry(Value)

		End Set
	End Property


	Public WriteOnly Property Cancelled() As Boolean
		Set(ByVal Value As Boolean)

			' Connection object passed in from the asp page
			If Value = True Then
				mobjEventLog.ChangeHeaderStatus(EventLog_Status.elsCancelled)
			Else
				mobjEventLog.ChangeHeaderStatus(EventLog_Status.elsSuccessful)
			End If

		End Set
	End Property

	Public ReadOnly Property NoRecords() As Boolean
		Get

			' Does the report have any records ?
			NoRecords = mblnNoRecords

		End Get
	End Property

	Public ReadOnly Property SQLSTRING() As String
		Get

			' Does the report have any records ?
			SQLSTRING = mstrSQL

		End Get
	End Property

	Public ReadOnly Property ReportHasPageBreak() As Boolean
		Get

			' Does the report have a page break
			ReportHasPageBreak = mblnReportHasPageBreak

		End Get
	End Property

	Public ReadOnly Property ReportHasSummaryInfo() As Boolean
		Get

			' Does the report have summary info
			ReportHasSummaryInfo = mblnReportHasSummaryInfo

		End Get
	End Property

	Public ReadOnly Property CustomReportsSummaryReport() As Boolean
		Get
			CustomReportsSummaryReport = mblnCustomReportsSummaryReport
		End Get
	End Property


	Public ReadOnly Property CustomReportsName() As String
		Get
			CustomReportsName = mstrCustomReportsName
		End Get
	End Property


	Public ReadOnly Property OutputArray_VisibleColumns() As Object
		Get

			OutputArray_VisibleColumns = VB6.CopyArray(mvarVisibleColumns)

		End Get
	End Property

	Public ReadOnly Property OutputArray_Columns() As String()
		Get
			Return mvarOutputArray_Columns
		End Get
	End Property

	Public ReadOnly Property OutputArray_PageBreakValues() As Object
		Get

			'Holds all the page break values for the report. The index is the same as the grids row number!
			OutputArray_PageBreakValues = VB6.CopyArray(mvarPageBreak)

		End Get
	End Property

	Public ReadOnly Property OutputArray_ExcelFormats() As Object
		Get
			Dim avTemp() As Object
			Dim iLoop As Short

			ReDim avTemp(UBound(mstrExcelFormats))

			For iLoop = LBound(mstrExcelFormats) To UBound(mstrExcelFormats)
				'UPGRADE_WARNING: Couldn't resolve default property of object avTemp(iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				avTemp(iLoop) = mstrExcelFormats(iLoop)
			Next iLoop

			OutputArray_ExcelFormats = VB6.CopyArray(avTemp)

		End Get
	End Property

	Public ReadOnly Property OutputArray_Heading(ByVal lngIndex As Object) As String
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object lngIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, lngIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			OutputArray_Heading = mvarColDetails(0, lngIndex)
		End Get
	End Property


	Public ReadOnly Property OutputArray_DataType(ByVal lngIndex As Object) As Integer
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object lngIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			OutputArray_DataType = IIf(mvarColDetails(3, lngIndex), SQLDataType.sqlNumeric, SQLDataType.sqlVarChar)
		End Get
	End Property

	Public ReadOnly Property OutputArray_Decimals(ByVal lngIndex As Object) As Integer
		Get
			'UPGRADE_WARNING: Couldn't resolve default property of object lngIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, lngIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			OutputArray_Decimals = mvarColDetails(2, lngIndex)
		End Get
	End Property

	Public WriteOnly Property Connection() As Connection
		Set(ByVal Value As Connection)
			gADOCon = Value
		End Set
	End Property

	Public WriteOnly Property Username() As String
		Set(ByVal Value As String)

			' Username passed in from the asp page
			gsUsername = Value

		End Set
	End Property

	'Public Property Get UserCancelled() As Boolean
	'  UserCancelled = mblnUserCancelled
	'End Property

	Public WriteOnly Property CustomReportID() As Integer
		Set(ByVal Value As Integer)

			' ID of the report to run passed in from the asp page
			mlngCustomReportID = Value
			mlngSingleRecordID = 0

		End Set
	End Property

	Public ReadOnly Property ErrorString() As String
		Get

			' Error information passed back to the asp page
			ErrorString = mstrErrorString

		End Get
	End Property

	Public ReadOnly Property OutputPreview() As Boolean
		Get
			OutputPreview = mblnOutputPreview
		End Get
	End Property

	Public ReadOnly Property OutputFormat() As Integer
		Get
			OutputFormat = mlngOutputFormat
		End Get
	End Property

	Public ReadOnly Property OutputScreen() As Boolean
		Get
			OutputScreen = mblnOutputScreen
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

	Public ReadOnly Property OutputFilename() As String
		Get
			OutputFilename = mstrOutputFilename
		End Get
	End Property


	Public Property EventLogID() As Integer
		Get
			EventLogID = mobjEventLog.EventLogID
		End Get
		Set(ByVal Value As Integer)
			mobjEventLog.EventLogID = Value
		End Set
	End Property

	Public Function Output_GridForm() As String

		Dim sTemp As String
		Dim lngCount As Integer
		Dim asStrings() As String

		Const STRINGLENGTH As Short = 5000

		ReDim asStrings(0)

		sTemp = vbNullString
		sTemp = sTemp & "   <form name=frmGridItems id=frmGridItems>" & vbNewLine

		For lngCount = 0 To mvarOutputArray_Data.Count - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarOutputArray_Data(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sTemp = sTemp & "      <input type=hidden id=" & "txtGridItem_" & CStr(lngCount) & " name=" & "txtGridItem_" & CStr(lngCount) & " value=""" & Replace(CStr(mvarOutputArray_Data(lngCount)), """", "&quot;") & """>" & vbNewLine

			If Len(sTemp) > STRINGLENGTH Then
				ReDim Preserve asStrings(UBound(asStrings) + 1)
				asStrings(UBound(asStrings)) = sTemp
				sTemp = vbNullString
			End If
		Next lngCount
		sTemp = sTemp & "</form>" & vbNewLine

		ReDim Preserve asStrings(UBound(asStrings) + 1)
		asStrings(UBound(asStrings)) = sTemp
		sTemp = vbNullString

		For lngCount = 1 To UBound(asStrings)
			sTemp = sTemp & asStrings(lngCount)
		Next lngCount

		Output_GridForm = sTemp

	End Function

	Public Function SetPromptedValues(ByRef pavPromptedValues As Object) As Boolean

		' Purpose : This function calls the individual functions that
		'           generate the components of the main SQL string.
		On Error GoTo ErrorTrap

		Dim fOK As Boolean
		Dim iLoop As Short
		Dim iDataType As Short
		Dim lngComponentID As Integer

		fOK = True

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
							'mvarPrompts(1, iLoop) = CDate(pavPromptedValues(1, iLoop))
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

		SetPromptedValues = fOK

		Exit Function

ErrorTrap:
		mstrErrorString = "Error setting prompted values." & vbNewLine & Err.Description
		mobjEventLog.AddDetailEntry(mstrErrorString)
		mobjEventLog.ChangeHeaderStatus(EventLog_Status.elsFailed)
		SetPromptedValues = False

	End Function

	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()

		' Initialise the the classes/arrays to be used
		mclsData = New clsDataAccess
		mclsGeneral = New clsGeneral
		'mclsUI = New clsUI
		mobjEventLog = New clsEventLog
		ReDim mvarSortOrder(2, 0)
		ReDim mvarColDetails(25, 0)
		ReDim mlngTableViews(2, 0)
		ReDim mstrViews(0)
		ReDim mvarOutputArray_Columns(0)
		mvarOutputArray_Data = New ArrayList()
		ReDim mvarVisibleColumns(3, 0)

		ReDim mvarGroupWith(1, 0)
		ReDim mvarPageBreak(0)

		' By default this is not a Bradford Index Report
		mbIsBradfordIndexReport = False

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
		'UPGRADE_NOTE: Object mclsUI may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		'mclsUI = Nothing
		'UPGRADE_NOTE: Object mobjEventLog may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mobjEventLog = Nothing
		' JPD20030313 Do not drop the tables & columns collections as they can be reused.
		'Set gcoTablePrivileges = Nothing
		'Set gcolColumnPrivilegesCollection = Nothing

	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub

	Public Function AddTempTableToSQL() As Boolean

		On Error GoTo AddTempTableToSQL_ERROR

		mstrTempTableName = datGeneral.UniqueSQLObjectName("ASRSysTempCustomReport", 3)

		mstrSQLSelect = mstrSQLSelect & " INTO [" & mstrTempTableName & "]"

		AddTempTableToSQL = True
		Exit Function

AddTempTableToSQL_ERROR:

		AddTempTableToSQL = False
		mstrErrorString = "Error retrieving unique temp table name." & vbNewLine & Err.Description
		mobjEventLog.AddDetailEntry(mstrErrorString)
		mobjEventLog.ChangeHeaderStatus(EventLog_Status.elsFailed)

	End Function

	Public Function MergeSQLStrings() As Boolean

		On Error GoTo MergeSQLStrings_ERROR

		mstrSQL = mstrSQLSelect & " FROM " & mstrSQLFrom & IIf(Len(mstrSQLJoin) = 0, "", " " & mstrSQLJoin) & IIf(Len(mstrSQLWhere) = 0, "", " " & mstrSQLWhere) & " " & mstrSQLOrderBy

		MergeSQLStrings = True

		Exit Function

MergeSQLStrings_ERROR:

		MergeSQLStrings = False
		mstrErrorString = "Error merging SQL string components." & vbNewLine & Err.Description
		mobjEventLog.AddDetailEntry(mstrErrorString)
		mobjEventLog.ChangeHeaderStatus(EventLog_Status.elsFailed)

	End Function

	Public Function ExecuteSql() As Boolean

		On Error GoTo ExecuteSQL_ERROR

		mclsData.ExecuteSql(mstrSQL)

		ExecuteSql = True
		Exit Function

ExecuteSQL_ERROR:

		ExecuteSql = False
		mstrErrorString = "Error executing SQL statement." & vbNewLine & Err.Description
		mobjEventLog.AddDetailEntry(mstrErrorString)
		mobjEventLog.ChangeHeaderStatus(EventLog_Status.elsFailed)

	End Function

	Public Function GetCustomReportDefinition() As Boolean

		On Error GoTo GetCustomReportDefinition_ERROR

		Dim rsDefinition As Recordset
		Dim strSQL As String
		Dim i As Short

		SetupTablesCollection()

		mbIsBradfordIndexReport = False

		strSQL = "EXEC spASRIntGetCustomReport " & mlngCustomReportID
		rsDefinition = mclsData.OpenRecordset(strSQL, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly)

		With rsDefinition

			' Dont run if its been deleted by another user.
			If .BOF And .EOF Then
				GetCustomReportDefinition = False
				mstrErrorString = "Report has been deleted by another user."
				Exit Function
			End If

			' RH 29/05/01 - Dont run if its been made hidden by another user.
			If LCase(CType(.Fields("Username").Value, String)) <> LCase(gsUsername) And CurrentUserAccess(UtilityType.utlCustomReport, mlngCustomReportID) = ACCESS_HIDDEN Then
				GetCustomReportDefinition = False
				mstrErrorString = "Report has been made hidden by another user."
				Exit Function
			End If

			mstrCustomReportsName = .Fields("Name").Value
			mlngCustomReportsBaseTable = .Fields("BaseTable").Value
			mstrCustomReportsBaseTableName = .Fields("TableName").Value
			mlngCustomReportsPickListID = .Fields("picklist").Value
			mlngCustomReportsFilterID = .Fields("Filter").Value
			mlngCustomReportsParent1Table = .Fields("parent1table").Value
			mlngCustomReportsParent1FilterID = .Fields("parent1filter").Value
			mlngCustomReportsParent2Table = .Fields("parent2table").Value
			mlngCustomReportsParent2FilterID = .Fields("parent2filter").Value

			mblnCustomReportsSummaryReport = .Fields("Summary").Value
			mblnIgnoreZerosInAggregates = .Fields("IgnoreZeros").Value
			mblnCustomReportsPrintFilterHeader = .Fields("PrintFilterHeader").Value
			mlngCustomReportsParent1PickListID = .Fields("parent1Picklist").Value
			mlngCustomReportsParent2PickListID = .Fields("parent2Picklist").Value

			'New Default Output Variables
			mblnOutputPreview = .Fields("OutputPreview").Value
			mlngOutputFormat = .Fields("OutputFormat").Value
			mblnOutputScreen = .Fields("OutputScreen").Value
			mblnOutputPrinter = .Fields("OutputPrinter").Value
			mstrOutputPrinterName = .Fields("OutputPrinterName").Value
			mblnOutputSave = .Fields("OutputSave").Value
			mlngOutputSaveExisting = .Fields("OutputSaveExisting").Value
			mblnOutputEmail = .Fields("OutputEmail").Value
			mlngOutputEmailID = .Fields("OutputEmailAddr").Value
			mstrOutputEmailName = .Fields("EmailGroupName").Value
			mstrOutputEmailSubject = .Fields("OutputEmailSubject").Value
			mstrOutputEmailAttachAs = .Fields("OutputEmailAttachAs").Value
			mstrOutputFilename = .Fields("OutputFilename").Value
			mblnOutputPreview = (.Fields("OutputPreview").Value Or (mlngOutputFormat = OutputFormats.fmtDataOnly And mblnOutputScreen))

		End With

		' Child data recordset
		rsDefinition = rsDefinition.NextRecordset()

		i = 0
		With rsDefinition
			Do Until .EOF
				ReDim Preserve mvarChildTables(5, i)
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(0, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarChildTables(0, i) = .Fields("ChildTable").Value	'Childs Table ID
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(1, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarChildTables(1, i) = .Fields("childFilter").Value 'Childs Filter ID (if any)
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(2, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarChildTables(2, i) = .Fields("ChildMaxRecords").Value 'Number of records to take from child
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(3, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarChildTables(3, i) = .Fields("TableName").Value 'Child Table Name
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(4, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarChildTables(4, i) = False	'Boolean - True if table is used, False if not
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(5, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarChildTables(5, i) = .Fields("ChildOrder").Value	'Childs Order ID (if any)
				i = i + 1
				.MoveNext()
			Loop
		End With

		miChildTablesCount = i

		If Not IsRecordSelectionValid() Then
			GetCustomReportDefinition = False
			Exit Function
		End If

		GetCustomReportDefinition = True

		mobjEventLog.AddHeader(EventLog_Type.eltCustomReport, mstrCustomReportsName)

TidyAndExit:

		'UPGRADE_NOTE: Object rsTemp_Definition may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsDefinition = Nothing

		Exit Function

GetCustomReportDefinition_ERROR:

		GetCustomReportDefinition = False
		mstrErrorString = "Error reading the Custom Report definition !" & vbNewLine & Err.Description
		mobjEventLog.AddDetailEntry(mstrErrorString)
		mobjEventLog.ChangeHeaderStatus(EventLog_Status.elsFailed)
		Resume TidyAndExit

	End Function

	Public Function GetDetailsRecordsets() As Boolean

		' Purpose : This function loads report details and sort details into
		'           arrays and leaves the details recordset reference there
		'           (dont remove it...used for summary info !)

		On Error GoTo GetDetailsRecordsets_ERROR

		Dim strTempSQL As String
		Dim intTemp As Short
		Dim prstCustomReportsSortOrder As Recordset
		Dim lngTableID As Integer
		Dim sMask As String

		' Get the column information from the Details table, in order
		strTempSQL = "EXEC spASRIntGetCustomReportDetails " & mlngCustomReportID
		mrstCustomReportsDetails = mclsData.OpenRecordset(strTempSQL, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly)


		Dim objExpr As clsExprExpression
		With mrstCustomReportsDetails
			If .BOF And .EOF Then
				GetDetailsRecordsets = False
				mstrErrorString = "No columns found in the specified Custom Report definition." & vbNewLine & "Please remove this definition and create a new one."
				Exit Function
			End If

			If Not CheckCalcsStillExist() Then
				GetDetailsRecordsets = False
				Exit Function
			End If

			Do Until .EOF
				intTemp = UBound(mvarColDetails, 2) + 1
				ReDim Preserve mvarColDetails(UBound(mvarColDetails, 1), intTemp)

				ReDim Preserve mstrExcelFormats(intTemp) 'MH20010307

				'*************************************************************************
				'Now we need to decide on what the heading needs to be because QA want to
				'be able to have similar headings for hidden columns...I warned them, but
				'NO...they thought that the best move was to spend ages fixing faults in
				'v2 and put HR Pro .NET on the back-burner so that we can release a
				'Limited Edition of HR Pro called HR Pro .NET 2012 Olympic Edition.
				'What twats!!!...Fault 10211.

				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If IIf((IsDBNull(.Fields("Hidden").Value) Or (.Fields("Hidden")).Value), True, False) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(0, intTemp) = "?ID_HD_" & .Fields("Type").Value & "_" & .Fields("ColExprID").Value
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(0, intTemp) = .Fields("Heading").Value
				End If

				'*************************************************************************

				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(1, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(1, intTemp) = .Fields("Size").Value
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(2, intTemp) = .Fields("dp").Value
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(3, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(3, intTemp) = .Fields("IsNumeric").Value
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(4, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(4, intTemp) = .Fields("Avge").Value
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(5, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(5, intTemp) = .Fields("cnt").Value
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(6, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(6, intTemp) = .Fields("tot").Value
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(7, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(7, intTemp) = .Fields("boc").Value
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(8, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(8, intTemp) = .Fields("poc").Value
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(9, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(9, intTemp) = .Fields("voc").Value
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(10, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(10, intTemp) = .Fields("srv").Value
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(11, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(11, intTemp) = ""
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(12, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(12, intTemp) = .Fields("ColExprID").Value
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(13, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(13, intTemp) = .Fields("Type").Value

				lngTableID = IIf(IsDBNull(.Fields("TableID").Value), 0, .Fields("TableID").Value)
				mvarColDetails(14, intTemp) = lngTableID
				mvarColDetails(15, intTemp) = .Fields("TableName").Value

				If .Fields("Type").Value = "C" Then
					mvarColDetails(16, intTemp) = .Fields("ColumnName").Value
					mvarColDetails(17, intTemp) = .Fields("IsDateColumn").Value
					mvarColDetails(18, intTemp) = .Fields("IsBooleanColumn").Value

				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(16, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(16, intTemp) = ""

					'MH20010307
					objExpr = New clsExprExpression

					objExpr.ExpressionID = CInt(.Fields("ColExprID").Value)
					objExpr.ConstructExpression()
					objExpr.ValidateExpression(True)

					lngTableID = objExpr.BaseTableID
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(14, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(14, intTemp) = lngTableID
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(15, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(15, intTemp) = objExpr.BaseTableName
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(16, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(16, intTemp) = ""

					'UPGRADE_NOTE: Object objExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					objExpr = Nothing

					mvarColDetails(17, intTemp) = mclsGeneral.DateColumn(.Fields("Type").Value, lngTableID, .Fields("ColExprID").Value)
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(18, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(18, intTemp) = mclsGeneral.BitColumn(.Fields("Type").Value, lngTableID, .Fields("ColExprID").Value)

				End If

				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(19, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(19, intTemp) = IIf((IsDBNull(.Fields("Hidden").Value) Or (.Fields("Hidden")).Value), True, False)

				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(20, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(20, intTemp) = IsReportChildTable(lngTableID) 'Indicates if column is a report child table.

				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(21, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(21, intTemp) = IIf(.Fields("repetition").Value = 1, True, False)

				mvarColDetails(22, intTemp) = .Fields("Use1000separator").Value

				' Format for this numeric column
				If mvarColDetails(3, intTemp) Then
					sMask = ""
					If mvarColDetails(22, intTemp) Then sMask = ",0"
					If mvarColDetails(2, intTemp) > 0 Then sMask = sMask & "." & New String("0", mvarColDetails(2, intTemp))
					mvarColDetails(23, intTemp) = "{0:0" & sMask & ";-0" & sMask & ";0." & New String("0", mvarColDetails(2, intTemp)) & "}"
				End If


				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(24, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(24, intTemp) = IIf((IsDBNull(.Fields("GroupWithNextColumn").Value) Or (Not .Fields("GroupWithNextColumn").Value)), False, True)

				.MoveNext()
			Loop
			.MoveFirst()
		End With

		'******************************************************************************
		' Add the ID columns for the tables so that we can re-select the child records
		' when we create the multiple child temp table.
		' NB. Is called only when there is more than one child in the report.
		'******************************************************************************

		intTemp = UBound(mvarColDetails, 2) + 1
		ReDim Preserve mvarColDetails(UBound(mvarColDetails, 1), intTemp)

		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(1, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(1, intTemp) = 99
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(2, intTemp) = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(3, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(3, intTemp) = False
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(4, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(4, intTemp) = False
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(5, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(5, intTemp) = False
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(6, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(6, intTemp) = False
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(7, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(7, intTemp) = False
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(8, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(8, intTemp) = False
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(9, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(9, intTemp) = False
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(10, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(10, intTemp) = False
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(11, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(11, intTemp) = ""
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(12, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(12, intTemp) = -1
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(13, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(13, intTemp) = "C"

		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(14, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(14, intTemp) = mlngCustomReportsBaseTable
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(15, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(15, intTemp) = datGeneral.GetTableName(CInt(mvarColDetails(14, intTemp)))

		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(0, intTemp) = "?ID"

		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(16, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(16, intTemp) = "ID"

		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(17, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(17, intTemp) = False
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(18, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(18, intTemp) = False

		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(19, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(19, intTemp) = True

		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(20, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(20, intTemp) = IsReportChildTable(lngTableID) 'Indicates if column is a report child table.

		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(21, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(21, intTemp) = True

		mvarColDetails(23, intTemp) = "0"


		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(24, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(24, intTemp) = False	'Group With Next Column.

		Dim iChildCount As Short
		Dim lngChildTableID As Integer
		If miChildTablesCount > 0 Then
			For iChildCount = 0 To UBound(mvarChildTables, 2) Step 1
				'TM20020409 Fault 3745 - only add the ID columns for tables that are actually used.
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(0, iChildCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngChildTableID = mvarChildTables(0, iChildCount)
				If IsChildTableUsed(lngChildTableID) Then
					intTemp = UBound(mvarColDetails, 2) + 2
					ReDim Preserve mvarColDetails(UBound(mvarColDetails, 1), intTemp)

					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(1, intTemp - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(1, intTemp - 1) = 99
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, intTemp - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(2, intTemp - 1) = 0
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(3, intTemp - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(3, intTemp - 1) = False
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(4, intTemp - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(4, intTemp - 1) = False
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(5, intTemp - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(5, intTemp - 1) = False
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(6, intTemp - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(6, intTemp - 1) = False
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(7, intTemp - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(7, intTemp - 1) = False
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(8, intTemp - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(8, intTemp - 1) = False
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(9, intTemp - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(9, intTemp - 1) = False
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(10, intTemp - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(10, intTemp - 1) = False
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(11, intTemp - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(11, intTemp - 1) = ""
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(12, intTemp - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(12, intTemp - 1) = -1
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(13, intTemp - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(13, intTemp - 1) = "C"
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(0, iChildCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(14, intTemp - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(14, intTemp - 1) = mvarChildTables(0, iChildCount)
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(3, iChildCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(15, intTemp - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(15, intTemp - 1) = mvarChildTables(3, iChildCount)
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, intTemp - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(0, intTemp - 1) = "?ID_" & mvarColDetails(14, intTemp - 1)
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(16, intTemp - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(16, intTemp - 1) = "ID_" & mlngCustomReportsBaseTable
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(17, intTemp - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(17, intTemp - 1) = False
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(18, intTemp - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(18, intTemp - 1) = False
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(19, intTemp - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(19, intTemp - 1) = True
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(20, intTemp - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(20, intTemp - 1) = True 'Indicates if column is a report child table.
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(21, intTemp - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(21, intTemp - 1) = True

					mvarColDetails(23, intTemp) = "0"

					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(24, intTemp - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(24, intTemp - 1) = False	'Group With Next Column.

					'*********************************************
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(1, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(1, intTemp) = 99
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(2, intTemp) = 0
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(3, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(3, intTemp) = False
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(4, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(4, intTemp) = False
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(5, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(5, intTemp) = False
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(6, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(6, intTemp) = False
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(7, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(7, intTemp) = False
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(8, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(8, intTemp) = False
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(9, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(9, intTemp) = False
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(10, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(10, intTemp) = False
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(11, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(11, intTemp) = ""
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(12, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(12, intTemp) = -1
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(13, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(13, intTemp) = "C"
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(0, iChildCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(14, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(14, intTemp) = mvarChildTables(0, iChildCount)
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(3, iChildCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(15, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(15, intTemp) = mvarChildTables(3, iChildCount)
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(0, intTemp) = "?ID_" & mvarColDetails(15, intTemp)
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(16, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(16, intTemp) = "ID"
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(17, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(17, intTemp) = False
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(18, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(18, intTemp) = False
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(19, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(19, intTemp) = True
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(20, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(20, intTemp) = True 'Indicates if column is a report child table.
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(21, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(21, intTemp) = True

					mvarColDetails(23, intTemp) = "0"

					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(24, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(24, intTemp) = False	'Group With Next Column.

				End If
			Next iChildCount
		End If

		If miChildTablesCount > 1 Then
			intTemp = UBound(mvarColDetails, 2) + 1
			ReDim Preserve mvarColDetails(UBound(mvarColDetails, 1), intTemp)

			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(1, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(1, intTemp) = 99
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(2, intTemp) = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(3, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(3, intTemp) = True
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(4, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(4, intTemp) = False
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(5, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(5, intTemp) = False
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(6, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(6, intTemp) = False
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(7, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(7, intTemp) = False
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(8, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(8, intTemp) = False
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(9, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(9, intTemp) = False
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(10, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(10, intTemp) = False
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(11, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(11, intTemp) = ""
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(12, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(12, intTemp) = -1
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(13, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(13, intTemp) = "C"
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(14, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(14, intTemp) = -1
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(15, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(15, intTemp) = ""
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(0, intTemp) = lng_SEQUENCECOLUMNNAME
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(16, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(16, intTemp) = ""
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(17, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(17, intTemp) = False
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(18, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(18, intTemp) = False
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(19, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(19, intTemp) = True
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(20, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(20, intTemp) = True 'Indicates if column is a report child table.
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(21, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(21, intTemp) = True

			mvarColDetails(23, intTemp) = "0"

			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(24, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarColDetails(24, intTemp) = False	'Group With Next Column.

		End If

		'******************************************************************************

		' Get those columns defined as a SortOrder and load into array

		strTempSQL = "SELECT * FROM ASRSysCustomReportsDetails WHERE " & "CustomReportID = " & mlngCustomReportID & " " & "AND SortOrderSequence > 0 " & "ORDER BY [SortOrderSequence]"
		prstCustomReportsSortOrder = mclsGeneral.GetReadOnlyRecords(strTempSQL)

		With prstCustomReportsSortOrder
			If .BOF And .EOF Then
				GetDetailsRecordsets = False
				mstrErrorString = "No columns have been defined as a sort order for the specified Custom Report definition." & vbNewLine & "Please remove this definition and create a new one."
				Exit Function
			End If
			Do Until .EOF
				intTemp = UBound(mvarSortOrder, 2) + 1
				ReDim Preserve mvarSortOrder(2, intTemp)
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(0, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarSortOrder(0, intTemp) = GetTableIDFromColumn(.Fields("ColExprID").Value)
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(1, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarSortOrder(1, intTemp) = .Fields("ColExprID").Value
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(2, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarSortOrder(2, intTemp) = .Fields("SortOrder").Value
				.MoveNext()
			Loop
		End With

		'UPGRADE_NOTE: Object prstCustomReportsSortOrder may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		prstCustomReportsSortOrder = Nothing

		GetDetailsRecordsets = True
		Exit Function

GetDetailsRecordsets_ERROR:

		GetDetailsRecordsets = False
		mstrErrorString = "Error retrieving the details recordsets'." & vbNewLine & Err.Description
		mobjEventLog.AddDetailEntry(mstrErrorString)
		mobjEventLog.ChangeHeaderStatus(EventLog_Status.elsFailed)

	End Function

	Private Function IsChildTableUsed(ByRef iChildTableID As Integer) As Boolean

		Dim i As Short

		IsChildTableUsed = False

		For i = 1 To UBound(mvarColDetails, 2) Step 1
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(14, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If mvarColDetails(14, i) = iChildTableID Then
				IsChildTableUsed = True
				Exit Function
			End If
		Next i

	End Function

	Public Function GenerateSQL() As Boolean

		' Purpose : This function calls the individual functions that
		'           generate the components of the main SQL string.

		Dim fOK As Boolean

		fOK = True

		If fOK Then fOK = GenerateSQLSelect()
		If fOK Then fOK = GenerateSQLFrom()
		If fOK Then fOK = GenerateSQLJoin()
		If fOK Then fOK = GenerateSQLWhere()
		If fOK Then fOK = GenerateSQLOrderBy()

		If fOK Then
			GenerateSQL = True
		Else
			GenerateSQL = False
		End If

	End Function

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
		Dim alngSourceTables(,) As Integer
		Dim objCalcExpr As clsExprExpression
		Dim objTableView As CTablePrivilege

		' Set flags with their starting values
		pblnOK = True
		pblnNoSelect = False

		ReDim mastrUDFsRequired(0)

		' JPD20030219 Fault 5068
		' Check the user has permission to read the base table.
		pblnOK = False
		For Each objTableView In gcoTablePrivileges.Collection
			If (objTableView.TableID = mlngCustomReportsBaseTable) And (objTableView.AllowSelect) Then
				pblnOK = True
				Exit For
			End If
		Next objTableView
		'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objTableView = Nothing

		If Not pblnOK Then
			GenerateSQLSelect = False
			mstrErrorString = "You do not have permission to read the base table" & vbNewLine & "either directly or through any views."
			Exit Function
		End If

		' Start off the select statement
		mstrSQLSelect = "SELECT TOP 1000000000000 "

		' Dimension an array of tables/views joined to the base table/view
		' Column 1 = 0 if this row is for a table, 1 if it is for a view
		' Column 2 = table/view ID
		' (should contain everything which needs to be joined to the base tbl/view)
		ReDim mlngTableViews(2, 0)

		' Loop thru the columns collection creating the SELECT and JOIN code
		For pintLoop = 1 To UBound(mvarColDetails, 2)

			' Clear temp vars
			plngTempTableID = 0
			pstrTempTableName = vbNullString
			pstrTempColumnName = vbNullString

			' If its a COLUMN then...
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(13, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If mvarColDetails(13, pintLoop) = "C" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If mvarColDetails(0, pintLoop) <> lng_SEQUENCECOLUMNNAME Then
					' Load the temp variables
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(14, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					plngTempTableID = mvarColDetails(14, pintLoop)
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(15, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					pstrTempTableName = mvarColDetails(15, pintLoop)
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(16, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					pstrTempColumnName = mvarColDetails(16, pintLoop)

					' Check permission on that column
					mobjColumnPrivileges = GetColumnPrivileges(pstrTempTableName)
					mstrRealSource = gcoTablePrivileges.Item(pstrTempTableName).RealSource

					If mbIsBradfordIndexReport Then
						If plngTempTableID <> mlngCustomReportsBaseTable Then
							mstrAbsenceRealSource = mstrRealSource
						End If
					End If

					pblnColumnOK = mobjColumnPrivileges.IsValid(pstrTempColumnName)

					If pblnColumnOK Then
						pblnColumnOK = mobjColumnPrivileges.Item(pstrTempColumnName).AllowSelect
					End If

					If pblnColumnOK Then

						' this column can be read direct from the tbl/view or from a parent table
						' JDM - 16/05/2005 - Fault 10018 - Pad out the duration field because it may not be long enough
						If mbIsBradfordIndexReport And (pintLoop = 12 Or pintLoop = 13) Then
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ",", "") & "convert(numeric(10,2)," & mstrRealSource & "." & Trim(pstrTempColumnName) & ")" & " AS [" & mvarColDetails(0, pintLoop) & "]"
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ",", "") & mstrRealSource & "." & Trim(pstrTempColumnName) & " AS [" & mvarColDetails(0, pintLoop) & "]"
						End If


						' If the table isnt the base table (or its realsource) then
						' Check if it has already been added to the array. If not, add it.
						If plngTempTableID <> mlngCustomReportsBaseTable Then
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
							pstrColumnCode = ""
							For pintNextIndex = 1 To UBound(mstrViews)
								If pintNextIndex = 1 Then
									pstrColumnCode = "CASE"
								End If

								pstrColumnCode = pstrColumnCode & " WHEN NOT " & mstrViews(pintNextIndex) & "." & pstrTempColumnName & " IS NULL THEN " & mstrViews(pintNextIndex) & "." & pstrTempColumnName

							Next pintNextIndex

							If Len(pstrColumnCode) > 0 Then
								'            pstrColumnCode = pstrColumnCode & _
								'" ELSE NULL" & _
								'" END AS '" & mvarColDetails(0, pintLoop) & "'"
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								pstrColumnCode = pstrColumnCode & " ELSE NULL" & " END AS [" & mvarColDetails(0, pintLoop) & "]"

								pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ",", "") & pstrColumnCode
							End If

						End If

						' If we cant see a column, then get outta here
						If pblnNoSelect Then
							GenerateSQLSelect = False
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mstrErrorString = vbNewLine & vbNewLine & "You do not have permission to see the column '" & mvarColDetails(16, pintLoop) & "'" & vbNewLine & "either directly or through any views."
							Exit Function
						End If


						If Not pblnOK Then
							GenerateSQLSelect = False
							Exit Function
						End If

					End If
				Else
					'Add the column which can store the sequence the records are added to the Temp table
					'when more than one Child table is selected.
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ",", "") & 0 & " AS [" & mvarColDetails(0, pintLoop) & "] "

				End If

			Else

				' UH OH ! Its an expression rather than a column

				' Get the calculation SQL, and the array of tables/views that are used to create it.
				' Column 1 = 0 if this row is for a table, 1 if it is for a view.
				' Column 2 = table/view ID.
				ReDim alngSourceTables(2, 0)
				objCalcExpr = New clsExprExpression
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				blnOK = objCalcExpr.Initialise(mlngCustomReportsBaseTable, CInt(mvarColDetails(12, pintLoop)), ExpressionTypes.giEXPR_RUNTIMECALCULATION, ExpressionValueTypes.giEXPRVALUE_UNDEFINED)
				If blnOK Then
					blnOK = objCalcExpr.RuntimeCalculationCode(alngSourceTables, sCalcCode, True, False, mvarPrompts)

					If blnOK Then
						blnOK = objCalcExpr.UDFCalculationCode(alngSourceTables, mastrUDFsRequired, True)
					End If

				End If

				'TM20030422 Fault 5244 - The "SELECT ... INTO..." statement errors when it trys to create a column for
				'and empty string. Therefore wrap this empty sting in a CONVERT(varchar... clause if an sql empty string
				'is returned.
				'TM20030521 Fault 5702 - Compare the empty string with the calc code value converted to varchar
				sCalcCode = "CASE WHEN CONVERT(varchar," & sCalcCode & ") = '' " & "THEN CONVERT(varchar," & sCalcCode & ") " & "ELSE " & sCalcCode & " END"

				'**************************************************************************
				'TM20020730 Fault 4253
				'
				'If there are no Table/View IDs returned in the alngSourceTables array and
				'the RuntimeCalculation code returned successfully (i.e. True) then the
				'current user can see all columns required by the calc on the CALC's basetable,
				'therefore must add the CALC'S BaseTableID to the mlngTableViews array so it
				'can be added to the SQLs Join code.
				'
				'NOTE: The above only applies to the REPORT'S parent tables 1 & 2 as the
				'expression code does not return the calc's BaseTableID in the alngSourceTables
				'array.
				'**************************************************************************

				If mlngCustomReportsParent1Table > 0 Or mlngCustomReportsParent2Table > 0 Then
					If blnOK Then
						If objCalcExpr.BaseTableID = mlngCustomReportsParent1Table Or objCalcExpr.BaseTableID = mlngCustomReportsParent2Table Then
							' Check if table has already been added to the array
							pblnFound = False
							For pintNextIndex = 1 To UBound(mlngTableViews, 2)
								If mlngTableViews(1, pintNextIndex) = 0 And mlngTableViews(2, pintNextIndex) = objCalcExpr.BaseTableID Then
									pblnFound = True
									Exit For
								End If
							Next pintNextIndex

							If Not pblnFound Then
								' View hasnt yet been added, so add it !
								pintNextIndex = UBound(mlngTableViews, 2) + 1
								ReDim Preserve mlngTableViews(2, pintNextIndex)
								mlngTableViews(1, pintNextIndex) = 0
								mlngTableViews(2, pintNextIndex) = objCalcExpr.BaseTableID
							End If
						End If
					End If
				End If

				'**************************************************************************

				'UPGRADE_NOTE: Object objCalcExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				objCalcExpr = Nothing

				If blnOK Then
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ",", "") & sCalcCode & " AS [" & mvarColDetails(0, pintLoop) & "]"

					' Add the required views to the JOIN code.
					For iLoop1 = 1 To UBound(alngSourceTables, 2)
						If alngSourceTables(1, iLoop1) = 1 Then
							' Check if view has already been added to the array
							pblnFound = False
							For pintNextIndex = 1 To UBound(mlngTableViews, 2)
								If mlngTableViews(1, pintNextIndex) = 1 And mlngTableViews(2, pintNextIndex) = alngSourceTables(2, iLoop1) Then
									pblnFound = True
									Exit For
								End If
							Next pintNextIndex

							If Not pblnFound Then

								' View hasnt yet been added, so add it !
								pintNextIndex = UBound(mlngTableViews, 2) + 1
								ReDim Preserve mlngTableViews(2, pintNextIndex)
								mlngTableViews(1, pintNextIndex) = 1
								mlngTableViews(2, pintNextIndex) = alngSourceTables(2, iLoop1)

							End If
							'********************************************************************************
						ElseIf alngSourceTables(1, iLoop1) = 0 Then
							' Check if table has already been added to the array
							pblnFound = False
							For pintNextIndex = 1 To UBound(mlngTableViews, 2)
								If mlngTableViews(1, pintNextIndex) = 0 And mlngTableViews(2, pintNextIndex) = alngSourceTables(2, iLoop1) Then
									pblnFound = True
									Exit For
								End If
							Next pintNextIndex

							' JPD20020514 Fault 3883 - Only want to check if the source table is the base table
							' if we have NOT just found the source table in the array of joined tables.
							If Not pblnFound Then
								pblnFound = (alngSourceTables(2, iLoop1) = mlngCustomReportsBaseTable)
							End If

							If Not pblnFound Then
								' table hasnt yet been added, so add it !
								pintNextIndex = UBound(mlngTableViews, 2) + 1
								ReDim Preserve mlngTableViews(2, pintNextIndex)
								mlngTableViews(1, pintNextIndex) = 0
								mlngTableViews(2, pintNextIndex) = alngSourceTables(2, iLoop1)
							End If
							'********************************************************************************
						End If
					Next iLoop1
				Else
					' Permission denied on something in the calculation.
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mstrErrorString = "You do not have permission to use the '" & mvarColDetails(0, pintLoop) & "' calculation."
					GenerateSQLSelect = False
					Exit Function
				End If

			End If

		Next pintLoop

		mstrSQLSelect = mstrSQLSelect & pstrColumnList

		GenerateSQLSelect = True

		Exit Function

GenerateSQLSelect_ERROR:

		GenerateSQLSelect = False
		mstrErrorString = "Error generating SQL Select statement." & vbNewLine & Err.Description
		mobjEventLog.AddDetailEntry(mstrErrorString)
		mobjEventLog.ChangeHeaderStatus(EventLog_Status.elsFailed)

	End Function

	Private Function IsReportChildTable(ByRef lngTableID As Integer) As Boolean

		Dim i As Short

		IsReportChildTable = False

		If miChildTablesCount > 0 Then
			For i = 0 To UBound(mvarChildTables, 2) Step 1
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(0, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If lngTableID = mvarChildTables(0, i) Then
					IsReportChildTable = True
					Exit Function
				End If
			Next i
		End If

	End Function

	Private Function GetMostChildsForParent(ByRef avChildRecs(,) As Recordset, ByRef iParentCount As Short) As Short

		Dim i As Short
		Dim iMostChildRecords As Short
		Dim iChildRecordCount As Short

		On Error GoTo Error_Trap

		iMostChildRecords = 0
		iChildRecordCount = 0

		For i = 0 To UBound(avChildRecs, 2) Step 1
			If (avChildRecs(iParentCount, i).BOF) And (avChildRecs(iParentCount, i).EOF) Then
				iChildRecordCount = 0
			Else
				iChildRecordCount = avChildRecs(iParentCount, i).RecordCount
			End If
			If iChildRecordCount > iMostChildRecords Then
				iMostChildRecords = iChildRecordCount
			End If
		Next i

		GetMostChildsForParent = iMostChildRecords

		Exit Function

Error_Trap:
		GetMostChildsForParent = 0

	End Function

	Private Function OrderBy(ByRef plngTableID As Integer) As Object

		' This function creates an ORDER BY statement by searching
		' through the columns defined as the reports sort order, then
		' uses the relevant alias name

		Dim iColCount As Short
		Dim iSortCount As Short
		Dim bHasOrder As Boolean

		bHasOrder = False

		For iSortCount = 1 To UBound(mvarSortOrder, 2)
			For iColCount = 1 To UBound(mvarColDetails, 2)

				'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(0, iSortCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(12, iColCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(1, iSortCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If mvarSortOrder(1, iSortCount) = mvarColDetails(12, iColCount) And mvarSortOrder(0, iSortCount) = plngTableID Then
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(2, iSortCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object OrderBy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					OrderBy = OrderBy & "[" & mvarColDetails(0, iColCount) & "] " & mvarSortOrder(2, iSortCount)
					'UPGRADE_WARNING: Couldn't resolve default property of object OrderBy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					OrderBy = OrderBy & ", "
					bHasOrder = True
					Exit For
				End If
			Next iColCount
		Next iSortCount

		If bHasOrder Then
			'UPGRADE_WARNING: Couldn't resolve default property of object OrderBy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			OrderBy = " ORDER BY " & Left(OrderBy, Len(OrderBy) - 2) & " "
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object OrderBy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			OrderBy = vbNullString
		End If

	End Function

	Public Function CreateMutipleChildTempTable() As Boolean

		Dim sMCTempTable As String
		Dim sSQL As String
		Dim iColCount As Short
		Dim sParentSelectSQL As String
		Dim rsParent As Recordset
		Dim lngColumnID As Integer
		Dim lngTableID As Integer
		Dim iChildCount As Short
		Dim rsChild As Recordset
		Dim iParentCount As Short
		Dim avChildRecordsets(,) As Recordset
		Dim sChildSelectSQL As String
		Dim sChildWhereSQL As String
		Dim iFields As Short
		Dim i As Short
		Dim iChildUsed As Short
		Dim iMostChilds As Short
		Dim lngCurrentTableID As Integer
		Dim lngSequenceCount As Integer

		Dim sFIELDS As String
		Dim sVALUES As String
		Dim SQLSTRING As String

		On Error GoTo Error_Trap

		'******************* Create multiple child temp table ***************************
		sMCTempTable = datGeneral.UniqueSQLObjectName("ASRSysTempCustomReport", 3)

		sSQL = "SELECT * INTO [" & sMCTempTable & "] FROM [" & mstrTempTableName & "]"
		mclsData.ExecuteSql(sSQL)

		sSQL = "DELETE FROM [" & sMCTempTable & "]"
		mclsData.ExecuteSql(sSQL)


		'************** Get the Parent SELECT SQL statment ******************************
		For iColCount = 1 To UBound(mvarColDetails, 2) Step 1
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(14, iColCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			lngTableID = mvarColDetails(14, iColCount)
			If IsReportParentTable(lngTableID) Or IsReportBaseTable(lngTableID) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sParentSelectSQL = sParentSelectSQL & "[" & mvarColDetails(0, iColCount) & "]"
				sParentSelectSQL = sParentSelectSQL & ", "
			End If
		Next iColCount

		sParentSelectSQL = Left(sParentSelectSQL, Len(sParentSelectSQL) - 2) & " "

		sSQL = "SELECT DISTINCT " & sParentSelectSQL
		sSQL = sSQL & " FROM [" & mstrTempTableName & "] "


		'Order the Parent recorset
		'UPGRADE_WARNING: Couldn't resolve default property of object OrderBy(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sSQL = sSQL & OrderBy(mlngCustomReportsBaseTable)

		rsParent = datGeneral.GetRecords(sSQL)

		lngColumnID = 0
		lngTableID = 0
		iChildUsed = 0

		'*************** Circle through the distinct list of parent records *************
		With rsParent

			'TM20020802 Fault 4273
			If (.BOF And .EOF) Then
				mstrErrorString = "No records meet selection criteria"
				CreateMutipleChildTempTable = False
				mobjEventLog.AddDetailEntry("Completed successfully. " & mstrErrorString)
				mobjEventLog.ChangeHeaderStatus(EventLog_Status.elsSuccessful)
				mblnNoRecords = True

				sMCTempTable = vbNullString
				'      Set rsTemp = Nothing
				'UPGRADE_NOTE: Object rsParent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				rsParent = Nothing
				'UPGRADE_NOTE: Object rsChild may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				rsChild = Nothing
				Exit Function
			End If

			.MoveFirst()
			iParentCount = 0
			lngSequenceCount = 1

			mbUseSequence = True

			Do Until .EOF

				iParentCount = iParentCount + 1

				ReDim avChildRecordsets(0, miUsedChildCount - 1)
				For iChildCount = 0 To UBound(mvarChildTables, 2) Step 1
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(0, iChildCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					lngCurrentTableID = mvarChildTables(0, iChildCount)
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(4, iChildCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If mvarChildTables(4, iChildCount) Then	'is the child table used???
						'UPGRADE_NOTE: Object rsChild may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
						rsChild = Nothing
						For iColCount = 1 To UBound(mvarColDetails, 2) Step 1
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(14, iColCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							lngTableID = mvarColDetails(14, iColCount)
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(16, iColCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(0, iChildCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If (mvarColDetails(20, iColCount)) And (lngTableID = mvarChildTables(0, iChildCount)) And (mvarColDetails(16, iColCount) <> ("?ID_" & CStr(mlngCustomReportsBaseTable))) Then
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sChildSelectSQL = sChildSelectSQL & "[" & mvarColDetails(0, iColCount) & "]"
								sChildSelectSQL = sChildSelectSQL & ", "
							End If
						Next iColCount
						sChildSelectSQL = Left(sChildSelectSQL, Len(sChildSelectSQL) - 2) & " "

						'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sChildWhereSQL = sChildWhereSQL & "[?ID_" & mvarChildTables(0, iChildCount) & "] = "
						sChildWhereSQL = sChildWhereSQL & .Fields("?ID").Value

						sSQL = "SELECT DISTINCT " & sChildSelectSQL
						sSQL = sSQL & " FROM [" & mstrTempTableName & "]"
						sSQL = sSQL & " WHERE " & sChildWhereSQL

						'Order the child recordset.
						'UPGRADE_WARNING: Couldn't resolve default property of object OrderBy(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sSQL = sSQL & OrderBy(lngCurrentTableID)

						sChildSelectSQL = vbNullString
						sChildWhereSQL = vbNullString

						rsChild = datGeneral.GetRecords(sSQL)

						'Add the child tables recordset to the array of child tables.
						avChildRecordsets(0, iChildUsed) = rsChild
						iChildUsed = iChildUsed + 1
					End If
				Next iChildCount

				'      With rsTemp
				iMostChilds = GetMostChildsForParent(avChildRecordsets, 0)
				If iMostChilds > 0 Then
					For i = 0 To iMostChilds - 1 Step 1
						'            .AddNew

						sFIELDS = vbNullString
						sVALUES = vbNullString
						SQLSTRING = vbNullString

						'<<<<<<<<<<<<<<<<<<< Add Values To Parent Fields >>>>>>>>>>>>>>>>>>>>>>>
						For iFields = 0 To rsParent.Fields.Count - 1 Step 1
							'              .Fields(rsParent.Fields(iFields).Name) = rsParent.Fields(iFields).Value

							sFIELDS = sFIELDS & "[" & rsParent.Fields(iFields).Name & "],"

							Select Case rsParent.Fields(iFields).Type
								Case DataTypeEnum.adNumeric, DataTypeEnum.adInteger, DataTypeEnum.adSingle, DataTypeEnum.adDouble
									'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
									sVALUES = sVALUES & IIf(IsDBNull(rsParent.Fields(iFields).Value), 0, rsParent.Fields(iFields).Value) & ","
								Case DataTypeEnum.adDBTimeStamp, DataTypeEnum.adDate, DataTypeEnum.adDBDate, DataTypeEnum.adDBTime
									'TM20030124 Fault 4974
									'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
									If Not IsDBNull(rsParent.Fields(iFields).Value) Then
										sVALUES = sVALUES & "'" & VB6.Format(rsParent.Fields(iFields).Value, "MM/dd/yyyy") & "',"
									Else
										sVALUES = sVALUES & "NULL,"
									End If
								Case DataTypeEnum.adBoolean
									sVALUES = sVALUES & IIf(rsParent.Fields(iFields).Value, 1, 0) & ","
								Case Else
									'MH20021119 Fault 4315
									'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
									If Not IsDBNull(rsParent.Fields(iFields).Value) Then
										sVALUES = sVALUES & "'" & Replace(rsParent.Fields(iFields).Value, "'", "''") & "',"
									Else
										sVALUES = sVALUES & "'',"
									End If
							End Select

						Next iFields

						For iChildCount = 0 To UBound(avChildRecordsets, 2) Step 1
							If Not avChildRecordsets(0, iChildCount).EOF Then
								'<<<<<<<<<<<<<<<<<<< Add Values To Child Fields >>>>>>>>>>>>>>>>>>>>>>>
								For iFields = 0 To avChildRecordsets(0, iChildCount).Fields.Count - 1 Step 1
									'                  .Fields(avChildRecordsets(0, iChildCount).Fields(iFields).Name) = avChildRecordsets(0, iChildCount).Fields(iFields).Value

									sFIELDS = sFIELDS & "[" & avChildRecordsets(0, iChildCount).Fields(iFields).Name & "],"

									Select Case avChildRecordsets(0, iChildCount).Fields(iFields).Type
										Case DataTypeEnum.adNumeric, DataTypeEnum.adInteger, DataTypeEnum.adSingle, DataTypeEnum.adDouble
											'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
											sVALUES = sVALUES & IIf(IsDBNull(avChildRecordsets(0, iChildCount).Fields(iFields).Value), 0, avChildRecordsets(0, iChildCount).Fields(iFields).Value) & ","
										Case DataTypeEnum.adDBTimeStamp, DataTypeEnum.adDate, DataTypeEnum.adDBDate, DataTypeEnum.adDBTime
											'TM20030124 Fault 4974
											'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
											If Not IsDBNull(avChildRecordsets(0, iChildCount).Fields(iFields).Value) Then
												sVALUES = sVALUES & "'" & VB6.Format(avChildRecordsets(0, iChildCount).Fields(iFields).Value, "MM/dd/yyyy") & "',"
											Else
												sVALUES = sVALUES & "NULL,"
											End If
										Case DataTypeEnum.adBoolean
											sVALUES = sVALUES & IIf(avChildRecordsets(0, iChildCount).Fields(iFields).Value, 1, 0) & ","
										Case Else
											'MH20021119 Fault 4315
											'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
											If Not IsDBNull(avChildRecordsets(0, iChildCount).Fields(iFields).Value) Then
												sVALUES = sVALUES & "'" & Replace(avChildRecordsets(0, iChildCount).Fields(iFields).Value, "'", "''") & "',"
											Else
												sVALUES = sVALUES & "'',"
											End If
									End Select

								Next iFields
								avChildRecordsets(0, iChildCount).MoveNext()
							End If
						Next iChildCount

						'Add the Sequence number to the sequence column for ordering the data later.
						'            .Fields(lng_SEQUENCECOLUMNNAME) = lngSequenceCount

						sFIELDS = sFIELDS & "[" & lng_SEQUENCECOLUMNNAME & "]"
						sVALUES = sVALUES & lngSequenceCount

						lngSequenceCount = lngSequenceCount + 1

						SQLSTRING = "INSERT INTO " & sMCTempTable & " (" & sFIELDS & ") "
						SQLSTRING = SQLSTRING & " VALUES (" & sVALUES & ") "

						gADOCon.Execute(SQLSTRING)

						'            .Update
					Next i
				Else
					'          .AddNew

					sFIELDS = vbNullString
					sVALUES = vbNullString
					SQLSTRING = vbNullString

					'<<<<<<<<<<<<<<<<<<< Add Values To Parent Fields >>>>>>>>>>>>>>>>>>>>>>>
					For iFields = 0 To rsParent.Fields.Count - 1 Step 1
						'            .Fields(rsParent.Fields(iFields).Name) = rsParent.Fields(iFields).Value

						sFIELDS = sFIELDS & "[" & rsParent.Fields(iFields).Name & "],"

						Select Case rsParent.Fields(iFields).Type
							Case DataTypeEnum.adNumeric, DataTypeEnum.adInteger, DataTypeEnum.adSingle, DataTypeEnum.adDouble
								'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
								sVALUES = sVALUES & IIf(IsDBNull(rsParent.Fields(iFields).Value), 0, rsParent.Fields(iFields).Value) & ","
							Case DataTypeEnum.adDBTimeStamp, DataTypeEnum.adDate, DataTypeEnum.adDBDate, DataTypeEnum.adDBTime
								'TM20030124 Fault 4974
								'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
								If Not IsDBNull(rsParent.Fields(iFields).Value) Then
									sVALUES = sVALUES & "'" & VB6.Format(rsParent.Fields(iFields).Value, "MM/dd/yyyy") & "',"
								Else
									sVALUES = sVALUES & "NULL,"
								End If
							Case DataTypeEnum.adBoolean
								sVALUES = sVALUES & IIf(rsParent.Fields(iFields).Value, 1, 0) & ","
							Case Else
								'MH20021119 Fault 4315
								'sVALUES = sVALUES & "'" & Replace(rsParent.Fields(iFields).Value, "'", "''") & "',"
								'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
								If Not IsDBNull(rsParent.Fields(iFields).Value) Then
									sVALUES = sVALUES & "'" & Replace(CStr(rsParent.Fields(iFields).Value), "'", "''") & "',"
								Else
									sVALUES = sVALUES & "'',"
								End If
						End Select

					Next iFields

					'Add the Sequence number to the sequence column for ordering the data later.
					'          .Fields(lng_SEQUENCECOLUMNNAME) = lngSequenceCount

					sFIELDS = sFIELDS & "[" & lng_SEQUENCECOLUMNNAME & "]"
					sVALUES = sVALUES & lngSequenceCount

					lngSequenceCount = lngSequenceCount + 1

					SQLSTRING = "INSERT INTO " & sMCTempTable & " (" & sFIELDS & ") "
					SQLSTRING = SQLSTRING & " VALUES (" & sVALUES & ") "

					gADOCon.Execute(SQLSTRING)

					'          .Update
				End If
				'      End With
				.MoveNext()
				iChildUsed = 0
			Loop
		End With

		'************ Re-Order the data using the defined sort orders. ******************
		sSQL = "DELETE FROM [" & mstrTempTableName & "]"
		mclsData.ExecuteSql(sSQL)

		sSQL = "INSERT INTO [" & mstrTempTableName & "] SELECT * FROM [" & sMCTempTable & "]"
		' Order the entire recordset.
		sSQL = sSQL & " ORDER BY [" & lng_SEQUENCECOLUMNNAME & "] ASC"
		mclsData.ExecuteSql(sSQL)


		'***************** Drop the multiple child temp table. **************************
		' Delete the temptable if exists, and then clear the variable
		'  If Len(sMCTempTable) > 0 Then
		'    mclsData.ExecuteSql ("IF EXISTS(SELECT * FROM sysobjects WHERE name = '" & sMCTempTable & "') " & _
		''                      "DROP TABLE [" & sMCTempTable & "]")
		'  End If
		datGeneral.DropUniqueSQLObject(sMCTempTable, 3)
		sMCTempTable = vbNullString


		'************ Drop the ID columns from the temp table. ******************
		'  With rsTemp
		'    'Remove the ".ID" & "ID" columns from the report.
		'    For iColCount = 1 To UBound(mvarColDetails, 2) Step 1
		'      If (mvarColDetails(16, iColCount) = "ID") Or (mvarColDetails(16, iColCount) = ("ID_" & CStr(mlngCustomReportsBaseTable))) Then
		'        sSQL = "ALTER TABLE [" & mstrTempTableName & "] DROP COLUMN [" & mvarColDetails(0, iColCount) & "]"
		'        mclsData.ExecuteSql sSQL
		'      End If
		'    Next iColCount
		'    .Close
		'  End With
		'  'remove the id columns from column details array.
		'  ReDim Preserve mvarColDetails(20, miColumnsInReport)


		'********************************************************************************
		CreateMutipleChildTempTable = True

		sMCTempTable = vbNullString
		'  Set rsTemp = Nothing
		'UPGRADE_NOTE: Object rsParent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsParent = Nothing
		'UPGRADE_NOTE: Object rsChild may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsChild = Nothing
		Exit Function

Error_Trap:
		CreateMutipleChildTempTable = False
		sMCTempTable = vbNullString
		'  Set rsTemp = Nothing
		'UPGRADE_NOTE: Object rsParent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsParent = Nothing
		'UPGRADE_NOTE: Object rsChild may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsChild = Nothing

		mstrErrorString = "Error creating temporary table for multiple childs." & vbNewLine & Err.Number & vbNewLine & Err.Description

		For i = 0 To gADOCon.Errors.Count - 1 Step 1
			mstrErrorString = mstrErrorString & "Err.Number = " & gADOCon.Errors(i).Number & " Err.Desc = " & gADOCon.Errors(i).Description
		Next i

		mobjEventLog.AddDetailEntry(mstrErrorString)
		mobjEventLog.ChangeHeaderStatus(EventLog_Status.elsFailed)

	End Function

	Private Function IsReportParentTable(ByRef lngTableID As Integer) As Boolean

		IsReportParentTable = False

		If lngTableID = mlngCustomReportsParent1Table Or lngTableID = mlngCustomReportsParent2Table Then
			IsReportParentTable = True
		End If

	End Function

	Private Function IsReportBaseTable(ByRef lngTableID As Integer) As Boolean

		IsReportBaseTable = False

		If lngTableID = mlngCustomReportsBaseTable Then
			IsReportBaseTable = True
		End If

	End Function

	Private Function GenerateSQLFrom() As Boolean

		mstrSQLFrom = gcoTablePrivileges.Item(mstrCustomReportsBaseTableName).RealSource
		Return True

	End Function

	Private Function GenerateSQLJoin() As Boolean

		' Purpose : Add the join strings for parent/child/views.
		'           Also adds filter clauses to the joins if used

		On Error GoTo GenerateSQLJoin_ERROR

		Dim pobjTableView As CTablePrivilege
		Dim objChildTable As CTablePrivilege
		Dim pintLoop As Short
		Dim sChildJoinCode As String
		Dim sChildOrderString As String
		Dim rsTemp As Recordset
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

		sOtherParentJoinCode = ""

		' First, do the join for all the views etc...

		For pintLoop = 1 To UBound(mlngTableViews, 2)

			' Get the table/view object from the id stored in the array
			If mlngTableViews(1, pintLoop) = 0 Then
				pobjTableView = gcoTablePrivileges.FindTableID(mlngTableViews(2, pintLoop))
			Else
				pobjTableView = gcoTablePrivileges.FindViewID(mlngTableViews(2, pintLoop))
			End If

			' Dont add a join here if its the child table...do that later
			'If pobjTableView.TableID <> mlngCustomReportsChildTable Then
			If Not IsReportChildTable((pobjTableView.TableID)) Then
				If pobjTableView.TableID <> mlngCustomReportsParent1Table Then
					If pobjTableView.TableID <> mlngCustomReportsParent2Table Then

						If (pobjTableView.TableID = mlngCustomReportsBaseTable) Then
							If (pobjTableView.ViewName <> mstrBaseTableRealSource) Then
								mstrSQLJoin = mstrSQLJoin & " LEFT OUTER JOIN " & pobjTableView.RealSource & " ON " & mstrBaseTableRealSource & ".ID = " & pobjTableView.RealSource & ".ID"
							End If
						Else
							'JPD 20031119 Fault 7660
							' This is a parent of a child of the report base table, not explicitly
							' included in the report, but referred to by a child table calculation.
							For iLoop2 = 1 To UBound(mlngTableViews, 2)
								If mlngTableViews(1, iLoop2) = 0 Then
									If mclsGeneral.IsAChildOf(mlngTableViews(2, iLoop2), (pobjTableView.TableID)) Then
										objChildTable = gcoTablePrivileges.FindTableID(mlngTableViews(2, iLoop2))

										sOtherParentJoinCode = sOtherParentJoinCode & " LEFT OUTER JOIN " & pobjTableView.RealSource & " ON " & objChildTable.RealSource & ".ID_" & CStr(pobjTableView.TableID) & " = " & pobjTableView.RealSource & ".ID"
										Exit For
									End If
								End If
							Next iLoop2
						End If
					End If
				End If
			End If

			If (pobjTableView.TableID = mlngCustomReportsParent1Table) Or (pobjTableView.TableID = mlngCustomReportsParent2Table) Then
				mstrSQLJoin = mstrSQLJoin & " LEFT OUTER JOIN " & pobjTableView.RealSource & " ON " & mstrBaseTableRealSource & ".ID_" & pobjTableView.TableID & " = " & pobjTableView.RealSource & ".ID"
			End If
		Next pintLoop

		'Now do the childview(s) bit, if required

		lngTempChildID = 0
		lngTempMaxRecords = 0
		lngTempFilterID = 0

		'  If mlngCustomReportsChildTable > 0 Then
		If miChildTablesCount > 0 Then
			For i = 0 To UBound(mvarChildTables, 2) Step 1
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(0, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngTempChildID = mvarChildTables(0, i)
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(1, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngTempFilterID = mvarChildTables(1, i)
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(5, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngTempOrderID = mvarChildTables(5, i)
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(2, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lngTempMaxRecords = mvarChildTables(2, i)

				' Only do the join if columns from the table are used.
				pblnChildUsed = IsChildTableUsed(lngTempChildID)

				'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(4, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarChildTables(4, i) = pblnChildUsed
				If pblnChildUsed Then miUsedChildCount = miUsedChildCount + 1

				If pblnChildUsed = True Then

					objChildTable = gcoTablePrivileges.FindTableID(lngTempChildID)

					If objChildTable.AllowSelect Then
						sChildJoinCode = sChildJoinCode & " LEFT OUTER JOIN " & objChildTable.RealSource & " ON " & mstrBaseTableRealSource & ".ID = " & objChildTable.RealSource & ".ID_" & mlngCustomReportsBaseTable

						sChildJoinCode = sChildJoinCode & " AND " & objChildTable.RealSource & ".ID IN"

						'TM20020328 Fault 3714 - ensure the maxrecords is >= zero.
						sChildJoinCode = sChildJoinCode & " (SELECT TOP" & IIf(lngTempMaxRecords < 1, " 100 PERCENT", " " & lngTempMaxRecords) & " " & objChildTable.RealSource & ".ID FROM " & objChildTable.RealSource

						' Now the child order by bit - done here in case tables need to be joined.
						'          Set rsTemp = datGeneral.GetOrderDefinition(datGeneral.GetDefaultOrder(mlngCustomReportsChildTable))
						If lngTempOrderID > 0 Then
							rsTemp = datGeneral.GetOrderDefinition(lngTempOrderID)
						Else
							rsTemp = datGeneral.GetOrderDefinition(datGeneral.GetDefaultOrder(lngTempChildID))
						End If

						sChildOrderString = DoChildOrderString(rsTemp, sChildJoin, lngTempChildID)
						'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
						rsTemp = Nothing

						sChildJoinCode = sChildJoinCode & sChildJoin

						sChildJoinCode = sChildJoinCode & " WHERE (" & objChildTable.RealSource & ".ID_" & mlngCustomReportsBaseTable & " = " & mstrBaseTableRealSource & ".ID)"

						' is the child filtered ?

						If lngTempFilterID > 0 Then
							blnOK = datGeneral.FilteredIDs(lngTempFilterID, strFilterIDs, mvarPrompts)

							' Generate any UDFs that are used in this filter
							If blnOK Then
								datGeneral.FilterUDFs(lngTempFilterID, mastrUDFsRequired)
							End If

							If blnOK Then
								sChildJoinCode = sChildJoinCode & " AND " & objChildTable.RealSource & ".ID IN (" & strFilterIDs & ")"
							Else
								' Permission denied on something in the filter.
								mstrErrorString = "You do not have permission to use the '" & datGeneral.GetFilterName(lngTempFilterID) & "' filter."
								GenerateSQLJoin = False
								Exit Function
							End If
						End If

					End If

					sChildJoinCode = sChildJoinCode & IIf(Len(sChildOrderString) > 0, " ORDER BY " & sChildOrderString & ")", "")

				End If
			Next i
		End If

		mstrSQLJoin = mstrSQLJoin & sChildJoinCode
		mstrSQLJoin = mstrSQLJoin & sOtherParentJoinCode

		GenerateSQLJoin = True
		Exit Function

GenerateSQLJoin_ERROR:

		GenerateSQLJoin = False
		mstrErrorString = "Error in GenerateSQLJoin." & vbNewLine & Err.Description
		mobjEventLog.AddDetailEntry(mstrErrorString)
		mobjEventLog.ChangeHeaderStatus(EventLog_Status.elsFailed)

	End Function

	Private Function DoChildOrderString(ByRef rsTemp As Recordset, ByRef psJoinCode As String, ByRef plngChildID As Integer) As String

		' This function loops through the child tables default order
		' checking if the user has privileges. If they do, add to the order string
		' if not, leave it out.

		On Error GoTo DoChildOrderString_ERROR

		Dim fColumnOK As Boolean
		Dim fFound As Boolean
		Dim iNextIndex As Short
		Dim sSource As String
		Dim sRealSource As String
		Dim sColumnCode As String
		Dim sCurrentTableViewName As String
		Dim objColumnPrivileges As CColumnPrivileges
		Dim pobjOrderCol As CTablePrivilege
		Dim objTableView As CTablePrivilege
		Dim alngTableViews(,) As Integer
		Dim asViews() As String
		Dim iTempCounter As Short

		' Dimension an array of tables/views joined to the base table/view.
		' Column 1 = 0 if this row is for a table, 1 if it is for a view.
		' Column 2 = table/view ID.
		ReDim alngTableViews(2, 0)

		pobjOrderCol = gcoTablePrivileges.FindTableID(plngChildID)
		sCurrentTableViewName = pobjOrderCol.RealSource
		'UPGRADE_NOTE: Object pobjOrderCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pobjOrderCol = Nothing

		Do Until rsTemp.EOF
			If rsTemp.Fields("Type").Value = "O" Then
				' Check if the user can read the column.
				pobjOrderCol = gcoTablePrivileges.FindTableID(rsTemp.Fields("TableID").Value)
				objColumnPrivileges = GetColumnPrivileges((pobjOrderCol.TableName))
				fColumnOK = objColumnPrivileges.Item(rsTemp.Fields("ColumnName").Value).AllowSelect
				'UPGRADE_NOTE: Object objColumnPrivileges may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				objColumnPrivileges = Nothing

				If fColumnOK Then
					'        If rsTemp!TableID = mlngCustomReportsChildTable Then
					If rsTemp.Fields("TableID").Value = plngChildID Then
						DoChildOrderString = DoChildOrderString & IIf(Len(DoChildOrderString) > 0, ",", "") & pobjOrderCol.RealSource & "." & rsTemp.Fields("ColumnName").Value & IIf(rsTemp.Fields("Ascending").Value, "", " DESC")
					Else
						' If the column comes from a parent table, then add the table to the Join code.
						' Check if the table has already been added to the join code.
						fFound = False
						iTempCounter = 0
						For iNextIndex = 1 To UBound(alngTableViews, 2)
							If alngTableViews(1, iNextIndex) = 0 And alngTableViews(2, iNextIndex) = rsTemp.Fields("TableID").Value Then
								iTempCounter = iNextIndex
								fFound = True
								Exit For
							End If
						Next iNextIndex

						If Not fFound Then
							' The table has not yet been added to the join code, so add it to the array and the join code.
							iNextIndex = UBound(alngTableViews, 2) + 1
							ReDim Preserve alngTableViews(2, iNextIndex)
							alngTableViews(1, iNextIndex) = 0
							alngTableViews(2, iNextIndex) = rsTemp.Fields("TableID").Value

							iTempCounter = iNextIndex

							psJoinCode = psJoinCode & " LEFT OUTER JOIN " & pobjOrderCol.RealSource & " ASRSysTemp_" & Trim(Str(iTempCounter)) & " ON " & sCurrentTableViewName & ".ID_" & Trim(Str(rsTemp.Fields("TableID").Value)) & " = ASRSysTemp_" & Trim(Str(iTempCounter)) & ".ID"
						End If

						DoChildOrderString = DoChildOrderString & IIf(Len(DoChildOrderString) > 0, ",", "") & "ASRSysTemp_" & Trim(Str(iTempCounter)) & "." & rsTemp.Fields("ColumnName").Value & IIf(rsTemp.Fields("Ascending").Value, "", " DESC")
					End If
				Else
					' The column cannot be read from the base table/view, or directly from a parent table.
					' If it is a column from a prent table, then try to read it from the views on the parent table.
					'        If rsTemp!TableID <> mlngCustomReportsChildTable Then
					If rsTemp.Fields("TableID").Value <> plngChildID Then
						' Loop through the views on the column's table, seeing if any have 'read' permission granted on them.
						ReDim asViews(0)
						For Each objTableView In gcoTablePrivileges.Collection
							If (Not objTableView.IsTable) And (objTableView.TableID = rsTemp.Fields("TableID").Value) And (objTableView.AllowSelect) Then

								sSource = objTableView.ViewName
								sRealSource = gcoTablePrivileges.Item(sSource).RealSource

								' Get the column permission for the view.
								objColumnPrivileges = GetColumnPrivileges(sSource)

								If objColumnPrivileges.IsValid(rsTemp.Fields("ColumnName").Value) Then
									If objColumnPrivileges.Item(rsTemp.Fields("ColumnName").Value).AllowSelect Then
										' Add the view info to an array to be put into the column list or order code below.
										iNextIndex = UBound(asViews) + 1
										ReDim Preserve asViews(iNextIndex)
										asViews(iNextIndex) = objTableView.ViewName

										' Add the view to the Join code.
										' Check if the view has already been added to the join code.
										fFound = False
										iTempCounter = 0
										For iNextIndex = 1 To UBound(alngTableViews, 2)
											If alngTableViews(1, iNextIndex) = 1 And alngTableViews(2, iNextIndex) = objTableView.ViewID Then
												fFound = True
												iTempCounter = iNextIndex
												Exit For
											End If
										Next iNextIndex

										If Not fFound Then
											' The view has not yet been added to the join code, so add it to the array and the join code.
											iNextIndex = UBound(alngTableViews, 2) + 1
											ReDim Preserve alngTableViews(2, iNextIndex)
											alngTableViews(1, iNextIndex) = 1
											alngTableViews(2, iNextIndex) = objTableView.ViewID

											iTempCounter = iNextIndex

											psJoinCode = psJoinCode & " LEFT OUTER JOIN " & sRealSource & " ASRSysTemp_" & Trim(Str(iTempCounter)) & " ON " & sCurrentTableViewName & ".ID_" & Trim(Str(objTableView.TableID)) & " = ASRSysTemp_" & Trim(Str(iTempCounter)) & ".ID"
										End If
									End If
								End If
								'UPGRADE_NOTE: Object objColumnPrivileges may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
								objColumnPrivileges = Nothing
							End If
						Next objTableView
						'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
						objTableView = Nothing

						' The current user does have permission to 'read' the column through a/some view(s) on the
						' table.
						If UBound(asViews) > 0 Then
							' Add the column to the column list.
							sColumnCode = ""
							For iNextIndex = 1 To UBound(asViews)
								If iNextIndex = 1 Then
									sColumnCode = "CASE "
								End If

								sColumnCode = sColumnCode & " WHEN NOT ASRSysTemp_" & Trim(Str(iNextIndex)) & "." & rsTemp.Fields("ColumnName").Value & " IS NULL THEN ASRSysTemp_" & Trim(Str(iNextIndex)) & "." & rsTemp.Fields("ColumnName").Value
							Next iNextIndex

							If Len(sColumnCode) > 0 Then
								sColumnCode = sColumnCode & " ELSE NULL" & " END"

								' Add the column to the order string.
								DoChildOrderString = DoChildOrderString & IIf(Len(DoChildOrderString) > 0, ", ", "") & sColumnCode & IIf(rsTemp.Fields("Ascending").Value, "", " DESC")
							End If
						End If
					End If
				End If

				'UPGRADE_NOTE: Object pobjOrderCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				pobjOrderCol = Nothing
			End If

			rsTemp.MoveNext()
		Loop

		' JIRA 3180 - Force the ID to be part of the sort order because the UDFs sort by ID too
		DoChildOrderString = DoChildOrderString & "," & sCurrentTableViewName & ".ID"

		Exit Function

DoChildOrderString_ERROR:

		'UPGRADE_NOTE: Object pobjOrderCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pobjOrderCol = Nothing
		mstrErrorString = "Error while generating child order string" & vbNewLine & Err.Description
		DoChildOrderString = ""
		mobjEventLog.AddDetailEntry(mstrErrorString)
		mobjEventLog.ChangeHeaderStatus(EventLog_Status.elsFailed)

	End Function

	Private Function GenerateSQLWhere() As Boolean

		' Purpose : Generate the where clauses that cope with the joins
		'           NB Need to add the where clauses for filters/picklists etc

		On Error GoTo GenerateSQLWhere_ERROR

		Dim pintLoop As Short
		Dim pobjTableView As CTablePrivilege
		Dim prstTemp As New Recordset
		Dim pstrPickListIDs As String
		Dim blnOK As Boolean
		Dim strFilterIDs As String
		Dim pstrParent1PickListIDs As String
		Dim pstrParent2PickListIDs As String

		pobjTableView = gcoTablePrivileges.FindTableID(mlngCustomReportsBaseTable)
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

		' Parent 1 filter
		' Parent 1 filter and picklist
		If mlngCustomReportsParent1PickListID > 0 Then
			pstrParent1PickListIDs = ""
			prstTemp = mclsData.OpenRecordset("EXEC sp_ASRGetPickListRecords " & mlngCustomReportsParent1PickListID, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly)

			If prstTemp.BOF And prstTemp.EOF Then
				mstrErrorString = "The first parent table picklist contains no records."
				GenerateSQLWhere = False
				Exit Function
			End If

			Do While Not prstTemp.EOF
				pstrParent1PickListIDs = pstrParent1PickListIDs & IIf(Len(pstrParent1PickListIDs) > 0, ", ", "") & prstTemp.Fields(0).Value
				prstTemp.MoveNext()
			Loop

			prstTemp.Close()
			'UPGRADE_NOTE: Object prstTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			prstTemp = Nothing

			mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & mstrBaseTableRealSource & ".ID_" & mlngCustomReportsParent1Table & " IN (" & pstrParent1PickListIDs & ") "
		ElseIf mlngCustomReportsParent1FilterID > 0 Then
			blnOK = True
			blnOK = datGeneral.FilteredIDs(mlngCustomReportsParent1FilterID, strFilterIDs, mvarPrompts)

			' Generate any UDFs that are used in this filter
			If blnOK Then
				datGeneral.FilterUDFs(mlngCustomReportsParent1FilterID, mastrUDFsRequired)
			End If

			If blnOK Then
				mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & mstrBaseTableRealSource & ".ID_" & mlngCustomReportsParent1Table & " IN (" & strFilterIDs & ") "
			Else
				mstrErrorString = "You do not have permission to use the '" & datGeneral.GetFilterName(mlngCustomReportsParent1FilterID) & "' filter."
				GenerateSQLWhere = False
				Exit Function
			End If
		End If

		' Parent 2 filter and picklist
		If mlngCustomReportsParent2PickListID > 0 Then
			pstrParent2PickListIDs = ""
			prstTemp = mclsData.OpenRecordset("EXEC sp_ASRGetPickListRecords " & mlngCustomReportsParent2PickListID, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly)

			If prstTemp.BOF And prstTemp.EOF Then
				mstrErrorString = "The second parent table picklist contains no records."
				GenerateSQLWhere = False
				Exit Function
			End If

			Do While Not prstTemp.EOF
				pstrParent2PickListIDs = pstrParent2PickListIDs & IIf(Len(pstrParent2PickListIDs) > 0, ", ", "") & prstTemp.Fields(0).Value
				prstTemp.MoveNext()
			Loop

			prstTemp.Close()
			'UPGRADE_NOTE: Object prstTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			prstTemp = Nothing

			mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & mstrBaseTableRealSource & ".ID_" & mlngCustomReportsParent2Table & " IN (" & pstrParent2PickListIDs & ") "
		ElseIf mlngCustomReportsParent2FilterID > 0 Then
			blnOK = True
			blnOK = datGeneral.FilteredIDs(mlngCustomReportsParent2FilterID, strFilterIDs, mvarPrompts)

			' Generate any UDFs that are used in this filter
			If blnOK Then
				datGeneral.FilterUDFs(mlngCustomReportsParent2FilterID, mastrUDFsRequired)
			End If

			If blnOK Then
				mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & mstrBaseTableRealSource & ".ID_" & mlngCustomReportsParent2Table & " IN (" & strFilterIDs & ") "
			Else
				mstrErrorString = "You do not have permission to use the '" & datGeneral.GetFilterName(mlngCustomReportsParent2FilterID) & "' filter."
				GenerateSQLWhere = False
				Exit Function
			End If
		End If

		If mlngSingleRecordID > 0 Then
			mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & mstrSQLFrom & ".ID IN (" & CStr(mlngSingleRecordID) & ")"

		ElseIf mlngCustomReportsPickListID > 0 Then
			' Now if we are using a picklist, add a where clause for that
			'Get List of IDs from Picklist
			prstTemp = mclsData.OpenRecordset("EXEC sp_ASRGetPickListRecords " & mlngCustomReportsPickListID, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly)

			If prstTemp.BOF And prstTemp.EOF Then
				mstrErrorString = "The selected picklist contains no records."
				GenerateSQLWhere = False
				Exit Function
			End If

			Do While Not prstTemp.EOF
				pstrPickListIDs = pstrPickListIDs & IIf(Len(pstrPickListIDs) > 0, ", ", "") & prstTemp.Fields(0).Value
				prstTemp.MoveNext()
			Loop

			prstTemp.Close()
			'UPGRADE_NOTE: Object prstTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			prstTemp = Nothing

			mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & mstrSQLFrom & ".ID IN (" & pstrPickListIDs & ")"

			' If we are running a Bradford Report on an individual person
		ElseIf mbIsBradfordIndexReport = True And mlngPersonnelID > 0 Then

			mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & mstrSQLFrom & ".ID IN (" & mlngPersonnelID & ")"

		ElseIf mlngCustomReportsFilterID > 0 Then

			blnOK = datGeneral.FilteredIDs(mlngCustomReportsFilterID, strFilterIDs, mvarPrompts)

			' Generate any UDFs that are used in this filter
			If blnOK Then
				datGeneral.FilterUDFs(mlngCustomReportsFilterID, mastrUDFsRequired)
			End If

			If blnOK Then
				mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & mstrSQLFrom & ".ID IN (" & strFilterIDs & ")"
			Else
				' Permission denied on something in the filter.
				mstrErrorString = "You do not have permission to use the '" & datGeneral.GetFilterName(mlngCustomReportsFilterID) & "' filter."
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
		mobjEventLog.AddDetailEntry(mstrErrorString)
		mobjEventLog.ChangeHeaderStatus(EventLog_Status.elsFailed)

	End Function

	Private Function GenerateSQLOrderBy() As Boolean

		' Purpose : Returns order by string from the sort order array
		Dim strOrder As String
		Dim pblnColumnOK As Boolean
		Dim pblnNoSelect As Boolean
		Dim pblnFound As Boolean
		Dim pstrSource As String
		Dim pstrOrderFrom1 As String
		Dim pstrOrderFrom2 As String
		Dim pintNextIndex As Short

		On Error GoTo GenerateSQLOrderBy_ERROR

		' Bradford Factor has it own sort order code
		If mbIsBradfordIndexReport Then
			'*********************************************************************************
			'TM20020605 Fault 3912 - check that the current user has permission to
			' see and therefore order by the selected order columns on the table.

			'First Order Column - Check the user has select access through a table or view.
			If mlngOrderByColumnID > 0 Then
				mobjColumnPrivileges = GetColumnPrivileges(mstrCustomReportsBaseTableName)
				pblnColumnOK = mobjColumnPrivileges.IsValid(mstrOrderByColumn)
				If pblnColumnOK Then
					pblnColumnOK = mobjColumnPrivileges.Item(mstrOrderByColumn).AllowSelect
				End If

				If Not pblnColumnOK Then
					' this column cannot be read direct. If its from a parent, try parent views
					' Loop thru the views on the table, seeing if any have read permis for the column
					ReDim mstrViews(0)
					For Each mobjTableView In gcoTablePrivileges.Collection
						If (Not mobjTableView.IsTable) And (mobjTableView.TableID = mlngCustomReportsBaseTable) And (mobjTableView.AllowSelect) Then

							pstrSource = mobjTableView.ViewName

							' Get the column permission for the view
							mobjColumnPrivileges = GetColumnPrivileges(pstrSource)

							' If we can see the column from this view
							If mobjColumnPrivileges.IsValid(mstrOrderByColumn) Then
								If mobjColumnPrivileges.Item(mstrOrderByColumn).AllowSelect Then

									ReDim Preserve mstrViews(UBound(mstrViews) + 1)
									mstrViews(UBound(mstrViews)) = mobjTableView.ViewName

									pstrOrderFrom1 = mobjTableView.ViewName

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
										Exit For
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
					End If

				Else
					pstrOrderFrom1 = mstrCustomReportsBaseTableName
				End If

				If pblnNoSelect Then
					GenerateSQLOrderBy = False
					mstrErrorString = vbNewLine & "You do not have permission to see the column '" & mstrOrderByColumn & "' " & vbNewLine & "either directly or through any views."
					Exit Function
				End If
			End If

			'Second Order Column - Check the user has select access through a table or view.
			If mlngGroupByColumnID > 0 Then
				pblnNoSelect = False
				mobjColumnPrivileges = GetColumnPrivileges(mstrCustomReportsBaseTableName)
				pblnColumnOK = mobjColumnPrivileges.IsValid(mstrGroupByColumn)
				If pblnColumnOK Then
					pblnColumnOK = mobjColumnPrivileges.Item(mstrGroupByColumn).AllowSelect
				End If

				If Not pblnColumnOK Then
					' this column cannot be read direct. If its from a parent, try parent views
					' Loop thru the views on the table, seeing if any have read permis for the column
					ReDim mstrViews(0)
					For Each mobjTableView In gcoTablePrivileges.Collection
						If (Not mobjTableView.IsTable) And (mobjTableView.TableID = mlngCustomReportsBaseTable) And (mobjTableView.AllowSelect) Then

							pstrSource = mobjTableView.ViewName

							' Get the column permission for the view
							mobjColumnPrivileges = GetColumnPrivileges(pstrSource)

							' If we can see the column from this view
							If mobjColumnPrivileges.IsValid(mstrOrderByColumn) Then
								If mobjColumnPrivileges.Item(mstrOrderByColumn).AllowSelect Then

									ReDim Preserve mstrViews(UBound(mstrViews) + 1)
									mstrViews(UBound(mstrViews)) = mobjTableView.ViewName

									pstrOrderFrom2 = mobjTableView.ViewName

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
										Exit For
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
					End If

				Else
					pstrOrderFrom2 = mstrCustomReportsBaseTableName
				End If

				If pblnNoSelect Then
					GenerateSQLOrderBy = False
					mstrErrorString = vbNewLine & "You do not have permission to see the column '" & mstrGroupByColumn & "' " & vbNewLine & "either directly or through any views."
					Exit Function
				End If
			End If
			'*********************************************************************************

			'TM24032004
			'      'MH20020521 Fault 3820
			'      'strOrder = "[" & mstrOrderByColumn & "] " & IIf(mbOrderBy1Asc = True, "Asc", "Desc")
			'      'If Not mstrGroupByColumn = "<None>" And Not mstrGroupByColumn = mstrOrderByColumn Then
			'      '  strOrder = strOrder & ",[" & mstrGroupByColumn & "] " & IIf(mbOrderBy2Asc = True, "Asc", "Desc")
			'      'End If
			'      'mstrSQLOrderBy = " ORDER BY " & strOrder & ", [Personnel_ID] Asc, [Start_Date] Asc"
			'      strOrder = "[" & pstrOrderFrom1 & "].[" & mstrOrderByColumn & "] " & IIf(mbOrderBy1Asc = True, "Asc", "Desc")
			'      If Not mstrGroupByColumn = "None" And Not mstrGroupByColumn = mstrOrderByColumn Then
			'        strOrder = strOrder & ", [" & pstrOrderFrom2 & "].[" & mstrGroupByColumn & "] " & IIf(mbOrderBy2Asc = True, "Asc", "Desc")
			'      End If
			'      mstrSQLOrderBy = " ORDER BY " & strOrder & ", [Personnel_ID] Asc"
			'      If InStr(strOrder, "[Start_Date]") = 0 Then
			'        mstrSQLOrderBy = mstrSQLOrderBy & ", [Start_Date] Asc"
			'      End If
			If mlngOrderByColumnID > 0 Then
				strOrder = "[Order_1] " & IIf(mbOrderBy1Asc = True, "Asc", "Desc")
			End If
			If mlngGroupByColumnID > 0 And (mlngOrderByColumnID <> mlngGroupByColumnID) Then
				If mlngOrderByColumnID > 0 Then
					strOrder = strOrder & ", "
					strOrder = strOrder & "[Order_2] " & IIf(mbOrderBy2Asc = True, "Asc", "Desc")
				Else
					strOrder = strOrder & "[Order_1] " & IIf(mbOrderBy2Asc = True, "Asc", "Desc")
				End If
			End If
			If (mlngOrderByColumnID = 0) And (mlngGroupByColumnID = 0) Then
				mstrSQLOrderBy = " ORDER BY [Personnel_ID] Asc"
			Else
				mstrSQLOrderBy = " ORDER BY " & strOrder & ", [Personnel_ID] Asc"
			End If

		Else

			If UBound(mvarSortOrder, 2) > 0 Then
				' Columns have been defined, so use these for the base table/view
				mstrSQLOrderBy = DoDefinedOrderBy()
			End If

			If Len(mstrSQLOrderBy) > 0 Then mstrSQLOrderBy = " ORDER BY " & mstrSQLOrderBy

		End If



		GenerateSQLOrderBy = True
		Exit Function

GenerateSQLOrderBy_ERROR:

		GenerateSQLOrderBy = False
		mstrErrorString = "Error in GenerateSQLOrderBy." & vbNewLine & Err.Description
		mobjEventLog.AddDetailEntry(mstrErrorString)
		mobjEventLog.ChangeHeaderStatus(EventLog_Status.elsFailed)

	End Function

	Private Function DoDefinedOrderBy() As String

		' This function creates the base ORDER BY statement by searching
		' through the columns defined as the reports sort order, then
		' uses the relevant alias name

		Dim iLoop As Short
		Dim iLoop2 As Short

		For iLoop = 1 To UBound(mvarSortOrder, 2)

			For iLoop2 = 1 To UBound(mvarColDetails, 2)

				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(12, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If mvarSortOrder(1, iLoop) = mvarColDetails(12, iLoop2) Then

					'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					DoDefinedOrderBy = DoDefinedOrderBy & IIf(Len(DoDefinedOrderBy) > 0, ",", "") & "[" & mvarColDetails(0, iLoop2) & "] " & mvarSortOrder(2, iLoop)

					Exit For

				End If

			Next iLoop2

		Next iLoop

	End Function

	Private Function GetTableIDFromColumn(ByVal lngColumnID As Integer) As Integer

		' Purpose : To return the table id for which the given column belongs

		Dim rsInfo As Recordset
		Dim strSQL As String

		strSQL = "SELECT ASRSysTables.TableID " & "FROM ASRSysColumns JOIN ASRSysTables " & "ON (ASRSysTables.TableID = ASRSysColumns.TableID) " & "WHERE ColumnID = " & CStr(lngColumnID)

		rsInfo = mclsData.OpenRecordset(strSQL, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly)

		If rsInfo.BOF And rsInfo.EOF Then
			GetTableIDFromColumn = 0
		Else
			GetTableIDFromColumn = rsInfo.Fields("TableID").Value
		End If

		'UPGRADE_NOTE: Object rsInfo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsInfo = Nothing

	End Function

	Public Function CheckRecordSet() As Boolean

		' Purpose : To get recordset from temptable and show recordcount

		Dim sSQL As String

		On Error GoTo CheckRecordSet_ERROR

		'  Set mrstCustomReportsOutput = mclsData.OpenRecordset("SELECT * FROM " & mstrTempTableName, adOpenStatic, adLockReadOnly)

		'TM20020429 Fault 3764
		If mbUseSequence Then
			sSQL = "SELECT * FROM [" & mstrTempTableName & "]"
			sSQL = sSQL & " ORDER BY [" & lng_SEQUENCECOLUMNNAME & "] ASC"
		Else
			sSQL = "SELECT * FROM " & mstrTempTableName

			If mbIsBradfordIndexReport Then
				sSQL = sSQL & mstrSQLOrderBy
			End If

		End If

		mrstCustomReportsOutput = mclsData.OpenRecordset(sSQL, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly)

		If mrstCustomReportsOutput.BOF And mrstCustomReportsOutput.EOF Then
			CheckRecordSet = False
			mstrErrorString = "No records meet selection criteria"
			mobjEventLog.AddDetailEntry("Completed successfully. " & mstrErrorString)
			mobjEventLog.ChangeHeaderStatus(EventLog_Status.elsSuccessful)
			mblnNoRecords = True
			Exit Function
		End If

		If mlngColumnLimit > 0 Then
			If mrstCustomReportsOutput.Fields.Count > mlngColumnLimit Then
				CheckRecordSet = False
				mstrErrorString = "Report contains more than " & mlngColumnLimit & " columns. It is not possible to run this report via the intranet."
				mobjEventLog.AddDetailEntry("Failed. " & mstrErrorString)
				mobjEventLog.ChangeHeaderStatus(EventLog_Status.elsFailed)
				mblnNoRecords = False
				Exit Function
			End If
		End If

		CheckRecordSet = True
		Exit Function

CheckRecordSet_ERROR:

		mstrErrorString = "Error while checking returned recordset." & vbNewLine & "(" & Err.Description & ")"
		CheckRecordSet = False
		mobjEventLog.AddDetailEntry(mstrErrorString)
		mobjEventLog.ChangeHeaderStatus(EventLog_Status.elsFailed)

	End Function

	Public Function PopulateGrid_LoadRecords() As Boolean
		' Purpose : Blimey ! This function does the actual work of populating the
		'           grid, calculating summary info, breaking, page breaking etc.
		'           Its a bit of a 'mare but it works ok. (JDM - I question that!)

		On Error GoTo LoadRecords_ERROR

		Dim sAddString As String
		Dim iLoop As Short
		Dim vDisplayData As Object
		Dim avColumns(,) As Object
		Dim vValue As Object
		Dim fBreak As Boolean
		Dim iLoop2 As Short
		Dim iLoop3 As Short
		Dim iColumnIndex As Short
		Dim iOtherColumnIndex As Short
		Dim fNotFirstTime As Boolean
		Dim bSuppress As Boolean

		Dim intColCounter As Short

		Dim sBreakValue As String

		'Group With Next Column variables
		Dim intRowIndex_GW As Short
		Dim intColIndex_GW As Short
		Dim intGroupCount As Short
		Dim blnHasGroupWithNext As Boolean
		Dim blnSkipped As Boolean
		Dim intSkippedIndex As Short
		Dim strGroupString As String
		Dim sLastValue As String = vbNullString

		intRowIndex_GW = 0
		intColIndex_GW = 0
		blnHasGroupWithNext = False
		blnSkipped = False
		intSkippedIndex = 0

		'Variables for Suppress Repeated Values within Table.
		Dim lngCurrentRecordID As Integer
		Dim bBaseRecordChanged As Boolean
		Dim isHiddenColumn As Boolean

		' Construct an array of the columns in the report. Basically this is an extension of the mvarColDetails array.
		' Col 1 = TRUE if the column is used for breaking/paging on change.
		' Col 2 = TRUE if the column is to be aggregated (average/count/total), else FALSE.
		' Col 3 = last column value.
		ReDim avColumns(3, UBound(mvarColDetails, 2))
		For iLoop = 1 To UBound(mvarColDetails, 2)
			'UPGRADE_WARNING: Couldn't resolve default property of object avColumns(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			avColumns(1, iLoop) = mvarColDetails(7, iLoop) Or mvarColDetails(8, iLoop)
			'UPGRADE_WARNING: Couldn't resolve default property of object avColumns(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			avColumns(2, iLoop) = mvarColDetails(4, iLoop) Or mvarColDetails(5, iLoop) Or mvarColDetails(6, iLoop)
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object avColumns(3, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			avColumns(3, iLoop) = DBNull.Value
		Next iLoop

		'Ensure we are at the beginning of the output recordset
		mrstCustomReportsOutput.MoveFirst()

		Do Until mrstCustomReportsOutput.EOF

			'bRecordChanged used for repetition funcionality.
			If Not mbIsBradfordIndexReport Then
				If mrstCustomReportsOutput.Fields("?ID").Value <> lngCurrentRecordID Then
					bBaseRecordChanged = True
					lngCurrentRecordID = mrstCustomReportsOutput.Fields("?ID").Value
				Else
					bBaseRecordChanged = False
				End If
			End If

			'offset the addstring by 2 columns (for pagebreak and summary info)
			sAddString = "<td class='hiddentablecolumn'></td><td class='summarytablecolumn'></td>"

			' Dont do summary info for first record (otherwise blank!)
			'If mrstCustomReportsOutput.AbsolutePosition > 1 Then
			If fNotFirstTime Then
				' Put the values from the previous record in the column array.
				For iLoop = 1 To UBound(mvarColDetails, 2)
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(11, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object avColumns(3, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					avColumns(3, iLoop) = mvarColDetails(11, iLoop)
				Next iLoop

				' From last column in the order to first, check changes.
				For iLoop = UBound(mvarSortOrder, 2) To 1 Step -1
					' Find the column in the details array.
					iColumnIndex = 0
					For iLoop2 = 1 To UBound(mvarColDetails, 2)
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(13, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(12, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If (mvarColDetails(12, iLoop2).ToString() = mvarSortOrder(1, iLoop).ToString()) And (mvarColDetails(13, iLoop2) = "C") Then
							iColumnIndex = iLoop2
							Exit For
						End If
					Next iLoop2

					If iColumnIndex > 0 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object avColumns(1, iColumnIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If avColumns(1, iColumnIndex) Then
							fBreak = False

							' The column breaks. Check if its changed.
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							If IsDBNull(mrstCustomReportsOutput.Fields(iColumnIndex - 1).Value) And (Not mvarColDetails(3, iColumnIndex)) And (Not mvarColDetails(17, iColumnIndex)) And (Not mvarColDetails(18, iColumnIndex)) Then
								' Field value is null but a character data type, so set it to be "".
								'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								vValue = ""
							Else
								'Dates need to be formatted with yyyy for boc to work ok
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(17, iColumnIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'

								If mvarColDetails(17, iColumnIndex) Then 'Date
									'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(1, iColumnIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									vValue = Left(VB6.Format(mrstCustomReportsOutput.Fields(iColumnIndex - 1).Value, DateFormat), mvarColDetails(1, iColumnIndex))

									'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(3, iColumnIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ElseIf mvarColDetails(3, iColumnIndex) Then	 'Numeric
									'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(1, iColumnIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									vValue = Left(mrstCustomReportsOutput.Fields(iColumnIndex - 1).Value, mvarColDetails(1, iColumnIndex))

									'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(18, iColumnIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ElseIf mvarColDetails(18, iColumnIndex) Then	'Bit
									'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									If (mrstCustomReportsOutput.Fields(iColumnIndex - 1).Value = True) Or (mrstCustomReportsOutput.Fields(iColumnIndex - 1).Value = 1) Then vValue = "Y"
									'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									If (mrstCustomReportsOutput.Fields(iColumnIndex - 1).Value = False) Or (mrstCustomReportsOutput.Fields(iColumnIndex - 1).Value = 0) Then vValue = "N"
									'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
									'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									If IsDBNull(mrstCustomReportsOutput.Fields(iColumnIndex - 1).Value) Then vValue = ""

								Else 'Varchar
									'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(1, iColumnIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									vValue = Left(mrstCustomReportsOutput.Fields(iColumnIndex - 1).Value, mvarColDetails(1, iColumnIndex))

								End If
							End If

							'Now that we store the formatted value in position (11) of the mcolDetails
							'Comparison made after adjusting the size of the field.
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							If IsDBNull(vValue) Or IsDBNull(mrstCustomReportsOutput.Fields(iColumnIndex - 1).Value) Then
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								fBreak = ("" <> mvarColDetails(11, iColumnIndex))
							Else
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(18, iColumnIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If mvarColDetails(18, iColumnIndex) Then 'Bit
									'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									fBreak = (RTrim(LCase(vValue)) <> RTrim(LCase(mvarColDetails(11, iColumnIndex))))
								Else
									'TM23112004 Fault 9072
									'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									fBreak = (RTrim(LCase(mrstCustomReportsOutput.Fields(iColumnIndex - 1).Value)) <> RTrim(LCase(mvarColDetails(11, iColumnIndex))))
								End If
							End If

							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(8, iColumnIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If mvarColDetails(8, iColumnIndex) Then
								sBreakValue = IIf(Len(mvarColDetails(11, iColumnIndex)) < 1, "<Empty>", mvarColDetails(11, iColumnIndex)) & IIf(Len(sBreakValue) > 0, " - ", "") & sBreakValue
							End If

							If Not fBreak Then
								' The value has not changed, but check if we need to do the summary due to another column changing.
								For iLoop2 = (iLoop - 1) To 1 Step -1
									iOtherColumnIndex = 0
									For iLoop3 = 1 To UBound(mvarColDetails, 2)
										'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(13, iLoop3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(1, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(12, iLoop3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										If mvarColDetails(12, iLoop3).ToString() = mvarSortOrder(1, iLoop2).ToString() And mvarColDetails(13, iLoop3) = "C" Then
											iOtherColumnIndex = iLoop3
											Exit For
										End If
									Next iLoop3

									If iOtherColumnIndex > 0 Then
										'UPGRADE_WARNING: Couldn't resolve default property of object avColumns(1, iOtherColumnIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										If avColumns(1, iOtherColumnIndex) Then
											' The column breaks. Check if its changed.
											'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
											If IsDBNull(mrstCustomReportsOutput.Fields(iOtherColumnIndex - 1).Value) And (Not mvarColDetails(3, iOtherColumnIndex)) And (Not mvarColDetails(17, iOtherColumnIndex)) And (Not mvarColDetails(18, iOtherColumnIndex)) Then
												' Field value is null but a character data type, so set it to be "".
												'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												vValue = ""

												'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(3, iOtherColumnIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											ElseIf mvarColDetails(3, iOtherColumnIndex) Then	'Numeric
												'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(1, iOtherColumnIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												vValue = Left(mrstCustomReportsOutput.Fields(iOtherColumnIndex - 1).Value, mvarColDetails(1, iOtherColumnIndex))

												'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(18, iOtherColumnIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											ElseIf mvarColDetails(18, iOtherColumnIndex) Then	 'Bit
												'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												If (mrstCustomReportsOutput.Fields(iOtherColumnIndex - 1).Value = True) Or (mrstCustomReportsOutput.Fields(iOtherColumnIndex - 1).Value = 1) Then vValue = "Y"
												'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												If (mrstCustomReportsOutput.Fields(iOtherColumnIndex - 1).Value = False) Or (mrstCustomReportsOutput.Fields(iOtherColumnIndex - 1).Value = 0) Then vValue = "N"
												'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
												'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												If IsDBNull(mrstCustomReportsOutput.Fields(iOtherColumnIndex - 1).Value) Then vValue = ""

											Else
												'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(1, iOtherColumnIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												vValue = Left(mrstCustomReportsOutput.Fields(iOtherColumnIndex - 1).Value, mvarColDetails(1, iOtherColumnIndex))

											End If

											'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
											If IsDBNull(vValue) Or IsDBNull(mrstCustomReportsOutput.Fields(iOtherColumnIndex - 1).Value) Then
												'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												fBreak = ("" <> mvarColDetails(11, iOtherColumnIndex))
											Else
												'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(18, iOtherColumnIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												If mvarColDetails(18, iOtherColumnIndex) Then	'Bit
													'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													fBreak = (RTrim(LCase(vValue)) <> RTrim(LCase(mvarColDetails(11, iOtherColumnIndex))))
												Else
													'TM23112004 Fault 9072
													'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													fBreak = (RTrim(LCase(mrstCustomReportsOutput.Fields(iOtherColumnIndex - 1).Value)) <> RTrim(LCase(mvarColDetails(11, iOtherColumnIndex))))
												End If
											End If

											If (fBreak = True) Then
												Exit For
											End If
										End If
									End If
								Next iLoop2
							End If

							' RH 09/02/01 - Report was doing summary info even when no aggregate was
							'               selected for the column, so check for aggregate too, and only
							'               do summary info if its true.
							If fBreak Then
								PopulateGrid_DoSummaryInfo(avColumns, iColumnIndex, iLoop)

								' Do the page break ?
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(8, iColumnIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If mvarColDetails(8, iColumnIndex) Then
									mblnPageBreak = True
									mblnReportHasPageBreak = True
								End If
							End If
						End If
					End If
				Next iLoop
			End If

			If mblnPageBreak Then
				AddToArray_Data("*", enum_RowType.Data)
				mintPageBreakRowIndex = mintPageBreakRowIndex + 1
				AddPageBreakValue(mintPageBreakRowIndex - 1, sBreakValue)
			End If
			mblnPageBreak = False
			sBreakValue = vbNullString

			intColCounter = 1
			' Loop thru each field, adding to the string to add to the grid
			For iLoop = 0 To (mrstCustomReportsOutput.Fields.Count - 1)

				intColCounter = intColCounter + 1
				isHiddenColumn = (mrstCustomReportsOutput.Fields(iLoop).Name.Substring(0, 1) = "?")		' there should be a cleaner way of deciding if this is an ID column, but would need more bigger changes. This will have to do for the moment. Sorry

				' yet another hack beacsue this is an over complex array instead of an easily modifyable class
				If mbIsBradfordIndexReport And iLoop > 12 Then
					isHiddenColumn = True
				End If

				' Only suppress values for new records in the Bradford Factor report
				bSuppress = IIf(mbIsBradfordIndexReport And fBreak, False, True)

				If (mvarColDetails(18, iLoop + 1)) Then	'Bit
					'UPGRADE_WARNING: Couldn't resolve default property of object tmpLogicValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If IsDBNull(mrstCustomReportsOutput.Fields(iLoop).Value) Then
						vDisplayData = ""
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object tmpLogicValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If (mrstCustomReportsOutput.Fields(iLoop).Value = "True") Or (mrstCustomReportsOutput.Fields(iLoop).Value = 1) Then vDisplayData = "Y"
						'UPGRADE_WARNING: Couldn't resolve default property of object tmpLogicValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If (mrstCustomReportsOutput.Fields(iLoop).Value = "False") Or (mrstCustomReportsOutput.Fields(iLoop).Value = 0) Then vDisplayData = "N"
					End If

				Else
					' Get the formatted data to display in the grid
					'UPGRADE_WARNING: Couldn't resolve default property of object PopulateGrid_FormatData(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object vDisplayData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					vDisplayData = PopulateGrid_FormatData((mrstCustomReportsOutput.Fields(iLoop).Name), mrstCustomReportsOutput.Fields(iLoop).Value, bSuppress, bBaseRecordChanged)
				End If

				If blnSkipped Then
					' Store the ACTUAL data in the array (previous value dimension)
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If IsDBNull(mrstCustomReportsOutput.Fields(intSkippedIndex).Value) And (Not mvarColDetails(3, intSkippedIndex + 1)) And (Not mvarColDetails(17, intSkippedIndex + 1)) And (Not mvarColDetails(18, intSkippedIndex + 1)) Then
						' Field value is null but a character data type, so set it to be "".
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(11, intSkippedIndex + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarColDetails(11, intSkippedIndex + 1) = ""

					Else
						'TM17052005 Fault 10086 - Need to store diffent values depending on the type.
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(17, intSkippedIndex + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If mvarColDetails(17, intSkippedIndex + 1) Then	'Date
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(11, intSkippedIndex + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mvarColDetails(11, intSkippedIndex + 1) = VB6.Format(mrstCustomReportsOutput.Fields(intSkippedIndex).Value, DateFormat)

						ElseIf (mvarColDetails(3, intSkippedIndex + 1)) Then	'Numeric
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(11, intSkippedIndex + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mvarColDetails(11, intSkippedIndex + 1) = IIf(IsDBNull(mrstCustomReportsOutput.Fields(intSkippedIndex).Value), "", mrstCustomReportsOutput.Fields(intSkippedIndex).Value)

						ElseIf (mvarColDetails(18, intSkippedIndex + 1)) Then	 'Bit
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(11, intSkippedIndex + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If (mrstCustomReportsOutput.Fields(intSkippedIndex).Value = "True") Or (mrstCustomReportsOutput.Fields(intSkippedIndex).Value = 1) Then mvarColDetails(11, intSkippedIndex + 1) = "Y"
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(11, intSkippedIndex + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If (mrstCustomReportsOutput.Fields(intSkippedIndex).Value = "False") Or (mrstCustomReportsOutput.Fields(intSkippedIndex).Value = 0) Then mvarColDetails(11, intSkippedIndex + 1) = "N"
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(11, intSkippedIndex + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If IsDBNull(mrstCustomReportsOutput.Fields(intSkippedIndex).Value) Then mvarColDetails(11, intSkippedIndex + 1) = ""

						Else 'Varchar
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(11, intSkippedIndex + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mvarColDetails(11, intSkippedIndex + 1) = mrstCustomReportsOutput.Fields(intSkippedIndex).Value
						End If

					End If

				Else
					' Store the ACTUAL data in the array (previous value dimension)
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If IsDBNull(mrstCustomReportsOutput.Fields(iLoop).Value) And (Not mvarColDetails(3, iLoop + 1)) And (Not mvarColDetails(17, iLoop + 1)) And (Not mvarColDetails(18, iLoop + 1)) Then
						' Field value is null but a character data type, so set it to be "".
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(11, iLoop + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarColDetails(11, iLoop + 1) = ""
					Else

						'TM17052005 Fault 10086 - Need to store diffent values depending on the type.
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(17, iLoop + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If mvarColDetails(17, iLoop + 1) Then	'Date
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(11, iLoop + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mvarColDetails(11, iLoop + 1) = VB6.Format(mrstCustomReportsOutput.Fields(iLoop).Value, DateFormat)

						ElseIf (mvarColDetails(3, iLoop + 1)) Then	'Numeric
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(11, iLoop + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mvarColDetails(11, iLoop + 1) = IIf(IsDBNull(mrstCustomReportsOutput.Fields(iLoop).Value), "", mrstCustomReportsOutput.Fields(iLoop).Value)

						ElseIf (mvarColDetails(18, iLoop + 1)) Then	 'Bit
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(11, iLoop + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(11, iLoop + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If IsDBNull(mrstCustomReportsOutput.Fields(iLoop).Value) Then
								mvarColDetails(11, iLoop + 1) = ""
							Else
								If (mrstCustomReportsOutput.Fields(iLoop).Value = "True") Or (mrstCustomReportsOutput.Fields(iLoop).Value = 1) Then mvarColDetails(11, iLoop + 1) = "Y"
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(11, iLoop + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If (mrstCustomReportsOutput.Fields(iLoop).Value = "False") Or (mrstCustomReportsOutput.Fields(iLoop).Value = 0) Then mvarColDetails(11, iLoop + 1) = "N"

							End If

						Else 'Varchar
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(11, iLoop + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mvarColDetails(11, iLoop + 1) = mrstCustomReportsOutput.Fields(iLoop).Value
						End If
					End If
				End If

				' Group with next column
				If (mvarColDetails(24, iLoop + 1) And (Not mvarColDetails(19, iLoop + 1))) Then
					sLastValue = sLastValue & vDisplayData & IIf(Not vDisplayData Is Nothing, "<br/>", "")

				Else
					vDisplayData = sLastValue & vDisplayData
					sLastValue = vbNullString

					If isHiddenColumn Then
						sAddString = sAddString & "<td class='hiddentablecolumn'>" & vDisplayData & "</td>"
					Else
						If mvarColDetails(3, iLoop + 1) Then
							sAddString = sAddString & "<td class='visibletablecolumn' align='right'>" & vDisplayData & "</td>"
						Else
							sAddString = sAddString & "<td class='visibletablecolumn'>" & vDisplayData & "</td>"
						End If
					End If

				End If

			Next iLoop

			' Only Add the addstring to the grid if its not a summary report
			If mblnCustomReportsSummaryReport = False Then
				If Not AddToArray_Data(sAddString, enum_RowType.Data) Then
					PopulateGrid_LoadRecords = False
					Exit Function
				Else
					mintPageBreakRowIndex = mintPageBreakRowIndex + 1

					If blnHasGroupWithNext Then
						strGroupString = vbNullString
						For intGroupCount = 0 To UBound(mvarGroupWith, 2) Step 1
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarGroupWith(0, intGroupCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strGroupString = strGroupString & vbNewLine & mvarGroupWith(0, intGroupCount)

							mintPageBreakRowIndex = mintPageBreakRowIndex + 1
						Next intGroupCount
						AddToArray_Data(strGroupString, enum_RowType.Data)

					End If

				End If
			End If

			'Clear the Group Arrays/Variables
			ReDim mvarGroupWith(1, 0)
			intRowIndex_GW = 0
			intColIndex_GW = 0
			blnHasGroupWithNext = False

			fNotFirstTime = True

			' Move to next row in the grid
			mrstCustomReportsOutput.MoveNext()
		Loop

		mblnPageBreak = False

		' Now do the final summary for the last bit (before the grand summary)
		' Put the values from the previous record in the column array.
		For iLoop = 1 To UBound(mvarColDetails, 2)
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(11, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object avColumns(3, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			avColumns(3, iLoop) = mvarColDetails(11, iLoop)
		Next iLoop
		' From last column in the order to first, check changes.
		For iLoop = UBound(mvarSortOrder, 2) To 1 Step -1
			' Find the column in the details array.
			iColumnIndex = 0
			For iLoop2 = 1 To UBound(mvarColDetails, 2)
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(13, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(12, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (mvarColDetails(12, iLoop2).ToString() = mvarSortOrder(1, iLoop).ToString()) And (mvarColDetails(13, iLoop2) = "C") Then

					iColumnIndex = iLoop2
					Exit For
				End If
			Next iLoop2

			If iColumnIndex > 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object avColumns(1, iColumnIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If avColumns(1, iColumnIndex) Then
					'        mblnPageBreak = True
					'        sBreakValue = IIf(Len(mvarColDetails(11, iColumnIndex)) < 1, "<Empty>", mvarColDetails(11, iColumnIndex)) & IIf(Len(sBreakValue) > 0, " - ", "") & sBreakValue

					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(8, iColumnIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If mvarColDetails(8, iColumnIndex) Then
						mblnPageBreak = True
						sBreakValue = IIf(Len(mvarColDetails(11, iColumnIndex)) < 1, "<Empty>", mvarColDetails(11, iColumnIndex)) & IIf(Len(sBreakValue) > 0, " - ", "") & sBreakValue
					End If

					PopulateGrid_DoSummaryInfo(avColumns, iColumnIndex, iLoop)
				End If
			End If
		Next iLoop

		If mblnPageBreak Then
			'AddToArray_Data "*"
			mintPageBreakRowIndex = mintPageBreakRowIndex + 1
			AddPageBreakValue(mintPageBreakRowIndex - 1, sBreakValue)
		End If
		sBreakValue = vbNullString

		' Now do the grand summary information
		If Not mbIsBradfordIndexReport Then
			PopulateGrid_DoGrandSummary()

			If mblnPageBreak And mblnDoesHaveGrandSummary Then
				AddPageBreakValue(mintPageBreakRowIndex - 1, sBreakValue)
				mintPageBreakRowIndex = mintPageBreakRowIndex + 1
			End If

		End If

		If Not AddToArray_Data(vbNullString, enum_RowType.Data) Then
			Return False
		End If

		Return True

LoadRecords_ERROR:

		mstrErrorString = mstrErrorString & "LOADRECORDS_ERROR (In Dll) - Error in PopulateGrid_LoadRecords." & vbNewLine & Err.Number & " - " & Err.Description
		mobjEventLog.AddDetailEntry(mstrErrorString)
		mobjEventLog.ChangeHeaderStatus(EventLog_Status.elsFailed)
		Return False

	End Function

	Private Sub AddPageBreakValue(ByVal pintRowIndex As Short, ByVal pvarValue As Object)

		ReDim Preserve mvarPageBreak(pintRowIndex)
		'UPGRADE_WARNING: Couldn't resolve default property of object pvarValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarPageBreak(pintRowIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarPageBreak(pintRowIndex) = pvarValue

	End Sub

	Private Function PopulateGrid_FormatData(ByVal sfieldname As String, ByVal vData As Object, ByVal mbSuppressRepeated As Boolean, ByVal pbNewBaseRecord As Boolean) As Object
		' Purpose : Format the data to the form the user has specified to see it
		'           in the grid
		' Input   : None
		' Output  : True/False
		Dim pintLoop As Short
		Dim vOriginalData As Object

		If IsDBNull(vData) Then Return ""

		'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object vOriginalData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		vOriginalData = vData

		For pintLoop = 1 To UBound(mvarColDetails, 2)
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If mvarColDetails(0, pintLoop) = sfieldname Then

				' Is it a boolean calculation ? If so, change to Y or N
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(18, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If mvarColDetails(18, pintLoop) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If vData = "True" Then vData = "Y"
					'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If vData = "False" Then vData = "N"
				End If

				' If its a date column, format it as dateformat
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(17, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If mvarColDetails(17, pintLoop) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					vData = VB6.Format(vData, mstrClientDateFormat)
				End If


				' Is Numeric
				If mvarColDetails(3, pintLoop) Then

					vData = CDbl(vData)

					' Overflow check (ignore decimals)
					If CLng(vData).ToString.Length > mvarColDetails(1, pintLoop) Then
						vData = New String("#", mvarColDetails(1, pintLoop))
					Else
						If Not mvarColDetails(23, pintLoop) Is Nothing Then
							vData = String.Format(mvarColDetails(23, pintLoop), vData)
						End If
					End If

				End If


				' SRV ?
				If Not mbIsBradfordIndexReport Then
					If mbSuppressRepeated = True Then
						'check if column value should be repeated or not.
						If Not mvarColDetails(21, pintLoop) And Not pbNewBaseRecord And Not mvarColDetails(10, pintLoop) And Not mvarColDetails(20, pintLoop) Then
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							If CStr(RTrim(IIf(IsDBNull(mvarColDetails(11, pintLoop)), vbNullString, mvarColDetails(11, pintLoop)))) = CStr(RTrim(IIf(IsDBNull(vOriginalData), vbNullString, vOriginalData))) Then
								'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								vData = ""
							End If
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(10, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ElseIf mvarColDetails(10, pintLoop) Then
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							If CStr(RTrim(IIf(IsDBNull(mvarColDetails(11, pintLoop)), vbNullString, mvarColDetails(11, pintLoop)))) = CStr(RTrim(IIf(IsDBNull(vOriginalData), vbNullString, vOriginalData))) Then
								'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								vData = ""
							End If
						End If
					End If
					Exit For
				Else
					'Bradford Factor does not use the repetition functionality.
					If mbSuppressRepeated = True Then
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(10, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If mvarColDetails(10, pintLoop) Then
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							If CStr(RTrim(IIf(IsDBNull(mvarColDetails(11, pintLoop)), vbNullString, mvarColDetails(11, pintLoop)))) = CStr(RTrim(IIf(IsDBNull(vOriginalData), vbNullString, vOriginalData))) Then
								'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								vData = ""
							End If
						End If
					End If
				End If

			End If
		Next pintLoop

		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If Not IsDBNull(vData) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			vData = Replace(vData, vbNewLine, " ")
			'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			vData = Replace(vData, vbCr, " ")
			'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			vData = Replace(vData, vbLf, " ")
			'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			vData = Replace(vData, vbTab, " ")
			'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			vData = Replace(vData, Chr(10), "")
			'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			vData = Replace(vData, Chr(13), "")
		End If

		'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PopulateGrid_FormatData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PopulateGrid_FormatData = vData

	End Function

	Private Function PopulateGrid_DoSummaryInfo(ByRef pavColumns As Object, ByRef piColumnIndex As Short, ByRef piSortIndex As Short) As Boolean

		Dim fDoValue As Boolean
		Dim iLoop As Short
		Dim iLoop2 As Short
		Dim iColumnIndex As Short
		Dim sSQL As String
		Dim rsTemp As Recordset
		Dim fHasAverage As Boolean
		Dim fHasCount As Boolean
		Dim fHasTotal As Boolean
		Dim sWhereCode As String
		Dim sFromCode As String
		Dim sCountAddString As String
		Dim sAverageAddString As String
		Dim sTotalAddString As String
		Dim sTotalBradfordAddString As String
		Dim iLogicValue As Short

		Dim miAmountOfRecords As Single
		Dim sBradfordSummary As String
		Dim asBradfordSummaryLine() As String

		Dim intColCounter As Short

		Dim strAggrValue As String

		intColCounter = 1
		strAggrValue = vbNullString

		' Construct the summary where clause.
		sWhereCode = ""
		For iLoop = 1 To piSortIndex
			iColumnIndex = 0
			For iLoop2 = 1 To UBound(mvarColDetails, 2)
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(13, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(12, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (mvarColDetails(12, iLoop2).ToString() = mvarSortOrder(1, iLoop).ToString()) And (mvarColDetails(13, iLoop2) = "C") Then
					iColumnIndex = iLoop2
					Exit For
				End If
			Next iLoop2

			If iColumnIndex > 0 Then
				If mvarColDetails(7, iColumnIndex) Or mvarColDetails(8, iColumnIndex) Then
					' The column is a break/page on change column so put it in the Where clause.
					sWhereCode = sWhereCode & IIf(Len(sWhereCode) = 0, " WHERE ", " AND ")

					If (Not mvarColDetails(3, iColumnIndex)) And (Not mvarColDetails(17, iColumnIndex)) And (Not mvarColDetails(18, iColumnIndex)) Then
						' Character column. Treat empty strings along with nulls.
						If Len(pavColumns(3, iColumnIndex)) = 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sWhereCode = sWhereCode & "(([" & CStr(mvarColDetails(0, iColumnIndex)) & "] = '') OR ([" & CStr(mvarColDetails(0, iColumnIndex)) & "] IS NULL))"
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object pavColumns(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sWhereCode = sWhereCode & "([" & CStr(mvarColDetails(0, iColumnIndex)) & "] = '" & Replace(pavColumns(3, iColumnIndex), "'", "''") & "')"
						End If
					Else
						If IsDBNull(pavColumns(3, iColumnIndex)) Or pavColumns(3, iColumnIndex).ToString() = "" Then
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sWhereCode = sWhereCode & "([" & CStr(mvarColDetails(0, iColumnIndex)) & "] IS NULL)"
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(17, iColumnIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If mvarColDetails(17, iColumnIndex) Then
								' Date column.
								'UPGRADE_WARNING: Couldn't resolve default property of object pavColumns(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sWhereCode = sWhereCode & "([" & CStr(mvarColDetails(0, iColumnIndex)) & "] = '" & VB6.Format(pavColumns(3, iColumnIndex), "MM/dd/yyyy") & "')"
							Else
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(18, iColumnIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If mvarColDetails(18, iColumnIndex) Then
									' Logic Column.
									'TM20020523 Fault 3910 - if logic column then convert the stored 'Y' or 'N' to 1 or 0.
									'UPGRADE_WARNING: Couldn't resolve default property of object pavColumns(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									iLogicValue = IIf(pavColumns(3, iColumnIndex) = "Y", 1, 0)
									'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									sWhereCode = sWhereCode & "([" & CStr(mvarColDetails(0, iColumnIndex)) & "] = " & iLogicValue & ")"
								Else
									' Numeric column.
									'UPGRADE_WARNING: Couldn't resolve default property of object pavColumns(3, iColumnIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									sWhereCode = sWhereCode & "([" & CStr(mvarColDetails(0, iColumnIndex)) & "] = " & pavColumns(3, iColumnIndex) & ")"
								End If
							End If
						End If
					End If
				End If
			End If
		Next iLoop

		' Construct the required select statement.
		sSQL = ""
		sFromCode = ""
		For iLoop = 1 To UBound(mvarColDetails, 2)

			If Not mvarColDetails(0, iLoop).ToString().Substring(0, 3) = "?ID" Then

				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(4, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If mvarColDetails(4, iLoop) Then
					' Average.
					mblnReportHasSummaryInfo = True
					sAverageAddString = "<td class='hiddentablecolumn'>*average*</td><td class='summarytablerow'>Sub Average</td>"

					If Not mbIsBradfordIndexReport Then
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(20, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If mvarColDetails(20, iLoop) Then
							' JPD20020712 Fault 4156
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(15, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sSQL = sSQL & ",(SELECT AVG(convert(float,[" & mvarColDetails(0, iLoop) & "])) " & "FROM (SELECT DISTINCT [?ID_" & mvarColDetails(15, iLoop) & "], [" & mvarColDetails(0, iLoop) & "] " & "FROM " & mstrTempTableName & " " & " " & sWhereCode & " "

							If mblnIgnoreZerosInAggregates And mvarColDetails(3, iLoop) Then
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sSQL = sSQL & " AND ([" & mvarColDetails(0, iLoop) & "] <> 0) "
							End If

							sSQL = sSQL & ") AS [vt." & Str(iLoop) & "]) AS 'avg_" & Trim(Str(iLoop)) & "'"
						Else
							' JPD20020712 Fault 4156
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sSQL = sSQL & ",(SELECT AVG(convert(float, [" & mvarColDetails(0, iLoop) & "])) " & "FROM (SELECT DISTINCT [?ID], [" & mvarColDetails(0, iLoop) & "] " & "FROM " & mstrTempTableName & " " & " " & sWhereCode & " "

							If mblnIgnoreZerosInAggregates And mvarColDetails(3, iLoop) Then
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sSQL = sSQL & " AND ([" & mvarColDetails(0, iLoop) & "] <> 0) "
							End If

							sSQL = sSQL & ") AS [vt." & Str(iLoop) & "]) AS 'avg_" & Trim(Str(iLoop)) & "'"
						End If
					Else
						'Bradford Index
						' JPD20020712 Fault 4156
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sSQL = sSQL & IIf(Len(sSQL) = 0, "SELECT", ",") & " avg(convert(float,[" & mvarColDetails(0, iLoop) & "])) AS avg_" & Trim(Str(iLoop))
					End If
				End If

				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(5, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If mvarColDetails(5, iLoop) Then
					' Count.
					mblnReportHasSummaryInfo = True
					sCountAddString = "<td class='hiddentablecolumn'>*count*</td><td class='summarytablerow'>Sub Count</td>"

					If Not mbIsBradfordIndexReport Then
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(20, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If mvarColDetails(20, iLoop) Then
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(15, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sSQL = sSQL & ",(SELECT COUNT([?ID_" & mvarColDetails(15, iLoop) & "]) " & "FROM (SELECT DISTINCT [?ID_" & mvarColDetails(15, iLoop) & "], [" & mvarColDetails(0, iLoop) & "] " & "FROM " & mstrTempTableName & " " & " " & sWhereCode & " "

							If mblnIgnoreZerosInAggregates And mvarColDetails(3, iLoop) Then
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sSQL = sSQL & " AND ([" & mvarColDetails(0, iLoop) & "] <> 0) "
							End If

							sSQL = sSQL & ") AS [vt." & Str(iLoop) & "]) AS 'cnt_" & Trim(Str(iLoop)) & "'"
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sSQL = sSQL & ",(SELECT COUNT([?ID]) " & "FROM (SELECT DISTINCT [?ID], [" & mvarColDetails(0, iLoop) & "] " & "FROM " & mstrTempTableName & " " & " " & sWhereCode & " "

							If mblnIgnoreZerosInAggregates And mvarColDetails(3, iLoop) Then
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sSQL = sSQL & " AND ([" & mvarColDetails(0, iLoop) & "] <> 0) "
							End If

							sSQL = sSQL & ") AS [vt." & Str(iLoop) & "]) AS 'cnt_" & Trim(Str(iLoop)) & "'"
						End If
					Else
						'Bradford Index
						sSQL = sSQL & IIf(Len(sSQL) = 0, "SELECT", ",") & " count(*) AS cnt_" & Trim(Str(iLoop))
					End If
				End If

				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(6, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If mvarColDetails(6, iLoop) Then
					' Total.
					mblnReportHasSummaryInfo = True
					sTotalAddString = "<td class='hiddentablecolumn'>*total*</td><td class='summarytablerow'>Sub Total</td>"
					sTotalBradfordAddString = "*total*" & vbTab & "Sub Total"

					If Not mbIsBradfordIndexReport Then
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(20, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If mvarColDetails(20, iLoop) Then
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(15, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sSQL = sSQL & ",(SELECT SUM([" & mvarColDetails(0, iLoop) & "]) " & "FROM (SELECT DISTINCT [?ID_" & mvarColDetails(15, iLoop) & "], [" & mvarColDetails(0, iLoop) & "] " & "FROM " & mstrTempTableName & " " & " " & sWhereCode & " "

							If mblnIgnoreZerosInAggregates And mvarColDetails(3, iLoop) Then
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sSQL = sSQL & " AND ([" & mvarColDetails(0, iLoop) & "] <> 0) "
							End If

							sSQL = sSQL & ") AS [vt." & Str(iLoop) & "]) AS 'ttl_" & Trim(Str(iLoop)) & "'"
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sSQL = sSQL & ",(SELECT SUM([" & mvarColDetails(0, iLoop) & "]) " & "FROM (SELECT DISTINCT [?ID], [" & mvarColDetails(0, iLoop) & "] " & "FROM " & mstrTempTableName & " " & " " & sWhereCode & " "

							If mblnIgnoreZerosInAggregates And mvarColDetails(3, iLoop) Then
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sSQL = sSQL & " AND ([" & mvarColDetails(0, iLoop) & "] <> 0) "
							End If

							sSQL = sSQL & ") AS [vt." & Str(iLoop) & "]) AS 'ttl_" & Trim(Str(iLoop)) & "'"
						End If
					Else
						'Bradford Index
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sSQL = sSQL & IIf(Len(sSQL) = 0, "SELECT", ",") & " sum([" & mvarColDetails(0, iLoop) & "])  AS ttl_" & Trim(Str(iLoop))
					End If
				End If
			End If

		Next iLoop

		If Len(sSQL) > 0 Then
			If Not mbIsBradfordIndexReport Then
				sSQL = "SELECT " & Right(sSQL, Len(sSQL) - 1)
			Else
				sSQL = sSQL & " FROM " & mstrTempTableName & IIf(Len(sFromCode) > 0, sFromCode, "") & IIf(Len(sWhereCode) > 0, sWhereCode, "")
			End If

			rsTemp = datGeneral.GetRecords(sSQL)

			For iLoop = 1 To UBound(mvarColDetails, 2)
				intColCounter = intColCounter + 1

				If Not mvarColDetails(0, iLoop).ToString().Substring(0, 3) = "?ID" Then

					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(4, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If mvarColDetails(4, iLoop) Then

						If Not mvarColDetails(19, iLoop) And (Not mvarColDetails(24, iLoop)) And (Not mvarColDetails(24, iLoop - 1)) Then
							fHasAverage = True
						End If

						' Average.
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						If IsDBNull(rsTemp.Fields("avg_" & Trim(Str(iLoop))).Value) Then
							strAggrValue = "0"
							'TM20020430 Fault 3810 - if the size and decimals of the report column are zero then
							'do not format the data, show it as it is.
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ElseIf mvarColDetails(2, iLoop) = 0 And mvarColDetails(1, iLoop) = 0 Then
							strAggrValue = rsTemp.Fields("avg_" & Trim(Str(iLoop))).Value
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ElseIf mvarColDetails(1, iLoop) > 0 And mvarColDetails(2, iLoop) = 0 Then
							strAggrValue = Format(rsTemp.Fields("avg_" & Trim(Str(iLoop))).Value, "#0")
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ElseIf mvarColDetails(2, iLoop) = 0 Then
							strAggrValue = rsTemp.Fields("avg_" & Trim(Str(iLoop))).Value
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strAggrValue = Format(rsTemp.Fields("avg_" & Trim(Str(iLoop))).Value, "0." & New String("0", mvarColDetails(2, iLoop)))
						End If

						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(22, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If mvarColDetails(22, iLoop) Then
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If mvarColDetails(2, iLoop) = 0 And (InStr(1, strAggrValue, ".") <= 0) Then
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								strAggrValue = Format(strAggrValue, "#,0" & New String("0", mvarColDetails(2, iLoop)))
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							ElseIf (mvarColDetails(1, iLoop) > 0) And (mvarColDetails(2, iLoop) = 0) Then
								strAggrValue = Format(strAggrValue, "#,0")
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							ElseIf mvarColDetails(2, iLoop) = 0 Then
								strAggrValue = Format(strAggrValue, "#,0.#")
							Else
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								strAggrValue = Format(strAggrValue, "#,0." & New String("0", mvarColDetails(2, iLoop)))
							End If
						End If

						sAverageAddString = sAverageAddString & "<td class='summarytablerow' align='right'>" & strAggrValue & "</td>"

						strAggrValue = vbNullString
					Else
						'        If (mvarColDetails(24, iLoop) = False) Then
						' Display the value ?
						fDoValue = False
						If (mvarColDetails(9, iLoop)) Then
							For iLoop2 = 1 To UBound(mvarSortOrder, 2)
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(12, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(1, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If mvarSortOrder(1, iLoop2).ToString() = mvarColDetails(12, iLoop).ToString() Then
									fDoValue = (iLoop2 <= piSortIndex)
									Exit For
								End If
							Next iLoop2
						End If

						If fDoValue Then
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object PopulateGrid_FormatData(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sAverageAddString = sAverageAddString & "<td class='summarytablerow'>" & PopulateGrid_FormatData(CStr(mvarColDetails(0, iLoop)), pavColumns(3, iLoop), False, True) & "</td>"
						Else
							sAverageAddString = sAverageAddString & "<td class='summarytablerow'></td>"
						End If

					End If

					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(5, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If mvarColDetails(5, iLoop) Then

						If Not mvarColDetails(19, iLoop) And (Not mvarColDetails(24, iLoop)) And (Not mvarColDetails(24, iLoop - 1)) Then
							fHasCount = True
						End If

						'JDM - Make a note of count the Bradford Index Report
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						If mbIsBradfordIndexReport Then miAmountOfRecords = IIf(Not IsDBNull(rsTemp.Fields("cnt_" & Trim(Str(iLoop))).Value), rsTemp.Fields("cnt_" & Trim(Str(iLoop))).Value, 0)

						' Count.
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						sCountAddString = sCountAddString & "<td class='summarytablerow' align='right'>" & IIf(IsDBNull(rsTemp.Fields("cnt_" & Trim(Str(iLoop))).Value), "0", Format(rsTemp.Fields("cnt_" & Trim(Str(iLoop))).Value, "0")) & "</td>"

					Else
						'        If (mvarColDetails(24, iLoop) = False) Then
						' Display the value ?
						fDoValue = False
						If (mvarColDetails(9, iLoop)) Then
							For iLoop2 = 1 To UBound(mvarSortOrder, 2)
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(12, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(1, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If mvarSortOrder(1, iLoop2).ToString() = mvarColDetails(12, iLoop).ToString() Then
									fDoValue = (iLoop2 <= piSortIndex)
									Exit For
								End If
							Next iLoop2
						End If

						If (mbIsBradfordIndexReport And mblnCustomReportsSummaryReport) And (mbBradfordCount) Then
							fDoValue = True
						End If

						If fDoValue Then
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object PopulateGrid_FormatData(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sCountAddString = sCountAddString & "<td class='summarytablerow'>" & PopulateGrid_FormatData(CStr(mvarColDetails(0, iLoop)), pavColumns(3, iLoop), False, True) & "</td>"
						Else
							sCountAddString = sCountAddString & "<td class='summarytablerow'></td>"
						End If
						'        End If
					End If

					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(6, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If mvarColDetails(6, iLoop) Then
						' Total.

						If Not mvarColDetails(19, iLoop) And (Not mvarColDetails(24, iLoop)) And (Not mvarColDetails(24, iLoop - 1)) Then
							fHasTotal = True
						End If

						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						If IsDBNull(rsTemp.Fields("ttl_" & Trim(Str(iLoop))).Value) Then
							strAggrValue = "0"
							'TM20020430 Fault 3810 - if the size and decimals of the report column are zero then
							'do not format the data, show it as it is.
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ElseIf mvarColDetails(2, iLoop) = 0 And mvarColDetails(1, iLoop) = 0 Then
							strAggrValue = rsTemp.Fields("ttl_" & Trim(Str(iLoop))).Value
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ElseIf mvarColDetails(2, iLoop) = 0 Then
							strAggrValue = Format(rsTemp.Fields("ttl_" & Trim(Str(iLoop))).Value, "0")
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strAggrValue = Format(rsTemp.Fields("ttl_" & Trim(Str(iLoop))).Value, "0." & New String("0", mvarColDetails(2, iLoop)))
						End If

						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(22, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If mvarColDetails(22, iLoop) Then
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If mvarColDetails(2, iLoop) = 0 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								strAggrValue = Format(strAggrValue, "#,0" & New String("0", mvarColDetails(2, iLoop)))
							Else
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								strAggrValue = Format(strAggrValue, "#,0." & New String("0", mvarColDetails(2, iLoop)))
							End If
						End If

						sTotalAddString = sTotalAddString & "<td class='summarytablerow' align='right'>" & strAggrValue & "</td>"
						sTotalBradfordAddString = sTotalBradfordAddString & strAggrValue & vbTab

						strAggrValue = vbNullString

					Else
						'        If (mvarColDetails(24, iLoop) = False) Then
						' Display the value ?
						fDoValue = False
						If (mvarColDetails(9, iLoop)) Then
							For iLoop2 = 1 To UBound(mvarSortOrder, 2)
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(12, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(1, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If mvarSortOrder(1, iLoop2).ToString() = mvarColDetails(12, iLoop).ToString() Then
									fDoValue = (iLoop2 <= piSortIndex)
									Exit For
								End If
							Next iLoop2
						End If

						If (mbIsBradfordIndexReport And mblnCustomReportsSummaryReport) Then
							If Not mbBradfordCount Then
								fDoValue = True
							End If
						End If

						If fDoValue Then
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object PopulateGrid_FormatData(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sTotalAddString = sTotalAddString & "<td class='summarytablerow'>" & PopulateGrid_FormatData(CStr(mvarColDetails(0, iLoop)), pavColumns(3, iLoop), False, True) & "</td>"
						Else
							sTotalAddString = sTotalAddString & "<td class='summarytablerow'></td>"
						End If

						sTotalBradfordAddString = sTotalBradfordAddString & vbTab
					End If
				End If
			Next iLoop

			rsTemp.Close()
			'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsTemp = Nothing
		End If

		' Do a different summary if we are a Bradford Index Report
		Dim iWidthTemp As Short
		If Not mbIsBradfordIndexReport Then

			' Put a blank line in here if its not a page break as well
			'    If Not mblnCustomReportsSummaryReport And (fHasAverage Or fHasCount Or fHasTotal) Then
			If ((Not mblnCustomReportsSummaryReport) And (fHasAverage Or fHasCount Or fHasTotal)) Or ((Not mblnCustomReportsSummaryReport) And Not (fHasAverage Or fHasCount Or fHasTotal Or (mvarColDetails(8, iColumnIndex)))) Then
				AddToArray_Data("*indicator*", enum_RowType.Data)
				mintPageBreakRowIndex = mintPageBreakRowIndex + 1
			End If

			If fHasAverage Then
				AddToArray_Data(sAverageAddString, enum_RowType.Average)
				mintPageBreakRowIndex = mintPageBreakRowIndex + 1
			End If

			If fHasCount Then
				AddToArray_Data(sCountAddString, enum_RowType.Count)
				mintPageBreakRowIndex = mintPageBreakRowIndex + 1
			End If

			If fHasTotal Then
				AddToArray_Data(sTotalAddString, enum_RowType.Total)
				mintPageBreakRowIndex = mintPageBreakRowIndex + 1
			End If

			If Not mblnCustomReportsSummaryReport Then
				If (Not mvarColDetails(8, iColumnIndex)) Then
					If fHasAverage Or fHasCount Or fHasTotal Then
						AddToArray_Data("*indicator*", enum_RowType.Data)
						mintPageBreakRowIndex = mintPageBreakRowIndex + 1
					End If
				End If
			End If

		Else

			mblnReportHasSummaryInfo = True

			asBradfordSummaryLine = Split(sTotalBradfordAddString, vbTab)

			' Build Bradford Total Summary
			asBradfordSummaryLine(10) = "Total"
			asBradfordSummaryLine(12) = CStr(Val(Str(CDbl(asBradfordSummaryLine(12)))))
			asBradfordSummaryLine(13) = CStr(Val(Str(CDbl(asBradfordSummaryLine(13)))))
			sTotalAddString = Join(asBradfordSummaryLine, vbTab)

			' Calculate Bradford index line
			asBradfordSummaryLine(10) = "Bradford Factor"

			If mbBradfordWorkings = True Then
				asBradfordSummaryLine(12) = CStr(Val(asBradfordSummaryLine(12)) * (miAmountOfRecords * miAmountOfRecords)) & " (" & Str(miAmountOfRecords) & Chr(178) & " * " & asBradfordSummaryLine(12) & ")"
				asBradfordSummaryLine(13) = CStr(Val(asBradfordSummaryLine(13)) * (miAmountOfRecords * miAmountOfRecords)) & " (" & Str(miAmountOfRecords) & Chr(178) & " * " & asBradfordSummaryLine(13) & ")"
			Else
				asBradfordSummaryLine(12) = CStr(CDbl(asBradfordSummaryLine(12)) * (miAmountOfRecords * miAmountOfRecords))
				asBradfordSummaryLine(13) = CStr(CDbl(asBradfordSummaryLine(13)) * (miAmountOfRecords * miAmountOfRecords))
			End If

			If (mblnCustomReportsSummaryReport) And (mbBradfordCount Or mbBradfordTotals) Then
				asBradfordSummaryLine(2) = vbNullString
				asBradfordSummaryLine(3) = vbNullString
				asBradfordSummaryLine(4) = vbNullString
				asBradfordSummaryLine(5) = vbNullString
			End If

			sBradfordSummary = Join(asBradfordSummaryLine, vbTab)

			' Add the summary lines
			If mbBradfordCount Then
				AddToArray_Data(sCountAddString, enum_RowType.Count)
				mintPageBreakRowIndex = mintPageBreakRowIndex + 1
			End If

			If mbBradfordTotals Then
				AddToArray_Data(sTotalAddString, enum_RowType.Total)
				mintPageBreakRowIndex = mintPageBreakRowIndex + 1
			End If

			AddToArray_Data(sBradfordSummary, enum_RowType.Total)
			mintPageBreakRowIndex = mintPageBreakRowIndex + 1
			AddToArray_Data("*indicator*", enum_RowType.Data)
			mintPageBreakRowIndex = mintPageBreakRowIndex + 1

		End If

	End Function

	Private Sub PopulateGrid_DoGrandSummary()

		' Purpose : To calculate the final grand summaries
		' Input   : None
		' Output  : True/False

		On Error GoTo PopulateGrid_DoGrandSummary_ERROR

		Dim iLoop As Short
		Dim rsTemp As Recordset

		Dim sAverageAddString As String
		Dim sCountAddString As String
		Dim sTotalAddString As String

		Dim fHasAverage As Boolean
		Dim fHasCount As Boolean
		Dim fHasTotal As Boolean

		Dim sSQL As String

		Dim intColCounter As Short

		Dim strAggrValue As String

		intColCounter = 1
		strAggrValue = vbNullString

		sAverageAddString = vbNullString
		sCountAddString = vbNullString
		sTotalAddString = vbNullString

		' Construct the required select statement.
		sSQL = vbNullString

		For iLoop = 1 To UBound(mvarColDetails, 2)
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(4, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If mvarColDetails(4, iLoop) Then
				' Average.

				'TM20020718 Fault 4170 - indicate in the hidden column that the row is an average row.
				sAverageAddString = "<td class='hiddentablecolumn'>*average*</td><td class='grandsummaryrow'>Average</td>"

				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(20, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If mvarColDetails(20, iLoop) Then
					' JPD20020712 Fault 4156
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(15, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sSQL = sSQL & ",(SELECT AVG(convert(float, [" & mvarColDetails(0, iLoop) & "])) " & "FROM (SELECT DISTINCT [?ID_" & mvarColDetails(15, iLoop) & "], [" & mvarColDetails(0, iLoop) & "] " & "FROM " & mstrTempTableName & " "

					If mblnIgnoreZerosInAggregates And mvarColDetails(3, iLoop) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sSQL = sSQL & "WHERE ([" & mvarColDetails(0, iLoop) & "] <> 0) "
					End If

					sSQL = sSQL & " ) AS [vt." & Str(iLoop) & "]) AS 'avg_" & Trim(Str(iLoop)) & "'"
				Else
					' JPD20020712 Fault 4156
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sSQL = sSQL & ",(SELECT AVG(convert(float, [" & mvarColDetails(0, iLoop) & "])) " & "FROM (SELECT DISTINCT [?ID], [" & mvarColDetails(0, iLoop) & "] " & "FROM " & mstrTempTableName & " "

					If mblnIgnoreZerosInAggregates And mvarColDetails(3, iLoop) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sSQL = sSQL & "WHERE ([" & mvarColDetails(0, iLoop) & "] <> 0) "
					End If

					sSQL = sSQL & " ) AS [vt." & Str(iLoop) & "]) AS 'avg_" & Trim(Str(iLoop)) & "'"
				End If

			End If

			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(5, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If mvarColDetails(5, iLoop) Then
				' Count.

				'Add a hidden key '*count*' so that when outputting to excel it does not format the
				'count to a date.
				sCountAddString = "<td class='hiddentablecolumn'>*count*</td><td class='grandsummaryrow'>Count</td>"

				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(20, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If mvarColDetails(20, iLoop) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(15, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sSQL = sSQL & ",(SELECT COUNT([?ID_" & mvarColDetails(15, iLoop) & "]) " & "FROM (SELECT DISTINCT [?ID_" & mvarColDetails(15, iLoop) & "], [" & mvarColDetails(0, iLoop) & "] " & "FROM " & mstrTempTableName & " "

					If mblnIgnoreZerosInAggregates And mvarColDetails(3, iLoop) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sSQL = sSQL & "WHERE ([" & mvarColDetails(0, iLoop) & "] <> 0) "
					End If

					sSQL = sSQL & " ) AS [vt." & Str(iLoop) & "]) AS 'cnt_" & Trim(Str(iLoop)) & "'"
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sSQL = sSQL & ",(SELECT COUNT([?ID]) " & "FROM (SELECT DISTINCT [?ID], [" & mvarColDetails(0, iLoop) & "] " & "FROM " & mstrTempTableName & " "

					If mblnIgnoreZerosInAggregates And mvarColDetails(3, iLoop) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sSQL = sSQL & "WHERE ([" & mvarColDetails(0, iLoop) & "] <> 0) "
					End If

					sSQL = sSQL & " ) AS [vt." & Str(iLoop) & "]) AS 'cnt_" & Trim(Str(iLoop)) & "'"
				End If

			End If

			'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(6, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If mvarColDetails(6, iLoop) Then
				' Total.

				'TM20020718 Fault 4170 - indicate in the hidden column that the row is a total row.
				sTotalAddString = "<td class='hiddentablecolumn'>*total*</td><td class='grandsummaryrow'>Total</td>"

				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(20, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If mvarColDetails(20, iLoop) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(15, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sSQL = sSQL & ",(SELECT SUM([" & mvarColDetails(0, iLoop) & "]) " & "FROM (SELECT DISTINCT [?ID_" & mvarColDetails(15, iLoop) & "], [" & mvarColDetails(0, iLoop) & "] " & "FROM " & mstrTempTableName & " "

					If mblnIgnoreZerosInAggregates And mvarColDetails(3, iLoop) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sSQL = sSQL & "WHERE ([" & mvarColDetails(0, iLoop) & "] <> 0) "
					End If

					sSQL = sSQL & " ) AS [vt." & Str(iLoop) & "]) AS 'ttl_" & Trim(Str(iLoop)) & "'"
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sSQL = sSQL & ",(SELECT SUM([" & mvarColDetails(0, iLoop) & "]) " & "FROM (SELECT DISTINCT [?ID], [" & mvarColDetails(0, iLoop) & "] " & "FROM " & mstrTempTableName & " "

					If mblnIgnoreZerosInAggregates And mvarColDetails(3, iLoop) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sSQL = sSQL & "WHERE ([" & mvarColDetails(0, iLoop) & "] <> 0) "
					End If

					sSQL = sSQL & " ) AS [vt." & Str(iLoop) & "]) AS 'ttl_" & Trim(Str(iLoop)) & "'"
				End If

			End If
		Next iLoop

		If Len(sSQL) > 0 Then
			sSQL = "SELECT " & Right(sSQL, Len(sSQL) - 1)

			rsTemp = datGeneral.GetRecords(sSQL)

			For iLoop = 1 To UBound(mvarColDetails, 2)

				intColCounter = intColCounter + 1

				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(4, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If mvarColDetails(4, iLoop) Then
					' Average.

					If Not mvarColDetails(19, iLoop) And (Not mvarColDetails(24, iLoop)) And (Not mvarColDetails(24, iLoop - 1)) Then
						fHasAverage = True
					End If

					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If IsDBNull(rsTemp.Fields("avg_" & Trim(Str(iLoop))).Value) Then
						strAggrValue = "0"
						'TM20020430 Fault 3810 - if the size and decimals of the report column are zero then
						'do not format the data, show it as it is.
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ElseIf mvarColDetails(2, iLoop) = 0 And mvarColDetails(1, iLoop) = 0 Then
						strAggrValue = rsTemp.Fields("avg_" & Trim(Str(iLoop))).Value
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ElseIf mvarColDetails(1, iLoop) > 0 And mvarColDetails(2, iLoop) = 0 Then
						strAggrValue = Format(rsTemp.Fields("avg_" & Trim(Str(iLoop))).Value, "#0")
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ElseIf mvarColDetails(2, iLoop) = 0 Then
						strAggrValue = rsTemp.Fields("avg_" & Trim(Str(iLoop))).Value
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						strAggrValue = Format(rsTemp.Fields("avg_" & Trim(Str(iLoop))).Value, "0." & New String("0", mvarColDetails(2, iLoop)))
					End If

					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(22, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If mvarColDetails(22, iLoop) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If mvarColDetails(2, iLoop) = 0 And (InStr(1, strAggrValue, ".") <= 0) Then
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strAggrValue = Format(strAggrValue, "#,0" & New String("0", mvarColDetails(2, iLoop)))
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ElseIf (mvarColDetails(1, iLoop) > 0) And (mvarColDetails(2, iLoop) = 0) Then
							strAggrValue = Format(strAggrValue, "#,0")
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ElseIf mvarColDetails(2, iLoop) = 0 Then
							strAggrValue = Format(strAggrValue, "#,0.#")
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strAggrValue = Format(strAggrValue, "#,0." & New String("0", mvarColDetails(2, iLoop)))
						End If
					End If

					sAverageAddString = sAverageAddString & "<td class='grandsummaryrow' align='right'>" & strAggrValue & "</td>"

					strAggrValue = vbNullString

				Else
					sAverageAddString = sAverageAddString & "<td class='grandsummaryrow'></td>"
				End If

				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(5, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If mvarColDetails(5, iLoop) Then
					' Count.

					If Not mvarColDetails(19, iLoop) And (Not mvarColDetails(24, iLoop)) And (Not mvarColDetails(24, iLoop - 1)) Then
						fHasCount = True
					End If

					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					sCountAddString = sCountAddString & "<td class='grandsummaryrow' align='right'>" & IIf(IsDBNull(rsTemp.Fields("cnt_" & Trim(Str(iLoop))).Value), "0", Format(rsTemp.Fields("cnt_" & Trim(Str(iLoop))).Value, "0")) & "</td>"

				Else
					sCountAddString = sCountAddString & "<td class='grandsummaryrow'></td>"
				End If

				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(6, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If mvarColDetails(6, iLoop) Then
					' Total.

					If Not mvarColDetails(19, iLoop) And (Not mvarColDetails(24, iLoop)) And (Not mvarColDetails(24, iLoop - 1)) Then
						fHasTotal = True
					End If

					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If IsDBNull(rsTemp.Fields("ttl_" & Trim(Str(iLoop))).Value) Then
						strAggrValue = "0"
						'TM20020430 Fault 3810 - if the size and decimals of the report column are zero then
						'do not format the data, show it as it is.
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ElseIf mvarColDetails(2, iLoop) = 0 And mvarColDetails(1, iLoop) = 0 Then
						strAggrValue = rsTemp.Fields("ttl_" & Trim(Str(iLoop))).Value
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ElseIf mvarColDetails(2, iLoop) = 0 Then
						strAggrValue = Format(rsTemp.Fields("ttl_" & Trim(Str(iLoop))).Value, "0")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						strAggrValue = Format(rsTemp.Fields("ttl_" & Trim(Str(iLoop))).Value, "0." & New String("0", mvarColDetails(2, iLoop)))
					End If

					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(22, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If mvarColDetails(22, iLoop) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If (mvarColDetails(2, iLoop) = 0) Then
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strAggrValue = Format(strAggrValue, "#,0" & New String("0", mvarColDetails(2, iLoop)))
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strAggrValue = Format(strAggrValue, "#,0." & New String("0", mvarColDetails(2, iLoop)))
						End If
					End If

					sTotalAddString = sTotalAddString & "<td class='grandsummaryrow' align='right'>" & strAggrValue & "</td>"

					strAggrValue = vbNullString
				Else
					sTotalAddString = sTotalAddString & "<td class='grandsummaryrow'></td>"
				End If

			Next iLoop

			rsTemp.Close()
			'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsTemp = Nothing
		End If

		mblnDoesHaveGrandSummary = (fHasAverage Or fHasCount Or fHasTotal)

		'Output the 4 lines of grand aggregates (blank,AVG,CNT,TTL)
		If mblnDoesHaveGrandSummary Then
			AddToArray_Data(IIf(mblnPageBreak, "*", "*indicator*"), enum_RowType.Data)
		End If

		If fHasAverage Then
			mblnReportHasSummaryInfo = True
			AddToArray_Data(sAverageAddString, enum_RowType.Average)
			mintPageBreakRowIndex = mintPageBreakRowIndex + 1
		End If

		If fHasCount Then
			mblnReportHasSummaryInfo = True
			AddToArray_Data(sCountAddString, enum_RowType.Count)
			mintPageBreakRowIndex = mintPageBreakRowIndex + 1
		End If

		If fHasTotal Then
			mblnReportHasSummaryInfo = True
			AddToArray_Data(sTotalAddString, enum_RowType.Total)
			mintPageBreakRowIndex = mintPageBreakRowIndex + 1
		End If

		Exit Sub

PopulateGrid_DoGrandSummary_ERROR:

		'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsTemp = Nothing
		mstrErrorString = "Error while calculating grand summary." & vbNewLine & "(" & Err.Description & ")"
		Return

	End Sub

	Public Function ClearUp() As Boolean

		' Purpose : To clear all variables/recordsets/references and drops temptable
		' Input   : None
		' Output  : True/False success

		' Definition variables

		On Error GoTo ClearUp_ERROR

		Call UtilUpdateLastRun(UtilityType.utlCustomReport, mlngCustomReportID)

		mlngCustomReportID = 0
		mstrCustomReportsName = vbNullString
		mlngCustomReportsBaseTable = 0
		mstrCustomReportsBaseTableName = vbNullString
		mlngCustomReportsPickListID = 0
		mlngCustomReportsFilterID = 0
		mlngCustomReportsParent1Table = 0
		mlngCustomReportsParent1PickListID = 0
		mlngCustomReportsParent1FilterID = 0
		mlngCustomReportsParent2Table = 0
		mlngCustomReportsParent2PickListID = 0
		mlngCustomReportsParent2FilterID = 0
		mblnCustomReportsSummaryReport = False
		mblnCustomReportsPrintFilterHeader = False
		mlngSingleRecordID = 0

		miChildTablesCount = 0

		' Recordsets

		'UPGRADE_NOTE: Object mrstCustomReportsDetails may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mrstCustomReportsDetails = Nothing
		'UPGRADE_NOTE: Object mrstCustomReportsOutput may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mrstCustomReportsOutput = Nothing

		'  ' Delete the temptable if exists, and then clear the variable
		'  If Len(mstrTempTableName) > 0 Then
		'    mclsData.ExecuteSql ("IF EXISTS(SELECT * FROM sysobjects WHERE name = '" & mstrTempTableName & "') " & _
		''                      "DROP TABLE " & mstrTempTableName)
		'  End If
		datGeneral.DropUniqueSQLObject(mstrTempTableName, 3)
		mstrTempTableName = vbNullString

		' SQL strings

		mstrSQLSelect = vbNullString
		mstrSQLFrom = vbNullString
		mstrSQLWhere = vbNullString
		mstrSQLJoin = vbNullString
		mstrSQLOrderBy = vbNullString
		mstrSQL = vbNullString

		' Class references

		'UPGRADE_NOTE: Object mclsData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mclsData = Nothing
		'UPGRADE_NOTE: Object mclsGeneral may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mclsGeneral = Nothing

		' Clear the connection reference
		'Set gADOCon = Nothing

		' Arrays

		'  ReDim mvarColDetails(24, 0)
		ReDim mvarSortOrder(2, 0)
		ReDim mvarChildTables(5, 0)

		' Flags

		mblnReportHasSummaryInfo = False
		mblnReportHasPageBreak = False

		' Column Privilege arrays / collections / variables

		mstrBaseTableRealSource = vbNullString
		mstrRealSource = vbNullString
		'UPGRADE_NOTE: Object mobjTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mobjTableView = Nothing
		'UPGRADE_NOTE: Object mobjColumnPrivileges may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mobjColumnPrivileges = Nothing
		ReDim mlngTableViews(2, 0)
		ReDim mstrViews(0)
		ReDim mvarPageBreak(0)
		ReDim mvarVisibleColumns(3, 0)

		ClearUp = True
		Exit Function

ClearUp_ERROR:

		mstrErrorString = "Error clearing data." & vbNewLine & "(" & Err.Description & ")"
		ClearUp = False

	End Function

	Private Function IsRecordSelectionValid() As Boolean
		Dim sSQL As String
		Dim rsTemp As Recordset
		Dim iResult As RecordSelectionValidityCodes
		Dim fCurrentUserIsSysSecMgr As Boolean
		Dim i As Short
		Dim lngFilterID As Integer

		fCurrentUserIsSysSecMgr = CurrentUserIsSysSecMgr()

		' Base Table First
		If mlngSingleRecordID = 0 Then
			If mlngCustomReportsFilterID > 0 Then
				iResult = ValidateRecordSelection(RecordSelectionTypes.REC_SEL_FILTER, mlngCustomReportsFilterID)
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
			ElseIf mlngCustomReportsPickListID > 0 Then
				iResult = ValidateRecordSelection(RecordSelectionTypes.REC_SEL_PICKLIST, mlngCustomReportsPickListID)
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
		End If

		If Len(mstrErrorString) = 0 Then
			' Parent 1 Table
			If mlngCustomReportsParent1FilterID > 0 Then
				iResult = ValidateRecordSelection(RecordSelectionTypes.REC_SEL_FILTER, mlngCustomReportsParent1FilterID)
				Select Case iResult
					Case RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
						mstrErrorString = "The first parent table filter used in this definition has been deleted by another user."
					Case RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
						mstrErrorString = "The first parent table filter used in this definition is invalid."
					Case RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
						If Not fCurrentUserIsSysSecMgr Then
							mstrErrorString = "The first parent table filter used in this definition has been made hidden by another user."
						End If
				End Select
			ElseIf mlngCustomReportsParent1PickListID > 0 Then
				iResult = ValidateRecordSelection(RecordSelectionTypes.REC_SEL_PICKLIST, mlngCustomReportsParent1PickListID)
				Select Case iResult
					Case RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
						mstrErrorString = "The first parent table picklist used in this definition has been deleted by another user."
					Case RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
						mstrErrorString = "The first parent table picklist used in this definition is invalid."
					Case RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
						If Not fCurrentUserIsSysSecMgr Then
							mstrErrorString = "The first parent table picklist used in this definition has been made hidden by another user."
						End If
				End Select
			End If
		End If

		' Parent 2 Table
		If Len(mstrErrorString) = 0 Then
			If mlngCustomReportsParent2FilterID > 0 Then
				iResult = ValidateRecordSelection(RecordSelectionTypes.REC_SEL_FILTER, mlngCustomReportsParent2FilterID)
				Select Case iResult
					Case RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
						mstrErrorString = "The second parent table filter used in this definition has been deleted by another user."
					Case RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
						mstrErrorString = "The second parent table filter used in this definition is invalid."
					Case RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
						If Not fCurrentUserIsSysSecMgr Then
							mstrErrorString = "The second parent table filter used in this definition has been made hidden by another user."
						End If
				End Select
			ElseIf mlngCustomReportsParent2PickListID > 0 Then
				iResult = ValidateRecordSelection(RecordSelectionTypes.REC_SEL_PICKLIST, mlngCustomReportsParent2PickListID)
				Select Case iResult
					Case RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
						mstrErrorString = "The second parent table picklist used in this definition has been deleted by another user."
					Case RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
						mstrErrorString = "The second parent table picklist used in this definition is invalid."
					Case RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
						If Not fCurrentUserIsSysSecMgr Then
							mstrErrorString = "The second parent table picklist used in this definition has been made hidden by another user."
						End If
				End Select
			End If
		End If

		' Child Table
		If Len(mstrErrorString) = 0 Then
			If miChildTablesCount > 0 Then
				For i = 0 To UBound(mvarChildTables, 2) Step 1
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(1, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					lngFilterID = mvarChildTables(1, i)
					If lngFilterID > 0 Then
						iResult = ValidateRecordSelection(RecordSelectionTypes.REC_SEL_FILTER, lngFilterID)
						Select Case iResult
							Case RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
								mstrErrorString = "The child table filter used in this definition has been deleted by another user."
							Case RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
								mstrErrorString = "The child table filter used in this definition is invalid."
							Case RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
								If Not fCurrentUserIsSysSecMgr Then
									mstrErrorString = "The child table filter used in this definition has been made hidden by another user."
								End If
						End Select
					End If

					If Len(mstrErrorString) > 0 Then
						Exit For
					End If
				Next i
			End If
		End If

		' JDM - 13/10/03 - Fault 7228 - Problems if somehow a customreportid of 0 gets in the database.
		If Not mbIsBradfordIndexReport Then

			'******* Check calculations for hidden/deleted elements *******
			If Len(mstrErrorString) = 0 Then
				sSQL = "SELECT * FROM ASRSYSCustomReportsDetails " & "WHERE CustomReportID = " & mlngCustomReportID & " AND LOWER(Type) = 'e' "

				rsTemp = mclsGeneral.GetRecords(sSQL)
				With rsTemp
					If Not (.EOF And .BOF) Then
						.MoveFirst()
						Do Until .EOF
							iResult = ValidateCalculation(.Fields("ColExprID").Value)
							Select Case iResult
								Case RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
									mstrErrorString = "A calculation used in this definition has been deleted by another user."
								Case RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
									mstrErrorString = "A calculation used in this definition is invalid."
								Case RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
									If Not fCurrentUserIsSysSecMgr Then
										mstrErrorString = "A calculation used in this definition has been made hidden by another user."
									End If
							End Select

							If Len(mstrErrorString) > 0 Then
								Exit Do
							End If

							.MoveNext()
						Loop
					End If
				End With

				'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				rsTemp = Nothing
			End If
		End If

		IsRecordSelectionValid = (Len(mstrErrorString) = 0)

	End Function

	Private Function CheckCalcsStillExist() As Boolean

		Dim pstrBadCalcs As String
		Dim prstTemp As Recordset

		On Error GoTo Check_ERROR

		Do Until mrstCustomReportsDetails.EOF
			If mrstCustomReportsDetails.Fields("Type").Value = "E" Then
				prstTemp = mclsGeneral.GetReadOnlyRecords("SELECT * FROM AsrSysExpressions WHERE ExprID = " & mrstCustomReportsDetails.Fields("ColExprID").Value)
				If prstTemp.BOF And prstTemp.EOF Then
					pstrBadCalcs = "One or more calculation(s) used in this report have been deleted" & vbNewLine & "by another user."
					Exit Do
				End If
			End If
			mrstCustomReportsDetails.MoveNext()
		Loop

		If Len(pstrBadCalcs) > 0 Then
			mstrErrorString = pstrBadCalcs
			CheckCalcsStillExist = False
			Exit Function
		End If

		CheckCalcsStillExist = True
		mrstCustomReportsDetails.MoveFirst()
		Exit Function

Check_ERROR:

		mstrErrorString = "Error checking if calcs still exist." & vbNewLine & Err.Description
		CheckCalcsStillExist = False

	End Function

	Private Function AddToArray_Columns(ByRef pstrRowToAdd As String) As Boolean

		On Error GoTo AddError

		ReDim Preserve mvarOutputArray_Columns(UBound(mvarOutputArray_Columns) + 1)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarOutputArray_Columns(UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarOutputArray_Columns(UBound(mvarOutputArray_Columns)) = pstrRowToAdd

		AddToArray_Columns = True
		Exit Function

AddError:

		AddToArray_Columns = False
		mstrErrorString = "Error adding to columns array:" & vbNewLine & Err.Description

	End Function

	Public Function OutputGridDefinition() As Boolean

		Dim pblnOK As Boolean
		Dim sCaption As String

		On Error GoTo ErrTrap

		pblnOK = True

		sCaption = mstrCustomReportsName

		If mblnCustomReportsPrintFilterHeader And (mlngSingleRecordID = 0) Then
			If (mlngCustomReportsFilterID > 0) Then
				sCaption = sCaption & " (Base Table filter : " & datGeneral.GetFilterName(mlngCustomReportsFilterID) & ")"
			ElseIf (mlngCustomReportsPickListID > 0) Then
				sCaption = sCaption & " (Base Table picklist : " & datGeneral.GetPicklistName(mlngCustomReportsPickListID) & ")"
			Else
				sCaption = sCaption & " (No Picklist or Filter Selected)"
			End If
		End If

		OutputGridDefinition = pblnOK
		Exit Function

ErrTrap:

		OutputGridDefinition = False
		mstrErrorString = "Error with OutputGridDefinition: " & vbNewLine & Err.Description

	End Function

	Public Function OutputGridColumns() As Boolean

		On Error GoTo ErrTrap

		Dim iLoop As Short
		Dim pblnOK As Boolean
		Dim intColCounter As Short

		pblnOK = True

		'Pagebreak
		intColCounter = 0

		AddToArray_Columns("<th class='hiddentablecolumn'>PageBreak</th>")

		'Summary Info
		intColCounter = intColCounter + 1
		AddToArray_Columns("<th class='summarytablecolumn'>Summary Info</th>")

		' Now loop through the recordset fields, adding the data columns
		For iLoop = 0 To (mrstCustomReportsOutput.Fields.Count - 1)

			intColCounter = intColCounter + 1

			If Not mvarColDetails(24, iLoop + 1) Then
				If (mrstCustomReportsOutput.Fields(iLoop).Name.Substring(0, 1) = "?" Or (mbIsBradfordIndexReport And iLoop > 12)) Then
					pblnOK = AddToArray_Columns("<th class='hiddentablecolumn'>" & Replace(Replace(mrstCustomReportsOutput.Fields(iLoop).Name, "_", " "), """", "&quot;") & "</th>")
				Else
					pblnOK = AddToArray_Columns("<th class='tablecolumn'>" & Replace(Replace(mrstCustomReportsOutput.Fields(iLoop).Name, "_", " "), """", "&quot;") & "</th>")
				End If
			End If

		Next iLoop

		OutputGridColumns = True

		Exit Function

ErrTrap:

		OutputGridColumns = False
		mstrErrorString = "Error with OutputGridColumns: " & vbNewLine & Err.Description

	End Function

	Private Function AddToArray_Data(ByRef pstrRowToAdd As String, rowType As enum_RowType) As Boolean

		On Error GoTo AddError

		Dim sClassName As String = "rowdata"

		Select Case rowType
			Case enum_RowType.Count, enum_RowType.Average, enum_RowType.Total
				sClassName = "rowsummary"

		End Select

		If pstrRowToAdd = "<blank>" Then
			mvarOutputArray_Data.Add("<tr></tr>")

		ElseIf pstrRowToAdd = "*" Then
			mvarOutputArray_Data.Add("<tr><td>*</td></tr>")

		ElseIf pstrRowToAdd = vbNullString Then
			mvarOutputArray_Data.Add("<tr></tr>")

		Else
			mvarOutputArray_Data.Add("<tr class='" & sClassName & "'>" & pstrRowToAdd & "</tr>")

		End If

		AddToArray_Data = True
		Exit Function

AddError:

		AddToArray_Data = False
		mstrErrorString = "Error adding to data array additem:" & vbNewLine & Err.Description

	End Function

	Public Function GenerateSQLBradford(ByRef pstrIncludeTypes As String) As Boolean

		' NOTE: Checks are made elsewhere to ensure that from and to dates are not blank
		' NOTE: Put in some code to handle blank end dates (do we include as an option on the main screen ?)

		On Error GoTo GenerateSQLBradford_ERROR
		Dim strAbsenceType As String
		Dim iCount As Short
		Dim astrIncludeTypes() As String
		Dim iUbound As Integer

		' Get the absence start/end field details
		'    strAbsenceStartField = mstrRealSource + "." + datGeneral.GetColumnName(Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESTARTDATE)))
		'    strAbsenceEndField = mstrRealSource + "." + datGeneral.GetColumnName(Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEENDDATE)))
		strAbsenceType = mstrAbsenceRealSource & "." & datGeneral.GetColumnName(Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPE)))

		' Force the inputted string into an array
		'UPGRADE_WARNING: Couldn't resolve default property of object pstrIncludeTypes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		astrIncludeTypes = Split(pstrIncludeTypes, ",")

		' Add the different reason types
		If UBound(astrIncludeTypes) > 0 Then
			mstrSQLWhere = IIf(mstrSQLWhere = vbNullString, "WHERE (", mstrSQLWhere & " AND (") & "UPPER(" & strAbsenceType & ") IN ("
			For iCount = 0 To UBound(astrIncludeTypes) - 1
				astrIncludeTypes(iCount) = Replace(astrIncludeTypes(iCount), "'", "''")
				mstrSQLWhere = mstrSQLWhere & "'" & UCase(astrIncludeTypes(iCount)) & "'"
				mstrSQLWhere = mstrSQLWhere & IIf(Not iCount = UBound(astrIncludeTypes) - 1, ",", "")
			Next iCount
			mstrSQLWhere = mstrSQLWhere & "))"
		End If

		' Add the ID to the select string
		' This is needed to re-calculate the duration amounts
		mstrSQLSelect = mstrSQLSelect & "," & mstrSQLFrom & ".ID AS 'Personnel_ID'," & mstrAbsenceRealSource & ".ID as 'Absence_ID'"

		' Redimension arrays (to handle the ID fields (Personel/absnece)
		ReDim Preserve mvarColDetails(UBound(mvarColDetails, 1), UBound(mvarColDetails, 2) + 2)

		'Personel ID
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, UBound() - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(0, UBound(mvarColDetails, 2) - 1) = "Personnel_ID"
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(1, UBound() - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(1, UBound(mvarColDetails, 2) - 1) = 99
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, UBound() - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(2, UBound(mvarColDetails, 2) - 1) = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(3, UBound() - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(3, UBound(mvarColDetails, 2) - 1) = False
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(4, UBound() - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(4, UBound(mvarColDetails, 2) - 1) = False
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(5, UBound() - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(5, UBound(mvarColDetails, 2) - 1) = False
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(6, UBound() - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(6, UBound(mvarColDetails, 2) - 1) = False
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(7, UBound() - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(7, UBound(mvarColDetails, 2) - 1) = True
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(8, UBound() - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(8, UBound(mvarColDetails, 2) - 1) = False
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(9, UBound() - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(9, UBound(mvarColDetails, 2) - 1) = True
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(10, UBound() - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(10, UBound(mvarColDetails, 2) - 1) = False
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(11, UBound() - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(11, UBound(mvarColDetails, 2) - 1) = ""
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(12, UBound() - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(12, UBound(mvarColDetails, 2) - 1) = -1
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(13, UBound() - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(13, UBound(mvarColDetails, 2) - 1) = "C"
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(14, UBound() - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(14, UBound(mvarColDetails, 2) - 1) = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(15, UBound() - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(15, UBound(mvarColDetails, 2) - 1) = ""
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(16, UBound() - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(16, UBound(mvarColDetails, 2) - 1) = ""
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(17, UBound() - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(17, UBound(mvarColDetails, 2) - 1) = False
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(18, UBound() - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(18, UBound(mvarColDetails, 2) - 1) = False
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(19, UBound() - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(19, UBound(mvarColDetails, 2) - 1) = True ' Is column hidden
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(20, UBound() - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(20, UBound(mvarColDetails, 2) - 1) = False
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(21, UBound() - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(21, UBound(mvarColDetails, 2) - 1) = False

		'Absence ID
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(0, UBound(mvarColDetails, 2)) = "Absence_ID"
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(1, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(1, UBound(mvarColDetails, 2)) = 99
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(2, UBound(mvarColDetails, 2)) = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(3, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(3, UBound(mvarColDetails, 2)) = False
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(4, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(4, UBound(mvarColDetails, 2)) = False
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(5, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(5, UBound(mvarColDetails, 2)) = False
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(6, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(6, UBound(mvarColDetails, 2)) = False
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(7, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(7, UBound(mvarColDetails, 2)) = True
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(8, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(8, UBound(mvarColDetails, 2)) = False
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(9, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(9, UBound(mvarColDetails, 2)) = True
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(10, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(10, UBound(mvarColDetails, 2)) = False
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(11, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(11, UBound(mvarColDetails, 2)) = ""
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(12, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(12, UBound(mvarColDetails, 2)) = -1
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(13, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(13, UBound(mvarColDetails, 2)) = "C"
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(14, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(14, UBound(mvarColDetails, 2)) = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(15, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(15, UBound(mvarColDetails, 2)) = ""
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(16, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(16, UBound(mvarColDetails, 2)) = ""
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(17, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(17, UBound(mvarColDetails, 2)) = False
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(18, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(18, UBound(mvarColDetails, 2)) = False
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(19, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(19, UBound(mvarColDetails, 2)) = True ' Is column hidden
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(20, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(20, UBound(mvarColDetails, 2)) = False
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(21, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarColDetails(21, UBound(mvarColDetails, 2)) = False

		iUbound = UBound(mvarSortOrder, 2)
		ReDim Preserve mvarSortOrder(2, iUbound)
		mvarSortOrder(0, iUbound) = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(1, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarSortOrder(1, iUbound) = -1
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(2, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarSortOrder(2, iUbound) = "ASC"

		' All done correctly
		'UPGRADE_WARNING: Couldn't resolve default property of object GenerateSQLBradford. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GenerateSQLBradford = True
		Exit Function


GenerateSQLBradford_ERROR:

		'UPGRADE_WARNING: Couldn't resolve default property of object GenerateSQLBradford. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GenerateSQLBradford = False
		mstrErrorString = "Error in GenerateSQLBradford." & vbNewLine & Err.Description

	End Function

	Public Function CalculateBradfordFactors() As Boolean

		' Purpose : To calculate any bradford factors, and place into the created temporary table
		Dim sSQL As String

		On Error GoTo CalculateBradfordFactors_ERROR

		' Merge the absence records if the continuous field is defined.
		If Not Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCECONTINUOUS)) = CDbl("0") Then
			sSQL = "EXECUTE sp_ASR_Bradford_MergeAbsences '" & mstrBradfordStartDate & "','" & mstrBradfordEndDate & "','" & mstrTempTableName & "'"
			mclsData.ExecuteSql(sSQL)
		End If

		' Delete unwanted absences from the table.
		sSQL = "EXECUTE sp_ASR_Bradford_DeleteAbsences '" & mstrBradfordStartDate & "','" & mstrBradfordEndDate & "'," + IIf(mbOmitBeforeStart, "1,", "0,") + IIf(mbOmitAfterEnd, "1,'", "0,'") + mstrTempTableName + "'"
		mclsData.ExecuteSql(sSQL)

		' Calculate the included durations for the absences.
		sSQL = "EXECUTE sp_ASR_Bradford_CalculateDurations '" & mstrBradfordStartDate & "','" & mstrBradfordEndDate & "','" & mstrTempTableName & "'"
		mclsData.ExecuteSql(sSQL)

		' Remove absences that are below the required Bradford Factor
		If mbMinBradford Then
			sSQL = "DELETE FROM " & mstrTempTableName & " WHERE personnel_id IN (SELECT personnel_id FROM " & mstrTempTableName & " GROUP BY personnel_id HAVING((count(duration)*count(duration))*sum(duration)) < " & Str(mlngMinBradfordAmount) & ")"
			mclsData.ExecuteSql(sSQL)
		End If

		CalculateBradfordFactors = True
		Exit Function

CalculateBradfordFactors_ERROR:

		mstrErrorString = "Error while checking calculating Bradford factors." & vbNewLine & "(" & Err.Description & ")"
		CalculateBradfordFactors = False

	End Function

	' Dates are in SQL (American format)
	Public Function GetBradfordReportDefinition(ByRef pstrAbsenceFrom As String, ByRef pstrAbsenceTo As String) As Boolean

		' Purpose : This function retrieves the basic definition details
		'           and stores it in module level variables

		On Error GoTo GetBradfordReportDefinition_ERROR

		mbIsBradfordIndexReport = True

		SetupTablesCollection()

		' Dates coming in are in American format (if they're not we have a problem)
		mstrBradfordStartDate = pstrAbsenceFrom
		mstrBradfordEndDate = pstrAbsenceTo

		'JPD 20041214 - ensure no injection can take place.
		mstrBradfordStartDate = Replace(mstrBradfordStartDate, "'", "''")
		mstrBradfordEndDate = Replace(mstrBradfordEndDate, "'", "''")

		If DateDiff(DateInterval.Day, ConvertSQLDateToLocale(pstrAbsenceFrom), ConvertSQLDateToLocale(pstrAbsenceTo)) < 0 Then
			mstrErrorString = "The report end date is before the report start date."
			mobjEventLog.AddDetailEntry(mstrErrorString)
			mobjEventLog.ChangeHeaderStatus(EventLog_Status.elsFailed)
			GetBradfordReportDefinition = False
			Exit Function
		End If


		'Set the grid header with no picklist/filter information
		mstrCustomReportsName = "Bradford Factor Report (" & ConvertSQLDateToLocale(mstrBradfordStartDate) & " - " & ConvertSQLDateToLocale(mstrBradfordEndDate) & ")"

		mlngCustomReportsBaseTable = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE))
		mstrCustomReportsBaseTableName = datGeneral.GetTableName(mlngCustomReportsBaseTable)
		mlngCustomReportsParent1Table = 0
		mlngCustomReportsParent1FilterID = 0
		mlngCustomReportsParent2Table = 0
		mlngCustomReportsParent2FilterID = 0

		ReDim Preserve mvarChildTables(5, 0)

		'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(0, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarChildTables(0, 0) = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETABLE))	'Childs Table ID
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(1, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarChildTables(1, 0) = 0	'Childs Filter ID (if any)
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(2, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarChildTables(2, 0) = 0	'Number of records to take from child
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(3, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarChildTables(3, 0) = datGeneral.GetTableName(mvarChildTables(0, 0)) 'Child Table Name
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(4, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarChildTables(4, 0) = True 'Boolean - True if table is used, False if not
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(5, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarChildTables(5, 0) = 0

		miChildTablesCount = 1
		'****************************************

		mblnCustomReportsSummaryReport = False
		mlngCustomReportsParent1PickListID = 0
		mlngCustomReportsParent2PickListID = 0

		If Not IsRecordSelectionValid() Then
			GetBradfordReportDefinition = False
			Exit Function
		End If

		GetBradfordReportDefinition = True
		mobjEventLog.AddHeader(EventLog_Type.eltStandardReport, "Bradford Factor")

TidyAndExit:

		Exit Function

GetBradfordReportDefinition_ERROR:

		GetBradfordReportDefinition = False
		mstrErrorString = "Error whilst reading the Bradford Factor Report definition !" & vbNewLine & Err.Description
		mobjEventLog.AddDetailEntry(mstrErrorString)
		mobjEventLog.ChangeHeaderStatus(EventLog_Status.elsFailed)
		Resume TidyAndExit

	End Function

	Public Function GetBradfordRecordSet() As Boolean

		' Purpose : This function loads report details and sort details into
		'           arrays and leaves the details recordset reference there
		'           (dont remove it...used for summary info !)

		On Error GoTo GetBradfordRecordSet_ERROR

		Dim intTemp As Short
		Dim lngTableID As Integer
		Dim iCount As Short
		Dim lngColumnID As Integer
		Dim sMask As String

		Dim lbHideStaffNumber As Boolean

		Dim aStrRequiredFields(15, 1) As String

		aStrRequiredFields(1, 1) = CStr(Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_EMPLOYEENUMBER)))
		aStrRequiredFields(2, 1) = CStr(Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_SURNAME)))
		aStrRequiredFields(3, 1) = CStr(Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_FORENAME)))
		aStrRequiredFields(4, 1) = CStr(Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_DEPARTMENT)))

		aStrRequiredFields(5, 1) = CStr(Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPE)))
		aStrRequiredFields(6, 1) = CStr(Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESTARTDATE)))
		aStrRequiredFields(7, 1) = CStr(Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESTARTSESSION)))
		aStrRequiredFields(8, 1) = CStr(Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEENDDATE)))
		aStrRequiredFields(9, 1) = CStr(Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEENDSESSION)))
		aStrRequiredFields(10, 1) = CStr(Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEREASON)))
		aStrRequiredFields(11, 1) = CStr(Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCECONTINUOUS)))
		aStrRequiredFields(12, 1) = CStr(Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEDURATION)))

		'This field is later recalculated for the included days
		aStrRequiredFields(13, 1) = CStr(Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEDURATION)))

		'****************************************************************************
		If mlngOrderByColumnID > 0 Then
			aStrRequiredFields(14, 1) = CStr(mlngOrderByColumnID)
		Else
			aStrRequiredFields(14, 1) = CStr(-1)
		End If

		If mlngGroupByColumnID > 0 Then
			aStrRequiredFields(15, 1) = CStr(mlngGroupByColumnID)
		Else
			aStrRequiredFields(15, 1) = CStr(-1)
		End If
		'****************************************************************************

		' Allow the staff number to be undefined (Let system read the surname field)
		lbHideStaffNumber = False
		If aStrRequiredFields(1, 1) = "0" Then
			aStrRequiredFields(1, 1) = CStr(Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_SURNAME)))
			lbHideStaffNumber = True
		End If

		' Allow the continuous field to be undefined (Let system read the absence reason)
		If aStrRequiredFields(11, 1) = "0" Then
			aStrRequiredFields(11, 1) = CStr(Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEDURATION)))
		End If

		' Ensure that module setup has been run
		For iCount = 1 To UBound(aStrRequiredFields, 1)
			If aStrRequiredFields(iCount, 1) = "0" Then
				GetBradfordRecordSet = False
				mstrErrorString = "Module setup has not been completed."
				Exit Function
			End If
		Next iCount

		mblnCustomReportsSummaryReport = (Not mbDisplayBradfordDetail)

		' Load the field list
		Dim objExpr As clsExprExpression
		For iCount = 1 To UBound(aStrRequiredFields, 1)

			If CDbl(aStrRequiredFields(iCount, 1)) <> -1 Then

				intTemp = UBound(mvarColDetails, 2) + 1
				ReDim Preserve mvarColDetails(UBound(mvarColDetails, 1), intTemp)

				ReDim Preserve mstrExcelFormats(intTemp) 'MH20010307

				lngColumnID = CInt(aStrRequiredFields(iCount, 1))
				lngTableID = GetTableIDFromColumn(lngColumnID)

				' Specify the column names and whether they are visible or not
				Select Case intTemp
					Case 1
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarColDetails(0, intTemp) = "Staff_No"
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(19, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarColDetails(19, intTemp) = lbHideStaffNumber
					Case 2
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarColDetails(0, intTemp) = "Surname"
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(19, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarColDetails(19, intTemp) = False
					Case 3
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarColDetails(0, intTemp) = "Forenames"
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(19, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarColDetails(19, intTemp) = False
					Case 4
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarColDetails(0, intTemp) = "Department"
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(19, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarColDetails(19, intTemp) = False
					Case 5
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarColDetails(0, intTemp) = "Type"
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(19, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarColDetails(19, intTemp) = Not mbDisplayBradfordDetail
					Case 6
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarColDetails(0, intTemp) = "Start_Date"
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(19, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarColDetails(19, intTemp) = Not mbDisplayBradfordDetail
					Case 7
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarColDetails(0, intTemp) = "Start_Session"
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(19, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarColDetails(19, intTemp) = Not mbDisplayBradfordDetail
					Case 8
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarColDetails(0, intTemp) = "End_Date"
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(19, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarColDetails(19, intTemp) = Not mbDisplayBradfordDetail
					Case 9
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarColDetails(0, intTemp) = "End_Session"
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(19, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarColDetails(19, intTemp) = Not mbDisplayBradfordDetail
					Case 10
						If mbDisplayBradfordDetail Then
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mvarColDetails(0, intTemp) = "Reason"
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(19, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mvarColDetails(19, intTemp) = False
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mvarColDetails(0, intTemp) = "Summary Info"
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(19, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mvarColDetails(19, intTemp) = False
						End If
					Case 11
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarColDetails(0, intTemp) = "Continuous"
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(19, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarColDetails(19, intTemp) = Not mbDisplayBradfordDetail
					Case 12
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarColDetails(0, intTemp) = "Duration"
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(19, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarColDetails(19, intTemp) = False
					Case 13
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarColDetails(0, intTemp) = "Included_Days"
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(19, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarColDetails(19, intTemp) = False

						'**********************************************************************
					Case 14
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarColDetails(0, intTemp) = "Order_1"
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(19, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarColDetails(19, intTemp) = True

					Case 15
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarColDetails(0, intTemp) = "Order_2"
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(19, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarColDetails(19, intTemp) = True
						'**********************************************************************

					Case Else

						'MH20020521 Fault 3820
						'mvarColDetails(0, intTemp) = datGeneral.GetColumnName(lngColumnID)
						If lngTableID = mlngCustomReportsBaseTable Then
							'Personnel
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mvarColDetails(0, intTemp) = mstrSQLFrom & "." & datGeneral.GetColumnName(lngColumnID)
						Else
							'Absence
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mvarColDetails(0, intTemp) = mstrRealSource & "." & datGeneral.GetColumnName(lngColumnID)
						End If

				End Select

				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(1, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(1, intTemp) = 99
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(2, intTemp) = IIf(intTemp = 12 Or intTemp = 13, 1, 0) 'Decimals
				'JDM - 02/07/01 - Fault 2144 - Needs to know if we're numeric or not.
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(3, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(3, intTemp) = IIf(datGeneral.GetDataType(lngTableID, lngColumnID) = 2, True, False) 'Is Numeric
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(4, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(4, intTemp) = False 'Average
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(5, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(5, intTemp) = IIf(intTemp = 12 Or intTemp = 13, True, False)	'Count
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(6, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(6, intTemp) = IIf(intTemp = 12 Or intTemp = 13, True, False)	'Total
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(7, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(7, intTemp) = False 'Break on change
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(8, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(8, intTemp) = False 'Page break on change
				If mblnCustomReportsSummaryReport Then
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(9, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(9, intTemp) = True	'Value on change
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(9, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(9, intTemp) = False 'Value on change
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(10, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(10, intTemp) = IIf(intTemp < 5 And mbBradfordSRV, True, False)	'Suppress repeated values
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(11, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(11, intTemp) = ""
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(12, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(12, intTemp) = lngColumnID

				' Set the expression/column type of this column
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(13, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(13, intTemp) = "C"

				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(14, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(14, intTemp) = lngTableID
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(15, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(15, intTemp) = datGeneral.GetTableName(CInt(lngTableID))

				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(13, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If mvarColDetails(13, intTemp) = "C" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(16, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(16, intTemp) = datGeneral.GetColumnName(CInt(mvarColDetails(12, intTemp)))

					'MH20010307
					Select Case mvarColDetails(12, intTemp)
						Case SQLDataType.sqlNumeric, SQLDataType.sqlInteger
							'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mstrExcelFormats(intTemp) = "0" & IIf(mvarColDetails(2, intTemp) > 0, "." & New String("0", mvarColDetails(2, intTemp)), "")
						Case SQLDataType.sqlDate
							mstrExcelFormats(intTemp) = DateFormat()
						Case Else
							mstrExcelFormats(intTemp) = "@"
					End Select



				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(16, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(16, intTemp) = ""

					'MH20010307
					objExpr = New clsExprExpression

					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					objExpr.ExpressionID = CInt(mvarColDetails(12, intTemp))
					objExpr.ConstructExpression()

					Select Case objExpr.ReturnType
						Case ExpressionValueTypes.giEXPRVALUE_NUMERIC
							mstrExcelFormats(intTemp) = "0.####"
						Case ExpressionValueTypes.giEXPRVALUE_DATE
							mstrExcelFormats(intTemp) = DateFormat()
						Case Else
							mstrExcelFormats(intTemp) = "@"
					End Select

					'UPGRADE_NOTE: Object objExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					objExpr = Nothing

				End If

				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(17, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(17, intTemp) = datGeneral.DateColumn("C", lngTableID, lngColumnID)	'??? - check these out 22/03/01
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(18, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(18, intTemp) = datGeneral.BitColumn("C", lngTableID, lngColumnID)

				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(22, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarColDetails(22, intTemp) = datGeneral.DoesColumnUseSeparators(lngColumnID)	'Does this column use 1000 separators?

				'Adjust the size of the field if digit separator is used
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(22, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If mvarColDetails(22, intTemp) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(1, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarColDetails(1, intTemp) = mvarColDetails(1, intTemp) + Int((mvarColDetails(1, intTemp) - mvarColDetails(2, intTemp)) / 3)
				End If

				' Format for this numeric column
				If mvarColDetails(3, intTemp) Then
					sMask = ""
					If mvarColDetails(22, intTemp) Then sMask = ",0"
					If mvarColDetails(2, intTemp) > 0 Then sMask = sMask & "." & New String("0", mvarColDetails(2, intTemp))
					mvarColDetails(23, intTemp) = "{0:0" & sMask & ";-0" & sMask & ";0." & New String("0", mvarColDetails(2, intTemp)) & "}"
				End If

			End If

		Next iCount

		' Get those columns defined as a SortOrder and load into array
		ReDim mvarSortOrder(2, 3)

		'Employee surname
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(1, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarSortOrder(1, 1) = mstrOrderByColumn	'Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_SURNAME))
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(2, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarSortOrder(2, 1) = "Asc"

		'Employee forename
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(1, 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarSortOrder(1, 2) = mstrGroupByColumn	'Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_FORENAME))
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(2, 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarSortOrder(2, 2) = "Asc"

		'    ' Absence start date
		'    mvarSortOrder(1, 3) = "Start_Date" 'Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESTARTDATE))"
		'    mvarSortOrder(2, 3) = "Asc"

		' Force duration and included days to be numeric format in Excel
		iCount = 11 - IIf(lbHideStaffNumber = True, 1, 0)
		mstrExcelFormats(iCount) = "0.0"
		mstrExcelFormats(iCount + 1) = "0.0"

		GetBradfordRecordSet = True
		Exit Function

GetBradfordRecordSet_ERROR:

		GetBradfordRecordSet = False
		mstrErrorString = "Error whilst retrieving the details recordsets'." & vbNewLine & Err.Description

	End Function

	Public Function SetBradfordDisplayOptions(ByVal pbSRV As Boolean, ByRef pbShowTotals As Boolean, ByVal pbShowCount As Boolean, ByVal pbShowWorkings As Boolean _
																						, ByVal pbShowBasePicklistFilter As Boolean, ByVal pbDisplayBradfordDetail As Boolean) As Boolean

		' Set Report Display Options
		mbBradfordSRV = pbSRV
		mbBradfordTotals = pbShowTotals
		mbBradfordCount = pbShowCount
		mbBradfordWorkings = pbShowWorkings
		mblnCustomReportsPrintFilterHeader = pbShowBasePicklistFilter
		mbDisplayBradfordDetail = pbDisplayBradfordDetail

		Return True

	End Function

	Public Function SetBradfordOrders(ByVal pstrOrderBy As String, ByVal pstrGroupBy As String, ByVal pbOrder1Asc As Boolean, ByRef pbOrder2Asc As Boolean _
																		, ByVal plngOrderByColumnID As Long, ByVal plngGroupByColumnID As Long) As Boolean

		' Set Report Order Options
		mstrOrderByColumn = pstrOrderBy
		mstrGroupByColumn = pstrGroupBy
		mbOrderBy1Asc = pbOrder1Asc
		mbOrderBy2Asc = pbOrder2Asc
		mlngOrderByColumnID = plngOrderByColumnID
		mlngGroupByColumnID = plngGroupByColumnID

		Return True

	End Function

	Public Function SetBradfordIncludeOptions(ByVal pbOmitBeforeStart As Boolean, ByVal pbOmitAfterEnd As Boolean, ByVal plngPersonnelID As Long _
																						, ByVal plngCustomReportsFilterID As Long, ByVal plngCustomReportsPickListID As Long, ByVal pbMinBradford As Boolean _
																						, ByVal plngMinBradfordAmount As Long) As Boolean

		' Include options for this report
		mbOmitBeforeStart = pbOmitBeforeStart
		mbOmitAfterEnd = pbOmitAfterEnd
		mlngPersonnelID = plngPersonnelID

		mlngCustomReportsFilterID = IIf(IsNumeric(plngCustomReportsFilterID), plngCustomReportsFilterID, 0)
		mlngCustomReportsPickListID = IIf(IsNumeric(plngCustomReportsPickListID), plngCustomReportsPickListID, 0)
		mbMinBradford = pbMinBradford
		mlngMinBradfordAmount = plngMinBradfordAmount

		Return True

	End Function

	Private Function ConvertSQLDateToLocale(ByRef psSQLDate As String) As Date
		' Convert the given date string (mm/dd/yyyy) into the locale format.
		' NB. This function assumes a sensible locale format is used.
		Dim fDaysDone As Boolean
		Dim fMonthsDone As Boolean
		Dim fYearsDone As Boolean
		Dim iLoop As Short
		Dim sFormattedDate As String

		sFormattedDate = ""

		' Get the locale's date format.
		fDaysDone = False
		fMonthsDone = False
		fYearsDone = False

		For iLoop = 1 To Len(mstrClientDateFormat)
			Select Case UCase(Mid(mstrClientDateFormat, iLoop, 1))
				Case "D"
					If Not fDaysDone Then
						sFormattedDate = sFormattedDate & Mid(psSQLDate, 4, 2)
						fDaysDone = True
					End If

				Case "M"
					If Not fMonthsDone Then
						sFormattedDate = sFormattedDate & Mid(psSQLDate, 1, 2)
						fMonthsDone = True
					End If

				Case "Y"
					If Not fYearsDone Then
						sFormattedDate = sFormattedDate & Mid(psSQLDate, 7, 4)
						fYearsDone = True
					End If

				Case Else
					sFormattedDate = sFormattedDate & Mid(mstrClientDateFormat, iLoop, 1)
			End Select
		Next iLoop

		ConvertSQLDateToLocale = sFormattedDate

	End Function

	' Function which we use to pass in the default output parameters (Standard reports read from the defintion table,
	'    which don't exist for standard reports)
	Public Function SetBradfordDefaultOutputOptions(ByRef pbOutputPreview As Boolean, ByRef plngOutputFormat As Long, ByRef pblnOutputScreen As Boolean, ByRef pblnOutputPrinter As Boolean _
																									, ByRef pstrOutputPrinterName As String, ByRef pblnOutputSave As Boolean, ByRef plngOutputSaveExisting As Long, ByRef pblnOutputEmail As Boolean _
																									, ByRef plngOutputEmailID As Long, ByRef pstrOutputEmailName As String, ByRef pstrOutputEmailSubject As String, ByRef pstrOutputEmailAttachAs As String _
																									, ByRef pstrOutputFilename As String) As Boolean

		mblnOutputPreview = pbOutputPreview
		mlngOutputFormat = plngOutputFormat
		mblnOutputScreen = pblnOutputScreen
		mblnOutputPrinter = pblnOutputPrinter
		mstrOutputPrinterName = pstrOutputPrinterName
		mblnOutputSave = pblnOutputSave
		mlngOutputSaveExisting = plngOutputSaveExisting
		mblnOutputEmail = pblnOutputEmail
		mlngOutputEmailID = plngOutputEmailID
		mstrOutputEmailName = GetEmailGroupName(CInt(plngOutputEmailID))
		mstrOutputEmailSubject = pstrOutputEmailSubject
		mstrOutputEmailAttachAs = IIf(IsDBNull(pstrOutputEmailAttachAs), vbNullString, pstrOutputEmailAttachAs)
		mstrOutputFilename = pstrOutputFilename
		mblnOutputPreview = (pbOutputPreview Or (mlngOutputFormat = OutputFormats.fmtDataOnly And mblnOutputScreen))

		Return True

	End Function

	Public Function UDFFunctions(ByRef pbCreate As Boolean) As Boolean
		Return mclsGeneral.UDFFunctions(mastrUDFsRequired, pbCreate)
	End Function


End Class