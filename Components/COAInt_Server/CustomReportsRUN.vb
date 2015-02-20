Option Strict Off
Option Explicit On

Imports System.Collections.Generic
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.Structures
Imports System.Collections.ObjectModel
Imports HR.Intranet.Server.Metadata
Imports HR.Intranet.Server.BaseClasses
Imports System.Data.SqlClient
Imports System.Linq
Imports HR.Intranet.Server.Expressions

Public Class Report
	Inherits BaseReport

	' To hold Properties
	Private mlngCustomReportID As Integer
	Private mstrErrorString As String

	' Variables to store definition
	'	Private mstrCustomReportsName As String
	Private mlngCustomReportsBaseTable As Integer
	Private mstrCustomReportsBaseTableName As String
	Private mlngCustomReportsPickListID As Integer
	Private mlngCustomReportsFilterID As Integer
	Private mlngCustomReportsParent1Table As Integer
	Private mlngCustomReportsParent1FilterID As Integer
	Private mlngCustomReportsParent2Table As Integer
	Private mlngCustomReportsParent2FilterID As Integer
	Public mblnCustomReportsSummaryReport As Boolean
	Private mblnIgnoreZerosInAggregates As Boolean
	Private mblnCustomReportsPrintFilterHeader As Boolean

	'New Default Output Variables

	'	Private mblnOutputScreen As Boolean
	Private mblnOutputPrinter As Boolean
	Private mstrOutputPrinterName As String
	Private mblnOutputSave As Boolean
	Private mlngOutputSaveExisting As Integer
	Private mblnOutputEmail As Boolean
	Private mlngOutputEmailID As Integer
	Private mstrOutputEmailName As String
	Private mstrOutputEmailSubject As String
	Private mstrOutputEmailAttachAs As String

	Private mvarChildTables(,) As Object
	Private miChildTablesCount As Integer
	Private miUsedChildCount As Integer

	Private mlngCustomReportsParent1PickListID As Integer
	Private mlngCustomReportsParent2PickListID As Integer

	' Recordsets to store the definition and column information
	Public mrstCustomReportsDetails As DataTable

	' TableViewsGuff
	Private mstrRealSource As String
	Private mstrBaseTableRealSource As String
	Private mlngTableViews(,) As Integer
	Private mstrViews() As String
	Private mobjTableView As TablePrivilege
	Private mobjColumnPrivileges As CColumnPrivileges

	' Strings to hold the SQL statement
	Private mstrSQLSelect As String
	Private mstrSQLFrom As String
	Private mstrSQLJoin As String
	Private mstrSQLWhere As String
	Private mstrSQLOrderBy As String
	Private mstrSQL As String

	' Collection to hold column definitions in the report
	Private ColumnDetails As List(Of ReportDetailItem)
	Private colSortOrder As List(Of ReportSortItem)
	Public DisplayColumns As List(Of ReportDetailItem)

	Dim mvarVisibleColumns(,) As Object

	'Array used to store the 'GroupWithNextColumn' option strings.
	Private mvarGroupWith(,) As Object

	'Array used to store the 'POC' values when outputting.
	Private mvarPageBreak As ArrayList
	Private mblnPageBreak As Boolean

	' String to hold the temp table name
	Private mstrTempTableName As String

	' Recordset to store the final data from the temp table
	Public mrstCustomReportsOutput As DataTable

	'Does the report generate no records ?
	Private mblnNoRecords As Boolean

	' Is this a Bradford Index Report
	Private mbIsBradfordIndexReport As Boolean

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

	Private mdBradfordStartDate As DateTime?
	Private mdBradfordEndDate As DateTime?

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

	Public ReportDataTable As DataTable

	Private Sub datCustomReportOutput_Start()

		ReportDataTable = New DataTable()

		Dim bGroupWithNext As Boolean = False

		ReportDataTable.Columns.Add("rowType", GetType(RowType))
		DisplayColumns.Add(New ReportDetailItem With {.IsHidden = True, .ColumnName = "rowType"})

		If mblnReportHasSummaryInfo Then
			DisplayColumns.Add(New ReportDetailItem With {.IsHidden = True, .ColumnName = "Summary Info"})
			ReportDataTable.Columns.Add("Summary Info", GetType(String))
		End If

		For Each objItem In ColumnDetails
			If Not (objItem.IsHidden Or bGroupWithNext) Then
				DisplayColumns.Add(objItem)
				ReportDataTable.Columns.Add(objItem.IDColumnName, GetType(String))
			End If
			bGroupWithNext = objItem.GroupWithNextColumn

		Next

	End Sub

	Public ReadOnly Property IsBradfordReport() As Boolean
		Get
			Return mbIsBradfordIndexReport
		End Get
	End Property

	Public ReadOnly Property HasSummaryColumns() As Boolean
		Get
			Return mblnReportHasSummaryInfo And Not mbIsBradfordIndexReport
		End Get
	End Property

	Public WriteOnly Property SingleRecordID() As Integer
		Set(ByVal Value As Integer)
			mlngSingleRecordID = Value
		End Set
	End Property

	Public ReadOnly Property BaseTableName() As String
		Get
			Return mstrCustomReportsBaseTableName
		End Get
	End Property

	Public ReadOnly Property ChildCount() As Integer
		Get
			Return miChildTablesCount
		End Get
	End Property

	Public ReadOnly Property UsedChildCount() As Integer
		Get
			Return miUsedChildCount
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
				Logs.ChangeHeaderStatus(EventLog_Status.elsFailed)
			End If

		End Set
	End Property

	Public WriteOnly Property FailedMessage() As String
		Set(ByVal Value As String)
			Logs.AddDetailEntry(Value)
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

	Public ReadOnly Property NoRecords() As Boolean
		Get
			Return mblnNoRecords
		End Get
	End Property

	'Public ReadOnly Property SQLSTRING() As String
	'	Get

	'		' Does the report have any records ?
	'		SQLSTRING = mstrSQL

	'	End Get
	'End Property

	Public ReadOnly Property ReportHasPageBreak() As Boolean
		Get
			Return mblnReportHasPageBreak
		End Get
	End Property

	Public ReadOnly Property ReportHasSummaryInfo() As Boolean
		Get
			Return mblnReportHasSummaryInfo And Not mbIsBradfordIndexReport
		End Get
	End Property

	Public ReadOnly Property CustomReportsSummaryReport() As Boolean
		Get
			Return mblnCustomReportsSummaryReport
		End Get
	End Property

	Public ReadOnly Property OutputArray_VisibleColumns() As Object
		Get
			Return VB6.CopyArray(mvarVisibleColumns)
		End Get
	End Property

	Public ReadOnly Property OutputArray_PageBreakValues() As String()
		Get
			Return CType(mvarPageBreak.ToArray(GetType(String)), String())
		End Get
	End Property

	Public WriteOnly Property CustomReportID() As Integer
		Set(ByVal Value As Integer)

			' ID of the report to run passed in from the asp page
			mlngCustomReportID = Value
			mlngSingleRecordID = 0

		End Set
	End Property

	Public ReadOnly Property ErrorString() As String
		Get
			Return mstrErrorString
		End Get
	End Property

	Public ReadOnly Property OutputPrinter() As Boolean
		Get
			Return mblnOutputPrinter
		End Get
	End Property

	Public ReadOnly Property OutputPrinterName() As String
		Get
			Return mstrOutputPrinterName
		End Get
	End Property

	Public Property OutputSave() As Boolean
		Get
			Return mblnOutputSave
		End Get
		Set(value As Boolean)
			mblnOutputSave = value
		End Set
	End Property

	Public ReadOnly Property OutputSaveExisting() As Integer
		Get
			Return mlngOutputSaveExisting
		End Get
	End Property

	Public ReadOnly Property OutputEmail() As Boolean
		Get
			Return mblnOutputEmail
		End Get
	End Property

	Public ReadOnly Property OutputEmailID() As Integer
		Get
			Return mlngOutputEmailID
		End Get
	End Property

	Public ReadOnly Property OutputEmailGroupName() As String
		Get
			Return mstrOutputEmailName
		End Get
	End Property

	Public ReadOnly Property OutputEmailSubject() As String
		Get
			Return mstrOutputEmailSubject
		End Get
	End Property

	Public ReadOnly Property OutputEmailAttachAs() As String
		Get
			Return mstrOutputEmailAttachAs
		End Get
	End Property

	Public Property EventLogID() As Integer
		Get
			EventLogID = Logs.EventLogID
		End Get
		Set(ByVal Value As Integer)
			Logs.EventLogID = Value
		End Set
	End Property

	Public Function SetPromptedValues(ByRef pavPromptedValues As Object) As Boolean

		' Purpose : This function calls the individual functions that
		'           generate the components of the main SQL string.

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

		Catch ex As Exception
			mstrErrorString = "Error setting prompted values." & vbNewLine & ex.Message.RemoveSensitive()
			Logs.AddDetailEntry(mstrErrorString)
			Logs.ChangeHeaderStatus(EventLog_Status.elsFailed)
			Return False

		End Try

		Return True

	End Function

	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()

		ReDim mlngTableViews(2, 0)
		ReDim mstrViews(0)
		mvarOutputArray_Data = New ArrayList()
		ReDim mvarVisibleColumns(3, 0)

		ReDim mvarGroupWith(1, 0)
		mvarPageBreak = New ArrayList()

		' By default this is not a Bradford Index Report
		mbIsBradfordIndexReport = False
		ColumnDetails = New List(Of ReportDetailItem)()
		DisplayColumns = New List(Of ReportDetailItem)

	End Sub

	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub

	Public Function AddTempTableToSQL() As Boolean

		Try

			mstrTempTableName = General.UniqueSQLObjectName("ASRSysTempCustomReport", 3)
			mstrSQLSelect = mstrSQLSelect & " INTO [" & mstrTempTableName & "]"

		Catch ex As Exception
			mstrErrorString = "Error retrieving unique temp table name." & vbNewLine & ex.Message.RemoveSensitive()
			Logs.AddDetailEntry(mstrErrorString)
			Logs.ChangeHeaderStatus(EventLog_Status.elsFailed)
			Return False

		End Try

		Return True

	End Function

	Public Function MergeSQLStrings() As Boolean

		mstrSQL = mstrSQLSelect & " FROM " & mstrSQLFrom & IIf(Len(mstrSQLJoin) = 0, "", " " & mstrSQLJoin) & IIf(Len(mstrSQLWhere) = 0, "", " " & mstrSQLWhere) & " " & mstrSQLOrderBy
		Return True

	End Function

	Public Function ExecuteSql() As Boolean

		Try
			DB.ExecuteSql(mstrSQL)

		Catch ex As Exception
			mblnNoRecords = True
			mstrErrorString = "Error executing SQL statement." & vbNewLine & ex.Message
			Logs.AddDetailEntry(mstrErrorString)
			Logs.ChangeHeaderStatus(EventLog_Status.elsFailed)
			Return False

		End Try

		Return True

	End Function

	Public Function GetCustomReportDefinition() As Boolean

		Dim objData As DataSet
		Dim rsDefinition As DataTable
		Dim prmID As New SqlParameter("ReportID", SqlDbType.Int)

		Dim i As Integer

		Try

			mbIsBradfordIndexReport = False

			prmID.Value = mlngCustomReportID
			objData = DB.GetDataSet("spASRIntGetCustomReport", prmID)
			rsDefinition = objData.Tables(0)

			With rsDefinition

				' Dont run if its been deleted by another user.
				If .Rows.Count = 0 Then
					mstrErrorString = "Report has been deleted by another user."
					Return False
				End If

				Dim rowData = rsDefinition.Rows(0)

				' RH 29/05/01 - Dont run if its been made hidden by another user.
				If LCase(rowData("Username").ToString()) <> LCase(_login.Username) And CurrentUserAccess(UtilityType.utlCustomReport, mlngCustomReportID) = ACCESS_HIDDEN Then
					mstrErrorString = "Report has been made hidden by another user."
					Return False
				End If

				Name = rowData("Name").ToString()
				mlngCustomReportsBaseTable = CInt(rowData("BaseTable"))
				mstrCustomReportsBaseTableName = rowData("TableName").ToString()
				mlngCustomReportsPickListID = CInt(rowData("picklist"))
				mlngCustomReportsFilterID = CInt(rowData("Filter"))
				mlngCustomReportsParent1Table = CInt(rowData("parent1table"))
				mlngCustomReportsParent1FilterID = CInt(rowData("parent1filter"))
				mlngCustomReportsParent2Table = CInt(rowData("parent2table"))
				mlngCustomReportsParent2FilterID = CInt(rowData("parent2filter"))

				mblnCustomReportsSummaryReport = CBool(rowData("Summary"))
				mblnIgnoreZerosInAggregates = CBool(rowData("IgnoreZeros"))
				mblnCustomReportsPrintFilterHeader = CBool(rowData("PrintFilterHeader"))
				mlngCustomReportsParent1PickListID = CInt(rowData("parent1Picklist"))
				mlngCustomReportsParent2PickListID = CInt(rowData("parent2Picklist"))

				'New Default Output Variables
				OutputFormat = CInt(rowData("OutputFormat"))
				OutputScreen = CBool(rowData("OutputScreen"))
				mblnOutputPrinter = CBool(rowData("OutputPrinter"))
				mstrOutputPrinterName = rowData("OutputPrinterName").ToString()
				mblnOutputSave = CBool(rowData("OutputSave"))
				mlngOutputSaveExisting = CInt(rowData("OutputSaveExisting"))
				mblnOutputEmail = CBool(rowData("OutputEmail"))
				mlngOutputEmailID = CInt(rowData("OutputEmailAddr"))
				mstrOutputEmailName = rowData("EmailGroupName").ToString()
				mstrOutputEmailSubject = rowData("OutputEmailSubject").ToString()
				mstrOutputEmailAttachAs = rowData("OutputEmailAttachAs").ToString()
				OutputFilename = rowData("OutputFilename").ToString()
				OutputPreview = CBool(rowData("OutputPreview"))

			End With

			' Child data recordset
			rsDefinition = objData.Tables(1)

			i = 0
			For Each rowData As DataRow In rsDefinition.Rows
				ReDim Preserve mvarChildTables(5, i)
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(0, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarChildTables(0, i) = rowData("ChildTable")	'Childs Table ID
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(1, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarChildTables(1, i) = rowData("childFilter") 'Childs Filter ID (if any)
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(2, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarChildTables(2, i) = rowData("ChildMaxRecords") 'Number of records to take from child
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(3, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarChildTables(3, i) = rowData("TableName") 'Child Table Name
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(4, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarChildTables(4, i) = False	'Boolean - True if table is used, False if not
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(5, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvarChildTables(5, i) = rowData("ChildOrder")	'Childs Order ID (if any)
				i = i + 1
			Next

			miChildTablesCount = i

			If Not IsRecordSelectionValid() Then
				Return False
			End If


			Logs.AddHeader(EventLog_Type.eltCustomReport, Name)

		Catch ex As Exception
			mstrErrorString = "Error reading the Custom Report definition !" & vbNewLine & ex.Message.RemoveSensitive()
			Logs.AddDetailEntry(mstrErrorString)
			Logs.ChangeHeaderStatus(EventLog_Status.elsFailed)
			Return False

		End Try

		Return True

	End Function

	Public Function GetDetailsRecordsets() As Boolean

		' Purpose : This function loads report details and sort details into
		'           arrays and leaves the details recordset reference there
		'           (dont remove it...used for summary info !)
		Try

			Dim strTempSQL As String
			Dim prstCustomReportsSortOrder As DataTable
			Dim lngTableID As Integer
			Dim objReportItemDetail As ReportDetailItem
			Dim objSortItem As ReportSortItem

			' Get the column information from the Details table, in order
			strTempSQL = "EXEC spASRIntGetCustomReportDetails " & mlngCustomReportID
			mrstCustomReportsDetails = DB.GetDataTable((strTempSQL))

			Dim objExpr As clsExprExpression
			With mrstCustomReportsDetails
				If .Rows.Count = 0 Then
					mstrErrorString = "No columns found in the specified Custom Report definition." & vbNewLine & "Please remove this definition and create a new one."
					Return False
				End If

				If Not CheckCalcsStillExist() Then
					Return False
				End If


				For Each objRow As DataRow In mrstCustomReportsDetails.Rows
					objReportItemDetail = New ReportDetailItem

					'*************************************************************************
					'Now we need to decide on what the heading needs to be because QA want to
					'be able to have similar headings for hidden columns...I warned them, but
					'NO...they thought that the best move was to spend ages fixing faults in
					'v2 and put HR Pro .NET on the back-burner so that we can release a
					'Limited Edition of HR Pro called HR Pro .NET 2012 Olympic Edition.
					'What twats!!!...Fault 10211.

					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If IIf((IsDBNull(objRow("Hidden")) Or (objRow("Hidden"))), True, False) Then
						objReportItemDetail.IDColumnName = "?ID_HD_" & objRow("Type") & "_" & objRow("ColExprID")
					Else
						objReportItemDetail.IDColumnName = objRow("Heading").ToString().Replace("'", "")
					End If

					'*************************************************************************

					objReportItemDetail.DataType = CType(objRow("DataType"), ColumnDataType)
					objReportItemDetail.Size = CInt(objRow("Size"))
					objReportItemDetail.Decimals = CInt(objRow("dp"))
					objReportItemDetail.IsNumeric = CBool(objRow("IsNumeric"))
					objReportItemDetail.IsAverage = CBool(objRow("Avge"))
					objReportItemDetail.IsCount = CBool(objRow("cnt"))
					objReportItemDetail.IsTotal = CBool(objRow("tot"))
					objReportItemDetail.IsBreakOnChange = CBool(objRow("boc"))
					objReportItemDetail.IsPageOnChange = CBool(objRow("poc"))
					objReportItemDetail.IsValueOnChange = CBool(objRow("voc"))
					objReportItemDetail.SuppressRepeated = CBool(objRow("srv"))
					objReportItemDetail.LastValue = ""
					objReportItemDetail.ID = CInt(objRow("ColExprID"))
					objReportItemDetail.ID = CInt(objRow("ColExprID"))

					objReportItemDetail.Type = objRow("Type").ToString()

					lngTableID = IIf(IsDBNull(objRow("TableID")), 0, objRow("TableID"))
					objReportItemDetail.TableID = lngTableID
					objReportItemDetail.TableName = objRow("TableName").ToString()

					If objRow("Type").ToString() = "C" Then
						objReportItemDetail.ColumnName = objRow("ColumnName").ToString()
						objReportItemDetail.IsDateColumn = CBool(objRow("IsDateColumn"))
						objReportItemDetail.IsBitColumn = CBool(objRow("IsBooleanColumn"))

					Else
						objReportItemDetail.ColumnName = ""

						'MH20010307
						objExpr = NewExpression()

						objExpr.ExpressionID = CInt(objRow("ColExprID"))
						objExpr.ConstructExpression()
						objExpr.ValidateExpression(True)

						'Sets the IsNumeric value for the calculated column. This will be used to display the content right alligned in the report preview
						objReportItemDetail.IsNumeric = (objExpr.ReturnType = ExpressionValueTypes.giEXPRVALUE_NUMERIC)

						lngTableID = objExpr.BaseTableID
						objReportItemDetail.TableID = lngTableID
						objReportItemDetail.TableName = objExpr.BaseTableName
						objReportItemDetail.ColumnName = ""

						'UPGRADE_NOTE: Object objExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
						objExpr = Nothing

						objReportItemDetail.IsDateColumn = IsDateColumn(objRow("Type"), lngTableID, objRow("ColExprID"))
						objReportItemDetail.IsBitColumn = IsBitColumn(objRow("Type"), lngTableID, objRow("ColExprID"))

					End If

					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					objReportItemDetail.IsHidden = IIf((IsDBNull(objRow("Hidden")) Or (objRow("Hidden"))), True, False)
					objReportItemDetail.IsReportChildTable = IsReportChildTable(lngTableID)	'Indicates if column is a report child table.
					objReportItemDetail.Repetition = IIf(objRow("repetition") = 1, True, False)
					objReportItemDetail.Use1000Separator = CBool(objRow("Use1000separator"))

					' Format for this numeric column
					If objReportItemDetail.IsNumeric Then

						If objReportItemDetail.Use1000Separator Then
							objReportItemDetail.Mask = "{0:#,0." & New String("0", objReportItemDetail.Decimals) & "}"
						Else
							objReportItemDetail.Mask = "{0:#0." & New String("0", objReportItemDetail.Decimals) & "}"
						End If
					End If

					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					objReportItemDetail.GroupWithNextColumn = IIf((IsDBNull(objRow("GroupWithNextColumn")) Or (Not objRow("GroupWithNextColumn"))), False, True)

					ColumnDetails.Add(objReportItemDetail)

				Next
			End With

			'******************************************************************************
			' Add the ID columns for the tables so that we can re-select the child records
			' when we create the multiple child temp table.
			' NB. Is called only when there is more than one child in the report.
			'******************************************************************************

			objReportItemDetail = New ReportDetailItem

			objReportItemDetail.Size = 99
			objReportItemDetail.Decimals = 0
			objReportItemDetail.IsNumeric = False
			objReportItemDetail.IsAverage = False
			objReportItemDetail.IsCount = False
			objReportItemDetail.IsTotal = False
			objReportItemDetail.IsBreakOnChange = False
			objReportItemDetail.IsPageOnChange = False
			objReportItemDetail.IsValueOnChange = False
			objReportItemDetail.SuppressRepeated = False
			objReportItemDetail.LastValue = ""
			objReportItemDetail.ID = -1
			objReportItemDetail.Type = "C"
			objReportItemDetail.TableID = mlngCustomReportsBaseTable
			objReportItemDetail.TableName = GetTableName(objReportItemDetail.TableID)
			objReportItemDetail.IDColumnName = "?ID"
			objReportItemDetail.ColumnName = "ID"
			objReportItemDetail.IsDateColumn = False
			objReportItemDetail.IsBitColumn = False
			objReportItemDetail.IsHidden = True
			objReportItemDetail.IsReportChildTable = IsReportChildTable(lngTableID)	'Indicates if column is a report child table.
			objReportItemDetail.Repetition = True
			objReportItemDetail.Mask = "0"
			objReportItemDetail.GroupWithNextColumn = False	'Group With Next Column.
			ColumnDetails.Add(objReportItemDetail)

			Dim iChildCount As Integer
			Dim lngChildTableID As Integer
			If miChildTablesCount > 0 Then
				For iChildCount = 0 To UBound(mvarChildTables, 2) Step 1
					'TM20020409 Fault 3745 - only add the ID columns for tables that are actually used.
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(0, iChildCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					lngChildTableID = mvarChildTables(0, iChildCount)
					If IsChildTableUsed(lngChildTableID) Then

						objReportItemDetail = New ReportDetailItem
						objReportItemDetail.Size = 99
						objReportItemDetail.Decimals = 0
						objReportItemDetail.IsNumeric = False
						objReportItemDetail.IsAverage = False
						objReportItemDetail.IsCount = False
						objReportItemDetail.IsTotal = False
						objReportItemDetail.IsBreakOnChange = False
						objReportItemDetail.IsPageOnChange = False
						objReportItemDetail.IsValueOnChange = False
						objReportItemDetail.SuppressRepeated = False
						objReportItemDetail.LastValue = ""
						objReportItemDetail.ID = -1
						objReportItemDetail.Type = "C"
						objReportItemDetail.TableID = mvarChildTables(0, iChildCount)
						objReportItemDetail.TableName = mvarChildTables(3, iChildCount)
						objReportItemDetail.IDColumnName = "?ID_" & objReportItemDetail.TableID
						objReportItemDetail.ColumnName = "ID_" & mlngCustomReportsBaseTable
						objReportItemDetail.IsDateColumn = False
						objReportItemDetail.IsBitColumn = False
						objReportItemDetail.IsHidden = True
						objReportItemDetail.IsReportChildTable = True	'Indicates if column is a report child table.
						objReportItemDetail.Repetition = True
						objReportItemDetail.Mask = "0"
						objReportItemDetail.GroupWithNextColumn = False	'Group With Next Column.
						ColumnDetails.Add(objReportItemDetail)

						objReportItemDetail = New ReportDetailItem
						objReportItemDetail.Size = 99
						objReportItemDetail.Decimals = 0
						objReportItemDetail.IsNumeric = False
						objReportItemDetail.IsAverage = False
						objReportItemDetail.IsCount = False
						objReportItemDetail.IsTotal = False
						objReportItemDetail.IsBreakOnChange = False
						objReportItemDetail.IsPageOnChange = False
						objReportItemDetail.IsValueOnChange = False
						objReportItemDetail.SuppressRepeated = False
						objReportItemDetail.LastValue = ""
						objReportItemDetail.ID = -1
						objReportItemDetail.Type = "C"
						objReportItemDetail.TableID = mvarChildTables(0, iChildCount)
						objReportItemDetail.TableName = mvarChildTables(3, iChildCount)
						objReportItemDetail.IDColumnName = "?ID_" & objReportItemDetail.TableName
						objReportItemDetail.ColumnName = "ID"
						objReportItemDetail.IsDateColumn = False
						objReportItemDetail.IsBitColumn = False
						objReportItemDetail.IsHidden = True
						objReportItemDetail.IsReportChildTable = True	'Indicates if column is a report child table.
						objReportItemDetail.Repetition = True

						objReportItemDetail.Mask = "0"
						objReportItemDetail.GroupWithNextColumn = False	'Group With Next Column.
						ColumnDetails.Add(objReportItemDetail)

					End If
				Next iChildCount
			End If

			If miChildTablesCount > 1 Then

				objReportItemDetail = New ReportDetailItem()
				objReportItemDetail.Size = 99
				objReportItemDetail.Decimals = 0
				objReportItemDetail.IsNumeric = True
				objReportItemDetail.IsAverage = False
				objReportItemDetail.IsCount = False
				objReportItemDetail.IsTotal = False
				objReportItemDetail.IsBreakOnChange = False
				objReportItemDetail.IsPageOnChange = False
				objReportItemDetail.IsValueOnChange = False
				objReportItemDetail.SuppressRepeated = False
				objReportItemDetail.LastValue = ""
				objReportItemDetail.ID = -1
				objReportItemDetail.Type = "C"
				objReportItemDetail.TableID = -1
				objReportItemDetail.TableName = ""
				objReportItemDetail.IDColumnName = lng_SEQUENCECOLUMNNAME
				objReportItemDetail.ColumnName = ""
				objReportItemDetail.IsDateColumn = False
				objReportItemDetail.IsBitColumn = False
				objReportItemDetail.IsHidden = True
				objReportItemDetail.IsReportChildTable = True	'Indicates if column is a report child table.
				objReportItemDetail.Repetition = True
				objReportItemDetail.Mask = "0"
				objReportItemDetail.GroupWithNextColumn = False	'Group With Next Column.
				ColumnDetails.Add(objReportItemDetail)

			End If

			'******************************************************************************

			' Calculate if we are going to need summary columns
			For Each objItem In ColumnDetails
				If objItem.IsAverage Or objItem.IsTotal Or objItem.IsCount Then
					mblnReportHasSummaryInfo = True
					Exit For
				End If
			Next

			' Get those columns defined as a SortOrder and load into array

			strTempSQL = "SELECT * FROM ASRSysCustomReportsDetails WHERE CustomReportID = " & mlngCustomReportID & " AND SortOrderSequence > 0 ORDER BY [SortOrderSequence]"
			prstCustomReportsSortOrder = DB.GetDataTable(strTempSQL)

			colSortOrder = New List(Of ReportSortItem)

			With prstCustomReportsSortOrder
				If .Rows.Count = 0 Then
					GetDetailsRecordsets = False
					mstrErrorString = "No columns have been defined as a sort order for the specified Custom Report definition." & vbNewLine & "Please remove this definition and create a new one."
					Exit Function
				End If

				For Each objRow As DataRow In .Rows
					objSortItem = New ReportSortItem
					objSortItem.TableID = GetTableIDFromColumn(CInt(objRow("ColExprID")))
					objSortItem.ColExprID = objRow("ColExprID")
					objSortItem.AscDesc = objRow("SortOrder")
					colSortOrder.Add(objSortItem)
				Next
			End With

			'UPGRADE_NOTE: Object prstCustomReportsSortOrder may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			prstCustomReportsSortOrder = Nothing


		Catch ex As Exception
			mstrErrorString = "Error retrieving the details recordsets'." & vbNewLine & ex.Message
			Logs.AddDetailEntry(mstrErrorString)
			Logs.ChangeHeaderStatus(EventLog_Status.elsFailed)
			Return False

		End Try

		Return True

	End Function

	Private Function IsChildTableUsed(iChildTableID As Integer) As Boolean
		Return ColumnDetails.Any(Function(objItem) objItem.TableID = iChildTableID)
	End Function

	Public Function GenerateSQL() As Boolean

		' Purpose : This function calls the individual functions that
		'           generate the components of the main SQL string.

		Dim fOK As Boolean

		fOK = True

		If fOK Then fOK = GenerateSQLSelect()
		GenerateSQLFrom()
		If fOK Then fOK = GenerateSQLJoin()
		If fOK Then fOK = GenerateSQLWhere()
		If fOK Then fOK = GenerateSQLOrderBy()

		If fOK Then
			Return True
		Else
			mblnNoRecords = True
			Return False
		End If

	End Function

	Private Function GenerateSQLSelect() As Boolean

		Dim plngTempTableID As Integer
		Dim pstrTempTableName As String
		Dim pstrTempColumnName As String

		Dim pblnOK As Boolean
		Dim pblnColumnOK As Boolean
		Dim iLoop1 As Integer
		Dim pblnNoSelect As Boolean
		Dim pblnFound As Boolean

		Dim pintLoop As Integer
		Dim pstrColumnList As String
		Dim pstrColumnCode As String
		Dim pstrSource As String
		Dim pintNextIndex As Integer

		Dim blnOK As Boolean
		Dim sCalcCode As String
		Dim alngSourceTables(,) As Integer
		Dim objCalcExpr As clsExprExpression
		Dim objTableView As TablePrivilege

		Try

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
			For Each objReportItem In ColumnDetails

				' Clear temp vars
				plngTempTableID = 0
				pstrTempTableName = vbNullString
				pstrTempColumnName = vbNullString

				' If its a COLUMN then...
				If objReportItem.Type = "C" Then
					If objReportItem.IDColumnName <> lng_SEQUENCECOLUMNNAME Then
						' Load the temp variables
						plngTempTableID = objReportItem.TableID
						pstrTempTableName = objReportItem.TableName
						pstrTempColumnName = objReportItem.ColumnName

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
								pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ",", "") & "convert(numeric(10,2)," & mstrRealSource & "." & Trim(pstrTempColumnName) & ")" & " AS [" & objReportItem.IDColumnName & "]"
							Else
								pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ",", "") & mstrRealSource & "." & Trim(pstrTempColumnName) & " AS [" & objReportItem.IDColumnName & "]"
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
									pstrColumnCode = pstrColumnCode & " ELSE NULL" & " END AS [" & objReportItem.IDColumnName & "]"
									pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ",", "") & pstrColumnCode
								End If

							End If

							' If we cant see a column, then get outta here
							If pblnNoSelect Then
								GenerateSQLSelect = False
								mstrErrorString = vbNewLine & vbNewLine & "You do not have permission to see the column '" & objReportItem.ColumnName & "'" & vbNewLine & "either directly or through any views."
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
						pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ",", "") & 0 & " AS [" & objReportItem.IDColumnName & "] "

					End If

				Else

					' UH OH ! Its an expression rather than a column

					' Get the calculation SQL, and the array of tables/views that are used to create it.
					' Column 1 = 0 if this row is for a table, 1 if it is for a view.
					' Column 2 = table/view ID.
					ReDim alngSourceTables(2, 0)
					objCalcExpr = NewExpression()
					blnOK = objCalcExpr.Initialise(mlngCustomReportsBaseTable, objReportItem.ID, ExpressionTypes.giEXPR_RUNTIMECALCULATION, ExpressionValueTypes.giEXPRVALUE_UNDEFINED)
					If blnOK Then
						blnOK = objCalcExpr.RuntimeCalculationCode(alngSourceTables, sCalcCode, mastrUDFsRequired, True, False, mvarPrompts)


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
						pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ",", "") & sCalcCode & " AS [" & objReportItem.IDColumnName & "]"

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
						mstrErrorString = "You do not have permission to use the '" & objReportItem.IDColumnName & "' calculation."
						Return False
					End If

				End If

			Next

			mstrSQLSelect = mstrSQLSelect & pstrColumnList

		Catch ex As Exception
			mstrErrorString = "Error generating SQL Select statement." & vbNewLine & ex.Message.RemoveSensitive()
			Logs.AddDetailEntry(mstrErrorString)
			Logs.ChangeHeaderStatus(EventLog_Status.elsFailed)
			Return False

		End Try

		Return True

	End Function

	Private Function IsReportChildTable(lngTableID As Integer) As Boolean

		Dim i As Integer

		IsReportChildTable = False

		If miChildTablesCount > 0 Then
			For i = 0 To UBound(mvarChildTables, 2) Step 1
				'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(0, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If lngTableID = mvarChildTables(0, i) Then
					Return True
				End If
			Next i
		End If

	End Function

	Private Function GetMostChildsForParent(avChildRecs(,) As DataTable, iParentCount As Short) As Integer

		Dim i As Integer
		Dim iMostChildRecords As Integer
		Dim iChildRecordCount As Integer

		Try

			iMostChildRecords = 0
			iChildRecordCount = 0

			For i = 0 To UBound(avChildRecs, 2) Step 1
				If (avChildRecs(iParentCount, i).Rows.Count = 0) Then
					iChildRecordCount = 0
				Else
					iChildRecordCount = avChildRecs(iParentCount, i).Rows.Count
				End If
				If iChildRecordCount > iMostChildRecords Then
					iMostChildRecords = iChildRecordCount
				End If
			Next i

		Catch ex As Exception
			Return 0

		End Try

		Return iMostChildRecords

	End Function

	Private Function OrderBy(plngTableID As Integer) As String

		' This function creates an ORDER BY statement by searching
		' through the columns defined as the reports sort order, then
		' uses the relevant alias name

		Dim bHasOrder As Boolean

		OrderBy = ""
		bHasOrder = False

		For Each objSort In colSortOrder
			For Each objItem In ColumnDetails
				If objSort.ColExprID = objItem.ID And objSort.TableID = plngTableID Then
					OrderBy = OrderBy & "[" & objItem.IDColumnName & "] " & objSort.AscDesc & ", "
					bHasOrder = True
					Exit For
				End If
			Next
		Next

		If bHasOrder Then
			OrderBy = " ORDER BY " & Left(OrderBy, Len(OrderBy) - 2) & " "
		Else
			OrderBy = vbNullString
		End If

	End Function

	Public Function CreateMutipleChildTempTable() As Boolean

		Dim sMCTempTable As String
		Dim sSQL As String
		Dim sParentSelectSQL As String
		Dim rsParent As DataTable
		Dim lngTableID As Integer
		Dim iChildCount As Integer
		Dim iParentCount As Integer
		Dim avChildRecordsets(,) As DataTable
		Dim sChildSelectSQL As String
		Dim sChildWhereSQL As String
		Dim iFields As Integer
		Dim iChildRowCount As Integer
		Dim iChildUsed As Integer
		Dim iMostChilds As Integer
		Dim lngCurrentTableID As Integer
		Dim lngSequenceCount As Integer

		Dim sFIELDS As String
		Dim sVALUES As String

		Dim aryInsertStatements As New List(Of String)

		Try

			'******************* Create multiple child temp table ***************************
			sMCTempTable = General.UniqueSQLObjectName("ASRSysTempCustomReport", 3)

			sSQL = "SELECT * INTO [" & sMCTempTable & "] FROM [" & mstrTempTableName & "]"
			DB.ExecuteSql(sSQL)

			sSQL = "DELETE FROM [" & sMCTempTable & "]"
			DB.ExecuteSql(sSQL)


			'************** Get the Parent SELECT SQL statment ******************************
			For Each objItem In ColumnDetails

				lngTableID = objItem.TableID
				If IsReportParentTable(lngTableID) Or IsReportBaseTable(lngTableID) Then
					sParentSelectSQL = sParentSelectSQL & "[" & objItem.IDColumnName & "], "
				End If
			Next

			sParentSelectSQL = Left(sParentSelectSQL, Len(sParentSelectSQL) - 2) & " "

			sSQL = "SELECT DISTINCT " & sParentSelectSQL & " FROM [" & mstrTempTableName & "] "

			'Order the Parent recorset
			sSQL = sSQL & OrderBy(mlngCustomReportsBaseTable)

			rsParent = DB.GetDataTable(sSQL)

			lngTableID = 0
			iChildUsed = 0

			'*************** Circle through the distinct list of parent records *************
			With rsParent

				'TM20020802 Fault 4273
				If (.Rows.Count = 0) Then
					mstrErrorString = "No records meet the selection criteria."
					CreateMutipleChildTempTable = False
					Logs.AddDetailEntry("Completed successfully. " & mstrErrorString)
					Logs.ChangeHeaderStatus(EventLog_Status.elsSuccessful)
					mblnNoRecords = True

					sMCTempTable = vbNullString
					'      Set rsTemp = Nothing
					'UPGRADE_NOTE: Object rsParent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					rsParent = Nothing
					Exit Function
				End If

				iParentCount = 0
				lngSequenceCount = 1

				mbUseSequence = True

				For Each objRow As DataRow In .Rows

					iParentCount = iParentCount + 1

					ReDim avChildRecordsets(0, miUsedChildCount - 1)
					For iChildCount = 0 To UBound(mvarChildTables, 2) Step 1
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(0, iChildCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						lngCurrentTableID = mvarChildTables(0, iChildCount)
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(4, iChildCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If mvarChildTables(4, iChildCount) Then	'is the child table used???

							For Each objItem In ColumnDetails
								lngTableID = objItem.TableID
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(0, iChildCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If objItem.IsReportChildTable And (lngTableID = mvarChildTables(0, iChildCount)) And (objItem.ColumnName <> ("?ID_" & CStr(mlngCustomReportsBaseTable))) Then
									'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									sChildSelectSQL = sChildSelectSQL & "[" & objItem.IDColumnName & "]" & ", "
								End If
							Next
							sChildSelectSQL = Left(sChildSelectSQL, Len(sChildSelectSQL) - 2) & " "

							'UPGRADE_WARNING: Couldn't resolve default property of object mvarChildTables(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sChildWhereSQL = sChildWhereSQL & "[?ID_" & mvarChildTables(0, iChildCount) & "] = "
							sChildWhereSQL = sChildWhereSQL & objRow("?ID")

							sSQL = "SELECT DISTINCT " & sChildSelectSQL & " FROM [" & mstrTempTableName & "] WHERE " & sChildWhereSQL

							'Order the child recordset.
							sSQL = sSQL & OrderBy(lngCurrentTableID)

							sChildSelectSQL = vbNullString
							sChildWhereSQL = vbNullString

							'Add the child tables recordset to the array of child tables.
							avChildRecordsets(0, iChildUsed) = DB.GetDataTable(sSQL)
							iChildUsed += 1
						End If
					Next iChildCount


					iMostChilds = GetMostChildsForParent(avChildRecordsets, 0)
					If iMostChilds > 0 Then
						For iChildRowCount = 0 To iMostChilds - 1 Step 1

							sFIELDS = vbNullString
							sVALUES = vbNullString

							'<<<<<<<<<<<<<<<<<<< Add Values To Parent Fields >>>>>>>>>>>>>>>>>>>>>>>
							For iFields = 0 To rsParent.Columns.Count - 1 Step 1

								sFIELDS = sFIELDS & "[" & rsParent.Columns(iFields).ColumnName & "],"

								Select Case rsParent.Columns(iFields).DataType.Name.ToLower()
									Case "int32"
										'	'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
										sVALUES = sVALUES & IIf(IsDBNull(objRow(iFields)), 0, objRow(iFields)) & ","

									Case "datetime"
										If Not IsDBNull(objRow(iFields)) Then
											sVALUES = sVALUES & "'" & VB6.Format(objRow(iFields), "MM/dd/yyyy") & "',"
										Else
											sVALUES = sVALUES & "NULL,"
										End If

									Case "boolean"
										sVALUES = sVALUES & IIf(CBool(objRow(iFields)), 1, 0) & ","

									Case Else
										'MH20021119 Fault 4315
										'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
										If Not IsDBNull(objRow(iFields)) Then
											sVALUES = sVALUES & "'" & Replace(objRow(iFields), "'", "''") & "',"
										Else
											sVALUES = sVALUES & "'',"
										End If
								End Select

							Next iFields

							For iChildCount = 0 To UBound(avChildRecordsets, 2) Step 1

								If avChildRecordsets(0, iChildCount).Rows.Count > iChildRowCount Then

									Dim rowFirstRow = avChildRecordsets(0, iChildCount).Rows(iChildRowCount)

									'<<<<<<<<<<<<<<<<<<< Add Values To Child Fields >>>>>>>>>>>>>>>>>>>>>>>
									For iFields = 0 To avChildRecordsets(0, iChildCount).Columns.Count - 1 Step 1
										sFIELDS = sFIELDS & "[" & avChildRecordsets(0, iChildCount).Columns(iFields).ColumnName & "],"

										Select Case avChildRecordsets(0, iChildCount).Columns(iFields).DataType.Name.ToLower()
											Case "int32"
												'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
												sVALUES = sVALUES & IIf(IsDBNull(rowFirstRow(iFields)), 0, rowFirstRow(iFields)) & ","
											Case "datetime"
												'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
												If Not IsDBNull(rowFirstRow(iFields)) Then
													sVALUES = sVALUES & "'" & VB6.Format(rowFirstRow(iFields), "MM/dd/yyyy") & "',"
												Else
													sVALUES = sVALUES & "NULL,"
												End If
											Case "boolean"
												sVALUES = sVALUES & IIf(rowFirstRow(iFields), 1, 0) & ","
											Case Else
												'MH20021119 Fault 4315
												'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
												If Not IsDBNull(rowFirstRow(iFields)) Then
													sVALUES = sVALUES & "'" & Replace(rowFirstRow(iFields), "'", "''") & "',"
												Else
													sVALUES = sVALUES & "'',"
												End If
										End Select

									Next iFields

								End If
							Next iChildCount

							sFIELDS = sFIELDS & "[" & lng_SEQUENCECOLUMNNAME & "]"
							sVALUES = sVALUES & lngSequenceCount

							lngSequenceCount += 1
							aryInsertStatements.Add(String.Format("INSERT INTO {0} ({1}) VALUES ({2});{3}", sMCTempTable, sFIELDS, sVALUES, vbNewLine))

						Next iChildRowCount
					Else

						sFIELDS = vbNullString
						sVALUES = vbNullString

						'<<<<<<<<<<<<<<<<<<< Add Values To Parent Fields >>>>>>>>>>>>>>>>>>>>>>>
						For iFields = 0 To rsParent.Columns.Count - 1 Step 1

							sFIELDS = sFIELDS & "[" & rsParent.Columns(iFields).ColumnName & "],"

							Select Case rsParent.Columns(iFields).DataType.Name.ToLower()

								Case "int32"
									'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
									sVALUES = sVALUES & IIf(IsDBNull(objRow(iFields)), 0, objRow(iFields)) & ","
								Case "datetime"
									'TM20030124 Fault 4974
									'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
									If Not IsDBNull(objRow(iFields)) Then
										sVALUES = sVALUES & "'" & VB6.Format(objRow(iFields), SQLDateFormat) & "',"
									Else
										sVALUES = sVALUES & "NULL,"
									End If
								Case "boolean"
									sVALUES = sVALUES & IIf(objRow(iFields), 1, 0) & ","
								Case Else
									'MH20021119 Fault 4315
									'sVALUES = sVALUES & "'" & Replace(rsParent.Fields(iFields).Value, "'", "''") & "',"
									'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
									If Not IsDBNull(objRow(iFields)) Then
										sVALUES = sVALUES & "'" & Replace(CStr(objRow(iFields)), "'", "''") & "',"
									Else
										sVALUES = sVALUES & "'',"
									End If
							End Select

						Next iFields

						'Add the Sequence number to the sequence column for ordering the data later.

						sFIELDS = sFIELDS & "[" & lng_SEQUENCECOLUMNNAME & "]"
						sVALUES = sVALUES & lngSequenceCount

						lngSequenceCount += 1
						aryInsertStatements.Add(String.Format("INSERT INTO {0} ({1}) VALUES ({2});{3}", sMCTempTable, sFIELDS, sVALUES, vbNewLine))

					End If
					'      End With

					iChildUsed = 0
				Next
			End With


			'************ Re-Order the data using the defined sort orders. ******************
			aryInsertStatements.Add(String.Format("DELETE FROM {0};{1}", mstrTempTableName, vbNewLine))
			aryInsertStatements.Add(String.Format("INSERT INTO [{0}] SELECT * FROM [{1}] ORDER BY [{2}] ASC;", mstrTempTableName, sMCTempTable, lng_SEQUENCECOLUMNNAME))

			sSQL = Join(aryInsertStatements.ToArray())

			DB.ExecuteSql(sSQL)


			'***************** Drop the multiple child temp table. **************************
			General.DropUniqueSQLObject(sMCTempTable, 3)
			sMCTempTable = vbNullString

		Catch ex As Exception

			mstrErrorString = "Error creating temporary table for multiple childs." & vbNewLine & Err.Number & vbNewLine & ex.Message
			Logs.AddDetailEntry(mstrErrorString)
			Logs.ChangeHeaderStatus(EventLog_Status.elsFailed)

			Return False

		End Try

		Return True

	End Function

	Private Function IsReportParentTable(lngTableID As Integer) As Boolean

		If lngTableID = mlngCustomReportsParent1Table Or lngTableID = mlngCustomReportsParent2Table Then
			Return True
		End If

		Return False

	End Function

	Private Function IsReportBaseTable(lngTableID As Integer) As Boolean

		If lngTableID = mlngCustomReportsBaseTable Then
			Return True
		End If

		Return False

	End Function

	Private Sub GenerateSQLFrom()

		mstrSQLFrom = gcoTablePrivileges.Item(mstrCustomReportsBaseTableName).RealSource

	End Sub

	Private Function GenerateSQLJoin() As Boolean

		' Purpose : Add the join strings for parent/child/views.
		'           Also adds filter clauses to the joins if used

		Dim pobjTableView As TablePrivilege
		Dim objChildTable As TablePrivilege
		Dim pintLoop As Integer
		Dim sChildJoinCode As String
		Dim sChildOrderString As String
		Dim rsTemp As DataTable
		Dim strFilterIDs As String
		Dim blnOK As Boolean
		Dim pblnChildUsed As Boolean
		Dim sChildJoin As String
		Dim lngTempChildID As Integer
		Dim lngTempMaxRecords As Integer
		Dim lngTempFilterID As Integer
		Dim lngTempOrderID As Integer
		Dim i As Integer
		Dim sOtherParentJoinCode As String
		Dim iLoop2 As Integer

		Try


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
										If IsAChildOf(mlngTableViews(2, iLoop2), (pobjTableView.TableID)) Then
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
							If lngTempOrderID > 0 Then
								rsTemp = GetOrderDefinition(lngTempOrderID)
							Else
								rsTemp = GetOrderDefinition(GetDefaultOrder(lngTempChildID))
							End If

							sChildOrderString = DoChildOrderString(rsTemp, sChildJoin, lngTempChildID)
							'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
							rsTemp = Nothing

							sChildJoinCode = sChildJoinCode & sChildJoin

							sChildJoinCode = sChildJoinCode & " WHERE (" & objChildTable.RealSource & ".ID_" & mlngCustomReportsBaseTable & " = " & mstrBaseTableRealSource & ".ID)"

							' is the child filtered ?

							If lngTempFilterID > 0 Then
								blnOK = FilteredIDs(lngTempFilterID, strFilterIDs, mastrUDFsRequired, mvarPrompts)

								If blnOK Then
									sChildJoinCode = sChildJoinCode & " AND " & objChildTable.RealSource & ".ID IN (" & strFilterIDs & ")"
								Else
									' Permission denied on something in the filter.
									mstrErrorString = "You do not have permission to use the '" & General.GetFilterName(lngTempFilterID) & "' filter."
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

		Catch ex As Exception
			mstrErrorString = "Error in GenerateSQLJoin." & vbNewLine & Err.Description
			Logs.AddDetailEntry(mstrErrorString)
			Logs.ChangeHeaderStatus(EventLog_Status.elsFailed)
			Return False

		End Try

		Return True

	End Function

	Private Function DoChildOrderString(rsTemp As DataTable, ByRef psJoinCode As String, plngChildID As Integer) As String

		' This function loops through the child tables default order
		' checking if the user has privileges. If they do, add to the order string
		' if not, leave it out.

		Dim sChildOrderString As String
		Dim fColumnOK As Boolean
		Dim fFound As Boolean
		Dim iNextIndex As Integer
		Dim sSource As String
		Dim sRealSource As String
		Dim sColumnCode As String
		Dim sCurrentTableViewName As String
		Dim objColumnPrivileges As CColumnPrivileges
		Dim pobjOrderCol As TablePrivilege
		Dim objTableView As TablePrivilege
		Dim alngTableViews(,) As Integer
		Dim asViews() As String
		Dim iTempCounter As Integer

		Try

			' Dimension an array of tables/views joined to the base table/view.
			' Column 1 = 0 if this row is for a table, 1 if it is for a view.
			' Column 2 = table/view ID.
			ReDim alngTableViews(2, 0)

			pobjOrderCol = gcoTablePrivileges.FindTableID(plngChildID)
			sCurrentTableViewName = pobjOrderCol.RealSource
			'UPGRADE_NOTE: Object pobjOrderCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			pobjOrderCol = Nothing

			For Each objRow As DataRow In rsTemp.Rows

				If objRow("Type") = "O" Then
					' Check if the user can read the column.
					pobjOrderCol = gcoTablePrivileges.FindTableID(objRow("TableID"))
					objColumnPrivileges = GetColumnPrivileges((pobjOrderCol.TableName))
					fColumnOK = objColumnPrivileges.Item(objRow("ColumnName")).AllowSelect
					'UPGRADE_NOTE: Object objColumnPrivileges may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					objColumnPrivileges = Nothing

					If fColumnOK Then
						'        If rsTemp!TableID = mlngCustomReportsChildTable Then
						If objRow("TableID") = plngChildID Then
							sChildOrderString &= IIf(Len(sChildOrderString) > 0, ",", "") & pobjOrderCol.RealSource & "." & objRow("ColumnName") & IIf(objRow("Ascending"), "", " DESC")
						Else
							' If the column comes from a parent table, then add the table to the Join code.
							' Check if the table has already been added to the join code.
							fFound = False
							iTempCounter = 0
							For iNextIndex = 1 To UBound(alngTableViews, 2)
								If alngTableViews(1, iNextIndex) = 0 And alngTableViews(2, iNextIndex) = objRow("TableID") Then
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
								alngTableViews(2, iNextIndex) = objRow("TableID")

								iTempCounter = iNextIndex

								psJoinCode = psJoinCode & " LEFT OUTER JOIN " & pobjOrderCol.RealSource & " ASRSysTemp_" & Trim(Str(iTempCounter)) & " ON " & sCurrentTableViewName & ".ID_" & Trim(Str(objRow("TableID"))) & " = ASRSysTemp_" & Trim(Str(iTempCounter)) & ".ID"
							End If

							sChildOrderString &= IIf(Len(sChildOrderString) > 0, ",", "") & "ASRSysTemp_" & Trim(Str(iTempCounter)) & "." & objRow("ColumnName").ToString() & IIf(objRow("Ascending"), "", " DESC")
						End If
					Else
						' The column cannot be read from the base table/view, or directly from a parent table.
						' If it is a column from a prent table, then try to read it from the views on the parent table.
						'        If rsTemp!TableID <> mlngCustomReportsChildTable Then
						If objRow("TableID") <> plngChildID Then
							' Loop through the views on the column's table, seeing if any have 'read' permission granted on them.
							ReDim asViews(0)
							For Each objTableView In gcoTablePrivileges.Collection
								If (Not objTableView.IsTable) And (objTableView.TableID = objRow("TableID")) And (objTableView.AllowSelect) Then

									sSource = objTableView.ViewName
									sRealSource = gcoTablePrivileges.Item(sSource).RealSource

									' Get the column permission for the view.
									objColumnPrivileges = GetColumnPrivileges(sSource)

									If objColumnPrivileges.IsValid(objRow("ColumnName").ToString()) Then
										If objColumnPrivileges.Item(objRow("ColumnName").ToString()).AllowSelect Then
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

									sColumnCode = sColumnCode & " WHEN NOT ASRSysTemp_" & Trim(Str(iNextIndex)) & "." & objRow("ColumnName").ToString & " IS NULL THEN ASRSysTemp_" & Trim(Str(iNextIndex)) & "." & objRow("ColumnName").ToString()
								Next iNextIndex

								If Len(sColumnCode) > 0 Then
									sColumnCode = sColumnCode & " ELSE NULL" & " END"

									' Add the column to the order string.
									sChildOrderString &= IIf(Len(sChildOrderString) > 0, ", ", "") & sColumnCode & IIf(objRow("Ascending").ToString(), "", " DESC")
								End If
							End If
						End If
					End If

					'UPGRADE_NOTE: Object pobjOrderCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					pobjOrderCol = Nothing
				End If

			Next

			' JIRA 3180 - Force the ID to be part of the sort order because the UDFs sort by ID too
			sChildOrderString &= "," & sCurrentTableViewName & ".ID"

		Catch ex As Exception
			mstrErrorString = "Error while generating child order string" & vbNewLine & ex.Message
			Logs.AddDetailEntry(mstrErrorString)
			Logs.ChangeHeaderStatus(EventLog_Status.elsFailed)
			Return ""

		End Try

		Return sChildOrderString

	End Function

	Private Function GenerateSQLWhere() As Boolean

		' Purpose : Generate the where clauses that cope with the joins
		'           NB Need to add the where clauses for filters/picklists etc

		Dim pintLoop As Integer
		Dim pobjTableView As TablePrivilege
		Dim prstTemp As DataTable
		Dim pstrPickListIDs As String
		Dim blnOK As Boolean
		Dim strFilterIDs As String
		Dim pstrParent1PickListIDs As String
		Dim pstrParent2PickListIDs As String

		Try

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
				prstTemp = DB.GetDataTable("EXEC sp_ASRGetPickListRecords " & mlngCustomReportsParent1PickListID)

				If prstTemp.Rows.Count = 0 Then
					mstrErrorString = "The first parent table picklist contains no records."
					Return False
				End If

				For Each objRow As DataRow In prstTemp.Rows
					pstrParent1PickListIDs = pstrParent1PickListIDs & IIf(Len(pstrParent1PickListIDs) > 0, ", ", "") & objRow(0)
				Next


				mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & mstrBaseTableRealSource & ".ID_" & mlngCustomReportsParent1Table & " IN (" & pstrParent1PickListIDs & ") "
			ElseIf mlngCustomReportsParent1FilterID > 0 Then
				blnOK = FilteredIDs(mlngCustomReportsParent1FilterID, strFilterIDs, mastrUDFsRequired, mvarPrompts)

				If blnOK Then
					mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & mstrBaseTableRealSource & ".ID_" & mlngCustomReportsParent1Table & " IN (" & strFilterIDs & ") "
				Else
					mstrErrorString = "You do not have permission to use the '" & General.GetFilterName(mlngCustomReportsParent1FilterID) & "' filter."
					Return False
				End If
			End If

			' Parent 2 filter and picklist
			If mlngCustomReportsParent2PickListID > 0 Then
				pstrParent2PickListIDs = ""
				prstTemp = DB.GetDataTable("EXEC sp_ASRGetPickListRecords " & mlngCustomReportsParent2PickListID)

				If prstTemp.Rows.Count = 0 Then
					mstrErrorString = "The second parent table picklist contains no records."
					Return False
				End If

				For Each objRow As DataRow In prstTemp.Rows
					pstrParent2PickListIDs = pstrParent2PickListIDs & IIf(Len(pstrParent2PickListIDs) > 0, ", ", "") & objRow(0)
				Next

				mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & mstrBaseTableRealSource & ".ID_" & mlngCustomReportsParent2Table & " IN (" & pstrParent2PickListIDs & ") "
			ElseIf mlngCustomReportsParent2FilterID > 0 Then
				blnOK = FilteredIDs(mlngCustomReportsParent2FilterID, strFilterIDs, mastrUDFsRequired, mvarPrompts)

				If blnOK Then
					mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & mstrBaseTableRealSource & ".ID_" & mlngCustomReportsParent2Table & " IN (" & strFilterIDs & ") "
				Else
					mstrErrorString = "You do not have permission to use the '" & General.GetFilterName(mlngCustomReportsParent2FilterID) & "' filter."
					Return False
				End If
			End If

			If mlngSingleRecordID > 0 Then
				mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & mstrSQLFrom & ".ID IN (" & CStr(mlngSingleRecordID) & ")"

			ElseIf mlngCustomReportsPickListID > 0 Then
				' Now if we are using a picklist, add a where clause for that
				'Get List of IDs from Picklist
				prstTemp = DB.GetDataTable("EXEC sp_ASRGetPickListRecords " & mlngCustomReportsPickListID)

				If prstTemp.Rows.Count = 0 Then
					mstrErrorString = "The selected picklist contains no records."
					Return False
				End If

				For Each objRow As DataRow In prstTemp.Rows
					pstrPickListIDs = pstrPickListIDs & IIf(Len(pstrPickListIDs) > 0, ", ", "") & objRow(0).ToString()
				Next

				mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & mstrSQLFrom & ".ID IN (" & pstrPickListIDs & ")"

				' If we are running a Bradford Report on an individual person
			ElseIf mbIsBradfordIndexReport = True And mlngPersonnelID > 0 Then

				mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & mstrSQLFrom & ".ID IN (" & mlngPersonnelID & ")"

			ElseIf mlngCustomReportsFilterID > 0 Then

				blnOK = FilteredIDs(mlngCustomReportsFilterID, strFilterIDs, mastrUDFsRequired, mvarPrompts)

				If blnOK Then
					mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & mstrSQLFrom & ".ID IN (" & strFilterIDs & ")"
				Else
					' Permission denied on something in the filter.
					mstrErrorString = "You do not have permission to use the '" & General.GetFilterName(mlngCustomReportsFilterID) & "' filter."
					Return False
				End If
			End If

			'UPGRADE_NOTE: Object prstTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			prstTemp = Nothing

		Catch ex As Exception
			mstrErrorString = "Error in GenerateSQLWhere." & vbNewLine & ex.Message.RemoveSensitive()
			Logs.AddDetailEntry(mstrErrorString)
			Logs.ChangeHeaderStatus(EventLog_Status.elsFailed)
			Return False

		End Try

		Return True

	End Function

	Private Function GenerateSQLOrderBy() As Boolean

		' Purpose : Returns order by string from the sort order array
		Dim strOrder As String
		Dim pblnColumnOK As Boolean
		Dim pblnNoSelect As Boolean
		Dim pblnFound As Boolean
		Dim pstrSource As String
		Dim pintNextIndex As Integer

		Try

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

					End If

					If pblnNoSelect Then
						GenerateSQLOrderBy = False
						mstrErrorString = vbNewLine & "You do not have permission to see the column '" & mstrGroupByColumn & "' " & vbNewLine & "either directly or through any views."
						Exit Function
					End If
				End If
				'*********************************************************************************

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

				If colSortOrder.Count > 0 Then
					' Columns have been defined, so use these for the base table/view
					mstrSQLOrderBy = DoDefinedOrderBy()
				End If

				If Len(mstrSQLOrderBy) > 0 Then mstrSQLOrderBy = " ORDER BY " & mstrSQLOrderBy

			End If

		Catch ex As Exception
			mstrErrorString = "Error in GenerateSQLOrderBy." & vbNewLine & ex.Message.RemoveSensitive()
			Logs.AddDetailEntry(mstrErrorString)
			Logs.ChangeHeaderStatus(EventLog_Status.elsFailed)
			Return False

		End Try

		Return True

	End Function

	Private Function DoDefinedOrderBy() As String

		' This function creates the base ORDER BY statement by searching
		' through the columns defined as the reports sort order, then
		' uses the relevant alias name

		For Each objSort In colSortOrder

			For Each objReportItem In ColumnDetails

				If objSort.ColExprID = objReportItem.ID Then
					DoDefinedOrderBy = DoDefinedOrderBy & IIf(Len(DoDefinedOrderBy) > 0, ",", "") & "[" & objReportItem.IDColumnName & "] " & objSort.AscDesc
					Exit For

				End If

			Next

		Next

	End Function

	Private Function GetTableIDFromColumn(ByVal lngColumnID As Integer) As Integer
		Return Columns.GetById(lngColumnID).TableID
	End Function

	Public Function CheckRecordSet() As Boolean

		' Purpose : To get recordset from temptable and show recordcount
		Dim sSQL As String

		Try

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

			mrstCustomReportsOutput = DB.GetDataTable(sSQL)

			If mrstCustomReportsOutput.Rows.Count = 0 Then
				CheckRecordSet = False
				mstrErrorString = "No records meet the selection criteria."
				Logs.AddDetailEntry("Completed successfully. " & mstrErrorString)
				Logs.ChangeHeaderStatus(EventLog_Status.elsSuccessful)
				mblnNoRecords = True
				Exit Function
			End If

			If mlngColumnLimit > 0 Then
				If mrstCustomReportsOutput.Columns.Count > mlngColumnLimit Then
					CheckRecordSet = False
					mstrErrorString = "Report contains more than " & mlngColumnLimit & " columns. It is not possible to run this report via the intranet."
					Logs.AddDetailEntry("Failed. " & mstrErrorString)
					Logs.ChangeHeaderStatus(EventLog_Status.elsFailed)
					mblnNoRecords = False
					Exit Function
				End If
			End If

		Catch ex As Exception
			mstrErrorString = "Error while checking returned recordset." & vbNewLine & "(" & ex.Message.RemoveSensitive() & ")"
			Logs.AddDetailEntry(mstrErrorString)
			Logs.ChangeHeaderStatus(EventLog_Status.elsFailed)
			Return False

		End Try

		Return True

	End Function

	Public Function PopulateGrid_LoadRecords() As Boolean
		' Purpose : Blimey ! This function does the actual work of populating the
		'           grid, calculating summary info, breaking, page breaking etc.
		'           Its a bit of a 'mare but it works ok. (JDM - I question that!)

		Dim aryAddString As ArrayList
		Dim vDisplayData As String
		Dim colColumns As ICollection(Of ReportColumn)

		Dim vValue As Object
		Dim fBreak As Boolean
		Dim iLoop2 As Integer
		Dim iColumnIndex As Integer
		Dim iOtherColumnIndex As Integer
		Dim fNotFirstTime As Boolean
		Dim bSuppress As Boolean

		Dim intColCounter As Integer

		Dim sBreakValue As String

		'Group With Next Column variables
		Dim intGroupCount As Integer
		Dim blnHasGroupWithNext As Boolean
		Dim blnSkipped As Boolean
		Dim intSkippedIndex As Integer
		Dim strGroupString As String
		Dim sLastValue As String = vbNullString

		blnHasGroupWithNext = False
		blnSkipped = False
		intSkippedIndex = 0

		'Variables for Suppress Repeated Values within Table.
		Dim lngCurrentRecordID As Integer
		Dim bBaseRecordChanged As Boolean
		Dim isHiddenColumn As Boolean

		Dim objReportItem As ReportDetailItem
		Dim otherColumnDetail As ReportDetailItem
		Dim objNextItem As ReportDetailItem
		Dim objReportColumn As ReportColumn

		Try

			' Construct a collection of the columns in the report.
			colColumns = New Collection(Of ReportColumn)
			For Each objReportItem In ColumnDetails
				objReportColumn = New ReportColumn()
				objReportColumn.ID = objReportItem.ID
				objReportColumn.BreakOrPageOnChange = objReportItem.IsBreakOnChange Or objReportItem.IsPageOnChange
				objReportColumn.HasSummaryLine = objReportItem.IsAverage Or objReportItem.IsCount Or objReportItem.IsTotal
				objReportColumn.LastValue = DBNull.Value
				colColumns.Add(objReportColumn)
			Next

			datCustomReportOutput_Start()

			For Each objRow As DataRow In mrstCustomReportsOutput.Rows

				aryAddString = New ArrayList()

				'bRecordChanged used for repetition funcionality.
				If Not mbIsBradfordIndexReport Then
					If CInt(objRow("?ID")) <> lngCurrentRecordID Then
						bBaseRecordChanged = True
						lngCurrentRecordID = CInt(objRow("?ID"))
					Else
						bBaseRecordChanged = False
					End If
				End If

				' Dont do summary info for first record (otherwise blank!)
				'If mrstCustomReportsOutput.AbsolutePosition > 1 Then
				If fNotFirstTime Then
					' Put the values from the previous record in the column array.

					For Each objReportItem In ColumnDetails
						colColumns.GetById(objReportItem.ID).LastValue = objReportItem.LastValue
					Next

					' From last column in the order to first, check changes.
					For iLoop7 = colSortOrder.Count - 1 To 0 Step -1
						' Find the column in the details array.
						iColumnIndex = 0
						iLoop2 = 1
						For Each objReportItem In ColumnDetails
							If (objReportItem.ID = colSortOrder.GetByIndex(iLoop7).ColExprID) And (objReportItem.Type = "C") Then
								iColumnIndex = iLoop2 - 1
								Exit For
							End If
							iLoop2 += 1
						Next

						'If iColumnIndex > 0 Then

						If colColumns.GetByIndex(iColumnIndex).BreakOrPageOnChange Then
							fBreak = False

							objReportItem = ColumnDetails.GetByIndex(iColumnIndex)

							' The column breaks. Check if its changed.
							vValue = PopulateGrid_FormatData(objReportItem, objRow(iColumnIndex), False, False, False)

							'Now that we store the formatted value in position (11) of the mcolDetails
							'Comparison made after adjusting the size of the field.
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							If IsDBNull(vValue) Or IsDBNull(objRow(iColumnIndex)) Then
								fBreak = ("" <> objReportItem.LastValue)
							Else
								If objReportItem.IsBitColumn Then	'Bit
									'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									fBreak = Not (vValue = objReportItem.LastValue)
								Else
									fBreak = (RTrim(LCase(vValue)) <> RTrim(LCase(objReportItem.LastValue)))
								End If
							End If


							If objReportItem.IsPageOnChange Then
								sBreakValue = IIf(Len(objReportItem.LastValue) < 1, "<Empty>", objReportItem.LastValue) & IIf(Len(sBreakValue) > 0, " - ", "").ToString() & sBreakValue
							End If

							If Not fBreak Then
								' The value has not changed, but check if we need to do the summary due to another column changing.
								For iLoop2 = iLoop7 - 1 To 0 Step -1
									iOtherColumnIndex = 0
									For Each objItem In ColumnDetails
										If objItem.ID = colSortOrder.GetByIndex(iLoop2).ColExprID And objItem.Type = "C" Then
											otherColumnDetail = objItem
											Exit For
										End If
										iOtherColumnIndex += 1
									Next

									If Not otherColumnDetail Is Nothing Then
										If colColumns.GetByIndex(iOtherColumnIndex).BreakOrPageOnChange Then
											' The column breaks. Check if its changed.
											If IsDBNull(objRow(iOtherColumnIndex)) And (Not otherColumnDetail.IsNumeric) And (Not otherColumnDetail.IsDateColumn) And (Not otherColumnDetail.IsBitColumn) Then
												' Field value is null but a character data type, so set it to be "".
												vValue = ""

											ElseIf otherColumnDetail.IsNumeric Then	'Numeric
												If IsDBNull(objRow(iOtherColumnIndex)) Then
													vValue = "0"
												Else
													vValue = Left(objRow(iOtherColumnIndex), otherColumnDetail.Size)
												End If


											ElseIf otherColumnDetail.IsBitColumn Then	 'Bit
												If IsDBNull(objRow(iOtherColumnIndex)) Then
													vValue = ""
												Else
													If (objRow(iOtherColumnIndex) = True) Or (objRow(iOtherColumnIndex) = 1) Then vValue = "Y"
													If (objRow(iOtherColumnIndex) = False) Or (objRow(iOtherColumnIndex) = 0) Then vValue = "N"
												End If

											Else
												'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(1, iOtherColumnIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												vValue = Left(objRow(iOtherColumnIndex).ToString(), otherColumnDetail.Size)

											End If

											'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
											If IsDBNull(vValue) Or IsDBNull(objRow(iOtherColumnIndex)) Then
												fBreak = ("" <> otherColumnDetail.LastValue)
											Else
												If otherColumnDetail.IsBitColumn Then	'Bit
													'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													fBreak = (RTrim(LCase(vValue)) <> RTrim(LCase(otherColumnDetail.LastValue)))
												Else
													'TM23112004 Fault 9072
													'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													fBreak = (RTrim(LCase(objRow(iOtherColumnIndex))) <> RTrim(LCase(otherColumnDetail.LastValue)))
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
								PopulateGrid_DoSummaryInfo(colColumns, iColumnIndex, iLoop7)

								' Do the page break ?

								If objReportItem.IsPageOnChange Then
									mblnPageBreak = True
									mblnReportHasPageBreak = True
								End If
							End If
						End If
						'End If
					Next
				End If


				If mblnPageBreak Then
					mvarPageBreak.Add(sBreakValue)
				End If

				mblnPageBreak = False
				sBreakValue = vbNullString

				intColCounter = 1
				' Loop thru each field, adding to the string to add to the grid
				For iLoop = 0 To (mrstCustomReportsOutput.Columns.Count - 1)

					intColCounter = intColCounter + 1
					isHiddenColumn = (mrstCustomReportsOutput.Columns(iLoop).ColumnName.Substring(0, 1) = "?")		' there should be a cleaner way of deciding if this is an ID column, but would need more bigger changes. This will have to do for the moment. Sorry

					' yet another hack beacsue this is an over complex array instead of an easily modifyable class
					If mbIsBradfordIndexReport And iLoop > 12 Then
						isHiddenColumn = True
					End If

					Dim objSkippedItem As ReportDetailItem = ColumnDetails.GetByIndex(intSkippedIndex + 1)

					If Not iLoop = (mrstCustomReportsOutput.Columns.Count - 1) Then
						objNextItem = ColumnDetails.GetByIndex(iLoop)

						' Only suppress values for new records in the Bradford Factor report
						bSuppress = IIf(mbIsBradfordIndexReport And fBreak, False, True)

						If objNextItem.IsBitColumn Then	 'Bit
							'UPGRADE_WARNING: Couldn't resolve default property of object tmpLogicValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If IsDBNull(objRow(iLoop)) Then
								vDisplayData = ""
							Else
								If (objRow(iLoop) = "True") Or (objRow(iLoop) = 1) Then vDisplayData = "Y"
								If (objRow(iLoop) = "False") Or (objRow(iLoop) = 0) Then vDisplayData = "N"
							End If

						Else
							' Get the formatted data to display in the grid
							vDisplayData = PopulateGrid_FormatData(objNextItem, objRow(iLoop), bSuppress, bBaseRecordChanged, True)
						End If

						If blnSkipped Then
							' Store the ACTUAL data in the array (previous value dimension)
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							If IsDBNull(objRow(intSkippedIndex)) And (Not objSkippedItem.IsNumeric) And (Not objSkippedItem.IsDateColumn) And (Not objSkippedItem.IsBitColumn) Then
								' Field value is null but a character data type, so set it to be "".
								objSkippedItem.LastValue = ""

							Else
								'TM17052005 Fault 10086 - Need to store diffent values depending on the type.
								If objSkippedItem.IsDateColumn Then	'Date
									objSkippedItem.LastValue = DateToString(objRow(intSkippedIndex), RegionalSettings)

								ElseIf objSkippedItem.IsNumeric Then	'Numeric
									'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
									objSkippedItem.LastValue = IIf(IsDBNull(objRow(intSkippedIndex)), "", objRow(intSkippedIndex))

								ElseIf (objSkippedItem.IsBitColumn) Then	 'Bit
									If (objRow(intSkippedIndex) = "True") Or (objRow(intSkippedIndex) = 1) Then objSkippedItem.LastValue = "Y"
									If (objRow(intSkippedIndex) = "False") Or (objRow(intSkippedIndex) = 0) Then objSkippedItem.LastValue = "N"
									'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
									If IsDBNull(objRow(intSkippedIndex)) Then objSkippedItem.LastValue = ""

								Else 'Varchar
									objSkippedItem.LastValue = objRow(intSkippedIndex)
								End If

							End If

						Else
							' Store the ACTUAL data in the array (previous value dimension)
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							If IsDBNull(objRow(iLoop)) And (Not objNextItem.IsNumeric) And (Not objNextItem.IsDateColumn) And (Not objNextItem.IsBitColumn) Then
								' Field value is null but a character data type, so set it to be "".
								objNextItem.LastValue = ""
							Else

								'TM17052005 Fault 10086 - Need to store diffent values depending on the type.
								If objNextItem.IsDateColumn Then	'Date
									objNextItem.LastValue = DateToString(objRow(iLoop).ToString(), RegionalSettings)

								ElseIf objNextItem.IsNumeric Then	'Numeric
									objNextItem.LastValue = IIf(IsDBNull(objRow(iLoop)), "", objRow(iLoop))

								ElseIf objNextItem.IsBitColumn Then	 'Bit
									If IsDBNull(objRow(iLoop)) Then
										objNextItem.LastValue = ""
									Else
										If (objRow(iLoop) = "True") Or (objRow(iLoop) = 1) Then objNextItem.LastValue = "Y"
										If (objRow(iLoop) = "False") Or (objRow(iLoop) = 0) Then objNextItem.LastValue = "N"
									End If

								Else 'Varchar
									objNextItem.LastValue = objRow(iLoop).ToString()
								End If
							End If
						End If

						' Group with next column
						If (objNextItem.GroupWithNextColumn And (Not objNextItem.IsHidden)) Then
							If Not vDisplayData = "" Then
								sLastValue = sLastValue & vDisplayData & IIf(Not vDisplayData Is "\n", vbNewLine, "")	'& "<br></br>\n"
							End If

						Else

							If Not isHiddenColumn Then ' hidden columns at end of recordset (hopefully)

								vDisplayData = sLastValue & vDisplayData
								aryAddString.Add(vDisplayData)
								sLastValue = vbNullString

							End If

						End If
					End If

				Next iLoop

				' Only Add the addstring to the grid if its not a summary report
				If mblnCustomReportsSummaryReport = False Then
					If Not NEW_AddToArray_Data(RowType.Data, aryAddString) Then

						Return False

					Else
						If blnHasGroupWithNext Then
							strGroupString = vbNullString
							For intGroupCount = 0 To UBound(mvarGroupWith, 2) Step 1
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarGroupWith(0, intGroupCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								strGroupString = strGroupString & vbNewLine & mvarGroupWith(0, intGroupCount)
							Next intGroupCount

							AddToArray_Data(strGroupString, RowType.Data)

						End If

					End If
				End If

				'Clear the Group Arrays/Variables
				ReDim mvarGroupWith(1, 0)
				blnHasGroupWithNext = False

				fNotFirstTime = True

			Next

			mblnPageBreak = False

			' Now do the final summary for the last bit (before the grand summary)
			' Put the values from the previous record in the column array.

			For Each objReportItem In ColumnDetails
				colColumns.GetById(objReportItem.ID).LastValue = objReportItem.LastValue
			Next


			Dim objColumnIndex As ReportDetailItem

			' From last column in the order to first, check changes.
			For iLoop = colSortOrder.Count - 1 To 0 Step -1
				' Find the column in the details array.
				iColumnIndex = 0
				For Each objItem In ColumnDetails

					If objItem.ID = colSortOrder.GetByIndex(iLoop).ColExprID And objItem.Type = "C" Then
						objColumnIndex = objItem
						iColumnIndex = iLoop2
						Exit For
					End If
				Next


				If objColumnIndex.IsBreakOnChange Or objColumnIndex.IsPageOnChange Then

					If objColumnIndex.IsPageOnChange Then
						mblnPageBreak = True
						sBreakValue = IIf(Len(objColumnIndex.LastValue) < 1, "<Empty>", objColumnIndex.LastValue) & IIf(Len(sBreakValue) > 0, " - ", "") & sBreakValue
					End If

					PopulateGrid_DoSummaryInfo(colColumns, iColumnIndex, iLoop)
				End If

			Next iLoop


			If mblnPageBreak Then
				mvarPageBreak.Add(sBreakValue)
			End If
			sBreakValue = vbNullString

			' Now do the grand summary information
			If Not mbIsBradfordIndexReport Then
				PopulateGrid_DoGrandSummary()
			End If

		Catch ex As Exception
			mstrErrorString = mstrErrorString & "LOADRECORDS_ERROR (In Dll) - Error in PopulateGrid_LoadRecords." & vbNewLine & ex.Message
			Logs.AddDetailEntry(mstrErrorString)
			Logs.ChangeHeaderStatus(EventLog_Status.elsFailed)
			Return False

		End Try

		Return True

	End Function

	Private Function PopulateGrid_FormatData(objReportItem As ReportDetailItem, vData As Object, mbSuppressRepeated As Boolean, pbNewBaseRecord As Boolean, trimData As Boolean) As String
		'Private Function PopulateGrid_FormatData(ByVal sfieldname As String, ByVal vData As Object, ByVal mbSuppressRepeated As Boolean, ByVal pbNewBaseRecord As Boolean) As Object
		' Purpose : Format the data to the form the user has specified to see it
		'           in the grid
		' Input   : None
		' Output  : True/False

		Dim vOriginalData As String

		If IsDBNull(vData) Then Return ""

		vOriginalData = vData

		' Is it a string
		If trimData And objReportItem.DataType = ColumnDataType.sqlVarChar And objReportItem.Size > 0 Then
			vData = Strings.Left(vData, objReportItem.Size)
		End If

		' Is it a boolean calculation ? If so, change to Y or N
		If objReportItem.IsBitColumn Then
			If vData = "True" Then vData = "Y"
			If vData = "False" Then vData = "N"
		End If

		' If its a date column, format it as dateformat
		'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(17, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If objReportItem.IsDateColumn Then
			vData = VB6.Format(vData, mstrClientDateFormat)
		End If


		' Is Numeric
		If objReportItem.IsNumeric Then

			If IsNumeric(vData) Then
				vData = CDbl(vData)

				' Overflow check (ignore decimals)
				If CLng(vData).ToString.Length > objReportItem.Size And objReportItem.Size > 0 Then
					vData = New String("#", objReportItem.Size)
				Else
					If Not objReportItem.Mask Is Nothing Then
						vData = String.Format(objReportItem.Mask, vData)
					End If
				End If
			Else
				vData = "######"
			End If

		End If


		' SRV ?
		If Not mbIsBradfordIndexReport Then
			If mbSuppressRepeated = True Then
				'check if column value should be repeated or not.
				If Not objReportItem.Repetition And Not pbNewBaseRecord And Not objReportItem.SuppressRepeated And Not objReportItem.IsReportChildTable Then
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If CStr(RTrim(IIf(IsDBNull(objReportItem.LastValue), vbNullString, objReportItem.LastValue))) = CStr(RTrim(IIf(IsDBNull(vOriginalData), vbNullString, vOriginalData))) Then
						vData = ""
					End If

				ElseIf objReportItem.SuppressRepeated Then
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If CStr(RTrim(IIf(IsDBNull(objReportItem.LastValue), vbNullString, objReportItem.LastValue))) = CStr(RTrim(IIf(IsDBNull(vOriginalData), vbNullString, vOriginalData))) Then
						vData = ""
					End If
				End If
			End If

		Else
			'Bradford Factor does not use the repetition functionality.
			If mbSuppressRepeated = True Then

				If objReportItem.SuppressRepeated Then
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If CStr(RTrim(IIf(IsDBNull(objReportItem.LastValue), vbNullString, objReportItem.LastValue))) = CStr(RTrim(IIf(IsDBNull(vOriginalData), vbNullString, vOriginalData))) Then
						vData = ""
					End If
				End If
			End If
		End If

		Return vData

	End Function

	Private Sub PopulateGrid_DoSummaryInfo(pavColumns As ICollection(Of ReportColumn), piColumnIndex As Integer, piSortIndex As Integer)

		Dim fDoValue As Boolean
		Dim iLoop As Integer
		Dim sSQL As String
		Dim rsTemp As DataTable
		Dim fHasAverage As Boolean
		Dim fHasCount As Boolean
		Dim fHasTotal As Boolean
		Dim sWhereCode As String = ""
		Dim sFromCode As String

		Dim iLogicValue As Integer

		Dim aryAverageAddString As ArrayList
		Dim aryTotalAddString As ArrayList
		Dim aryCountAddString As ArrayList

		Dim miAmountOfRecords As Integer
		Dim bIsColumnVisible As Boolean
		Dim strAggrValue As String
		Dim objLastItem As New ReportDetailItem
		Dim objThisColumn As ReportColumn

		Dim objReportItem As ReportDetailItem
		Dim objSortItem As ReportSortItem
		Dim iCount As Integer

		strAggrValue = vbNullString

		' Construct the summary where clause.

		For iLoop = 0 To piSortIndex

			objSortItem = colSortOrder.GetByIndex(iLoop)
			objReportItem = ColumnDetails.GetById(objSortItem.ColExprID)

			If objReportItem.IsBreakOnChange Or objReportItem.IsPageOnChange Then
				' The column is a break/page on change column so put it in the Where clause.
				sWhereCode = sWhereCode & IIf(Len(sWhereCode) = 0, " WHERE ", " AND ")

				objThisColumn = pavColumns.GetById(objSortItem.ColExprID)

				If (Not objReportItem.IsNumeric) And (Not objReportItem.IsDateColumn) And (Not objReportItem.IsBitColumn) Then
					' Character column. Treat empty strings along with nulls.
					If Len(objThisColumn.LastValue) = 0 Then
						sWhereCode = sWhereCode & "(([" & CStr(objReportItem.IDColumnName) & "] = '') OR ([" & CStr(objReportItem.IDColumnName) & "] IS NULL))"
					Else
						sWhereCode = sWhereCode & "([" & CStr(objReportItem.IDColumnName) & "] = '" & Replace(objThisColumn.LastValue, "'", "''") & "')"
					End If
				Else
					If IsDBNull(objThisColumn.LastValue) Or objThisColumn.LastValue = "" Then
						sWhereCode = sWhereCode & "([" & CStr(objReportItem.IDColumnName) & "] IS NULL)"
					Else

						If objReportItem.IsDateColumn Then
							' Date column.
							sWhereCode = sWhereCode & "([" & CStr(objReportItem.IDColumnName) & "] = '" & VB6.Format(objThisColumn.LastValue, SQLDateFormat) & "')"
						Else

							If objReportItem.IsBitColumn Then
								' Logic Column.
								iLogicValue = IIf(objThisColumn.LastValue = "Y", 1, 0)
								sWhereCode = sWhereCode & "([" & CStr(objReportItem.IDColumnName) & "] = " & iLogicValue & ")"
							Else
								' Numeric column.
								sWhereCode = sWhereCode & "([" & CStr(objReportItem.IDColumnName) & "] = " & ConvertNumberForSQL(objThisColumn.LastValue) & ")"
							End If
						End If
					End If
				End If
			End If

		Next iLoop

		' Construct the required select statement.
		sSQL = ""
		sFromCode = ""
		aryAverageAddString = New ArrayList()
		aryTotalAddString = New ArrayList()
		aryCountAddString = New ArrayList()

		If Not mbIsBradfordIndexReport Then
			aryAverageAddString.Add("Sub Average")
			aryCountAddString.Add("Sub Count")
			aryTotalAddString.Add("Sub Total")
		End If

		iLoop = 0
		For Each objReportItem In ColumnDetails

			If Not objReportItem.IDColumnName.ToString().Substring(0, 3) = "?ID" Then

				If objReportItem.IsAverage Then
					' Average.

					If Not mbIsBradfordIndexReport Then
						If objReportItem.IsReportChildTable Then
							sSQL = sSQL & ",(SELECT AVG(convert(float,[" & objReportItem.IDColumnName & "])) " & "FROM (SELECT DISTINCT [?ID_" & objReportItem.TableName & "], [" & objReportItem.IDColumnName & "] FROM " & mstrTempTableName & " " & sWhereCode & " "

							If mblnIgnoreZerosInAggregates And objReportItem.IsNumeric Then
								sSQL = sSQL & " AND ([" & objReportItem.IDColumnName & "] <> 0) "
							End If

							sSQL = sSQL & ") AS [vt." & Str(iLoop) & "]) AS 'avg_" & Trim(Str(iLoop)) & "'"
						Else
							sSQL = sSQL & ",(SELECT AVG(convert(float, [" & objReportItem.IDColumnName & "])) " & "FROM (SELECT DISTINCT [?ID], [" & objReportItem.IDColumnName & "] " & "FROM " & mstrTempTableName & " " & " " & sWhereCode & " "

							If mblnIgnoreZerosInAggregates And objReportItem.IsNumeric Then
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sSQL = sSQL & " AND ([" & objReportItem.IDColumnName & "] <> 0) "
							End If

							sSQL = sSQL & ") AS [vt." & Str(iLoop) & "]) AS 'avg_" & Trim(Str(iLoop)) & "'"
						End If
					Else
						'Bradford Index
						sSQL = sSQL & IIf(Len(sSQL) = 0, "SELECT", ",") & " avg(convert(float,[" & objReportItem.IDColumnName & "])) AS avg_" & Trim(Str(iLoop))
					End If
				End If

				If objReportItem.IsCount Then
					' Count.
					If Not mbIsBradfordIndexReport Then
						If objReportItem.IsReportChildTable Then
							sSQL = sSQL & ",(SELECT COUNT([?ID_" & objReportItem.TableName & "]) " & "FROM (SELECT DISTINCT [?ID_" & objReportItem.TableName & "], [" & objReportItem.IDColumnName & "] " & "FROM " & mstrTempTableName & " " & " " & sWhereCode & " "

							If mblnIgnoreZerosInAggregates And objReportItem.IsNumeric Then
								sSQL = sSQL & " AND ([" & objReportItem.IDColumnName & "] <> 0) "
							End If

							sSQL = sSQL & ") AS [vt." & Str(iLoop) & "]) AS 'cnt_" & Trim(Str(iLoop)) & "'"
						Else
							sSQL = sSQL & ",(SELECT COUNT([?ID]) " & "FROM (SELECT DISTINCT [?ID], [" & objReportItem.IDColumnName & "] " & "FROM " & mstrTempTableName & " " & sWhereCode & " "

							If mblnIgnoreZerosInAggregates And objReportItem.IsNumeric Then
								sSQL = sSQL & " AND ([" & objReportItem.IDColumnName & "] <> 0) "
							End If

							sSQL = sSQL & ") AS [vt." & Str(iLoop) & "]) AS 'cnt_" & Trim(Str(iLoop)) & "'"
						End If
					Else
						'Bradford Index
						sSQL = sSQL & IIf(Len(sSQL) = 0, "SELECT", ",") & " count(*) AS cnt_" & Trim(Str(iLoop))
					End If
				End If

				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(6, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If objReportItem.IsTotal Then
					' Total.
					If Not mbIsBradfordIndexReport Then
						If objReportItem.IsReportChildTable Then
							sSQL = sSQL & ",(SELECT SUM([" & objReportItem.IDColumnName & "]) " & "FROM (SELECT DISTINCT [?ID_" & objReportItem.TableName & "], [" & objReportItem.IDColumnName & "] FROM " & mstrTempTableName & " " & " " & sWhereCode & " "

							If mblnIgnoreZerosInAggregates And objReportItem.IsNumeric Then
								'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sSQL = sSQL & " AND ([" & objReportItem.IDColumnName & "] <> 0) "
							End If

							sSQL = sSQL & ") AS [vt." & Str(iLoop) & "]) AS 'ttl_" & Trim(Str(iLoop)) & "'"
						Else
							sSQL = sSQL & ",(SELECT SUM([" & objReportItem.IDColumnName & "]) " & "FROM (SELECT DISTINCT [?ID], [" & objReportItem.IDColumnName & "] " & "FROM " & mstrTempTableName & " " & sWhereCode & " "

							If mblnIgnoreZerosInAggregates And objReportItem.IsNumeric Then
								sSQL = sSQL & " AND ([" & objReportItem.IDColumnName & "] <> 0) "
							End If

							sSQL = sSQL & ") AS [vt." & Str(iLoop) & "]) AS 'ttl_" & Trim(Str(iLoop)) & "'"
						End If
					Else
						'Bradford Index
						sSQL = sSQL & IIf(Len(sSQL) = 0, "SELECT", ",") & " sum([" & objReportItem.IDColumnName & "])  AS ttl_" & Trim(Str(iLoop))
					End If
				End If
			End If

			iLoop += 1
		Next

		If Len(sSQL) > 0 Then
			If Not mbIsBradfordIndexReport Then
				sSQL = "SELECT " & Right(sSQL, Len(sSQL) - 1)
			Else
				sSQL = sSQL & " FROM " & mstrTempTableName & IIf(Len(sFromCode) > 0, sFromCode, "") & IIf(Len(sWhereCode) > 0, sWhereCode, "")
			End If

			rsTemp = DB.GetDataTable(sSQL)

			iLoop = 0
			For Each objReportItem In ColumnDetails

				bIsColumnVisible = Not objReportItem.IsHidden And (Not objReportItem.GroupWithNextColumn) And (Not ColumnDetails.GetByIndex(iLoop).GroupWithNextColumn)

				Dim rowData As DataRow = rsTemp.Rows(0)

				If Not objReportItem.IDColumnName.ToString().Substring(0, 3) = "?ID" Then

					If objReportItem.IsAverage Then

						If Not objReportItem.IsHidden And (Not objReportItem.GroupWithNextColumn) And (Not ColumnDetails.GetByIndex(iLoop).GroupWithNextColumn) Then
							fHasAverage = True
						End If

						' Average.
						If IsDBNull(rowData("avg_" & Trim(Str(iLoop)))) Then
							strAggrValue = "0"
						Else
							strAggrValue = FormatNumber(rowData("avg_" & iLoop.ToString), objReportItem.Decimals, , , objReportItem.Use1000Separator)
						End If

						aryAverageAddString.Add(strAggrValue)

						strAggrValue = vbNullString
					Else

						' Display the value ?
						fDoValue = False
						If (objReportItem.IsValueOnChange) Then
							If colSortOrder.Any(Function(objSort) objSort.ColExprID = objReportItem.ID) Then
								fDoValue = True
							End If
						End If

						If (fDoValue Or mbIsBradfordIndexReport) And bIsColumnVisible Then
							aryAverageAddString.Add(PopulateGrid_FormatData(objReportItem, objReportItem.LastValue, False, True, True))
						ElseIf bIsColumnVisible Then
							aryAverageAddString.Add("")
						End If

					End If

					If objReportItem.IsCount Then

						If Not objReportItem.IsHidden And (Not objReportItem.GroupWithNextColumn) And (Not ColumnDetails.GetByIndex(iLoop).GroupWithNextColumn) Then
							fHasCount = True
						End If

						'JDM - Make a note of count the Bradford Index Report
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						If mbIsBradfordIndexReport Then miAmountOfRecords = IIf(Not IsDBNull(rowData("cnt_" & Trim(Str(iLoop)))), rowData("cnt_" & Trim(Str(iLoop))), 0)

						' Count.
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						aryCountAddString.Add(IIf(IsDBNull(rowData("cnt_" & Trim(Str(iLoop)))), "0", Format(rowData("cnt_" & Trim(Str(iLoop))), "0")))
					Else

						' Display the value ?
						fDoValue = False
						If (objReportItem.IsValueOnChange) Then
							If colSortOrder.Any(Function(objSort) objSort.ColExprID = objReportItem.ID) Then
								fDoValue = True
							End If
						End If

						If (fDoValue Or mbIsBradfordIndexReport) And bIsColumnVisible Then
							aryCountAddString.Add(PopulateGrid_FormatData(objReportItem, objReportItem.LastValue, False, True, True))
						ElseIf bIsColumnVisible Then
							aryCountAddString.Add("")
						End If

					End If

					If objReportItem.IsTotal Then
						' Total.

						If Not objReportItem.IsHidden And (Not objReportItem.GroupWithNextColumn) And (Not objLastItem.GroupWithNextColumn) Then
							fHasTotal = True
						End If

						If IsDBNull(rowData("ttl_" & iLoop.ToString)) Then
							strAggrValue = "0"
						Else
							strAggrValue = FormatNumber(rowData("ttl_" & iLoop.ToString), objReportItem.Decimals, , , objReportItem.Use1000Separator)
						End If

						aryTotalAddString.Add(strAggrValue)
						strAggrValue = vbNullString

					Else
						' Display the value ?
						fDoValue = False
						If objReportItem.IsValueOnChange Then
							If colSortOrder.Any(Function(objSort) objSort.ColExprID = objReportItem.ID) Then
								fDoValue = True
							End If
						End If

						If (fDoValue Or mbIsBradfordIndexReport) And bIsColumnVisible Then
							aryTotalAddString.Add(PopulateGrid_FormatData(objReportItem, objReportItem.LastValue, False, True, True))
						ElseIf bIsColumnVisible Then
							aryTotalAddString.Add("")
						End If

					End If
				End If

				objLastItem = objReportItem
				iLoop += 1
			Next

			'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsTemp = Nothing
		End If

		' Do a different summary if we are a Bradford Index Report
		If Not mbIsBradfordIndexReport Then

			If fHasAverage Then
				NEW_AddToArray_Data(RowType.Average, aryAverageAddString)
			End If

			If fHasCount Then
				NEW_AddToArray_Data(RowType.Count, aryCountAddString)
			End If

			If fHasTotal Then
				NEW_AddToArray_Data(RowType.Total, aryTotalAddString)
			End If

			If Not mblnCustomReportsSummaryReport Then
				NEW_AddToArray_Data(RowType.PageBreak, "")
			End If

		Else

			Dim iSummaryColumn As Integer = CInt(If(mbDisplayBradfordDetail, 9, 4))
			Dim iDurationColumn As Integer = CInt(If(mbDisplayBradfordDetail, 11, 5))
			Dim iIncludedDaysColumn As Integer = CInt(If(mbDisplayBradfordDetail, 12, 6))

			If mbDisplayBradfordDetail Then
				For iCount = 0 To iSummaryColumn
					aryCountAddString(iCount) = vbNullString
					aryTotalAddString(iCount) = vbNullString
				Next
			End If

			' Add the summary lines
			If mbBradfordCount Then
				aryCountAddString(iSummaryColumn) = "Instances"
				NEW_AddToArray_Data(RowType.Count, aryCountAddString)
				For iCount = 0 To iSummaryColumn
					aryCountAddString(iCount) = vbNullString
					aryTotalAddString(iCount) = vbNullString
				Next
			End If

			If mbBradfordTotals Then
				aryTotalAddString(iSummaryColumn) = "Total"
				NEW_AddToArray_Data(RowType.Total, aryTotalAddString)
				For iCount = 0 To iSummaryColumn
					aryCountAddString(iCount) = vbNullString
					aryTotalAddString(iCount) = vbNullString
				Next
			End If

			' Calculate Bradford index line
			aryTotalAddString(iSummaryColumn) = "Bradford Factor"

			If mbBradfordWorkings = True Then
				aryTotalAddString(iDurationColumn) = CStr(Val(aryTotalAddString(iDurationColumn)) * (miAmountOfRecords * miAmountOfRecords)) & " (" & Str(miAmountOfRecords) & Chr(178) & " * " & aryTotalAddString(iDurationColumn) & ")"
				aryTotalAddString(iIncludedDaysColumn) = CStr(Val(aryTotalAddString(iIncludedDaysColumn)) * (miAmountOfRecords * miAmountOfRecords)) & " (" & Str(miAmountOfRecords) & Chr(178) & " * " & aryTotalAddString(iIncludedDaysColumn) & ")"
			Else
				aryTotalAddString(iDurationColumn) = CStr(CDbl(aryTotalAddString(iDurationColumn)) * (miAmountOfRecords * miAmountOfRecords))
				aryTotalAddString(iIncludedDaysColumn) = CStr(CDbl(aryTotalAddString(iIncludedDaysColumn)) * (miAmountOfRecords * miAmountOfRecords))
			End If

			NEW_AddToArray_Data(RowType.BradfordCalculation, aryTotalAddString)
			NEW_AddToArray_Data(RowType.PageBreak, "")

			End If

	End Sub

	Private Sub PopulateGrid_DoGrandSummary()

		' Purpose : To calculate the final grand summaries
		' Input   : None
		' Output  : True/False

		Dim iLoop As Integer = 0
		Dim rsTemp As DataTable

		Dim aryAverageAddString As New ArrayList
		Dim aryCountAddString As New ArrayList
		Dim aryTotalAddString As New ArrayList

		Dim objPrevious As ReportDetailItem

		Dim fHasAverage As Boolean
		Dim fHasCount As Boolean
		Dim fHasTotal As Boolean
		Dim bIsColumnVisible As Boolean

		Dim sSQL As String

		Dim strAggrValue As String

		strAggrValue = vbNullString

		' Construct the required select statement.
		sSQL = vbNullString

		aryAverageAddString.Add("Average")
		aryCountAddString.Add("Count")
		aryTotalAddString.Add("Total")

		Try

			For Each objReportItem In ColumnDetails

				'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(4, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If objReportItem.IsAverage Then
					' Average.

					If objReportItem.IsReportChildTable Then
						sSQL = sSQL & ",(SELECT AVG(convert(float, [" & objReportItem.IDColumnName & "])) FROM (SELECT DISTINCT [?ID_" & objReportItem.TableName & "], [" & objReportItem.IDColumnName & "] FROM " & mstrTempTableName & " "

						If mblnIgnoreZerosInAggregates And objReportItem.IsNumeric Then
							sSQL = sSQL & "WHERE ([" & objReportItem.IDColumnName & "] <> 0) "
						End If

						sSQL = sSQL & " ) AS [vt." & Str(iLoop) & "]) AS 'avg_" & Trim(Str(iLoop)) & "'"
					Else
						sSQL = sSQL & ",(SELECT AVG(convert(float, [" & objReportItem.IDColumnName & "])) FROM (SELECT DISTINCT [?ID], [" & objReportItem.IDColumnName & "] FROM " & mstrTempTableName & " "

						If mblnIgnoreZerosInAggregates And objReportItem.IsNumeric Then
							sSQL = sSQL & "WHERE ([" & objReportItem.IDColumnName & "] <> 0) "
						End If

						sSQL = sSQL & " ) AS [vt." & Str(iLoop) & "]) AS 'avg_" & Trim(Str(iLoop)) & "'"
					End If

				End If

				If objReportItem.IsCount Then
					' Count.

					If objReportItem.IsReportChildTable Then
						sSQL = sSQL & ",(SELECT COUNT([?ID_" & objReportItem.TableName & "]) FROM (SELECT DISTINCT [?ID_" & objReportItem.TableName & "], [" & objReportItem.IDColumnName & "] FROM " & mstrTempTableName & " "

						If mblnIgnoreZerosInAggregates And objReportItem.IsNumeric Then
							sSQL = sSQL & "WHERE ([" & objReportItem.IDColumnName & "] <> 0) "
						End If

						sSQL = sSQL & " ) AS [vt." & Str(iLoop) & "]) AS 'cnt_" & Trim(Str(iLoop)) & "'"
					Else
						sSQL = sSQL & ",(SELECT COUNT([?ID]) FROM (SELECT DISTINCT [?ID], [" & objReportItem.IDColumnName & "] FROM " & mstrTempTableName & " "

						If mblnIgnoreZerosInAggregates And objReportItem.IsNumeric Then
							sSQL = sSQL & "WHERE ([" & objReportItem.IDColumnName & "] <> 0) "
						End If

						sSQL = sSQL & " ) AS [vt." & Str(iLoop) & "]) AS 'cnt_" & Trim(Str(iLoop)) & "'"
					End If

				End If

				If objReportItem.IsTotal Then
					' Total.

					If objReportItem.IsReportChildTable Then
						sSQL = sSQL & ",(SELECT SUM([" & objReportItem.IDColumnName & "]) FROM (SELECT DISTINCT [?ID_" & objReportItem.TableName & "], [" & objReportItem.IDColumnName & "] FROM " & mstrTempTableName & " "

						If mblnIgnoreZerosInAggregates And objReportItem.IsNumeric Then
							sSQL = sSQL & "WHERE ([" & objReportItem.IDColumnName & "] <> 0) "
						End If

						sSQL = sSQL & " ) AS [vt." & Str(iLoop) & "]) AS 'ttl_" & Trim(Str(iLoop)) & "'"
					Else
						sSQL = sSQL & ",(SELECT SUM([" & objReportItem.IDColumnName & "]) FROM (SELECT DISTINCT [?ID], [" & objReportItem.IDColumnName & "] FROM " & mstrTempTableName & " "

						If mblnIgnoreZerosInAggregates And objReportItem.IsNumeric Then
							sSQL = sSQL & "WHERE ([" & objReportItem.IDColumnName & "] <> 0) "
						End If

						sSQL = sSQL & " ) AS [vt." & Str(iLoop) & "]) AS 'ttl_" & Trim(Str(iLoop)) & "'"
					End If

				End If

				iLoop += 1

			Next

			iLoop = 0
			If Len(sSQL) > 0 Then
				sSQL = "SELECT " & Right(sSQL, Len(sSQL) - 1)

				rsTemp = DB.GetDataTable(sSQL)
				Dim rowData As DataRow = rsTemp.Rows(0)

				For Each objReportItem In ColumnDetails

					bIsColumnVisible = Not objReportItem.IsHidden And (Not objReportItem.GroupWithNextColumn) And (Not ColumnDetails.GetByIndex(iLoop).GroupWithNextColumn)
					If iLoop > 0 Then objPrevious = ColumnDetails.GetByIndex(iLoop - 1) Else objPrevious = New ReportDetailItem

					If objReportItem.IsAverage Then
						' Average.

						If Not objReportItem.IsHidden And (Not objReportItem.GroupWithNextColumn) And (Not objPrevious.GroupWithNextColumn) Then
							fHasAverage = True
						End If

						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						If IsDBNull(rowData("avg_" & Trim(Str(iLoop)))) Then
							strAggrValue = "0"
						Else
							strAggrValue = FormatNumber(rowData("avg_" & iLoop.ToString), objReportItem.Decimals, , , objReportItem.Use1000Separator)
						End If

						aryAverageAddString.Add(strAggrValue)
						strAggrValue = vbNullString

					ElseIf bIsColumnVisible Then
						aryAverageAddString.Add("")
					End If


					If objReportItem.IsCount Then
						' Count.

						If Not objReportItem.IsHidden And (Not objReportItem.GroupWithNextColumn) And (Not objPrevious.GroupWithNextColumn) Then
							fHasCount = True
						End If

						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						aryCountAddString.Add(IIf(IsDBNull(rowData("cnt_" & Trim(Str(iLoop)))), "0", Format(rowData("cnt_" & Trim(Str(iLoop))), "0")))
					ElseIf bIsColumnVisible Then
						aryCountAddString.Add("")
					End If

					'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(6, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If objReportItem.IsTotal Then
						' Total.

						If Not objReportItem.IsHidden And (Not objReportItem.GroupWithNextColumn) And (Not objPrevious.GroupWithNextColumn) Then
							fHasTotal = True
						End If

						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						If IsDBNull(rowData("ttl_" & iLoop.ToString())) Then
							strAggrValue = "0"
						Else
							strAggrValue = FormatNumber(rowData("ttl_" & iLoop.ToString), objReportItem.Decimals, , , objReportItem.Use1000Separator)
						End If

						aryTotalAddString.Add(strAggrValue)
						strAggrValue = vbNullString
					ElseIf bIsColumnVisible Then
						aryTotalAddString.Add("")
					End If

					iLoop += 1

				Next

				'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				rsTemp = Nothing
			End If

			mblnDoesHaveGrandSummary = (fHasAverage Or fHasCount Or fHasTotal)

			'Output the grand aggregates (AVG,CNT,TTL)
			If fHasAverage Then
				NEW_AddToArray_Data(RowType.GrandSummary, aryAverageAddString)
			End If

			If fHasCount Then
				NEW_AddToArray_Data(RowType.GrandSummary, aryCountAddString)
			End If

			If fHasTotal Then
				NEW_AddToArray_Data(RowType.GrandSummary, aryTotalAddString)
			End If

		Catch ex As Exception
			mstrErrorString = "Error while calculating grand summary." & vbNewLine & "(" & ex.Message & ")"

		End Try

	End Sub

	Public Function PopulateGrid_HideColumns() As Boolean

		' Purpose : This function hides any columns we don't want the user to see.
		Dim pblnOK As Boolean
		Dim intColCounter As Short
		Dim intVisColCount As Short
		Dim objLastItem As New ReportDetailItem

		pblnOK = True

		intVisColCount = 0
		intColCounter = 0

		'If report contains no summary info, hide the column
		intColCounter = intColCounter + 1
		If (Not mblnReportHasSummaryInfo) Or (mbIsBradfordIndexReport) Then
		Else

			'UPGRADE_WARNING: Couldn't resolve default property of object mvarVisibleColumns(0, intVisColCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarVisibleColumns(0, intVisColCount) = "Summary Info" 'Heading
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarVisibleColumns(1, intVisColCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarVisibleColumns(1, intVisColCount) = ColumnDataType.sqlVarChar	'DataType
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarVisibleColumns(2, intVisColCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarVisibleColumns(2, intVisColCount) = 0	'Decimals
			'UPGRADE_WARNING: Couldn't resolve default property of object mvarVisibleColumns(3, intVisColCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvarVisibleColumns(3, intVisColCount) = 0	'1000 Separator
			intVisColCount = intVisColCount + 1
		End If

		For Each objReportItem In ColumnDetails
			intColCounter = intColCounter + 1

			If Not objLastItem.GroupWithNextColumn Then

				If (Not objReportItem.IsHidden) Then
					ReDim Preserve mvarVisibleColumns(3, intVisColCount)
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarVisibleColumns(0, intVisColCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarVisibleColumns(0, intVisColCount) = objReportItem.IDColumnName	 'Heading

					'TM20050901 - Fault 10291
					'If the column is is grouped then don't force the date format.
					If objReportItem.IsDateColumn And objReportItem.GroupWithNextColumn = False Then
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarVisibleColumns(1, intVisColCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarVisibleColumns(1, intVisColCount) = ColumnDataType.sqlDate
					ElseIf objReportItem.IsNumeric And objReportItem.GroupWithNextColumn = False Then
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarVisibleColumns(1, intVisColCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarVisibleColumns(1, intVisColCount) = ColumnDataType.sqlNumeric
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object mvarVisibleColumns(1, intVisColCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvarVisibleColumns(1, intVisColCount) = ColumnDataType.sqlVarChar
					End If

					'UPGRADE_WARNING: Couldn't resolve default property of object mvarVisibleColumns(2, intVisColCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarVisibleColumns(2, intVisColCount) = objReportItem.Decimals	'Decimals
					'UPGRADE_WARNING: Couldn't resolve default property of object mvarVisibleColumns(3, intVisColCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvarVisibleColumns(3, intVisColCount) = IIf(objReportItem.Use1000Separator, 1, 0)	'1000 Separator.
					intVisColCount = intVisColCount + 1
				End If

			End If

			objLastItem = objReportItem

		Next

		'	mintVisibleColumnCount = intVisColCount - 1

		Return True

	End Function

	Public Function ClearUp() As Boolean

		Try
			AccessLog.UtilUpdateLastRun(UtilityType.utlCustomReport, mlngCustomReportID)

			' Delete the temptable if exists
			General.DropUniqueSQLObject(mstrTempTableName, 3)

			Return True

		Catch ex As Exception
			Throw

		End Try

	End Function

	Private Function IsRecordSelectionValid() As Boolean
		Dim sSQL As String
		Dim rsTemp As DataTable
		Dim iResult As RecordSelectionValidityCodes
		Dim fCurrentUserIsSysSecMgr As Boolean
		Dim i As Short
		Dim lngFilterID As Integer

		fCurrentUserIsSysSecMgr = CurrentUserIsSysSecMgr()

		' Base Table First
		If mlngSingleRecordID = 0 Then
			If mlngCustomReportsFilterID > 0 Then
				iResult = ValidateRecordSelection(RecordSelectionType.Filter, mlngCustomReportsFilterID)
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
				iResult = ValidateRecordSelection(RecordSelectionType.Picklist, mlngCustomReportsPickListID)
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
				iResult = ValidateRecordSelection(RecordSelectionType.Filter, mlngCustomReportsParent1FilterID)
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
				iResult = ValidateRecordSelection(RecordSelectionType.Picklist, mlngCustomReportsParent1PickListID)
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
				iResult = ValidateRecordSelection(RecordSelectionType.Filter, mlngCustomReportsParent2FilterID)
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
				iResult = ValidateRecordSelection(RecordSelectionType.Picklist, mlngCustomReportsParent2PickListID)
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
						iResult = ValidateRecordSelection(RecordSelectionType.Filter, lngFilterID)
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

				rsTemp = DB.GetDataTable(sSQL)
				For Each objRow As DataRow In rsTemp.Rows

					iResult = ValidateCalculation(CInt(objRow("ColExprID")))
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
						Exit For
					End If

				Next
			End If

			'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsTemp = Nothing
		End If

		Return (Len(mstrErrorString) = 0)

	End Function

	Private Function CheckCalcsStillExist() As Boolean

		Dim pstrBadCalcs As String
		Dim prstTemp As DataTable

		Try

			For Each objRow As DataRow In mrstCustomReportsDetails.Rows

				If objRow("Type").ToString() = "E" Then
					prstTemp = DB.GetDataTable("SELECT * FROM AsrSysExpressions WHERE ExprID = " & objRow("ColExprID"))
					If prstTemp.Rows.Count = 0 Then
						pstrBadCalcs = "One or more calculation(s) used in this report have been deleted" & vbNewLine & "by another user."
						Exit For
					End If
				End If
			Next

			If Len(pstrBadCalcs) > 0 Then
				mstrErrorString = pstrBadCalcs
				Return False
			End If

		Catch ex As Exception
			mstrErrorString = "Error checking if calcs still exist." & vbNewLine & ex.Message.RemoveSensitive()
			Return False

		End Try

		Return True

	End Function

	Public ReadOnly Property ReportCaption() As String
		Get

			Dim sCaption As String = Name

			Try

				If mblnCustomReportsPrintFilterHeader And (mlngSingleRecordID = 0) Then
					If (mlngCustomReportsFilterID > 0) Then
						sCaption = sCaption & " (Base Table filter : " & General.GetFilterName(mlngCustomReportsFilterID) & ")"
					ElseIf (mlngCustomReportsPickListID > 0) Then
						sCaption = sCaption & " (Base Table picklist : " & General.GetPicklistName(mlngCustomReportsPickListID) & ")"
					Else
						sCaption = sCaption & " (All records)"
					End If
				End If

				Return sCaption

			Catch ex As Exception
				Return String.Format("{0} {1})", sCaption, ex.Message)

			End Try
		End Get
	End Property

	Private Function NEW_AddToArray_Data(RowType As RowType, data As IEnumerable) As Boolean

		Dim dr As DataRow
		Dim iColumn As Integer

		dr = ReportDataTable.Rows.Add()
		dr(0) = RowType

		Select Case RowType
			Case RowType.Data
				iColumn = IIf(mblnReportHasSummaryInfo, 2, 1)
			Case RowType.PageBreak
				iColumn = 1
			Case Else
				iColumn = 1
		End Select

		If Not data Is Nothing Then
			For Each objData In data
				dr(iColumn) = objData
				iColumn += 1
			Next
		End If

		Return True

	End Function

	Private Sub AddToArray_Data(pstrRowToAdd As String, rowType As RowType)

		Dim sClassName As String = "rowdata"

		Select Case rowType
			Case rowType.Count, rowType.Average, rowType.Total
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

	End Sub

	Public Function GenerateSQLBradford(pstrIncludeTypes As String) As Boolean

		' NOTE: Checks are made elsewhere to ensure that from and to dates are not blank
		' NOTE: Put in some code to handle blank end dates (do we include as an option on the main screen ?)

		Dim strAbsenceType As String
		Dim iCount As Short
		Dim astrIncludeTypes() As String
		Dim objBradfordDetail As ReportDetailItem

		Try

			' Get the absence start/end field details
			strAbsenceType = mstrAbsenceRealSource & "." & GetColumnName(Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPE)))

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


			'Personel ID
			objBradfordDetail = New ReportDetailItem
			objBradfordDetail.IDColumnName = "Personnel_ID"
			objBradfordDetail.Size = 99
			objBradfordDetail.Decimals = 0
			objBradfordDetail.IsNumeric = False
			objBradfordDetail.IsAverage = False
			objBradfordDetail.IsCount = False
			objBradfordDetail.IsTotal = False
			objBradfordDetail.IsBreakOnChange = True
			objBradfordDetail.IsPageOnChange = False
			objBradfordDetail.IsValueOnChange = True
			objBradfordDetail.SuppressRepeated = False
			objBradfordDetail.LastValue = ""
			objBradfordDetail.ID = -1
			objBradfordDetail.Type = "C"
			objBradfordDetail.TableID = 0
			objBradfordDetail.TableName = ""
			objBradfordDetail.ColumnName = ""
			objBradfordDetail.IsDateColumn = False
			objBradfordDetail.IsBitColumn = False
			objBradfordDetail.IsHidden = True	' Is column hidden
			objBradfordDetail.IsReportChildTable = False
			objBradfordDetail.Repetition = False
			ColumnDetails.Add(objBradfordDetail)

			'Absence ID
			objBradfordDetail = New ReportDetailItem
			objBradfordDetail.IDColumnName = "Absence_ID"
			objBradfordDetail.Size = 99
			objBradfordDetail.Decimals = 0
			objBradfordDetail.IsNumeric = False
			objBradfordDetail.IsAverage = False
			objBradfordDetail.IsCount = False
			objBradfordDetail.IsTotal = False
			objBradfordDetail.IsBreakOnChange = True
			objBradfordDetail.IsPageOnChange = False
			objBradfordDetail.IsValueOnChange = True
			objBradfordDetail.SuppressRepeated = False
			objBradfordDetail.LastValue = ""
			objBradfordDetail.ID = -1
			objBradfordDetail.Type = "C"
			objBradfordDetail.TableID = 0
			objBradfordDetail.TableName = ""
			objBradfordDetail.ColumnName = ""
			objBradfordDetail.IsDateColumn = False
			objBradfordDetail.IsBitColumn = False
			objBradfordDetail.IsHidden = True	' Is column hidden
			objBradfordDetail.IsReportChildTable = False
			objBradfordDetail.Repetition = False
			ColumnDetails.Add(objBradfordDetail)

			' All done correctly

		Catch ex As Exception
			mstrErrorString = "Error in GenerateSQLBradford." & vbNewLine & ex.Message.RemoveSensitive()
			Return False

		End Try

		Return True


	End Function

	Public Function CalculateBradfordFactors() As Boolean

		' Purpose : To calculate any bradford factors, and place into the created temporary table
		Dim sSQL As String

		Try

			' Merge the absence records if the continuous field is defined.
			If Not CInt(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCECONTINUOUS)) = 0 Then
				sSQL = "EXECUTE sp_ASR_Bradford_MergeAbsences '" & CDate(mdBradfordStartDate).ToString(SQLDateFormat) & "','" & CDate(mdBradfordEndDate).ToString(SQLDateFormat) & "','" & mstrTempTableName & "'"
				DB.ExecuteSql(sSQL)
			End If

			' Delete unwanted absences from the table.
			sSQL = "EXECUTE sp_ASR_Bradford_DeleteAbsences '" & CDate(mdBradfordStartDate).ToString(SQLDateFormat) & "','" & CDate(mdBradfordEndDate).ToString(SQLDateFormat) & "'," + IIf(mbOmitBeforeStart, "1,", "0,") + IIf(mbOmitAfterEnd, "1,'", "0,'") + mstrTempTableName + "'"
			DB.ExecuteSql(sSQL)

			' Calculate the included durations for the absences.
			sSQL = "EXECUTE sp_ASR_Bradford_CalculateDurations '" & CDate(mdBradfordStartDate).ToString(SQLDateFormat) & "','" & CDate(mdBradfordEndDate).ToString(SQLDateFormat) & "','" & mstrTempTableName & "'"
			DB.ExecuteSql(sSQL)

			' Remove absences that are below the required Bradford Factor
			If mbMinBradford Then
				sSQL = "DELETE FROM " & mstrTempTableName & " WHERE personnel_id IN (SELECT personnel_id FROM " & mstrTempTableName & " GROUP BY personnel_id HAVING((count(duration)*count(duration))*sum(duration)) < " & Str(mlngMinBradfordAmount) & ")"
				DB.ExecuteSql(sSQL)
			End If

		Catch ex As Exception
			mstrErrorString = "Error while checking calculating Bradford factors." & vbNewLine & "(" & ex.Message & ")"
			Return False

		End Try

		Return True

	End Function

	' Dates are in SQL (American format)
	Public Function GetBradfordReportDefinition(pdtAbsenceFrom As Date, pdtAbsenceTo As Date) As Boolean

		' Purpose : This function retrieves the basic definition details
		'           and stores it in module level variables

		Try

			mbIsBradfordIndexReport = True

			' Dates coming in are in American format (if they're not we have a problem)
			mdBradfordStartDate = pdtAbsenceFrom
			mdBradfordEndDate = pdtAbsenceTo

			If DateDiff(DateInterval.Day, pdtAbsenceFrom, pdtAbsenceTo) < 0 Then
				mstrErrorString = "The report end date is before the report start date."
				Logs.AddDetailEntry(mstrErrorString)
				Logs.ChangeHeaderStatus(EventLog_Status.elsFailed)
				Return False
			End If


			'Set the grid header with no picklist/filter information
			Name = "Bradford Factor Report (" & CDate(mdBradfordStartDate).ToString(LocaleDateFormat) & " - " & CDate(mdBradfordEndDate).ToString(LocaleDateFormat) & ")"

			mlngCustomReportsBaseTable = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE))
			mstrCustomReportsBaseTableName = GetTableName(mlngCustomReportsBaseTable)
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
			mvarChildTables(3, 0) = GetTableName(mvarChildTables(0, 0))	'Child Table Name
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
				Return False
			End If

			Logs.AddHeader(EventLog_Type.eltStandardReport, "Bradford Factor")
			Return True

		Catch ex As Exception
			mstrErrorString = "Error whilst reading the Bradford Factor Report definition !" & vbNewLine & ex.Message
			Logs.AddDetailEntry(mstrErrorString)
			Logs.ChangeHeaderStatus(EventLog_Status.elsFailed)
			Return False

		End Try

	End Function

	Public Function GetBradfordRecordSet() As Boolean

		' Purpose : This function loads report details and sort details into
		'           arrays and leaves the details recordset reference there
		'           (dont remove it...used for summary info !)


		Dim lngTableID As Integer
		Dim iCount As Short
		Dim lngColumnID As Integer

		Dim lbHideStaffNumber As Boolean

		Dim aStrRequiredFields(15, 1) As String

		Try

			aStrRequiredFields(1, 1) = GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_EMPLOYEENUMBER)
			aStrRequiredFields(2, 1) = GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_SURNAME)
			aStrRequiredFields(3, 1) = GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_FORENAME)
			aStrRequiredFields(4, 1) = GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_DEPARTMENT)

			aStrRequiredFields(5, 1) = GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPE)
			aStrRequiredFields(6, 1) = GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESTARTDATE)
			aStrRequiredFields(7, 1) = GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESTARTSESSION)
			aStrRequiredFields(8, 1) = GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEENDDATE)
			aStrRequiredFields(9, 1) = GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEENDSESSION)
			aStrRequiredFields(10, 1) = GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEREASON)
			aStrRequiredFields(11, 1) = GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCECONTINUOUS)
			aStrRequiredFields(12, 1) = GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEDURATION)

			'This field is later recalculated for the included days
			aStrRequiredFields(13, 1) = GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEDURATION)

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
				aStrRequiredFields(1, 1) = GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_SURNAME)
				lbHideStaffNumber = True
			End If

			' Allow the continuous field to be undefined (Let system read the absence reason)
			If aStrRequiredFields(11, 1) = "0" Then
				aStrRequiredFields(11, 1) = GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEDURATION)
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
			Dim objReportItem As ReportDetailItem

			ColumnDetails = New List(Of ReportDetailItem)
			For iCount = 1 To UBound(aStrRequiredFields, 1)

				If CInt(aStrRequiredFields(iCount, 1)) <> -1 Then

					objReportItem = New ReportDetailItem

					lngColumnID = CInt(aStrRequiredFields(iCount, 1))
					lngTableID = GetTableIDFromColumn(lngColumnID)

					objReportItem.IsBreakOnChange = False
					objReportItem.Size = 99
					objReportItem.Decimals = 0
					objReportItem.IsNumeric = (GetDataType(lngTableID, lngColumnID) = 2)
					objReportItem.IsAverage = False
					objReportItem.IsCount = False
					objReportItem.IsTotal = False


					' Specify the column names and whether they are visible or not
					Select Case iCount
						Case 1
							objReportItem.IDColumnName = "Staff_No"
							objReportItem.IsHidden = lbHideStaffNumber
							objReportItem.IsBreakOnChange = True
						Case 2
							objReportItem.IDColumnName = "Surname"
							objReportItem.IsHidden = False
						Case 3
							objReportItem.IDColumnName = "Forenames"
							objReportItem.IsHidden = False
						Case 4
							objReportItem.IDColumnName = "Department"
							objReportItem.IsHidden = False
						Case 5
							objReportItem.IDColumnName = "Type"
							objReportItem.IsHidden = Not mbDisplayBradfordDetail
						Case 6
							objReportItem.IDColumnName = "Start_Date"
							objReportItem.IsHidden = Not mbDisplayBradfordDetail
						Case 7
							objReportItem.IDColumnName = "Start_Session"
							objReportItem.IsHidden = Not mbDisplayBradfordDetail
						Case 8
							objReportItem.IDColumnName = "End_Date"
							objReportItem.IsHidden = Not mbDisplayBradfordDetail
						Case 9
							objReportItem.IDColumnName = "End_Session"
							objReportItem.IsHidden = Not mbDisplayBradfordDetail
						Case 10
							If mbDisplayBradfordDetail Then
								objReportItem.IDColumnName = "Reason"
								objReportItem.IsHidden = False
							Else
								objReportItem.IDColumnName = "Summary"
								objReportItem.IsHidden = False
							End If
						Case 11
							objReportItem.IDColumnName = "Continuous"
							objReportItem.IsHidden = Not mbDisplayBradfordDetail
						Case 12
							objReportItem.IDColumnName = "Duration"
							objReportItem.IsHidden = False
							objReportItem.IsCount = True
							objReportItem.IsTotal = True
							objReportItem.Decimals = 1
							objReportItem.IsNumeric = True
						Case 13
							objReportItem.IDColumnName = "Included_Days"
							objReportItem.IsHidden = False
							objReportItem.IsCount = True
							objReportItem.IsTotal = True
							objReportItem.Decimals = 1
							objReportItem.IsNumeric = True
							'**********************************************************************
						Case 14
							objReportItem.IDColumnName = "Order_1"
							objReportItem.IsHidden = True

						Case 15
							objReportItem.IDColumnName = "Order_2"
							objReportItem.IsHidden = True
							'**********************************************************************

						Case Else

							If lngTableID = mlngCustomReportsBaseTable Then
								'Personnel
								objReportItem.IDColumnName = mstrSQLFrom & "." & GetColumnName(lngColumnID)
							Else
								'Absence
								objReportItem.IDColumnName = mstrRealSource & "." & GetColumnName(lngColumnID)
							End If

					End Select


					objReportItem.IsPageOnChange = False	'Page break on change
					If mblnCustomReportsSummaryReport Then
						objReportItem.IsValueOnChange = True	'Value on change
					Else
						objReportItem.IsValueOnChange = False	'Value on change
					End If
					objReportItem.SuppressRepeated = IIf(iCount < 5 And mbBradfordSRV, True, False)	'Suppress repeated values
					objReportItem.LastValue = ""
					objReportItem.ID = lngColumnID

					' Set the expression/column type of this column
					objReportItem.Type = "C"
					objReportItem.TableID = lngTableID
					objReportItem.TableName = GetTableName(lngTableID)
					objReportItem.ColumnName = GetColumnName(objReportItem.ID)
					objReportItem.IsDateColumn = IsDateColumn("C", lngTableID, lngColumnID)	'??? - check these out 22/03/01
					objReportItem.IsBitColumn = IsBitColumn("C", lngTableID, lngColumnID)
					objReportItem.Use1000Separator = DoesColumnUseSeparators(lngColumnID)	'Does this column use 1000 separators?

					'Adjust the size of the field if digit separator is used
					If objReportItem.Use1000Separator Then
						objReportItem.Size = objReportItem.Size + Int((objReportItem.Size - objReportItem.Decimals) / 3)
					End If

					' Format for this numeric column
					If objReportItem.IsNumeric Then
						If objReportItem.Use1000Separator Then
							objReportItem.Mask = "{0:#,0." & New String("0", objReportItem.Decimals) & "}"
						Else
							objReportItem.Mask = "{0:#0." & New String("0", objReportItem.Decimals) & "}"
						End If

					End If

					ColumnDetails.Add(objReportItem)

				End If

			Next iCount

			' Get those columns defined as a SortOrder and load into array
			Dim objSortItem As ReportSortItem
			colSortOrder = New List(Of ReportSortItem)()

			'Employee surname
			objSortItem = New ReportSortItem
			objSortItem.ColExprID = aStrRequiredFields(2, 1) ' mstrOrderByColumn
			objSortItem.AscDesc = "Asc"
			colSortOrder.Add(objSortItem)

			'Employee forename
			objSortItem = New ReportSortItem
			objSortItem.ColExprID = aStrRequiredFields(3, 1)	'mstrGroupByColumn
			objSortItem.AscDesc = "Asc"
			colSortOrder.Add(objSortItem)

			'Employee staff number
			objSortItem = New ReportSortItem
			objSortItem.ColExprID = aStrRequiredFields(1, 1)	'mstrGroupByColumn
			objSortItem.AscDesc = "Asc"
			colSortOrder.Add(objSortItem)

			' Force duration and included days to be numeric format in Excel
			iCount = 11 - IIf(lbHideStaffNumber = True, 1, 0)

		Catch ex As Exception
			mstrErrorString = "Error whilst retrieving the details recordsets'." & vbNewLine & ex.Message.RemoveSensitive()
			Return False

		End Try

		Return True

	End Function

	Public Function SetBradfordDisplayOptions(pbSRV As Boolean, pbShowTotals As Boolean, pbShowCount As Boolean, pbShowWorkings As Boolean _
																						, pbShowBasePicklistFilter As Boolean, pbDisplayBradfordDetail As Boolean) As Boolean

		' Set Report Display Options
		mbBradfordSRV = pbSRV
		mbBradfordTotals = pbShowTotals
		mbBradfordCount = pbShowCount
		mbBradfordWorkings = pbShowWorkings
		mblnCustomReportsPrintFilterHeader = pbShowBasePicklistFilter
		mbDisplayBradfordDetail = pbDisplayBradfordDetail

		Return True

	End Function

	Public Function SetBradfordOrders(pstrOrderBy As String, pstrGroupBy As String, pbOrder1Asc As Boolean, pbOrder2Asc As Boolean _
																		, plngOrderByColumnID As Integer, plngGroupByColumnID As Integer) As Boolean

		' Set Report Order Options
		mstrOrderByColumn = pstrOrderBy
		mstrGroupByColumn = pstrGroupBy
		mbOrderBy1Asc = pbOrder1Asc
		mbOrderBy2Asc = pbOrder2Asc
		mlngOrderByColumnID = plngOrderByColumnID
		mlngGroupByColumnID = plngGroupByColumnID

		Return True

	End Function

	Public Function SetBradfordIncludeOptions(pbOmitBeforeStart As Boolean, pbOmitAfterEnd As Boolean, plngPersonnelID As Integer _
																						, plngCustomReportsFilterID As Integer, plngCustomReportsPickListID As Integer, pbMinBradford As Boolean _
																						, plngMinBradfordAmount As Integer) As Boolean

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

	Public Function UDFFunctions(pbCreate As Boolean) As Boolean
		Return General.UDFFunctions(mastrUDFsRequired, pbCreate)
	End Function

	Private Function GetOrderDefinition(plngOrderID As Integer) As DataTable
		' Return a recordset of the order items (both Find Window and Sort Order columns)
		' for the given order.
		Dim prmID As New SqlParameter("piOrderID", SqlDbType.Int)
		prmID.Value = plngOrderID

		Return DB.GetDataTable("sp_ASRGetOrderDefinition", CommandType.StoredProcedure, prmID)

	End Function

	Private Function GetDefaultOrder(plngTableID As Integer) As Integer
		Return Tables.GetById(plngTableID).DefaultOrderID
	End Function


End Class