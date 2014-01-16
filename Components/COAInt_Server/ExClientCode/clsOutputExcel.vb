Option Strict Off
Option Explicit On

Imports Aspose.Cells.Charts
Imports System.Collections.Generic
Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Enums
Imports Aspose.Cells

Namespace ExClientCode

	Friend Class clsOutputExcel
		Inherits BaseOutputFormat

		Private _mxlWorkBook As Workbook
		Private _mxlWorkSheet As Worksheet
		Private _mxlTemplateBook As Workbook
		Private _mxlTemplateSheet As Worksheet
		Private _mxlFirstSheet As Worksheet
		Private _mxlDeleteSheet As Worksheet

		Private _mobjParent As clsOutputRun

		Private _mlngHeaderRows As Integer
		Private _mlngHeaderCols As Integer
		Private _mlngDataCurrentRow As Integer
		Private _mlngDataStartRow As Integer
		Private _mlngDataStartCol As Integer

		Private _mblnScreen As Boolean
		Private _mblnPrinter As Boolean
		Private _mstrPrinterName As String
		Private _mblnSave As Boolean
		Private _mlngSaveExisting As Integer
		Private _mblnEmail As Boolean
		Private _mstrFileName As String
		Private _mblnSizeColumnsIndependently As Boolean
		Private _mblnApplyStyles As Boolean

		Private _mstrSheetMode As String
		Private _mblnAppending As Boolean
		Private _mlngAppendStartRow As Integer

		Private _mstrDefTitle As String
		Private _mstrErrorMessage As String

		Private _mblnChart As Boolean
		Private _mblnPivotTable As Boolean
		'Private mstrIntersectionFormat As String

		Private _mstrXlTemplate As String
		Private _mblnXlExcelGridlines As Boolean
		Private _mblnXlExcelHeaders As Boolean
		Private _mblnXlExcelOmitTopRow As Boolean
		Private _mblnXlExcelOmitLeftCol As Boolean
		Private _mblnXlAutoFitCols As Boolean
		Private _mblnXlLandscape As Boolean
		Const OfficeVersion As Integer = 12

		Public Sub ClearUp()

			'Reset all references to ensure that Excel closes cleanly...			
			'_mxlTemplateSheet = Nothing
			'_mxlTemplateBook = Nothing
			_mxlDeleteSheet = Nothing
			_mxlFirstSheet = Nothing
			_mxlWorkSheet = Nothing
			_mxlWorkBook = Nothing

		End Sub

		Private Sub Class_Initialize_Renamed()

			_mstrXlTemplate = gstrSettingExcelTemplate
			_mblnXlExcelGridlines = gblnSettingExcelGridlines
			_mblnXlExcelHeaders = gblnSettingExcelHeaders
			_mblnXlExcelOmitTopRow = gblnSettingExcelOmitSpacerRow
			_mblnXlExcelOmitLeftCol = gblnSettingExcelOmitSpacerCol
			_mblnXlAutoFitCols = gblnSettingAutoFitCols
			_mblnXlLandscape = gblnSettingLandscape

			_mlngDataStartRow = glngSettingDataRow
			_mlngDataStartCol = glngSettingDataCol

			_mblnSizeColumnsIndependently = False
			_mblnApplyStyles = True

		End Sub

		Public Sub New()
			MyBase.New()
			Class_Initialize_Renamed()
		End Sub

		Public WriteOnly Property Screen() As Boolean
			Set(ByVal value As Boolean)
				_mblnScreen = value
			End Set
		End Property

		Public WriteOnly Property DestPrinter() As Boolean
			Set(ByVal value As Boolean)
				_mblnPrinter = value
			End Set
		End Property

		Public WriteOnly Property PrinterName() As String
			Set(ByVal value As String)
				_mstrPrinterName = value
			End Set
		End Property

		Public WriteOnly Property Save() As Boolean
			Set(ByVal value As Boolean)
				_mblnSave = value
			End Set
		End Property

		Public Property SaveExisting() As Integer
			Get
				SaveExisting = _mlngSaveExisting
			End Get
			Set(ByVal value As Integer)
				_mlngSaveExisting = value
			End Set
		End Property

		Public WriteOnly Property Email() As Boolean
			Set(ByVal value As Boolean)
				_mblnEmail = value
			End Set
		End Property

		Public WriteOnly Property FileName() As String
			Set(ByVal value As String)
				_mstrFileName = value
			End Set
		End Property

		Public WriteOnly Property Chart() As Boolean
			Set(ByVal value As Boolean)
				_mblnChart = value
			End Set
		End Property

		Public WriteOnly Property PivotTable() As Boolean
			Set(ByVal value As Boolean)
				_mblnPivotTable = value
			End Set
		End Property

		Public WriteOnly Property HeaderRows() As Integer
			Set(ByVal value As Integer)
				_mlngHeaderRows = value
			End Set
		End Property

		Public WriteOnly Property HeaderCols() As Integer
			Set(ByVal value As Integer)
				_mlngHeaderCols = value
			End Set
		End Property

		Public WriteOnly Property SizeColumnsIndependently() As Boolean
			Set(ByVal value As Boolean)
				_mblnSizeColumnsIndependently = value
			End Set
		End Property

		Public WriteOnly Property ApplyStyles() As Boolean
			Set(ByVal value As Boolean)
				_mblnApplyStyles = value
			End Set
		End Property

		Public WriteOnly Property Parent() As clsOutputRun
			Set(ByVal value As clsOutputRun)
				_mobjParent = value
			End Set
		End Property

		Public ReadOnly Property ErrorMessage() As String
			Get
				ErrorMessage = _mstrErrorMessage
			End Get
		End Property

		Private Function CreateExcelApplication() As Boolean
			Return True
		End Function

		Public Function GetFile(ByRef objParent As clsOutputRun, ByRef colStyles As Collection) As Boolean

			If Not CreateExcelApplication() Then
				GetFile = False
				Exit Function
			End If

			'Check if file already exists...
			If Dir(_mstrFileName) <> vbNullString And _mstrFileName <> vbNullString Then
				' TODO: We only create new now - no append to file or saveas or owt like that...
				Select Case _mlngSaveExisting
					Case 0 'Overwrite
						If Not objParent.KillFile(_mstrFileName) Then
							GetFile = False
							Exit Function
						End If
						GetWorkBook(strWorkbook:="New", strWorksheet:="New")
					Case 1 'Do not overwrite (fail)
						_mstrErrorMessage = "File already exists."
					Case 2 'Add Sequential number to file
						_mstrFileName = _mobjParent.GetSequentialNumberedFile(_mstrFileName)
						GetWorkBook(strWorkbook:="New", strWorksheet:="New")
					Case 3 'Append to existing file
						GetWorkBook(strWorkbook:="Open", strWorksheet:="Existing")
					Case 4 'Create new worksheet within existing workbook...
						GetWorkBook(strWorkbook:="Open", strWorksheet:="New")
				End Select
			Else
				GetWorkBook(strWorkbook:="New", strWorksheet:="New")
			End If

			GetFile = (_mstrErrorMessage = vbNullString)

		End Function

		Private Sub GetWorkBook(ByRef strWorkbook As String, ByRef strWorksheet As String)

			Dim strFormat As String
			Dim strTempFile As String
			Dim lngCount As Integer
			' Dim lngOriginalFormat As Integer


			If _mblnApplyStyles And _mstrXlTemplate <> "" And Dir(_mstrXlTemplate) <> "" Then
				If Not IsFileCompatibleWithExcelVersion(_mstrXlTemplate, OfficeVersion) Then
					_mstrErrorMessage = "Your User Configuration Output Options are set to use a template file which is not compatible with your version of Microsoft Office."
					Exit Sub
				End If

			End If

			_mstrSheetMode = strWorksheet
			Select Case strWorkbook
				Case "New"
					If _mstrXlTemplate <> vbNullString Then
						_mxlWorkBook = New Workbook(_mstrXlTemplate)
					Else
						_mxlWorkBook = New Workbook()
						' remove ALL worksheets.
						For lngCount = 0 To _mxlWorkBook.Worksheets.Count - 1
							_mxlWorkBook.Worksheets.RemoveAt(lngCount)
						Next

					End If

				Case "Open"
					If Not IsFileCompatibleWithExcelVersion(_mstrFileName, OfficeVersion) Then
						_mstrErrorMessage = "This definition is set to append to a file which is not compatible with your version of Microsoft Office."
						Exit Sub
					End If

					_mxlWorkBook = New Workbook(_mstrFileName)

			End Select

		End Sub


		Private Sub GetWorksheet(ByRef strSheetName As String)

			Dim blnFound As Boolean

			_mblnAppending = False
			_mlngAppendStartRow = 0

			'If we are appending, then see if there is an existing worksheet with this name...
			blnFound = False
			If _mstrSheetMode = "Existing" Then
				For Each workSheet As Worksheet In _mxlWorkBook.Worksheets
					If Trim(workSheet.Name) = FormatSheetName(strSheetName) Then
						_mxlWorkBook.Worksheets.ActiveSheetIndex = workSheet.Index
						blnFound = True
						Exit For
					End If
				Next workSheet
			End If

			If blnFound Then
				StartAtBottomOfSheet()
				_mblnAppending = True
			Else
				If _mstrXlTemplate <> vbNullString Then
					'_mxlWorkBook.Worksheets(_mxlWorkBook.Worksheets.Count).Copy(_mxlTemplateSheet)
					'_mxlWorkSheet = _mxlWorkBook.Worksheets(_mxlWorkBook.Worksheets.Count + 1)
					_mxlWorkSheet = _mxlWorkBook.Worksheets(0)
					StartAtBottomOfSheet()
				Else
					_mxlWorkSheet = _mxlWorkBook.Worksheets(_mxlWorkBook.Worksheets.Add())
				End If
				SetSheetName(_mxlWorkSheet, strSheetName)
			End If

			If Not (_mxlDeleteSheet Is Nothing) Then
				'mxlDeleteSheet.Delete()
				_mxlDeleteSheet = Nothing
			End If

		End Sub

		Public Sub AddPage(ByRef strDefTitle As String, ByRef strSheetName As String, ByRef colStyles As Collection)

			_mstrDefTitle = strDefTitle

			If _mblnPivotTable Then
				GetWorksheet("Data " & strSheetName)
			Else
				GetWorksheet(strSheetName)
			End If

			If Not _mblnChart And Not _mblnPivotTable Then
				If _mxlFirstSheet Is Nothing Then
					_mxlFirstSheet = _mxlWorkSheet
				End If
			End If

			If _mlngAppendStartRow = 0 Then
				_mlngDataCurrentRow = _mlngDataStartRow
			End If

			If _mblnApplyStyles = False Then
				If Not _mblnAppending Then
					_mlngDataCurrentRow = 1
				End If
				_mlngDataStartCol = 1
				_mlngHeaderCols = 0
				_mlngHeaderRows = 0
			End If

		End Sub


		Public Sub DataArray(ByRef strArray(,) As String, ByRef colColumns As List(Of Metadata.Column), ByRef colStyles As Collection, ByRef colMerges As Collection)

			Dim lngGridCol As Integer
			Dim lngGridRow As Integer
			Dim lngExcelCol As Integer
			Dim lngExcelRow As Integer

			If _mstrErrorMessage <> vbNullString Then
				Exit Sub
			End If

			If UBound(strArray, 1) > 255 Then
				_mstrErrorMessage = "Maximum of 255 columns exceeded"
				Exit Sub
			End If

			'Instantiate the error checking options
			Dim opts As ErrorCheckOptionCollection = _mxlWorkSheet.ErrorCheckOptions
			Dim index As Integer = opts.Add()
			Dim opt As ErrorCheckOption = opts(index)
			'Disable the numbers stored as text option
			opt.SetErrorCheck(ErrorCheckType.TextNumber, False)

			' PrepareRows sets the datatype for each column. However, they're overwritten at present - need to rethink.
			PrepareRows(UBound(strArray, 2), colColumns, colStyles)

			lngExcelCol = _mlngDataStartCol
			lngExcelRow = _mlngDataCurrentRow

			For lngGridRow = 0 To UBound(strArray, 2)
				For lngGridCol = 0 To UBound(strArray, 1)

					With _mxlWorkSheet.Cells(lngExcelRow + lngGridRow - 1, lngExcelCol + lngGridCol - 1)

						Dim stlNumeric As Style = _mxlWorkBook.Styles(_mxlWorkBook.Styles.Add())
						Dim stlGeneral As Style = _mxlWorkBook.Styles(_mxlWorkBook.Styles.Add())
						Dim stlDate As Style = _mxlWorkBook.Styles(_mxlWorkBook.Styles.Add())
						Dim flag As StyleFlag = New StyleFlag()

						stlNumeric.Number = 4
						stlGeneral.Number = 49
						stlDate.Number = 14

						Select Case colColumns.Item(lngGridCol).DataType

							Case SQLDataType.sqlNumeric, SQLDataType.sqlInteger
								' .NumberFormat = IIf(objColumn.ThousandSeparator, "#,##0", "0") & IIf(objColumn.DecPlaces, "." & New String("0", objColumn.DecPlaces), "")
								.SetStyle(stlNumeric)
								If lngGridRow = 0 Then
									.PutValue(strArray(lngGridCol, lngGridRow))
								Else
									.PutValue(CLng(NullSafeInteger(strArray(lngGridCol, lngGridRow))))
								End If
							Case SQLDataType.sqlBoolean
								.SetStyle(stlGeneral)
								.PutValue(strArray(lngGridCol, lngGridRow))
							Case SQLDataType.sqlUnknown
								'Leave it alone! (Required for percentages on Standard Reports)
								.SetStyle(stlGeneral)
								.PutValue(strArray(lngGridCol, lngGridRow))
							Case SQLDataType.sqlDate
								.SetStyle(stlDate)
								'MH20050104 Fault 9695 & 9696
								'Adding ;@ to the end formats it as "short date" so excel will look at the
								'regional settings when opening the workbook rather than force it to always
								'be in the format of the user who created the workbook.
								.PutValue(strArray(lngGridCol, lngGridRow))
							Case Else
								.SetStyle(stlGeneral)
								.PutValue(strArray(lngGridCol, lngGridRow))
						End Select


						'MH20031113 Fault 7602
						' .Value = IIf(Left(strArray(lngGridCol, lngGridRow), 1) = "'", "'", vbNullString) & strArray(lngGridCol, lngGridRow)
						'If lngGridRow < mlngHeaderRows Then
						'  .HorizontalAlignment = xlCenter
						'End If
					End With
				Next
			Next

			If _mblnChart Then
				ApplyStyle(UBound(strArray, 1), UBound(strArray, 2), colStyles)
				ApplyCellOptions(UBound(strArray, 1), colStyles, True)

				CreateChart(_mlngDataCurrentRow + UBound(strArray, 2), _mlngDataStartCol + UBound(strArray, 1), colStyles)
				ApplyCellOptions(UBound(strArray, 1), colStyles, False)

				'Delete superfluous rows and cols if setup in User Config reports section
				If _mblnXlExcelOmitLeftCol Then _mxlWorkSheet.Cells.DeleteColumn(0)
				If _mblnXlExcelOmitTopRow Then _mxlWorkSheet.Cells.DeleteRows(0, 1)

			ElseIf _mblnPivotTable Then

				If UBound(strArray, 1) < 1 Then
					_mstrErrorMessage = "Unable to create a pivot table for a single column of data."
				Else
					ApplyStyle(UBound(strArray, 1), UBound(strArray, 2), colStyles)

					' CreatePivotTable(_mlngDataCurrentRow + UBound(strArray, 2), _mlngDataStartCol + UBound(strArray, 1), strArray(0, 0), strArray(1, 0), strArray(UBound(strArray), 0), colStyles, colColumns)
				End If

			Else
				If _mblnApplyStyles Then
					ApplyStyle(UBound(strArray, 1), UBound(strArray, 2), colStyles)
					ApplyMerges(colMerges)
				End If
				ApplyCellOptions(UBound(strArray, 1), colStyles, True)

				'Delete superfluous rows and cols if setup in User Config reports section
				If _mblnXlExcelOmitLeftCol Then _mxlWorkSheet.Cells.DeleteColumn(0)
				If _mblnXlExcelOmitTopRow Then _mxlWorkSheet.Cells.DeleteRows(0, 1)

			End If

			_mlngDataCurrentRow = _mlngDataCurrentRow + UBound(strArray, 2) + IIf(_mblnApplyStyles, 2, 1)

		End Sub

		Private Sub CreateChart(ByRef lngMaxRows As Integer, ByRef lngMaxCols As Integer, ByRef colStyles As Collection)



			Dim strSheetName As String

			On Error GoTo LocalErr

			strSheetName = _mxlWorkSheet.Name & " Chart"
			Dim mxlChartWorkSheet = _mxlWorkBook.Worksheets(_mxlWorkBook.Worksheets.Add())
			mxlChartWorkSheet.Name = strSheetName
			mxlChartWorkSheet.MoveTo(0)

			Dim chartIndex As Integer = mxlChartWorkSheet.Charts.Add(Charts.ChartType.Bar, 0, 0, 30, 20)
			Dim xlChart As Charts.Chart = mxlChartWorkSheet.Charts(chartIndex)	'	 Microsoft.Office.Interop.Excel.Chart
			Dim xlData As Range		' Microsoft.Office.Interop.Excel.Range
			Dim xlCategories As Range

			Dim dataFirstRow As Integer = _mlngDataCurrentRow
			Dim dataFirstCol As Integer = _mlngDataStartCol - 1
			Dim dataRowCount As Integer = lngMaxRows - _mlngDataCurrentRow
			Dim dataColumnCount As Integer = lngMaxCols - _mlngDataStartCol

			' xlData = mxlWorkSheet.Range(mxlWorkSheet.Cells._Default(mlngDataCurrentRow, mlngDataStartCol), mxlWorkSheet.Cells._Default(lngMaxRows, lngMaxCols))
			xlCategories = _mxlWorkSheet.Cells.CreateRange(dataFirstRow, dataFirstCol, dataRowCount, 1)
			xlCategories.Name = "xlCategories"
			xlData = _mxlWorkSheet.Cells.CreateRange(dataFirstRow, dataFirstCol + dataColumnCount, dataRowCount, 1)
			xlData.Name = "xlData"


			' xlChart = mxlApp.Charts.Add(After:=mxlWorkSheet)

			With xlChart
				'.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xl3DColumnClustered
				.Type = ChartType.Column3DClustered
				' .SetSourceData(Source:=xlData, PlotBy:=Microsoft.Office.Interop.Excel.XlRowCol.xlColumns)
				.NSeries.Add("=xlData", True)
				.NSeries.CategoryData = "=xlCategories"
				' .Location(Where:=Microsoft.Office.Interop.Excel.XlChartLocation.xlLocationAsNewSheet)

				'.ChartTitle.Caption = mstrDefTitle
				.Title.Text = _mstrDefTitle
				'.HasTitle = True
				.Title.IsVisible = True
				'MH20061204 Fault 11230
				'.ChartTitle.Characters.Text = mstrDefTitle
				'.ChartTitle.Text = mstrDefTitle
				'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().Bold. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'.ChartTitle.Font.Bold = colStyles.Item("Title").Bold
				.Title.Font.IsBold = True
				'.ChartTitle.Font.Size = 12
				.Title.Font.Size = 12
				'MH20050113 Fault 9376
				'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().Underline. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'.ChartTitle.Font.Underline = colStyles.Item("Title").Underline
				.Title.Font.Underline = colStyles.Item("Title").Underline
				'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().ForeCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'.ChartTitle.Font.Color = colStyles.Item("Title").ForeCol
				.Title.Font.Color = ColorTranslator.FromWin32(colStyles.Item("Title").ForeCol)
			End With


			' SetSheetName((mxlWorkBook.ActiveChart), strSheetName)


			' SetSheetName(_mxlWorkSheet, strSheetName)
			'If mxlFirstSheet Is Nothing Then
			'	mxlFirstSheet = mxlWorkBook.ActiveChart
			'End If


			Exit Sub

LocalErr:
			_mstrErrorMessage = Err.Description

		End Sub


		Private Sub CreatePivotTable(ByRef lngMaxRows As Integer, ByRef lngMaxCols As Integer, ByRef strHor As String, ByRef strVer As String, ByRef strInt As String, ByRef colStyles As Collection, ByRef colColumns As Collection)

			''Adding a new sheet
			'Dim sheet2 As Worksheet = _mxlWorkBook.Worksheets(_mxlWorkBook.Worksheets.Add())
			''Naming the sheet
			'sheet2.Name = "PivotTable"
			''Getting the pivottables collection in the sheet
			'Dim pivotTables As Aspose.Cells.Pivot.PivotTableCollection = sheet2.PivotTables
			''Adding a PivotTable to the worksheet
			'Dim index As Integer = pivotTables.Add("=Data!B4:J206", "B3", "PivotTable1")
			''Accessing the instance of the newly added PivotTable
			'Dim mxlpivotTable As Aspose.Cells.Pivot.PivotTable = pivotTables(index)
			''Showing the grand totals
			'mxlpivotTable.RowGrand = True
			'mxlpivotTable.ColumnGrand = True
			''Setting the PivotTable report is automatically formatted
			'mxlpivotTable.IsAutoFormat = True
			''Setting the PivotTable autoformat type.
			'mxlpivotTable.AutoFormatType = Aspose.Cells.Pivot.PivotTableAutoFormatType.Report6
			''Draging the first field to the row area.
			'mxlpivotTable.AddFieldToArea(Aspose.Cells.Pivot.PivotFieldType.Row, 0)
			''Draging the third field to the row area.
			'mxlpivotTable.AddFieldToArea(Aspose.Cells.Pivot.PivotFieldType.Row, 2)
			''Draging the second field to the row area.
			'mxlpivotTable.AddFieldToArea(Aspose.Cells.Pivot.PivotFieldType.Row, 1)
			''Draging the fourth field to the column area.
			'mxlpivotTable.AddFieldToArea(Aspose.Cells.Pivot.PivotFieldType.Column, 3)
			''Draging the fifth field to the data area.
			'mxlpivotTable.AddFieldToArea(Aspose.Cells.Pivot.PivotFieldType.Data, 5)
			''Setting the number format of the first data field
			'mxlpivotTable.DataFields(0).NumberFormat = "$#,##0.00"
			''Saving the Excel file
			'' Workbook.Save("f:\test\pivotTable_test.xls")


			'		Dim xlPivot As Microsoft.Office.Interop.Excel.PivotTable
			'		Dim xlDataSheet As Microsoft.Office.Interop.Excel.Worksheet
			'		Dim xlData As Microsoft.Office.Interop.Excel.Range
			'		Dim xlStart As Microsoft.Office.Interop.Excel.Range
			'		Dim objColumn As clsColumn
			'		Dim strSheetName As String
			'		Dim xlFunc As Microsoft.Office.Interop.Excel.XlConsolidationFunction

			'		On Error GoTo LocalErr

			'		mxlApp.DisplayAlerts = True

			'		xlData = mxlWorkSheet.Range(mxlWorkSheet.Cells._Default(mlngDataCurrentRow, mlngDataStartCol), mxlWorkSheet.Cells._Default(lngMaxRows, lngMaxCols))
			'		strSheetName = Mid(mxlWorkSheet.Name, 6)
			'		'SetSheetName mxlWorkSheet, "Data " & mxlWorkSheet.Name
			'		xlDataSheet = mxlWorkSheet

			'		GetWorksheet(strSheetName)
			'		If mxlFirstSheet Is Nothing Then
			'			mxlFirstSheet = mxlWorkSheet
			'		End If

			'		xlDataSheet.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden
			'		xlStart = mxlWorkSheet.Cells._Default(mlngDataCurrentRow, mlngDataStartCol)

			'		mxlApp.DisplayAlerts = False


			'		'MH20100628
			'		If OfficeVersion > 12 Then
			'			xlPivot = mxlWorkBook.PivotCaches.Create(SourceType:=Microsoft.Office.Interop.Excel.XlPivotTableSourceType.xlDatabase, SourceData:=xlData).CreatePivotTable(TableDestination:=xlStart)
			'		Else
			'			xlPivot = mxlWorkSheet.PivotTableWizard(SourceType:=Microsoft.Office.Interop.Excel.XlPivotTableSourceType.xlDatabase, SourceData:=xlData, TableDestination:=xlStart)
			'		End If


			'		With xlPivot
			'			.AddFields(RowFields:=strVer, ColumnFields:=strHor)

			'			'AE20071017 Fault #12540
			'			Select Case mobjParent.PivotDataFunction
			'				Case "Count"
			'					xlFunc = Microsoft.Office.Interop.Excel.XlConsolidationFunction.xlCount
			'				Case "Average"
			'					xlFunc = Microsoft.Office.Interop.Excel.XlConsolidationFunction.xlAverage
			'				Case "Maximum"
			'					xlFunc = Microsoft.Office.Interop.Excel.XlConsolidationFunction.xlMax
			'				Case "Minimum"
			'					xlFunc = Microsoft.Office.Interop.Excel.XlConsolidationFunction.xlMin
			'				Case "Total"
			'					xlFunc = Microsoft.Office.Interop.Excel.XlConsolidationFunction.xlSum
			'			End Select

			'			'.PivotFields(strInt).Orientation = xlDataField

			'			With .PivotFields(strInt)
			'				'UPGRADE_WARNING: Couldn't resolve default property of object xlPivot.PivotFields().Orientation. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'				.Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlDataField
			'				'UPGRADE_WARNING: Couldn't resolve default property of object xlPivot.PivotFields().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'				.Name = mobjParent.PivotDataFunction & " of " & strInt
			'				'UPGRADE_WARNING: Couldn't resolve default property of object xlPivot.PivotFields().Function. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'				.Function = xlFunc
			'			End With

			'			.NullString = IIf(mobjParent.PivotSuppressBlanks, "", "0")

			'			objColumn = colColumns.Item(colColumns.Count())
			'			If objColumn.DecPlaces > 0 Then
			'				If objColumn.DecPlaces > 100 Then objColumn.DecPlaces = 100
			'				.DataBodyRange.NumberFormat = IIf(objColumn.ThousandSeparator, "#,##0", "0") & IIf(objColumn.DecPlaces, "." & New String("0", objColumn.DecPlaces), "")
			'			End If
			'			'UPGRADE_NOTE: Object objColumn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			'			objColumn = Nothing

			'			mxlApp.DisplayAlerts = True
			'		End With

			'		'UPGRADE_WARNING: Couldn't resolve default property of object colStyles(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'		ApplyStyleToRange(xlStart, colStyles.Item("Heading"))
			'		'UPGRADE_WARNING: Couldn't resolve default property of object colStyles(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'		ApplyStyleToRange(xlPivot.RowRange, colStyles.Item("Heading"))
			'		'UPGRADE_WARNING: Couldn't resolve default property of object colStyles(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'		ApplyStyleToRange(xlPivot.ColumnRange, colStyles.Item("Heading"))
			'		'UPGRADE_WARNING: Couldn't resolve default property of object colStyles(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'		ApplyStyleToRange(xlPivot.DataBodyRange, colStyles.Item("Data"))
			'		mxlWorkSheet.Range("A1").Select()
			'		mlngHeaderCols = 1

			'		ApplyCellOptions(xlPivot.ColumnRange.Columns.Count, colStyles, True)

			'		'Delete superfluous rows and cols if setup in User Config reports section
			'		If mblnXLExcelOmitLeftCol Then mxlWorkSheet.Range("A:A").Delete()
			'		If mblnXLExcelOmitTopRow Then mxlWorkSheet.Range("1:1").Delete()

			'		'UPGRADE_NOTE: Object xlPivot may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			'		xlPivot = Nothing
			'		mxlApp.DisplayAlerts = False

			'		Exit Sub

			'LocalErr:
			'		mstrErrorMessage = Err.Description

		End Sub


		Private Sub PrepareRows(lngRowCount As Integer, ByRef colColumns As List(Of Metadata.Column), ByRef colStyles As Collection)

			Dim objColumn As Metadata.Column
			Dim lngCount As Integer = 0

			On Error GoTo LocalErr

			With _mxlWorkSheet

				If _mlngHeaderRows > 0 Then
					' Define style for header row and apply it.
					Dim stlHeaderRow As Style = _mxlWorkBook.Styles(_mxlWorkBook.Styles.Add())
					stlHeaderRow.Custom = "@"
					Dim range As Range = _mxlWorkSheet.Cells.CreateRange(_mlngDataCurrentRow - 1, _mlngDataStartCol - 1, _mlngHeaderRows, colColumns.Count())
					'Name the range.
					range.Name = "HeaderRange"
					range.SetStyle(stlHeaderRow)
				End If

				Dim stlColumnStyleTmp As Style = _mxlWorkBook.Styles(_mxlWorkBook.Styles.Add())

				For Each objColumn In colColumns

					Dim columnRange As Range = _mxlWorkSheet.Cells.CreateRange(_mlngDataCurrentRow + _mlngHeaderRows - 1, _mlngDataStartCol - 1 + lngCount, _mlngDataCurrentRow + lngRowCount, 1)
					Select Case objColumn.DataType
						Case SQLDataType.sqlNumeric, SQLDataType.sqlInteger
							If objColumn.Decimals > 0 Then
								If objColumn.Decimals > 100 Then objColumn.Decimals = 100
								' .NumberFormat = IIf(objColumn.ThousandSeparator, "#,##0", "0") & IIf(objColumn.DecPlaces, "." & New String("0", objColumn.DecPlaces), "")
								stlColumnStyleTmp.Custom = IIf(objColumn.Use1000Separator, "#,##0", "0") & IIf(objColumn.Decimals, "." & New String("0", objColumn.Decimals), "")
								' 								stlColumnStyleTmp.Number = 4
							Else
								stlColumnStyleTmp.Custom = "@"
							End If
						Case SQLDataType.sqlBoolean
							' .HorizontalAlignment = TextAlignmentType.Center
							stlColumnStyleTmp.Custom = "@"
						Case SQLDataType.sqlUnknown
							'Leave it alone! (Required for percentages on Standard Reports)
							stlColumnStyleTmp.Custom = "@"
						Case SQLDataType.sqlDate
							'MH20050104 Fault 9695 & 9696
							'Adding ;@ to the end formats it as "short date" so excel will look at the
							'regional settings when opening the workbook rather than force it to always
							'be in the format of the user who created the workbook.
							stlColumnStyleTmp.Custom = DateFormat() & ";@"
						Case Else
							stlColumnStyleTmp.Custom = "@"
					End Select

					' Apply style.
					columnRange.SetStyle(stlColumnStyleTmp)

				Next

			End With

			Exit Sub

LocalErr:
			_mstrErrorMessage = Err.Description

		End Sub


		Private Sub ApplyCellOptions(ByRef lngColCount As Integer, ByRef colStyles As Collection, ByRef blnGridLines As Boolean)

			Dim objRange As Range	' Microsoft.Office.Interop.Excel.Range
			Dim lngCount As Integer

			Dim lngMaxWidth As Double
			Dim lngTitleColWidth As Double
			Dim lngTitleSize As Double

			On Error GoTo LocalErr

			If _mblnXlAutoFitCols Then
				' mxlWorkSheet.Range(mxlWorkSheet.Cells._Default(mlngDataCurrentRow, mlngDataStartCol), mxlWorkSheet.Cells._Default(mlngDataCurrentRow, mlngDataStartCol + lngColCount)).EntireColumn.AutoFit()
				_mxlWorkSheet.AutoFitColumns()

				' TODO: NOt required?
				'If Not mblnSizeColumnsIndependently Then
				'	lngMaxWidth = 0
				'	For lngCount = mlngDataStartCol + mlngHeaderCols To mlngDataStartCol + lngColCount
				'		'UPGRADE_WARNING: Couldn't resolve default property of object mxlWorkSheet.Columns(lngCount).ColumnWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'		If lngMaxWidth < mxlWorkSheet.Columns._Default(lngCount).ColumnWidth Then
				'			'UPGRADE_WARNING: Couldn't resolve default property of object mxlWorkSheet.Columns().ColumnWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'			lngMaxWidth = mxlWorkSheet.Columns._Default(lngCount).ColumnWidth
				'		End If
				'	Next

				'	For lngCount = mlngDataStartCol + mlngHeaderCols To mlngDataStartCol + lngColCount
				'		'UPGRADE_WARNING: Couldn't resolve default property of object mxlWorkSheet.Columns().ColumnWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'		mxlWorkSheet.Columns._Default(lngCount).ColumnWidth = lngMaxWidth
				'	Next
				'End If

			End If


			If _mblnApplyStyles Then
				If blnGridLines Then
					'With mxlApp.ActiveWindow
					'.DisplayGridlines = mblnXLExcelGridlines
					_mxlWorkSheet.IsGridlinesVisible = _mblnXlExcelGridlines
					' .DisplayHeadings = mblnXLExcelHeaders
					_mxlWorkSheet.IsRowColumnHeadersVisible = _mblnXlExcelHeaders
					'End With
				End If

				With colStyles.Item("Title")
					.StartCol = glngSettingTitleCol
					.StartRow = IIf(_mlngAppendStartRow > 0, _mlngAppendStartRow, glngSettingTitleRow)
					.EndCol = .StartCol
					.EndRow = .StartRow
				End With

				'Put title in after autofit...
				If colStyles.Item("Title").StartCol <> 0 And colStyles.Item("Title").StartRow <> 0 Then
					_mxlWorkSheet.Cells(colStyles.Item("Title").StartRow - 1, colStyles.Item("Title").StartCol - 1).Value = _mstrDefTitle
					objRange = _mxlWorkSheet.Cells.CreateRange(colStyles.Item("Title").StartRow - 1, colStyles.Item("Title").StartCol - 1, 1, 1)
					ApplyStyleToRange(objRange, colStyles.Item("Title"))
					' objRange = mxlWorkSheet.Cells._Default(colStyles.Item("Title").StartRow, colStyles.Item("Title").StartCol)
					' ApplyStyleToRange(objRange, colStyles.Item("Title"))

					'MH20020807 Fault 6562
					'Merge cells for the title column so that if you append to the file
					'then the title is not taken into account during column sizing.
					' TODO: 
					'With mxlWorkSheet.Columns._Default(colStyles.Item("Title").StartCol)
					'	lngTitleColWidth = .ColumnWidth
					'	lngMaxWidth = .Width
					'	.AutoFit()
					'	lngTitleSize = .Width

					'	lngCount = colStyles.Item("Title").StartCol
					'	Do
					'		lngCount = lngCount + 1
					'		lngMaxWidth = lngMaxWidth + mxlWorkSheet.Columns._Default(lngCount).Width
					'	Loop While lngMaxWidth < lngTitleSize

					'	With mxlWorkSheet
					'		.Range(.Cells._Default(objRange.Row, objRange.Column), .Cells._Default(objRange.Row, lngCount)).Merge()
					'	End With
					'	.ColumnWidth = lngTitleColWidth
					'End With
					'NHRD09072012 Jira HRPRO-2308
					'      'Delete superfluous rows and cols if setup in User Config reports section
					'      If mblnXLExcelOmitLeftCol Then mxlWorkSheet.Range("A:A").Delete
					'      If mblnXLExcelOmitTopRow Then mxlWorkSheet.Range("1:1").Delete
				End If
			End If

			Exit Sub

LocalErr:
			_mstrErrorMessage = Err.Description

		End Sub


		Private Sub ApplyStyle(ByRef lngNumCols As Integer, ByRef lngNumRows As Integer, ByRef colStyles As Collection)

			Dim objStyle As clsOutputStyle
			Dim objRange As Range
			Dim lngCol As Integer
			Dim lngRow As Integer

			On Error GoTo LocalErr

			lngCol = _mlngDataStartCol
			lngRow = _mlngDataCurrentRow

			With colStyles.Item("Title")
				.StartCol = glngSettingTitleCol
				.StartRow = IIf(_mlngAppendStartRow > 0, _mlngAppendStartRow, glngSettingTitleRow)
				.EndCol = .StartCol
				.EndRow = .StartRow
			End With

			With colStyles.Item("Heading")
				.StartCol = 0
				.StartRow = 0
				.EndCol = lngNumCols
				.EndRow = _mlngHeaderRows - 1
			End With

			If _mlngHeaderCols > 0 Then
				With colStyles.Item("HeadingCols")
					.StartCol = 0
					.StartRow = 0
					.EndCol = _mlngHeaderCols - 1
					.EndRow = lngNumRows
				End With
			End If

			With colStyles.Item("Data")
				.StartCol = _mlngHeaderCols
				.StartRow = _mlngHeaderRows
				.EndCol = lngNumCols
				.EndRow = lngNumRows
			End With

			For Each objStyle In colStyles
				If objStyle.Name <> "Title" Then
					If objStyle.StartRow + lngRow > 0 And objStyle.StartCol + lngCol > 0 Then
						Dim totalRows = (objStyle.EndRow + lngRow) - (objStyle.StartRow + lngRow - 1)
						Dim totalCols = (objStyle.EndCol + lngCol) - (objStyle.StartCol + lngCol - 1)

						objRange = _mxlWorkSheet.Cells.CreateRange(objStyle.StartRow + lngRow - 1, objStyle.StartCol + lngCol - 1, totalRows, totalCols)
						ApplyStyleToRange(objRange, objStyle)
					End If
				End If
			Next objStyle

			Exit Sub

LocalErr:
			_mstrErrorMessage = Err.Description

		End Sub


		Private Sub ApplyMerges(ByRef colMerges As Collection)

			Dim objMerge As clsOutputStyle
			Dim objRange As Range
			Dim lngCol As Integer
			Dim lngRow As Integer

			On Error GoTo LocalErr

			lngCol = _mlngDataStartCol
			lngRow = _mlngDataCurrentRow

			For Each objMerge In colMerges
				If objMerge.StartRow + lngRow > 0 And objMerge.StartCol + lngCol > 0 Then
					Dim totalRows = (objMerge.EndRow + lngRow) - objMerge.StartRow
					Dim totalCols = (objMerge.EndCol + lngCol) - objMerge.StartCol
					objRange = _mxlWorkSheet.Cells.CreateRange(objMerge.StartRow + lngRow - 1, objMerge.StartCol + lngCol - 1, totalRows, totalCols)
					' objRange.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlTop
					objRange.Merge()
				End If
			Next objMerge

			Exit Sub

LocalErr:
			_mstrErrorMessage = Err.Description

		End Sub


		Private Sub ApplyStyleToRange(ByRef objRange As Range, ByRef objStyle As clsOutputStyle)

			On Error GoTo LocalErr

			Dim rangeStyle As Style = _mxlWorkBook.Styles(_mxlWorkBook.Styles.Add())
			rangeStyle.Name = objStyle.Name

			With objRange

				If objStyle.CenterText Then
					rangeStyle.HorizontalAlignment = TextAlignmentType.Center
				Else
					rangeStyle.HorizontalAlignment = TextAlignmentType.Left
				End If

				rangeStyle.Font.IsBold = objStyle.Bold
				rangeStyle.Font.Underline = objStyle.Underline
				rangeStyle.Font.Color = ColorTranslator.FromWin32(objStyle.ForeCol)

				'Don't do the backcol nor gridlines for the title...
				If objStyle.Name <> "Title" Then
					rangeStyle.ForegroundColor = ColorTranslator.FromWin32(objStyle.BackCol)
					rangeStyle.Pattern = BackgroundType.Solid

					On Error Resume Next

					If objStyle.Gridlines Then
						.SetOutlineBorders(CellBorderType.Thin, Color.Black)
						rangeStyle.Borders(BorderType.LeftBorder).Color = Color.Black
						rangeStyle.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.Thin
						rangeStyle.Borders(BorderType.RightBorder).Color = Color.Black
						rangeStyle.Borders(BorderType.RightBorder).LineStyle = CellBorderType.Thin
						rangeStyle.Borders(BorderType.TopBorder).Color = Color.Black
						rangeStyle.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Thin
						rangeStyle.Borders(BorderType.BottomBorder).Color = Color.Black
						rangeStyle.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Thin
					Else
						.SetOutlineBorders(CellBorderType.None, Color.Transparent)
					End If

				End If

				objRange.SetStyle(rangeStyle)

			End With

			Exit Sub

LocalErr:
			_mstrErrorMessage = Err.Description

		End Sub


		Public Sub Complete()

			'TODO: Dim objChart As Microsoft.Office.Interop.Excel.Chart
			Dim objWorksheet As Worksheet		' Microsoft.Office.Interop.Excel.Worksheet
			Dim strFormat As String
			Dim strTempFile As String
			Dim strExtension As String
			Dim aryFileBits() As String

			On Error GoTo LocalErr

			If _mstrErrorMessage <> vbNullString Then
				Exit Sub
			End If

			_mxlWorkBook.Worksheets.ActiveSheetIndex = 0

			'SAVE
			'If _mblnSave Then
			' Always Save - we need the temporary file. 
			_mstrErrorMessage = "Error saving file <" & _mstrFileName & ">"

			' calculate the appropriate output type
			'	aryFileBits = Split(_mstrFileName, ".")
			'	strExtension = aryFileBits(UBound(aryFileBits))

			_mxlWorkBook.Save(_mstrFileName, SaveAsFormat(DownloadExtension))

			'Select Case UCase(strExtension)
			'	Case "XLSX"
			'		_mxlWorkBook.Save(_mstrFileName, SaveFormat.Xlsx)
			'	Case "XLS"
			'		_mxlWorkBook.Save(_mstrFileName, SaveFormat.Excel97To2003)
			'	Case "HTML"
			'		_mxlWorkBook.Save(_mstrFileName, SaveFormat.Html)
			'	Case "PDF"
			'		_mxlWorkBook.Save(_mstrFileName, SaveFormat.Pdf)
			'	Case "CSV"
			'		_mxlWorkBook.Save(_mstrFileName, SaveFormat.CSV)
			'End Select
			'End If

			'EMAIL
			If _mblnEmail Then
				_mstrErrorMessage = "Error sending email"
				_mobjParent.SendEmail(_mstrFileName)
			End If

			'PRINTER
			Dim strCurrentPrinter As String
			If _mblnPrinter Then
				_mstrErrorMessage = "Error printing"

				If _mblnChart Then
					' TODO: Charts
					'For Each objChart In mxlWorkBook.Charts
					'	objChart.PrintOut(, , , , mstrPrinterName)
					'Next objChart
				Else
					' TODO: 
					'For Each objWorksheet In mxlWorkBook.Worksheets
					'	If objWorksheet.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible Then
					'		objWorksheet.PrintOut(, , , , mstrPrinterName)
					'	End If
					'Next objWorksheet
				End If

				'mobjParent.ResetDefaultPrinter
			End If


			'SCREEN
			' TODO: Stream it out!
			'If mblnScreen Then
			'	mstrErrorMessage = "Error displaying Excel"
			'	mxlApp.DisplayAlerts = True
			'	mxlApp.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMaximized
			'	mxlApp.Visible = True
			'	mxlWorkBook.Activate()
			'	mxlWorkBook = Nothing
			'	mxlApp = Nothing 'Stops Excel quitting...
			'Else
			'	mxlWorkBook.Saved = True
			'	mxlWorkBook.Close()
			'	mxlApp.Quit()
			'End If

			_mstrErrorMessage = vbNullString

TidyAndExit:
			ClearUp()

			Exit Sub

LocalErr:
			_mstrErrorMessage = _mstrErrorMessage & IIf(Err.Description <> vbNullString, vbCrLf & " (" & Err.Description & ")", vbNullString).ToString()
			Resume TidyAndExit

		End Sub

		Private Sub Class_Terminate_Renamed()
			_mxlFirstSheet = Nothing
			' NOTE: No templates in intranet
			'_mxlTemplateSheet = Nothing
			'_mxlTemplateBook = Nothing
			_mxlWorkSheet = Nothing
			_mxlWorkBook = Nothing
		End Sub
		Protected Overrides Sub Finalize()
			Class_Terminate_Renamed()
			MyBase.Finalize()
		End Sub

		Private Sub StartAtBottomOfSheet()

			'Start at the bottom of the sheet
			' mlngAppendStartRow = mxlWorkSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell).Row + IIf(mblnApplyStyles, 2, 1)
			_mlngAppendStartRow = _mxlWorkSheet.Cells.MaxDataRow + IIf(_mblnApplyStyles, 2, 1)
			_mlngDataCurrentRow = _mlngAppendStartRow + IIf(_mblnApplyStyles, 2, 1)

		End Sub


		Private Function FormatSheetName(ByRef strSheetName As String) As String

			Dim strInvalidChars As String
			Dim lngCount As Integer

			If Left(strSheetName, 1) = "'" Then
				strSheetName = " " & strSheetName
			End If

			strInvalidChars = "\/*:[]?,"
			For lngCount = 1 To Len(strInvalidChars)
				strSheetName = Replace(strSheetName, Mid(strInvalidChars, lngCount, 1), " ")
			Next

			Do While InStr(strSheetName, "  ") > 0
				strSheetName = Replace(strSheetName, "  ", " ")
			Loop

			FormatSheetName = Left(Trim(strSheetName), 31)

		End Function


		Private Function SetSheetName(ByRef objObject As Worksheet, ByVal strSheetName As String) As Boolean

			Dim strNumber As String
			Dim lngCount As Integer

			strSheetName = FormatSheetName(strSheetName)

			On Error Resume Next
			Err.Clear()
			If strSheetName <> vbNullString Then
				'Sheet may already exist so add sequential number
				objObject.Name = strSheetName
			Else
				strSheetName = "Sheet"
			End If

			If objObject.Name <> strSheetName Then
				lngCount = 1
				Do
					lngCount = lngCount + 1
					Err.Clear()
					strNumber = "(" & CStr(lngCount) & ")"
					'		objObject.Name = Left(strSheetName, 31 - Len(strNumber)) & strNumber
					If lngCount > 256 Then
						_mstrErrorMessage = "Error naming sheet"
						Exit Function
					End If
				Loop While Err.Number > 0
			End If


			On Error Resume Next
			'MH20031117 Fault 7628
			' NOTE: No templates in intranet
			'If _mxlTemplateSheet Is Nothing Then
			If _mstrXlTemplate = vbNullString Then
				With objObject.PageSetup
					' .LeftFooter = "Created on &D at &T by " & gsUsername
					.SetFooter(0, "Created on &D at &T by " & gsUsername)
					' .RightFooter = "Page &P"
					.SetFooter(2, "Page &P")
					.Orientation = IIf(_mblnXlLandscape, PageOrientationType.Landscape, PageOrientationType.Portrait)
					' .DisplayPageBreaks = False
				End With
			End If

			SetSheetName = True

		End Function


		'****************************************************************
		' NullSafeInteger
		'****************************************************************
		Function NullSafeInteger(ByVal arg As Object, _
		Optional ByVal returnIfEmpty As Integer = 0) As String

			Dim returnValue As Integer

			If (arg Is DBNull.Value) OrElse (arg Is Nothing) _
				OrElse (arg Is String.Empty) Then
				returnValue = returnIfEmpty
			Else
				Try
					returnValue = CInt(arg)
				Catch
					returnValue = returnIfEmpty
				End Try

			End If

			Return returnValue

		End Function

		Private Shared Function SaveAsFormat(strExtension As String) As SaveFormat

			strExtension = strExtension.Replace(".", "")

			Select Case UCase(strExtension)
				Case "XLS"
					Return SaveFormat.Excel97To2003
				Case "HTML"
					Return SaveFormat.Html
				Case "PDF"
					Return SaveFormat.Pdf
				Case "CSV"
					Return SaveFormat.CSV
				Case "TIFF"
					Return SaveFormat.TIFF
				Case Else
					Return SaveFormat.Xlsx

			End Select

		End Function


	End Class
End Namespace