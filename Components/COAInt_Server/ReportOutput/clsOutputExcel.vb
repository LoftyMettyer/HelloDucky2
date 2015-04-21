Option Strict Off
Option Explicit On

Imports Aspose.Cells.Charts
Imports System.Collections.Generic
Imports Aspose.Cells.Drawing
Imports HR.Intranet.Server.Enums
Imports Aspose.Cells
Imports Aspose.Cells.Pivot
Imports System.Linq
Imports System.Text.RegularExpressions

Namespace ReportOutput

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
		Private _mlngSaveExisting As ExistingFile
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

		Private _mstrXlTemplate As String
		Private _mblnXlExcelGridlines As Boolean
		Private _mblnXlExcelHeaders As Boolean
		Private _mblnXlExcelOmitTopRow As Boolean
		Private _mblnXlExcelOmitLeftCol As Boolean
		Private _mblnXlAutoFitCols As Boolean
		Private _mblnXlLandscape As Boolean

		Private _mblnSummaryReport As Boolean

		Const OfficeVersion As Integer = 12

		Public IntersectionType As IntersectionType

		Private _mcolColumns As List(Of Metadata.Column)

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
				Return _mstrErrorMessage
			End Get
		End Property

		Public Property SummaryReport() As Boolean
			Get
				SummaryReport = _mblnSummaryReport
			End Get
			Set(ByVal value As Boolean)
				_mblnSummaryReport = value
			End Set
		End Property


		Public Function GetFile(ByRef objParent As clsOutputRun, ByRef colStyles As Collection) As Boolean

			'Check if file already exists...
			If Dir(_mstrFileName) <> vbNullString And _mstrFileName <> vbNullString Then
				' TODO: We only create new now - no append to file or saveas or owt like that...
				Select Case _mlngSaveExisting
					Case ExistingFile.Overwrite

						Try
							Kill(_mstrFileName)
							GetWorkBook(strWorkbook:="New", strWorksheet:="New")

						Catch ex As Exception
							Return False

						End Try

					Case ExistingFile.DoNotOverwrite
						_mstrErrorMessage = "File already exists."

					Case ExistingFile.AddSequentialToName
						_mstrFileName = _mobjParent.GetSequentialNumberedFile(_mstrFileName)
						GetWorkBook(strWorkbook:="New", strWorksheet:="New")

					Case ExistingFile.AppendToFile
						GetWorkBook(strWorkbook:="Open", strWorksheet:="Existing")

					Case ExistingFile.CreateNewSheet
						GetWorkBook(strWorkbook:="Open", strWorksheet:="New")
				End Select
			Else
				GetWorkBook(strWorkbook:="New", strWorksheet:="New")
			End If

			GetFile = (_mstrErrorMessage = vbNullString)

		End Function

		Private Sub GetWorkBook(strWorkbook As String, strWorksheet As String)
			Dim lngCount As Integer

			Dim objCellsLicense As New License
			objCellsLicense.SetLicense("Aspose.Cells.lic")

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

		Public Sub AddPage(strDefTitle As String, strSheetName As String, ByRef colStyles As Collection)

			_mstrDefTitle = strDefTitle

			If _mblnPivotTable Then
				GetWorksheet("Data_" & strSheetName)
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

			_mcolColumns = colColumns

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

			lngExcelCol = _mlngDataStartCol
			lngExcelRow = _mlngDataCurrentRow

			If _mblnApplyStyles And Not (_mblnPivotTable) Then
				ApplyStyle(UBound(strArray, 1), UBound(strArray, 2), colStyles)
			End If

			For lngGridRow = 0 To UBound(strArray, 2)
				For lngGridCol = 0 To UBound(strArray, 1)

					With _mxlWorkSheet.Cells(lngExcelRow + lngGridRow - 1, lngExcelCol + lngGridCol - 1)

						Dim stlNumeric As Style = .GetStyle()
						Dim stlDecimal As Style = .GetStyle()
						Dim stlGeneral As Style = .GetStyle()
						Dim stlDate As Style = .GetStyle()
						Dim stlPercentage As Style = .GetStyle()

						' Replicate style formats from ActiveX...
						stlNumeric.Number = 1	' Numeric style
						stlNumeric.VerticalAlignment = TextAlignmentType.Top
						stlNumeric.HorizontalAlignment = TextAlignmentType.Right

						stlDecimal.Number = 2	' Numeric style
						stlDecimal.VerticalAlignment = TextAlignmentType.Top
						stlDecimal.HorizontalAlignment = TextAlignmentType.Right

						stlPercentage.Number = 10	' Percentage style
						stlPercentage.Custom = "0.00%"
						stlPercentage.VerticalAlignment = TextAlignmentType.Top
						stlPercentage.HorizontalAlignment = TextAlignmentType.Right

						stlGeneral.Number = 49	' Text style		
						stlGeneral.VerticalAlignment = TextAlignmentType.Top
						stlGeneral.HorizontalAlignment = TextAlignmentType.Left

						stlDate.Number = 14	' Date style
						stlDate.VerticalAlignment = TextAlignmentType.Top
						stlDate.HorizontalAlignment = TextAlignmentType.Right

						If Not strArray(lngGridCol, lngGridRow) Is Nothing Then
							Select Case colColumns.Item(lngGridCol).DataType

								Case ColumnDataType.sqlInteger

									If lngGridRow = 0 Then
										.SetStyle(stlGeneral)
										.PutValue(strArray(lngGridCol, lngGridRow))
									Else
										.SetStyle(stlNumeric)
										.PutValue(NullSafeInteger(strArray(lngGridCol, lngGridRow)))
									End If

								Case ColumnDataType.sqlNumeric

									If lngGridRow = 0 Then
										' header, so leave as a string
										.SetStyle(stlGeneral)
										.PutValue(strArray(lngGridCol, lngGridRow))
									Else
										' format as a number
										Dim numberAsString As String = strArray(lngGridCol, lngGridRow).ToString()
										Dim indexOfDecimalPoint As Integer = numberAsString.IndexOf(".", StringComparison.Ordinal)
										Dim numberOfDecimals As Integer = 0
										If indexOfDecimalPoint > 0 Then numberOfDecimals = numberAsString.Substring(indexOfDecimalPoint + 1).Length

										If numberOfDecimals > 0 Then
											If numberOfDecimals > 100 Then numberOfDecimals = 100
											stlDecimal.Custom = "0" & "." & New String("0", numberOfDecimals)
										Else
											stlDecimal.Custom = "@"
										End If


										Dim dblNumber As Double
										If numberAsString.Contains("%") Then

											dblNumber = CDbl(numberAsString.Replace("%", "")) / 100
											.SetStyle(stlPercentage)
											.PutValue(dblNumber)

										Else
											If IsNumeric(strArray(lngGridCol, lngGridRow)) Then
												.SetStyle(stlDecimal)
												.PutValue(CDbl(strArray(lngGridCol, lngGridRow)))
											Else
												.SetStyle(stlGeneral)
												.PutValue(strArray(lngGridCol, lngGridRow))
											End If
										End If


									End If
								Case ColumnDataType.sqlBoolean
									.SetStyle(stlGeneral)
									.PutValue(strArray(lngGridCol, lngGridRow))
								Case ColumnDataType.sqlUnknown
									'Leave it alone! (Required for percentages on Standard Reports)
									.SetStyle(stlGeneral)
									.PutValue(strArray(lngGridCol, lngGridRow))
								Case ColumnDataType.sqlDate
									.SetStyle(stlDate)
									.PutValue(strArray(lngGridCol, lngGridRow))
								Case Else
									Dim strValue As String = strArray(lngGridCol, lngGridRow).TrimEnd()
									If lngGridRow = 0 Then strValue = strValue.Replace("_", " ")
									If InStr(strValue, vbNewLine) > 0 Then stlGeneral.IsTextWrapped = True
									.SetStyle(stlGeneral)
									.PutValue(strValue.Replace(vbNewLine, Microsoft.VisualBasic.Constants.vbLf))
							End Select
						End If

					End With
				Next
			Next

			If _mblnChart Then	' Excel chart?
				ApplyCellOptions(_mxlWorkSheet, colStyles, True)
				CreateChart(_mxlWorkSheet, UBound(strArray, 2), UBound(strArray, 1), colStyles)
				ApplyCellOptions(_mxlWorkSheet, colStyles, False)

				'Delete superfluous rows and cols if setup in User Config reports section
				If _mblnXlExcelOmitLeftCol Then _mxlWorkSheet.Cells.DeleteColumn(0)
				If _mblnXlExcelOmitTopRow Then _mxlWorkSheet.Cells.DeleteRows(0, 1)

			ElseIf _mblnPivotTable Then

				If UBound(strArray, 1) < 1 Then
					_mstrErrorMessage = "Unable to create a pivot table for a single column of data."
				Else

					Dim pivotSheet = CreatePivotTable(_mxlWorkSheet, _mlngDataCurrentRow + UBound(strArray, 2), _mlngDataStartCol + UBound(strArray, 1), colColumns, colStyles)
					_mxlWorkSheet.VisibilityType = VisibilityType.Hidden
					ApplyCellOptions(pivotSheet, colStyles, True)

				End If
			Else
				If _mblnApplyStyles Then
					ApplyMerges(colMerges)
				End If
				ApplyCellOptions(_mxlWorkSheet, colStyles, True)

				'Delete superfluous rows and cols if setup in User Config reports section
				If _mblnXlExcelOmitLeftCol Then _mxlWorkSheet.Cells.DeleteColumn(0)
				If _mblnXlExcelOmitTopRow Then _mxlWorkSheet.Cells.DeleteRows(0, 1)
			End If


			If _mblnXlAutoFitCols Then
				_mxlWorkSheet.AutoFitColumns()
			End If

			_mlngDataCurrentRow += UBound(strArray, 2) + IIf(_mblnApplyStyles, 2, 1)

		End Sub

		Public Sub DataArrayNineBoxGrid(ByRef strArray(,) As String, ByRef colColumns As List(Of Metadata.Column), ByRef colStyles As Collection, ByRef colMerges As Collection)
			Dim lngGridCol As Integer
			Dim lngGridRow As Integer
			Dim lngExcelCol As Integer
			Dim lngExcelRow As Integer
			Dim cell As String
			Dim cellDescription As String
			Dim cellValue As String
			Dim cellColour As String
			Dim stlGeneral As Style

			_mcolColumns = colColumns

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

			lngExcelCol = _mlngDataStartCol
			lngExcelRow = _mlngDataCurrentRow

			If _mblnApplyStyles And Not (_mblnPivotTable) Then
				ApplyStyle(UBound(strArray, 1), UBound(strArray, 2), colStyles)
			End If

			For lngGridRow = 0 To UBound(strArray, 2)
				For lngGridCol = 0 To UBound(strArray, 1)
					cell = strArray(lngGridCol, lngGridRow)
					cellDescription = cell.Substring(0, cell.IndexOf("¦"))
					cellValue = cell.Substring(cell.IndexOf("¦") + 1, cell.IndexOf("|") - cell.IndexOf("¦") - 1)
					cellColour = cell.Substring(cell.IndexOf("|") + 1)

					With _mxlWorkSheet.Cells(lngExcelRow + lngGridRow - 1, lngExcelCol + lngGridCol - 1)
						stlGeneral = .GetStyle()
						With stlGeneral
							.Number = 49	'Text style
							.VerticalAlignment = TextAlignmentType.Center
							.HorizontalAlignment = TextAlignmentType.Center
							.ForegroundColor = ColorTranslator.FromHtml("#" & cellColour)
							.IsTextWrapped = True
						End With

						If cellValue Is Nothing Then
							cellValue = ""
						End If

						.SetStyle(stlGeneral)
						.PutValue(cellDescription & vbLf & cellValue)
						.Characters(cellDescription.Length + 1, cellValue.Length).Font.Size = 20 'Set the font size of the value
					End With
				Next
			Next

			'Insert rows and columns to contain the axis labels
			_mxlWorkSheet.Cells.InsertColumns(0, 2)
			_mxlWorkSheet.Cells.InsertRows(_mxlWorkSheet.Cells.Rows.Count, 2)
			'X axis
			With _mxlWorkSheet.Cells(9, 4) 'row,col
				stlGeneral = .GetStyle
				stlGeneral.HorizontalAlignment = TextAlignmentType.Center
				stlGeneral.VerticalAlignment = TextAlignmentType.Center
				stlGeneral.Font.Size = 14
				.SetStyle(stlGeneral)
				.PutValue(_mobjParent.AxisLabelsAsArray(0))
				'.PutValue("9,4")
			End With
			With _mxlWorkSheet.Cells(8, 3) 'row,col
				stlGeneral = .GetStyle
				stlGeneral.HorizontalAlignment = TextAlignmentType.Center
				stlGeneral.VerticalAlignment = TextAlignmentType.Center
				stlGeneral.Font.Size = 12
				SetCellBorders(stlGeneral)
				.SetStyle(stlGeneral)
				.PutValue(_mobjParent.AxisLabelsAsArray(1))
				'.PutValue("8,3")
			End With
			With _mxlWorkSheet.Cells(8, 4) 'row,col
				stlGeneral = .GetStyle
				stlGeneral.HorizontalAlignment = TextAlignmentType.Center
				stlGeneral.VerticalAlignment = TextAlignmentType.Center
				stlGeneral.Font.Size = 12
				SetCellBorders(stlGeneral)
				.SetStyle(stlGeneral)
				.PutValue(_mobjParent.AxisLabelsAsArray(2))
				'.PutValue("8,4")
			End With
			With _mxlWorkSheet.Cells(8, 5) 'row,col
				stlGeneral = .GetStyle
				stlGeneral.HorizontalAlignment = TextAlignmentType.Center
				stlGeneral.VerticalAlignment = TextAlignmentType.Center
				stlGeneral.Font.Size = 12
				SetCellBorders(stlGeneral)
				.SetStyle(stlGeneral)
				.PutValue(_mobjParent.AxisLabelsAsArray(3))
				'.PutValue("8,5")
			End With
			'Y axis
			With _mxlWorkSheet.Cells(6, 1) 'row,col
				stlGeneral = .GetStyle
				stlGeneral.RotationAngle = 90
				stlGeneral.HorizontalAlignment = TextAlignmentType.Center
				stlGeneral.VerticalAlignment = TextAlignmentType.Center
				stlGeneral.Font.Size = 14
				.SetStyle(stlGeneral)
				.PutValue(_mobjParent.AxisLabelsAsArray(4))
				'.PutValue("6,1")
			End With
			With _mxlWorkSheet.Cells(5, 2) 'row,col
				stlGeneral = .GetStyle
				stlGeneral.RotationAngle = 90
				stlGeneral.HorizontalAlignment = TextAlignmentType.Center
				stlGeneral.VerticalAlignment = TextAlignmentType.Center
				stlGeneral.Font.Size = 12
				SetCellBorders(stlGeneral)
				.SetStyle(stlGeneral)
				.PutValue(_mobjParent.AxisLabelsAsArray(5))
				'.PutValue("5,2")
			End With
			With _mxlWorkSheet.Cells(6, 2) 'row,col
				stlGeneral = .GetStyle
				stlGeneral.RotationAngle = 90
				stlGeneral.HorizontalAlignment = TextAlignmentType.Center
				stlGeneral.VerticalAlignment = TextAlignmentType.Center
				stlGeneral.Font.Size = 12
				SetCellBorders(stlGeneral)
				.SetStyle(stlGeneral)
				.PutValue(_mobjParent.AxisLabelsAsArray(6))
				'.PutValue("6,2")
			End With
			With _mxlWorkSheet.Cells(7, 2) 'row,col
				stlGeneral = .GetStyle
				stlGeneral.RotationAngle = 90
				stlGeneral.HorizontalAlignment = TextAlignmentType.Center
				stlGeneral.VerticalAlignment = TextAlignmentType.Center
				stlGeneral.Font.Size = 12
				SetCellBorders(stlGeneral)
				.SetStyle(stlGeneral)
				.PutValue(_mobjParent.AxisLabelsAsArray(7))
				'.PutValue("7,2")
			End With

			'Merge the X axis label into three columns and add borders to it
			_mxlWorkSheet.Cells.Merge(9, 3, 1, 3)
			With _mxlWorkSheet.Cells(9, 3)
				stlGeneral = .GetStyle
				SetCellBorders(stlGeneral)
				.SetStyle(stlGeneral)
			End With
			With _mxlWorkSheet.Cells(9, 4)
				stlGeneral = .GetStyle
				SetCellBorders(stlGeneral)
				.SetStyle(stlGeneral)
			End With
			With _mxlWorkSheet.Cells(9, 5)
				stlGeneral = .GetStyle
				SetCellBorders(stlGeneral)
				.SetStyle(stlGeneral)
			End With

			'Merge the Y axis label into three rows and add borders to it
			_mxlWorkSheet.Cells.Merge(5, 1, 3, 1)
			With _mxlWorkSheet.Cells(5, 1)
				stlGeneral = .GetStyle
				SetCellBorders(stlGeneral)
				.SetStyle(stlGeneral)
			End With
			With _mxlWorkSheet.Cells(6, 1)
				stlGeneral = .GetStyle
				SetCellBorders(stlGeneral)
				.SetStyle(stlGeneral)
			End With
			With _mxlWorkSheet.Cells(7, 1)
				stlGeneral = .GetStyle
				SetCellBorders(stlGeneral)
				.SetStyle(stlGeneral)
			End With

			If _mblnApplyStyles Then
				ApplyMerges(colMerges)
			End If
			ApplyCellOptions(_mxlWorkSheet, colStyles, True)

			With _mxlWorkSheet.Cells(1, 2) 'Increase the size of the title
				stlGeneral = .GetStyle
				stlGeneral.Font.Size = 14
				.SetStyle(stlGeneral)
			End With
			_mxlWorkSheet.Cells.Merge(1, 2, 1, 5)	'Merge the title into 5 columns

			_mxlWorkSheet.AutoFitColumns()
			_mxlWorkSheet.AutoFitRows()

			'Set the cells' height and width
			_mxlWorkSheet.Cells.SetRowHeight(5, 110)
			_mxlWorkSheet.Cells.SetRowHeight(6, 110)
			_mxlWorkSheet.Cells.SetRowHeight(7, 110)
			_mxlWorkSheet.Cells.SetRowHeight(8, 40)
			_mxlWorkSheet.Cells.SetRowHeight(9, 40)
			_mxlWorkSheet.Cells.SetColumnWidth(1, 8)
			_mxlWorkSheet.Cells.SetColumnWidth(2, 8)
			_mxlWorkSheet.Cells.SetColumnWidth(3, 25)
			_mxlWorkSheet.Cells.SetColumnWidth(4, 25)
			_mxlWorkSheet.Cells.SetColumnWidth(5, 25)

			_mlngDataCurrentRow += UBound(strArray, 2) + IIf(_mblnApplyStyles, 2, 1)
		End Sub

		Private Sub CreateChart(ByRef objDataSheet As Worksheet, lngMaxRows As Integer, lngMaxCols As Integer, colStyles As Collection)

			Dim strSheetName As String

			Try

				strSheetName = GetSheetName(objDataSheet.Name.Replace("Data_", "") & " chart")

				Dim mxlChartWorkSheet = _mxlWorkBook.Worksheets.Insert(objDataSheet.Index, SheetType.Chart, strSheetName)

				Dim chartIndex As Integer = mxlChartWorkSheet.Charts.Add(ChartType.Column3DClustered, 0, 0, 30, 20)
				Dim xlChart As Chart = mxlChartWorkSheet.Charts(chartIndex)
				Dim xlData As Range
				Dim xlCategories As Range

				Const colTitleRowCount As Integer = 1

				_mlngDataCurrentRow -= 1

				Dim dataFirstRow As Integer = _mlngDataCurrentRow
				Dim dataFirstCol As Integer = _mlngDataStartCol - 1
				Dim dataRowCount As Integer = lngMaxRows + colTitleRowCount
				Dim dataColumnCount As Integer = IIf(lngMaxCols < 1, 1, lngMaxCols)

				If SummaryReport Then

					' some logic to add right most numeric columns
					Dim bNumericFound As Boolean = False

					dataFirstCol = 0
					For Each objColumn In _mcolColumns
						If (objColumn.DataType = ColumnDataType.sqlNumeric Or objColumn.DataType = ColumnDataType.sqlInteger) And dataFirstCol > 0 Then
							bNumericFound = True
							Exit For
						End If

						dataFirstCol += 1
					Next

					xlCategories = _mxlWorkSheet.Cells.CreateRange(dataFirstRow + 1, 1, dataRowCount - 1, dataFirstCol)
					xlData = _mxlWorkSheet.Cells.CreateRange(dataFirstRow + 1, dataFirstCol + 1, dataRowCount - 1, IIf(bNumericFound, _mcolColumns.Count() - dataFirstCol, 1))


				Else
					xlCategories = _mxlWorkSheet.Cells.CreateRange(dataFirstRow + 1, dataFirstCol, dataRowCount - 1, 1)
					xlData = _mxlWorkSheet.Cells.CreateRange(dataFirstRow + 1, dataFirstCol + 1, dataRowCount - 1, dataColumnCount)
				End If

				xlCategories.Name = "xlCategories"
				xlData.Name = "xlData"

				With xlChart

					.NSeries.Add("=xlData", True)
					.NSeries.CategoryData = "=xlCategories"

					Dim iSeries As Integer = dataFirstCol + 2
					For Each objSeries In .NSeries
						FormatSeries(iSeries, objSeries)
						objSeries.Name = String.Format("={0}!{1}{2}", _mxlWorkSheet.Name, NumberToExcelColumn(iSeries), dataFirstRow + 1)
						iSeries += 1
					Next

					.Title.Text = _mstrDefTitle
					.Title.IsVisible = True
					.Title.Font.IsBold = True
					.Title.Font.Size = 12
					If colStyles.Item("Title").Underline Then .Title.Font.Underline = FontUnderlineType.Single
					.Title.Font.Color = ColorTranslator.FromWin32(colStyles.Item("Title").ForeCol)

					.RightAngleAxes = False
					.Perspective = 15

					.PlotArea.Area.ForegroundColor = Color.White
					.ChartArea.Area.ForegroundColor = Color.White
					.Walls.ForegroundColor = Color.White

					.Calculate()

					.PlotArea.Border.IsVisible = False
					.ChartArea.Border.IsVisible = False

				End With

			Catch ex As Exception
				_mstrErrorMessage = ex.Message

			End Try

		End Sub

		Private Sub FormatSeries(iNumber As Integer, ByRef objSeries As Series)

			objSeries.Shadow = True
			objSeries.Smooth = True

			Dim spPr As ShapePropertyCollection = objSeries.ShapeProperties
			Dim fmt3d As Format3D = spPr.Format3D
			Dim bevel As Bevel = fmt3d.TopBevel

			bevel.Type = BevelPresetType.Circle
			bevel.Height = 2
			bevel.Width = 5

			fmt3d.SurfaceMaterialType = PresetMaterialType.WarmMatte
			fmt3d.SurfaceLightingType = LightRigType.ThreePoint
			fmt3d.LightingAngle = 20

		End Sub


		Private Function CreatePivotTable(ByRef objDataSheet As Worksheet, lngMaxRows As Integer, lngMaxCols As Integer, colColumns As List(Of Metadata.Column), colStyles As Collection) As Worksheet

			Dim pivotSheet As Worksheet = _mxlWorkBook.Worksheets(_mxlWorkBook.Worksheets.Add())
			pivotSheet.Name = objDataSheet.Name.Replace("Data_", "")

			Try

				Dim pivotTables As PivotTableCollection = pivotSheet.PivotTables

				Dim sRange = String.Format("={0}!{1}:{2}{3}", objDataSheet.Name, objDataSheet.Cells.FirstCell.Name, NumberToExcelColumn(lngMaxCols), lngMaxRows)
				Dim index As Integer = pivotTables.Add(sRange, "B4", "PivotTable1")

				With pivotTables(index)
					.AddFieldToArea(PivotFieldType.Row, 1)
					.AddFieldToArea(PivotFieldType.Column, 0)
					.AddFieldToArea(PivotFieldType.Data, pivotTables(index).BaseFields.Count - 1)

					Select Case IntersectionType
						Case IntersectionType.Average
							.DataFields(0).Function = ConsolidationFunction.Average

						Case IntersectionType.Minimum
							.DataFields(0).Function = ConsolidationFunction.Min

						Case IntersectionType.Maximum
							.DataFields(0).Function = ConsolidationFunction.Max

						Case IntersectionType.Count
							.DataFields(0).Function = ConsolidationFunction.Count

						Case Else
							.DataFields(0).Function = ConsolidationFunction.Sum
					End Select

					.IsAutoFormat = False
					.AutoFormatType = PivotTableAutoFormatType.None
					.PivotTableStyleType = PivotTableStyleType.None
					.PreserveFormatting = True

					.RowGrand = True
					.ColumnGrand = True
					.PageFieldOrder = PrintOrderType.DownThenOver
					.RowFields(0).IsAscendSort = True
					.ColumnFields(0).IsAscendSort = True

					.ShowPivotStyleRowHeader = True
					.ShowPivotStyleColumnHeader = True
					.DisplayNullString = True
					.CalculateData()
					.RefreshDataOnOpeningFile = False

					ApplyStyleToRange(.RowRange.ToRange(pivotSheet), colStyles("Heading"))
					ApplyStyleToRange(.ColumnRange.ToRange(pivotSheet), colStyles("Heading"))
					ApplyStyleToRange(.DataBodyRange.ToRange(pivotSheet), colStyles("Data"))

					Dim stlDecimal = pivotSheet.Cells(.DataBodyRange.StartRow, .DataBodyRange.StartColumn).GetStyle()
					stlDecimal.Number = 2

					Dim objColumn = colColumns(colColumns.Count - 1)
					If objColumn.Decimals > 0 Then
						stlDecimal.Custom = "#,##0" & "." & New String("0", objColumn.Decimals)
					Else
						stlDecimal.Custom = "@"
					End If

					.DataBodyRange.ToRange(pivotSheet).SetStyle(stlDecimal)

				End With

				pivotSheet.AutoFitColumns()


			Catch ex As Exception
				_mstrErrorMessage = ex.Message
				Throw

			End Try

			Return pivotSheet

		End Function

		Private Sub ApplyCellOptions(worksheet As Worksheet, ByRef colStyles As Collection, blnGridLines As Boolean)

			Dim objRange As Range

			Try

				If _mblnApplyStyles Then
					If blnGridLines Then
						worksheet.IsGridlinesVisible = _mblnXlExcelGridlines
						worksheet.IsRowColumnHeadersVisible = _mblnXlExcelHeaders
					End If

					With colStyles.Item("Title")
						.StartCol = glngSettingTitleCol
						.StartRow = IIf(_mlngAppendStartRow > 0, _mlngAppendStartRow, glngSettingTitleRow)
						.EndCol = .StartCol
						.EndRow = .StartRow
					End With

					'Put title in after autofit...
					If colStyles.Item("Title").StartCol <> 0 And colStyles.Item("Title").StartRow <> 0 Then
						worksheet.Cells(colStyles.Item("Title").StartRow - 1, colStyles.Item("Title").StartCol - 1).Value = _mstrDefTitle
						objRange = worksheet.Cells.CreateRange(colStyles.Item("Title").StartRow - 1, colStyles.Item("Title").StartCol - 1, 1, 1)
						ApplyStyleToRange(objRange, colStyles.Item("Title"))
					End If
				End If

			Catch ex As Exception
				_mstrErrorMessage = ex.Message
			End Try

		End Sub


		Private Sub ApplyStyle(lngNumCols As Integer, lngNumRows As Integer, ByRef colStyles As Collection)

			Dim objStyle As clsOutputStyle
			Dim objRange As Range
			Dim lngCol As Integer = _mlngDataStartCol
			Dim lngRow As Integer = _mlngDataCurrentRow

			Try

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

							objRange = _mxlWorkSheet.Cells.CreateRange(objStyle.StartRow + (lngRow - 1), objStyle.StartCol + lngCol - 1, totalRows, totalCols)
							ApplyStyleToRange(objRange, objStyle)
						End If
					End If
				Next objStyle

			Catch ex As Exception
				_mstrErrorMessage = ex.Message
			End Try

		End Sub


		Private Sub ApplyMerges(ByRef colMerges As Collection)

			Dim objMerge As clsOutputStyle
			Dim objRange As Range
			Dim lngCol As Integer
			Dim lngRow As Integer

			Try

				lngCol = _mlngDataStartCol
				lngRow = _mlngDataCurrentRow

				For Each objMerge In colMerges
					If objMerge.StartRow + lngRow > 0 And objMerge.StartCol + lngCol > 0 Then
						Dim totalRows = (objMerge.EndRow + lngRow) - objMerge.StartRow
						Dim totalCols = (objMerge.EndCol + lngCol) - objMerge.StartCol
						objRange = _mxlWorkSheet.Cells.CreateRange(objMerge.StartRow + lngRow - 1, objMerge.StartCol + lngCol - 1, totalRows, totalCols)
						objRange.Merge()
					End If
				Next objMerge

			Catch ex As Exception
				_mstrErrorMessage = ex.Message.RemoveSensitive()

			End Try

		End Sub

		Private Sub ApplyStyleToRange(ByRef objRange As Range, objStyle As clsOutputStyle)

			Try

				Dim rangeStyle As Style = _mxlWorkBook.Styles(_mxlWorkBook.Styles.Add())
				rangeStyle.Name = objStyle.Name

				With objRange

					If objStyle.CenterText Then
						rangeStyle.HorizontalAlignment = TextAlignmentType.Center
					Else
						rangeStyle.HorizontalAlignment = TextAlignmentType.Left
					End If

					rangeStyle.Font.Name = objStyle.Font.Name
					rangeStyle.Font.Size = objStyle.Font.Size
					rangeStyle.Font.IsBold = objStyle.Bold
					If objStyle.Underline Then rangeStyle.Font.Underline = FontUnderlineType.Single
					rangeStyle.Font.Color = ColorTranslator.FromWin32(objStyle.ForeCol)

					'Don't do the backcol nor gridlines for the title...
					If objStyle.Name <> "Title" Then
						' We use foregroundColor for the background...
						rangeStyle.ForegroundColor = ColorTranslator.FromWin32(objStyle.BackCol)
						rangeStyle.Pattern = BackgroundType.Solid

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

			Catch ex As Exception
				_mstrErrorMessage = ex.Message

			End Try

		End Sub

		Public Sub Complete()

			Try

				If _mstrErrorMessage <> vbNullString Then
					Exit Sub
				End If

				_mxlWorkBook.Worksheets.ActiveSheetIndex = 0
				_mxlWorkBook.Settings.FirstVisibleTab = 0

				'SAVE
				_mstrErrorMessage = "Error saving file <" & _mstrFileName & ">"
				_mxlWorkBook.Save(_mstrFileName, SaveAsFormat(DownloadExtension))
				_mstrErrorMessage = vbNullString

				ClearUp()

			Catch ex As Exception
				_mstrErrorMessage &= ex.Message.RemoveSensitive()

			End Try

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

		Private Shared Function FormatSheetName(strSheetName As String) As String

			If Left(strSheetName, 1) = "'" Then
				strSheetName = " " & strSheetName
			End If

			Dim separators As Char() = New Char() {"\"c, "/"c, "*"c, ":"c, "["c, "]"c, "?"c, ","c, ControlChars.Quote}
			strSheetName.ReplaceMultiple(separators, "")
			strSheetName = strSheetName.Replace("&amp;", "&")

			Do While InStr(strSheetName, "  ") > 0
				strSheetName = Replace(strSheetName, "  ", " ")
			Loop

			Return strSheetName

		End Function

		Private Function GetSheetName(desiredName As String) As String

			Dim bNameOK As Boolean = False
			Dim iCount As Integer = 0
			Dim sSheetName As String = desiredName

			If sSheetName = "" Then sSheetName = "Sheet"
			sSheetName = Left(Trim(sSheetName), 31)

			Do While Not bNameOK

				If Not _mxlWorkBook.Worksheets.Any(Function(oSheet) oSheet.Name = sSheetName) Then
					bNameOK = True
				Else

					iCount += 1

					sSheetName = If(sSheetName.Length <= 25, sSheetName, sSheetName.Substring(0, 25))
					sSheetName = String.Format("{0} {1}", sSheetName, iCount)
				End If
			Loop

			Return sSheetName

		End Function

		Private Sub SetSheetName(ByRef objObject As Worksheet, strSheetName As String)

			Try
				strSheetName = FormatSheetName(Regex.Replace(strSheetName, "[:\\\/?\*\[\]]", " ")) 'Replace invalid characters with space so Aspose doesn't throw a wobbly when creating the Excel tabs

				If _mxlWorkBook.Worksheets.Count < 255 Then
					objObject.Name = GetSheetName(strSheetName)
				End If

				If _mstrXlTemplate = vbNullString Then
					With objObject.PageSetup
						' .LeftFooter = "Created on &D at &T by " & gsUsername
						.SetFooter(0, "Created on &D at &T by " & UserName)
						' .RightFooter = "Page &P"
						.SetFooter(2, "Page &P")
						.Orientation = IIf(_mblnXlLandscape, PageOrientationType.Landscape, PageOrientationType.Portrait)
						' .DisplayPageBreaks = False
					End With
				End If


			Catch ex As Exception
				Throw

			End Try

		End Sub


		'****************************************************************
		' NullSafeInteger
		'****************************************************************

		Function NullSafeInteger(ByVal arg As Object, _
		Optional ByVal returnIfEmpty As Integer = 0) As Integer

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

		Private Shared Function NumberToExcelColumn(num As Integer) As String
			' Subtract one to make modulo/divide cleaner. '

			num = num - 1

			' Select return value based on invalid/one-char/two-char input. '

			If num < 0 Or num >= 27 * 26 Then
				Return "-"
			Else
				' Single char, just get the letter. '

				If num < 26 Then
					Return Chr(num + 65)
				Else
					' Double char, get letters based on integer divide and modulus. '

					Return Chr(num \ 26 + 64) + Chr(num Mod 26 + 65)
				End If
			End If
		End Function

		Private Sub SetCellBorders(ByRef stlGeneral As Style)
			stlGeneral.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black) 'Top border
			stlGeneral.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black)	'Bottom border
			stlGeneral.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black)	'Left border
			stlGeneral.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black) 'Right border
		End Sub
	End Class
End Namespace