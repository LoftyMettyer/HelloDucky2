Option Strict Off
Option Explicit On

Imports HR.Intranet.Server.ExClientCode
Imports HR.Intranet.Server.Enums
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6

Public Class clsOutputRun


	Private mobjOutputType As Object
	'Private mfrmOutput As frmOutputOptions
	Private mcolStyles As Collection
	Private mcolMerges As Collection
	Private mcolColumns As Collection
	Private mstrErrorMessage As String

	Private mlngFormat As Integer
	Private mstrFunction As String
	Private mstrEmailAddresses As String
	Private mstrEmailSubject As String
	Private mstrEmailAttachAs As String
	Private mblnPageRange As Boolean
	Private mblnPrintData As Boolean
	Private mstrPrinterName As String

	Private mstrDefaultPrinter As String
	Private mblnSizeColumnsIndependently As Boolean
	Private mlngHeaderRows As Integer
	Private mlngHeaderCols As Integer
	Private mstrArray(,) As String

	Private mstrSaveAsValues As String

	Private mblnData As Boolean
	Private mblnCSV As Boolean
	Private mblnHTML As Boolean
	Private mblnWord As Boolean
	Private mblnExcel As Boolean
	Private mblnChart As Boolean
	Private mblnPivot As Boolean
	Private mblnPivotSuppressBlanks As Boolean
	Private mblnIndicatorColumn As Boolean
	Private mblnPageTitles As Boolean

	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()

		'Set mfrmOutput = New frmOutputOptions
		mcolColumns = New Collection

		mlngHeaderRows = 1
		mlngHeaderCols = 0
		mblnPageTitles = True
		mblnPivotSuppressBlanks = True
		mblnIndicatorColumn = False

		mblnData = True
		mblnCSV = True
		mblnHTML = True
		mblnWord = True
		mblnExcel = True
		mblnChart = True
		mblnPivot = True

		'InitialiseStyles
		mcolMerges = New Collection

	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub

	Public Sub InitialiseStyles()

		Dim objStyle As clsOutputStyle

		mcolStyles = New Collection

		objStyle = New clsOutputStyle
		With objStyle
			.Name = "Title"
			.StartCol = glngSettingTitleCol
			.StartRow = glngSettingTitleRow
			.Gridlines = gblnSettingTitleGridlines
			.Bold = gblnSettingTitleBold
			.Underline = gblnSettingTitleUnderline
			.BackCol = glngSettingTitleBackcolour
			.ForeCol = glngSettingTitleForecolour
			.BackCol97 = glngSettingTitleBackcolour97
			.ForeCol97 = glngSettingTitleForecolour97
		End With

		mcolStyles.Add(objStyle, objStyle.Name)


		objStyle = New clsOutputStyle
		With objStyle
			.Name = "Heading"
			.Gridlines = gblnSettingHeadingGridlines
			.Bold = gblnSettingHeadingBold
			.Underline = gblnSettingHeadingUnderline
			.BackCol = glngSettingHeadingBackcolour
			.ForeCol = glngSettingHeadingForecolour
			.BackCol97 = glngSettingHeadingBackcolour97
			.ForeCol97 = glngSettingHeadingForecolour97
			.CenterText = True
		End With

		mcolStyles.Add(objStyle, objStyle.Name)


		objStyle = New clsOutputStyle
		With objStyle
			.Name = "HeadingCols"
			.StartCol = 0
			.StartRow = 0
			.Gridlines = gblnSettingHeadingGridlines
			.Bold = gblnSettingHeadingBold
			.Underline = gblnSettingHeadingUnderline
			.BackCol = glngSettingHeadingBackcolour
			.ForeCol = glngSettingHeadingForecolour
			.BackCol97 = glngSettingHeadingBackcolour97
			.ForeCol97 = glngSettingHeadingForecolour97
		End With

		mcolStyles.Add(objStyle, objStyle.Name)


		objStyle = New clsOutputStyle
		With objStyle
			.Name = "Data"
			.StartCol = glngSettingDataCol
			.StartRow = glngSettingDataRow
			.Gridlines = gblnSettingDataGridlines
			.Bold = gblnSettingDataBold
			.Underline = gblnSettingDataUnderline
			.BackCol = glngSettingDataBackcolour
			.ForeCol = glngSettingDataForecolour
			.BackCol97 = glngSettingDataBackcolour97
			.ForeCol97 = glngSettingDataForecolour97
		End With

		mcolStyles.Add(objStyle, objStyle.Name)


		'UPGRADE_NOTE: Object objStyle may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objStyle = Nothing

	End Sub


	Public Function AddColumn(ByRef strHeading As String, ByRef lngDataType As Integer, ByRef lngDecimals As Integer, ByRef bThousandSeparator As Boolean) As Boolean

		Dim objColumn As clsColumn

		On Error GoTo LocalErr
		AddColumn = True

		objColumn = New clsColumn
		objColumn.Heading = strHeading
		objColumn.DataType = lngDataType
		objColumn.DecPlaces = lngDecimals
		objColumn.ThousandSeparator = bThousandSeparator

		mcolColumns.Add(objColumn)

		Exit Function

LocalErr:
		AddColumn = False

	End Function


	Public Function AddStyle(ByRef strType As String, ByRef lngStartCol As Integer, ByRef lngStartRow As Integer, ByRef lngEndCol As Integer, ByRef lngEndRow As Integer, Optional ByRef lngBackCol As Object = Nothing, Optional ByRef lngForeCol As Object = Nothing, Optional ByRef blnBold As Object = Nothing, Optional ByRef blnUnderline As Object = Nothing, Optional ByRef blnGridLines As Object = Nothing, Optional ByRef lngBackCol97 As Object = Nothing, Optional ByRef lngForeCol97 As Object = Nothing) As Boolean

		Dim objStyle As clsOutputStyle

		On Error GoTo LocalErr
		AddStyle = True

		objStyle = New clsOutputStyle

		With objStyle
			.StartCol = lngStartCol	'(mcolStyles("Data").StartCol)
			.StartRow = lngStartRow	'(mcolStyles("Data").StartRow)
			.EndCol = lngEndCol	'(mcolStyles("Data").StartCol)
			.EndRow = lngEndRow	'(mcolStyles("Data").StartRow)

			Select Case strType
				Case "Title", "Heading", "Data"
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolStyles().BackCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.BackCol = mcolStyles.Item(strType).BackCol
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolStyles().ForeCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.ForeCol = mcolStyles.Item(strType).ForeCol
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolStyles().Bold. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Bold = mcolStyles.Item(strType).Bold
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolStyles().Underline. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Underline = mcolStyles.Item(strType).Underline
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolStyles().Gridlines. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Gridlines = mcolStyles.Item(strType).Gridlines
			End Select

			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(lngBackCol) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object lngBackCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BackCol = lngBackCol
			End If

			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(lngForeCol) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object lngForeCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.ForeCol = lngForeCol
			End If

			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(blnBold) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object blnBold. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Bold = blnBold
			End If

			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(blnUnderline) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object blnUnderline. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Underline = blnUnderline
			End If

			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(blnGridLines) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object blnGridLines. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Gridlines = blnGridLines
			End If

			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(lngBackCol97) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object lngBackCol97. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BackCol97 = lngBackCol97
			End If

			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(lngForeCol97) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object lngForeCol97. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.ForeCol97 = lngForeCol97
			End If

		End With

		mcolStyles.Add(objStyle)
		'UPGRADE_NOTE: Object objStyle may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objStyle = Nothing


		Exit Function

LocalErr:
		AddStyle = False

	End Function


	Public Function AddMerge(ByRef lngStartCol As Integer, ByRef lngStartRow As Integer, ByRef lngEndCol As Integer, ByRef lngEndRow As Integer) As Boolean

		Dim objMerge As clsOutputStyle

		On Error GoTo LocalErr
		AddMerge = True

		If lngStartCol <> lngEndCol Or lngStartRow <> lngEndRow Then
			objMerge = New clsOutputStyle
			objMerge.StartCol = lngStartCol
			objMerge.StartRow = lngStartRow
			objMerge.EndCol = lngEndCol
			objMerge.EndRow = lngEndRow

			mcolMerges.Add(objMerge)
			'UPGRADE_NOTE: Object objMerge may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objMerge = Nothing
		End If

		Exit Function

LocalErr:
		AddMerge = False

	End Function

	Public Function ResetColumns() As Object
		'UPGRADE_NOTE: Object mcolColumns may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mcolColumns = Nothing
		mcolColumns = New Collection
	End Function

	Public Function ResetStyles() As Object
		'UPGRADE_NOTE: Object mcolStyles may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mcolStyles = Nothing
		InitialiseStyles()
	End Function

	Public Function ResetMerges() As Object
		'UPGRADE_NOTE: Object mcolMerges may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mcolMerges = Nothing
		mcolMerges = New Collection
	End Function

	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'Set mfrmOutput = Nothing
		'UPGRADE_NOTE: Object mcolStyles may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mcolStyles = Nothing
		'UPGRADE_NOTE: Object mcolMerges may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mcolMerges = Nothing
		'UPGRADE_NOTE: Object mcolColumns may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mcolColumns = Nothing
		'UPGRADE_NOTE: Object mobjOutputType may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mobjOutputType = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub


	Public Function GetFile() As Object
		On Error Resume Next
		'UPGRADE_WARNING: Couldn't resolve default property of object GetFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetFile = True
		'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.GetFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetFile = mobjOutputType.GetFile(Me, mcolStyles)
	End Function

	Public Function AddPage(ByRef strDefTitle As String, Optional ByRef mstrSheetName As String = "") As Object
		On Error Resume Next
		'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.AddPage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mobjOutputType.AddPage(strDefTitle, mstrSheetName, mcolStyles)
	End Function

	Public WriteOnly Property SizeColumnsIndependently() As Boolean
		Set(ByVal Value As Boolean)
			mblnSizeColumnsIndependently = Value
		End Set
	End Property

	''Public Function DataArrayToGrid(strArray() As String, pgrdNew As SSDBGrid)
	''
	''  'Not all classes support all properties so catch any errors...
	''  On Local Error Resume Next
	''  mobjOutputType.SizeColumnsIndependently = mblnSizeColumnsIndependently
	''  mobjOutputType.HeaderCols = mlngHeaderCols
	''  mobjOutputType.HeaderRows = mlngHeaderRows
	''
	''  'On Local Error GoTo 0
	''  If mblnPrintData Then
	''    DataGrid ConvertToGrid(strArray(), pgrdNew)
	''  Else
	''    mobjOutputType.DataArray strArray(), mcolColumns, mcolStyles, mcolMerges
	''  End If
	''
	''End Function

	'Public Function RecordProfilePage(pfrmRecProfile As Form, _
	''  piPageNumber As Integer) As Boolean
	'  'Not all classes support all properties so catch any errors...
	'  On Local Error Resume Next
	''''  mobjOutputType.SizeColumnsIndependently = mblnSizeColumnsIndependently
	''''  mobjOutputType.HeaderCols = mlngHeaderCols
	''''  mobjOutputType.HeaderRows = mlngHeaderRows
	'
	'  On Local Error GoTo 0
	'
	''''  mobjOutputType.DataArray strArray(), mcolColumns, mcolStyles, mcolMerges
	'  RecordProfilePage = mobjOutputType.RecordProfilePage(pfrmRecProfile, piPageNumber, mcolStyles)
	'
	'End Function

	Public ReadOnly Property PrintData() As Boolean
		Get
			PrintData = mblnPrintData
		End Get
	End Property

	Public ReadOnly Property PrinterName() As String
		Get
			PrinterName = mstrPrinterName
		End Get
	End Property

	Public WriteOnly Property HeaderRows() As Integer
		Set(ByVal Value As Integer)
			mlngHeaderRows = Value
		End Set
	End Property

	Public WriteOnly Property HeaderCols() As Integer
		Set(ByVal Value As Integer)
			mlngHeaderCols = Value
		End Set
	End Property

	Public WriteOnly Property FileDelimiter() As String
		Set(ByVal Value As String)
			'clsOutputCSV only
			On Error Resume Next
			'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.FileDelimiter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mobjOutputType.FileDelimiter = Value
		End Set
	End Property

	Public WriteOnly Property EncloseInQuotes() As Boolean
		Set(ByVal Value As Boolean)
			'clsOutputCSV only
			On Error Resume Next
			'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.EncloseInQuotes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mobjOutputType.EncloseInQuotes = Value
		End Set
	End Property

	Public WriteOnly Property ApplyStyles() As Boolean
		Set(ByVal Value As Boolean)
			On Error Resume Next
			'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.ApplyStyles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mobjOutputType.ApplyStyles = Value
		End Set
	End Property

	Public ReadOnly Property ErrorMessage() As String
		Get
			If mstrErrorMessage <> vbNullString Then
				ErrorMessage = mstrErrorMessage
			Else
				If Not (mobjOutputType Is Nothing) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.ErrorMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ErrorMessage = mobjOutputType.ErrorMessage
				End If
			End If
		End Get
	End Property


	Public Property UserName() As String
		Get
			UserName = gsUsername
		End Get
		Set(ByVal Value As String)
			gsUsername = Value
		End Set
	End Property

	'Public Property Get Format() As OutputFormats
	'  Format = mlngFormat
	'End Property
	'
	'Public Property Let Format(lngNewValue As OutputFormats)
	'  mlngFormat = Format
	'End Property


	Public Property PageTitles() As Boolean
		Get
			PageTitles = mblnPageTitles
		End Get
		Set(ByVal Value As Boolean)
			mblnPageTitles = Value
		End Set
	End Property

	Public ReadOnly Property EmailAttachAs() As String
		Get
			EmailAttachAs = mstrEmailAttachAs
		End Get
	End Property


	Public ReadOnly Property PrintPrompt() As Boolean
		Get

			'Dim blnPromptConfig As Boolean
			'blnPromptConfig = (GetPCSetting("Printer", "Prompt", False) = True)
			'PrintPrompt = (blnPromptConfig And Not gblnBatchMode)

		End Get
	End Property


	Public Property PivotSuppressBlanks() As Boolean
		Get
			PivotSuppressBlanks = mblnPivotSuppressBlanks
		End Get
		Set(ByVal Value As Boolean)
			mblnPivotSuppressBlanks = Value
		End Set
	End Property


	Public Property PivotDataFunction() As String
		Get
			If mstrFunction = vbNullString Then mstrFunction = "Total"
			PivotDataFunction = mstrFunction
		End Get
		Set(ByVal Value As String)
			mstrFunction = Value
		End Set
	End Property


	Public Property IndicatorColumn() As Boolean
		Get
			IndicatorColumn = mblnIndicatorColumn
		End Get
		Set(ByVal Value As Boolean)
			mblnIndicatorColumn = Value
		End Set
	End Property


	Public Property SaveAsValues() As String
		Get
			SaveAsValues = mstrSaveAsValues
		End Get
		Set(ByVal Value As String)
			mstrSaveAsValues = Value
		End Set
	End Property


	Public Function DataArray() As Object

		If UBound(mstrArray, 2) < 1 Then
			mstrErrorMessage = "No data to output."
			Exit Function
		End If

		'Not all classes support all properties so catch any errors...
		On Error Resume Next
		'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.SizeColumnsIndependently. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mobjOutputType.SizeColumnsIndependently = mblnSizeColumnsIndependently
		'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.HeaderCols. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mobjOutputType.HeaderCols = mlngHeaderCols
		'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.HeaderRows. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mobjOutputType.HeaderRows = mlngHeaderRows

		On Error GoTo 0
		'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.DataArray. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mobjOutputType.DataArray(mstrArray, mcolColumns, mcolStyles, mcolMerges)

	End Function


	'Public Function DataGrid(objNewValue As SSDBGrid)
	'
	'  Dim strArray() As String
	'  Dim lngGridGrp As Long
	'  Dim lngGridCol As Long
	'  Dim lngGridRow As Long
	'  Dim lngGrp As Long
	'  Dim lngCol As Long
	'  Dim lngRow As Long
	'  Dim blnGroupHeaders As Boolean
	'
	'  On Local Error GoTo LocalErr
	'
	'  If mblnPrintData Then
	'    mobjOutputType.DataGrid objNewValue
	'    Exit Function
	'  End If
	'
	'
	'  With objNewValue
	'
	'    ResetMerges
	'    blnGroupHeaders = (.GroupHeaders And .Groups.Count > 0)
	'
	'    .Redraw = False
	'
	'    'Get count of visible columns
	'    lngCol = 0
	'    For lngGridCol = 0 To .Cols - 1
	'      If .Columns(lngGridCol).Visible Then
	'        lngCol = lngCol + 1
	'
	'        'Check if this is a header column...
	'        If .Columns(lngGridCol).ButtonsAlways Then
	'          If mlngHeaderCols = lngCol - 1 Then
	'            mlngHeaderCols = lngCol
	'          End If
	'        End If
	'      End If
	'    Next
	'    ReDim Preserve strArray(lngCol - 1, .Rows + IIf(blnGroupHeaders, 1, 0))
	'
	'    'GROUP HEADERS
	'    lngCol = 0
	'    If blnGroupHeaders Then
	'      mlngHeaderRows = 2
	'      For lngGridGrp = 0 To .Groups.Count - 1
	'        If .Groups(lngGridGrp).Visible Then
	'          strArray(lngCol, 0) = .Groups(lngGridGrp).Caption
	'          AddMerge lngCol, 0, lngCol + .Groups(lngGridGrp).Columns.Count - 1, 0
	'          lngCol = lngCol + .Groups(lngGridGrp).Columns.Count
	'        End If
	'      Next
	'    End If
	'
	'    'COLUMN HEADERS
	'    lngCol = 0
	'    lngRow = IIf(blnGroupHeaders, 1, 0)
	'    For lngGridCol = 0 To .Cols - 1
	'      If .Columns(lngGridCol).Visible Then
	'        strArray(lngCol, lngRow) = .Columns(lngGridCol).Caption
	'        lngCol = lngCol + 1
	'      End If
	'    Next
	'
	'    'DATA ROWS
	'    For lngGridRow = 0 To .Rows - 1
	'      lngCol = 0
	'      lngRow = lngRow + 1
	'      For lngGridCol = 0 To .Cols - 1
	'
	'        If .Columns(lngGridCol).Visible Then
	'          strArray(lngCol, lngRow) = .Columns(lngGridCol).CellText(.AddItemBookmark(lngGridRow))
	'          lngCol = lngCol + 1
	'        End If
	'
	'      Next
	'
	'    Next
	'    .Redraw = True
	'
	'  End With
	'
	'  DataArray strArray()
	'
	'Exit Function
	'
	'LocalErr:
	'  mstrErrorMessage = Err.Description
	'
	'End Function

	Public Sub Complete()

		If mstrErrorMessage = vbNullString Then
			On Error Resume Next
			'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.Complete. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mobjOutputType.Complete()
		End If
		ClearUp()

	End Sub

	Public Sub ClearUp()
		On Error Resume Next
		'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.ClearUp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mobjOutputType.ClearUp()
	End Sub


	Private Function CheckEmailAttachment(ByRef strExt As String) As Object

		Dim lngFound As Integer

		If mstrEmailAttachAs <> vbNullString Then
			lngFound = InStrRev(mstrEmailAttachAs, ".")
			If lngFound = 0 Then
				mstrEmailAttachAs = mstrEmailAttachAs & "." & strExt
			End If
		End If

	End Function


	Public Function SendEmail(ByRef strAttachment As String) As Boolean

		Dim objOutputEmail As clsGeneral
		'Dim strAddress() As String
		'Dim lngCount As Long

		On Error GoTo LocalErr

		If gblnEmailSystemPermission = False Then
			mstrErrorMessage = "You do not have permission to use email groups."
			SendEmail = False
			Exit Function
		End If


		'  Screen.MousePointer = vbHourglass
		'
		'  If Trim(Replace(mstrEmailAddresses, ";", "")) = vbNullString Then
		'    mstrErrorMessage = "Error sending email (invalid email address)"
		'    SendEmail = False
		'    Exit Function
		'  End If
		'
		'  frmEmailSel.MAPISignon
		'  If frmEmailSel.MAPISession1.SessionID <> 0 Then
		'    With frmEmailSel.MAPIMessages1
		'      .Compose
		'
		'      strAddress = Split(mstrEmailAddresses, ";")
		'
		'      For lngCount = 0 To UBound(strAddress)
		'        If Trim(strAddress(lngCount)) <> vbNullString Then
		'          .RecipIndex = .RecipCount
		'          .RecipAddress = Trim(strAddress(lngCount))
		'          .RecipType = mapToList
		'          .ResolveName
		'        End If
		'      Next
		'
		'
		'      .MsgSubject = mstrEmailSubject
		'      .MsgNoteText = " "
		'      If strAttachment <> "" Then
		'        .AttachmentPosition = 0
		'        .AttachmentType = 0
		'        .AttachmentPathName = strAttachment
		'        .AttachmentName = mstrEmailAttachAs
		'      End If
		'      .Send False
		'    End With
		'  End If
		'  'frmEmailSel.MAPIsignoff

		objOutputEmail = New clsGeneral

		'TODO email stuff
	'	mstrErrorMessage = objOutputEmail.SendEmailFromClientUsingMAPI(mstrEmailAddresses, "", "", mstrEmailSubject, "", strAttachment, False)
		'UPGRADE_NOTE: Object objOutputEmail may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objOutputEmail = Nothing

		'Set frmEmailSel = Nothing
		SendEmail = True

		Exit Function

LocalErr:
		mstrErrorMessage = "Error sending email" & IIf(Err.Description <> vbNullString, " (" & Err.Description & ")", vbNullString)
		'On Error Resume Next
		'frmEmailSel.MAPIsignoff
		'Set frmEmailSel = Nothing
		SendEmail = False

	End Function


	Public Function GetTempFileName(ByRef strFilename As String) As String

		Dim strTempFileName As String

		On Error GoTo LocalErr

		'Get temp path
		strTempFileName = Space(1024)
		Call GetTempPath(1024, strTempFileName)
		strTempFileName = GetTmpFName()
		If InStr(strTempFileName, Chr(0)) > 0 Then
			strTempFileName = Left(strTempFileName, InStr(strTempFileName, Chr(0)) - 1)
		End If

		'temp path + "\" + file name
		If strFilename <> vbNullString Then
			strFilename = Left(strTempFileName, InStrRev(strTempFileName, "\")) & Mid(strFilename, InStrRev(strFilename, "\") + 1)
		Else
			strFilename = strTempFileName
		End If

		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If Dir(strFilename) <> vbNullString Then
			Kill(strFilename)
		End If

		GetTempFileName = strFilename

		Exit Function

LocalErr:
		mstrErrorMessage = Err.Description

	End Function


	Public Sub ShowFormats(ByRef blnData As Boolean, ByRef blnCSV As Boolean, ByRef blnHTML As Boolean, ByRef blnWord As Boolean, ByRef blnExcel As Boolean, ByRef blnChart As Boolean, ByRef blnPivot As Boolean)

		mblnData = blnData
		mblnCSV = blnCSV
		mblnHTML = blnHTML
		mblnWord = blnWord
		mblnExcel = blnExcel
		mblnChart = blnChart
		mblnPivot = blnPivot

	End Sub


	'Public Property Get cboPageBreak() As ComboBox
	'Set cboPageBreak = mfrmOutput.cboPageBreak
	'  mblnPageRange = True
	'End Property


	Public Function SetOptions(ByRef blnPrompt As Boolean, ByRef lngFormat As Integer, ByRef blnScreen As Boolean, ByRef blnPrinter As Boolean, ByRef strPrinterName As String, ByRef blnSave As Boolean, ByRef lngSaveExisting As Integer, ByRef blnEmail As Boolean, ByRef strEmailAddresses As String, ByRef strEmailSubject As String, ByRef strEmailAttachAs As String, ByRef strFilename As String) As Boolean
		Dim Printer As New Printing.PrinterSettings

		Dim blnCancelled As Boolean
		'Dim lngSaveExisting As Long

		On Error GoTo LocalErr


		blnCancelled = False

		'mlngEmailGroupID = lngEmailAddr
		mstrEmailAddresses = strEmailAddresses
		mstrEmailSubject = strEmailSubject
		mstrEmailAttachAs = strEmailAttachAs


		mlngFormat = lngFormat
		Select Case mlngFormat
			Case OutputFormats.fmtDataOnly
				mobjOutputType = New clsOutputGrid

			Case OutputFormats.fmtCSV, OutputFormats.fmtFixedLengthFile
				mobjOutputType = New clsOutputCSV
				CheckEmailAttachment("csv")

			Case OutputFormats.fmtHTML
				mobjOutputType = New clsOutputHTML
				CheckEmailAttachment("htm")

			Case OutputFormats.fmtWordDoc
				mobjOutputType = New clsOutputWord
				CheckEmailAttachment("doc")

			Case OutputFormats.fmtExcelWorksheet
				mobjOutputType = New clsOutputExcel
				CheckEmailAttachment("xls")

			Case OutputFormats.fmtExcelGraph
				mobjOutputType = New clsOutputExcel
				'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.Chart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mobjOutputType.Chart = True
				CheckEmailAttachment("xls")

			Case OutputFormats.fmtExcelPivotTable
				mobjOutputType = New clsOutputExcel
				'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.PivotTable. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mobjOutputType.PivotTable = True
				CheckEmailAttachment("xls")

		End Select


		'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.Parent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Me. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mobjOutputType.Parent = Me
		'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.Screen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mobjOutputType.Screen = blnScreen
		'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.DestPrinter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mobjOutputType.DestPrinter = blnPrinter
		'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.Save. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mobjOutputType.Save = blnSave
		'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.SaveExisting. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mobjOutputType.SaveExisting = lngSaveExisting
		'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.Email. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mobjOutputType.Email = blnEmail
		'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.FileName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mobjOutputType.FileName = strFilename

		mblnPrintData = (mlngFormat = OutputFormats.fmtDataOnly And blnPrinter)

		If strPrinterName = "<Default Printer>" Then
			mstrPrinterName = printer.PrinterName
		Else
			mstrPrinterName = strPrinterName
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.PrinterName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mobjOutputType.PrinterName = mstrPrinterName

		If blnPrinter And Printers.Count = 0 Then
			mstrErrorMessage = "Unable to print as no printers are installed."
		End If


		'MH20040209 Fault 8024
		If Not ValidPrinter(mstrPrinterName) Then
			mstrErrorMessage = "This definition is set to output to printer " & mstrPrinterName & " which is not set up on your PC."
		End If


		SetOptions = (mstrErrorMessage = vbNullString And Not blnCancelled)

		Exit Function

LocalErr:
		If mstrErrorMessage = vbNullString Then
			mstrErrorMessage = Err.Description
		End If

	End Function


	Public Sub SetPrinter()
		Dim Printer As New Printing.PrinterSettings

		Dim objDefPrinter As cSetDfltPrinter

		If mstrPrinterName <> "<Default Printer>" Then
			mstrDefaultPrinter = Printer.PrinterName
			objDefPrinter = New cSetDfltPrinter
			objDefPrinter.SetPrinterAsDefault(mstrPrinterName)
			'UPGRADE_NOTE: Object objDefPrinter may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objDefPrinter = Nothing
		End If

	End Sub

	Public Sub ResetDefaultPrinter()

		Dim objDefPrinter As cSetDfltPrinter

		If mstrPrinterName <> "<Default Printer>" Then
			objDefPrinter = New cSetDfltPrinter
			objDefPrinter.SetPrinterAsDefault(mstrDefaultPrinter)
			'UPGRADE_NOTE: Object objDefPrinter may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objDefPrinter = Nothing
		End If

	End Sub


	Public Function SettingOptions(ByRef strSettingWordTemplate As String, ByRef strSettingExcelTemplate As String, ByRef blnSettingExcelGridlines As Boolean, ByRef blnSettingExcelHeaders As Boolean, ByRef blnSettingExcelOmitSpacerRow As Boolean, ByRef blnSettingExcelOmitSpacerCol As Boolean, ByRef blnSettingAutoFitCols As Boolean, ByRef blnSettingLandscape As Boolean, ByRef blnEmailSystemPermission As Boolean) As Boolean

		gstrSettingWordTemplate = strSettingWordTemplate
		gstrSettingExcelTemplate = strSettingExcelTemplate
		gblnSettingExcelGridlines = blnSettingExcelGridlines
		gblnSettingExcelHeaders = blnSettingExcelHeaders
		gblnSettingExcelOmitSpacerRow = blnSettingExcelOmitSpacerRow
		gblnSettingExcelOmitSpacerCol = blnSettingExcelOmitSpacerCol
		gblnSettingAutoFitCols = blnSettingAutoFitCols
		gblnSettingLandscape = blnSettingLandscape
		gblnEmailSystemPermission = blnEmailSystemPermission

	End Function


	Public Function SettingLocations(ByRef lngSettingTitleCol As Integer, ByRef lngSettingTitleRow As Integer, ByRef lngSettingDataCol As Integer, ByRef lngSettingDataRow As Integer) As Boolean

		glngSettingTitleCol = lngSettingTitleCol
		glngSettingTitleRow = lngSettingTitleRow
		glngSettingDataCol = lngSettingDataCol
		glngSettingDataRow = lngSettingDataRow

	End Function


	Public Function SettingTitle(ByRef blnSettingTitleGridlines As Boolean, ByRef blnSettingTitleBold As Boolean, ByRef blnSettingTitleUnderline As Boolean, ByRef lngSettingTitleBackcolour As Integer, ByRef lngSettingTitleForecolour As Integer, ByRef lngSettingTitleBackcolour97 As Integer, ByRef lngSettingTitleForecolour97 As Integer) As Boolean

		gblnSettingTitleGridlines = blnSettingTitleGridlines
		gblnSettingTitleBold = blnSettingTitleBold
		gblnSettingTitleUnderline = blnSettingTitleUnderline
		glngSettingTitleBackcolour = lngSettingTitleBackcolour
		glngSettingTitleForecolour = lngSettingTitleForecolour
		glngSettingTitleBackcolour97 = lngSettingTitleBackcolour97
		glngSettingTitleForecolour97 = lngSettingTitleForecolour97

	End Function


	Public Function SettingHeading(ByRef blnSettingHeadingGridlines As Boolean, ByRef blnSettingHeadingBold As Boolean, ByRef blnSettingHeadingUnderline As Boolean, ByRef lngSettingHeadingBackcolour As Integer, ByRef lngSettingHeadingForecolour As Integer, ByRef lngSettingHeadingBackcolour97 As Integer, ByRef lngSettingHeadingForecolour97 As Integer) As Boolean

		gblnSettingHeadingGridlines = blnSettingHeadingGridlines
		gblnSettingHeadingBold = blnSettingHeadingBold
		gblnSettingHeadingUnderline = blnSettingHeadingUnderline
		glngSettingHeadingBackcolour = lngSettingHeadingBackcolour
		glngSettingHeadingForecolour = lngSettingHeadingForecolour
		glngSettingHeadingBackcolour97 = lngSettingHeadingBackcolour97
		glngSettingHeadingForecolour97 = lngSettingHeadingForecolour97

	End Function


	Public Function SettingData(ByRef blnSettingDataGridlines As Boolean, ByRef blnSettingDataBold As Boolean, ByRef blnSettingDataUnderline As Boolean, ByRef lngSettingDataBackcolour As Integer, ByRef lngSettingDataForecolour As Integer, ByRef lngSettingDataBackcolour97 As Integer, ByRef lngSettingDataForecolour97 As Integer) As Boolean

		gblnSettingDataGridlines = blnSettingDataGridlines
		gblnSettingDataBold = blnSettingDataBold
		gblnSettingDataUnderline = blnSettingDataUnderline
		glngSettingDataBackcolour = lngSettingDataBackcolour
		glngSettingDataForecolour = lngSettingDataForecolour
		glngSettingDataBackcolour97 = lngSettingDataBackcolour97
		glngSettingDataForecolour97 = lngSettingDataForecolour97

	End Function

	Public Function ArrayDim(ByRef lngCol As Integer, ByRef lngRow As Integer) As Boolean
		ReDim mstrArray(lngCol, lngRow)
'				ReDim mstrArray(50, 12)
	End Function

	Public Function ArrayReDim() As Boolean
		ReDim Preserve mstrArray(UBound(mstrArray, 1), UBound(mstrArray, 2) + 1)
	End Function

	Public Function ArrayAddTo(ByRef lngCol As Integer, ByRef lngRow As Integer, ByRef strInput As Object) As Boolean
		'UPGRADE_WARNING: Couldn't resolve default property of object strInput. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mstrArray(lngCol, lngRow) = strInput
	End Function

	Public Function KillFile(ByRef strFilename As String) As Boolean

		On Error GoTo LocalErr

		Kill(strFilename)
		Return True

LocalErr:
		mstrErrorMessage = "Error overwriting file '" & strFilename & "'" & IIf(Err.Description <> vbNullString, vbCrLf & "(" & Err.Description & ")", vbNullString)
		'mstrErrorMessage = "Cannot access read-only document '" & Mid(strFileName, InStrRev(strFileName, "\") + 1) & "'."
		KillFile = False

	End Function


	Public Function GetSequentialNumberedFile(ByVal strFilename As String) As String

		Dim lngFound As Integer
		Dim lngCount As Integer

		lngCount = 2
		lngFound = InStrRev(strFilename, ".")
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Do While Dir(Left(strFilename, lngFound - 1) & "(" & CStr(lngCount) & ")" & Mid(strFilename, lngFound)) <> vbNullString
			lngCount = lngCount + 1
		Loop
		GetSequentialNumberedFile = Left(strFilename, lngFound - 1) & "(" & CStr(lngCount) & ")" & Mid(strFilename, lngFound)

	End Function

	Private Function ValidPrinter(ByRef strName As String) As Boolean

		Dim objPrinter As Printer
		Dim blnFound As Boolean

		If strName <> vbNullString And strName <> "<Default Printer>" Then
			blnFound = False
			For Each objPrinter In Printers
				If objPrinter.DeviceName = strName Then
					blnFound = True
					Exit For
				End If
			Next objPrinter
		Else
			blnFound = True
		End If

		ValidPrinter = blnFound

	End Function

	Public Function GetSaveAsFormat(ByRef strFilename As String) As String
		GetSaveAsFormat = GetSaveAsFormat2(strFilename, mstrSaveAsValues)
	End Function
End Class