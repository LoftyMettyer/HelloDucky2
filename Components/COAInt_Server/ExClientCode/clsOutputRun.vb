Option Strict Off
Option Explicit On

Imports System.Collections.Generic
Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.ExClientCode
Imports HR.Intranet.Server.Enums
Imports System.Net.Mail
Imports System.Net.Mime
Imports System.IO
Imports HR.Intranet.Server.Metadata

Public Class clsOutputRun
	Inherits BaseForDMI

	Private mobjOutputType As Object
	Private mcolStyles As Collection
	Private mcolStylesNineGridBox As Collection
	Private mcolMerges As Collection
	Private mcolColumns As List(Of Column)
	Private mstrErrorMessage As String

	Private mlngFormat As OutputFormats
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
	Private mstrArrayNineBoxGrid(,) As String

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

	Private mblnSummaryReport As Boolean

	Private _axisLabelsAsArray As ArrayList
	Public Property AxisLabelsAsArray As ArrayList
		Get
			Return _axisLabelsAsArray
		End Get
		Set(value As ArrayList)
			_axisLabelsAsArray = value
		End Set
	End Property

	Public GeneratedFile As String

	Public IntersectionType As IntersectionType = IntersectionType.Total

	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()

		'Set mfrmOutput = New frmOutputOptions
		mcolColumns = New List(Of Column)

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


		'Nine box grid styles
		mcolStylesNineGridBox = New Collection

		objStyle = New clsOutputStyle
		objStyle.Name = "Title"

		mcolStylesNineGridBox.Add(objStyle, objStyle.Name)
		
		objStyle = New clsOutputStyle
		objStyle.Name = "Heading"

		mcolStylesNineGridBox.Add(objStyle, objStyle.Name)
		
		objStyle = New clsOutputStyle
		With objStyle
			.Name = "HeadingCols"
			.StartCol = 0
			.StartRow = 0
			.Gridlines = gblnSettingHeadingGridlines
		End With

		mcolStylesNineGridBox.Add(objStyle, objStyle.Name)
		
		objStyle = New clsOutputStyle
		With objStyle
			.Name = "Data"
			.StartCol = glngSettingDataCol
			.StartRow = glngSettingDataRow
			.Gridlines = gblnSettingDataGridlines
		End With

		mcolStylesNineGridBox.Add(objStyle, objStyle.Name)

		objStyle = Nothing
	End Sub

	Public Sub AddColumn(Heading As String, DataType As ColumnDataType, Decimals As Integer, ThousandSeparator As Boolean)

		Dim objColumn As New Column With {
					.Name = Heading,
					.DataType = DataType,
					.Decimals = Decimals,
					.Use1000Separator = ThousandSeparator}
		mcolColumns.Add(objColumn)

	End Sub

	Public Function AddStyle(strType As String, lngStartCol As Integer, lngStartRow As Integer, lngEndCol As Integer, lngEndRow As Integer, Optional lngBackCol As Object = Nothing _
													, Optional lngForeCol As Object = Nothing, Optional blnBold As Object = Nothing, Optional blnUnderline As Object = Nothing _
													, Optional blnGridLines As Object = Nothing, Optional lngBackCol97 As Object = Nothing, Optional lngForeCol97 As Object = Nothing) As Boolean

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

	Public Sub ResetColumns()
		'UPGRADE_NOTE: Object mcolColumns may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mcolColumns = Nothing
		mcolColumns = New List(Of Column)
	End Sub

	Public Sub ResetStyles()
		'UPGRADE_NOTE: Object mcolStyles may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mcolStyles = Nothing
		InitialiseStyles()
	End Sub

	'Public Function ResetMerges() As Object
	'	'UPGRADE_NOTE: Object mcolMerges may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	'	mcolMerges = Nothing
	'	mcolMerges = New Collection
	'End Function

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
		Return mobjOutputType.GetFile(Me, mcolStyles)
	End Function

	Public Function AddPage(strDefTitle As String, Optional mstrSheetName As String = "") As Object

		'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.AddPage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mobjOutputType.AddPage(strDefTitle, mstrSheetName, mcolStyles)
	End Function

	Public WriteOnly Property SizeColumnsIndependently() As Boolean
		Set(ByVal Value As Boolean)
			mblnSizeColumnsIndependently = Value
		End Set
	End Property

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

			'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.FileDelimiter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mobjOutputType.FileDelimiter = Value
		End Set
	End Property

	Public WriteOnly Property EncloseInQuotes() As Boolean
		Set(ByVal Value As Boolean)
			'clsOutputCSV only

			'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.EncloseInQuotes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mobjOutputType.EncloseInQuotes = Value
		End Set
	End Property

	Public WriteOnly Property ApplyStyles() As Boolean
		Set(ByVal Value As Boolean)

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

	Public WriteOnly Property SummaryReport() As Boolean
		Set(ByVal Value As Boolean)
			mobjOutputType.SummaryReport = Value
		End Set
	End Property

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

	Public Sub DataArray()

		If UBound(mstrArray, 2) < 1 Then
			mstrErrorMessage = "No data to output."
			Exit Sub
		End If

		'Not all classes support all properties so catch any errors...
		'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.SizeColumnsIndependently. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mobjOutputType.SizeColumnsIndependently = mblnSizeColumnsIndependently
		'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.HeaderCols. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mobjOutputType.HeaderCols = mlngHeaderCols
		'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.HeaderRows. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mobjOutputType.HeaderRows = mlngHeaderRows

		mobjOutputType.IntersectionType = IntersectionType

		'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.DataArray. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mobjOutputType.DataArray(mstrArray, mcolColumns, mcolStyles, mcolMerges)

	End Sub

	Public Sub DataArrayNineBoxGrid()
		If UBound(mstrArrayNineBoxGrid, 2) < 1 Then
			mstrErrorMessage = "No data to output."
			Exit Sub
		End If

		mobjOutputType.SizeColumnsIndependently = mblnSizeColumnsIndependently

		mobjOutputType.HeaderCols = mlngHeaderCols

		mobjOutputType.HeaderRows = mlngHeaderRows

		mobjOutputType.IntersectionType = IntersectionType


		mobjOutputType.DataArrayNineBoxGrid(mstrArrayNineBoxGrid, mcolColumns, mcolStylesNineGridBox, mcolMerges)
	End Sub

	Public Sub Complete()

		If mstrErrorMessage = vbNullString Then

			'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.Complete. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mobjOutputType.Complete()
		End If
		ClearUp()

	End Sub

	Public Sub ClearUp()

		'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.ClearUp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mobjOutputType.ClearUp()
	End Sub

	Private Function CheckEmailAttachment(strExt As String) As Object

		Dim lngFound As Integer

		If mstrEmailAttachAs <> vbNullString Then
			lngFound = InStrRev(mstrEmailAttachAs, ".")
			If lngFound = 0 Then
				mstrEmailAttachAs = mstrEmailAttachAs & "." & strExt
			End If
		End If

	End Function

	Public Function SendEmail(strAttachment As String) As Boolean

		'If Not Permissions.GetByKey("EMAILGROUPS_VIEW") Then
		'	mstrErrorMessage = "You do not have permission to use email groups."
		'	SendEmail = False
		'	Exit Function
		'End If

		If Trim(Replace(mstrEmailAddresses, ";", "")) = vbNullString Then
			mstrErrorMessage = "Error sending email (invalid email address)"
			SendEmail = False
			Exit Function
		End If

		SendMailWithAttachment(strAttachment, mstrEmailAddresses, mstrEmailAttachAs)

		Return True

	End Function

	'' The following example sends a binary file as an e-mail attachment.
	Public Shared Sub SendMailWithAttachment(strAttachment As String, recipientList As String, mstrEmailAttachAs As String)

		Dim message As New MailMessage()
		message.Subject = "OpenHR Report"

		If recipientList.Contains(";") = True Then
			Dim aRecipientList = Split(recipientList, ";")

			For iLoop = 0 To UBound(aRecipientList) - 1
				message.To.Add(aRecipientList(iLoop))
			Next
		Else
			message.To.Add(recipientList)
		End If

		Dim fileName As String = strAttachment
		' Get the file stream for the error log.
		' Requires the System.IO namespace.
		Dim fs As New FileStream(fileName, FileMode.Open, FileAccess.Read)
		message.Body = "Your report is attached."
		' Make a contentType indicating that the file is octet
		Dim ct As New ContentType(MediaTypeNames.Application.Octet)
		' Attach the file stream to the e-mail message.
		Dim data As New Attachment(fs, ct)
		Dim disposition As ContentDisposition = data.ContentDisposition
		' Suggest a file name for the attachment.
		disposition.FileName = mstrEmailAttachAs
		' Add the attachment to the message.
		message.Attachments.Add(data)
		' Send the message.
		' The smtpClient settings come from web.config
		Dim client As New SmtpClient()

		Try
			client.Send(message)
		Catch ex As Exception
			' Console.WriteLine("Exception caught in SendErrorLog: {0}", ex.ToString())
		End Try
		data.Dispose()
		' Close the log file.
		fs.Close()

	End Sub

	Public Function GetTempFileName(strFilename As String) As String

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

	Public Sub ShowFormats(blnData As Boolean, blnCSV As Boolean, blnHTML As Boolean, blnWord As Boolean, blnExcel As Boolean, blnChart As Boolean, blnPivot As Boolean)

		mblnData = blnData
		mblnCSV = blnCSV
		mblnHTML = blnHTML
		mblnWord = blnWord
		mblnExcel = blnExcel
		mblnChart = blnChart
		mblnPivot = blnPivot

	End Sub

	Public Function SetOptions(blnPrompt As Boolean, lngFormat As OutputFormats, blnScreen As Boolean, blnPrinter As Boolean, strPrinterName As String _
															, blnSave As Boolean, lngSaveExisting As Integer, blnEmail As Boolean, strEmailAddresses As String, strEmailSubject As String _
															, strEmailAttachAs As String, strDownloadExtension As String) As Boolean
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
			Case OutputFormats.DataOnly
				mobjOutputType = New clsOutputGrid
				GeneratedFile = Path.GetTempFileName.Replace(".tmp", ".txt")

			Case OutputFormats.CSV, OutputFormats.FixedLengthFile
				mobjOutputType = New clsOutputCSV
				CheckEmailAttachment("csv")
				GeneratedFile = Path.GetTempFileName.Replace(".tmp", ".csv")

			Case OutputFormats.HTML
				mobjOutputType = New clsOutputHTML
				CheckEmailAttachment("htm")
				GeneratedFile = Path.GetTempFileName.Replace(".tmp", ".htm")

			Case OutputFormats.WordDoc
				mobjOutputType = New clsOutputWord
				CheckEmailAttachment("doc")
				GeneratedFile = Path.GetTempFileName.Replace(".tmp", ".doc")

			Case OutputFormats.ExcelWorksheet
				mobjOutputType = New clsOutputExcel
				CheckEmailAttachment("xls")
				GeneratedFile = Path.GetTempFileName.Replace(".tmp", ".xlsx")

			Case OutputFormats.ExcelGraph
				mobjOutputType = New clsOutputExcel
				'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.Chart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mobjOutputType.Chart = True
				CheckEmailAttachment("xls")
				GeneratedFile = Path.GetTempFileName.Replace(".tmp", ".xlsx")

			Case OutputFormats.ExcelPivotTable
				mobjOutputType = New clsOutputExcel
				'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.PivotTable. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mobjOutputType.PivotTable = True
				CheckEmailAttachment("xls")
				GeneratedFile = Path.GetTempFileName.Replace(".tmp", ".xlsx")


		End Select

		mobjOutputType.UserName = _login.Username

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
		mobjOutputType.FileName = GeneratedFile
		mobjOutputType.DownloadExtension = strDownloadExtension

		mblnPrintData = (mlngFormat = OutputFormats.DataOnly And blnPrinter)

		If strPrinterName = "<Default Printer>" Then
			mstrPrinterName = Printer.PrinterName
		Else
			mstrPrinterName = strPrinterName
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object mobjOutputType.PrinterName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mobjOutputType.PrinterName = mstrPrinterName

		'If blnPrinter And Printers.Count = 0 Then
		'	mstrErrorMessage = "Unable to print as no printers are installed."
		'End If


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
		'Dim Printer As New Printing.PrinterSettings

		'If mstrPrinterName <> "<Default Printer>" Then
		'	mstrDefaultPrinter = Printer.PrinterName
		'	objDefPrinter = New cSetDfltPrinter
		'	objDefPrinter.SetPrinterAsDefault(mstrPrinterName)
		'	'UPGRADE_NOTE: Object objDefPrinter may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		'	objDefPrinter = Nothing
		'End If

	End Sub

	Public Sub SettingOptions(strSettingWordTemplate As String, strSettingExcelTemplate As String, blnSettingExcelGridlines As Boolean, blnSettingExcelHeaders As Boolean _
														 , blnSettingExcelOmitSpacerRow As Boolean, blnSettingExcelOmitSpacerCol As Boolean, blnSettingAutoFitCols As Boolean _
														 , blnSettingLandscape As Boolean, blnEmailSystemPermission As Boolean)

		gstrSettingWordTemplate = strSettingWordTemplate
		gstrSettingExcelTemplate = ""	'strSettingExcelTemplate
		gblnSettingExcelGridlines = blnSettingExcelGridlines
		gblnSettingExcelHeaders = blnSettingExcelHeaders
		gblnSettingExcelOmitSpacerRow = blnSettingExcelOmitSpacerRow
		gblnSettingExcelOmitSpacerCol = blnSettingExcelOmitSpacerCol
		gblnSettingAutoFitCols = blnSettingAutoFitCols
		gblnSettingLandscape = blnSettingLandscape
		gblnEmailSystemPermission = blnEmailSystemPermission

	End Sub

	Public Sub SettingLocations(lngSettingTitleCol As Integer, lngSettingTitleRow As Integer, lngSettingDataCol As Integer, lngSettingDataRow As Integer)

		glngSettingTitleCol = lngSettingTitleCol
		glngSettingTitleRow = lngSettingTitleRow
		glngSettingDataCol = lngSettingDataCol
		glngSettingDataRow = lngSettingDataRow

	End Sub

	Public Sub SettingTitle(blnSettingTitleGridlines As Boolean, blnSettingTitleBold As Boolean, blnSettingTitleUnderline As Boolean _
																, lngSettingTitleBackcolour As Integer, lngSettingTitleForecolour As Integer, lngSettingTitleBackcolour97 As Integer _
																, lngSettingTitleForecolour97 As Integer)

		gblnSettingTitleGridlines = blnSettingTitleGridlines
		gblnSettingTitleBold = blnSettingTitleBold
		gblnSettingTitleUnderline = blnSettingTitleUnderline
		glngSettingTitleBackcolour = lngSettingTitleBackcolour
		glngSettingTitleForecolour = lngSettingTitleForecolour
		glngSettingTitleBackcolour97 = lngSettingTitleBackcolour97
		glngSettingTitleForecolour97 = lngSettingTitleForecolour97

	End Sub

	Public Sub SettingHeading(blnSettingHeadingGridlines As Boolean, blnSettingHeadingBold As Boolean, blnSettingHeadingUnderline As Boolean _
														 , lngSettingHeadingBackcolour As Integer, lngSettingHeadingForecolour As Integer, lngSettingHeadingBackcolour97 As Integer _
														 , lngSettingHeadingForecolour97 As Integer)

		gblnSettingHeadingGridlines = blnSettingHeadingGridlines
		gblnSettingHeadingBold = blnSettingHeadingBold
		gblnSettingHeadingUnderline = blnSettingHeadingUnderline
		glngSettingHeadingBackcolour = lngSettingHeadingBackcolour
		glngSettingHeadingForecolour = lngSettingHeadingForecolour
		glngSettingHeadingBackcolour97 = lngSettingHeadingBackcolour97
		glngSettingHeadingForecolour97 = lngSettingHeadingForecolour97

	End Sub

	Public Sub SettingData(blnSettingDataGridlines As Boolean, blnSettingDataBold As Boolean, blnSettingDataUnderline As Boolean, lngSettingDataBackcolour As Integer _
													, lngSettingDataForecolour As Integer, lngSettingDataBackcolour97 As Integer, lngSettingDataForecolour97 As Integer)

		gblnSettingDataGridlines = blnSettingDataGridlines
		gblnSettingDataBold = blnSettingDataBold
		gblnSettingDataUnderline = blnSettingDataUnderline
		glngSettingDataBackcolour = lngSettingDataBackcolour
		glngSettingDataForecolour = lngSettingDataForecolour
		glngSettingDataBackcolour97 = lngSettingDataBackcolour97
		glngSettingDataForecolour97 = lngSettingDataForecolour97

	End Sub

	Public Sub ArrayDim(lngCol As Integer, lngRow As Integer)
		ReDim mstrArray(lngCol, lngRow)
		ReDim mstrArrayNineBoxGrid(lngCol, lngRow)
	End Sub

	Public Sub ArrayReDim()
		ReDim Preserve mstrArray(UBound(mstrArray, 1), UBound(mstrArray, 2) + 1)
		ReDim Preserve mstrArrayNineBoxGrid(UBound(mstrArray, 1), UBound(mstrArray, 2) + 1)
	End Sub

	Public Sub ArrayAddTo(lngCol As Integer, lngRow As Integer, strInput As String)
		mstrArray(lngCol, lngRow) = strInput
	End Sub

	Public Sub ArrayAddToNineBoxGrid(lngCol As Integer, lngRow As Integer, desc As String, value As String, colour As String)
		mstrArrayNineBoxGrid(lngCol, lngRow) = desc & "¦" & value & "|" & colour
	End Sub

	Public Function KillFile(strFilename As String) As Boolean

		On Error GoTo LocalErr

		Kill(strFilename)
		Return True

LocalErr:
		mstrErrorMessage = "Error overwriting file '" & strFilename & "'" & IIf(Err.Description <> vbNullString, vbCrLf & "(" & Err.Description & ")", vbNullString).ToString()
		'mstrErrorMessage = "Cannot access read-only document '" & Mid(strFileName, InStrRev(strFileName, "\") + 1) & "'."
		KillFile = False

	End Function

	Public Function GetSequentialNumberedFile(strFilename As String) As String

		Dim lngFound As Integer
		Dim lngCount As Integer

		lngCount = 2
		lngFound = InStrRev(strFilename, ".")
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Do While Dir(Left(strFilename, lngFound - 1) & "(" & CStr(lngCount) & ")" & Mid(strFilename, lngFound)) <> vbNullString
			lngCount = lngCount + 1
		Loop
		Return Left(strFilename, lngFound - 1) & "(" & CStr(lngCount) & ")" & Mid(strFilename, lngFound)

	End Function

	Private Function ValidPrinter(strName As String) As Boolean
		Return True
		'TODO Implement printing
		'Dim objPrinter As New Printing.PrinterSettings
		'Dim blnFound As Boolean

		'If strName <> vbNullString And strName <> "<Default Printer>" Then
		'	blnFound = False
		'	For Each objPrinter In Printers
		'		If objPrinter.PrinterName = strName Then
		'			blnFound = True
		'			Exit For
		'		End If
		'	Next objPrinter
		'Else
		'	blnFound = True
		'End If

		'ValidPrinter = blnFound

	End Function

End Class