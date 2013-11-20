Option Strict Off
Option Explicit On
Friend Class clsOutputCSV

	Private mobjParent As clsOutputRun

	Private mblnScreen As Boolean
	Private mblnPrinter As Boolean
	Private mstrPrinterName As String
	Private mblnSave As Boolean
	Private mlngSaveExisting As Integer
	Private mblnEmail As Boolean
	Private mstrFileName As String
	Private mstrDelim As String
	Private mblnQuotes As Boolean

	Private mstrErrorMessage As String

	Public WriteOnly Property Screen() As Boolean
		Set(ByVal Value As Boolean)
			mblnScreen = Value
		End Set
	End Property

	Public WriteOnly Property DestPrinter() As Boolean
		Set(ByVal Value As Boolean)
			mblnPrinter = Value
		End Set
	End Property

	Public WriteOnly Property PrinterName() As String
		Set(ByVal Value As String)
			mstrPrinterName = Value
		End Set
	End Property

	Public WriteOnly Property Save() As Boolean
		Set(ByVal Value As Boolean)
			mblnSave = Value
		End Set
	End Property


	Public Property SaveExisting() As Integer
		Get
			SaveExisting = mlngSaveExisting
		End Get
		Set(ByVal Value As Integer)
			mlngSaveExisting = Value
		End Set
	End Property

	Public WriteOnly Property Email() As Boolean
		Set(ByVal Value As Boolean)
			mblnEmail = Value
		End Set
	End Property

	Public WriteOnly Property FileName() As String
		Set(ByVal Value As String)
			mstrFileName = Value
		End Set
	End Property

	Public WriteOnly Property FileDelimiter() As String
		Set(ByVal Value As String)
			mstrDelim = Value
		End Set
	End Property

	Public WriteOnly Property EncloseInQuotes() As Boolean
		Set(ByVal Value As Boolean)
			mblnQuotes = Value
		End Set
	End Property

	Public WriteOnly Property Parent() As clsOutputRun
		Set(ByVal Value As clsOutputRun)
			mobjParent = Value
		End Set
	End Property

	Public ReadOnly Property ErrorMessage() As String
		Get
			ErrorMessage = mstrErrorMessage
		End Get
	End Property


	Public Function GetFile(ByRef objParent As clsOutputRun, ByRef colStyles As Collection) As Boolean

		Dim strTempFileName As String
		Dim lngFound As Integer
		Dim lngCount As Integer

		On Error GoTo LocalErr


		If mstrFileName = vbNullString Then
			mstrFileName = objParent.GetTempFileName((mobjParent.EmailAttachAs))
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If Dir(mstrFileName) <> vbNullString Then
				objParent.KillFile(mstrFileName)
			End If


			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		ElseIf Dir(mstrFileName) <> vbNullString Then
			'Check if file already exists...

			Select Case mlngSaveExisting
				Case 0 'Overwrite
					Kill(mstrFileName)

				Case 1 'Do not overwrite (fail)
					mstrErrorMessage = "File already exists"

				Case 2 'Add Sequential number to file
					mstrFileName = mobjParent.GetSequentialNumberedFile(mstrFileName)


				Case 3 'Append to existing file


				Case 4 'Create new worksheet within existing workbook...
					'N/A (EXCEL ONLY)

			End Select

		End If

		GetFile = (mstrErrorMessage = vbNullString)

		Exit Function

LocalErr:
		GetFile = False
		mstrErrorMessage = Err.Description

	End Function


	'Public Sub DataGrid(objNewValue As SSDBGrid, colColumns As Collection, colStyles As Collection)
	'
	'  Dim strOutput As String
	'  Dim lngGridCol As Long
	'  Dim lngGridRow As Long
	'
	'  Open mstrFileName For Append As #1
	'
	'  With objNewValue
	'
	'    .Redraw = False
	'
	'    'If not appending to existing file then add headers...
	'    If LOF(1) = 0 Then
	'      strOutput = vbNullString
	'      For lngGridCol = 0 To .Cols - 1
	'        If .Columns(lngGridCol).Visible Then
	'          strOutput = _
	''            IIf(strOutput <> vbNullString, strOutput & ",", "") & _
	''            .Columns(lngGridCol).Caption
	'        End If
	'      Next
	'      Print #1, strOutput
	'    End If
	'
	'    For lngGridRow = 0 To .Rows - 1
	'      strOutput = vbNullString
	'      .Bookmark = .AddItemBookmark(lngGridRow)
	'      For lngGridCol = 0 To .Cols - 1
	'        If .Columns(lngGridCol).Visible Then
	'          strOutput = _
	''            IIf(strOutput <> vbNullString, strOutput & ",", "") & _
	''            .Columns(lngGridCol).CellText(.Bookmark)
	'        End If
	'      Next
	'      Print #1, strOutput
	'    Next
	'
	'    .Redraw = True
	'  End With
	'
	'  Close
	'
	'End Sub


	Public Sub DataArray(ByRef strArray(,) As String, ByRef colColumns As Collection, ByRef colStyles As Collection, ByRef colMerges As Collection)

		Dim objColumn As clsColumn
		Dim strOutput As String
		Dim strTemp As String
		Dim lngGridCol As Integer
		Dim lngGridRow As Integer

		On Error GoTo LocalErr

		FileOpen(1, mstrFileName, OpenMode.Append)

		For lngGridRow = 0 To UBound(strArray, 2)
			strOutput = vbNullString
			For lngGridCol = 0 To UBound(strArray, 1)

				strTemp = strArray(lngGridCol, lngGridRow)

				If mblnQuotes Then
					If InStr(strTemp, ",") > 0 Or InStr(strTemp, Chr(34)) Then
						strTemp = Chr(34) & Replace(strTemp, New String(Chr(34), 1), New String(Chr(34), 2)) & Chr(34)
					End If
				End If

				strOutput = IIf(lngGridCol > 0, strOutput & mstrDelim, "") & strTemp

				'      If gobjProgress.Visible And gobjProgress.Cancelled Then
				'        mstrErrorMessage = "Cancelled by User"
				'        Close
				'        Exit Sub
				'      End If

			Next

			PrintLine(1, strOutput)
		Next

		FileClose()

		Exit Sub

LocalErr:
		mstrErrorMessage = Err.Description

	End Sub


	Public Sub Complete()

		On Error GoTo LocalErr

		If mstrErrorMessage <> vbNullString Then
			Exit Sub
		End If

		'EMAIL
		If mblnEmail Then
			mstrErrorMessage = "Error sending email"
			mobjParent.SendEmail(mstrFileName)
		End If

		mstrErrorMessage = vbNullString

		Exit Sub

LocalErr:
		mstrErrorMessage = mstrErrorMessage & IIf(Err.Description <> vbNullString, " (" & Err.Description & ")", vbNullString)

	End Sub

	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		mstrDelim = ","
		mblnQuotes = True
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub

	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		FileClose()
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub

	Public Sub ClearUp()
		FileClose()
	End Sub
End Class