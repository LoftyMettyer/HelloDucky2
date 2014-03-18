Option Strict Off
Option Explicit On

Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Enums
Imports VB = Microsoft.VisualBasic
Friend Class clsOutputHTML
	Inherits BaseOutputFormat

	Private mobjParent As clsOutputRun

	Private mstrDefTitle As String
	Private mblnScreen As Boolean
	Private mblnPrinter As Boolean
	Private mstrPrinterName As String
	Private mblnSave As Boolean
	Private mlngSaveExisting As Integer
	Private mblnEmail As Boolean
	Private mstrFileName As String
	Private mlngHeaderRows As Integer
	Private mlngHeaderCols As Integer
	Private mblnHeaderVertical As Boolean
	Private mblnApplyStyles As Boolean

	'Private mstrHTMLTemplate As String

	Private mstrHTMLOutput As String
	Private mlngPageCount As Integer
	Private mstrErrorMessage As String

	Public Sub ClearUp()
		FileClose()
	End Sub

	'Public Function RecordProfilePage(pfrmRecProfile As Form, _
	''  piPageNumber As Integer, _
	''  pcolStyles As Collection)
	'  ' Output the record profile page to Excel.
	'
	'  On Error GoTo ErrorTrap
	'  gobjErrorStack.PushStack "clsOutputHTML.RecordProfilePage()"
	'
	'  Dim fOK As Boolean
	'  Dim iLoop As Integer
	'  Dim iLoop2 As Integer
	'  Dim iLoop3 As Integer
	'  Dim iLoop4 As Integer
	'  Dim ctlTemp As Control
	'  Dim varBookmark As Variant
	'  Dim fGridPreceded As Boolean
	'  Dim fGridFollowed As Boolean
	'  Dim sTemp As String
	'  Dim fPhotoDone As Boolean
	'  Dim objRecProfTable As clsRecordProfileTabDtl
	'  Dim sTempName As String
	'  Dim iTemp As Integer
	'  Dim fIsHeading As Boolean
	'  Dim sSubFolderPath As String
	'
	'  Const RECPROFFOLLOWONCORRECTION = 10
	'
	'  Const COLUMN_ISHEADING = "IsHeading"
	'  Const COLUMN_ISPHOTO = "IsPhoto"
	'  Const PHOTOSTYLESET = "PhotoSS_"
	'
	'  fOK = True
	'
	'  sSubFolderPath = mstrFileName
	'  If InStrRev(sSubFolderPath, ".") > 0 Then
	'    sSubFolderPath = Left(sSubFolderPath, InStrRev(sSubFolderPath, ".") - 1)
	'  End If
	'  sSubFolderPath = sSubFolderPath & "_files"
	'
	'  For Each ctlTemp In pfrmRecProfile.Controls
	'    If ctlTemp.Container Is pfrmRecProfile.picOutput(piPageNumber) Then
	'      '
	'      ' LABEL control
	'      '
	'      If TypeOf ctlTemp Is Label Then
	'        If ctlTemp.Visible Then
	'          mstrHTMLOutput = mstrHTMLOutput & _
	''            "<BR>" & _
	''            HTMLText("SPAN", ctlTemp.Caption, pcolStyles("Title")) & _
	''            "<BR>"
	'        End If
	'      End If
	'
	'      '
	'      ' GRID control
	'      '
	'      If TypeOf ctlTemp Is SSDBGrid Then
	'        Set objRecProfTable = pfrmRecProfile.Definition.Item(ctlTemp.Tag)
	'
	'        ' Check if this grid is preceded or followed IMMEDIATELY by other grids.
	'        ' ie. if this grid is part of a group of grids that are used to display
	'        ' data (including pictures) vertically.
	'        ' NB. Grids have only one row height value. To display pictures with their
	'        ' own row height, we actually put them in their own grid, and position this grid
	'        ' IMMEDIATELY after the normal data grid. Subsequent data is put in its own grid
	'        ' IMMEDIATELY after the picture's grid.
	'        ' This is what's meant by 'following' and 'preceding' grids.
	'        fGridFollowed = False
	'        fGridPreceded = False
	'
	'        If ctlTemp.Index > 1 Then
	'          If (ctlTemp.Container = pfrmRecProfile.grdOutput(ctlTemp.Index - 1).Container) And _
	''            (ctlTemp.Top = (pfrmRecProfile.grdOutput(ctlTemp.Index - 1).Top + pfrmRecProfile.grdOutput(ctlTemp.Index - 1).Height - RECPROFFOLLOWONCORRECTION)) Then
	'
	'            fGridPreceded = True
	'          End If
	'        End If
	'
	'        If ctlTemp.Index < pfrmRecProfile.grdOutput.Count - 1 Then
	'          If (ctlTemp.Container = pfrmRecProfile.grdOutput(ctlTemp.Index + 1).Container) And _
	''            ((ctlTemp.Top + ctlTemp.Height - RECPROFFOLLOWONCORRECTION) = pfrmRecProfile.grdOutput(ctlTemp.Index + 1).Top) Then
	'
	'            fGridFollowed = True
	'          End If
	'        End If
	'
	'        If Not fGridPreceded Then
	'          mstrHTMLOutput = mstrHTMLOutput & _
	''              "<CENTER><TABLE border=1 cellspacing=0 cellpadding=1" & _
	''              " bordercolordark=" & HexColour(pcolStyles("Data").BackCol) & _
	''              " bordercolorlight=000000>" & vbCrLf
	'        End If
	'
	'        ' Send the group headers to the HTML document.
	'        If (ctlTemp.GroupHeaders) And (ctlTemp.ColumnHeaders) Then
	'          mstrHTMLOutput = mstrHTMLOutput & "<TR>"
	'          For iLoop = 0 To ctlTemp.Groups.Count - 1
	'            If ctlTemp.Groups(iLoop).Visible Then
	'              iTemp = 0
	'              For iLoop2 = 0 To ctlTemp.Groups(iLoop).Columns.Count - 1
	'                If (ctlTemp.Columns(iLoop).Visible) Then
	'                  iTemp = iTemp + 1
	'                End If
	'              Next iLoop2
	'
	'              ' Send the group header to the HTML document.
	'              mstrHTMLOutput = mstrHTMLOutput & _
	''                  HTMLText("TD", ctlTemp.Groups(iLoop).Caption, pcolStyles("Heading"), IIf(iTemp > 1, " COLSPAN=" & CStr(iTemp), ""))
	'            End If
	'          Next iLoop
	'          mstrHTMLOutput = mstrHTMLOutput & _
	''              "</TR>" & vbCrLf
	'        End If
	'
	'        ' Send the column headers to the HTML document.
	'        If (ctlTemp.ColumnHeaders) Then
	'          mstrHTMLOutput = mstrHTMLOutput & "<TR>"
	'          For iLoop = 0 To ctlTemp.Columns.Count - 1
	'            If (ctlTemp.Columns(iLoop).Visible) Then
	'              ' Send the column header to the HTML document.
	'              mstrHTMLOutput = mstrHTMLOutput & _
	''                  HTMLText("TD", ctlTemp.Columns(iLoop).Caption, pcolStyles("Heading"), vbNullString)
	'            End If
	'          Next iLoop
	'          mstrHTMLOutput = mstrHTMLOutput & "</TR>" & vbCrLf
	'        End If
	'
	'        ' Send data rows and columns to Excel.
	'        For iLoop = 0 To ctlTemp.Rows - 1
	'          varBookmark = ctlTemp.AddItemBookmark(iLoop)
	'
	'          mstrHTMLOutput = mstrHTMLOutput & "<TR>"
	'
	'          For iLoop2 = 0 To ctlTemp.Columns.Count - 1
	'            If ctlTemp.Columns(iLoop2).Visible Then
	'              ' Send the text or picture to the HTML document.
	'              fPhotoDone = False
	'              If (ctlTemp.TagVariant = COLUMN_ISPHOTO) And _
	''                (ctlTemp.Columns(iLoop2).Style <> 4) Then
	'
	'                For iLoop3 = 0 To ctlTemp.Columns.Count - 1
	'                  If ctlTemp.Columns(iLoop3).Visible Then
	'                    sTemp = PHOTOSTYLESET & CStr(iLoop3 + 1)
	'
	'                    For iLoop4 = 0 To ctlTemp.StyleSets.Count - 1
	'                      If ctlTemp.StyleSets(iLoop4).Name = sTemp Then
	'                        sTempName = GetTmpFNameInFolder(sSubFolderPath)
	'
	'                        SavePicture ctlTemp.StyleSets(iLoop4).Picture, sTempName
	'                        mstrHTMLOutput = mstrHTMLOutput & _
	''                              "<TD" & _
	''                              " bgcolor=" & HexColour(pcolStyles("Data").BackCol) & ">" & _
	''                              "<IMG alt="""" src=""file://" & sTempName & """>" & _
	''                              "</TD>"
	'
	'                        fPhotoDone = True
	'                        Exit For
	'                      End If
	'                    Next iLoop4
	'                  End If
	'
	'                  If fPhotoDone Then
	'                    Exit For
	'                  End If
	'                Next iLoop3
	'              End If
	'
	'              If (Not fPhotoDone) And _
	''                ctlTemp.Columns(iLoop2).TagVariant = COLUMN_ISPHOTO Then
	'
	'                sTemp = PHOTOSTYLESET & CStr(iLoop2 + 1) & "_" & ctlTemp.Columns(CStr(objRecProfTable.IDPosition)).Value
	'
	'                For iLoop4 = 0 To ctlTemp.StyleSets.Count - 1
	'                  If ctlTemp.StyleSets(iLoop4).Name = sTemp Then
	'                    sTempName = GetTmpFNameInFolder(sSubFolderPath)
	'                    SavePicture ctlTemp.StyleSets(iLoop4).Picture, sTempName
	'                    mstrHTMLOutput = mstrHTMLOutput & _
	''                          "<TD" & _
	''                          " bgcolor=" & HexColour(pcolStyles("Data").BackCol) & ">" & _
	''                          "<IMG alt="""" src=""file://" & sTempName & """>" & _
	''                          "</TD>"
	'                    fPhotoDone = True
	'                    Exit For
	'                  End If
	'                Next iLoop4
	'              End If
	'
	'              If Not fPhotoDone Then
	'                ' Send the data to the HTML document.
	'                varBookmark = ctlTemp.AddItemBookmark(iLoop)
	'
	'                fIsHeading = ((ctlTemp.Columns(iLoop2).Style = 4) And _
	''                  (ctlTemp.Columns(iLoop2).TagVariant <> COLUMN_ISPHOTO)) Or _
	''                  (ctlTemp.Columns(iLoop2).StyleSet = "Separator")
	'                If Not ctlTemp.ColumnHeaders Then
	'                  If ctlTemp.Columns(COLUMN_ISHEADING).CellText(varBookmark) = "1" Then
	'                    fIsHeading = True
	'                  End If
	'                End If
	'
	'                mstrHTMLOutput = mstrHTMLOutput & _
	''                    HTMLText("TD", ctlTemp.Columns(iLoop2).CellText(varBookmark), IIf(fIsHeading, pcolStyles("HeadingCols"), pcolStyles("Data")), vbNullString)
	'              End If
	'
	'            End If
	'          Next iLoop2
	'
	'          mstrHTMLOutput = mstrHTMLOutput & "</TR>" & vbCrLf
	'        Next iLoop
	'
	'        If Not fGridFollowed Then
	'          mstrHTMLOutput = mstrHTMLOutput & _
	''              "</TABLE></CENTER><BR>" & vbCrLf
	'        End If
	'      End If
	'    End If
	'  Next ctlTemp
	'  Set ctlTemp = Nothing
	'
	'TidyUpAndExit:
	'  gobjErrorStack.PopStack
	'  RecordProfilePage = fOK
	'  Exit Function
	'
	'ErrorTrap:
	'  gobjErrorStack.HandleError
	'  fOK = False
	'  Resume TidyUpAndExit
	'
	'End Function

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

	Public Function GetFile(ByRef objParent As clsOutputRun, ByRef colSytles As Collection) As Boolean

		Dim strTempFileName As String
		Dim lngFound As Integer
		Dim lngCount As Integer

		Dim strLineInput As String
		Dim blnAppending As Boolean
		Dim blnFound As Boolean

		On Error GoTo LocalErr


		blnAppending = False

		'Just in case we are emailing but not saving...
		If Not mblnSave Then
			If mblnEmail Then
				mstrFileName = objParent.GetTempFileName((mobjParent.EmailAttachAs))
				'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				If Dir(mstrFileName) <> vbNullString Then
					objParent.KillFile(mstrFileName)
				End If
			End If

		Else

			'Check if file already exists...
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If Dir(mstrFileName) <> vbNullString And mstrFileName <> vbNullString Then

				Select Case mlngSaveExisting
					Case 0 'Overwrite
						If Not objParent.KillFile(mstrFileName) Then
							GetFile = False
							Exit Function
						End If

					Case 1 'Do not overwrite (fail)
						mstrErrorMessage = "File already exists."
						GetFile = False
						Exit Function

					Case 2 'Add Sequential number to file
						mstrFileName = mobjParent.GetSequentialNumberedFile(mstrFileName)

					Case 3 'Append to existing file


					Case 4 'Create new worksheet within existing workbook...
						'N/A (EXCEL ONLY)

				End Select

			End If

			If OpenFile Then
				blnAppending = ((LOF(1) > 0) And (mlngSaveExisting = 3))
			End If

		End If

		'  If mstrHTMLTemplate <> vbNullString Then
		'    If Dir(mstrHTMLTemplate) <> vbNullString Then
		'      Open mstrHTMLTemplate For Input As #2
		'
		'      blnFound = False
		'      Do While Not blnFound And Not EOF(2)
		'        Input #2, strLineInput
		'        blnFound = (LCase(strLineInput) = "<hrprodata>")
		'        If Not blnFound Then
		'          mstrHTMLOutput = mstrHTMLOutput & strLineInput
		'        End If
		'      Loop
		'    End If
		'
		'  Else
		If Not blnAppending Then
			mstrHTMLOutput = mstrHTMLOutput & "<HTML><BODY>"
		End If
		mstrHTMLOutput = mstrHTMLOutput & "<FONT face=Verdana size=2>"

		'  End If

		GetFile = (mstrErrorMessage = vbNullString)

		Exit Function

LocalErr:
		mstrErrorMessage = Err.Description
		GetFile = False

	End Function


	Public Sub AddPage(ByRef strDefTitle As String, ByRef mstrSheetName As String, ByRef colStyles As Collection)

		Dim strTitle As String

		On Error GoTo LocalErr

		mstrHTMLOutput = mstrHTMLOutput & "<BR>"

		mlngPageCount = mlngPageCount + 1
		If mlngPageCount = 1 Then
			mstrDefTitle = strDefTitle
			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strTitle = HTMLText("SPAN", strDefTitle, colStyles.Item("Title"))
			mstrHTMLOutput = mstrHTMLOutput & "<CENTER>" & strTitle & "</CENTER>"
		Else
			mstrHTMLOutput = mstrHTMLOutput & "<HR>" & vbCrLf
		End If

		If mstrSheetName <> vbNullString And mobjParent.PageTitles Then
			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mstrHTMLOutput = mstrHTMLOutput & "<BR>" & HTMLText("SPAN", mstrSheetName, colStyles.Item("Title")) & "<BR>"
		End If

		mstrHTMLOutput = mstrHTMLOutput & "<BR>"

		Exit Sub

LocalErr:
		mstrErrorMessage = Err.Description

	End Sub


	Public Sub DataArray(ByRef strArray(,) As String, ByRef colColumns As Collection, ByRef colStyles As Collection, ByRef colMerges As Collection)

		Dim strOutput As String
		Dim lngGridCol As Integer
		Dim lngGridRow As Integer

		On Error GoTo LocalErr


		With colStyles.Item("Title")
			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.StartCol = glngSettingTitleCol
			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.StartRow = glngSettingTitleRow
			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles(Title).EndCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.EndCol = .StartCol
			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles(Title).EndRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.EndRow = .StartRow
		End With

		With colStyles.Item("Heading")
			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.StartCol = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.StartRow = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles(Heading).EndCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.EndCol = UBound(strArray, 1)
			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().EndRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.EndRow = mlngHeaderRows - 1
		End With

		If mlngHeaderCols > 0 Then
			With colStyles.Item("HeadingCols")
				'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.StartCol = 0
				'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.StartRow = 0
				'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().EndCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.EndCol = mlngHeaderCols - 1
				'UPGRADE_WARNING: Couldn't resolve default property of object colStyles(HeadingCols).EndRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.EndRow = UBound(strArray, 2)
			End With
		End If

		With colStyles.Item("Data")
			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.StartCol = mlngHeaderCols
			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.StartRow = mlngHeaderRows
			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles(Data).EndCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.EndCol = UBound(strArray, 1)
			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles(Data).EndRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.EndRow = UBound(strArray, 2)
		End With


		'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().BackCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mstrHTMLOutput = mstrHTMLOutput & "<CENTER><TABLE border=1 cellspacing=0 cellpadding=1" & " bordercolordark=" & HexColour(colStyles.Item("Data").BackCol) & " bordercolorlight=000000>" & vbCrLf

		'strOutput = vbNullString
		'lngGridCol = 0
		'For Each objColumn In colColumns
		'  strOutput = strOutput & _
		''      HTMLText("TD", objColumn.Heading, colStyles("Heading"))
		'  lngGridCol = lngGridCol + 1
		'Next
		'mstrHTMLOutput = mstrHTMLOutput & "<TR>" & strOutput & "</TR>" & vbCrLf

		For lngGridRow = 0 To UBound(strArray, 2)
			strOutput = vbNullString
			For lngGridCol = 0 To UBound(strArray, 1)
				strOutput = strOutput & CheckHTMLText("TD", strArray(lngGridCol, lngGridRow), lngGridCol, lngGridRow, colStyles, colMerges, colColumns)

				'      If gobjProgress.Visible And gobjProgress.Cancelled Then
				'        mstrErrorMessage = "Cancelled by user."
				'        Exit Sub
				'      End If

			Next
			mstrHTMLOutput = mstrHTMLOutput & "<TR>" & strOutput & "</TR>" & vbCrLf
		Next

		mstrHTMLOutput = mstrHTMLOutput & "</TABLE></CENTER>" & vbCrLf

		Exit Sub

LocalErr:
		mstrErrorMessage = Err.Description

	End Sub


	Private Function OpenFile() As Boolean

		On Error GoTo LocalErr

		FileOpen(1, mstrFileName, OpenMode.Append)

		OpenFile = True

		Exit Function

LocalErr:
		mstrErrorMessage = "Error saving file <" & mstrFileName & ">" & IIf(Err.Description <> vbNullString, vbCrLf & " (" & Err.Description & ")", vbNullString)
		OpenFile = False

	End Function


	Private Function CheckHTMLText(ByRef strTag As String, ByRef strInput As String, ByRef lngCol As Integer, ByRef lngRow As Integer, ByRef colStyles As Collection, ByRef colMerges As Collection, ByRef colColumns As Collection) As String

		Dim objStyle As clsOutputStyle
		Dim objMerge As clsOutputStyle
		Dim objTemp As clsOutputStyle
		Dim strTemp As String

		On Error GoTo LocalErr


		For Each objTemp In colStyles
			If (objTemp.StartCol <= lngCol And objTemp.EndCol >= lngCol) And (objTemp.StartRow <= lngRow And objTemp.EndRow >= lngRow) Then
				objStyle = objTemp
			End If
		Next objTemp

		For Each objTemp In colMerges
			If (objTemp.StartCol <= lngCol And objTemp.EndCol >= lngCol) And (objTemp.StartRow <= lngRow And objTemp.EndRow >= lngRow) Then
				objMerge = objTemp
			End If
		Next objTemp


		'UPGRADE_WARNING: Couldn't resolve default property of object colColumns(lngCol + 1).DataType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Select Case colColumns.Item(lngCol + 1).DataType
			Case SQLDataType.sqlNumeric, SQLDataType.sqlInteger
				strTemp = " ALIGN=Right "
			Case SQLDataType.sqlBoolean
				strTemp = " ALIGN=Center "
			Case Else
				strTemp = " ALIGN=Left "
		End Select

		If objMerge Is Nothing Then
			'No merging required...
			CheckHTMLText = HTMLText(strTag, strInput, objStyle, strTemp)

		Else
			If objMerge.StartCol = lngCol And objMerge.StartRow = lngRow Then
				'This is the top left of a merged range...
				strTemp = strTemp & " COLSPAN=" & CStr(objMerge.EndCol - objMerge.StartCol + 1) & " ROWSPAN=" & CStr(objMerge.EndRow - objMerge.StartRow + 1) & " VALIGN=Top"
				CheckHTMLText = HTMLText(strTag, strInput, objStyle, strTemp)
			Else
				'part of a range so don't bother...
				CheckHTMLText = vbNullString
			End If
		End If

		Exit Function

LocalErr:
		mstrErrorMessage = Err.Description

	End Function


	Private Function HTMLText(ByRef strTag As String, ByRef strInput As String, ByRef objStyle As clsOutputStyle, Optional ByRef strExtraTag As String = "") As String

		Dim strOutput As String

		On Error GoTo LocalErr

		strOutput = Replace(strInput, "<", "&LT;")
		strOutput = Replace(strOutput, ">", "&GT;")
		strOutput = Replace(strOutput, vbTab, "</TD><TD>")
		strOutput = Replace(strOutput, " ", "&nbsp;")

		If strOutput = vbNullString Then
			strOutput = "&nbsp;"
		End If

		If mblnApplyStyles And Not (objStyle Is Nothing) Then
			If objStyle.CenterText Then
				strOutput = "<CENTER>" & strOutput & "</CENTER>"
			End If
			If objStyle.Bold Then
				strOutput = "<B>" & strOutput & "</B>"
			End If
			If objStyle.Underline Then
				strOutput = "<U>" & strOutput & "</U>"
			End If

			strOutput = "<" & strTag & strExtraTag & " bgcolor=" & HexColour((objStyle.BackCol)) & ">" & "<FONT size=2 color=" & HexColour((objStyle.ForeCol)) & ">" & strOutput & "</FONT>" & "</" & strTag & ">"
		End If

		HTMLText = strOutput

		Exit Function

LocalErr:
		mstrErrorMessage = Err.Description

	End Function


	Private Function HexColour(ByRef lngColour As Object) As String
		'UPGRADE_WARNING: Couldn't resolve default property of object lngColour. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Mod has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		HexColour = Right("0" & Hex(lngColour Mod 256), 2) & Right("0" & Hex(lngColour \ 256), 2) & Right("0" & Hex(lngColour \ 65536), 2)
	End Function

	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		'mstrHTMLTemplate = GetUserSetting("Output", "HTMLTemplate", vbNullString)
		mblnApplyStyles = True
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

	Public Sub Complete()

		Dim strLineInput As String
		Dim blnOK As Boolean


		If mstrErrorMessage <> vbNullString Then
			FileClose()
			Exit Sub
		End If

		On Error GoTo LocalErr

		blnOK = True

		'  If mstrHTMLTemplate <> vbNullString Then
		'    If Dir(mstrHTMLTemplate) <> vbNullString Then
		'      Do While Not EOF(2)
		'        Input #2, strLineInput
		'        mstrHTMLOutput = mstrHTMLOutput & strLineInput
		'      Loop
		'    End If
		'
		'  Else
		mstrHTMLOutput = mstrHTMLOutput & "</CENTER><BR><HR>" & vbCrLf & "Created on " & VB6.Format(Now, DateFormat() & " hh:nn") & " by " & UserName & vbCrLf & "</FONT></BODY></HTML>"
		'  End If

		If mblnSave Then
			PrintLine(1, mstrHTMLOutput)
		End If

		FileClose()


		'EMAIL
		If mblnEmail Then
			mstrErrorMessage = "Error sending email"

			If mblnSave Then
				mobjParent.SendEmail(mstrFileName)
			Else
				'mstrFileName = GetTmpFName
				'mstrFileName = Left(mstrFileName, Len(mstrFileName) - 3) & "htm"
				FileOpen(1, mstrFileName, OpenMode.Output)
				PrintLine(1, mstrHTMLOutput)
				FileClose()
				mobjParent.SendEmail(mstrFileName)
				Kill(mstrFileName)
			End If

		End If


		If mblnScreen Then
			mstrErrorMessage = "Error displaying HTML"
			blnOK = DisplayInBrowser
		End If

		If blnOK Then
			mstrErrorMessage = vbNullString
		End If

TidyAndExit:

		Exit Sub

LocalErr:
		mstrErrorMessage = mstrErrorMessage & IIf(Err.Description <> vbNullString, " (" & Err.Description & ")", vbNullString)
		Resume TidyAndExit

	End Sub

	Private Function DisplayInBrowser() As Boolean

		Dim IE As System.Windows.Forms.WebBrowser
		Dim dblWait As Double
		Dim dblWait2 As Double
		Dim blnOK As Boolean

		On Error GoTo LocalErr

		blnOK = True
		dblWait = VB.Timer() + 10

		IE = New System.Windows.Forms.WebBrowser

		'JPD 20091007 Fault HRPRO-31, HRPRO-33, HRPRO-34
		' New SHDocVw.InternetExplorer sometimes gets a handle on the existing SSI/DMI browser instance.
		' If this happens, do it again toensure you get a fresh instance.
		If IE.DocumentTitle = "OpenHR Self-service Intranet" Or IE.DocumentTitle = "OpenHR Intranet" Then

			IE = New System.Windows.Forms.WebBrowser
		End If

		If mblnSave Then
			IE.Navigate(New System.URI(mstrFileName))
			Do While IE.IsBusy
				System.Windows.Forms.Application.DoEvents()
			Loop

		Else
RetryDisplay:
			'AE20071129 Fault #12111 / #12112
			'    IE.Navigate ""    'Creates a blank document
			IE.Navigate(New System.URI("about:blank"))
			Do While IE.IsBusy
				System.Windows.Forms.Application.DoEvents()
			Loop

			'Keep trying for 10 seconds then error
			blnOK = False
			Do

				Err.Clear()
				On Error GoTo LocalErr
				'UPGRADE_ISSUE: SHDocVw.InternetExplorer property IE.AddressBar was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
	'			IE.AddressBar = False
				If Not IE.Document.DomDocument Is Nothing Then
					'UPGRADE_WARNING: Couldn't resolve default property of object IE.Document.Title. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					IE.Document.DomDocument.Title = mstrDefTitle
					'UPGRADE_WARNING: Couldn't resolve default property of object IE.Document.Body. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If Not IE.Document.DomDocument.Body Is Nothing Then
						'UPGRADE_WARNING: Couldn't resolve default property of object IE.Document.Body. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						IE.Document.DomDocument.Body.InnerHtml = mstrHTMLOutput
						'UPGRADE_WARNING: Couldn't resolve default property of object IE.Document.Body. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						IE.Document.DomDocument.Body.Bgcolor = "white"
						blnOK = (Err.Number = 0)
					End If
				End If
				System.Windows.Forms.Application.DoEvents()

			Loop While Not blnOK

		End If

		If blnOK Then
			IE.Visible = True
		End If
		'UPGRADE_NOTE: Object IE may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		IE = Nothing

		DisplayInBrowser = blnOK

		Exit Function

LocalErr:
		dblWait2 = VB.Timer() + 2
		Do While dblWait2 > VB.Timer()
			System.Windows.Forms.Application.DoEvents()
		Loop

		If dblWait > VB.Timer() Then
			Err.Clear()
			On Error GoTo LocalErr
			GoTo RetryDisplay
		End If
		DisplayInBrowser = False
		'UPGRADE_NOTE: Object IE may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		IE = Nothing

	End Function


End Class