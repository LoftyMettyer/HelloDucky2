Option Strict Off
Option Explicit On

Imports HR.Intranet.Server.Enums

Friend Class clsOutputExcel

	Private mxlApp As Microsoft.Office.Interop.Excel.Application
	Private mxlWorkBook As Microsoft.Office.Interop.Excel.Workbook
	Private mxlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
	Private mxlTemplateBook As Microsoft.Office.Interop.Excel.Workbook
	Private mxlTemplateSheet As Microsoft.Office.Interop.Excel.Worksheet
	Private mxlFirstSheet As Object
	Private mxlDeleteSheet As Object
	Private mobjParent As clsOutputRun

	Private mlngHeaderRows As Integer
	Private mlngHeaderCols As Integer
	Private mblnHeaderVertical As Boolean
	Private mlngDataCurrentRow As Integer
	Private mlngDataStartRow As Integer
	Private mlngDataStartCol As Integer

	Private mblnScreen As Boolean
	Private mblnPrinter As Boolean
	Private mstrPrinterName As String
	Private mblnSave As Boolean
	Private mlngSaveExisting As Integer
	Private mblnEmail As Boolean
	Private mstrFileName As String
	Private mblnSizeColumnsIndependently As Boolean
	Private mblnApplyStyles As Boolean

	Private mstrSheetMode As String
	Private mblnAppending As Boolean
	Private mlngAppendStartRow As Integer

	Private mstrDefTitle As String
	Private mstrErrorMessage As String

	Private mblnChart As Boolean
	Private mblnPivotTable As Boolean
	'Private mstrIntersectionFormat As String

	Private mstrXLTemplate As String
	Private mblnXLExcelGridlines As Boolean
	Private mblnXLExcelHeaders As Boolean
	Private mblnXLExcelOmitTopRow As Boolean
	Private mblnXLExcelOmitLeftCol As Boolean
	Private mblnXLAutoFitCols As Boolean
	Private mblnXLLandscape As Boolean


	Public Sub ClearUp()

		On Error Resume Next

		'Always close the template...
		If Not mxlTemplateBook Is Nothing Then
			mxlTemplateBook.Saved = True
			mxlTemplateBook.Close()
		End If

		'If error then close the workbook and app...
		If Not mxlWorkBook Is Nothing Then
			mxlWorkBook.Saved = True
			'mxlWorkBook.Close()
		End If
		'If Not mxlApp Is Nothing Then
		'	mxlApp.Quit()
		'End If

		'Reset all references to ensure that Excel closes cleanly...
		'UPGRADE_NOTE: Object mxlTemplateSheet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mxlTemplateSheet = Nothing
		'UPGRADE_NOTE: Object mxlTemplateBook may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mxlTemplateBook = Nothing
		'UPGRADE_NOTE: Object mxlDeleteSheet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mxlDeleteSheet = Nothing
		'UPGRADE_NOTE: Object mxlFirstSheet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mxlFirstSheet = Nothing
		'UPGRADE_NOTE: Object mxlWorkSheet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mxlWorkSheet = Nothing
		'UPGRADE_NOTE: Object mxlWorkBook may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mxlWorkBook = Nothing
		'UPGRADE_NOTE: Object mxlApp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mxlApp = Nothing

	End Sub

	'Public Function RecordProfilePage(pfrmRecProfile As Form, _
	''  piPageNumber As Integer, _
	''  pcolStyles As Collection)
	'  ' Output the record profile page to Excel.
	'
	'  On Error GoTo ErrorTrap
	'  'gobjErrorStack.PushStack "clsOutputExcel.RecordProfilePage()"
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
	'  Dim sTitle As String
	'  Dim sTemp As String
	'  Dim fPhotoDone As Boolean
	'  Dim objRecProfTable As clsRecordProfileTabDtl
	'  Dim sTempName As String
	'  Dim lngBorder As Long
	'  Dim iLastGroup As Integer
	'  Dim colMerges As Collection
	'  Dim objMerge As clsOutputStyle
	'  Dim iGroupStart As Integer
	'  Dim iTemp As Integer
	'  Dim objTemp As clsOutputStyle
	'  Dim iOriginalStyles As Integer
	'  Dim fHasHeadingColumn As Boolean
	'  Dim lngXLCol As Long
	'  Dim lngXLRow As Long
	'  Dim lngXLGridStartRow As Long
	'  Dim lngXLGridEndRow As Long
	'  Dim lngXLGridStartCol As Long
	'  Dim lngXLGridEndCol As Long
	'  Dim objRange As Excel.Range
	'  Dim lngMaxXLCol As Long
	'  Dim lngMaxXLRow As Long
	'  Dim lngMaxWidth As Long
	'  Dim lngCount As Long
	'  Dim alngPictureRows() As Long
	'  Dim alngPictureCols() As Long
	'  Dim fFound As Boolean
	'  Dim lngPictureHeight As Long
	'  Dim lngPictureWidth As Long
	'  Dim iOriginalScaleMode As ScaleModeConstants
	'  Dim iDecPlaces As Integer
	'
	'  Const RECPROFFOLLOWONCORRECTION = 10
	'
	'  Const COLUMN_ISHEADING = "IsHeading"
	'  Const COLUMN_ISPHOTO = "IsPhoto"
	'  Const PHOTOSTYLESET = "PhotoSS_"
	'  Const COLUMN_DECPLACES = "DecPlaces"
	'
	'  fOK = True
	'  sTitle = pfrmRecProfile.Caption
	'  iOriginalStyles = pcolStyles.Count
	'  lngXLCol = 2
	'
	'  'lngXLRow = 4
	'  lngXLRow = mlngDataCurrentRow
	'
	'  lngMaxXLRow = lngXLRow
	'  lngMaxXLCol = lngXLCol
	'
	'  ' Initialise the hidden column sizing label in the preview screen
	'  ' with the same font as the Excel worksheet. We use this label to
	'  ' work out the required column width for pictures that are output to Excel,
	'  ' as we need to translate the picture's width into the number of characters
	'  ' that fit into that width.
	'  iOriginalScaleMode = pfrmRecProfile.ScaleMode
	'  With mxlWorkSheet.Cells(1, 1)
	'    .Font.Bold = pcolStyles("Data").Bold
	'    pfrmRecProfile.lblColumnSizingLabel.Font.Name = .Font.Name
	'    pfrmRecProfile.lblColumnSizingLabel.Font.Size = .Font.Size
	'    pfrmRecProfile.lblColumnSizingLabel.Font.Bold = .Font.Bold
	'  End With
	'
	'  ' Dimension an array to hold info for the Excel rows that contain pictures.
	'  ' Column 1 - Excel row
	'  ' Column 2 - Max picture height
	'  ReDim alngPictureRows(2, 0)
	'
	'  ' Dimension an array to hold info for the Excel columns that contain pictures.
	'  ' Column 1 - Excel col
	'  ' Column 2 - Max picture width
	'  ReDim alngPictureCols(2, 0)
	'
	'  For Each ctlTemp In pfrmRecProfile.Controls
	'    If ctlTemp.Container Is pfrmRecProfile.picOutput(piPageNumber) Then
	'      '
	'      ' LABEL control
	'      '
	'      If TypeOf ctlTemp Is Label Then
	'        If ctlTemp.Visible Then
	'          ' Write the label's caption to Excel.
	'          lngXLCol = 2
	'
	'          mxlWorkSheet.Cells(lngXLRow, lngXLCol).FormulaR1C1 = ctlTemp.Caption
	'
	'          lngMaxXLRow = IIf(lngMaxXLRow < lngXLRow, lngXLRow, lngMaxXLRow)
	'          lngMaxXLCol = IIf(lngMaxXLCol < lngXLCol, lngXLCol, lngMaxXLCol)
	'
	'          Set objRange = mxlWorkSheet.Cells(lngXLRow, lngXLCol)
	'          ApplyStyleToRange objRange, pcolStyles("Title")
	'
	''          With objRange
	''            .Font.Bold = pcolStyles("Title").Bold
	''            .Font.Underline = pcolStyles("Title").Underline
	''            .Font.Color = pcolStyles("Title").ForeCol
	''
	''            If pcolStyles("Title").CenterText Then
	''              .HorizontalAlignment = xlCenter
	''            End If
	''            .VerticalAlignment = xlCenter
	''
	''            If objStyle.Name <> "Title" Then
	''              .Interior.Color = pcolStyles("Title").BackCol
	''              If pcolStyles("Title").Gridlines Then
	''                .BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
	''              Else
	''                .BorderAround xlNone
	''                .Borders(xlDiagonalDown).LineStyle = xlNone
	''                .Borders(xlDiagonalUp).LineStyle = xlNone
	''                .Borders(xlEdgeLeft).LineStyle = xlNone
	''                .Borders(xlEdgeTop).LineStyle = xlNone
	''                .Borders(xlEdgeBottom).LineStyle = xlNone
	''                .Borders(xlEdgeRight).LineStyle = xlNone
	''                .Borders(xlInsideVertical).LineStyle = xlNone
	''                .Borders(xlInsideHorizontal).LineStyle = xlNone
	''              End If
	''            End If
	''          End With
	'
	'          lngXLRow = lngXLRow + 2
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
	'        lngXLCol = 2
	'
	'        If Not fGridPreceded Then
	'          lngXLGridStartRow = lngXLRow
	'          lngXLGridEndRow = -1
	'          lngXLGridStartCol = lngXLCol
	'          lngXLGridEndCol = -1
	'          Set colMerges = New Collection
	'        End If
	'
	'        ' Send the column/group headers to the Excel document.
	'        fHasHeadingColumn = False
	'        iLastGroup = -1
	'        iTemp = lngXLCol
	'        iGroupStart = lngXLCol
	'
	'        For iLoop = 0 To ctlTemp.Columns.Count - 1
	'          If ctlTemp.Columns(iLoop).Name = COLUMN_ISHEADING Then
	'            fHasHeadingColumn = True
	'          End If
	'
	'          If (ctlTemp.Columns(iLoop).Visible) Then
	'            If (ctlTemp.ColumnHeaders) Then
	'              If (ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders) Then
	'                If (iLastGroup <> ctlTemp.Columns(iLoop).Group) Then
	'                  ' Send the group header to the Excel document.
	'                  mxlWorkSheet.Cells(lngXLRow, lngXLCol).FormulaR1C1 = ctlTemp.Groups(ctlTemp.Columns(iLoop).Group).Caption
	'
	'                  ' Remember if the group heading cells need to be merged.
	'                  If iGroupStart < iTemp - 1 Then
	'                    Set objMerge = New clsOutputStyle
	'                    objMerge.StartCol = iGroupStart
	'                    objMerge.StartRow = lngXLRow
	'                    objMerge.EndCol = iTemp - 1
	'                    objMerge.EndRow = lngXLRow
	'
	'                    colMerges.Add objMerge
	'                    Set objMerge = Nothing
	'                  End If
	'
	'                  iGroupStart = iTemp
	'                End If
	'
	'                iLastGroup = ctlTemp.Columns(iLoop).Group
	'              End If
	'
	'              ' Check if the column is a Separator column.
	'              ' If so we'll need to create a 'style' object for it.
	'              If ctlTemp.Columns(iLoop).StyleSet = "Separator" Then
	'                Set objTemp = New clsOutputStyle
	'
	'                With objTemp
	'                  .StartCol = lngXLCol
	'                  .StartRow = lngXLRow + IIf(ctlTemp.ColumnHeaders, 1, 0) + IIf((ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders), 1, 0)
	'                  .EndCol = lngXLCol
	'                  .EndRow = -1
	'
	'                  .BackCol = pcolStyles("HeadingCols").BackCol
	'                  .ForeCol = pcolStyles("HeadingCols").ForeCol
	'                  .Bold = pcolStyles("HeadingCols").Bold
	'                  .Underline = pcolStyles("HeadingCols").Underline
	'                  .Gridlines = pcolStyles("HeadingCols").Gridlines
	'                  .Name = "RECPROFCOL_" & CStr(iLoop)
	'                End With
	'
	'                pcolStyles.Add objTemp
	'                Set objTemp = Nothing
	'              End If
	'
	'              ' Send the column header to the Excel document.
	'              mxlWorkSheet.Cells(lngXLRow + IIf((ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders), 1, 0), lngXLCol).FormulaR1C1 = ctlTemp.Columns(iLoop).Caption
	'              iTemp = iTemp + 1
	'            End If
	'
	'            lngXLGridEndCol = lngXLCol
	'            lngXLCol = lngXLCol + 1
	'          End If
	'        Next iLoop
	'
	'        ' Remember if the group heading cells need to be merged.
	'        If (ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders) Then
	'          If iGroupStart < iTemp - 1 Then
	'            Set objMerge = New clsOutputStyle
	'            objMerge.StartCol = iGroupStart
	'            objMerge.StartRow = lngXLRow
	'            objMerge.EndCol = iTemp - 1
	'            objMerge.EndRow = lngXLRow
	'
	'            colMerges.Add objMerge
	'            Set objMerge = Nothing
	'          End If
	'        End If
	'
	'        lngXLRow = lngXLRow + IIf(ctlTemp.ColumnHeaders, 1, 0) + IIf((ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders), 1, 0)
	'
	'        ' Send data rows and columns to Excel.
	'        For iLoop = 0 To ctlTemp.Rows - 1
	'          varBookmark = ctlTemp.AddItemBookmark(iLoop)
	'
	'          ' Check if the row is a Separator/Heading row.
	'          ' If so we'll need to create a 'style' object for it.
	'          If fHasHeadingColumn Then
	'            If ctlTemp.Columns(COLUMN_ISHEADING).CellText(varBookmark) = "1" Then
	'              Set objTemp = New clsOutputStyle
	'
	'              With objTemp
	'                .StartCol = lngXLGridStartCol + 1
	'                .StartRow = lngXLRow
	'                .EndCol = -1
	'                .EndRow = lngXLRow
	'
	'                .BackCol = pcolStyles("HeadingCols").BackCol
	'                .ForeCol = pcolStyles("HeadingCols").ForeCol
	'                .Bold = pcolStyles("HeadingCols").Bold
	'                .Underline = pcolStyles("HeadingCols").Underline
	'                .Gridlines = pcolStyles("HeadingCols").Gridlines
	'                .Name = "RECPROFROW_" & CStr(iLoop)
	'              End With
	'
	'              pcolStyles.Add objTemp
	'              Set objTemp = Nothing
	'            End If
	'          End If
	'
	'          lngXLCol = 2
	'
	'          For iLoop2 = 0 To ctlTemp.Columns.Count - 1
	'            If ctlTemp.Columns(iLoop2).Visible Then
	'              ' Send the text or picture to Excel.
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
	'                        sTempName = GetTmpFName
	'                        SavePicture ctlTemp.StyleSets(iLoop4).Picture, sTempName
	'                        mxlWorkBook.ActiveSheet.Range(mxlWorkBook.ActiveSheet.Cells(lngXLRow, lngXLCol), mxlWorkBook.ActiveSheet.Cells(lngXLRow, lngXLCol)).Select
	'                        mxlWorkBook.ActiveSheet.Pictures.Insert (sTempName)
	'                        lngPictureHeight = mxlWorkBook.ActiveSheet.Shapes(mxlWorkBook.ActiveSheet.Shapes.Count).Height
	'
	'                        ' Use the hidden column sizing label in the preview form to
	'                        ' translate the picture's width into the number of characters that
	'                        ' fit into that width.
	'                        pfrmRecProfile.ScaleMode = vbPoints
	'                        pfrmRecProfile.lblColumnSizingLabel.Caption = "a"
	'                        Do While pfrmRecProfile.lblColumnSizingLabel.Width < mxlWorkBook.ActiveSheet.Shapes(mxlWorkBook.ActiveSheet.Shapes.Count).Width
	'                          pfrmRecProfile.lblColumnSizingLabel.Caption = pfrmRecProfile.lblColumnSizingLabel.Caption & "a"
	'                        Loop
	'                        lngPictureWidth = Len(pfrmRecProfile.lblColumnSizingLabel.Caption)
	'                        pfrmRecProfile.ScaleMode = iOriginalScaleMode
	'
	'                        Kill sTempName
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
	'                    sTempName = GetTmpFName
	'                    SavePicture ctlTemp.StyleSets(iLoop4).Picture, sTempName
	'                    mxlWorkBook.ActiveSheet.Range(mxlWorkBook.ActiveSheet.Cells(lngXLRow, lngXLCol), mxlWorkBook.ActiveSheet.Cells(lngXLRow, lngXLCol)).Select
	'                    mxlWorkBook.ActiveSheet.Pictures.Insert (sTempName)
	'                    lngPictureHeight = mxlWorkBook.ActiveSheet.Shapes(mxlWorkBook.ActiveSheet.Shapes.Count).Height
	'                    lngPictureWidth = mxlWorkBook.ActiveSheet.Shapes(mxlWorkBook.ActiveSheet.Shapes.Count).Width
	'
	'                    ' Use the hidden column sizing label in the preview form to
	'                    ' translate the picture's width into the number of characters that
	'                    ' fit into that width.
	'                    pfrmRecProfile.ScaleMode = vbPoints
	'                    pfrmRecProfile.lblColumnSizingLabel.Caption = "a"
	'                    Do While pfrmRecProfile.lblColumnSizingLabel.Width < mxlWorkBook.ActiveSheet.Shapes(mxlWorkBook.ActiveSheet.Shapes.Count).Width
	'                      pfrmRecProfile.lblColumnSizingLabel.Caption = pfrmRecProfile.lblColumnSizingLabel.Caption & "a"
	'                    Loop
	'                    lngPictureWidth = Len(pfrmRecProfile.lblColumnSizingLabel.Caption)
	'                    pfrmRecProfile.ScaleMode = iOriginalScaleMode
	'
	'                    Kill sTempName
	'                    fPhotoDone = True
	'                    Exit For
	'                  End If
	'                Next iLoop4
	'              End If
	'
	'              If fPhotoDone Then
	'                ' Remember the height/width of pictures, so that the rows/cols
	'                ' can be sized to fit the pictures later.
	'                fFound = False
	'                For iLoop4 = 1 To UBound(alngPictureRows, 2)
	'                  If alngPictureRows(1, iLoop4) = lngXLRow Then
	'                    alngPictureRows(2, iLoop4) = IIf(alngPictureRows(2, iLoop4) < lngPictureHeight, lngPictureHeight, alngPictureRows(2, iLoop4))
	'                    fFound = True
	'                    Exit For
	'                  End If
	'                Next iLoop4
	'
	'                If Not fFound Then
	'                  ReDim Preserve alngPictureRows(2, UBound(alngPictureRows, 2) + 1)
	'                  alngPictureRows(1, UBound(alngPictureRows, 2)) = lngXLRow
	'                  alngPictureRows(2, UBound(alngPictureRows, 2)) = lngPictureHeight
	'                End If
	'
	'                fFound = False
	'                For iLoop4 = 1 To UBound(alngPictureCols, 2)
	'                  If alngPictureCols(1, iLoop4) = lngXLCol Then
	'                    alngPictureCols(2, iLoop4) = IIf(alngPictureCols(2, iLoop4) < lngPictureWidth, lngPictureWidth, alngPictureCols(2, iLoop4))
	'                    fFound = True
	'                    Exit For
	'                  End If
	'                Next iLoop4
	'
	'                If Not fFound Then
	'                  ReDim Preserve alngPictureCols(2, UBound(alngPictureCols, 2) + 1)
	'                  alngPictureCols(1, UBound(alngPictureCols, 2)) = lngXLCol
	'                  alngPictureCols(2, UBound(alngPictureCols, 2)) = lngPictureWidth
	'                End If
	'              Else
	'                varBookmark = ctlTemp.AddItemBookmark(iLoop)
	'                With mxlWorkSheet.Cells(lngXLRow, lngXLCol)
	'                  .NumberFormat = "@"
	'
	'                  ' Format the cell for the required size & decimals if the column is numeric.
	'                  If (ctlTemp.Columns(iLoop2).Style <> 4) Then
	'                    If ctlTemp.ColumnHeaders Then
	'                      ' Horizontal grid.
	'                      iDecPlaces = IIf(ctlTemp.Columns(iLoop2).TagVariant = COLUMN_ISPHOTO, -1, CInt(ctlTemp.Columns(iLoop2).TagVariant))
	'                    Else
	'                      ' Vertical grid.
	'                      iDecPlaces = CInt(ctlTemp.Columns(COLUMN_DECPLACES).CellText(varBookmark))
	'                    End If
	'
	'                    If iDecPlaces >= 0 Then
	'                      If iDecPlaces > 127 Then iDecPlaces = 127
	'                      .NumberFormat = "0" & IIf(iDecPlaces > 0, "." & String(iDecPlaces, "0"), "")
	'                    End If
	'                  End If
	'
	'                  ' Send the data to the cell.
	'                  .FormulaR1C1 = ctlTemp.Columns(iLoop2).CellText(varBookmark)
	'                End With
	'              End If
	'
	'              lngXLCol = lngXLCol + 1
	'            End If
	'          Next iLoop2
	'
	'          lngXLGridEndRow = lngXLRow
	'          lngXLRow = lngXLRow + 1
	'        Next iLoop
	'
	'        If Not fGridFollowed Then
	'          lngMaxXLRow = IIf(lngMaxXLRow < lngXLGridEndRow, lngXLGridEndRow, lngMaxXLRow)
	'          lngMaxXLCol = IIf(lngMaxXLCol < lngXLGridEndCol, lngXLGridEndCol, lngMaxXLCol)
	'
	'          ' Apply styles to the table.
	'          With pcolStyles("Heading")
	'            .StartCol = lngXLGridStartCol
	'            .StartRow = lngXLGridStartRow
	'            .EndCol = lngXLGridEndCol
	'            .EndRow = IIf(ctlTemp.ColumnHeaders, lngXLGridStartRow + IIf((ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders), 1, 0), -1)
	'          End With
	'
	'          With pcolStyles("HeadingCols")
	'            .StartCol = lngXLGridStartCol
	'            .StartRow = lngXLGridStartRow
	'            .EndCol = IIf(ctlTemp.ColumnHeaders, -1, lngXLGridStartCol)
	'            .EndRow = lngXLGridEndRow
	'          End With
	'
	'          With pcolStyles("Data")
	'            .StartCol = lngXLGridStartCol + IIf(ctlTemp.ColumnHeaders, 0, 1)
	'            .StartRow = lngXLGridStartRow + IIf(ctlTemp.ColumnHeaders, 1, 0) + IIf((ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders), 1, 0)
	'            .EndCol = lngXLGridEndCol
	'            .EndRow = lngXLGridEndRow
	'          End With
	'
	'          With pcolStyles("Title")
	'            .StartCol = -1
	'            .StartRow = -1
	'            .EndCol = -1
	'            .EndRow = -1
	'          End With
	'
	'          ' Set the endRow/endCol values for the separator/heading rows/cols.
	'          For Each objTemp In pcolStyles
	'            If Left(objTemp.Name, 11) = "RECPROFCOL_" Then
	'              objTemp.EndRow = lngXLGridEndRow
	'            End If
	'            If Left(objTemp.Name, 11) = "RECPROFROW_" Then
	'              objTemp.EndCol = lngXLGridEndCol
	'            End If
	'          Next objTemp
	'          Set objTemp = Nothing
	'
	'          ' Apply styles and merging to the table.
	'          If mblnApplyStyles Then
	'            For Each objTemp In pcolStyles
	'              If (objTemp.EndRow >= 0) And _
	''                (objTemp.EndCol >= 0) And _
	''                (objTemp.EndRow >= objTemp.StartRow) And _
	''                (objTemp.EndCol >= objTemp.StartCol) Then
	'
	'                Set objRange = mxlWorkSheet.Range(mxlWorkSheet.Cells(objTemp.StartRow, objTemp.StartCol), mxlWorkSheet.Cells(objTemp.EndRow, objTemp.EndCol))
	'                ApplyStyleToRange objRange, objTemp
	'                objRange.VerticalAlignment = xlCenter
	'              End If
	'            Next
	'
	'            For Each objMerge In colMerges
	'              If (objMerge.EndRow > objMerge.StartRow) Or _
	''                (objMerge.EndCol > objMerge.StartCol) Then
	'
	'                Set objRange = mxlWorkSheet.Range(mxlWorkSheet.Cells(objMerge.StartRow, objMerge.StartCol), mxlWorkSheet.Cells(objMerge.EndRow, objMerge.EndCol))
	'                objRange.MergeCells = True
	'                objRange.VerticalAlignment = xlCenter
	'              End If
	'            Next
	'          End If
	'          Set colMerges = Nothing
	'
	'          ' Get rid of any separator/heading styles that have been created
	'          ' for this table, so that they are not carried over and applied
	'          ' to other tables.
	'          Do While pcolStyles.Count > iOriginalStyles
	'            pcolStyles.Remove pcolStyles.Count
	'          Loop
	'
	'          lngXLRow = lngXLRow + 1
	'        End If
	'      End If
	'    End If
	'  Next ctlTemp
	'  Set ctlTemp = Nothing
	'
	'  With mxlApp.ActiveWindow
	'    .DisplayGridlines = mblnXLExcelGridlines
	'    .DisplayHeadings = mblnXLExcelHeaders
	'  End With
	'
	'  mxlWorkSheet.Range(mxlWorkSheet.Cells(2, 2), mxlWorkSheet.Cells(lngMaxXLRow, lngMaxXLCol)).EntireColumn.AutoFit
	'
	'  ' Size rows/cols with pictures to fit the pictures.
	'  For iLoop4 = 1 To UBound(alngPictureRows, 2)
	'    If mxlWorkSheet.Rows(alngPictureRows(1, iLoop4)).RowHeight < alngPictureRows(2, iLoop4) Then
	'      mxlWorkSheet.Rows(alngPictureRows(1, iLoop4)).RowHeight = alngPictureRows(2, iLoop4)
	'    End If
	'  Next iLoop4
	'
	'  For iLoop4 = 1 To UBound(alngPictureCols, 2)
	'    If mxlWorkSheet.Columns(alngPictureCols(1, iLoop4)).ColumnWidth < alngPictureCols(2, iLoop4) Then
	'      mxlWorkSheet.Columns(alngPictureCols(1, iLoop4)).ColumnWidth = alngPictureCols(2, iLoop4)
	'    End If
	'  Next iLoop4
	'
	'  'Put title in after autofit...
	'  With pcolStyles("Title")
	'    '.StartCol = 2
	'    '.StartRow = 2
	'    '.EndCol = 2
	'    '.EndRow = 2
	'    .StartCol = Val(GetUserSetting("Output", "TitleCol", "3"))
	'    .StartRow = IIf(mlngAppendStartRow > 0, mlngAppendStartRow, Val(GetUserSetting("Output", "TitleRow", "2")))
	'    .EndCol = .StartCol
	'    .EndRow = .StartRow
	'  End With
	'  mxlWorkSheet.Cells(pcolStyles("Title").StartRow, pcolStyles("Title").StartCol).FormulaR1C1 = mstrDefTitle
	'  Set objRange = mxlWorkSheet.Cells(pcolStyles("Title").StartRow, pcolStyles("Title").StartCol)
	'  ApplyStyleToRange objRange, pcolStyles("Title")
	'  objRange.VerticalAlignment = xlCenter
	'
	'  mxlWorkSheet.PageSetup.Orientation = IIf(mblnXLLandscape, xlLandscape, xlPortrait)
	'  mxlWorkSheet.DisplayPageBreaks = False
	'
	'  mxlWorkBook.ActiveSheet.Range("A1").Select
	'  mxlWorkBook.Sheets(1).Select
	'
	'TidyUpAndExit:
	'  'gobjErrorStack.PopStack
	'  RecordProfilePage = fOK
	'  Exit Function
	'
	'ErrorTrap:
	'  'gobjErrorStack.HandleError
	'  fOK = False
	'  Resume TidyUpAndExit
	'
	'End Function



	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()

		mstrXLTemplate = gstrSettingExcelTemplate
		mblnXLExcelGridlines = gblnSettingExcelGridlines
		mblnXLExcelHeaders = gblnSettingExcelHeaders
		mblnXLExcelOmitTopRow = gblnSettingExcelOmitSpacerRow
		mblnXLExcelOmitLeftCol = gblnSettingExcelOmitSpacerCol
		mblnXLAutoFitCols = gblnSettingAutoFitCols
		mblnXLLandscape = gblnSettingLandscape

		mlngDataStartRow = glngSettingDataRow
		mlngDataStartCol = glngSettingDataCol

		mblnSizeColumnsIndependently = False
		mblnApplyStyles = True

	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub

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

	Public WriteOnly Property Chart() As Boolean
		Set(ByVal Value As Boolean)
			mblnChart = Value
		End Set
	End Property

	Public WriteOnly Property PivotTable() As Boolean
		Set(ByVal Value As Boolean)
			mblnPivotTable = Value
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

	Public WriteOnly Property SizeColumnsIndependently() As Boolean
		Set(ByVal Value As Boolean)
			mblnSizeColumnsIndependently = Value
		End Set
	End Property

	Public WriteOnly Property ApplyStyles() As Boolean
		Set(ByVal Value As Boolean)
			mblnApplyStyles = Value
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

	Private Function CreateExcelApplication() As Boolean

		On Error GoTo LocalErr

		mxlApp = CreateObject("Excel.Application")

		CreateExcelApplication = True

		Exit Function

LocalErr:
		mstrErrorMessage = "Error opening Excel Application"
		CreateExcelApplication = False

	End Function



	Public Function GetFile(ByRef objParent As clsOutputRun, ByRef colStyles As Collection) As Boolean

		On Error GoTo LocalErr


		If Not CreateExcelApplication() Then
			GetFile = False
			Exit Function
		End If


		''Just in case we are emailing but not saving...
		'If mblnEmail And Not mblnSave Then
		'  mstrFileName = objParent.GetTempFileName(mstrFileName)
		'End If


		' Leave the app there after user has closed the worksheet
		mxlApp.UserControl = True
		mxlApp.DisplayAlerts = False

		'Check if file already exists...
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If Dir(mstrFileName) <> vbNullString And mstrFileName <> vbNullString Then

			Select Case mlngSaveExisting
				Case 0 'Overwrite
					If Not objParent.KillFile(mstrFileName) Then
						GetFile = False
						Exit Function
					End If

					GetWorkBook(strWorkbook:="New", strWorksheet:="New")

				Case 1 'Do not overwrite (fail)
					mxlApp.Quit()
					'UPGRADE_NOTE: Object mxlApp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					mxlApp = Nothing
					mstrErrorMessage = "File already exists."

				Case 2 'Add Sequential number to file
					mstrFileName = mobjParent.GetSequentialNumberedFile(mstrFileName)
					GetWorkBook(strWorkbook:="New", strWorksheet:="New")

				Case 3 'Append to existing file
					GetWorkBook(strWorkbook:="Open", strWorksheet:="Existing")

				Case 4 'Create new worksheet within existing workbook...
					GetWorkBook(strWorkbook:="Open", strWorksheet:="New")

			End Select

		Else
			GetWorkBook(strWorkbook:="New", strWorksheet:="New")

		End If

		GetFile = (mstrErrorMessage = vbNullString)

		Exit Function

LocalErr:
		mstrErrorMessage = Err.Description
		GetFile = False

	End Function


	Private Sub GetWorkBook(ByRef strWorkbook As String, ByRef strWorksheet As String)

		Dim strFormat As String
		Dim strTempFile As String
		Dim lngCount As Integer
		Dim lngOriginalFormat As Integer


		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If mblnApplyStyles And mstrXLTemplate <> "" And Dir(mstrXLTemplate) <> "" Then

			If Not IsFileCompatibleWithExcelVersion(mstrXLTemplate, Val(mxlApp.Version)) Then
				mstrErrorMessage = "Your User Configuration Output Options are set to use a template file which is not compatible with your version of Microsoft Office."
				Exit Sub
			End If

			mxlTemplateBook = mxlApp.Workbooks.Open(mstrXLTemplate, ReadOnly:=True)


			'Save a temp template in the format of the output...
			If mstrFileName <> vbNullString Then
				strFormat = GetOfficeSaveAsFormat(mstrFileName, Val(mxlApp.Version), modIntClient.OfficeApp.oaExcel)
				If strFormat <> "" Then
					strTempFile = mobjParent.GetTempFileName("")
					mxlTemplateBook.SaveAs(strTempFile, Val(strFormat))
					mxlTemplateBook.Close()
					mxlTemplateBook = mxlApp.Workbooks.Open(strTempFile, ReadOnly:=True)
				End If
			End If

			mxlTemplateSheet = mxlTemplateBook.ActiveSheet

		End If

		mstrSheetMode = strWorksheet
		Select Case strWorkbook
			Case "New"

				If Val(mxlApp.Version) >= 12 And mstrXLTemplate <> vbNullString Then
					'Make sure the new workbook is in the same format as the template
					'otherwise we won't be able to copy sheets into the new workbook.
					lngOriginalFormat = mxlApp.DefaultSaveFormat
					mxlApp.DefaultSaveFormat = CShort(GetOfficeSaveAsFormat(mstrXLTemplate, Val(mxlApp.Version), modIntClient.OfficeApp.oaExcel))
					mxlWorkBook = mxlApp.Workbooks.Add
					mxlApp.DefaultSaveFormat = lngOriginalFormat
				Else
					mxlWorkBook = mxlApp.Workbooks.Add
				End If

				For lngCount = 1 To mxlWorkBook.Sheets.Count - 1
					'UPGRADE_WARNING: Couldn't resolve default property of object mxlWorkBook.Sheets().Delete. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mxlWorkBook.Sheets(1).Delete()
				Next
				mxlDeleteSheet = mxlWorkBook.Sheets(1)

			Case "Open"
				If Not IsFileCompatibleWithExcelVersion(mstrFileName, Val(mxlApp.Version)) Then
					mstrErrorMessage = "This definition is set to append to a file which is not compatible with your version of Microsoft Office."
					Exit Sub
				End If

				mxlWorkBook = mxlApp.Workbooks.Open(mstrFileName)

		End Select

		Exit Sub

LocalErr:
		mstrErrorMessage = "Error getting new workbook (" & Err.Description & ")"

	End Sub


	Private Sub GetWorksheet(ByRef strSheetName As String)

		Dim blnFound As Boolean

		On Error GoTo LocalErr

		mblnAppending = False
		mlngAppendStartRow = 0


		'If we are appending, then see if there is an existing worksheet with this name...
		blnFound = False
		If mstrSheetMode = "Existing" Then
			For Each mxlWorkSheet In mxlWorkBook.Worksheets
				If Trim(mxlWorkSheet.Name) = FormatSheetName(strSheetName) Then
					mxlWorkSheet.Activate()
					blnFound = True
					Exit For
				End If
			Next mxlWorkSheet
		End If


		If blnFound Then
			StartAtBottomOfSheet()
			mblnAppending = True
		Else
			If Not mxlTemplateSheet Is Nothing Then
				mxlTemplateSheet.Copy(After:=mxlWorkBook.Sheets(mxlWorkBook.Sheets.Count))
				mxlWorkSheet = mxlWorkBook.ActiveSheet
				StartAtBottomOfSheet()
			Else
				mxlWorkSheet = mxlWorkBook.Sheets.Add(After:=mxlWorkBook.Sheets(mxlWorkBook.Sheets.Count))
			End If
			SetSheetName(mxlWorkSheet, strSheetName)
		End If


		If Not (mxlDeleteSheet Is Nothing) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object mxlDeleteSheet.Delete. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mxlDeleteSheet.Delete()
			'UPGRADE_NOTE: Object mxlDeleteSheet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			mxlDeleteSheet = Nothing
		End If

		Exit Sub

LocalErr:
		mstrErrorMessage = "Error getting new worksheet (" & Err.Description & ")"

	End Sub

	Public Function AddPage(ByRef strDefTitle As String, ByRef strSheetName As String, ByRef colStyles As Collection) As Object

		On Error GoTo LocalErr

		mstrDefTitle = strDefTitle

		If mblnPivotTable Then
			GetWorksheet("Data " & strSheetName)
		Else
			GetWorksheet(strSheetName)
		End If

		If Not mblnChart And Not mblnPivotTable Then
			If mxlFirstSheet Is Nothing Then
				mxlFirstSheet = mxlWorkSheet
			End If
		End If

		If mlngAppendStartRow = 0 Then
			mlngDataCurrentRow = mlngDataStartRow
		End If

		If mblnApplyStyles = False Then
			If Not mblnAppending Then
				mlngDataCurrentRow = 1
			End If
			mlngDataStartCol = 1
			mlngHeaderCols = 0
			mlngHeaderRows = 0
		End If

		Exit Function

LocalErr:
		mstrErrorMessage = Err.Description

	End Function


	Public Sub DataArray(ByRef strArray(,) As String, ByRef colColumns As Collection, ByRef colStyles As Collection, ByRef colMerges As Collection)

		Dim objColumn As clsColumn
		Dim lngGridCol As Integer
		Dim lngGridRow As Integer
		Dim lngXLCol As Integer
		Dim lngXLRow As Integer
		Dim strCell As String

		On Error GoTo LocalErr


		If mstrErrorMessage <> vbNullString Then
			Exit Sub
		End If

		If UBound(strArray, 1) > 255 Then
			mstrErrorMessage = "Maximum of 255 columns exceeded"
			Exit Sub
		End If


		PrepareRows(UBound(strArray, 2), colColumns, colStyles)

		lngXLCol = mlngDataStartCol
		lngXLRow = mlngDataCurrentRow
		For lngGridRow = 0 To UBound(strArray, 2)
			For lngGridCol = 0 To UBound(strArray, 1)

				With mxlWorkSheet.Cells._Default(lngXLRow + lngGridRow, lngXLCol + lngGridCol)

					'MH20050104 Fault 9695 & 9696
					'UPGRADE_WARNING: Couldn't resolve default property of object mxlWorkSheet.Cells(lngXLRow + lngGridRow, lngXLCol + lngGridCol).NumberFormat. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .NumberFormat = DateFormat() & ";@" Then
						If strArray(lngGridCol, lngGridRow) Like "??/??/????" Or strArray(lngGridCol, lngGridRow) = vbNullString Then
							strArray(lngGridCol, lngGridRow) = Replace(VB6.Format(strArray(lngGridCol, lngGridRow), "mm/dd/yyyy"), GetSystemDateSeparator, "/")
						Else
							'A non-date in a date column (Report sub totals for example)...
							'UPGRADE_WARNING: Couldn't resolve default property of object mxlWorkSheet.Cells().NumberFormat. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							.NumberFormat = "General"
						End If
					End If


					'MH20031113 Fault 7602
					'.FormulaR1C1 = strArray(lngGridCol, lngGridRow)
					'UPGRADE_WARNING: Couldn't resolve default property of object mxlWorkSheet.Cells(lngXLRow + lngGridRow, lngXLCol + lngGridCol).FormulaR1C1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.FormulaR1C1 = IIf(Left(strArray(lngGridCol, lngGridRow), 1) = "'", "'", vbNullString) & strArray(lngGridCol, lngGridRow)

					'If lngGridRow < mlngHeaderRows Then
					'  .HorizontalAlignment = xlCenter
					'End If
				End With

				'      If gobjProgress.Visible And gobjProgress.Cancelled Then
				'        mstrErrorMessage = "Cancelled by user."
				'        Exit Sub
				'      End If

			Next
		Next


		If mblnChart Then
			ApplyStyle(UBound(strArray, 1), UBound(strArray, 2), colStyles)
			ApplyCellOptions(UBound(strArray, 1), colStyles, True)

			CreateChart(mlngDataCurrentRow + UBound(strArray, 2), mlngDataStartCol + UBound(strArray, 1), colStyles)
			ApplyCellOptions(UBound(strArray, 1), colStyles, False)

			'Delete superfluous rows and cols if setup in User Config reports section
			If mblnXLExcelOmitLeftCol Then mxlWorkSheet.Range("A:A").Delete()
			If mblnXLExcelOmitTopRow Then mxlWorkSheet.Range("1:1").Delete()

		ElseIf mblnPivotTable Then

			If UBound(strArray, 1) < 1 Then
				mstrErrorMessage = "Unable to create a pivot table for a single column of data."
			Else
				ApplyStyle(UBound(strArray, 1), UBound(strArray, 2), colStyles)

				CreatePivotTable(mlngDataCurrentRow + UBound(strArray, 2), mlngDataStartCol + UBound(strArray, 1), strArray(0, 0), strArray(1, 0), strArray(UBound(strArray), 0), colStyles, colColumns)
			End If

		Else
			If mblnApplyStyles Then
				ApplyStyle(UBound(strArray, 1), UBound(strArray, 2), colStyles)
				ApplyMerges(colMerges)
			End If
			ApplyCellOptions(UBound(strArray, 1), colStyles, True)

			'Delete superfluous rows and cols if setup in User Config reports section
			If mblnXLExcelOmitLeftCol Then mxlWorkSheet.Range("A:A").Delete()
			If mblnXLExcelOmitTopRow Then mxlWorkSheet.Range("1:1").Delete()

		End If

		mlngDataCurrentRow = mlngDataCurrentRow + UBound(strArray, 2) + IIf(mblnApplyStyles, 2, 1)

		Exit Sub

LocalErr:
		mstrErrorMessage = Err.Description

	End Sub


	Private Sub CreateChart(ByRef lngMaxRows As Integer, ByRef lngMaxCols As Integer, ByRef colStyles As Collection)

		Dim xlChart As Microsoft.Office.Interop.Excel.Chart
		Dim xlData As Microsoft.Office.Interop.Excel.Range
		Dim strSheetName As String

		On Error GoTo LocalErr

		xlData = mxlWorkSheet.Range(mxlWorkSheet.Cells._Default(mlngDataCurrentRow, mlngDataStartCol), mxlWorkSheet.Cells._Default(lngMaxRows, lngMaxCols))
		strSheetName = mxlWorkSheet.Name & " Chart"

		xlChart = mxlApp.Charts.Add(After:=mxlWorkSheet)
		With xlChart
			.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xl3DColumnClustered
			.SetSourceData(Source:=xlData, PlotBy:=Microsoft.Office.Interop.Excel.XlRowCol.xlColumns)
			.Location(Where:=Microsoft.Office.Interop.Excel.XlChartLocation.xlLocationAsNewSheet)
			'.ChartTitle.Caption = mstrDefTitle
			.HasTitle = True
			'MH20061204 Fault 11230
			'.ChartTitle.Characters.Text = mstrDefTitle
			.ChartTitle.Text = mstrDefTitle
			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().Bold. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.ChartTitle.Font.Bold = colStyles.Item("Title").Bold
			.ChartTitle.Font.Size = 12
			'MH20050113 Fault 9376
			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().Underline. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.ChartTitle.Font.Underline = colStyles.Item("Title").Underline
			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().ForeCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.ChartTitle.Font.Color = colStyles.Item("Title").ForeCol
		End With

		SetSheetName((mxlWorkBook.ActiveChart), strSheetName)
		If mxlFirstSheet Is Nothing Then
			mxlFirstSheet = mxlWorkBook.ActiveChart
		End If


		Exit Sub

LocalErr:
		mstrErrorMessage = Err.Description

	End Sub


	Private Sub CreatePivotTable(ByRef lngMaxRows As Integer, ByRef lngMaxCols As Integer, ByRef strHor As String, ByRef strVer As String, ByRef strInt As String, ByRef colStyles As Collection, ByRef colColumns As Collection)

		Dim xlPivot As Microsoft.Office.Interop.Excel.PivotTable
		Dim xlDataSheet As Microsoft.Office.Interop.Excel.Worksheet
		Dim xlData As Microsoft.Office.Interop.Excel.Range
		Dim xlStart As Microsoft.Office.Interop.Excel.Range
		Dim objColumn As clsColumn
		Dim strSheetName As String
		Dim xlFunc As Microsoft.Office.Interop.Excel.XlConsolidationFunction

		On Error GoTo LocalErr

		mxlApp.DisplayAlerts = True

		'EXCEL 97
		'    ActiveSheet.PivotTableWizard SourceType:=xlDatabase, SourceData:= _
		''        "Personnel_Records!R5C2:R20C10", TableDestination:="", TableName:= _
		''        "PivotTable1"
		'    ActiveSheet.PivotTables("PivotTable1").AddFields RowFields:="Forenames", _
		''        ColumnFields:="Surname"
		'    ActiveSheet.PivotTables("PivotTable1").PivotFields("Copy of abs dur"). _
		''        Orientation = xlDataField

		'EXCEL 2000
		'    ActiveWorkbook.PivotCaches.Add(SourceType:=xlDatabase, SourceData:= _
		''        "'<Blank>'!R8C2:R11C9").CreatePivotTable TableDestination:=Range("B17"), _
		''        TableName:="PivotTable1"
		'    ActiveSheet.PivotTables("PivotTable1").SmallGrid = False
		'    ActiveSheet.PivotTables("PivotTable1").AddFields RowFields:="Department'", _
		''        ColumnFields:="Forenames"
		'    ActiveSheet.PivotTables("PivotTable1").PivotFields("include_salary_column"). _
		''        Orientation = xlDataField

		'EXCEL XP
		'  mxlWorkBook.PivotCaches.Add(SourceType:=xlDatabase, SourceData:= _
		''      mxlWorkSheet.Range(mxlWorkSheet.Cells(mlngDataCurrentRow, mlngDataStartCol), mxlWorkSheet.Cells(lngMaxRows, lngMaxCols))).CreatePivotTable TableDestination:="", TableName:= _
		''      "PivotTable1"

		'Excel 2010
		'    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
		''        "Data Personnel_Records!R5C2:R59C11", Version:=xlPivotTableVersion14). _
		''        CreatePivotTable TableDestination:="Personnel_Records!R5C2", TableName:= _
		''        "PivotTable1", DefaultVersion:=xlPivotTableVersion14

		xlData = mxlWorkSheet.Range(mxlWorkSheet.Cells._Default(mlngDataCurrentRow, mlngDataStartCol), mxlWorkSheet.Cells._Default(lngMaxRows, lngMaxCols))
		strSheetName = Mid(mxlWorkSheet.Name, 6)
		'SetSheetName mxlWorkSheet, "Data " & mxlWorkSheet.Name
		xlDataSheet = mxlWorkSheet

		GetWorksheet(strSheetName)
		If mxlFirstSheet Is Nothing Then
			mxlFirstSheet = mxlWorkSheet
		End If

		xlDataSheet.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden
		xlStart = mxlWorkSheet.Cells._Default(mlngDataCurrentRow, mlngDataStartCol)

		mxlApp.DisplayAlerts = False


		'MH20100628
		If Val(mxlApp.Version) > 12 Then
			xlPivot = mxlWorkBook.PivotCaches.Create(SourceType:=Microsoft.Office.Interop.Excel.XlPivotTableSourceType.xlDatabase, SourceData:=xlData).CreatePivotTable(TableDestination:=xlStart)
		Else
			xlPivot = mxlWorkSheet.PivotTableWizard(SourceType:=Microsoft.Office.Interop.Excel.XlPivotTableSourceType.xlDatabase, SourceData:=xlData, TableDestination:=xlStart)
		End If


		With xlPivot
			.AddFields(RowFields:=strVer, ColumnFields:=strHor)

			'AE20071017 Fault #12540
			Select Case mobjParent.PivotDataFunction
				Case "Count"
					xlFunc = Microsoft.Office.Interop.Excel.XlConsolidationFunction.xlCount
				Case "Average"
					xlFunc = Microsoft.Office.Interop.Excel.XlConsolidationFunction.xlAverage
				Case "Maximum"
					xlFunc = Microsoft.Office.Interop.Excel.XlConsolidationFunction.xlMax
				Case "Minimum"
					xlFunc = Microsoft.Office.Interop.Excel.XlConsolidationFunction.xlMin
				Case "Total"
					xlFunc = Microsoft.Office.Interop.Excel.XlConsolidationFunction.xlSum
			End Select

			'.PivotFields(strInt).Orientation = xlDataField

			With .PivotFields(strInt)
				'UPGRADE_WARNING: Couldn't resolve default property of object xlPivot.PivotFields().Orientation. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlDataField
				'UPGRADE_WARNING: Couldn't resolve default property of object xlPivot.PivotFields().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Name = mobjParent.PivotDataFunction & " of " & strInt
				'UPGRADE_WARNING: Couldn't resolve default property of object xlPivot.PivotFields().Function. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Function = xlFunc
			End With

			.NullString = IIf(mobjParent.PivotSuppressBlanks, "", "0")

			objColumn = colColumns.Item(colColumns.Count())
			If objColumn.DecPlaces > 0 Then
				If objColumn.DecPlaces > 100 Then objColumn.DecPlaces = 100
				.DataBodyRange.NumberFormat = IIf(objColumn.ThousandSeparator, "#,##0", "0") & IIf(objColumn.DecPlaces, "." & New String("0", objColumn.DecPlaces), "")
			End If
			'UPGRADE_NOTE: Object objColumn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objColumn = Nothing

			mxlApp.DisplayAlerts = True
		End With

		'UPGRADE_WARNING: Couldn't resolve default property of object colStyles(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ApplyStyleToRange(xlStart, colStyles.Item("Heading"))
		'UPGRADE_WARNING: Couldn't resolve default property of object colStyles(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ApplyStyleToRange(xlPivot.RowRange, colStyles.Item("Heading"))
		'UPGRADE_WARNING: Couldn't resolve default property of object colStyles(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ApplyStyleToRange(xlPivot.ColumnRange, colStyles.Item("Heading"))
		'UPGRADE_WARNING: Couldn't resolve default property of object colStyles(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ApplyStyleToRange(xlPivot.DataBodyRange, colStyles.Item("Data"))
		mxlWorkSheet.Range("A1").Select()
		mlngHeaderCols = 1

		ApplyCellOptions(xlPivot.ColumnRange.Columns.Count, colStyles, True)

		'Delete superfluous rows and cols if setup in User Config reports section
		If mblnXLExcelOmitLeftCol Then mxlWorkSheet.Range("A:A").Delete()
		If mblnXLExcelOmitTopRow Then mxlWorkSheet.Range("1:1").Delete()

		'UPGRADE_NOTE: Object xlPivot may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		xlPivot = Nothing
		mxlApp.DisplayAlerts = False

		Exit Sub

LocalErr:
		mstrErrorMessage = Err.Description

	End Sub


	Private Sub PrepareRows(ByRef lngRowCount As Integer, ByRef colColumns As Collection, ByRef colStyles As Collection)

		Dim objColumn As clsColumn
		Dim lngCount As Integer

		On Error GoTo LocalErr

		With mxlWorkSheet

			'    If mstrXLTemplate <> vbNullString Then
			'      For lngCount = 1 To lngRowCount
			'        .Rows(mlngDataCurrentRow + mlngHeaderRows).Select
			'        mxlApp.Selection.Copy
			'        mxlApp.Selection.Insert Shift:=xlDown
			'      Next
			'      mxlApp.CutCopyMode = False
			'    End If
			If .Visible Then
				.Range("A1").Select()
			End If


			If mlngHeaderRows > 0 Then
				With .Range(.Cells._Default(mlngDataCurrentRow, mlngDataStartCol), .Cells._Default(mlngDataCurrentRow + mlngHeaderRows - 1, mlngDataStartCol + colColumns.Count()))
					.NumberFormat = "@"
				End With
			End If

			For lngCount = 0 To colColumns.Count() - 1

				objColumn = colColumns.Item(lngCount + 1)

				With .Range(.Cells._Default(mlngDataCurrentRow + mlngHeaderRows, mlngDataStartCol + lngCount), .Cells._Default(mlngDataCurrentRow + lngRowCount, mlngDataStartCol + lngCount))
					Select Case objColumn.DataType
						Case SQLDataType.sqlNumeric, SQLDataType.sqlInteger

							If objColumn.DecPlaces > 0 Then
								If objColumn.DecPlaces > 100 Then objColumn.DecPlaces = 100
								.NumberFormat = IIf(objColumn.ThousandSeparator, "#,##0", "0") & IIf(objColumn.DecPlaces, "." & New String("0", objColumn.DecPlaces), "")
							End If

						Case SQLDataType.sqlBoolean
							.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
							.NumberFormat = "@"
						Case SQLDataType.sqlUnknown
							'Leave it alone! (Required for percentages on Standard Reports)
						Case SQLDataType.sqlDate
							'MH20050104 Fault 9695 & 9696
							'Adding ;@ to the end formats it as "short date" so excel will look at the
							'regional settings when opening the workbook rather than force it to always
							'be in the format of the user who created the workbook.
							.NumberFormat = DateFormat() & ";@"
						Case Else
							.NumberFormat = "@"
					End Select
				End With

			Next

		End With

		Exit Sub

LocalErr:
		mstrErrorMessage = Err.Description

	End Sub


	Private Sub ApplyCellOptions(ByRef lngColCount As Integer, ByRef colStyles As Collection, ByRef blnGridLines As Boolean)

		Dim objRange As Microsoft.Office.Interop.Excel.Range
		Dim lngCount As Integer

		Dim lngMaxWidth As Double
		Dim lngTitleColWidth As Double
		Dim lngTitleSize As Double

		On Error GoTo LocalErr

		If mblnXLAutoFitCols Then
			mxlWorkSheet.Range(mxlWorkSheet.Cells._Default(mlngDataCurrentRow, mlngDataStartCol), mxlWorkSheet.Cells._Default(mlngDataCurrentRow, mlngDataStartCol + lngColCount)).EntireColumn.AutoFit()

			If Not mblnSizeColumnsIndependently Then
				lngMaxWidth = 0
				For lngCount = mlngDataStartCol + mlngHeaderCols To mlngDataStartCol + lngColCount
					'UPGRADE_WARNING: Couldn't resolve default property of object mxlWorkSheet.Columns(lngCount).ColumnWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If lngMaxWidth < mxlWorkSheet.Columns._Default(lngCount).ColumnWidth Then
						'UPGRADE_WARNING: Couldn't resolve default property of object mxlWorkSheet.Columns().ColumnWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						lngMaxWidth = mxlWorkSheet.Columns._Default(lngCount).ColumnWidth
					End If
				Next

				For lngCount = mlngDataStartCol + mlngHeaderCols To mlngDataStartCol + lngColCount
					'UPGRADE_WARNING: Couldn't resolve default property of object mxlWorkSheet.Columns().ColumnWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mxlWorkSheet.Columns._Default(lngCount).ColumnWidth = lngMaxWidth
				Next
			End If

		End If


		If mblnApplyStyles Then
			If blnGridLines Then
				With mxlApp.ActiveWindow
					.DisplayGridlines = mblnXLExcelGridlines
					.DisplayHeadings = mblnXLExcelHeaders
				End With
			End If

			With colStyles.Item("Title")
				'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.StartCol = glngSettingTitleCol
				'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.StartRow = IIf(mlngAppendStartRow > 0, mlngAppendStartRow, glngSettingTitleRow)
				'UPGRADE_WARNING: Couldn't resolve default property of object colStyles(Title).EndCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.EndCol = .StartCol
				'UPGRADE_WARNING: Couldn't resolve default property of object colStyles(Title).EndRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.EndRow = .StartRow
			End With

			'Put title in after autofit...
			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles(Title).StartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles(Title).StartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If colStyles.Item("Title").StartCol <> 0 And colStyles.Item("Title").StartRow <> 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object colStyles(Title).StartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mxlWorkSheet.Cells().FormulaR1C1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mxlWorkSheet.Cells._Default(colStyles.Item("Title").StartRow, colStyles.Item("Title").StartCol).FormulaR1C1 = mstrDefTitle
				'UPGRADE_WARNING: Couldn't resolve default property of object colStyles(Title).StartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				objRange = mxlWorkSheet.Cells._Default(colStyles.Item("Title").StartRow, colStyles.Item("Title").StartCol)
				'UPGRADE_WARNING: Couldn't resolve default property of object colStyles(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ApplyStyleToRange(objRange, colStyles.Item("Title"))


				'MH20020807 Fault 6562
				'Merge cells for the title column so that if you append to the file
				'then the title is not taken into account during column sizing.
				'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				With mxlWorkSheet.Columns._Default(colStyles.Item("Title").StartCol)
					'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mxlWorkSheet.Columns().ColumnWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					lngTitleColWidth = .ColumnWidth
					'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mxlWorkSheet.Columns().Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					lngMaxWidth = .Width
					'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mxlWorkSheet.Columns().AutoFit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.AutoFit()
					'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mxlWorkSheet.Columns().Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					lngTitleSize = .Width

					'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					lngCount = colStyles.Item("Title").StartCol
					Do
						lngCount = lngCount + 1
						'UPGRADE_WARNING: Couldn't resolve default property of object mxlWorkSheet.Columns().Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						lngMaxWidth = lngMaxWidth + mxlWorkSheet.Columns._Default(lngCount).Width
					Loop While lngMaxWidth < lngTitleSize

					With mxlWorkSheet
						.Range(.Cells._Default(objRange.Row, objRange.Column), .Cells._Default(objRange.Row, lngCount)).Merge()
					End With
					'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mxlWorkSheet.Columns().ColumnWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.ColumnWidth = lngTitleColWidth
				End With
				'NHRD09072012 Jira HRPRO-2308
				'      'Delete superfluous rows and cols if setup in User Config reports section
				'      If mblnXLExcelOmitLeftCol Then mxlWorkSheet.Range("A:A").Delete
				'      If mblnXLExcelOmitTopRow Then mxlWorkSheet.Range("1:1").Delete
			End If
		End If

		Exit Sub

LocalErr:
		mstrErrorMessage = Err.Description

	End Sub


	Private Sub ApplyStyle(ByRef lngNumCols As Integer, ByRef lngNumRows As Integer, ByRef colStyles As Collection)

		Dim objStyle As clsOutputStyle
		Dim objRange As Microsoft.Office.Interop.Excel.Range
		Dim lngCol As Integer
		Dim lngRow As Integer

		On Error GoTo LocalErr

		lngCol = mlngDataStartCol
		lngRow = mlngDataCurrentRow

		With colStyles.Item("Title")
			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.StartCol = glngSettingTitleCol
			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.StartRow = IIf(mlngAppendStartRow > 0, mlngAppendStartRow, glngSettingTitleRow)
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
			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().EndCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.EndCol = lngNumCols
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
				'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().EndRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.EndRow = lngNumRows
			End With
		End If

		With colStyles.Item("Data")
			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.StartCol = mlngHeaderCols
			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.StartRow = mlngHeaderRows
			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().EndCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.EndCol = lngNumCols
			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().EndRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.EndRow = lngNumRows
		End With


		For Each objStyle In colStyles
			If objStyle.Name <> "Title" Then
				If objStyle.StartRow + lngRow > 0 And objStyle.StartCol + lngCol > 0 Then
					objRange = mxlWorkSheet.Range(mxlWorkSheet.Cells._Default(objStyle.StartRow + lngRow, objStyle.StartCol + lngCol), mxlWorkSheet.Cells._Default(objStyle.EndRow + lngRow, objStyle.EndCol + lngCol))
					ApplyStyleToRange(objRange, objStyle)
				End If
			End If
		Next objStyle

		Exit Sub

LocalErr:
		mstrErrorMessage = Err.Description

	End Sub


	Private Sub ApplyMerges(ByRef colMerges As Collection)

		Dim objMerge As clsOutputStyle
		Dim objRange As Microsoft.Office.Interop.Excel.Range
		Dim lngCol As Integer
		Dim lngRow As Integer

		On Error GoTo LocalErr

		lngCol = mlngDataStartCol
		lngRow = mlngDataCurrentRow

		For Each objMerge In colMerges
			If objMerge.StartRow + lngRow > 0 And objMerge.StartCol + lngCol > 0 Then
				objRange = mxlWorkSheet.Range(mxlWorkSheet.Cells._Default(objMerge.StartRow + lngRow, objMerge.StartCol + lngCol), mxlWorkSheet.Cells._Default(objMerge.EndRow + lngRow, objMerge.EndCol + lngCol))
				objRange.MergeCells = True
				objRange.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlTop
			End If
		Next objMerge

		Exit Sub

LocalErr:
		mstrErrorMessage = Err.Description

	End Sub


	Private Sub ApplyStyleToRange(ByRef objRange As Microsoft.Office.Interop.Excel.Range, ByRef objStyle As clsOutputStyle)

		On Error GoTo LocalErr

		With objRange

			If objStyle.CenterText Then
				.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
			End If

			.Font.Bold = objStyle.Bold
			.Font.Underline = objStyle.Underline
			.Font.Color = objStyle.ForeCol

			'Don't do the backcol nor gridlines for the title...
			If objStyle.Name <> "Title" Then
				.Interior.Color = objStyle.BackCol

				On Error Resume Next

				If objStyle.Gridlines Then
					.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic)
					If objStyle.StartCol <> objStyle.EndCol Then
						.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
					End If
					If objStyle.StartRow <> objStyle.EndRow Then
						.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
					End If
				Else
					.BorderAround(Microsoft.Office.Interop.Excel.Constants.xlNone)
					.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
					.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
					.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
					.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
					.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
					.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
					.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
					.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
				End If

			End If

		End With

		Exit Sub

LocalErr:
		mstrErrorMessage = Err.Description

	End Sub


	Public Sub Complete()

		Dim objChart As Microsoft.Office.Interop.Excel.Chart
		Dim objWorksheet As Microsoft.Office.Interop.Excel.Worksheet
		Dim strFormat As String
		Dim strTempFile As String
		Dim strExtension As String
		Dim aryFileBits() As String

		On Error GoTo LocalErr

		If mstrErrorMessage <> vbNullString Then
			Exit Sub
		End If

		'blnOffice2007 = (Val(mxlApp.Version) >= 12)
		'UPGRADE_WARNING: Couldn't resolve default property of object mxlFirstSheet.Activate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mxlFirstSheet.Activate()

		'SAVE
		If mblnSave Then
			mstrErrorMessage = "Error saving file <" & mstrFileName & ">"

			strFormat = GetOfficeSaveAsFormat(mstrFileName, Val(mxlApp.Version), modIntClient.OfficeApp.oaExcel)
			If strFormat = "" Then
				mstrErrorMessage = "This definition is set to save in a file format which is not compatible with your version of Microsoft Office."
				GoTo TidyAndExit
			End If

			' calculate the appropriate output type
			aryFileBits = Split(mstrFileName, ".")
			strExtension = aryFileBits(UBound(aryFileBits))

			Select Case UCase(strExtension)
				Case "XLSX"
					mxlWorkBook.SaveAs(mstrFileName, FileFormat:=Val(strFormat))
				Case "XLS"
					mxlWorkBook.SaveAs(mstrFileName, FileFormat:=Val(strFormat))
				Case "HTML"
					mxlWorkBook.SaveAs(mstrFileName, FileFormat:=Val(CStr(Microsoft.Office.Interop.Excel.XlFileFormat.xlHtml)))
			End Select

		End If

		'EMAIL
		If mblnEmail Then
			mstrErrorMessage = "Error sending email"

			strFormat = GetOfficeSaveAsFormat((mobjParent.EmailAttachAs), Val(mxlApp.Version), modIntClient.OfficeApp.oaExcel)
			If strFormat = "" Then
				mstrErrorMessage = "This definition is set to email an attachment in a file format which is not compatible with your version of Microsoft Office."
				GoTo TidyAndExit
			End If

			strTempFile = mobjParent.GetTempFileName((mobjParent.EmailAttachAs))
			mxlWorkBook.SaveAs(strTempFile, Val(strFormat))
			mxlWorkBook.Close(False)
			mobjParent.SendEmail(strTempFile)
			mxlWorkBook = mxlApp.Workbooks.Open(strTempFile)
		End If

		'PRINTER
		Dim strCurrentPrinter As String
		If mblnPrinter Then
			'TM23122003 FAULT - DEFAULT PRINTER
			mstrErrorMessage = "Error printing"
			'mobjParent.SetPrinter

			If mblnChart Then
				For Each objChart In mxlWorkBook.Charts
					objChart.PrintOut(, , , , mstrPrinterName)
				Next objChart
			Else
				For Each objWorksheet In mxlWorkBook.Worksheets
					If objWorksheet.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible Then
						objWorksheet.PrintOut(, , , , mstrPrinterName)
					End If
				Next objWorksheet
			End If

			'mobjParent.ResetDefaultPrinter
		End If


		'SCREEN
		If mblnScreen Then
			mstrErrorMessage = "Error displaying Excel"
			mxlApp.DisplayAlerts = True
			mxlApp.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMaximized
			mxlApp.Visible = True
			mxlWorkBook.Activate()
			'UPGRADE_NOTE: Object mxlWorkBook may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			mxlWorkBook = Nothing
			'UPGRADE_NOTE: Object mxlApp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			mxlApp = Nothing 'Stops Excel quitting...
		Else
			mxlWorkBook.Saved = True
			mxlWorkBook.Close()
			mxlApp.Quit()
		End If

		mstrErrorMessage = vbNullString

TidyAndExit:
		ClearUp()

		Exit Sub

LocalErr:
		mstrErrorMessage = mstrErrorMessage & IIf(Err.Description <> vbNullString, vbCrLf & " (" & Err.Description & ")", vbNullString)
		Resume TidyAndExit

	End Sub

	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object mxlFirstSheet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mxlFirstSheet = Nothing
		'UPGRADE_NOTE: Object mxlTemplateSheet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mxlTemplateSheet = Nothing
		'UPGRADE_NOTE: Object mxlTemplateBook may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mxlTemplateBook = Nothing
		'UPGRADE_NOTE: Object mxlWorkSheet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mxlWorkSheet = Nothing
		'UPGRADE_NOTE: Object mxlWorkBook may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mxlWorkBook = Nothing
		'UPGRADE_NOTE: Object mxlApp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mxlApp = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub

	Private Sub StartAtBottomOfSheet()

		'Start at the bottom of the sheet
		mlngAppendStartRow = mxlWorkSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell).Row + IIf(mblnApplyStyles, 2, 1)
		mlngDataCurrentRow = mlngAppendStartRow + IIf(mblnApplyStyles, 2, 1)

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


	Private Function SetSheetName(ByRef objObject As Object, ByVal strSheetName As String) As Boolean

		Dim strNumber As String
		Dim lngCount As Integer

		strSheetName = FormatSheetName(strSheetName)

		On Error Resume Next
		Err.Clear()
		If strSheetName <> vbNullString Then
			'Sheet may already exist so add sequential number
			'UPGRADE_WARNING: Couldn't resolve default property of object objObject.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			objObject.Name = strSheetName
		Else
			strSheetName = "Sheet"
		End If

		'UPGRADE_WARNING: Couldn't resolve default property of object objObject.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If objObject.Name <> strSheetName Then
			lngCount = 1
			Do
				lngCount = lngCount + 1
				Err.Clear()
				strNumber = "(" & CStr(lngCount) & ")"
				'UPGRADE_WARNING: Couldn't resolve default property of object objObject.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'		objObject.Name = Left(strSheetName, 31 - Len(strNumber)) & strNumber
				If lngCount > 256 Then
					mstrErrorMessage = "Error naming sheet"
					Exit Function
				End If
			Loop While Err.Number > 0
		End If


		On Error Resume Next
		'MH20031117 Fault 7628
		If mxlTemplateSheet Is Nothing Then
			'UPGRADE_WARNING: Couldn't resolve default property of object objObject.PageSetup. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			With objObject.PageSetup
				'UPGRADE_WARNING: Couldn't resolve default property of object objObject.PageSetup. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.LeftFooter = "Created on &D at &T by " & gsUsername
				'UPGRADE_WARNING: Couldn't resolve default property of object objObject.PageSetup. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.RightFooter = "Page &P"
				'UPGRADE_WARNING: Couldn't resolve default property of object objObject.PageSetup. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Orientation = IIf(mblnXLLandscape, Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape, Microsoft.Office.Interop.Excel.XlPageOrientation.xlPortrait)
				mxlWorkSheet.DisplayPageBreaks = False
			End With
		End If

		SetSheetName = True

	End Function
End Class