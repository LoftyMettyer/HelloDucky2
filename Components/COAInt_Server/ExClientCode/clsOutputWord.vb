Option Strict Off
Option Explicit On

Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Enums

Friend Class clsOutputWord
	Inherits BaseOutputFormat

	'	Private mwrdApp As Microsoft.Office.Interop.Word.Application
	'	Private mwrdDoc As Microsoft.Office.Interop.Word.Document
	'	Private mwrdTable As Microsoft.Office.Interop.Word.Table
	'	Private mobjParent As clsOutputRun

	'	Private mblnScreen As Boolean
	'	Private mblnPrinter As Boolean
	'	Private mstrPrinterName As String
	'	Private mblnSave As Boolean
	'	Private mlngSaveExisting As Integer
	'	Private mblnEmail As Boolean
	'	Private mstrFileName As String

	'	Private mstrDefTitle As String
	'	Private mstrErrorMessage As String
	'	Private mlngPageCount As Integer
	'	Private mlngHeaderRows As Integer
	'	Private mlngHeaderCols As Integer
	'	Private mblnHeaderVertical As Boolean
	'	Private mblnApplyStyles As Boolean

	'	Private mstrWrdTemplate As String
	'	Private mblnWrdWordGridlines As Boolean
	'	'Private mblnWrdWordHeaders As Boolean
	'	Private mblnWrdAutoFitCols As Boolean
	'	Private mblnWrdLandscape As Boolean

	'	Public Sub ClearUp()
	'		
	'		'UPGRADE_NOTE: Object mwrdTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	'		mwrdTable = Nothing
	'		'UPGRADE_NOTE: Object mwrdDoc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	'		mwrdDoc = Nothing
	'		If Not mwrdApp Is Nothing Then mwrdApp.Quit(False)
	'		'UPGRADE_NOTE: Object mwrdApp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	'		mwrdApp = Nothing
	'	End Sub

	'	'Public Function RecordProfilePage(pfrmRecProfile As Form, _
	'	''  piPageNumber As Integer, _
	'	''  pcolStyles As Collection)
	'	'  ' Output the record profile page to Word.
	'	'
	'	'  On Error GoTo ErrorTrap
	'	'  'gobjErrorStack.PushStack "clsOutputWord.RecordProfilePage()"
	'	'
	'	'  Dim fOK As Boolean
	'	'  Dim fFirstColumnDone As Boolean
	'	'  Dim iLoop As Integer
	'	'  Dim iLoop2 As Integer
	'	'  Dim iLoop3 As Integer
	'	'  Dim iLoop4 As Integer
	'	'  Dim ctlTemp As Control
	'	'  Dim varBookmark As Variant
	'	'  Dim fGridPreceded As Boolean
	'	'  Dim fGridFollowed As Boolean
	'	'  Dim sTitle As String
	'	'  Dim sTemp As String
	'	'  Dim fPhotoDone As Boolean
	'	'  Dim objRecProfTable As clsRecordProfileTabDtl
	'	'  Dim fPageBreak As Boolean
	'	'  Dim sTempName As String
	'	'  Dim lngBorder As Long
	'	'  Dim iLastGroup As Integer
	'	'  Dim colMerges As Collection
	'	'  Dim objMerge As clsOutputStyle
	'	'  Dim iGroupStart As Integer
	'	'  Dim iTemp As Integer
	'	'  Dim objTemp As clsOutputStyle
	'	'  Dim iOriginalStyles As Integer
	'	'  Dim fHasHeadingColumn As Boolean
	'	'  Dim lngPrecedingRows As Long
	'	'
	'	'  Const strBookMark As String = "ASRSysTableStart"
	'	'
	'	'  Const RECPROFFOLLOWONCORRECTION = 10
	'	'
	'	'  Const COLUMN_ISHEADING = "IsHeading"
	'	'  Const COLUMN_ISPHOTO = "IsPhoto"
	'	'  Const PHOTOSTYLESET = "PhotoSS_"
	'	'
	'	'  fOK = True
	'	'  sTitle = pfrmRecProfile.Caption
	'	'  iOriginalStyles = pcolStyles.Count
	'	'  fPageBreak = False
	'	'
	'	'  For Each ctlTemp In pfrmRecProfile.Controls
	'	'    If ctlTemp.Container Is pfrmRecProfile.picOutput(piPageNumber) Then
	'	'      '
	'	'      ' LABEL control
	'	'      '
	'	'      If TypeOf ctlTemp Is Label Then
	'	'        If ctlTemp.Visible Then
	'	'          ' If we are page-breaking after the previous grid output
	'	'          ' then get Word to create a new page.
	'	'          If fPageBreak Then
	'	'            ' Get Word to create a new page.
	'	'            AddPage sTitle, "", pcolStyles
	'	'          End If
	'	'
	'	'          ' Check if we need to page break after this table's output.
	'	'          Set objRecProfTable = pfrmRecProfile.Definition.Item(ctlTemp.Tag)
	'	'          fPageBreak = objRecProfTable.PageBreak
	'	'
	'	'          ' Write the label's caption to Word.
	'	'          With mwrdDoc.ActiveWindow.Selection
	'	'            .TypeText ctlTemp.Caption
	'	'            mwrdApp.Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
	'	'            mwrdApp.Selection.Font.Bold = pcolStyles("Title").Bold
	'	'            mwrdApp.Selection.Font.Underline = IIf(pcolStyles("Title").Underline, wdUnderlineSingle, wdUnderlineNone)
	'	'            mwrdApp.Selection.EndKey Unit:=wdStory
	'	'            .TypeParagraph
	'	'            .TypeParagraph
	'	'          End With
	'	'
	'	'          mwrdApp.Selection.Font.Bold = False
	'	'          mwrdApp.Selection.Font.Underline = wdUnderlineNone
	'	'        End If
	'	'      End If
	'	'
	'	'      '
	'	'      ' GRID control
	'	'      '
	'	'      If TypeOf ctlTemp Is SSDBGrid Then
	'	'        ' Check if we need to page break after this table's output.
	'	'        Set objRecProfTable = pfrmRecProfile.Definition.Item(ctlTemp.Tag)
	'	'        fPageBreak = objRecProfTable.PageBreak
	'	'
	'	'        ' Check if this grid is preceded or followed IMMEDIATELY by other grids.
	'	'        ' ie. if this grid is part of a group of grids that are used to display
	'	'        ' data (including pictures) vertically.
	'	'        ' NB. Grids have only one row height value. To display pictures with their
	'	'        ' own row height, we actually put them in their own grid, and position this grid
	'	'        ' IMMEDIATELY after the normal data grid. Subsequent data is put in its own grid
	'	'        ' IMMEDIATELY after the picture's grid.
	'	'        ' This is what's meant by 'following' and 'preceding' grids.
	'	'        fGridFollowed = False
	'	'        fGridPreceded = False
	'	'
	'	'        If ctlTemp.Index > 1 Then
	'	'          If (ctlTemp.Container = pfrmRecProfile.grdOutput(ctlTemp.Index - 1).Container) And _
	'	''            (ctlTemp.Top = (pfrmRecProfile.grdOutput(ctlTemp.Index - 1).Top + pfrmRecProfile.grdOutput(ctlTemp.Index - 1).Height - RECPROFFOLLOWONCORRECTION)) Then
	'	'
	'	'            fGridPreceded = True
	'	'          End If
	'	'        End If
	'	'
	'	'        If ctlTemp.Index < pfrmRecProfile.grdOutput.Count - 1 Then
	'	'          If (ctlTemp.Container = pfrmRecProfile.grdOutput(ctlTemp.Index + 1).Container) And _
	'	''            ((ctlTemp.Top + ctlTemp.Height - RECPROFFOLLOWONCORRECTION) = pfrmRecProfile.grdOutput(ctlTemp.Index + 1).Top) Then
	'	'
	'	'            fGridFollowed = True
	'	'          End If
	'	'        End If
	'	'
	'	'        If Not fGridPreceded Then
	'	'          lngPrecedingRows = 0
	'	'          mwrdDoc.Bookmarks.Add strBookMark
	'	'          Set colMerges = New Collection
	'	'        End If
	'	'
	'	'        ' Send the column/group headers to the word document.
	'	'        fHasHeadingColumn = False
	'	'        sTemp = ""
	'	'        iLastGroup = -1
	'	'        iTemp = 0
	'	'        iGroupStart = 0
	'	'        fFirstColumnDone = False
	'	'
	'	'        For iLoop = 0 To ctlTemp.Columns.Count - 1
	'	'          If ctlTemp.Columns(iLoop).Name = COLUMN_ISHEADING Then
	'	'            fHasHeadingColumn = True
	'	'          End If
	'	'
	'	'          If (ctlTemp.ColumnHeaders) And (ctlTemp.Columns(iLoop).Visible) Then
	'	'            If (ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders) Then
	'	'              If (iLastGroup <> ctlTemp.Columns(iLoop).Group) Then
	'	'                ' Send the group headers to the word document.
	'	'                mwrdDoc.ActiveWindow.Selection.TypeText IIf(fFirstColumnDone, vbTab, "") & ctlTemp.Groups(ctlTemp.Columns(iLoop).Group).Caption
	'	'
	'	'                ' Remember if the group heading cells need to be merged.
	'	'                If iGroupStart < iTemp - 1 Then
	'	'                  Set objMerge = New clsOutputStyle
	'	'                  objMerge.StartCol = iGroupStart
	'	'                  objMerge.StartRow = 0
	'	'                  objMerge.EndCol = iTemp - 1
	'	'                  objMerge.EndRow = 0
	'	'
	'	'                  colMerges.Add objMerge
	'	'                  Set objMerge = Nothing
	'	'                End If
	'	'
	'	'                iGroupStart = iTemp
	'	'              Else
	'	'                mwrdDoc.ActiveWindow.Selection.TypeText IIf(fFirstColumnDone, vbTab, "")
	'	'              End If
	'	'
	'	'              iLastGroup = ctlTemp.Columns(iLoop).Group
	'	'            End If
	'	'
	'	'            ' Check if the column is a Separator column.
	'	'            ' If so we'll need to create a 'style' object for it.
	'	'            If ctlTemp.Columns(iLoop).StyleSet = "Separator" Then
	'	'              Set objTemp = New clsOutputStyle
	'	'
	'	'              With objTemp
	'	'                .StartCol = iTemp
	'	'                .StartRow = IIf((ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders), 2, 1)
	'	'                .EndCol = CLng(iTemp)
	'	'                .EndRow = -1
	'	'
	'	'                .BackCol = pcolStyles("HeadingCols").BackCol
	'	'                .ForeCol = pcolStyles("HeadingCols").ForeCol
	'	'                .Bold = pcolStyles("HeadingCols").Bold
	'	'                .Underline = pcolStyles("HeadingCols").Underline
	'	'                .Gridlines = pcolStyles("HeadingCols").Gridlines
	'	'                .Name = "RECPROFCOL_" & CStr(iLoop)
	'	'              End With
	'	'
	'	'              pcolStyles.Add objTemp
	'	'              Set objTemp = Nothing
	'	'            End If
	'	'
	'	'            ' Send the column/group headers to the word document.
	'	'            sTemp = sTemp & IIf(fFirstColumnDone, vbTab, "") & ctlTemp.Columns(iLoop).Caption
	'	'            fFirstColumnDone = True
	'	'            iTemp = iTemp + 1
	'	'          End If
	'	'        Next iLoop
	'	'
	'	'        ' Remember if the group heading cells need to be merged.
	'	'        If (ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders) Then
	'	'          If iGroupStart < iTemp - 1 Then
	'	'            Set objMerge = New clsOutputStyle
	'	'            objMerge.StartCol = iGroupStart
	'	'            objMerge.StartRow = 0
	'	'            objMerge.EndCol = iTemp - 1
	'	'            objMerge.EndRow = 0
	'	'
	'	'            colMerges.Add objMerge
	'	'            Set objMerge = Nothing
	'	'          End If
	'	'
	'	'          mwrdDoc.ActiveWindow.Selection.TypeParagraph
	'	'        End If
	'	'
	'	'        If Len(sTemp) > 0 Then
	'	'          ' Send the column headers to the word document.
	'	'          mwrdDoc.ActiveWindow.Selection.TypeText sTemp
	'	'          mwrdDoc.ActiveWindow.Selection.TypeParagraph
	'	'        End If
	'	'
	'	'        ' Send data rows and columns to Word.
	'	'        For iLoop = 0 To ctlTemp.Rows - 1
	'	'          fFirstColumnDone = False
	'	'          varBookmark = ctlTemp.AddItemBookmark(iLoop)
	'	'
	'	'          ' Check if the row is a Separator/Heading row.
	'	'          ' If so we'll need to create a 'style' object for it.
	'	'          If fHasHeadingColumn Then
	'	'            If ctlTemp.Columns(COLUMN_ISHEADING).CellText(varBookmark) = "1" Then
	'	'              Set objTemp = New clsOutputStyle
	'	'
	'	'              With objTemp
	'	'                .StartCol = IIf(ctlTemp.ColumnHeaders, 0, 1)
	'	'                .StartRow = lngPrecedingRows + iLoop + _
	'	''                  IIf(ctlTemp.ColumnHeaders, 1, 0) + _
	'	''                  IIf((ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders), 1, 0)
	'	'                .EndCol = -1
	'	'                .EndRow = .StartRow
	'	'
	'	'                .BackCol = pcolStyles("HeadingCols").BackCol
	'	'                .ForeCol = pcolStyles("HeadingCols").ForeCol
	'	'                .Bold = pcolStyles("HeadingCols").Bold
	'	'                .Underline = pcolStyles("HeadingCols").Underline
	'	'                .Gridlines = pcolStyles("HeadingCols").Gridlines
	'	'                .Name = "RECPROFROW_" & CStr(iLoop)
	'	'              End With
	'	'
	'	'              pcolStyles.Add objTemp
	'	'              Set objTemp = Nothing
	'	'            End If
	'	'          End If
	'	'
	'	'          For iLoop2 = 0 To ctlTemp.Columns.Count - 1
	'	'            If ctlTemp.Columns(iLoop2).Visible Then
	'	'              ' Send the text or picture to Word.
	'	'              fPhotoDone = False
	'	'              If (ctlTemp.TagVariant = COLUMN_ISPHOTO) And _
	'	''                (ctlTemp.Columns(iLoop2).Style <> 4) Then
	'	'
	'	'                For iLoop3 = 0 To ctlTemp.Columns.Count - 1
	'	'                  If ctlTemp.Columns(iLoop3).Visible Then
	'	'                    sTemp = PHOTOSTYLESET & CStr(iLoop3 + 1)
	'	'
	'	'                    For iLoop4 = 0 To ctlTemp.StyleSets.Count - 1
	'	'                      If ctlTemp.StyleSets(iLoop4).Name = sTemp Then
	'	'                        sTempName = GetTmpFName
	'	'                        SavePicture ctlTemp.StyleSets(iLoop4).Picture, sTempName
	'	'                        mwrdDoc.ActiveWindow.Selection.TypeText IIf(fFirstColumnDone, vbTab, "")
	'	'                        mwrdDoc.ActiveWindow.Selection.InlineShapes.AddPicture FileName:=sTempName, LinkToFile:=False, SaveWithDocument:=True
	'	'                        Kill sTempName
	'	'                        fPhotoDone = True
	'	'                        Exit For
	'	'                      End If
	'	'                    Next iLoop4
	'	'                  End If
	'	'
	'	'                  If fPhotoDone Then
	'	'                    Exit For
	'	'                  End If
	'	'                Next iLoop3
	'	'              End If
	'	'
	'	'              If (Not fPhotoDone) And _
	'	''                ctlTemp.Columns(iLoop2).TagVariant = COLUMN_ISPHOTO Then
	'	'
	'	'                sTemp = PHOTOSTYLESET & CStr(iLoop2 + 1) & "_" & ctlTemp.Columns(CStr(objRecProfTable.IDPosition)).Value
	'	'
	'	'                For iLoop4 = 0 To ctlTemp.StyleSets.Count - 1
	'	'                  If ctlTemp.StyleSets(iLoop4).Name = sTemp Then
	'	'                    sTempName = GetTmpFName
	'	'                    SavePicture ctlTemp.StyleSets(iLoop4).Picture, sTempName
	'	'                    mwrdDoc.ActiveWindow.Selection.TypeText IIf(fFirstColumnDone, vbTab, "")
	'	'                    mwrdDoc.ActiveWindow.Selection.InlineShapes.AddPicture FileName:=sTempName, LinkToFile:=False, SaveWithDocument:=True
	'	'                    Kill sTempName
	'	'                    fPhotoDone = True
	'	'                    Exit For
	'	'                  End If
	'	'                Next iLoop4
	'	'              End If
	'	'
	'	'              If Not fPhotoDone Then
	'	'                varBookmark = ctlTemp.AddItemBookmark(iLoop)
	'	'
	'	'                mwrdDoc.ActiveWindow.Selection.TypeText IIf(fFirstColumnDone, vbTab, "") & ctlTemp.Columns(iLoop2).CellText(varBookmark)
	'	'              End If
	'	'
	'	'              fFirstColumnDone = True
	'	'            End If
	'	'          Next iLoop2
	'	'
	'	'          mwrdDoc.ActiveWindow.Selection.TypeParagraph
	'	'        Next iLoop
	'	'
	'	'        If fGridFollowed Then
	'	'          lngPrecedingRows = lngPrecedingRows + _
	'	''            IIf(ctlTemp.ColumnHeaders, 1, 0) + _
	'	''            IIf((ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders), 1, 0) + _
	'	''            ctlTemp.Rows
	'	'        Else
	'	'          ' Convert the text into a table.
	'	'          mwrdDoc.ActiveWindow.Selection.GoTo What:=wdGoToBookmark, Name:=strBookMark
	'	'          mwrdDoc.ActiveWindow.Selection.EndKey Unit:=wdStory, Extend:=wdExtend
	'	'
	'	'          Set mwrdTable = mwrdDoc.ActiveWindow.Selection.ConvertToTable _
	'	''            (Separator:=wdSeparateByTabs, _
	'	''            NumRows:=ctlTemp.Rows + IIf(ctlTemp.ColumnHeaders, 1, 0) + IIf((ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders), 1, 0), _
	'	''            Format:=wdTableFormatNone, _
	'	''            ApplyFont:=False, _
	'	''            ApplyColor:=False, _
	'	''            AutoFit:=False)
	'	'
	'	'          With mwrdTable
	'	'            lngBorder = IIf(mblnWrdWordGridlines, wdLineStyleSingle, wdLineStyleNone)
	'	'            .Borders(wdBorderLeft).LineStyle = lngBorder
	'	'            .Borders(wdBorderRight).LineStyle = lngBorder
	'	'            .Borders(wdBorderTop).LineStyle = lngBorder
	'	'            .Borders(wdBorderBottom).LineStyle = lngBorder
	'	'            .Borders(wdBorderVertical).LineStyle = lngBorder
	'	'            .Borders(wdBorderHorizontal).LineStyle = lngBorder
	'	'          End With
	'	'
	'	'          ' Apply styles to the table.
	'	'          With pcolStyles("Heading")
	'	'            .StartCol = 0
	'	'            .StartRow = 0
	'	'            .EndCol = IIf(ctlTemp.ColumnHeaders, mwrdTable.Columns.Count, 0) - 1
	'	'            .EndRow = IIf(ctlTemp.ColumnHeaders, 1, 0) - 1 + IIf((ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders), 1, 0)
	'	'          End With
	'	'
	'	'          With pcolStyles("HeadingCols")
	'	'            .StartCol = 0
	'	'            .StartRow = 0
	'	'            .EndCol = IIf(ctlTemp.ColumnHeaders, 0, 1) - 1
	'	'            .EndRow = IIf(ctlTemp.ColumnHeaders, -1, mwrdTable.Rows.Count)
	'	'          End With
	'	'
	'	'          With pcolStyles("Data")
	'	'            .StartCol = IIf(ctlTemp.ColumnHeaders, 0, 1)
	'	'            .StartRow = IIf(ctlTemp.ColumnHeaders, 1, 0) + IIf((ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders), 1, 0)
	'	'            .EndCol = mwrdTable.Columns.Count - 1
	'	'            .EndRow = mwrdTable.Rows.Count - 1
	'	'          End With
	'	'
	'	'          With pcolStyles("Title")
	'	'            .StartCol = -1
	'	'            .StartRow = -1
	'	'            .EndCol = -1
	'	'            .EndRow = -1
	'	'          End With
	'	'
	'	'          ' Set the endRow/endCol values for the separator/heading rows/cols.
	'	'          For Each objTemp In pcolStyles
	'	'            If Left(objTemp.Name, 11) = "RECPROFCOL_" Then
	'	'              objTemp.EndRow = mwrdTable.Rows.Count - 1
	'	'            End If
	'	'            If Left(objTemp.Name, 11) = "RECPROFROW_" Then
	'	'              objTemp.EndCol = mwrdTable.Columns.Count - 1
	'	'            End If
	'	'          Next objTemp
	'	'          Set objTemp = Nothing
	'	'
	'	'          ' Apply styles and merging to the table.
	'	'          If mblnApplyStyles Then
	'	'            ApplyStyle pcolStyles
	'	'            ApplyMerges colMerges
	'	'          End If
	'	'          Set colMerges = Nothing
	'	'
	'	'          ' Get rid of any separator/heading styles that have been created
	'	'          ' for this table, so that they are not carried over and applied
	'	'          ' to other tables.
	'	'          Do While pcolStyles.Count > iOriginalStyles
	'	'            pcolStyles.Remove pcolStyles.Count
	'	'          Loop
	'	'
	'	'          mwrdApp.Selection.EndKey Unit:=wdStory
	'	'
	'	'          mwrdDoc.ActiveWindow.Selection.TypeParagraph
	'	'        End If
	'	'      End If
	'	'    End If
	'	'  Next ctlTemp
	'	'  Set ctlTemp = Nothing
	'	'
	'	'TidyUpAndExit:
	'	'  'gobjErrorStack.PopStack
	'	'  RecordProfilePage = fOK
	'	'  Exit Function
	'	'
	'	'ErrorTrap:
	'	'  'gobjErrorStack.HandleError
	'	'  fOK = False
	'	'  Resume TidyUpAndExit
	'	'
	'	'End Function


	'	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'	Private Sub Class_Initialize_Renamed()
	'		mstrWrdTemplate = gstrSettingWordTemplate
	'		mblnWrdAutoFitCols = gblnSettingAutoFitCols
	'		mblnWrdLandscape = gblnSettingLandscape
	'		mblnWrdWordGridlines = gblnSettingDataGridlines	'gblnSettingExcelGridlines
	'		mblnApplyStyles = True
	'	End Sub
	'	Public Sub New()
	'		MyBase.New()
	'		Class_Initialize_Renamed()
	'	End Sub

	'	Public WriteOnly Property Screen() As Boolean
	'		Set(ByVal Value As Boolean)
	'			mblnScreen = Value
	'		End Set
	'	End Property

	'	Public WriteOnly Property DestPrinter() As Boolean
	'		Set(ByVal Value As Boolean)
	'			mblnPrinter = Value
	'		End Set
	'	End Property

	'	Public WriteOnly Property PrinterName() As String
	'		Set(ByVal Value As String)
	'			mstrPrinterName = Value
	'		End Set
	'	End Property

	'	Public WriteOnly Property Save() As Boolean
	'		Set(ByVal Value As Boolean)
	'			mblnSave = Value
	'		End Set
	'	End Property


	'	Public Property SaveExisting() As Integer
	'		Get
	'			SaveExisting = mlngSaveExisting
	'		End Get
	'		Set(ByVal Value As Integer)
	'			mlngSaveExisting = Value
	'		End Set
	'	End Property

	'	Public WriteOnly Property Email() As Boolean
	'		Set(ByVal Value As Boolean)
	'			mblnEmail = Value
	'		End Set
	'	End Property

	'	Public WriteOnly Property FileName() As String
	'		Set(ByVal Value As String)
	'			mstrFileName = Value
	'		End Set
	'	End Property


	'	'Private Sub GetWidthFirstCol(wrdDoc As Word.Document)
	'	'
	'	'  Const strBookMark As String = "TableStart"
	'	'  Dim lngCount As Long
	'	'
	'	'  With wrdDoc.ActiveWindow.Selection.End
	'	'
	'	'    '.Font.Size = 8
	'	'    .Bookmarks.Add Range:=.Range, Name:=strBookMark
	'	'
	'	'    For lngCount = 0 To UBound(mvarHeadings(VER))
	'	'      .TypeText Trim(mvarHeadings(VER)(lngCount))
	'	'      .TypeParagraph
	'	'    Next
	'	'
	'	'    'go to start of table highlight to end of document
	'	'    .GoTo What:=wdGoToBookmark, Name:=strBookMark
	'	'    .Bookmarks.Item(strBookMark).Delete
	'	'    .EndKey Unit:=wdStory, Extend:=wdExtend
	'	'
	'	'    .ConvertToTable _
	'	''      NumRows:=UBound(mvarHeadings(VER)) + 1, _
	'	''      NumColumns:=1, _
	'	''      Format:=wdTableFormatNone, _
	'	''      ApplyFont:=False, _
	'	''      ApplyColor:=False, _
	'	''      AutoFit:=False
	'	'
	'	'    .Cells.AutoFit
	'	'    mlngFirstColWidth = .Columns(1).Width + 2
	'	'    .Tables(1).Delete
	'	'
	'	'  End With
	'	'
	'	'End Sub

	'	Public WriteOnly Property HeaderRows() As Integer
	'		Set(ByVal Value As Integer)
	'			mlngHeaderRows = Value
	'		End Set
	'	End Property

	'	Public WriteOnly Property HeaderCols() As Integer
	'		Set(ByVal Value As Integer)
	'			mlngHeaderCols = Value
	'		End Set
	'	End Property

	'	Public WriteOnly Property ApplyStyles() As Boolean
	'		Set(ByVal Value As Boolean)
	'			mblnApplyStyles = Value
	'		End Set
	'	End Property

	'	Public WriteOnly Property Parent() As clsOutputRun
	'		Set(ByVal Value As clsOutputRun)
	'			mobjParent = Value
	'		End Set
	'	End Property

	'	Public ReadOnly Property ErrorMessage() As String
	'		Get
	'			ErrorMessage = mstrErrorMessage
	'		End Get
	'	End Property


	'	Private Function CreateWordApplication() As Boolean

	'		On Error GoTo LocalErr

	'		mwrdApp = CreateObject("Word.Application")
	'		CreateWordApplication = True

	'		Exit Function

	'LocalErr:
	'		mstrErrorMessage = "Error opening Word Application"
	'		CreateWordApplication = False

	'	End Function


	'	Public Function GetFile(ByRef objParent As clsOutputRun, ByRef colStyles As Collection) As Boolean

	'		Dim strTempFileName As String
	'		Dim lngFound As Integer
	'		Dim lngCount As Integer

	'		On Error GoTo LocalErr

	'		If Not CreateWordApplication() Then
	'			GetFile = False
	'			Exit Function
	'		End If

	'		'Just in case we are emailing but not saving...
	'		If mblnEmail And Not mblnSave Then
	'			mstrFileName = objParent.GetTempFileName(mstrFileName)
	'		End If


	'		' Leave the app there after user has closed the worksheet
	'		'mwrdApp.UserControl = True
	'		mwrdApp.DisplayAlerts = False

	'		'Check if file already exists...
	'		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
	'		If Dir(mstrFileName) <> vbNullString And mstrFileName <> vbNullString Then

	'			Select Case mlngSaveExisting
	'				Case 0 'Overwrite
	'					If Not objParent.KillFile(mstrFileName) Then
	'						GetFile = False
	'						Exit Function
	'					End If
	'					GetNewDocument()

	'				Case 1 'Do not overwrite (fail)
	'					mwrdApp.Quit()
	'					'UPGRADE_NOTE: Object mwrdApp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	'					mwrdApp = Nothing
	'					mstrErrorMessage = "File already exists."

	'				Case 2 'Add Sequential number to file
	'					mstrFileName = mobjParent.GetSequentialNumberedFile(mstrFileName)
	'					GetNewDocument()

	'				Case 3 'Append to existing file

	'					If Not IsFileCompatibleWithWordVersion(mstrFileName, Val(mwrdApp.Version)) Then
	'						mstrErrorMessage = "This definition is set to append to a file which is not compatible with your version of Microsoft Office."
	'						GetFile = False
	'						Exit Function
	'					End If

	'					'Start at the bottom of the document
	'					mwrdDoc = mwrdApp.Documents.Open(mstrFileName)
	'					mwrdApp.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
	'					mwrdApp.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak)

	'				Case 4 'Create new worksheet within existing workbook...
	'					'N/A (EXCEL ONLY)

	'			End Select

	'		Else
	'			GetNewDocument()

	'		End If

	'		GetFile = (mstrErrorMessage = vbNullString)

	'		Exit Function

	'LocalErr:
	'		mstrErrorMessage = Err.Description
	'		GetFile = False

	'	End Function


	'	Private Sub GetNewDocument()

	'		Dim mwrdFooterTable As Microsoft.Office.Interop.Word.Table
	'		Dim strTempFile As String
	'		Dim lngView As Integer
	'		Dim lngCount As Integer

	'		On Error GoTo LocalErr

	'		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
	'		If mstrWrdTemplate <> "" And Dir(mstrWrdTemplate) <> "" Then
	'			'Set mwrdDoc = mwrdApp.Documents.Open(mstrWrdTemplate, ReadOnly:=True)

	'			If Not IsFileCompatibleWithWordVersion(mstrWrdTemplate, Val(mwrdApp.Version)) Then
	'				mstrErrorMessage = "Your User Configuration Output options are set to use a template file which is not compatible with your version of Microsoft Office."
	'				Exit Sub
	'			End If

	'			'MH20030905 Fault 6911
	'			'If Word 2000 then make a copy of the template
	'			If Val(mwrdApp.Version) >= 9 And Val(mwrdApp.Version) < 10 Then
	'				strTempFile = mobjParent.GetTempFileName(vbNullString)
	'				FileCopy(mstrWrdTemplate, strTempFile)
	'			Else
	'				strTempFile = mstrWrdTemplate
	'			End If

	'			mwrdDoc = mwrdApp.Documents.Add(strTempFile)
	'			mwrdApp.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
	'		Else
	'			mwrdDoc = mwrdApp.Documents.Add

	'			''Insert heading (slightly bigger, bold, underline and centered)
	'			'.Font.Bold = True
	'			'.Font.Size = .Font.Size + 2
	'			'.Font.Underline = wdUnderlineSingle
	'			'.ParagraphFormat.Alignment = wdAlignParagraphCenter

	'			mwrdDoc.PageSetup.Orientation = IIf(mblnWrdLandscape, Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape, Microsoft.Office.Interop.Word.WdOrientation.wdOrientPortrait)

	'			With mwrdApp
	'				lngView = .ActiveWindow.View.Type
	'				.ActiveDocument.PageSetup.OddAndEvenPagesHeaderFooter = False
	'				.ActiveDocument.PageSetup.DifferentFirstPageHeaderFooter = False
	'				.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewTypeOld.wdPageView
	'				.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekPrimaryFooter

	'				mwrdFooterTable = mwrdDoc.Tables.Add(Range:=.Selection.Range, NumRows:=1, NumColumns:=2)

	'				.Selection.TypeText(Text:="Created on ")
	'				'.Selection.Fields.Add .Selection.Range, Type:=wdFieldDate,
	'				'Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:=
	'				.Selection.Fields.Add(Range:=.Selection.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldDate, Text:="CREATEDATE \@ ""dd/MM/yyyy at hh:mm"" ", PreserveFormatting:=True)
	'				.Selection.TypeText(Text:=" by " & UserName)
	'				.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)

	'				.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight
	'				.Selection.TypeText(Text:="Page ")
	'				.Selection.Fields.Add(.Selection.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage)

	'				With mwrdFooterTable
	'					.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
	'					.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
	'				End With

	'				.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument
	'				.ActiveWindow.View.Type = lngView

	'			End With

	'		End If

	'		'UPGRADE_NOTE: Object mwrdFooterTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	'		mwrdFooterTable = Nothing

	'		Exit Sub

	'LocalErr:
	'		'UPGRADE_NOTE: Object mwrdFooterTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	'		mwrdFooterTable = Nothing
	'		mstrErrorMessage = Err.Description

	'	End Sub


	'	Public Sub AddPage(ByRef strDefTitle As String, ByRef mstrSheetName As String, ByRef colStyles As Collection)

	'		Const strBookMark As String = "ASRSysTitleStart"
	'		Dim objRange As Microsoft.Office.Interop.Word.Range
	'		Dim objStyle As clsOutputStyle

	'		On Error GoTo LocalErr

	'		mwrdDoc.Bookmarks.Add(strBookMark)

	'		With mwrdDoc.ActiveWindow.Selection

	'			'Remember the current setting to restore them later...
	'			objStyle = GetCurrentStyle(mwrdDoc.ActiveWindow.Selection)


	'			'If its the first page then add the definition title...
	'			mlngPageCount = mlngPageCount + 1
	'			If mlngPageCount = 1 Then
	'				.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
	'				.TypeText(strDefTitle)
	'				.TypeParagraph()
	'				.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
	'				.TypeParagraph()
	'			Else
	'				mwrdApp.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak)
	'			End If

	'			If mstrSheetName <> vbNullString And mobjParent.PageTitles Then
	'				.TypeText(mstrSheetName)
	'				.TypeParagraph()
	'				.TypeParagraph()
	'			End If

	'			'Move to the begining of the title, then highlight to end of document (minus one line)
	'			.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:=strBookMark)
	'			.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
	'			'.MoveUp Unit:=wdLine, Count:=1, Extend:=wdExtend

	'			'Apply the "title" style to this range
	'			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'			ApplyStylesToRange(mwrdApp.Selection.Range, colStyles.Item("Title"), True)

	'			'Now move to the end of the document, store the formatting, ready for the data...
	'			.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
	'			ApplyStylesToRange(mwrdApp.Selection.Range, objStyle, False)

	'		End With

	'		'UPGRADE_NOTE: Object objRange may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	'		objRange = Nothing

	'		Exit Sub

	'LocalErr:
	'		'UPGRADE_NOTE: Object objRange may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	'		objRange = Nothing
	'		mstrErrorMessage = Err.Description

	'	End Sub

	'	Private Function GetCurrentStyle(ByRef objRange As Object) As clsOutputStyle

	'		Dim objStyle As clsOutputStyle

	'		

	'		objStyle = New clsOutputStyle

	'		With objRange
	'			'UPGRADE_WARNING: Couldn't resolve default property of object objRange.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'			objStyle.Bold = .Font.Bold
	'			'UPGRADE_WARNING: Couldn't resolve default property of object objRange.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'			objStyle.Underline = .Font.Underline
	'			'UPGRADE_WARNING: Couldn't resolve default property of object objRange.ParagraphFormat. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'			objStyle.CenterText = (.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter)

	'			If Val(mwrdApp.Version) < 9 Then
	'				'UPGRADE_WARNING: Couldn't resolve default property of object objRange.Shading. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'				objStyle.BackCol = .Shading.BackgroundPatternColorIndex
	'				'UPGRADE_WARNING: Couldn't resolve default property of object objRange.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'				objStyle.ForeCol = .Font.ForegroundPatternColorIndex
	'			Else
	'				'UPGRADE_WARNING: Couldn't resolve default property of object objRange.Shading. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'				objStyle.BackCol = .Shading.BackgroundPatternColor
	'				'UPGRADE_WARNING: Couldn't resolve default property of object objRange.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'				objStyle.ForeCol = .Font.ForegroundPatternColor
	'			End If

	'		End With

	'		GetCurrentStyle = objStyle
	'		'UPGRADE_NOTE: Object objStyle may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	'		objStyle = Nothing

	'	End Function

	'	Public Sub DataArray(ByRef strArray(,) As String, ByRef colColumns As Collection, ByRef colStyles As Collection, ByRef colMerges As Collection)

	'		Const strBookMark As String = "ASRSysTableStart"
	'		Dim strOutput As String
	'		Dim lngGridCol As Integer
	'		Dim lngGridRow As Integer
	'		Dim lngBorder As Integer

	'		On Error GoTo LocalErr

	'		mwrdDoc.Bookmarks.Add(strBookMark)

	'		For lngGridRow = 0 To UBound(strArray, 2)
	'			strOutput = vbNullString
	'			For lngGridCol = 0 To UBound(strArray, 1)
	'				strOutput = IIf(lngGridCol > 0, strOutput & vbTab, "") & strArray(lngGridCol, lngGridRow)

	'				'      If gobjProgress.Visible And gobjProgress.Cancelled Then
	'				'        mstrErrorMessage = "Cancelled by user."
	'				'        Exit Sub
	'				'      End If

	'			Next
	'			mwrdDoc.ActiveWindow.Selection.TypeText(strOutput)
	'			mwrdDoc.ActiveWindow.Selection.TypeParagraph()

	'		Next

	'		mwrdDoc.ActiveWindow.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:=strBookMark)
	'		mwrdDoc.ActiveWindow.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)

	'		'convert selected text into a table
	'		mwrdTable = mwrdDoc.ActiveWindow.Selection.ConvertToTable(Separator:=Microsoft.Office.Interop.Word.WdTableFieldSeparator.wdSeparateByTabs, NumRows:=UBound(strArray, 2), Format:=Microsoft.Office.Interop.Word.WdTableFormat.wdTableFormatNone, ApplyFont:=False, ApplyColor:=False, AutoFit:=False)


	'		For lngGridCol = 0 To UBound(strArray, 1)
	'			mwrdTable.Columns.Item(lngGridCol + 1).Select()
	'			'UPGRADE_WARNING: Couldn't resolve default property of object colColumns(lngGridCol + 1).DataType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'			Select Case colColumns.Item(lngGridCol + 1).DataType
	'				Case SQLDataType.sqlNumeric, SQLDataType.sqlInteger
	'					mwrdApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight
	'				Case SQLDataType.sqlBoolean
	'					mwrdApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
	'			End Select
	'		Next


	'		With mwrdTable
	'			lngBorder = IIf(mblnWrdWordGridlines, Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle, Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone)
	'			.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle = lngBorder
	'			.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle = lngBorder
	'			.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = lngBorder
	'			.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = lngBorder
	'			.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical).LineStyle = lngBorder
	'			.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = lngBorder
	'		End With

	'		With colStyles.Item("Title")
	'			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'			.StartCol = -1
	'			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'			.StartRow = -1
	'			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().EndCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'			.EndCol = -1
	'			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().EndRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'			.EndRow = -1
	'		End With


	'		With colStyles.Item("Heading")
	'			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'			.StartCol = 0
	'			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'			.StartRow = 0
	'			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().EndCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'			.EndCol = mwrdTable.Columns.Count - 1
	'			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().EndRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'			.EndRow = mlngHeaderRows - 1
	'		End With


	'		If mlngHeaderCols > 0 Then
	'			With colStyles.Item("HeadingCols")
	'				'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'				.StartCol = 0
	'				'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'				.StartRow = 0
	'				'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().EndCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'				.EndCol = mlngHeaderCols - 1
	'				'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().EndRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'				.EndRow = mwrdTable.Rows.Count - 1
	'			End With
	'		End If


	'		With colStyles.Item("Data")
	'			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'			.StartCol = mlngHeaderCols
	'			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().StartRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'			.StartRow = mlngHeaderRows
	'			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().EndCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'			.EndCol = mwrdTable.Columns.Count - 1
	'			'UPGRADE_WARNING: Couldn't resolve default property of object colStyles().EndRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'			.EndRow = mwrdTable.Rows.Count - 1
	'		End With


	'		'If mstrWrdTemplate = vbNullString Then
	'		If mblnApplyStyles Then
	'			ApplyStyle(colStyles)
	'			ApplyMerges(colMerges)
	'		End If


	'		'mwrdTable.Columns(0).Select

	'		'mwrdApp.Selection.Tables(1).AutoFitBehavior wdAutoFitWindow

	'		'With colStyles("Data")
	'		'  mwrdDoc.Range(mwrdTable.Cell(.StartRow + 1, .StartCol + 1).Range.Start, mwrdTable.Cell(.EndRow + 1, .EndCol + 1).Range.End).Select
	'		'  mwrdApp.Selection.Cells.DistributeWidth
	'		'End With

	'		mwrdApp.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
	'		mwrdDoc.ActiveWindow.Selection.TypeParagraph()

	'		Exit Sub

	'LocalErr:
	'		If Err.Number = 4608 Then
	'			mstrErrorMessage = "Cannot output more than 60 columns to Word"
	'		Else
	'			mstrErrorMessage = Err.Description
	'		End If

	'	End Sub


	'	Private Sub PrepareRows(ByRef lngStartRow As Integer, ByRef lngRowCount As Integer)

	'		'  Dim objColumn As clsColumn
	'		'  Dim lngStartCol As Long
	'		'  Dim lngEndRow As Long
	'		'  Dim lngCount As Long
	'		'
	'		'  With mwrdWorkSheet
	'		'
	'		'    If mstrWrdTemplate <> vbNullString Then
	'		'      For lngCount = 1 To lngRowCount
	'		'        .Rows(lngStartRow).Select
	'		'        mwrdApp.Selection.Copy
	'		'        mwrdApp.Selection.Insert Shift:=-4121 'xlDown
	'		'      Next
	'		'      mwrdApp.CutCopyMode = False
	'		'    End If
	'		'    .Range("A1").Select
	'		'
	'		'
	'		'    lngStartCol = mwrdData.StartCol - 1
	'		'    lngEndRow = lngStartRow + lngRowCount
	'		'    For lngCount = 1 To mcolColumns.Count
	'		'
	'		'      Set objColumn = mcolColumns(lngCount)
	'		'      With .Range(.Cells(lngStartRow, lngStartCol + lngCount), .Cells(lngEndRow, lngStartCol + lngCount))
	'		'        Select Case objColumn.DataType
	'		'        Case sqlNumeric, sqlInteger
	'		'          .NumberFormat = "0" & IIf(objColumn.DecPlaces, "." & String(objColumn.DecPlaces, "0"), "")
	'		'        Case Else
	'		'          .NumberFormat = "@"   'Dates need to be exported as text as Word insists on changing some dates!
	'		'        End Select
	'		'      End With
	'		'
	'		'    Next
	'		'
	'		'  End With

	'	End Sub


	'	Private Sub ApplyStyle(ByRef colStyles As Collection)

	'		Dim objTemp As clsOutputStyle
	'		'Dim objRange As Object
	'		'Dim lngBorder As Long
	'		'Dim lngCol As Long
	'		'Dim lngRow As Long


	'		For Each objTemp In colStyles

	'			'Must do it row by row otherwise other cells get included in the formatting...
	'			'For lngRow = objTemp.StartRow To objTemp.EndRow
	'			'  Set objRange = mwrdDoc.Range(mwrdTable.Cell(lngRow + 1, objTemp.StartCol + 1).Range.Start, mwrdTable.Cell(lngRow + 1, objTemp.EndCol + 1).Range.End)

	'			'JPD 20030404 Changed the endCol/endRow conditions as we might
	'			'want to apply a style to the top-left (0,0) cell, as we do sometimes in Record Profile.
	'			'If objTemp.EndCol > 0 Or objTemp.EndRow > 0 Then
	'			If objTemp.EndCol >= 0 Or objTemp.EndRow >= 0 Then
	'				mwrdDoc.Range(mwrdTable.Cell(objTemp.StartRow + 1, objTemp.StartCol + 1).Range.Start, mwrdTable.Cell(objTemp.EndRow + 1, objTemp.EndCol + 1).Range.End).Select()
	'				ApplyStylesToRange((mwrdApp.Selection), objTemp, True)
	'			End If

	'		Next objTemp

	'	End Sub


	'	Private Sub ApplyStylesToRange(ByRef objRange As Object, ByRef objStyle As clsOutputStyle, ByRef blnColourIndex As Boolean)

	'		

	'		With objRange
	'			'UPGRADE_WARNING: Couldn't resolve default property of object objRange.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'			.Font.Bold = objStyle.Bold
	'			'UPGRADE_WARNING: Couldn't resolve default property of object objRange.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'			.Font.Underline = objStyle.Underline

	'			If objStyle.CenterText Then
	'				'UPGRADE_WARNING: Couldn't resolve default property of object objRange.ParagraphFormat. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'				.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
	'			End If

	'			If Val(mwrdApp.Version) < 9 Then
	'				If blnColourIndex Then
	'					'UPGRADE_WARNING: Couldn't resolve default property of object objRange.Shading. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'					.Shading.BackgroundPatternColorIndex = objStyle.BackCol97
	'					'UPGRADE_WARNING: Couldn't resolve default property of object objRange.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'					.Font.ColorIndex = objStyle.ForeCol97
	'				Else
	'					'UPGRADE_WARNING: Couldn't resolve default property of object objRange.Shading. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'					.Shading.BackgroundPatternColorIndex = objStyle.BackCol
	'					'UPGRADE_WARNING: Couldn't resolve default property of object objRange.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'					.Font.ColorIndex = objStyle.ForeCol
	'				End If
	'			Else
	'				'UPGRADE_WARNING: Couldn't resolve default property of object objRange.Shading. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'				.Shading.BackgroundPatternColor = objStyle.BackCol
	'				'UPGRADE_WARNING: Couldn't resolve default property of object objRange.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'				.Font.Color = objStyle.ForeCol
	'			End If

	'			'MH20050907 Fault 10319 & 10320
	'			If objStyle.Name = "Heading" Then
	'				'UPGRADE_WARNING: Couldn't resolve default property of object objRange.Rows. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'				.Rows.HeadingFormat = True
	'			End If

	'		End With

	'	End Sub

	'	Private Sub ApplyMerges(ByRef colMerges As Collection)

	'		Dim objTemp As clsOutputStyle
	'		Dim objRange As Object
	'		Dim lngBorder As Integer
	'		Dim lngCol As Integer
	'		Dim lngRow As Integer
	'		Dim lngOffset As Integer

	'		On Error GoTo LocalErr

	'		lngOffset = 1
	'		For Each objTemp In colMerges
	'			objRange = mwrdDoc.Range(mwrdTable.Cell(objTemp.StartRow + 1, objTemp.StartCol + lngOffset).Range.Start, mwrdTable.Cell(objTemp.EndRow + 1, objTemp.EndCol + lngOffset).Range.End)
	'			'UPGRADE_WARNING: Couldn't resolve default property of object objRange.Select. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'			objRange.Select()
	'			mwrdApp.ActiveWindow.Selection.Cells.Merge()
	'			lngOffset = lngOffset - (objTemp.EndCol - objTemp.StartCol)
	'		Next objTemp

	'		'UPGRADE_NOTE: Object objRange may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	'		objRange = Nothing

	'		Exit Sub

	'LocalErr:
	'		'UPGRADE_NOTE: Object objRange may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	'		objRange = Nothing
	'		mstrErrorMessage = Err.Description

	'	End Sub


	'	Public Sub Complete()

	'		Dim strDefaultPrinter As String
	'		Dim strFormat As String
	'		Dim strTempFile As String
	'		Dim strExtension As String
	'		Dim aryFileBits() As String

	'		On Error GoTo LocalErr

	'		If mstrErrorMessage <> vbNullString Then
	'			Exit Sub
	'		End If

	'		'SAVE
	'		If mblnSave Then
	'			mstrErrorMessage = "Error saving file <" & mstrFileName & ">"

	'			strFormat = GetOfficeSaveAsFormat(mstrFileName, Val(mwrdApp.Version), modIntClient.OfficeApp.oaWord)
	'			If strFormat = "" Then
	'				mstrErrorMessage = "This definition is set to save in a file format which is not compatible with your version of Microsoft Office."
	'				GoTo TidyAndExit
	'			End If

	'			' calculate the appropriate output type
	'			aryFileBits = Split(mstrFileName, ".")
	'			strExtension = aryFileBits(UBound(aryFileBits))

	'			Select Case LCase(strExtension)
	'				Case "pdf"
	'					mwrdDoc.SaveAs(mstrFileName, FileFormat:=Val(CStr(modIntClient.WordOutputType.wdFormatPDF)))
	'				Case "docx"
	'					mwrdDoc.SaveAs(mstrFileName, FileFormat:=Val(CStr(modIntClient.WordOutputType.wdFormatXMLDocument)))
	'				Case "doc"
	'					mwrdDoc.SaveAs(mstrFileName, FileFormat:=Val(CStr(modIntClient.WordOutputType.wdFormatDocument97)))
	'				Case "txt"
	'					mwrdDoc.SaveAs(mstrFileName, FileFormat:=Val(CStr(modIntClient.WordOutputType.wdFormatText)))
	'				Case "rtf"
	'					mwrdDoc.SaveAs(mstrFileName, FileFormat:=Val(CStr(modIntClient.WordOutputType.wdFormatRTF)))
	'				Case "html"
	'					mwrdDoc.SaveAs(mstrFileName, FileFormat:=Val(CStr(modIntClient.WordOutputType.wdFormatHTML)))
	'			End Select

	'		End If

	'		'EMAIL
	'		If mblnEmail Then
	'			mstrErrorMessage = "Error sending email"

	'			strFormat = GetOfficeSaveAsFormat((mobjParent.EmailAttachAs), Val(mwrdApp.Version), modIntClient.OfficeApp.oaWord)
	'			If strFormat = "" Then
	'				mstrErrorMessage = "This definition is set to email an attachment in a file format which is not compatible with your version of Microsoft Office."
	'				GoTo TidyAndExit
	'			End If

	'			strTempFile = mobjParent.GetTempFileName((mobjParent.EmailAttachAs))
	'			mwrdApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone
	'			mwrdDoc.SaveAs(strTempFile, FileFormat:=Val(strFormat))

	'			mwrdDoc.Close(False)
	'			mobjParent.SendEmail(strTempFile)
	'			mwrdDoc = mwrdApp.Documents.Open(strTempFile)
	'		End If

	'		'PRINTER
	'		Dim strCurrentPrinter As String
	'		If mblnPrinter Then
	'			'TM23122003 FAULT - DEFAULT PRINTER
	'			mstrErrorMessage = "Error printing"
	'			'mobjParent.SetPrinter
	'			strCurrentPrinter = mwrdApp.ActivePrinter
	'			mwrdApp.ActivePrinter = mstrPrinterName
	'			mwrdDoc.PrintOut()
	'			mwrdApp.ActivePrinter = strCurrentPrinter
	'		End If


	'		'SCREEN
	'		If mblnScreen Then
	'			mstrErrorMessage = "Error displaying Word"
	'			mwrdApp.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
	'			mwrdApp.Visible = True
	'			mwrdApp.WindowState = Microsoft.Office.Interop.Word.WdWindowState.wdWindowStateMaximize
	'			mwrdApp.Activate()
	'			'UPGRADE_NOTE: Object mwrdApp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	'			mwrdApp = Nothing	'Stops word quitting!
	'		Else
	'			Do While mwrdApp.BackgroundPrintingStatus > 0 Or mwrdApp.BackgroundSavingStatus > 0
	'				System.Windows.Forms.Application.DoEvents()
	'			Loop

	'			mwrdDoc.Close(False)
	'			'mwrdApp.Quit
	'		End If

	'		mstrErrorMessage = vbNullString

	'TidyAndExit:
	'		ClearUp()

	'		Exit Sub

	'LocalErr:
	'		mstrErrorMessage = mstrErrorMessage & IIf(Err.Description <> vbNullString, vbCrLf & " (" & Err.Description & ")", vbNullString)
	'		Resume TidyAndExit

	'	End Sub
End Class