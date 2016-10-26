VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOutputExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mxlApp As Excel.Application
Private mxlWorkBook As Excel.Workbook
Private mxlWorkSheet As Excel.Worksheet
Private mxlTemplateBook As Excel.Workbook
Private mxlTemplateSheet As Excel.Worksheet
Private mxlFirstSheet As Object
Private mxlDeleteSheet As Object
Private mobjParent As clsOutputRun

Private mlngHeaderRows As Long
Private mlngHeaderCols As Long
Private mlngDataCurrentRow As Long
Private mlngDataStartRow As Long
Private mlngDataStartCol As Long

Private mblnScreen As Boolean
Private mblnPrinter As Boolean
Private mstrPrinterName As String
Private mblnSave As Boolean
Private mlngSaveExisting As Long
Private mblnEmail As Boolean
Private mstrFileName As String
Private mblnSizeColumnsIndependently As Boolean
Private mblnApplyStyles As Boolean

Private mstrSheetMode As String
Private mblnAppending As Boolean
Private mlngAppendStartRow As Long

Private mstrDefTitle As String
Private mstrErrorMessage As String

Private mblnChart As Boolean
Private mblnPivotTable As Boolean
Private mstrPackTemplate As String
Private mstrXLTemplate As String
Private mblnXLExcelGridlines As Boolean
Private mblnXLExcelOmitTopRow As Boolean
Private mblnXLExcelOmitLeftCol As Boolean
Private mblnXLExcelHeaders As Boolean
Private mblnXLAutoFitCols As Boolean
Private mblnXLLandscape As Boolean

Public Sub ClearUp()
  On Local Error Resume Next
  'Always close the template...
  If Not mxlTemplateBook Is Nothing Then
    mxlTemplateBook.Saved = True
    mxlTemplateBook.Close
  End If

  'If error then close the workbook and app...
  If Not mxlWorkBook Is Nothing Then
    mxlWorkBook.Saved = True
    mxlWorkBook.Close
  End If
  If Not mxlApp Is Nothing Then
    mxlApp.Quit
  End If
  
  'Reset all references to ensure that Excel closes cleanly...
  Set mxlTemplateSheet = Nothing
  Set mxlTemplateBook = Nothing
  Set mxlDeleteSheet = Nothing
  Set mxlFirstSheet = Nothing
  Set mxlWorkSheet = Nothing
  Set mxlWorkBook = Nothing
  Set mxlApp = Nothing

End Sub

Public Function RecordProfilePage(pfrmRecProfile As frmRecordProfilePreview, _
  piPageNumber As Integer, _
  pcolStyles As Collection)
  ' Output the record profile page to Excel.
  
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "clsOutputExcel.RecordProfilePage()"

  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim iLoop3 As Integer
  Dim iLoop4 As Integer
  Dim ctlTemp As Control
  Dim varBookmark As Variant
  Dim fGridPreceded As Boolean
  Dim fGridFollowed As Boolean
  Dim sTitle As String
  Dim sTemp As String
  Dim fPhotoDone As Boolean
  Dim objRecProfTable As clsRecordProfileTabDtl
  Dim sTempName As String
  Dim lngBorder As Long
  Dim iLastGroup As Integer
  Dim colMerges As Collection
  Dim objMerge As clsOutputStyle
  Dim iGroupStart As Integer
  Dim iTemp As Integer
  Dim objTemp As clsOutputStyle
  Dim iOriginalStyles As Integer
  Dim fHasHeadingColumn As Boolean
  Dim lngXLCol As Long
  Dim lngXLRow As Long
  Dim lngXLGridStartRow As Long
  Dim lngXLGridEndRow As Long
  Dim lngXLGridStartCol As Long
  Dim lngXLGridEndCol As Long
  Dim objRange As Excel.Range
  Dim lngMaxXLCol As Long
  Dim lngMaxXLRow As Long
  Dim lngMaxWidth As Long
  Dim lngCount As Long
  Dim alngPictureRows() As Long
  Dim alngPictureCols() As Long
  Dim fFound As Boolean
  Dim lngPictureHeight As Long
  Dim lngPictureWidth As Long
  Dim iOriginalScaleMode As ScaleModeConstants
  Dim iDecPlaces As Integer
  Dim bBlankIfZero As Boolean
  Dim bUse1000Separators As Boolean
  Dim lngTabIndex As Long         'MH20040625 Fault 8782
  Dim alngControlOrder() As Long
  Dim fPageDone As Boolean

  Const RECPROFFOLLOWONCORRECTION = 10

  Const COLUMN_ISHEADING = "IsHeading"
  Const COLUMN_ISPHOTO = "IsPhoto"
  Const PHOTOSTYLESET = "PhotoSS_"
  Const COLUMN_DECPLACES = "DecPlaces"
  Const COLUMN_BLANKIFZERO = "BlankIfZero"
  Const COLUMN_USE1000SEPARATOR = "ThousandSeparator"

  Const EXCEL_MAXROWHEIGHT = 409

  fOK = True
  sTitle = pfrmRecProfile.Caption
  iOriginalStyles = pcolStyles.Count
  lngXLCol = 2
  
  lngXLRow = 4
  'lngXLRow = mlngDataCurrentRow
  'lngXLRow = IIf(mlngDataCurrentRow < 1, 1, mlngDataCurrentRow)

  lngMaxXLRow = lngXLRow
  lngMaxXLCol = lngXLCol
  
  ' Initialise the hidden column sizing label in the preview screen
  ' with the same font as the Excel worksheet. We use this label to
  ' work out the required column width for pictures that are output to Excel,
  ' as we need to translate the picture's width into the number of characters
  ' that fit into that width.
  iOriginalScaleMode = pfrmRecProfile.ScaleMode
  With mxlWorkBook.ActiveSheet.Cells(1, 1)
  'With mxlWorkBook.Worksheets(1).Cells(1, 1)
    .Font.Bold = pcolStyles("Data").Bold
    pfrmRecProfile.lblColumnSizingLabel.Font.Name = .Font.Name
    pfrmRecProfile.lblColumnSizingLabel.Font.Size = .Font.Size
    pfrmRecProfile.lblColumnSizingLabel.Font.Bold = .Font.Bold
  End With

  ' Dimension an array to hold info for the Excel rows that contain pictures.
  ' Column 1 - Excel row
  ' Column 2 - Max picture height
  ReDim alngPictureRows(2, 0)
  
  ' Dimension an array to hold info for the Excel columns that contain pictures.
  ' Column 1 - Excel col
  ' Column 2 - Max picture width
  ReDim alngPictureCols(2, 0)
  
  alngControlOrder = pfrmRecProfile.ControlOrder
  fPageDone = False
  
  'Force the controls to be processed in the correct tab order...
  For lngTabIndex = 0 To UBound(alngControlOrder, 2)
    If alngControlOrder(1, lngTabIndex) = piPageNumber Then
      fPageDone = True
  
      '
      ' LABEL control
      '
      If alngControlOrder(2, lngTabIndex) = 0 Then
        Set ctlTemp = pfrmRecProfile.lblCaption(alngControlOrder(3, lngTabIndex))
        
        ' Write the label's caption to Excel.
        lngXLCol = 2
        
        'mxlWorkBook.Worksheets(1).Cells(lngXLRow, lngXLCol).FormulaR1C1 = ctlTemp.Caption
        mxlWorkBook.ActiveSheet.Cells(lngXLRow, lngXLCol).FormulaR1C1 = ctlTemp.Caption
        
        lngMaxXLRow = IIf(lngMaxXLRow < lngXLRow, lngXLRow, lngMaxXLRow)
        lngMaxXLCol = IIf(lngMaxXLCol < lngXLCol, lngXLCol, lngMaxXLCol)
        
        'Set objRange = mxlWorkBook.Worksheets(1).Cells(lngXLRow, lngXLCol)
        Set objRange = mxlWorkBook.ActiveSheet.Cells(lngXLRow, lngXLCol)
        ApplyStyleToRange objRange, pcolStyles("Title")

    
        lngXLRow = lngXLRow + 2
      Else
        Set ctlTemp = pfrmRecProfile.grdOutput(alngControlOrder(3, lngTabIndex))
        '
        ' GRID control
        '
        Set objRecProfTable = pfrmRecProfile.Definition.item(ctlTemp.Tag)

        ' Check if this grid is preceded or followed IMMEDIATELY by other grids.
        ' ie. if this grid is part of a group of grids that are used to display
        ' data (including pictures) vertically.
        ' NB. Grids have only one row height value. To display pictures with their
        ' own row height, we actually put them in their own grid, and position this grid
        ' IMMEDIATELY after the normal data grid. Subsequent data is put in its own grid
        ' IMMEDIATELY after the picture's grid.
        ' This is what's meant by 'following' and 'preceding' grids.
        fGridFollowed = False
        fGridPreceded = False
    
        If ctlTemp.Index > 1 Then
          If (ctlTemp.Container = pfrmRecProfile.grdOutput(ctlTemp.Index - 1).Container) And _
            (ctlTemp.Top = (pfrmRecProfile.grdOutput(ctlTemp.Index - 1).Top + pfrmRecProfile.grdOutput(ctlTemp.Index - 1).Height - RECPROFFOLLOWONCORRECTION)) Then

            fGridPreceded = True
          End If
        End If

        If ctlTemp.Index < pfrmRecProfile.grdOutput.Count - 1 Then
          If (ctlTemp.Container = pfrmRecProfile.grdOutput(ctlTemp.Index + 1).Container) And _
            ((ctlTemp.Top + ctlTemp.Height - RECPROFFOLLOWONCORRECTION) = pfrmRecProfile.grdOutput(ctlTemp.Index + 1).Top) Then

            fGridFollowed = True
          End If
        End If

        lngXLCol = 2
            
        If Not fGridPreceded Then
          lngXLGridStartRow = lngXLRow
          lngXLGridEndRow = -1
          lngXLGridStartCol = lngXLCol
          lngXLGridEndCol = -1
          Set colMerges = New Collection
        End If
        
        ' Send the column/group headers to the Excel document.
        fHasHeadingColumn = False
        iLastGroup = -1
        iTemp = lngXLCol
        iGroupStart = lngXLCol
        
        For iLoop = 0 To ctlTemp.Columns.Count - 1
          If ctlTemp.Columns(iLoop).Name = COLUMN_ISHEADING Then
            fHasHeadingColumn = True
          End If
            
          If (ctlTemp.Columns(iLoop).Visible) Then
            If (ctlTemp.ColumnHeaders) Then
              If (ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders) Then
                If (iLastGroup <> ctlTemp.Columns(iLoop).Group) Then
                  ' Send the group header to the Excel document.
                  'mxlWorkBook.Worksheets(1).Cells(lngXLRow, lngXLCol).FormulaR1C1 = ctlTemp.Groups(ctlTemp.Columns(iLoop).Group).Caption
                  mxlWorkBook.ActiveSheet.Cells(lngXLRow, lngXLCol).FormulaR1C1 = ctlTemp.Groups(ctlTemp.Columns(iLoop).Group).Caption
                  ' Remember if the group heading cells need to be merged.
                  If iGroupStart < iTemp - 1 Then
                    Set objMerge = New clsOutputStyle
                    objMerge.StartCol = iGroupStart
                    objMerge.StartRow = lngXLRow
                    objMerge.EndCol = iTemp - 1
                    objMerge.EndRow = lngXLRow
                      
                    colMerges.Add objMerge
                    Set objMerge = Nothing
                  End If
                  
                  iGroupStart = iTemp
                End If
                
                iLastGroup = ctlTemp.Columns(iLoop).Group
              End If
  
              ' Check if the column is a Separator column.
              ' If so we'll need to create a 'style' object for it.
              If ctlTemp.Columns(iLoop).StyleSet = "Separator" Then
                Set objTemp = New clsOutputStyle
              
                With objTemp
                  .StartCol = lngXLCol
                  .StartRow = lngXLRow + IIf(ctlTemp.ColumnHeaders, 1, 0) + IIf((ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders), 1, 0)
                  .EndCol = lngXLCol
                  .EndRow = -1
              
                  .BackCol = pcolStyles("HeadingCols").BackCol
                  .ForeCol = pcolStyles("HeadingCols").ForeCol
                  .Bold = pcolStyles("HeadingCols").Bold
                  .Underline = pcolStyles("HeadingCols").Underline
                  .Gridlines = pcolStyles("HeadingCols").Gridlines
                  .Name = "RECPROFCOL_" & CStr(iLoop)
                End With
              
                pcolStyles.Add objTemp
                Set objTemp = Nothing
              End If
              
              ' Send the column header to the Excel document.
              'mxlWorkBook.Worksheets(1).Cells(lngXLRow + IIf((ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders), 1, 0), lngXLCol).FormulaR1C1 = ctlTemp.Columns(iLoop).Caption
              mxlWorkBook.ActiveSheet.Cells(lngXLRow + IIf((ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders), 1, 0), lngXLCol).FormulaR1C1 = ctlTemp.Columns(iLoop).Caption
              iTemp = iTemp + 1
            End If
            
            lngXLGridEndCol = lngXLCol
            lngXLCol = lngXLCol + 1
          End If
        Next iLoop
          
        ' Remember if the group heading cells need to be merged.
        If (ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders) Then
          If iGroupStart < iTemp - 1 Then
            Set objMerge = New clsOutputStyle
            objMerge.StartCol = iGroupStart
            objMerge.StartRow = lngXLRow
            objMerge.EndCol = iTemp - 1
            objMerge.EndRow = lngXLRow
              
            colMerges.Add objMerge
            Set objMerge = Nothing
          End If
        End If
        
        lngXLRow = lngXLRow + IIf(ctlTemp.ColumnHeaders, 1, 0) + IIf((ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders), 1, 0)
          
        ' Send data rows and columns to Excel.
        For iLoop = 0 To ctlTemp.Rows - 1
          varBookmark = ctlTemp.AddItemBookmark(iLoop)
          
          ' Check if the row is a Separator/Heading row.
          ' If so we'll need to create a 'style' object for it.
          If fHasHeadingColumn Then
            If ctlTemp.Columns(COLUMN_ISHEADING).CellText(varBookmark) = "1" Then
              Set objTemp = New clsOutputStyle
            
              With objTemp
                .StartCol = lngXLGridStartCol + 1
                .StartRow = lngXLRow
                .EndCol = -1
                .EndRow = lngXLRow
            
                .BackCol = pcolStyles("HeadingCols").BackCol
                .ForeCol = pcolStyles("HeadingCols").ForeCol
                .Bold = pcolStyles("HeadingCols").Bold
                .Underline = pcolStyles("HeadingCols").Underline
                .Gridlines = pcolStyles("HeadingCols").Gridlines
                .Name = "RECPROFROW_" & CStr(iLoop)
              End With
            
              pcolStyles.Add objTemp
              Set objTemp = Nothing
            End If
          End If
          
          lngXLCol = 2

          For iLoop2 = 0 To ctlTemp.Columns.Count - 1
            If ctlTemp.Columns(iLoop2).Visible Then
              ' Send the text or picture to Excel.
              fPhotoDone = False
              If (ctlTemp.TagVariant = COLUMN_ISPHOTO) And _
                (ctlTemp.Columns(iLoop2).Style <> 4) Then
                
                For iLoop3 = 0 To ctlTemp.Columns.Count - 1
                  If ctlTemp.Columns(iLoop3).Visible Then
                    sTemp = PHOTOSTYLESET & CStr(iLoop3 + 1)

                    For iLoop4 = 0 To ctlTemp.StyleSets.Count - 1
                      If ctlTemp.StyleSets(iLoop4).Name = sTemp Then
                        'JPD 20040719 Fault 8725
                        If Not ctlTemp.StyleSets(iLoop4).Picture Is Nothing Then
                          sTempName = GetTmpFName
                          SavePicture ctlTemp.StyleSets(iLoop4).Picture, sTempName
                          mxlWorkBook.ActiveSheet.Range(mxlWorkBook.ActiveSheet.Cells(lngXLRow, lngXLCol), mxlWorkBook.ActiveSheet.Cells(lngXLRow, lngXLCol)).Select
                          mxlWorkBook.ActiveSheet.Pictures.Insert (sTempName)
                          
                          'JPD 20040805 Fault 8984
                          If mxlWorkBook.ActiveSheet.Shapes(mxlWorkBook.ActiveSheet.Shapes.Count).Height > EXCEL_MAXROWHEIGHT Then
                            mxlWorkBook.ActiveSheet.Shapes(mxlWorkBook.ActiveSheet.Shapes.Count).Height = EXCEL_MAXROWHEIGHT
                          End If
                          lngPictureHeight = mxlWorkBook.ActiveSheet.Shapes(mxlWorkBook.ActiveSheet.Shapes.Count).Height
  
                          ' Use the hidden column sizing label in the preview form to
                          ' translate the picture's width into the number of characters that
                          ' fit into that width.
                          pfrmRecProfile.ScaleMode = vbPoints
                          pfrmRecProfile.lblColumnSizingLabel.Caption = "a"
                          Do While pfrmRecProfile.lblColumnSizingLabel.Width < mxlWorkBook.ActiveSheet.Shapes(mxlWorkBook.ActiveSheet.Shapes.Count).Width
                            pfrmRecProfile.lblColumnSizingLabel.Caption = pfrmRecProfile.lblColumnSizingLabel.Caption & "a"
                          Loop
                          lngPictureWidth = Len(pfrmRecProfile.lblColumnSizingLabel.Caption)
                          pfrmRecProfile.ScaleMode = iOriginalScaleMode
                          
                          Kill sTempName
                          fPhotoDone = True
                          Exit For
                        End If
                      End If
                    Next iLoop4
                  End If

                  If fPhotoDone Then
                    Exit For
                  End If
                Next iLoop3
              End If

              If (Not fPhotoDone) And _
                ctlTemp.Columns(iLoop2).TagVariant = COLUMN_ISPHOTO Then

                sTemp = PHOTOSTYLESET & CStr(iLoop2 + 1) & "_" & ctlTemp.Columns(CStr(objRecProfTable.IDPosition)).Value

                For iLoop4 = 0 To ctlTemp.StyleSets.Count - 1
                  If ctlTemp.StyleSets(iLoop4).Name = sTemp Then
                    'JPD 20040719 Fault 8725
                    If Not ctlTemp.StyleSets(iLoop4).Picture Is Nothing Then
                      sTempName = GetTmpFName
                      SavePicture ctlTemp.StyleSets(iLoop4).Picture, sTempName
                      mxlWorkBook.ActiveSheet.Range(mxlWorkBook.ActiveSheet.Cells(lngXLRow, lngXLCol), mxlWorkBook.ActiveSheet.Cells(lngXLRow, lngXLCol)).Select
                      mxlWorkBook.ActiveSheet.Pictures.Insert (sTempName)
                      
                      'JPD 20040805 Fault 8984
                      If mxlWorkBook.ActiveSheet.Shapes(mxlWorkBook.ActiveSheet.Shapes.Count).Height > EXCEL_MAXROWHEIGHT Then
                        mxlWorkBook.ActiveSheet.Shapes(mxlWorkBook.ActiveSheet.Shapes.Count).Height = EXCEL_MAXROWHEIGHT
                      End If
                      
                      lngPictureHeight = mxlWorkBook.ActiveSheet.Shapes(mxlWorkBook.ActiveSheet.Shapes.Count).Height
                      lngPictureWidth = mxlWorkBook.ActiveSheet.Shapes(mxlWorkBook.ActiveSheet.Shapes.Count).Width
                      
                      ' Use the hidden column sizing label in the preview form to
                      ' translate the picture's width into the number of characters that
                      ' fit into that width.
                      pfrmRecProfile.ScaleMode = vbPoints
                      pfrmRecProfile.lblColumnSizingLabel.Caption = "a"
                      Do While pfrmRecProfile.lblColumnSizingLabel.Width < mxlWorkBook.ActiveSheet.Shapes(mxlWorkBook.ActiveSheet.Shapes.Count).Width
                        pfrmRecProfile.lblColumnSizingLabel.Caption = pfrmRecProfile.lblColumnSizingLabel.Caption & "a"
                      Loop
                      lngPictureWidth = Len(pfrmRecProfile.lblColumnSizingLabel.Caption)
                      pfrmRecProfile.ScaleMode = iOriginalScaleMode
                      
                      Kill sTempName
                      fPhotoDone = True
                      Exit For
                    End If
                  End If
                Next iLoop4
              End If

              If fPhotoDone Then
                ' Remember the height/width of pictures, so that the rows/cols
                ' can be sized to fit the pictures later.
                fFound = False
                For iLoop4 = 1 To UBound(alngPictureRows, 2)
                  If alngPictureRows(1, iLoop4) = lngXLRow Then
                    alngPictureRows(2, iLoop4) = IIf(alngPictureRows(2, iLoop4) < lngPictureHeight, lngPictureHeight, alngPictureRows(2, iLoop4))
                    fFound = True
                    Exit For
                  End If
                Next iLoop4
                
                If Not fFound Then
                  ReDim Preserve alngPictureRows(2, UBound(alngPictureRows, 2) + 1)
                  alngPictureRows(1, UBound(alngPictureRows, 2)) = lngXLRow
                  alngPictureRows(2, UBound(alngPictureRows, 2)) = lngPictureHeight
                End If
              
                fFound = False
                For iLoop4 = 1 To UBound(alngPictureCols, 2)
                  If alngPictureCols(1, iLoop4) = lngXLCol Then
                    alngPictureCols(2, iLoop4) = IIf(alngPictureCols(2, iLoop4) < lngPictureWidth, lngPictureWidth, alngPictureCols(2, iLoop4))
                    fFound = True
                    Exit For
                  End If
                Next iLoop4

                If Not fFound Then
                  ReDim Preserve alngPictureCols(2, UBound(alngPictureCols, 2) + 1)
                  alngPictureCols(1, UBound(alngPictureCols, 2)) = lngXLCol
                  alngPictureCols(2, UBound(alngPictureCols, 2)) = lngPictureWidth
                End If
              Else
                varBookmark = ctlTemp.AddItemBookmark(iLoop)
                With mxlWorkBook.ActiveSheet.Cells(lngXLRow, lngXLCol)
                  .NumberFormat = "@"
                  
                  ' Format the cell for the required size & decimals if the column is numeric.
                  If (ctlTemp.Columns(iLoop2).Style <> 4) Then
                    If ctlTemp.ColumnHeaders Then
                      ' Horizontal grid.
                      iDecPlaces = IIf(ctlTemp.Columns(iLoop2).TagVariant = COLUMN_ISPHOTO, -1, CInt(ctlTemp.Columns(iLoop2).TagVariant))
                      bBlankIfZero = CInt(ctlTemp.Columns(iLoop2).AllowSizing)
                      bUse1000Separators = CInt(ctlTemp.Columns(iLoop2).ButtonsAlways)
                    Else
                      ' Vertical grid.
                      iDecPlaces = CInt(ctlTemp.Columns(COLUMN_DECPLACES).CellText(varBookmark))
                      'bBlankIfZero = ctlTemp.Columns(COLUMN_BLANKIFZERO).CellText(varBookmark)
                      'bUse1000Separators = ctlTemp.Columns(COLUMN_USE1000SEPARATOR).CellText(varBookmark)
                      bBlankIfZero = Val(ctlTemp.Columns(COLUMN_BLANKIFZERO).CellText(varBookmark))
                      bUse1000Separators = Val(ctlTemp.Columns(COLUMN_USE1000SEPARATOR).CellText(varBookmark))
                    End If
                    
                    If iDecPlaces >= 0 Then
                      If iDecPlaces > 127 Then iDecPlaces = 127
                      
                      .NumberFormat = _
                        IIf(bUse1000Separators, "#,##", "") & _
                        IIf(bBlankIfZero, "#", "0") & _
                        IIf(iDecPlaces, "." & String(iDecPlaces, "0"), "")
                      '.NumberFormat = "0" & IIf(iDecPlaces > 0, "." & String(iDecPlaces, "0"), "")
                    End If
                    
                    'JPD 20041119 Fault 9381
                    If iDecPlaces = 0 Then
                      If IsDate(ctlTemp.Columns(iLoop2).CellText(varBookmark)) Then
                        .NumberFormat = "@"
                      End If
                    End If
                  End If
                  
                  ' Send the data to the cell.
                  .FormulaR1C1 = ctlTemp.Columns(iLoop2).CellText(varBookmark)
                End With
              End If

              lngXLCol = lngXLCol + 1
            End If
          Next iLoop2
            
          lngXLGridEndRow = lngXLRow
          lngXLRow = lngXLRow + 1
        Next iLoop
      
        If Not fGridFollowed Then
          lngMaxXLRow = IIf(lngMaxXLRow < lngXLGridEndRow, lngXLGridEndRow, lngMaxXLRow)
          lngMaxXLCol = IIf(lngMaxXLCol < lngXLGridEndCol, lngXLGridEndCol, lngMaxXLCol)
          
          ' Apply styles to the table.
          With pcolStyles("Heading")
            .StartCol = lngXLGridStartCol
            .StartRow = lngXLGridStartRow
            .EndCol = lngXLGridEndCol
            .EndRow = IIf(ctlTemp.ColumnHeaders, lngXLGridStartRow + IIf((ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders), 1, 0), -1)
          End With
          
          With pcolStyles("HeadingCols")
            .StartCol = lngXLGridStartCol
            .StartRow = lngXLGridStartRow
            .EndCol = IIf(ctlTemp.ColumnHeaders, -1, lngXLGridStartCol)
            .EndRow = lngXLGridEndRow
          End With

          With pcolStyles("Data")
            .StartCol = lngXLGridStartCol + IIf(ctlTemp.ColumnHeaders, 0, 1)
            .StartRow = lngXLGridStartRow + IIf(ctlTemp.ColumnHeaders, 1, 0) + IIf((ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders), 1, 0)
            .EndCol = lngXLGridEndCol
            .EndRow = lngXLGridEndRow
          End With
          
          With pcolStyles("Title")
            .StartCol = -1
            .StartRow = -1
            .EndCol = -1
            .EndRow = -1
          End With

          ' Set the endRow/endCol values for the separator/heading rows/cols.
          For Each objTemp In pcolStyles
            If Left(objTemp.Name, 11) = "RECPROFCOL_" Then
              objTemp.EndRow = lngXLGridEndRow
            End If
            If Left(objTemp.Name, 11) = "RECPROFROW_" Then
              objTemp.EndCol = lngXLGridEndCol
            End If
          Next objTemp
          Set objTemp = Nothing

          ' Apply styles and merging to the table.
          If mblnApplyStyles Then
            For Each objTemp In pcolStyles
              If (objTemp.EndRow >= 0) And _
                (objTemp.EndCol >= 0) And _
                (objTemp.EndRow >= objTemp.StartRow) And _
                (objTemp.EndCol >= objTemp.StartCol) Then
                
                Set objRange = mxlWorkBook.ActiveSheet.Range(mxlWorkBook.ActiveSheet.Cells(objTemp.StartRow, objTemp.StartCol), mxlWorkBook.ActiveSheet.Cells(objTemp.EndRow, objTemp.EndCol))
                ApplyStyleToRange objRange, objTemp
                objRange.VerticalAlignment = xlCenter
              End If
            Next

            For Each objMerge In colMerges
              If (objMerge.EndRow > objMerge.StartRow) Or _
                (objMerge.EndCol > objMerge.StartCol) Then
                
                Set objRange = mxlWorkBook.ActiveSheet.Range(mxlWorkBook.ActiveSheet.Cells(objMerge.StartRow, objMerge.StartCol), mxlWorkBook.ActiveSheet.Cells(objMerge.EndRow, objMerge.EndCol))
                objRange.MergeCells = True
                objRange.VerticalAlignment = xlCenter
              End If
            Next
          End If
          Set colMerges = Nothing

          ' Get rid of any separator/heading styles that have been created
          ' for this table, so that they are not carried over and applied
          ' to other tables.
          Do While pcolStyles.Count > iOriginalStyles
            pcolStyles.Remove pcolStyles.Count
          Loop
          
          lngXLRow = lngXLRow + 1
        End If
      End If
    
      Set ctlTemp = Nothing
    Else
      If fPageDone Then
        Exit For
      End If
    End If
  Next lngTabIndex
        
  'mxlApp.Visible = True
  With mxlApp.ActiveWindow
    .DisplayGridLines = True 'mblnXLExcelGridlines
    .DisplayHeadings = True 'mblnXLExcelHeaders
  End With
  
  mxlWorkBook.ActiveSheet.Range(mxlWorkBook.ActiveSheet.Cells(2, 2), mxlWorkBook.ActiveSheet.Cells(lngMaxXLRow, lngMaxXLCol)).EntireColumn.AutoFit
  
  ' Size rows/cols with pictures to fit the pictures.
  For iLoop4 = 1 To UBound(alngPictureRows, 2)
    If mxlWorkBook.ActiveSheet.Rows(alngPictureRows(1, iLoop4)).RowHeight < alngPictureRows(2, iLoop4) Then
      mxlWorkBook.ActiveSheet.Rows(alngPictureRows(1, iLoop4)).RowHeight = alngPictureRows(2, iLoop4)
    End If
  Next iLoop4

  For iLoop4 = 1 To UBound(alngPictureCols, 2)
    If mxlWorkBook.ActiveSheet.Columns(alngPictureCols(1, iLoop4)).ColumnWidth < alngPictureCols(2, iLoop4) Then
      mxlWorkBook.ActiveSheet.Columns(alngPictureCols(1, iLoop4)).ColumnWidth = alngPictureCols(2, iLoop4)
    End If
  Next iLoop4
  
  'Put title in after autofit...
  With pcolStyles("Title")
    .StartCol = Val(GetUserSetting("Output", "TitleCol", "3"))
    .StartRow = IIf(mlngAppendStartRow > 0, mlngAppendStartRow, Val(GetUserSetting("Output", "TitleRow", "2")))
    .EndCol = .StartCol
    .EndRow = .StartRow
  End With
  mxlWorkBook.ActiveSheet.Cells(pcolStyles("Title").StartRow, pcolStyles("Title").StartCol).FormulaR1C1 = mstrDefTitle
  Set objRange = mxlWorkBook.ActiveSheet.Cells(pcolStyles("Title").StartRow, pcolStyles("Title").StartCol)
  ApplyStyleToRange objRange, pcolStyles("Title")
  objRange.VerticalAlignment = xlCenter
  
  mxlWorkBook.ActiveSheet.DisplayPageBreaks = False
  
  'Delete superfluous spacerrow and spacer cols if setup in User Config reports section
  If mblnXLExcelOmitLeftCol Then mxlWorkBook.ActiveSheet.Range("A:A").Delete
  If mblnXLExcelOmitTopRow Then mxlWorkBook.ActiveSheet.Range("1:1").Delete
      
  mxlWorkBook.ActiveSheet.Range("A1").Select
  'mxlWorkBook.Sheets(1).Activate

  On Local Error Resume Next  'MH20040120 Fault 7955
  mxlWorkBook.ActiveSheet.PageSetup.Orientation = IIf(mblnXLLandscape, xlLandscape, xlPortrait)

TidyUpAndExit:
  gobjErrorStack.PopStack
  RecordProfilePage = fOK
  Exit Function
  
ErrorTrap:
  gobjErrorStack.HandleError
  fOK = False
  Resume TidyUpAndExit

End Function



Private Sub Class_Initialize()
  mstrPackTemplate = GetUserSetting("Output", "PackTemplate", "")
  mstrXLTemplate = GetUserSetting("Output", "ExcelTemplate", vbNullString)
  mblnXLExcelGridlines = (GetUserSetting("Output", "ExcelGridlines", "0") = "1")
  mblnXLExcelOmitTopRow = (GetUserSetting("Output", "ExcelOmitSpacerRow", "0") = "1")
  mblnXLExcelOmitLeftCol = (GetUserSetting("Output", "ExcelOmitSpacerCol", "0") = "1")
  mblnXLExcelHeaders = (GetUserSetting("Output", "ExcelHeaders", "0") = "1")
  mblnXLAutoFitCols = (GetUserSetting("Output", "AutoFitCols", "1") = "1")
  mblnXLLandscape = (GetUserSetting("Output", "Landscape", "1") = "1")
  
  mblnXLExcelOmitTopRow = False
  mblnXLExcelOmitLeftCol = False
   
  mlngDataStartRow = Val(GetUserSetting("Output", "DataRow", "4"))
  mlngDataStartCol = Val(GetUserSetting("Output", "DataCol", "2"))
 
  mblnSizeColumnsIndependently = False
  mblnApplyStyles = True

End Sub

Public Property Let Screen(blnScreen As Boolean)
  mblnScreen = blnScreen
End Property

Public Property Let DestPrinter(blnPrinter As Boolean)
  mblnPrinter = blnPrinter
End Property

Public Property Let PrinterName(strPrinterName As String)
  mstrPrinterName = strPrinterName
End Property

Public Property Let Save(blnSave As Boolean)
  mblnSave = blnSave
End Property

'Public Property Let SaveFileFormat(lngSaveFileFormat As Long)
'  mlngSaveFileFormat = lngSaveFileFormat
'End Property
'
'Public Property Get SaveFileFormat() As Long
'  SaveFileFormat = mlngSaveFileFormat
'End Property

Public Property Let SaveExisting(lngSaveExisting As Long)
  mlngSaveExisting = lngSaveExisting
End Property

Public Property Get SaveExisting() As Long
  SaveExisting = mlngSaveExisting
End Property

Public Property Let Email(blnEmail As Boolean)
  mblnEmail = blnEmail
End Property

Public Property Let FileName(strFileName As String)
  mstrFileName = strFileName
End Property

Private Function CreateExcelApplication() As Boolean

  On Local Error GoTo LocalErr

  Set mxlApp = CreateObject("Excel.Application")
  mxlApp.DisplayAlerts = False
  
  If GetSystemSetting("Output", "ExcelVisible", "0") = "1" Then
    mxlApp.Visible = True
    mxlApp.DisplayAlerts = True
  Else
    mxlApp.DisplayAlerts = False
  End If
    
  CreateExcelApplication = True

Exit Function

LocalErr:
  mstrErrorMessage = "Error opening Excel Application"
  CreateExcelApplication = False

End Function



Public Function GetFile(objParent As clsOutputRun, colStyles As Collection) As Boolean

  On Local Error GoTo LocalErr

  If Not CreateExcelApplication Then
    GetFile = False
    Exit Function
  End If
  
  ' Leave the app there after user has closed the worksheet
  mxlApp.UserControl = True

  If gblnReportPackMode And mlngSaveExisting = 0 And gobjProgress.Bar1Value > 0 Then
      'reset write mode to append in order that the previous report job doesnt get overwritten
      mlngSaveExisting = 4
  End If
  'mxlApp.Visible = True
  
  'Check if file already exists...
  If Dir(mstrFileName) <> vbNullString And mstrFileName <> vbNullString Then

    Select Case mlngSaveExisting
    Case 0  'Overwrite
      If Not objParent.KillFile(mstrFileName) Then
        GetFile = False
        Exit Function
      End If
      
      GetWorkBook strWorkbook:="New", strWorksheet:="New"

    Case 1  'Do not overwrite (fail)
      mxlApp.Quit
      Set mxlApp = Nothing
      mstrErrorMessage = "File already exists."
    
    Case 2  'Add Sequential number to file
      mstrFileName = mobjParent.GetSequentialNumberedFile(mstrFileName)
      GetWorkBook strWorkbook:="New", strWorksheet:="New"
    
    Case 3  'Append to existing file
      GetWorkBook strWorkbook:="Open", strWorksheet:="Existing"

    Case 4  'Create new worksheet within existing workbook...
      GetWorkBook strWorkbook:="Open", strWorksheet:="New"

    End Select

  Else
    GetWorkBook strWorkbook:="New", strWorksheet:="New"

  End If

  GetFile = (mstrErrorMessage = vbNullString)

Exit Function

LocalErr:
  mstrErrorMessage = Err.Description
  GetFile = False

End Function


Private Sub GetWorkBook(strWorkbook As String, strWorksheet As String)

  Dim strFormat As String
  Dim strTempFile As String
  Dim lngCount As Long
  Dim lngOriginalFormat As Long
  
  If gblnReportPackMode Then mstrXLTemplate = mstrPackTemplate
    
  If mblnApplyStyles And mstrXLTemplate <> "" And Dir(mstrXLTemplate) <> "" Then
    
    If Not IsFileCompatibleWithExcelVersion(mstrXLTemplate) Then
      mstrErrorMessage = "Your User Configuration Output Options are set to use a template file which is not compatible with your version of Microsoft Office."
      Exit Sub
    End If

    Set mxlTemplateBook = mxlApp.Workbooks.Open(mstrXLTemplate, ReadOnly:=True)
    
    'Save a temp template in the format of the output...
    If mstrFileName <> vbNullString Then
      strFormat = GetOfficeSaveAsFormat(mstrFileName, GetOfficeExcelVersion, oaExcel)
      If strFormat <> "" Then
        strTempFile = mobjParent.GetTempFileName("")
        mxlTemplateBook.SaveAs strTempFile, Val(strFormat)
        mxlTemplateBook.Close
        Set mxlTemplateBook = mxlApp.Workbooks.Open(strTempFile, ReadOnly:=True)
      End If
    End If
    
    Set mxlTemplateSheet = mxlTemplateBook.ActiveSheet
  
  End If

  mstrSheetMode = strWorksheet
  Select Case strWorkbook
  Case "New"
    'Create new workbook and ensure only 1 worksheet
    
    If GetOfficeExcelVersion >= 12 And mstrXLTemplate <> vbNullString Then
      'Make sure the new workbook is in the same format as the template
      'otherwise we won't be able to copy sheets into the new workbook.
      lngOriginalFormat = mxlApp.DefaultSaveFormat
      mxlApp.DefaultSaveFormat = GetOfficeSaveAsFormat(mstrXLTemplate, GetOfficeExcelVersion, oaExcel)
      Set mxlWorkBook = mxlApp.Workbooks.Add
      mxlApp.DefaultSaveFormat = lngOriginalFormat
    Else
      Set mxlWorkBook = mxlApp.Workbooks.Add
    End If
    
    For lngCount = 1 To mxlWorkBook.Sheets.Count - 1
      mxlWorkBook.Sheets(1).Delete
    Next
    Set mxlDeleteSheet = mxlWorkBook.Sheets(1)

  Case "Open"
    If Not IsFileCompatibleWithExcelVersion(mstrFileName) Then
      mstrErrorMessage = "This definition is set to append to a file which is not compatible with your version of Microsoft Office."
      Exit Sub
    End If
    
    Set mxlWorkBook = mxlApp.Workbooks.Open(mstrFileName)

  End Select

Exit Sub

LocalErr:
  mstrErrorMessage = "Error getting new workbook (" & Err.Description & ")"

End Sub


Private Sub GetWorksheet(strSheetName As String)
  
  Dim blnFound As Boolean
  
  On Local Error GoTo LocalErr
  
  mblnAppending = False
  mlngAppendStartRow = 0
   
  'If we are appending, then see if there is an existing worksheet with this name...
  blnFound = False
  If mstrSheetMode = "Existing" Then
    For Each mxlWorkSheet In mxlWorkBook.Worksheets
      If Trim(mxlWorkSheet.Name) = FormatSheetName(strSheetName) Then
        mxlWorkSheet.Activate
        blnFound = True
        Exit For
      End If
    Next
  End If
    
  If blnFound Then
    StartAtBottomOfSheet
    mblnAppending = True
  Else
    If Not mxlTemplateSheet Is Nothing Then
      mxlTemplateSheet.Copy After:=mxlWorkBook.Sheets(mxlWorkBook.Sheets.Count)
      Set mxlWorkSheet = mxlWorkBook.ActiveSheet
      StartAtBottomOfSheet
    Else
      Set mxlWorkSheet = mxlWorkBook.Sheets.Add(After:=mxlWorkBook.Sheets(mxlWorkBook.Sheets.Count))
    End If

    SetSheetName mxlWorkSheet, strSheetName

  End If

  If Not (mxlDeleteSheet Is Nothing) Then
    mxlDeleteSheet.Delete
    Set mxlDeleteSheet = Nothing
  End If

Exit Sub

LocalErr:
  mstrErrorMessage = "Error getting new worksheet (" & Err.Description & ")"

End Sub

Public Function AddPage(strDefTitle As String, strSheetName As String, colStyles As Collection, Optional mblnOutputCoverSheet As Boolean)

  On Local Error GoTo LocalErr

  mstrDefTitle = strDefTitle

  If mblnPivotTable Then
    GetWorksheet "Data " & strSheetName
  Else
    GetWorksheet IIf(gblnReportPackMode, strDefTitle, strSheetName)
  End If
  
  If Not mblnChart And Not mblnPivotTable Then
    If mxlFirstSheet Is Nothing Then
      Set mxlFirstSheet = mxlWorkSheet
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


Public Sub DataArray(strArray() As String, colColumns As Collection, colStyles As Collection, colMerges As Collection)

  Dim objColumn As clsColumn
  Dim lngGridCol As Long
  Dim lngGridRow As Long
  Dim lngXLCol As Long
  Dim lngXLRow As Long
  Dim strCell As String
  Dim sDateFormat As String

  On Local Error GoTo LocalErr

  If mstrErrorMessage <> vbNullString Then
    Exit Sub
  End If
  
  'If UBound(strArray, 1) > 255 Then
  If colColumns.Count + mlngDataStartCol > 255 Then
    mstrErrorMessage = "Maximum of 255 columns exceeded"
    Exit Sub
  End If
  
  sDateFormat = DateFormat
  
  PrepareRows UBound(strArray, 2), colColumns, colStyles

  lngXLCol = mlngDataStartCol
  lngXLRow = mlngDataCurrentRow
  For lngGridRow = 0 To UBound(strArray, 2)
    For lngGridCol = 0 To UBound(strArray, 1)

      With mxlWorkSheet.Cells(lngXLRow + lngGridRow, lngXLCol + lngGridCol)

        If .NumberFormat = sDateFormat & ";@" Then
          If strArray(lngGridCol, lngGridRow) Like "??/??/????" Or strArray(lngGridCol, lngGridRow) = vbNullString Then
            strArray(lngGridCol, lngGridRow) = Replace(Format(strArray(lngGridCol, lngGridRow), "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/")
          Else
            'A non-date in a date column (Report sub totals for example)...
            .NumberFormat = "General"
          End If
        End If

        .FormulaR1C1 = _
          IIf(Left$(strArray(lngGridCol, lngGridRow), 1) = "'", "'", vbNullString) & _
          strArray(lngGridCol, lngGridRow)
      End With
    Next

    If lngGridRow Mod 100 = 0 Then
      If gobjProgress.Visible And gobjProgress.Cancelled Then
        mstrErrorMessage = "Cancelled by user."
        Exit Sub
      End If
    End If

    'MH20071219 Fault 12727
    If mobjParent.UpdateProgressPerRow Then
      gobjProgress.UpdateProgress gblnBatchMode
    End If
  Next


  If mblnChart Then
    ApplyStyle UBound(strArray, 1), UBound(strArray, 2), colStyles
    ApplyCellOptions UBound(strArray, 1), colStyles, True
    
    CreateChart mlngDataCurrentRow + UBound(strArray, 2), mlngDataStartCol + UBound(strArray, 1), colStyles
    ApplyCellOptions UBound(strArray, 1), colStyles, False
    
    'Delete superfluous rows and cols if setup in User Config reports section
    If mblnXLExcelOmitLeftCol Then mxlWorkSheet.Range("A:A").Delete
    If mblnXLExcelOmitTopRow Then mxlWorkSheet.Range("1:1").Delete
  
  ElseIf mblnPivotTable Then

    If UBound(strArray, 1) < 1 Then
      mstrErrorMessage = "Unable to create a pivot table for a single column of data."
    Else
      ApplyStyle UBound(strArray, 1), UBound(strArray, 2), colStyles
      
      CreatePivotTable mlngDataCurrentRow + UBound(strArray, 2), mlngDataStartCol + UBound(strArray, 1), strArray(0, 0), strArray(1, 0), strArray(UBound(strArray), 0), colStyles, colColumns
    End If

  Else
    If mblnApplyStyles Then
      ApplyStyle UBound(strArray, 1), UBound(strArray, 2), colStyles
      ApplyMerges colMerges
    End If
    ApplyCellOptions UBound(strArray, 1), colStyles, True
      
    'Delete superfluous rows and cols if setup in User Config reports section
    If mblnXLExcelOmitLeftCol Then mxlWorkSheet.Range("A:A").Delete
    If mblnXLExcelOmitTopRow Then mxlWorkSheet.Range("1:1").Delete
      
  End If

  mlngDataCurrentRow = mlngDataCurrentRow + UBound(strArray, 2) + IIf(mblnApplyStyles, 2, 1)

Exit Sub

LocalErr:
  mstrErrorMessage = Err.Description
  
  Select Case Err.Number
  Case 7  'Out of Memory
    If Val(mxlApp.Version) < 12 And Len(strArray(lngGridCol, lngGridRow)) > 65535 Then
      mstrErrorMessage = "Multi-line data is too large for this version of Excel."
    End If
  End Select

End Sub


Private Sub CreateChart(lngMaxRows As Long, lngMaxCols As Long, colStyles As Collection)
  
  Dim xlChart As Excel.Chart
  Dim xlData As Excel.Range
  Dim strSheetName As String

  On Local Error GoTo LocalErr

  Set xlData = mxlWorkSheet.Range(mxlWorkSheet.Cells(mlngDataCurrentRow, mlngDataStartCol), mxlWorkSheet.Cells(lngMaxRows, lngMaxCols))
  strSheetName = mxlWorkSheet.Name & " Chart"
  
  Set xlChart = mxlApp.Charts.Add(After:=mxlWorkSheet)
  With xlChart
    .ChartType = xl3DColumnClustered
    .SetSourceData Source:=xlData, PlotBy:=xlColumns
    .Location Where:=xlLocationAsNewSheet
    .HasTitle = True
    .ChartTitle.Text = mstrDefTitle
    .ChartTitle.Font.Bold = colStyles("Title").Bold
    .ChartTitle.Font.Size = 12
    .ChartTitle.Font.Underline = colStyles("Title").Underline
    .ChartTitle.Font.Color = colStyles("Title").ForeCol
  End With

  SetSheetName mxlWorkBook.ActiveChart, strSheetName
  If mxlFirstSheet Is Nothing Then
    Set mxlFirstSheet = mxlWorkBook.ActiveChart
  End If
Exit Sub

LocalErr:
  mstrErrorMessage = Err.Description

End Sub


Private Sub CreatePivotTable(lngMaxRows As Long, lngMaxCols As Long, strHor As String, strVer As String, strInt As String, colStyles As Collection, colColumns As Collection)

  Dim xlPivot As Excel.PivotTable
  Dim xlDataSheet As Excel.Worksheet
  Dim xlData As Excel.Range
  Dim xlStart As Excel.Range
  Dim objColumn As clsColumn
  Dim strSheetName As String
  Dim xlFunc As XlConsolidationFunction

  If ASRDEVELOPMENT Then
  On Local Error GoTo 0
  Else
  On Local Error GoTo LocalErr
  End If

  mxlApp.DisplayAlerts = True
  
  Set xlData = mxlWorkSheet.Range(mxlWorkSheet.Cells(mlngDataCurrentRow, mlngDataStartCol), mxlWorkSheet.Cells(lngMaxRows, lngMaxCols))
  strSheetName = Mid(mxlWorkSheet.Name, 6)
  Set xlDataSheet = mxlWorkSheet

  GetWorksheet strSheetName
  If mxlFirstSheet Is Nothing Then
    Set mxlFirstSheet = mxlWorkSheet
  End If

  xlDataSheet.Visible = xlSheetHidden 'IIf(gblnReportPackMode, xlSheetVisible, xlSheetHidden)
  Set xlStart = mxlWorkSheet.Cells(mlngDataCurrentRow, mlngDataStartCol)

  mxlApp.DisplayAlerts = False
  
  'MH20100628
  If Val(mxlApp.Version) > 12 Then
    Set xlPivot = mxlWorkBook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=xlData). _
                  CreatePivotTable(TableDestination:=xlStart)
  Else
    Set xlPivot = mxlWorkSheet.PivotTableWizard(SourceType:=xlDatabase, SourceData:=xlData, _
                  TableDestination:=xlStart)
  End If

  With xlPivot
    .AddFields RowFields:=strVer, ColumnFields:=strHor
    
    'AE20071017 Fault #12540
    'Changed ".PivotFields().Orientation = xlDataField" to ".AddDataField" so we can specify
    'which function the user selected for the output
    Select Case mobjParent.PivotDataFunction
    Case "Count"
      xlFunc = xlCount
    Case "Average"
      xlFunc = xlAverage
    Case "Maximum"
      xlFunc = xlMax
    Case "Minimum"
      xlFunc = xlMin
    Case "Total"
      xlFunc = xlSum
    End Select
    
    With .PivotFields(strInt)
        .Orientation = xlDataField
        .Name = mobjParent.PivotDataFunction & " of " & strInt
        .Function = xlFunc
    End With
    
    .NullString = IIf(mobjParent.PivotSuppressBlanks, "", "0")
      
    Set objColumn = colColumns(colColumns.Count)
    If objColumn.DecPlaces > 0 Then
      If objColumn.DecPlaces > 100 Then objColumn.DecPlaces = 100
      .DataBodyRange.NumberFormat = _
        IIf(objColumn.ThousandSeparator, "#,##0", "0") & _
        IIf(objColumn.DecPlaces, "." & String(objColumn.DecPlaces, "0"), "")
    End If
    Set objColumn = Nothing
      
    mxlApp.DisplayAlerts = True
  End With

  ApplyStyleToRange xlStart, colStyles("Heading")
  ApplyStyleToRange xlPivot.RowRange, colStyles("Heading")
  ApplyStyleToRange xlPivot.ColumnRange, colStyles("Heading")
  ApplyStyleToRange xlPivot.DataBodyRange, colStyles("Data")
  mxlWorkSheet.Range("A1").Select
  mlngHeaderCols = 1

  ApplyCellOptions xlPivot.ColumnRange.Columns.Count, colStyles, True
  
  'Delete superfluous rows and cols if setup in User Config reports section
  If mblnXLExcelOmitLeftCol Then mxlWorkSheet.Range("A:A").Delete
  If mblnXLExcelOmitTopRow Then mxlWorkSheet.Range("1:1").Delete
      
  Set xlPivot = Nothing
  mxlApp.DisplayAlerts = False

Exit Sub

LocalErr:
  mstrErrorMessage = Err.Description

End Sub


Private Sub PrepareRows(lngRowCount As Long, colColumns As Collection, colStyles As Collection)

  Dim objColumn As clsColumn
  Dim lngCount As Long

  On Local Error GoTo LocalErr

  With mxlWorkSheet
    If .Visible Then
      .Range("A1").Select
    End If
    
    If mlngHeaderRows > 0 Then
      With .Range(.Cells(mlngDataCurrentRow, mlngDataStartCol), .Cells(mlngDataCurrentRow + mlngHeaderRows - 1, mlngDataStartCol + colColumns.Count))
        .NumberFormat = "@"
      End With
    End If

    For lngCount = 0 To colColumns.Count - 1
      Set objColumn = colColumns(lngCount + 1)

      With .Range(.Cells(mlngDataCurrentRow + mlngHeaderRows, mlngDataStartCol + lngCount), .Cells(mlngDataCurrentRow + lngRowCount, mlngDataStartCol + lngCount))
        Select Case objColumn.DataType
        Case sqlNumeric, sqlInteger

          If objColumn.DecPlaces > 0 Then
            If objColumn.DecPlaces > 100 Then objColumn.DecPlaces = 100
            .NumberFormat = _
              IIf(objColumn.ThousandSeparator, "#,##0", "0") & _
              IIf(objColumn.DecPlaces, "." & String(objColumn.DecPlaces, "0"), "")
          End If

        Case sqlBoolean
          .HorizontalAlignment = xlCenter
          .NumberFormat = "@"
        Case sqlUnknown
          'Leave it alone! (Required for percentages on Standard Reports)
        Case sqlDate
          'Adding ;@ to the end formats it as "short date" so excel will look at the
          'regional settings when opening the workbook rather than force it to always
          'be in the format of the user who created the workbook.
          .NumberFormat = DateFormat & ";@"
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


Private Sub ApplyCellOptions(lngColCount As Long, colStyles As Collection, blnGridlines As Boolean)
  
  Dim objRange As Excel.Range
  Dim lngCount As Long
  
  Dim lngMaxWidth As Double
  Dim lngTitleColWidth As Double
  Dim lngTitleSize As Double

  On Local Error GoTo LocalErr

  If mblnXLAutoFitCols Then
    mxlWorkSheet.Range(mxlWorkSheet.Cells(mlngDataCurrentRow, mlngDataStartCol), mxlWorkSheet.Cells(mlngDataCurrentRow, mlngDataStartCol + lngColCount)).EntireColumn.AutoFit
  
    If Not mblnSizeColumnsIndependently Then
      lngMaxWidth = 0
      For lngCount = mlngDataStartCol + mlngHeaderCols To mlngDataStartCol + lngColCount
        If lngMaxWidth < mxlWorkSheet.Columns(lngCount).ColumnWidth Then
          lngMaxWidth = mxlWorkSheet.Columns(lngCount).ColumnWidth
        End If
      Next
  
      For lngCount = mlngDataStartCol + mlngHeaderCols To mlngDataStartCol + lngColCount
        mxlWorkSheet.Columns(lngCount).ColumnWidth = lngMaxWidth
      Next
    End If

  End If
  
  
  If mblnApplyStyles Then
    If blnGridlines Then
      With mxlApp.ActiveWindow
        .DisplayGridLines = mblnXLExcelGridlines
        .DisplayHeadings = mblnXLExcelHeaders
      End With
    End If

    With colStyles("Title")
      .StartCol = Val(GetUserSetting("Output", "TitleCol", "3"))
      .StartRow = IIf(mlngAppendStartRow > 0, mlngAppendStartRow, Val(GetUserSetting("Output", "TitleRow", "2")))
      .EndCol = .StartCol
      .EndRow = .StartRow
    End With

    'Put title in after autofit...
    If colStyles("Title").StartCol <> 0 And colStyles("Title").StartRow <> 0 Then
      mxlWorkSheet.Cells(colStyles("Title").StartRow, colStyles("Title").StartCol).FormulaR1C1 = mstrDefTitle
      Set objRange = mxlWorkSheet.Cells(colStyles("Title").StartRow, colStyles("Title").StartCol)
      ApplyStyleToRange objRange, colStyles("Title")
      
      'Merge cells for the title column so that if you append to the file
      'then the title is not taken into account during column sizing.
      With mxlWorkSheet.Columns(colStyles("Title").StartCol)
        lngTitleColWidth = .ColumnWidth
        lngMaxWidth = .Width
        .AutoFit
        lngTitleSize = .Width
  
        lngCount = colStyles("Title").StartCol
        Do
          lngCount = lngCount + 1
          lngMaxWidth = lngMaxWidth + mxlWorkSheet.Columns(lngCount).Width
        Loop While lngMaxWidth < lngTitleSize
  
        With mxlWorkSheet
          .Range(.Cells(objRange.Row, objRange.Column), .Cells(objRange.Row, lngCount)).Merge
        End With
        .ColumnWidth = lngTitleColWidth
      End With
'      'Delete superfluous rows and cols if setup in User Config reports section
'      If mblnXLExcelOmitLeftCol Then mxlWorkSheet.Range("A:A").Delete
'      If mblnXLExcelOmitTopRow Then mxlWorkSheet.Range("1:1").Delete
    End If
  End If

Exit Sub

LocalErr:
  mstrErrorMessage = Err.Description

End Sub



Private Sub ApplyStyle(lngNumCols As Long, lngNumRows As Long, colStyles As Collection)

  Dim objStyle As clsOutputStyle
  Dim objRange As Excel.Range
  Dim lngCol As Long
  Dim lngRow As Long

  On Local Error GoTo LocalErr

  lngCol = mlngDataStartCol
  lngRow = mlngDataCurrentRow

  With colStyles("Title")
    .StartCol = Val(GetUserSetting("Output", "TitleCol", "3"))
    .StartRow = IIf(mlngAppendStartRow > 0, mlngAppendStartRow, Val(GetUserSetting("Output", "TitleRow", "2")))
    .EndCol = .StartCol
    .EndRow = .StartRow
  End With
  
  With colStyles("Heading")
    .StartCol = 0
    .StartRow = 0
    .EndCol = lngNumCols
    .EndRow = mlngHeaderRows - 1
  End With
  
  If mlngHeaderCols > 0 Then
    With colStyles("HeadingCols")
      .StartCol = 0
      .StartRow = 0
      .EndCol = mlngHeaderCols - 1
      .EndRow = lngNumRows
    End With
  End If

  With colStyles("Data")
    .StartCol = mlngHeaderCols
    .StartRow = mlngHeaderRows
    .EndCol = lngNumCols
    .EndRow = lngNumRows
  End With

  For Each objStyle In colStyles
    If objStyle.Name <> "Title" Then
      If objStyle.StartRow + lngRow > 0 And objStyle.StartCol + lngCol > 0 Then
        Set objRange = mxlWorkSheet.Range(mxlWorkSheet.Cells(objStyle.StartRow + lngRow, objStyle.StartCol + lngCol), mxlWorkSheet.Cells(objStyle.EndRow + lngRow, objStyle.EndCol + lngCol))
        ApplyStyleToRange objRange, objStyle
      End If
    End If
  Next

Exit Sub

LocalErr:
  mstrErrorMessage = Err.Description

End Sub


Private Sub ApplyMerges(colMerges As Collection)

  Dim objMerge As clsOutputStyle
  Dim objRange As Excel.Range
  Dim lngCol As Long
  Dim lngRow As Long

  On Local Error GoTo LocalErr

  lngCol = mlngDataStartCol
  lngRow = mlngDataCurrentRow

  For Each objMerge In colMerges
    If objMerge.StartRow + lngRow > 0 And objMerge.StartCol + lngCol > 0 Then
      Set objRange = mxlWorkSheet.Range(mxlWorkSheet.Cells(objMerge.StartRow + lngRow, objMerge.StartCol + lngCol), mxlWorkSheet.Cells(objMerge.EndRow + lngRow, objMerge.EndCol + lngCol))
      objRange.MergeCells = True
      objRange.VerticalAlignment = xlTop
    End If
  Next

Exit Sub

LocalErr:
  mstrErrorMessage = Err.Description

End Sub


Private Sub ApplyStyleToRange(objRange As Excel.Range, objStyle As clsOutputStyle)

  On Local Error GoTo LocalErr

  With objRange

    If objStyle.CenterText Then
      .HorizontalAlignment = xlCenter
    End If

    .Font.Bold = objStyle.Bold
    .Font.Underline = objStyle.Underline
    .Font.Color = objStyle.ForeCol
    
    'Don't do the backcol nor gridlines for the title...
    If objStyle.Name <> "Title" Then
      .Interior.Color = objStyle.BackCol
  
      On Error Resume Next
  
      If objStyle.Gridlines Then
        .BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
        If objStyle.StartCol <> objStyle.EndCol Then
          .Borders(xlInsideVertical).LineStyle = xlContinuous
        End If
        If objStyle.StartRow <> objStyle.EndRow Then
          .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        End If
      Else
        .BorderAround xlNone
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
      End If

    End If

  End With

Exit Sub

LocalErr:
  mstrErrorMessage = Err.Description

End Sub


Public Sub Complete()

  Dim objChart As Excel.Chart
  Dim objWorksheet As Excel.Worksheet
  Dim strFormat As String
  Dim blnOffice2007 As Boolean
  Dim strTempFile As String
  Dim strExtension As String
  Dim aryFileBits() As String

  On Local Error GoTo LocalErr
  
  If mstrErrorMessage <> vbNullString Then
    Exit Sub
  End If
  
  If gblnReportPackMode Then
    mxlWorkBook.Sheets(1).Activate
  Else
    mxlFirstSheet.Activate
  End If

  'SAVE
  If mblnSave Then
    mstrErrorMessage = "Error saving file <" & mstrFileName & ">"
    
    strFormat = GetOfficeSaveAsFormat(mstrFileName, GetOfficeExcelVersion, oaExcel)
    If strFormat = "" Then
      mstrErrorMessage = "This definition is set to save in a file format which is not compatible with your version of Microsoft Office."
      GoTo TidyAndExit
    End If
    
    ' calculate the appropriate output type
    aryFileBits = Split(mstrFileName, ".")
    strExtension = aryFileBits(UBound(aryFileBits))
   
    Select Case UCase(strExtension)
      Case "XLSX"
        mxlWorkBook.SaveAs mstrFileName, FileFormat:=Val(strFormat)
      Case "XLS"
        mxlWorkBook.SaveAs mstrFileName, FileFormat:=Val(strFormat)
      Case "HTML"
        mxlWorkBook.SaveAs mstrFileName, FileFormat:=Val(xlHtml)
    End Select

  End If


  'EMAIL
  If mblnEmail And Not gblnReportPackMode Then
    mstrErrorMessage = "Error sending email"

    strFormat = GetOfficeSaveAsFormat(mobjParent.EmailAttachAs, GetOfficeExcelVersion, oaExcel)
    If strFormat = "" Then
      mstrErrorMessage = "This definition is set to email an attachment in a file format which is not compatible with your version of Microsoft Office."
      GoTo TidyAndExit
    End If

    strTempFile = mobjParent.GetTempFileName(mobjParent.EmailAttachAs)
    mxlWorkBook.SaveAs strTempFile, Val(strFormat)
    mxlWorkBook.Close False
    mobjParent.SendEmail strTempFile
    Set mxlWorkBook = mxlApp.Workbooks.Open(strTempFile)
  End If
  

  'PRINTER
  If mblnPrinter Then
    Dim strCurrentPrinter As String
    mstrErrorMessage = "Error printing"
    
    If mblnChart Then
      For Each objChart In mxlWorkBook.Charts
        objChart.PrintOut , , , , mstrPrinterName
      Next
    Else
      For Each objWorksheet In mxlWorkBook.Worksheets
        If objWorksheet.Visible = xlSheetVisible Then
          objWorksheet.PrintOut , , , , mstrPrinterName
        End If
      Next
    End If
    
  End If


  'SCREEN
  If mblnScreen Then
    mstrErrorMessage = "Error displaying Excel"
    mxlApp.DisplayAlerts = True
    mxlApp.WindowState = xlMaximized
    mxlApp.Visible = True
    mxlWorkBook.Activate
    Set mxlWorkBook = Nothing
    Set mxlApp = Nothing        'Stops Excel quitting...
  Else
    mxlWorkBook.Saved = True
    mxlWorkBook.Close
    mxlApp.Quit
  End If

  mstrErrorMessage = vbNullString

TidyAndExit:
  ClearUp

Exit Sub

LocalErr:
  mstrErrorMessage = mstrErrorMessage & _
    IIf(Err.Description <> vbNullString, vbCrLf & " (" & Err.Description & ")", vbNullString)
  Resume TidyAndExit

End Sub

Public Property Let Chart(ByVal blnNewValue As Boolean)
  mblnChart = blnNewValue
End Property

Public Property Let PivotTable(ByVal blnNewValue As Boolean)
  mblnPivotTable = blnNewValue
End Property

Public Property Let HeaderRows(ByVal lngNewValue As Long)
  mlngHeaderRows = lngNewValue
End Property

Public Property Let HeaderCols(ByVal lngNewValue As Long)
  mlngHeaderCols = lngNewValue
End Property

Public Property Let SizeColumnsIndependently(ByVal blnNewValue As Boolean)
  mblnSizeColumnsIndependently = blnNewValue
End Property

Public Property Let ApplyStyles(ByVal blnNewValue As Boolean)
  mblnApplyStyles = blnNewValue
End Property

Public Property Let Parent(objNewValue As clsOutputRun)
  Set mobjParent = objNewValue
End Property

Private Sub Class_Terminate()
  Set mxlFirstSheet = Nothing
  Set mxlTemplateSheet = Nothing
  Set mxlTemplateBook = Nothing
  Set mxlWorkSheet = Nothing
  Set mxlWorkBook = Nothing
  Set mxlApp = Nothing
End Sub

Public Property Get ErrorMessage() As String
  ErrorMessage = mstrErrorMessage
End Property

Private Sub StartAtBottomOfSheet()

  'Start at the bottom of the sheet
  mlngAppendStartRow = mxlWorkSheet.Cells.SpecialCells(xlCellTypeLastCell).Row + IIf(mblnApplyStyles, 2, 1)
  mlngDataCurrentRow = mlngAppendStartRow + IIf(mblnApplyStyles, 2, 1)

End Sub


Private Function FormatSheetName(strSheetName As String) As String

  Dim strInvalidChars As String
  Dim lngCount As Long

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



Private Function SetSheetName(objObject As Object, ByVal strSheetName As String) As Boolean

  Dim strNumber As String
  Dim lngCount As Long

  strSheetName = FormatSheetName(strSheetName)

  On Local Error Resume Next
  Err.Clear
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
      Err.Clear
      strNumber = "(" & CStr(lngCount) & ")"
      objObject.Name = Left(strSheetName, 31 - Len(strNumber)) & strNumber
      If lngCount > 256 Then
        mstrErrorMessage = "Error naming sheet"
        Exit Function
      End If
    Loop While Err.Number > 0
  End If

  On Local Error Resume Next
  If mxlTemplateSheet Is Nothing Then
    With objObject.PageSetup
      .LeftFooter = "Created on &D at &T by " & gsUserName
      .RightFooter = "Page &P"
      .Orientation = IIf(mblnXLLandscape, xlLandscape, xlPortrait)
      mxlWorkSheet.DisplayPageBreaks = False
    End With
  End If

  SetSheetName = True

End Function



