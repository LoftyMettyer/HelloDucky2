VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrintGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mblnSuppressPrompt As Boolean
Private mgrdTemp As SSDataWidgets_B.SSDBGrid
Private mstrHeading As String
Private mintNumberOfCopies As Integer
Private mblnCollate As Boolean
Private mblnPortrait As Boolean
Private mblnGridlines As Boolean
Private mblnShading As Boolean
Private mblnHeadingsOnEveryPage As Boolean
Private mintMarginTop As Integer
Private mintMarginLeft As Integer
Private mintMarginRight As Integer
Private mintMarginBottom As Integer
Private mblnPrintSelectedRowsOnly As Boolean
Private mblnCancelled As Boolean
Private mgrdSelected As SSDataWidgets_B.SSDBGrid

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As Rect, ByVal wFormat As Long) As Long
Private Const DT_LEFT = &H0
Private Const DT_CENTER = &H1
Private Const DT_NOPREFIX = &H800
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_VCENTER = &H4
Private Const DT_CALCRECT = &H400
Private Const DT_WORDBREAK = &H10
Private Const c_DTDefFmt = DT_NOPREFIX 'Or DT_SINGLELINE Or DT_VCENTER

Private Type Rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Enum EAlignment
  eRight = DT_RIGHT
  eLeft = DT_LEFT
  eCenter = DT_CENTER
End Enum

Private miRecordProfilePageCounter As Integer
Private mlngRecordProfileTopOffset As Long
Private mlngRecordProfileBottomOffset As Long
Private mfntRecordProfileHeaderFont As StdFont
Private mfntRecordProfileFooterFont As StdFont

Private Const TWIPSPERMM = 56.7
Private Const OFFSET_X = 20
Private Const OFFSET_Y = 5
Private Const MARGINCORRECTION = 5 * TWIPSPERMM
Private Const OVERLAPCOMPENSATION = 500

Private Const PRINTERFACTOR = 5
Private Const THINLINE = 1 * PRINTERFACTOR
Private Const THICKLINE = 2 * PRINTERFACTOR
  
Public Enum GridCellType
  cellGROUPHEADING = 0
  cellCOLUMNHEADING = 1
  cellDATA = 2
  cellSEPARATOR = 3
End Enum

Private Const COLUMN_ISHEADING = "IsHeading"
Private Const COLUMN_ISPHOTO = "IsPhoto"
Private Const PHOTOSTYLESET = "PhotoSS_"
Private Const RECPROFFOLLOWONCORRECTION = 10
Private mstrSelectedPrinter As String

Private mfrmRecProfileForm As Form
Private msErrorMessage As String



Public Property Get Grid_Selected() As SSDataWidgets_B.SSDBGrid
  Set Grid_Selected = mgrdSelected
End Property

Public Property Let Grid_Selected(ByVal blnNewValue As SSDataWidgets_B.SSDBGrid)
  Set mgrdSelected = blnNewValue
End Property

Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property

Public Property Let Cancelled(ByVal blnNewValue As Boolean)
  mblnCancelled = blnNewValue
End Property

Private Function PrintRecordProfileColumnHeadings( _
  pctlGrid As SSDBGrid, _
  plngMarginLeft_Twips As Long, _
  plngPageTopPictureBoxPosition As Long, _
  plngGroupColumnHeadingHeight As Long, _
  palngOverflows As Variant, _
  plngStartColumn As Long, _
  plngStartColumnPages As Long) As Boolean
  
  On Error GoTo ErrorTrap
  
  ' Print the group and column headings of the given grid.
  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim rctDraw As Rect
  Dim iSM As Integer
  Dim lngRightMostLine As Long
  Dim iLastGroup As Integer
  Dim iLastGroupToPrint As Integer
  Dim lngLeftMostLine As Long
  Dim fOntoNextPage
  Dim fSkipGroup As Boolean
  Dim fHasGroups As Boolean
  
  fOK = True
  iSM = Printer.ScaleMode
  
  fHasGroups = (pctlGrid.Groups.Count > 0) And (pctlGrid.GroupHeaders)
  
  iLastGroup = -1
  iLastGroupToPrint = -1
  lngLeftMostLine = -1
  lngRightMostLine = 0
    
  For iLoop = plngStartColumn To pctlGrid.Columns.Count - 1
    fOntoNextPage = False
    fSkipGroup = False
    
    If pctlGrid.Columns(iLoop).Visible Then
      If fHasGroups Then
        If (iLastGroup >= 0) And _
          (iLastGroup <> pctlGrid.Columns(iLoop).Group) Then
          
          ' Calculate the print coordinates of the group heading box
          rctDraw.Left = lngLeftMostLine
          rctDraw.Right = lngRightMostLine
          rctDraw.Top = pctlGrid.Top _
            + pctlGrid.Groups(iLastGroup).Top _
            + mlngRecordProfileTopOffset _
            - plngPageTopPictureBoxPosition
          rctDraw.Bottom = rctDraw.Top _
            + plngGroupColumnHeadingHeight
            
          ' Call the function to print the group heading box.
          If Not PrintRecordProfileGridCell(rctDraw, _
            pctlGrid.StyleSets(pctlGrid.HeadStyleSet).Font, _
            pctlGrid.Groups(iLastGroup).Caption, _
            cellGROUPHEADING, _
            pctlGrid, _
            -1, _
            plngStartColumnPages, _
            (plngStartColumnPages > 0) And (iLastGroup = pctlGrid.Columns(plngStartColumn).Group)) Then
            
            fOK = False
            Resume ErrorTrap
          End If
        
          lngLeftMostLine = -1
        End If
          
        iLastGroup = pctlGrid.Columns(iLoop).Group
      End If
      
      ' Calculate the print coordinates of the column heading box
      rctDraw.Left = IIf((plngStartColumn = 0) And (plngStartColumnPages = 0), pctlGrid.Left, -pctlGrid.Columns(plngStartColumn).Left) _
        + pctlGrid.Columns(iLoop).Left _
        + plngMarginLeft_Twips
      If plngStartColumnPages > 0 Then
        If iLoop > plngStartColumn Then
          rctDraw.Left = rctDraw.Left - (plngStartColumnPages * (Printer.ScaleWidth - plngMarginLeft_Twips))
        End If
      End If

      rctDraw.Right = rctDraw.Left + pctlGrid.Columns(iLoop).Width
      If plngStartColumnPages > 0 Then
        If iLoop = plngStartColumn Then
          rctDraw.Right = rctDraw.Right _
            - (plngStartColumnPages * (Printer.ScaleWidth - plngMarginLeft_Twips))
        End If
      End If
      
      rctDraw.Top = pctlGrid.Top _
        + pctlGrid.Columns(iLoop).Top _
        + mlngRecordProfileTopOffset _
        - plngPageTopPictureBoxPosition
      rctDraw.Bottom = rctDraw.Top _
        + plngGroupColumnHeadingHeight
      
      If plngStartColumnPages > 0 Then
        If (plngStartColumn <> iLoop) Then
          rctDraw.Left = rctDraw.Left + (plngStartColumnPages * OVERLAPCOMPENSATION)
        End If
      
        rctDraw.Right = rctDraw.Right + (plngStartColumnPages * OVERLAPCOMPENSATION)
      End If
      
      ' Check if this column fits onto the current page.
      If rctDraw.Right > Printer.ScaleWidth Then
        If iLoop > plngStartColumn Then
          fSkipGroup = (iLastGroup <> iLastGroupToPrint)
        
          If fHasGroups Then
            If (iLastGroup >= 0) Then
              If ((iLastGroupToPrint > 0) _
                And (iLastGroup <> iLastGroupToPrint)) _
                And (lngLeftMostLine > 0) Then
                ' Calculate the print coordinates of the group heading box
                rctDraw.Left = lngLeftMostLine
                rctDraw.Right = lngRightMostLine
                rctDraw.Top = pctlGrid.Top _
                  + pctlGrid.Groups(pctlGrid.Columns(iLoop).Group).Top _
                  + mlngRecordProfileTopOffset _
                  - plngPageTopPictureBoxPosition
                rctDraw.Bottom = rctDraw.Top _
                  + plngGroupColumnHeadingHeight
  
                ' Call the function to print the column heading box.
                If Not PrintRecordProfileGridCell(rctDraw, _
                  pctlGrid.StyleSets(pctlGrid.HeadStyleSet).Font, _
                  pctlGrid.Groups(pctlGrid.Columns(iLoop).Group).Caption, _
                  cellGROUPHEADING, _
                  pctlGrid, _
                  -1, _
                  plngStartColumnPages, _
                  False) Then
                  
                  fOK = False
                  GoTo ErrorTrap
                End If
              End If
            End If
          End If
          
          ' Remember which columns of the grid need to be carried over to the overflow page.
          ReDim Preserve palngOverflows(3, UBound(palngOverflows, 2) + 1)
          palngOverflows(1, UBound(palngOverflows, 2)) = pctlGrid.Index
          palngOverflows(2, UBound(palngOverflows, 2)) = iLoop
          palngOverflows(3, UBound(palngOverflows, 2)) = 0
          Exit For
        Else
          fOntoNextPage = True
          rctDraw.Right = Printer.ScaleWidth
        
          ReDim Preserve palngOverflows(3, UBound(palngOverflows, 2) + 1)
          palngOverflows(1, UBound(palngOverflows, 2)) = pctlGrid.Index
          palngOverflows(2, UBound(palngOverflows, 2)) = iLoop
          palngOverflows(3, UBound(palngOverflows, 2)) = plngStartColumnPages + 1
        End If
      End If

      If pctlGrid.ColumnHeaders Then
        lngRightMostLine = rctDraw.Right
        If lngLeftMostLine < 0 Then
          lngLeftMostLine = rctDraw.Left
        End If

        iLastGroupToPrint = pctlGrid.Columns(iLoop).Group
        
        ' Call the function to print the column heading box.
        If Not PrintRecordProfileGridCell(rctDraw, _
          pctlGrid.StyleSets(pctlGrid.HeadStyleSet).Font, _
          pctlGrid.Columns(iLoop).Caption, _
          cellCOLUMNHEADING, _
          pctlGrid, _
          -1, _
          IIf(iLoop = plngStartColumn, plngStartColumnPages, 0), _
          (plngStartColumnPages > 0) And (plngStartColumn = iLoop)) Then
          
          fOK = False
          GoTo ErrorTrap
        End If

        If fOntoNextPage Then
          Exit For
        End If
      End If
    End If
  Next iLoop
  
  If fHasGroups And (Not fSkipGroup) Then
    ' Calculate the print coordinates of the group heading box
    rctDraw.Left = lngLeftMostLine
    rctDraw.Right = lngRightMostLine
    rctDraw.Top = pctlGrid.Top _
      + pctlGrid.Groups(iLastGroupToPrint).Top _
      + mlngRecordProfileTopOffset _
      - plngPageTopPictureBoxPosition
    rctDraw.Bottom = rctDraw.Top _
      + plngGroupColumnHeadingHeight

    ' Call the function to print the group heading box.
    If Not PrintRecordProfileGridCell(rctDraw, _
      pctlGrid.StyleSets(pctlGrid.HeadStyleSet).Font, _
      pctlGrid.Groups(iLastGroupToPrint).Caption, _
      cellGROUPHEADING, _
      pctlGrid, _
      -1, _
      IIf(iLoop = plngStartColumn, plngStartColumnPages, 0), _
      (plngStartColumnPages > 0) And (plngStartColumn = iLoop)) Then
      
      fOK = False
      GoTo ErrorTrap
    End If
  End If
  
TidyUpAndExit:

  PrintRecordProfileColumnHeadings = fOK
  Exit Function

ErrorTrap:
  If Err.Number <> 0 Then
    msErrorMessage = Err.Description
  End If
  
  PrintRecordProfileColumnHeadings = False
  
End Function


Private Function PrintRecordProfileGridCell( _
  prctDraw As Rect, _
  pfntFont As StdFont, _
  psCaption As String, _
  piCellType As GridCellType, _
  pctlGrid As SSDBGrid, _
  plngPictureIndex As Long, _
  plngStartColumnPages As Long, _
  pfFirstOverlapColumn As Boolean) As Boolean
  
  On Error GoTo ErrorTrap
    
  ' Print the griven cell of the grid.
  Dim rctDraw As Rect
  Dim rctDraw2 As Rect
  Dim iSM As Integer
  Dim iDrawWidth As Integer
  Dim iDrawStyle As Integer
  Dim iFillStyle As Integer
  'Dim fntTemp As StdFont
  Dim lngLoop As Long
  Dim lngTextLength As Long
  Dim lngTextDoneWidth As Long
  Dim sTemp As String
  Dim lngTempWidth As Long
  Dim iComparisonResult As Integer
  Dim lngLower As Long
  Dim lngUpper As Long
  Dim imgTemp As Image
  Dim sngWidthRatio As Single
  Dim sngHeightRatio As Single
  Dim sngMinRatio As Single

  'JPD 20050609 Fault 10060
  ' We used to create a font object (fntTemp), assign the required properties,
  ' and then Set Printer.font = fntTemp (see below).
  ' This caused printer errors on huge reports, so now we simply assign the required
  ' properties directly and individually to the Printer.Font object.
  
  ' Copy the font parameters to a temporary variable.
  ' By directly setting the printer's font property to be the grids font,
  ' and then changing the printer's font size, I actually managed to change the
  ' grid's font size as well. Using the temporary font avoids this.
  'Set fntTemp = New StdFont
  'With fntTemp
  '  .Name = pfntFont.Name
  '  .Size = pfntFont.Size
  '  .Bold = pfntFont.Bold
  '  .Underline = pfntFont.Underline
  '  .Strikethrough = pfntFont.Strikethrough
  '  .Italic = pfntFont.Italic
  'End With
  
  iSM = Printer.ScaleMode
  
  iDrawWidth = THINLINE
  If (piCellType = cellGROUPHEADING) Or _
    (piCellType = cellCOLUMNHEADING) Then
    iDrawWidth = THICKLINE
  End If
  
  iDrawStyle = IIf(mblnGridlines, vbSolid, vbInvisible)
  
  iFillStyle = vbFSTransparent
  If (piCellType = cellGROUPHEADING) Or _
    (piCellType = cellCOLUMNHEADING) Or _
    (piCellType = cellSEPARATOR) Then
    iFillStyle = IIf(mblnShading, vbFSSolid, vbFSTransparent)
  End If
  
  ' DRAW THE CELL BOX
  ' Set the printer's line and fill style, and draw the box
  Printer.DrawWidth = iDrawWidth
  Printer.DrawStyle = iDrawStyle
  Printer.FillStyle = iFillStyle
  Printer.FillColor = vbButtonFace
  
  Printer.Line (prctDraw.Left, prctDraw.Top)-(prctDraw.Right, prctDraw.Bottom), , B

  ' Ensure the right border is in bold if the cell overflows onto the next page.
  If prctDraw.Right = Printer.ScaleWidth Then
    rctDraw2.Left = prctDraw.Right - 10
    rctDraw2.Right = rctDraw2.Left
    rctDraw2.Top = prctDraw.Top
    rctDraw2.Bottom = prctDraw.Bottom
    Printer.DrawWidth = THICKLINE
    Printer.Line (rctDraw2.Left, rctDraw2.Top)-(rctDraw2.Right, rctDraw2.Bottom), , B
  End If

  Printer.DrawWidth = THINLINE
  Printer.DrawStyle = vbSolid
  Printer.FillStyle = vbFSTransparent
    
  If (piCellType = cellDATA) And _
    (plngPictureIndex >= 0) Then
    ' PRINT THE CELL PICTURE
    If (plngStartColumnPages <= 0) Or (Not pfFirstOverlapColumn) Then
      If Not pctlGrid.StyleSets(plngPictureIndex).Picture Is Nothing Then
        Set mfrmRecProfileForm.imgTemp.Picture = pctlGrid.StyleSets(plngPictureIndex).Picture
        
        If mfrmRecProfileForm.imgTemp.Height > (prctDraw.Bottom - prctDraw.Top - 30) Or _
          mfrmRecProfileForm.imgTemp.Width > (prctDraw.Right - prctDraw.Left - 70) Then
  
          sngWidthRatio = (prctDraw.Right - prctDraw.Left - 140) / mfrmRecProfileForm.imgTemp.Width
          sngHeightRatio = (prctDraw.Bottom - prctDraw.Top - 60) / mfrmRecProfileForm.imgTemp.Height
          sngMinRatio = IIf(sngWidthRatio < sngHeightRatio, sngWidthRatio, sngHeightRatio)
  
          Printer.PaintPicture pctlGrid.StyleSets(plngPictureIndex).Picture, prctDraw.Left + 70, prctDraw.Top + 30, (mfrmRecProfileForm.imgTemp.Width * sngMinRatio), (mfrmRecProfileForm.imgTemp.Height * sngMinRatio)
        Else
            Printer.PaintPicture pctlGrid.StyleSets(plngPictureIndex).Picture, prctDraw.Left + 70, prctDraw.Top + 30, , , , , prctDraw.Right, prctDraw.Bottom
        End If
      End If
    End If
  Else
    ' PRINT THE CELL TEXT
    ' Scale the print coordinates for use with the DrawText function.
    rctDraw.Left = Printer.ScaleX(prctDraw.Left, iSM, vbPixels) + IIf(piCellType <> cellGROUPHEADING, OFFSET_X, 0)
    rctDraw.Right = Printer.ScaleX(prctDraw.Right, iSM, vbPixels)
    rctDraw.Top = Printer.ScaleY(prctDraw.Top, iSM, vbPixels) + OFFSET_Y
    rctDraw.Bottom = Printer.ScaleY(prctDraw.Bottom, iSM, vbPixels)
            
    'JPD 20050609 Fault 10060
    ' Set the printer's font to match the cell's font.
    'Set Printer.Font = fntTemp
    'Set fntTemp = Nothing

    With Printer.Font
      .Name = pfntFont.Name
      .Size = pfntFont.Size
      .Bold = pfntFont.Bold
      .Underline = pfntFont.Underline
      .Strikethrough = pfntFont.Strikethrough
      .Italic = pfntFont.Italic
    End With

    If prctDraw.Right = Printer.ScaleWidth Then
      If Printer.Orientation = vbPRORPortrait Then
        ' Portrait grid
        rctDraw.Right = Printer.ScaleX(prctDraw.Right - 100, iSM, vbPixels)
      Else
        ' Landscape grid
        rctDraw.Right = Printer.ScaleX(prctDraw.Right - 125, iSM, vbPixels)
      End If
    End If
    
    If (plngStartColumnPages > 0) And pfFirstOverlapColumn Then
      lngTextDoneWidth = (Printer.ScaleWidth - 100 - prctDraw.Left) * plngStartColumnPages
      
      For lngLoop = 1 To Len(psCaption)
        lngTempWidth = TextWidth(Left(psCaption, lngLoop))
        
        If lngTempWidth > lngTextDoneWidth Then
          Exit For
        End If
      Next lngLoop
      
      psCaption = Mid(psCaption, lngLoop)
    End If

    ' Command the printer to print the cell text.
    DrawText Printer.hDC, Replace(psCaption, "&", "&&"), -1, rctDraw, IIf(piCellType = cellGROUPHEADING, DT_CENTER, DT_LEFT)
  End If
    
  PrintRecordProfileGridCell = True
  
  Exit Function

ErrorTrap:
  If Err.Number <> 0 Then
    msErrorMessage = Err.Description
  End If
  
  PrintRecordProfileGridCell = False

End Function



Function TextWidth(psInString As String) As Long
  ' TextWidth method of the Printer object causes a runtime error when used with long strings.
  ' MSDN article Q298825 suggests using the following function instead.
  Const iBLOCKSIZE = 500
  
  If Len(psInString) > iBLOCKSIZE Then
    TextWidth = Printer.TextWidth(Left(psInString, iBLOCKSIZE)) + _
      TextWidth(Right(psInString, Len(psInString) - iBLOCKSIZE))
  Else
    TextWidth = TextWidth + Printer.TextWidth(psInString)
  End If
    
End Function

Private Function PrintRecordProfileOverflowPage( _
  pfrmRecProfile As Form, _
  palngOverflows As Variant, _
  psTitle As String, _
  psLeftFooter As String, _
  plngMarginLeft_Twips As Long, _
  plngMarginRight_Twips As Long, _
  plngMarginTop_Twips As Long, _
  plngMarginBottom_Twips As Long, _
  plngPageTopPictureBoxPosition As Long) As Boolean
  
  On Error GoTo ErrorTrap
    
  ' Print the columns of the grids that did not fit on the previous page.
  Dim fOK As Boolean
  Dim blnTempHeadFontUnderline As Boolean
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim iLoop3 As Integer
  Dim iLoop4 As Integer
  Dim iLoop5 As Integer
  Dim iLoop6 As Integer
  Dim ctlTemp As Control
  Dim rctDraw As Rect
  Dim rctDraw2 As Rect
  Dim iSM As Integer
  Dim varBookmark As Variant
  Dim fGridPreceded As Boolean
  Dim fGridFollowed As Boolean
  Dim sTitle As String
  Dim sTemp As String
  Dim iOldScaleMode As Integer
  Dim fHasHeadingColumn As Boolean
  Dim fPhotoDone As Boolean
  Dim iLastVisibleColumn As Integer
  Dim fNewPage As Boolean
  Dim sLeftFooter As String
  Dim sRightFooter As String
  Dim lngGroupColumnHeadingHeight As Long
  Dim objRecProfTable As clsRecordProfileTabDtl
  Dim fPageBreak As Boolean
  Dim alngOverflows() As Long
  Dim fFound As Boolean
  Dim fntFont As StdFont
  Dim iCellType As GridCellType
  Dim lngPictureIndex As Long
  Dim fHeadingsDone As Boolean
  Dim lngPageTopPictureBoxPosition As Long
  Dim lngTemp As Long
  Dim fFirstRow As Long
  Dim fCellsPrinted As Boolean

  fOK = True
  iSM = Printer.ScaleMode
  fCellsPrinted = False

  ' Column 1 = grid index
  ' Column 2 = grid column being carried over onto overflow page
  ' Column 3 = number of pages the first column has already been printed on (in case it too wide for a single page)
  ReDim alngOverflows(3, 0)
  
  ' Get the printer to start printing on a new page.
  miRecordProfilePageCounter = miRecordProfilePageCounter + 1
  Printer.NewPage

  sRightFooter = "Page " & CStr(miRecordProfilePageCounter) & " (continued from page " & CStr(miRecordProfilePageCounter - 1) & ")"
  If Not PrintRecordProfilePageHeaderAndFooter(psTitle, _
    psLeftFooter, _
    sRightFooter, _
    plngMarginLeft_Twips, _
    plngMarginRight_Twips, _
    plngMarginTop_Twips, _
    plngMarginBottom_Twips) Then
    
    fOK = False
    Resume TidyUpAndExit
  End If

  For iLoop6 = 1 To UBound(palngOverflows, 2)
    Set ctlTemp = pfrmRecProfile.grdOutput(palngOverflows(1, iLoop6))
    
    ' Check if we need to page break after this table's output.
    Set objRecProfTable = pfrmRecProfile.Definition.Item(ctlTemp.Tag)
    fPageBreak = objRecProfTable.PageBreak
          
    ' Check if this grid has Heading columns,
    ' and get the indexes of the first and last visible columns.
    fHasHeadingColumn = False
    iLastVisibleColumn = 0
    For iLoop2 = palngOverflows(2, iLoop6) To ctlTemp.Columns.Count - 1
      If ctlTemp.Columns(iLoop2).Name = COLUMN_ISHEADING Then
        fHasHeadingColumn = True
      End If
            
      If ctlTemp.Columns(iLoop2).Visible Then
        iLastVisibleColumn = iLoop2
      End If
    Next iLoop2
          
    ' Check if this grid is preceded or followed IMMEDIATELY by other grids.
    ' ie. if this grid is part of a group of grids that are used to display
    ' data (including pictures) vertically.
    ' NB. Grids have only one row height value. To display pictures with their
    ' own row height, we actually put them in their own grid, and position this grid
    ' IMMEDIATELY after the normal data grid. Subsequent data is put in its own grid
    ' IMMEDIATELY after the picture's grid.
    ' This is waht's meant by 'following' and 'preceding' grids.
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
          
    ' Determine the height of the Group and Column Heading rows,
    ' as this is not the same as the grid's rowHeight.
    lngTemp = 0
    lngGroupColumnHeadingHeight = 0
    
    If ctlTemp.ColumnHeaders Then
      lngTemp = lngGroupColumnHeadingHeight
      lngGroupColumnHeadingHeight = ctlTemp.Height - (ctlTemp.RowHeight * ctlTemp.Rows)
      
      If ((ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders)) Then
        lngGroupColumnHeadingHeight = ctlTemp.Columns(0).Top
        lngTemp = 2 * lngGroupColumnHeadingHeight
      End If
    End If
    
    ' Draw data rows and columns.
    fHeadingsDone = False
    fFirstRow = Not ctlTemp.ColumnHeaders
    
    For iLoop = 0 To ctlTemp.Rows - 1

      ' Check if the next row will go beyond the bottom of the page.
      rctDraw.Top = ctlTemp.Top + (iLoop * ctlTemp.RowHeight) + mlngRecordProfileTopOffset - plngPageTopPictureBoxPosition
      If (ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders) Then
        rctDraw.Top = rctDraw.Top + lngGroupColumnHeadingHeight
      End If
      If ctlTemp.ColumnHeaders Then
        rctDraw.Top = rctDraw.Top + lngGroupColumnHeadingHeight
      End If
      rctDraw.Bottom = rctDraw.Top + ctlTemp.RowHeight

      If (Not fCellsPrinted) And (rctDraw.Bottom > mlngRecordProfileBottomOffset) Then
        rctDraw.Bottom = mlngRecordProfileBottomOffset
      End If

      fNewPage = (rctDraw.Bottom > mlngRecordProfileBottomOffset)
      If fNewPage Then
        ' Make sure the last row on a page has it's bottom line in bold.
        If mblnGridlines Then
          If iLoop > 0 Then
            For iLoop2 = palngOverflows(2, iLoop6) To ctlTemp.Columns.Count - 1
              If ctlTemp.Columns(iLoop2).Visible Then
                rctDraw.Left = ctlTemp.Columns(iLoop2).Left - ctlTemp.Columns(CLng(palngOverflows(2, iLoop6))).Left + plngMarginLeft_Twips
                rctDraw.Right = rctDraw.Left + ctlTemp.Columns(iLoop2).Width
                rctDraw.Top = ctlTemp.Top + ((iLoop - 1) * ctlTemp.RowHeight) + mlngRecordProfileTopOffset - plngPageTopPictureBoxPosition
                If (ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders) Then
                  rctDraw.Top = rctDraw.Top + lngGroupColumnHeadingHeight
                End If
                If ctlTemp.ColumnHeaders Then
                  rctDraw.Top = rctDraw.Top + lngGroupColumnHeadingHeight
                End If
                rctDraw.Bottom = rctDraw.Top + ctlTemp.RowHeight

                If CLng(palngOverflows(3, iLoop6)) > 0 Then
                  If iLoop2 = CLng(palngOverflows(2, iLoop6)) Then
                    rctDraw.Right = rctDraw.Right _
                      - (CLng(palngOverflows(3, iLoop6)) * (Printer.ScaleWidth - plngMarginLeft_Twips))
                  End If
                End If
                                
                If CLng(palngOverflows(3, iLoop6)) > 0 Then
                  If Not ((CLng(palngOverflows(3, iLoop6)) > 0) And (CLng(palngOverflows(2, iLoop6)) = iLoop2)) Then
                    rctDraw.Left = rctDraw.Left + (CLng(palngOverflows(3, iLoop6)) * OVERLAPCOMPENSATION)
                  End If
                
                  rctDraw.Right = rctDraw.Right + (CLng(palngOverflows(3, iLoop6)) * OVERLAPCOMPENSATION)
                End If
                
                rctDraw2.Left = rctDraw.Left
                rctDraw2.Right = rctDraw.Right
                rctDraw2.Top = rctDraw.Bottom
                rctDraw2.Bottom = rctDraw.Bottom
                
                If rctDraw2.Right > Printer.ScaleWidth Then
                  If iLoop2 > palngOverflows(2, iLoop6) Then
                    Exit For
                  Else
                    rctDraw2.Right = Printer.ScaleWidth
                  End If
                End If
                
                Printer.DrawWidth = THICKLINE
                Printer.Line (rctDraw2.Left, rctDraw2.Top)-(rctDraw2.Right, rctDraw2.Bottom), , B
                Printer.DrawWidth = THINLINE
              End If
            Next iLoop2
          End If
        End If
        
        Exit For
      End If
      
      If (rctDraw.Top - lngTemp) >= mlngRecordProfileTopOffset Then
        If Not fHeadingsDone Then
          lngPageTopPictureBoxPosition = ctlTemp.Top + (iLoop * ctlTemp.RowHeight)
          
          ' Print group and column headers if required.
          If Not PrintRecordProfileColumnHeadings(ctlTemp, _
            plngMarginLeft_Twips, _
            plngPageTopPictureBoxPosition - (iLoop * ctlTemp.RowHeight), _
            lngGroupColumnHeadingHeight, _
            alngOverflows, _
            CLng(palngOverflows(2, iLoop6)), _
            CLng(palngOverflows(3, iLoop6))) Then
            
            fOK = False
            GoTo ErrorTrap
          End If
          
          fHeadingsDone = True
        End If
  
        For iLoop2 = palngOverflows(2, iLoop6) To ctlTemp.Columns.Count - 1
          If ctlTemp.Columns(iLoop2).Visible Then

            rctDraw.Left = ctlTemp.Columns(iLoop2).Left _
              - ctlTemp.Columns(CLng(palngOverflows(2, iLoop6))).Left _
              + plngMarginLeft_Twips
            If CLng(palngOverflows(3, iLoop6)) > 0 Then
              If iLoop2 > palngOverflows(2, iLoop6) Then
                rctDraw.Left = rctDraw.Left - (CLng(palngOverflows(3, iLoop6)) * (Printer.ScaleWidth - plngMarginLeft_Twips))
              End If
            End If

            rctDraw.Right = rctDraw.Left + ctlTemp.Columns(iLoop2).Width
            If CLng(palngOverflows(3, iLoop6)) > 0 Then
              If iLoop2 = CLng(palngOverflows(2, iLoop6)) Then
                rctDraw.Right = rctDraw.Right _
                  - (CLng(palngOverflows(3, iLoop6)) * (Printer.ScaleWidth - plngMarginLeft_Twips))
              End If
            End If
          
            rctDraw.Top = ctlTemp.Top _
              + (iLoop * ctlTemp.RowHeight) _
              + mlngRecordProfileTopOffset _
              - plngPageTopPictureBoxPosition
            If (ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders) Then
              rctDraw.Top = rctDraw.Top + lngGroupColumnHeadingHeight
            End If
            If ctlTemp.ColumnHeaders Then
              rctDraw.Top = rctDraw.Top + lngGroupColumnHeadingHeight
            End If
            rctDraw.Bottom = rctDraw.Top + ctlTemp.RowHeight

            If (rctDraw.Bottom > mlngRecordProfileBottomOffset) Then
              rctDraw.Bottom = mlngRecordProfileBottomOffset
            End If

            If CLng(palngOverflows(3, iLoop6)) > 0 Then
              If Not ((CLng(palngOverflows(3, iLoop6)) > 0) And (CLng(palngOverflows(2, iLoop6)) = iLoop2)) Then
                rctDraw.Left = rctDraw.Left + (CLng(palngOverflows(3, iLoop6)) * OVERLAPCOMPENSATION)
              End If
            
              rctDraw.Right = rctDraw.Right + (CLng(palngOverflows(3, iLoop6)) * OVERLAPCOMPENSATION)
            End If
  
            If rctDraw.Right > Printer.ScaleWidth Then
              If iLoop2 > palngOverflows(2, iLoop6) Then
                If mblnGridlines Then
                  ' Ensure the last printed data cell has a bold line on the right-hand side.
                  rctDraw2.Left = rctDraw.Left
                  rctDraw2.Right = rctDraw.Left
                  rctDraw2.Top = rctDraw.Top
                  rctDraw2.Bottom = rctDraw.Bottom
                  Printer.DrawWidth = THICKLINE
                  Printer.Line (rctDraw2.Left, rctDraw2.Top)-(rctDraw2.Right, rctDraw2.Bottom), , B
                  Printer.DrawWidth = THINLINE
                End If
  
                Exit For
              Else
                rctDraw.Right = Printer.ScaleWidth
              End If
            End If

            If (ctlTemp.Columns(iLoop2).Style = 4) And _
              (Not ctlTemp.ColumnHeaders) Then

              ' Button
              Set fntFont = ctlTemp.Font
              varBookmark = ctlTemp.AddItemBookmark(iLoop)
              If fHasHeadingColumn Then
                If ctlTemp.Columns(COLUMN_ISHEADING).CellText(varBookmark) = "1" Then
                  Set fntFont = ctlTemp.StyleSets("Heading").Font
                End If
              End If

              fCellsPrinted = True
              
              If Not PrintRecordProfileGridCell(rctDraw, _
                fntFont, _
                ctlTemp.Columns(iLoop2).CellText(varBookmark), _
                cellCOLUMNHEADING, _
                ctlTemp, _
                -1, _
                CLng(palngOverflows(3, iLoop6)), _
                (CLng(palngOverflows(3, iLoop6)) > 0) And (CLng(palngOverflows(2, iLoop6)) = iLoop2)) Then
                
                fOK = False
                GoTo ErrorTrap
              End If
            
              Set fntFont = Nothing
            Else
              ' Normal data cell
              iCellType = cellDATA

              If fHasHeadingColumn Then
                varBookmark = ctlTemp.AddItemBookmark(iLoop)
                If ctlTemp.Columns(COLUMN_ISHEADING).CellText(varBookmark) = "1" Then
                  iCellType = cellSEPARATOR
                End If
              End If

              If ctlTemp.Columns(iLoop2).StyleSet = "Separator" Then
                iCellType = cellSEPARATOR
              End If

              lngPictureIndex = -1
              fPhotoDone = False
              If ctlTemp.TagVariant = COLUMN_ISPHOTO Then
                For iLoop3 = 0 To ctlTemp.Columns.Count - 1
                  If ctlTemp.Columns(iLoop3).Visible Then
                    sTemp = PHOTOSTYLESET & CStr(iLoop3 + 1)

                    For iLoop4 = 0 To ctlTemp.StyleSets.Count - 1
                      If ctlTemp.StyleSets(iLoop4).Name = sTemp Then
                        lngPictureIndex = iLoop4
                        fPhotoDone = True
                        Exit For
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
                    lngPictureIndex = iLoop4
                    fPhotoDone = True
                    Exit For
                  End If
                Next iLoop4
              End If

              Set fntFont = ctlTemp.Font
              varBookmark = ctlTemp.AddItemBookmark(iLoop)
              If fHasHeadingColumn Then
                If ctlTemp.Columns(COLUMN_ISHEADING).CellText(varBookmark) = "1" Then
                  Set fntFont = ctlTemp.StyleSets("Heading").Font
                End If
              End If
              
              fCellsPrinted = True

              If Not PrintRecordProfileGridCell(rctDraw, _
                fntFont, _
                ctlTemp.Columns(iLoop2).CellText(varBookmark), _
                iCellType, _
                ctlTemp, _
                lngPictureIndex, _
                CLng(palngOverflows(3, iLoop6)), _
                (CLng(palngOverflows(3, iLoop6)) > 0) And (CLng(palngOverflows(2, iLoop6)) = iLoop2)) Then
                
                fOK = False
                GoTo ErrorTrap
              End If

              Set fntFont = Nothing
              
              If mblnGridlines Then
                ' Top rows of the grid must have the top line in bold.
                If (fFirstRow And (Not fGridPreceded)) Or _
                  fNewPage Then
                  
                  fFirstRow = False
                  
                  rctDraw2.Left = rctDraw.Left
                  rctDraw2.Right = rctDraw.Right
                  rctDraw2.Top = rctDraw.Top
                  rctDraw2.Bottom = rctDraw.Top
                  Printer.DrawWidth = THICKLINE
                  Printer.Line (rctDraw2.Left, rctDraw2.Top)-(rctDraw2.Right, rctDraw2.Bottom), , B
                  Printer.DrawWidth = THINLINE
                End If

                ' Last columns on the page must have the right line in bold.
                If iLoop2 = iLastVisibleColumn Then
                  rctDraw2.Left = rctDraw.Right
                  rctDraw2.Right = rctDraw.Right
                  rctDraw2.Top = rctDraw.Top
                  rctDraw2.Bottom = rctDraw.Bottom
                  Printer.DrawWidth = THICKLINE
                  Printer.Line (rctDraw2.Left, rctDraw2.Top)-(rctDraw2.Right, rctDraw2.Bottom), , B
                  Printer.DrawWidth = THINLINE
                End If

                ' First columns on the page must have the left line in bold.
                If iLoop2 = palngOverflows(2, iLoop6) Then
                  rctDraw2.Left = rctDraw.Left
                  rctDraw2.Right = rctDraw.Left
                  rctDraw2.Top = rctDraw.Top
                  rctDraw2.Bottom = rctDraw.Bottom
                  Printer.DrawWidth = THICKLINE
                  Printer.Line (rctDraw2.Left, rctDraw2.Top)-(rctDraw2.Right, rctDraw2.Bottom), , B
                  Printer.DrawWidth = THINLINE
                End If

                ' Bottom rows of the grid must have the bottom line in bold.
                If (iLoop = (ctlTemp.Rows - 1)) And (Not fGridFollowed) Then
                  rctDraw2.Left = rctDraw.Left
                  rctDraw2.Right = rctDraw.Right
                  rctDraw2.Top = rctDraw.Bottom
                  rctDraw2.Bottom = rctDraw.Bottom
                  Printer.DrawWidth = THICKLINE
                  Printer.Line (rctDraw2.Left, rctDraw2.Top)-(rctDraw2.Right, rctDraw2.Bottom), , B
                  Printer.DrawWidth = THINLINE
                End If
              End If
            End If
          End If
        Next iLoop2
      End If
    Next iLoop
  
    Set objRecProfTable = Nothing
    Set ctlTemp = Nothing
  Next iLoop6
  
  ' Print any columns that did not even fit on this overflow page.
  If UBound(alngOverflows, 2) > 0 Then
    If Not PrintRecordProfileOverflowPage(pfrmRecProfile, _
      alngOverflows, _
      psTitle, _
      psLeftFooter, _
      plngMarginLeft_Twips, _
      plngMarginRight_Twips, _
      plngMarginTop_Twips, _
      plngMarginBottom_Twips, _
      plngPageTopPictureBoxPosition) Then
      
      fOK = False
      GoTo TidyUpAndExit
    End If
  End If
  
TidyUpAndExit:
  PrintRecordProfileOverflowPage = fOK
  Exit Function

ErrorTrap:
  If Err.Number <> 0 Then
    msErrorMessage = Err.Description
  End If
  
  PrintRecordProfileOverflowPage = False
  
End Function

Private Function PrintRecordProfilePageHeaderAndFooter( _
  psHeader As String, _
  psLeftFooter As String, _
  psRightFooter As String, _
  plngMarginLeft_Twips As Long, _
  plngMarginRight_Twips As Long, _
  plngMarginTop_Twips As Long, _
  plngMarginBottom_Twips As Long) As Boolean
  
  On Error GoTo ErrorTrap
    
  Const iGAPUNDERTITLE = 5 * TWIPSPERMM
  
  With Printer
    ' Print the header.
    If (miRecordProfilePageCounter = 1) Or mblnHeadingsOnEveryPage Then
      'JPD 20050609 Fault 10060
      ' Printer errors caused when assigning the Printer.Font object with another Font object.
      ' Found it better to simply assign values to the individual properties.
      
      'Set .Font = mfntRecordProfileHeaderFont
      .Font.Name = mfntRecordProfileHeaderFont.Name
      .Font.Size = mfntRecordProfileHeaderFont.Size
      .Font.Bold = mfntRecordProfileHeaderFont.Bold
      .Font.Underline = mfntRecordProfileHeaderFont.Underline
      .Font.Strikethrough = mfntRecordProfileHeaderFont.Strikethrough
      .Font.Italic = mfntRecordProfileHeaderFont.Italic
      
      .CurrentX = plngMarginLeft_Twips + _
        ((.ScaleWidth - plngMarginLeft_Twips - plngMarginRight_Twips) / 2) - _
        (.TextWidth(psHeader) / 2)
      .CurrentY = plngMarginTop_Twips
      Printer.Print psHeader
      mlngRecordProfileTopOffset = .CurrentY + iGAPUNDERTITLE
    End If
    
    ' Print the footer.
    'JPD 20050609 Fault 10060
    ' Printer errors caused when assigning the Printer.Font object with another Font object.
    ' Found it better to simply assign values to the individual properties.
    
    'Set .Font = mfntRecordProfileFooterFont
    .Font.Name = mfntRecordProfileFooterFont.Name
    .Font.Size = mfntRecordProfileFooterFont.Size
    .Font.Bold = mfntRecordProfileFooterFont.Bold
    .Font.Underline = mfntRecordProfileFooterFont.Underline
    .Font.Strikethrough = mfntRecordProfileFooterFont.Strikethrough
    .Font.Italic = mfntRecordProfileFooterFont.Italic
    
    .CurrentY = .ScaleHeight - plngMarginBottom_Twips - .TextHeight(psLeftFooter)
    mlngRecordProfileBottomOffset = .CurrentY - iGAPUNDERTITLE
    .CurrentX = plngMarginLeft_Twips
    Printer.Print psLeftFooter
    
    .CurrentX = (.ScaleWidth - plngMarginRight_Twips) - .TextWidth(psRightFooter)
    Printer.Print psRightFooter
  End With
    
  PrintRecordProfilePageHeaderAndFooter = True
  Exit Function
  
  Exit Function

ErrorTrap:
  If Err.Number <> 0 Then
    msErrorMessage = Err.Description
  End If
  
  PrintRecordProfilePageHeaderAndFooter = False

End Function

Public Property Get PrintSelectedRowsOnly() As Boolean
  PrintSelectedRowsOnly = mblnPrintSelectedRowsOnly
End Property

Public Property Let PrintSelectedRowsOnly(ByVal blnNewValue As Boolean)
  mblnPrintSelectedRowsOnly = blnNewValue
End Property

Public Property Get SuppressPrompt() As Boolean
  SuppressPrompt = mblnSuppressPrompt
End Property

Public Property Let SuppressPrompt(ByVal blnNewValue As Boolean)
  mblnSuppressPrompt = blnNewValue
End Property

Public Property Get Heading() As String
  Heading = mstrHeading
End Property

Public Property Let Heading(ByVal strNewValue As String)
  mstrHeading = strNewValue
End Property

Public Property Get Gridlines() As Boolean
  Gridlines = mblnGridlines
End Property

Public Property Let Gridlines(ByVal intNewValue As Boolean)
  mblnGridlines = intNewValue
End Property

Public Property Get NumberOfCopies() As Integer
  NumberOfCopies = mintNumberOfCopies
End Property

Public Property Let NumberOfCopies(ByVal intNewValue As Integer)
  mintNumberOfCopies = intNewValue
End Property

Public Property Get CollateCopies() As Boolean
  CollateCopies = mblnCollate
End Property

Public Property Let CollateCopies(ByVal intNewValue As Boolean)
  mblnCollate = intNewValue
End Property

Public Property Get Portrait() As Boolean
  Portrait = mblnPortrait
End Property

Public Property Let Portrait(ByVal blnNewValue As Boolean)
  mblnPortrait = blnNewValue
End Property

Public Property Get Grid() As SSDataWidgets_B.SSDBGrid
  Set Grid = mgrdTemp
End Property

Public Property Let Grid(ByVal blnNewValue As SSDataWidgets_B.SSDBGrid)
  Set mgrdTemp = blnNewValue
End Property

Public Property Get MarginTop() As Integer
  MarginTop = mintMarginTop
End Property

Public Property Let MarginTop(ByVal intNewValue As Integer)
  mintMarginTop = intNewValue
End Property

Public Property Get MarginBottom() As Integer
  MarginBottom = mintMarginBottom
End Property

Public Property Let MarginBottom(ByVal intNewValue As Integer)
  mintMarginBottom = intNewValue
End Property

Public Property Get MarginLeft() As Integer
  MarginLeft = mintMarginLeft
End Property

Public Property Let MarginLeft(ByVal intNewValue As Integer)
  mintMarginLeft = intNewValue
End Property

Public Property Get MarginRight() As Integer
  MarginRight = mintMarginRight
End Property

Public Property Let MarginRight(ByVal intNewValue As Integer)
  mintMarginRight = intNewValue
End Property

Private Sub Class_Initialize()
  mblnSuppressPrompt = False
  mintNumberOfCopies = 1
  mblnCollate = True
  mblnPortrait = False
  mblnGridlines = True
  mblnShading = True
  mblnHeadingsOnEveryPage = True
  
  If UI.GetSystemMeasurement = "us" Then ' Inches
    mintMarginTop = 1
    mintMarginBottom = 1
    mintMarginLeft = 1
    mintMarginRight = 1
  Else
    mintMarginTop = 10 ' MM
    mintMarginBottom = 10
    mintMarginLeft = 10
    mintMarginRight = 10
  End If
  
  mblnPrintSelectedRowsOnly = False
  
  Set mfntRecordProfileHeaderFont = New StdFont
  With mfntRecordProfileHeaderFont
    .Name = "Verdana"
    .Size = 12
    .Bold = True
    .Underline = True
  End With
  
  Set mfntRecordProfileFooterFont = New StdFont
  With mfntRecordProfileFooterFont
    .Name = "Verdana"
    .Size = 8
    .Bold = False
    .Underline = False
  End With
  
End Sub

Public Sub PrintGrid(Optional pvarLastPage As Variant)

  Dim blnHideProgress As Boolean

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "clsPrintGrid.PrintGrid()"

  Dim blnTempHeadFontUnderline As Boolean
  
  If Not mblnSuppressPrompt Then
    
    With frmPrintOptions
      Screen.MousePointer = vbDefault
     
      blnHideProgress = gobjProgress.Visible
      If blnHideProgress Then
        If gobjProgress.Cancelled Then
          GoTo TidyUpAndExit
        End If
        gobjProgress.Visible = False
      End If

     .Show vbModal
      
      If .Cancelled Then
        'COAMsgBox "Printing cancelled by user", vbExclamation
        mblnCancelled = True
        Unload frmPrintOptions
        Set frmPrintOptions = Nothing
        Exit Sub
      End If

      If blnHideProgress Then
        gobjProgress.Visible = True
      End If

      mblnGridlines = .PrintGrid
      mintNumberOfCopies = .PrintCopies
      mblnPortrait = .PrintPortrait
      mblnCollate = .CollateCopies
      mblnShading = .Shading
      mblnHeadingsOnEveryPage = .HeadingsOnEveryPage

      mintMarginTop = .MarginTop
      mintMarginBottom = .MarginBottom
      mintMarginLeft = .MarginLeft
      mintMarginRight = .MarginRight

'      Unload frmPrintOptions
'      Set frmPrintOptions = Nothing

    End With
  
  End If

  mstrSelectedPrinter = Printer.DeviceName

  'TM20020325 Fault 2756 - if we are printing a selection of the report we need to use the
  'mgrdSelected grid control as this is populated for Page on change functionality, else
  'use mgrdTemp for printing all rows on the grid.
  If mblnPrintSelectedRowsOnly = True Then
    With mgrdSelected
      .Redraw = False
      Screen.MousePointer = vbHourglass
      
      blnTempHeadFontUnderline = .HeadFont.Underline
      .HeadFont.Underline = Not (mblnGridlines)
      
      .PageHeaderFont.Name = "Verdana"
      .PageHeaderFont.Size = 12
      .PageHeaderFont.Bold = True
      .PageHeaderFont.Underline = True
      
      .PageFooterFont.Name = "Verdana"
      .PageFooterFont.Size = 8
      .PageFooterFont.Bold = False
      .PageFooterFont.Underline = False
      
      .PrintData ssPrintAllRows + ssPrintFieldOrder, False, False
      
      .HeadFont.Underline = blnTempHeadFontUnderline
      .Redraw = True
      Screen.MousePointer = vbDefault
    End With

  Else
    With mgrdTemp
      .Redraw = False
      Screen.MousePointer = vbHourglass
      
      blnTempHeadFontUnderline = .HeadFont.Underline
      .HeadFont.Underline = Not (mblnGridlines)
      
      .PageHeaderFont.Name = "Verdana"
      .PageHeaderFont.Size = 12
      .PageHeaderFont.Bold = True
      .PageHeaderFont.Underline = True
      
      .PageFooterFont.Name = "Verdana"
      .PageFooterFont.Size = 8
      .PageFooterFont.Bold = False
      .PageFooterFont.Underline = False
      
  '    If mblnPrintSelectedRowsOnly = True Then
  '      .PrintData ssPrintSelectedRows + ssPrintFieldOrder, False, False
  '    Else
        .PrintData ssPrintAllRows + ssPrintFieldOrder, False, False
  '    End If
    
      .HeadFont.Underline = blnTempHeadFontUnderline
      .Redraw = True
      Screen.MousePointer = vbDefault
    End With
  End If

'TM20030217 Fault 5059
'  If Not IsMissing(pbLastPage) Then
'        If pbLastPage Then
'      ' JDM - 29/08/01 - Fault 2758 - Should display prompt after printing
'      PrintConfirm
'
'      Unload frmPrintOptions
'      Set frmPrintOptions = Nothing
'    End If
'
'  Else
'    ' JDM - 29/08/01 - Fault 2758 - Should display prompt after printing
'    PrintConfirm
'
'    Unload frmPrintOptions
'    Set frmPrintOptions = Nothing
'
'  End If
  If IsMissing(pvarLastPage) Then
    ' JDM - 29/08/01 - Fault 2758 - Should display prompt after printing
    PrintConfirm

    Unload frmPrintOptions
    Set frmPrintOptions = Nothing

  Else
    If pvarLastPage Then
      ' JDM - 29/08/01 - Fault 2758 - Should display prompt after printing
      PrintConfirm

      Unload frmPrintOptions
      Set frmPrintOptions = Nothing
    End If

  End If


  'MH20030922 Fault 6124 (Q257688)
  'Printer.TrackDefault = True

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub

Public Function PrintRecordProfilePage( _
  pfrmRecProfile As Form, _
  piPageNumber As Integer) As Boolean
  ' Print the record profile page.
  ' NOTE - all positioning is made using TWIPS.
  
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "clsPrintGrid.PrintRecordProfilePage()"
  
  msErrorMessage = ""
  
  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim iLoop3 As Integer
  Dim iLoop4 As Integer
  Dim iLoop5 As Integer
  Dim ctlTemp As Control
  Dim rctDraw As Rect
  Dim rctDraw2 As Rect
  Dim iSM As Integer
  Dim varBookmark As Variant
  Dim fGridPreceded As Boolean
  Dim fGridFollowed As Boolean
  Dim sTitle As String
  Dim sTemp As String
  Dim lngMarginTop_Twips As Long
  Dim lngMarginBottom_Twips As Long
  Dim lngMarginLeft_Twips As Long
  Dim lngMarginRight_Twips As Long
  Dim iOldScaleMode As Integer
  Dim fHasHeadingColumn As Boolean
  Dim fPhotoDone As Boolean
  Dim iFirstVisibleColumn As Integer
  Dim iLastVisibleColumn As Integer
  Dim lngPageTopPictureBoxPosition As Long
  Dim fNewPage As Boolean
  Dim sLeftFooter As String
  Dim sRightFooter As String
  Dim lngGroupColumnHeadingHeight As Long
  Dim objRecProfTable As clsRecordProfileTabDtl
  Dim fPageBreak As Boolean
  Dim alngOverflows() As Long
  Dim fFound As Boolean
  Dim ctlTempGrid As SSDBGrid
  Dim lngTemp As Long
  Dim fntFont As StdFont
  Dim iCellType As GridCellType
  Dim lngPictureIndex As Long
  Dim blnHideProgress As Boolean    'MH20030905 Fault 6819
  Dim fCellsPrinted As Boolean
  
  fOK = True
  sTitle = pfrmRecProfile.Caption
  iOldScaleMode = vbMillimeters
  
  Set mfrmRecProfileForm = pfrmRecProfile
  
  fCellsPrinted = False

  If Not mblnSuppressPrompt Then
    With frmPrintOptions
      .DenyCollate = True
      Screen.MousePointer = vbDefault
     
      blnHideProgress = gobjProgress.Visible
      If blnHideProgress Then
        If gobjProgress.Cancelled Then
          GoTo TidyUpAndExit
        End If
        gobjProgress.Visible = False
      End If
     
      .Show vbModal
      If .Cancelled Then
        'COAMsgBox "Printing cancelled by user", vbExclamation
        mblnCancelled = True
        Unload frmPrintOptions
        Set frmPrintOptions = Nothing
        Exit Function
      End If

      If blnHideProgress Then
        gobjProgress.Visible = True
      End If

      mstrSelectedPrinter = Printer.DeviceName

      mblnGridlines = .PrintGrid
      Printer.Copies = .PrintCopies
      Printer.Orientation = IIf(.PrintPortrait, vbPRORPortrait, vbPRORLandscape)
      mblnShading = .Shading
      mblnHeadingsOnEveryPage = .HeadingsOnEveryPage

      mintMarginTop = .MarginTop
      mintMarginBottom = .MarginBottom
      mintMarginLeft = .MarginLeft
      mintMarginRight = .MarginRight
    End With
  End If
  
  ' Calculate the page margins.
  lngMarginTop_Twips = (mintMarginTop * TWIPSPERMM) - MARGINCORRECTION
  lngMarginBottom_Twips = (mintMarginBottom * TWIPSPERMM)
  lngMarginLeft_Twips = (mintMarginLeft * TWIPSPERMM) - MARGINCORRECTION
  lngMarginRight_Twips = (mintMarginRight * TWIPSPERMM) - MARGINCORRECTION
  
  ' Initialise the page counter, or create a new page, as required.
  If piPageNumber = 1 Then
    miRecordProfilePageCounter = 1
  Else
    miRecordProfilePageCounter = miRecordProfilePageCounter + 1
    Printer.NewPage
  End If
  lngPageTopPictureBoxPosition = 0

  ' Remember the printer's current scalemode.
  ' We'll be using twips but will restore the original setting when we've finished.
  iOldScaleMode = Printer.ScaleMode
  Printer.ScaleMode = vbTwips
    
  ' Create the footer text.
  sLeftFooter = "Printed on " & Format(Now, DateFormat) & _
    " at " & Format(Now, "hh:nn") & " by " & gsUserName
  sRightFooter = "Page " & CStr(miRecordProfilePageCounter)
  
  ' Print the header and footer text on the current page.
  If Not PrintRecordProfilePageHeaderAndFooter(sTitle, _
    sLeftFooter, _
    sRightFooter, _
    lngMarginLeft_Twips, _
    lngMarginRight_Twips, _
    lngMarginTop_Twips, _
    lngMarginBottom_Twips) Then
    
    fOK = False
    GoTo SubErrorTrap
  End If

  iSM = Printer.ScaleMode
  
  fPageBreak = False

  ' Initialise an array of the grids that are too wide for the current page.
  ' Column 1 = grid index
  ' Column 2 = first column of the grid being carried over onto the overflow page
  ' Column 3 = number of pages the first column has already been printed on (in case it too wide for a single page)
  ReDim alngOverflows(3, 0)

  For Each ctlTemp In pfrmRecProfile.Controls
    If ctlTemp.Container Is pfrmRecProfile.picOutput(piPageNumber) Then
      '
      ' LABEL control
      '
      If TypeOf ctlTemp Is Label Then
        If ctlTemp.Visible Then
          ' Calculate the print coordinates of the label
          rctDraw.Left = ctlTemp.Left + lngMarginLeft_Twips
          rctDraw.Right = rctDraw.Left + ctlTemp.Width
          rctDraw.Top = ctlTemp.Top + mlngRecordProfileTopOffset - lngPageTopPictureBoxPosition
          rctDraw.Bottom = rctDraw.Top + ctlTemp.Height
          
          ' If we are page-breaking after the previous grid output,
          ' or this label does not fit on the current page,
          ' then get the printer to print it on a new page.
          fNewPage = fPageBreak Or (rctDraw.Bottom > mlngRecordProfileBottomOffset)
          
          If (Not fNewPage) And _
            (Not pfrmRecProfile.IsEmptyTableCaption(ctlTemp.Index)) Then
            ' Get the printer to print on a new page if the label DOES fit on the current page,
            ' BUT the grid that follows the label DOES NOT have any data rows that fit on it.
            
            ' Get the grid that follows this label.
            lngTemp = pfrmRecProfile.AssociatedControlIndex(False, ctlTemp.Index)
            
            If lngTemp > 0 Then
              Set ctlTempGrid = pfrmRecProfile.grdOutput(lngTemp)
            
              ' Get the bottom coordinate of the first data row in the grid,
              ' to see if it fits in the current page.
              lngGroupColumnHeadingHeight = ctlTempGrid.Height - (ctlTempGrid.RowHeight * ctlTempGrid.Rows)
              If ((ctlTempGrid.Groups.Count > 0) And (ctlTempGrid.GroupHeaders)) And _
                ctlTempGrid.ColumnHeaders Then
                lngGroupColumnHeadingHeight = ctlTempGrid.Columns(0).Top
              End If
              
              rctDraw2.Top = ctlTempGrid.Top + mlngRecordProfileTopOffset - lngPageTopPictureBoxPosition
              If (ctlTempGrid.Groups.Count > 0) And (ctlTempGrid.GroupHeaders) Then
                rctDraw2.Top = rctDraw2.Top + lngGroupColumnHeadingHeight
              End If
              If ctlTempGrid.ColumnHeaders Then
                rctDraw2.Top = rctDraw2.Top + lngGroupColumnHeadingHeight
              End If
              rctDraw2.Bottom = rctDraw2.Top + ctlTempGrid.RowHeight

              If (lngPageTopPictureBoxPosition = 0) And (rctDraw2.Bottom > mlngRecordProfileBottomOffset) Then
                rctDraw2.Bottom = mlngRecordProfileBottomOffset
              End If
             
              fNewPage = (rctDraw2.Bottom > mlngRecordProfileBottomOffset)
            
              Set ctlTempGrid = Nothing
            End If
          End If
          
          If fNewPage Then
            ' Before starting to print on a new page, print any overflow pages
            ' (ie. any pages for grid columns that did not fit on the current page).
            If UBound(alngOverflows, 2) > 0 Then
              If Not PrintRecordProfileOverflowPage(pfrmRecProfile, _
                alngOverflows, _
                sTitle, _
                sLeftFooter, _
                lngMarginLeft_Twips, _
                lngMarginRight_Twips, _
                lngMarginTop_Twips, _
                lngMarginBottom_Twips, _
                lngPageTopPictureBoxPosition) Then
                
                fOK = False
                GoTo SubErrorTrap
              End If

              ReDim alngOverflows(3, 0)
            End If
            
            ' Get the printer to start printing on a new page.
            miRecordProfilePageCounter = miRecordProfilePageCounter + 1
            Printer.NewPage
            
            ' Remember the picturebox position that is printed at the top of the page
            ' (all other vertical position is offset from this).
            lngPageTopPictureBoxPosition = ctlTemp.Top
            
            ' Print the new page header and footers.
            sRightFooter = "Page " & CStr(miRecordProfilePageCounter)
            If Not PrintRecordProfilePageHeaderAndFooter(sTitle, _
              sLeftFooter, _
              sRightFooter, _
              lngMarginLeft_Twips, _
              lngMarginRight_Twips, _
              lngMarginTop_Twips, _
              lngMarginBottom_Twips) Then
              
              fOK = False
              GoTo SubErrorTrap
            End If
          
            ' Re-calculate the vertical print coordinates of the label,
            ' as they will have changed as we created the new page.
            rctDraw.Top = ctlTemp.Top + mlngRecordProfileTopOffset - lngPageTopPictureBoxPosition
            rctDraw.Bottom = rctDraw.Top + ctlTemp.Height
          End If
          
          ' Check if we need to page break after this table's output.
          Set objRecProfTable = pfrmRecProfile.Definition.Item(ctlTemp.Tag)
          fPageBreak = objRecProfTable.PageBreak
          Set objRecProfTable = Nothing
          
          ' Scale the print coordinates for use with the DrawText function.
          rctDraw.Left = Printer.ScaleY(rctDraw.Left, iSM, vbPixels)
          rctDraw.Right = Printer.ScaleY(rctDraw.Right, iSM, vbPixels)
          rctDraw.Top = Printer.ScaleY(rctDraw.Top, iSM, vbPixels)
          rctDraw.Bottom = Printer.ScaleY(rctDraw.Bottom, iSM, vbPixels)
          
          'JPD 20050609 Fault 10060
          ' Printer errors caused when assigning the Printer.Font object with another Font object.
          ' Found it better to simply assign values to the individual properties.
          
          ' Set the printer's font to match the label's font.
          'Set Printer.Font = ctlTemp.Font
          
          Printer.Font.Name = ctlTemp.Font.Name
          Printer.Font.Size = ctlTemp.Font.Size
          Printer.Font.Bold = ctlTemp.Font.Bold
          Printer.Font.Underline = ctlTemp.Font.Underline
          Printer.Font.Strikethrough = ctlTemp.Font.Strikethrough
          Printer.Font.Italic = ctlTemp.Font.Italic
          
          ' Command the printer to print the label's caption.
          DrawText Printer.hDC, ctlTemp.Caption, -1, rctDraw, c_DTDefFmt Or DT_LEFT
        End If
      End If
        
      '
      ' GRID control
      '
      If TypeOf ctlTemp Is SSDBGrid Then
        ' Check if we need to page break after this table's output.
        Set objRecProfTable = pfrmRecProfile.Definition.Item(ctlTemp.Tag)
        fPageBreak = objRecProfTable.PageBreak
        
        ' Check if this grid has Heading columns,
        ' and get the indexes of the first and last visible columns.
        fHasHeadingColumn = False
        iFirstVisibleColumn = -1
        iLastVisibleColumn = 0
        For iLoop2 = 0 To ctlTemp.Columns.Count - 1
          If ctlTemp.Columns(iLoop2).Name = COLUMN_ISHEADING Then
            fHasHeadingColumn = True
          End If
          If ctlTemp.Columns(iLoop2).Visible Then
            iLastVisibleColumn = iLoop2
            
            If iFirstVisibleColumn < 0 Then
              iFirstVisibleColumn = iLoop2
            End If
          End If
        Next iLoop2
        
        ' Check if this grid is preceded or followed IMMEDIATELY by other grids.
        ' ie. if this grid is part of a group of grids that are used to display
        ' data (including pictures) vertically.
        ' NB. Grids have only one row height value. To display pictures with their
        ' own row height, we actually put them in their own grid, and position this grid
        ' IMMEDIATELY after the normal data grid. Subsequent data is put in its own grid
        ' IMMEDIATELY after the picture's grid.
        ' This is waht's meant by 'following' and 'preceding' grids.
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
        
        ' Determine the height of the Group and Column Heading rows,
        ' as this is not the same as the grid's rowHeight.
        lngGroupColumnHeadingHeight = ctlTemp.Height - (ctlTemp.RowHeight * ctlTemp.Rows)
        If ((ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders)) And _
          ctlTemp.ColumnHeaders Then
          lngGroupColumnHeadingHeight = ctlTemp.Columns(0).Top
        End If
        
        ' Draw group and column headers if required.
        If Not PrintRecordProfileColumnHeadings(ctlTemp, _
          lngMarginLeft_Twips, _
          lngPageTopPictureBoxPosition, _
          lngGroupColumnHeadingHeight, _
          alngOverflows, _
          0, 0) Then
          
          fOK = False
          GoTo SubErrorTrap
        End If
        
        ' Draw data rows and columns.
        For iLoop = 0 To ctlTemp.Rows - 1
          ' Check if the next row will go beyond the bottom of the page.
          rctDraw.Top = ctlTemp.Top + (iLoop * ctlTemp.RowHeight) + mlngRecordProfileTopOffset - lngPageTopPictureBoxPosition
          If (ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders) Then
            rctDraw.Top = rctDraw.Top + lngGroupColumnHeadingHeight
          End If
          If ctlTemp.ColumnHeaders Then
            rctDraw.Top = rctDraw.Top + lngGroupColumnHeadingHeight
          End If
          rctDraw.Bottom = rctDraw.Top + ctlTemp.RowHeight

          If (Not fCellsPrinted) And (rctDraw.Bottom > mlngRecordProfileBottomOffset) Then
            rctDraw.Bottom = mlngRecordProfileBottomOffset
          End If

          fNewPage = (rctDraw.Bottom > mlngRecordProfileBottomOffset)
          If fNewPage Then
            ' Make sure the last row on a page has it's bottom line in bold.
            If mblnGridlines Then
              If iLoop > 0 Then
                For iLoop2 = 0 To ctlTemp.Columns.Count - 1
                  If ctlTemp.Columns(iLoop2).Visible Then
                    rctDraw.Left = ctlTemp.Left + ctlTemp.Columns(iLoop2).Left + lngMarginLeft_Twips
                    rctDraw.Right = rctDraw.Left + ctlTemp.Columns(iLoop2).Width
                    rctDraw.Top = ctlTemp.Top + ((iLoop - 1) * ctlTemp.RowHeight) + mlngRecordProfileTopOffset - lngPageTopPictureBoxPosition
                    If (ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders) Then
                      rctDraw.Top = rctDraw.Top + lngGroupColumnHeadingHeight
                    End If
                    If ctlTemp.ColumnHeaders Then
                      rctDraw.Top = rctDraw.Top + lngGroupColumnHeadingHeight
                    End If
                    rctDraw.Bottom = rctDraw.Top + ctlTemp.RowHeight

                    rctDraw2.Left = rctDraw.Left
                    rctDraw2.Right = rctDraw.Right
                    rctDraw2.Top = rctDraw.Bottom
                    rctDraw2.Bottom = rctDraw.Bottom

                    If rctDraw2.Right > Printer.ScaleWidth Then
                      If iLoop2 > 0 Then
                        Exit For
                      Else
                        rctDraw2.Right = Printer.ScaleWidth
                      End If
                    End If

                    Printer.DrawWidth = THICKLINE
                    Printer.Line (rctDraw2.Left, rctDraw2.Top)-(rctDraw2.Right, rctDraw2.Bottom), , B
                    Printer.DrawWidth = THINLINE
                  End If
                Next iLoop2
              End If
            End If

            ' Before starting to print on a new page, print any overflow pages
            ' (ie. any pages for grid columns that did not fit on the current page).
            If UBound(alngOverflows, 2) > 0 Then
              If Not PrintRecordProfileOverflowPage(pfrmRecProfile, _
                alngOverflows, _
                sTitle, _
                sLeftFooter, _
                lngMarginLeft_Twips, _
                lngMarginRight_Twips, _
                lngMarginTop_Twips, _
                lngMarginBottom_Twips, _
                lngPageTopPictureBoxPosition) Then
                
                fOK = False
                GoTo SubErrorTrap
              End If

              ReDim alngOverflows(3, 0)
            End If

            miRecordProfilePageCounter = miRecordProfilePageCounter + 1
            Printer.NewPage

            lngPageTopPictureBoxPosition = ctlTemp.Top + (iLoop * ctlTemp.RowHeight)

            sRightFooter = "Page " & CStr(miRecordProfilePageCounter)
            If Not PrintRecordProfilePageHeaderAndFooter(sTitle, _
              sLeftFooter, _
              sRightFooter, _
              lngMarginLeft_Twips, _
              lngMarginRight_Twips, _
              lngMarginTop_Twips, _
              lngMarginBottom_Twips) Then
              
              fOK = False
              GoTo SubErrorTrap
            End If

            ' Print the group & column headings again if required.
            If Not PrintRecordProfileColumnHeadings(ctlTemp, _
              lngMarginLeft_Twips, _
              lngPageTopPictureBoxPosition - (iLoop * ctlTemp.RowHeight), _
              lngGroupColumnHeadingHeight, _
              alngOverflows, _
              0, 0) Then
              
              fOK = False
              GoTo SubErrorTrap
            End If
          End If

          For iLoop2 = 0 To ctlTemp.Columns.Count - 1
            If ctlTemp.Columns(iLoop2).Visible Then
              rctDraw.Left = ctlTemp.Left + ctlTemp.Columns(iLoop2).Left + lngMarginLeft_Twips
              rctDraw.Right = rctDraw.Left + ctlTemp.Columns(iLoop2).Width
              rctDraw.Top = ctlTemp.Top + (iLoop * ctlTemp.RowHeight) + mlngRecordProfileTopOffset - lngPageTopPictureBoxPosition
              If (ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders) Then
                rctDraw.Top = rctDraw.Top + lngGroupColumnHeadingHeight
              End If
              If ctlTemp.ColumnHeaders Then
                rctDraw.Top = rctDraw.Top + lngGroupColumnHeadingHeight
              End If
              rctDraw.Bottom = rctDraw.Top + ctlTemp.RowHeight

              If rctDraw.Bottom > mlngRecordProfileBottomOffset Then
                rctDraw.Bottom = mlngRecordProfileBottomOffset
              End If
              
              If rctDraw.Right > Printer.ScaleWidth Then
                If iLoop2 > 0 Then
                  If mblnGridlines Then
                    ' Ensure the last printed data cell has a bold line on the right-hand side.
                    rctDraw2.Left = rctDraw.Left
                    rctDraw2.Right = rctDraw.Left
                    rctDraw2.Top = rctDraw.Top
                    rctDraw2.Bottom = rctDraw.Bottom
                    Printer.DrawWidth = THICKLINE
                    Printer.Line (rctDraw2.Left, rctDraw2.Top)-(rctDraw2.Right, rctDraw2.Bottom), , B
                    Printer.DrawWidth = THINLINE
                  End If
  
                  Exit For
                Else
                  rctDraw.Right = Printer.ScaleWidth
                End If
              End If

              If (ctlTemp.Columns(iLoop2).Style = 4) And _
                (Not ctlTemp.ColumnHeaders) Then
                ' Button
                Set fntFont = ctlTemp.Font
                varBookmark = ctlTemp.AddItemBookmark(iLoop)
                If fHasHeadingColumn Then
                  If ctlTemp.Columns(COLUMN_ISHEADING).CellText(varBookmark) = "1" Then
                    Set fntFont = ctlTemp.StyleSets("Heading").Font
                  End If
                End If

                fCellsPrinted = True
                
                If Not PrintRecordProfileGridCell(rctDraw, _
                  fntFont, _
                  ctlTemp.Columns(iLoop2).CellText(varBookmark), _
                  cellCOLUMNHEADING, _
                  ctlTemp, _
                  -1, _
                  0, _
                  False) Then
                  
                  fOK = False
                  GoTo SubErrorTrap
                End If
              
                Set fntFont = Nothing
              Else
                ' Normal data cell
                iCellType = cellDATA

                ' DRAW THE BOX
                If fHasHeadingColumn Then
                  varBookmark = ctlTemp.AddItemBookmark(iLoop)
                  If ctlTemp.Columns(COLUMN_ISHEADING).CellText(varBookmark) = "1" Then
                    iCellType = cellSEPARATOR
                  End If
                End If

                If ctlTemp.Columns(iLoop2).StyleSet = "Separator" Then
                  iCellType = cellSEPARATOR
                End If

                lngPictureIndex = -1
                fPhotoDone = False
                If ctlTemp.TagVariant = COLUMN_ISPHOTO Then
                  For iLoop3 = 0 To ctlTemp.Columns.Count - 1
                    If ctlTemp.Columns(iLoop3).Visible Then
                      sTemp = PHOTOSTYLESET & CStr(iLoop3 + 1)

                      For iLoop4 = 0 To ctlTemp.StyleSets.Count - 1
                        If ctlTemp.StyleSets(iLoop4).Name = sTemp Then
                          lngPictureIndex = iLoop4
                          fPhotoDone = True
                          Exit For
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
                      lngPictureIndex = iLoop4
                      fPhotoDone = True
                      Exit For
                    End If
                  Next iLoop4
                End If

                Set fntFont = ctlTemp.Font
                varBookmark = ctlTemp.AddItemBookmark(iLoop)
                If fHasHeadingColumn Then
                  If ctlTemp.Columns(COLUMN_ISHEADING).CellText(varBookmark) = "1" Then
                    Set fntFont = ctlTemp.StyleSets("Heading").Font
                  End If
                End If
                
                fCellsPrinted = True

                If Not PrintRecordProfileGridCell(rctDraw, _
                  fntFont, _
                  ctlTemp.Columns(iLoop2).CellText(varBookmark), _
                  iCellType, _
                  ctlTemp, _
                  lngPictureIndex, _
                  0, _
                  False) Then
                  
                  fOK = False
                  GoTo SubErrorTrap
                End If

                Set fntFont = Nothing

                If mblnGridlines Then
                  ' Top rows of the grid must have the top line in bold.
                  If ((iLoop = 0) And (Not fGridPreceded)) Or _
                    fNewPage Then
                    rctDraw2.Left = rctDraw.Left
                    rctDraw2.Right = rctDraw.Right
                    rctDraw2.Top = rctDraw.Top
                    rctDraw2.Bottom = rctDraw.Top
                    Printer.DrawWidth = THICKLINE
                    Printer.Line (rctDraw2.Left, rctDraw2.Top)-(rctDraw2.Right, rctDraw2.Bottom), , B
                    Printer.DrawWidth = THINLINE
                  End If

                  ' Last columns on the page must have the right line in bold.
                  If iLoop2 = iLastVisibleColumn Then
                    rctDraw2.Left = rctDraw.Right
                    rctDraw2.Right = rctDraw.Right
                    rctDraw2.Top = rctDraw.Top
                    rctDraw2.Bottom = rctDraw.Bottom
                    Printer.DrawWidth = THICKLINE
                    Printer.Line (rctDraw2.Left, rctDraw2.Top)-(rctDraw2.Right, rctDraw2.Bottom), , B
                    Printer.DrawWidth = THINLINE
                  End If

                  ' First columns on the page must have the left line in bold.
                  If iLoop2 = iFirstVisibleColumn Then
                    rctDraw2.Left = rctDraw.Left
                    rctDraw2.Right = rctDraw.Left
                    rctDraw2.Top = rctDraw.Top
                    rctDraw2.Bottom = rctDraw.Bottom
                    Printer.DrawWidth = THICKLINE
                    Printer.Line (rctDraw2.Left, rctDraw2.Top)-(rctDraw2.Right, rctDraw2.Bottom), , B
                    Printer.DrawWidth = THINLINE
                  End If

                  ' Bottom rows of the grid must have the bottom line in bold.
                  If (iLoop = (ctlTemp.Rows - 1)) And (Not fGridFollowed) Then
                    rctDraw2.Left = rctDraw.Left
                    rctDraw2.Right = rctDraw.Right
                    rctDraw2.Top = rctDraw.Bottom
                    rctDraw2.Bottom = rctDraw.Bottom
                    Printer.DrawWidth = THICKLINE
                    Printer.Line (rctDraw2.Left, rctDraw2.Top)-(rctDraw2.Right, rctDraw2.Bottom), , B
                    Printer.DrawWidth = THINLINE
                  End If
                End If
              End If
            End If
          Next iLoop2
        Next iLoop
      
        Set objRecProfTable = Nothing
      End If
    End If
  Next ctlTemp
  Set ctlTemp = Nothing

  ' Print any columns that did not fit on this page.
  If UBound(alngOverflows, 2) > 0 Then
    If Not PrintRecordProfileOverflowPage(pfrmRecProfile, _
      alngOverflows, _
      sTitle, _
      sLeftFooter, _
      lngMarginLeft_Twips, _
      lngMarginRight_Twips, _
      lngMarginTop_Twips, _
      lngMarginBottom_Twips, _
      lngPageTopPictureBoxPosition) Then
      
      fOK = False
      GoTo SubErrorTrap
    End If

    ReDim alngOverflows(3, 0)
  End If

  If (piPageNumber = (pfrmRecProfile.picOutput.Count - 1)) Then
    Printer.EndDoc
    PrintConfirm
    Unload frmPrintOptions
    Set frmPrintOptions = Nothing
  End If


TidyUpAndExit:
  Printer.ScaleMode = iOldScaleMode

SubErrorTrap:
  If fOK Then
    gobjErrorStack.PopStack
  End If
  
  PrintRecordProfilePage = fOK
  Exit Function

ErrorTrap:
  If Err.Number <> 0 Then
    msErrorMessage = Err.Description
  End If
  
  gobjErrorStack.HandleError
  PrintRecordProfilePage = False
  
End Function


Public Sub PrintInitialise(ssPrintInfo As SSDataWidgets_B.ssPrintInfo)

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "clsPrintGrid.PrintInitialise()"

  With ssPrintInfo
    'Underline headings if not printing gridlines
    .PrintGridlines = IIf(mblnGridlines, ssPrintGridLinesAll, ssPrintGridLinesNone)
    
    .PrintHeaders = IIf(mblnHeadingsOnEveryPage, ssTopOfPage, ssTopOfReport)
    .Portrait = mblnPortrait
    .Copies = mintNumberOfCopies
    .Collate = mblnCollate
    .PrintColors = mblnShading

    .RowAutoSize = True
    'NHRD20030425 Fault 2470 Modified first line from  .MaxLinesPerRow = 2 to  .MaxLinesPerRow = 3
    'Added the clipping override property too.  Hopefully shouldn't mess up any other subs that use this Sub
    .MaxLinesPerRow = 3
    .ClippingOverride = ssClippingOverrideAuto
    
    .PrintColumnHeaders = ssUseCaption
  

    .PageHeader = vbTab & mstrHeading & vbTab
    .PageFooter = "Printed on <date> at <time> by " & gsUserName & vbTab & vbTab & "Page <page number>"

    ' 15/02/2000 RH - Add ability to specify margins
    .MarginTop = mintMarginTop
    .MarginBottom = mintMarginBottom
    .MarginLeft = mintMarginLeft
    .MarginRight = mintMarginRight
    
  End With

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub

Public Sub PrintConfirm()

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "clsPrintGrid.PrintConfirm()"

' Display a printing complete prompt

  Dim strMBText As String
  Dim msb As frmMessageBox
  Dim iShowMeAgain As Integer
  
  If gbPrinterConfirm And Not gblnBatchMode Then
    
    'MH20040127 Fault 7941
    'strMBText = mstrHeading & " printing complete." _
      & vbCrLf & vbCrLf & "(" & Printer.DeviceName & ")"
    strMBText = mstrHeading & " printing complete." _
      & vbCrLf & vbCrLf & "(" & mstrSelectedPrinter & ")"
    
  
    iShowMeAgain = IIf(gbPrinterConfirm, 1, 0)
    If iShowMeAgain = 1 Then
      Set msb = New frmMessageBox
      'TM20020829 Fault 4293 - the checkbox should not be checked.
      iShowMeAgain = 0
      msb.MessageBox strMBText, vbInformation, App.ProductName, iShowMeAgain, "Don't show me this confirmation again."
      gbPrinterConfirm = IIf(iShowMeAgain = 0, True, False)
      SavePCSetting "Printer", "Confirm", gbPrinterConfirm
      Set msb = Nothing
    End If
  
  End If

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub

Private Sub Class_Terminate()
  Set mfntRecordProfileHeaderFont = Nothing
  Set mfntRecordProfileFooterFont = Nothing

End Sub



