VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOutputWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mwrdApp As Word.Application
Private mwrdDoc As Word.Document
Private mwrdTable As Word.Table
Private mobjParent As clsOutputRun

Private mblnScreen As Boolean
Private mblnPrinter As Boolean
Private mstrPrinterName As String
Private mblnSave As Boolean
Private mlngSaveExisting As Long
Private mblnEmail As Boolean
Private mstrFileName As String
'Private mlngSaveFileFormat As Long
'Private mlngEmailFileFormat As Long

Private mstrErrorMessage As String
Private mlngPageCount As Long
Private mlngHeaderRows As Long
Private mlngHeaderCols As Long
Private mblnApplyStyles As Boolean

Private mstrWrdTemplate As String
Private mblnWrdWordGridlines As Boolean
Private mblnWrdAutoFitCols As Boolean
Private mblnWrdLandscape As Boolean

Public Sub ClearUp()
  On Local Error Resume Next
  Set mwrdTable = Nothing
  Set mwrdDoc = Nothing
  If Not mwrdApp Is Nothing Then mwrdApp.Quit False
  Set mwrdApp = Nothing
End Sub

Public Function RecordProfilePage(pfrmRecProfile As Form, _
  piPageNumber As Integer, _
  pcolStyles As Collection)
  ' Output the record profile page to Word.
  
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim fFirstColumnDone As Boolean
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
  Dim fPageBreak As Boolean
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
  Dim lngPrecedingRows As Long
  Dim lngTabIndex As Long         'MH20040625 Fault 8782
  Dim alngControlOrder() As Long
  Dim fPageDone As Boolean
  Dim sTableName As String

  Const strBookMark As String = "ASRSysTableStart"
  
  Const RECPROFFOLLOWONCORRECTION = 10

  Const COLUMN_ISHEADING = "IsHeading"
  Const COLUMN_ISPHOTO = "IsPhoto"
  Const PHOTOSTYLESET = "PhotoSS_"

  fOK = True
  sTitle = pfrmRecProfile.Caption
  iOriginalStyles = pcolStyles.Count
  fPageBreak = False

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

        ' If we are page-breaking after the previous grid output
        ' then get Word to create a new page.
        If fPageBreak Then
          ' Get Word to create a new page.
          AddPage sTitle, "", pcolStyles
        End If

        ' Check if we need to page break after this table's output.
        Set objRecProfTable = pfrmRecProfile.Definition.Item(ctlTemp.Tag)
        fPageBreak = objRecProfTable.PageBreak
        Set objRecProfTable = Nothing

        ' Write the label's caption to Word.
        With mwrdDoc.ActiveWindow.Selection
          mwrdApp.Selection.Font.Bold = pcolStyles("Title").Bold
          mwrdApp.Selection.Font.Underline = IIf(pcolStyles("Title").Underline, wdUnderlineSingle, wdUnderlineNone)
          .TypeText ctlTemp.Caption
          mwrdApp.Selection.Font.Bold = False
          mwrdApp.Selection.Font.Underline = wdUnderlineNone
          .TypeParagraph
          .TypeParagraph
        End With
      Else
        Set ctlTemp = pfrmRecProfile.grdOutput(alngControlOrder(3, lngTabIndex))
        '
        ' GRID control
        '
        ' Check if we need to page break after this table's output.
        Set objRecProfTable = pfrmRecProfile.Definition.Item(ctlTemp.Tag)
        fPageBreak = objRecProfTable.PageBreak
        sTableName = objRecProfTable.TableName
        
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

        If Not fGridPreceded Then
          lngPrecedingRows = 0
          mwrdDoc.Bookmarks.Add strBookMark
          Set colMerges = New Collection
        End If

        ' Send the column/group headers to the word document.
        fHasHeadingColumn = False
        sTemp = ""
        iLastGroup = -1
        iTemp = 0
        iGroupStart = 0
        fFirstColumnDone = False

        For iLoop = 0 To ctlTemp.Columns.Count - 1
          If ctlTemp.Columns(iLoop).Name = COLUMN_ISHEADING Then
            fHasHeadingColumn = True
          End If

          If (ctlTemp.ColumnHeaders) And (ctlTemp.Columns(iLoop).Visible) Then
            If (ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders) Then
              If (iLastGroup <> ctlTemp.Columns(iLoop).Group) Then
                ' Send the group headers to the word document.
                mwrdDoc.ActiveWindow.Selection.TypeText IIf(fFirstColumnDone, vbTab, "") & ctlTemp.Groups(ctlTemp.Columns(iLoop).Group).Caption

                ' Remember if the group heading cells need to be merged.
                If iGroupStart < iTemp - 1 Then
                  Set objMerge = New clsOutputStyle
                  objMerge.StartCol = iGroupStart
                  objMerge.StartRow = 0
                  objMerge.EndCol = iTemp - 1
                  objMerge.EndRow = 0

                  colMerges.Add objMerge
                  Set objMerge = Nothing
                End If

                iGroupStart = iTemp
              Else
                mwrdDoc.ActiveWindow.Selection.TypeText IIf(fFirstColumnDone, vbTab, "")
              End If

              iLastGroup = ctlTemp.Columns(iLoop).Group
            End If

            ' Check if the column is a Separator column.
            ' If so we'll need to create a 'style' object for it.
            If ctlTemp.Columns(iLoop).StyleSet = "Separator" Then
              Set objTemp = New clsOutputStyle

              With objTemp
                .StartCol = iTemp
                .StartRow = IIf((ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders), 2, 1)
                .EndCol = CLng(iTemp)
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

            ' Send the column/group headers to the word document.
            sTemp = sTemp & IIf(fFirstColumnDone, vbTab, "") & ctlTemp.Columns(iLoop).Caption
            fFirstColumnDone = True
            iTemp = iTemp + 1
          End If
        Next iLoop

        ' Remember if the group heading cells need to be merged.
        If (ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders) Then
          If iGroupStart < iTemp - 1 Then
            Set objMerge = New clsOutputStyle
            objMerge.StartCol = iGroupStart
            objMerge.StartRow = 0
            objMerge.EndCol = iTemp - 1
            objMerge.EndRow = 0

            colMerges.Add objMerge
            Set objMerge = Nothing
          End If

          mwrdDoc.ActiveWindow.Selection.TypeParagraph
        End If

        If Len(sTemp) > 0 Then
          ' Send the column headers to the word document.
          mwrdDoc.ActiveWindow.Selection.TypeText sTemp
          mwrdDoc.ActiveWindow.Selection.TypeParagraph
        End If

        ' Send data rows and columns to Word.
        For iLoop = 0 To ctlTemp.Rows - 1

          fFirstColumnDone = False
          varBookmark = ctlTemp.AddItemBookmark(iLoop)

          ' Check if the row is a Separator/Heading row.
          ' If so we'll need to create a 'style' object for it.
          If fHasHeadingColumn Then
            If ctlTemp.Columns(COLUMN_ISHEADING).CellText(varBookmark) = "1" Then

              Set objTemp = New clsOutputStyle

              With objTemp
                .StartCol = IIf(ctlTemp.ColumnHeaders, 0, 1)
                .StartRow = lngPrecedingRows + iLoop + _
                  IIf(ctlTemp.ColumnHeaders, 1, 0) + _
                  IIf((ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders), 1, 0)
                .EndCol = -1
                .EndRow = .StartRow

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

          For iLoop2 = 0 To ctlTemp.Columns.Count - 1
            If ctlTemp.Columns(iLoop2).Visible Then
              ' Send the text or picture to Word.

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
                          mwrdDoc.ActiveWindow.Selection.TypeText IIf(fFirstColumnDone, vbTab, "")
                          mwrdDoc.ActiveWindow.Selection.InlineShapes.AddPicture FileName:=sTempName, LinkToFile:=False, SaveWithDocument:=True
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
                      mwrdDoc.ActiveWindow.Selection.TypeText IIf(fFirstColumnDone, vbTab, "")
                      mwrdDoc.ActiveWindow.Selection.InlineShapes.AddPicture FileName:=sTempName, LinkToFile:=False, SaveWithDocument:=True
                      Kill sTempName
                      fPhotoDone = True

                      Exit For
                    End If
                  End If
                Next iLoop4

              End If

              If Not fPhotoDone Then

                    
                varBookmark = ctlTemp.AddItemBookmark(iLoop)

                mwrdDoc.ActiveWindow.Selection.TypeText IIf(fFirstColumnDone, vbTab, "") & ctlTemp.Columns(iLoop2).CellText(varBookmark)

              End If

              fFirstColumnDone = True

            End If
          Next iLoop2

          mwrdDoc.ActiveWindow.Selection.TypeParagraph
        
        Next iLoop

        Set objRecProfTable = Nothing

        If fGridFollowed Then

          lngPrecedingRows = lngPrecedingRows + _
            IIf(ctlTemp.ColumnHeaders, 1, 0) + _
            IIf((ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders), 1, 0) + _
            ctlTemp.Rows

        Else

          ' Convert the text into a table.
          mwrdDoc.ActiveWindow.Selection.GoTo What:=wdGoToBookmark, Name:=strBookMark
          mwrdDoc.ActiveWindow.Selection.EndKey Unit:=wdStory, Extend:=wdExtend

          Set mwrdTable = mwrdDoc.ActiveWindow.Selection.ConvertToTable _
            (Separator:=wdSeparateByTabs, _
            NumRows:=ctlTemp.Rows + IIf(ctlTemp.ColumnHeaders, 1, 0) + IIf((ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders), 1, 0), _
            Format:=wdTableFormatNone, _
            ApplyFont:=False, _
            ApplyColor:=False, _
            AutoFit:=False)

          With mwrdTable
            lngBorder = IIf(mblnWrdWordGridlines, wdLineStyleSingle, wdLineStyleNone)
            .Borders(wdBorderLeft).LineStyle = lngBorder
            .Borders(wdBorderRight).LineStyle = lngBorder
            .Borders(wdBorderTop).LineStyle = lngBorder
            .Borders(wdBorderBottom).LineStyle = lngBorder
            .Borders(wdBorderVertical).LineStyle = lngBorder
            .Borders(wdBorderHorizontal).LineStyle = lngBorder
          End With

          ' Apply styles to the table.
          With pcolStyles("Heading")
            .StartCol = 0
            .StartRow = 0
            .EndCol = IIf(ctlTemp.ColumnHeaders, mwrdTable.Columns.Count, 0) - 1
            .EndRow = IIf(ctlTemp.ColumnHeaders, 1, 0) - 1 + IIf((ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders), 1, 0)
          End With

          With pcolStyles("HeadingCols")
            .StartCol = 0
            .StartRow = 0
            .EndCol = IIf(ctlTemp.ColumnHeaders, 0, 1) - 1
            .EndRow = IIf(ctlTemp.ColumnHeaders, -1, mwrdTable.Rows.Count)
          End With

          With pcolStyles("Data")
            .StartCol = IIf(ctlTemp.ColumnHeaders, 0, 1)
            .StartRow = IIf(ctlTemp.ColumnHeaders, 1, 0) + IIf((ctlTemp.Groups.Count > 0) And (ctlTemp.GroupHeaders), 1, 0)
            .EndCol = mwrdTable.Columns.Count - 1
            .EndRow = mwrdTable.Rows.Count - 1
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
              objTemp.EndRow = mwrdTable.Rows.Count - 1
            End If
            If Left(objTemp.Name, 11) = "RECPROFROW_" Then
              objTemp.EndCol = mwrdTable.Columns.Count - 1
            End If
          Next objTemp
          Set objTemp = Nothing

          ' Apply styles and merging to the table.
          If mblnApplyStyles Then
            ApplyStyle pcolStyles
            ApplyMerges colMerges
          End If
          Set colMerges = Nothing

          ' Get rid of any separator/heading styles that have been created
          ' for this table, so that they are not carried over and applied
          ' to other tables.
          Do While pcolStyles.Count > iOriginalStyles
            pcolStyles.Remove pcolStyles.Count
          Loop

          mwrdApp.Selection.EndKey Unit:=wdStory

          mwrdDoc.ActiveWindow.Selection.TypeParagraph

        End If

      End If
      
      Set ctlTemp = Nothing
    Else
      If fPageDone Then
        Exit For
      End If
    End If
  Next lngTabIndex

TidyUpAndExit:
  RecordProfilePage = fOK
  Exit Function
  
ErrorTrap:
  If Err.Number = 4608 Then
    mstrErrorMessage = "Cannot output more than 63 columns to Word." & vbCrLf & _
      "Try changing the Data Orientation of the '" & Replace(sTableName, "_", " ") & "' table to be 'Vertical'."
  Else
    mstrErrorMessage = Err.Description
  End If
  
  fOK = False
  Resume TidyUpAndExit

End Function


Private Sub Class_Initialize()
  mstrWrdTemplate = GetUserSetting("Output", "WordTemplate", "")
  mblnWrdAutoFitCols = (GetUserSetting("Output", "AutoFitCols", "1") = "1")
  mblnWrdLandscape = (GetUserSetting("Output", "Landscape", "1") = "1")
  'mblnWrdWordGridlines = (GetUserSetting("Output", "Gridlines", "1") = "1")
  mblnWrdWordGridlines = (GetUserSetting("Output", "DataGridLines", "1") = "1")
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


Private Function CreateWordApplication() As Boolean

  On Local Error GoTo LocalErr

  Set mwrdApp = CreateObject("Word.Application")
  
  If GetSystemSetting("Output", "WordVisible", "0") = "1" Then
    mwrdApp.Visible = True
    mwrdApp.DisplayAlerts = wdAlertsAll
  Else
    mwrdApp.DisplayAlerts = wdAlertsNone
  End If
  
  CreateWordApplication = True

Exit Function

LocalErr:
  mstrErrorMessage = "Error opening Word Application"
  CreateWordApplication = False

End Function


Public Function GetFile(objParent As clsOutputRun, colStyles As Collection) As Boolean

  Dim strTempFileName As String
  Dim lngFound As Long
  Dim lngCount As Long

  On Local Error GoTo LocalErr

  If Not CreateWordApplication Then
    GetFile = False
    Exit Function
  End If

  'Just in case we are emailing but not saving...
  If mblnEmail And Not mblnSave Then
    mstrFileName = objParent.GetTempFileName(mstrFileName)
  End If
  
  
  'Check if file already exists...
  If Dir(mstrFileName) <> vbNullString And mstrFileName <> vbNullString Then

    Select Case mlngSaveExisting
    Case 0  'Overwrite
      If Not objParent.KillFile(mstrFileName) Then
        GetFile = False
        Exit Function
      End If
      GetNewDocument

    Case 1  'Do not overwrite (fail)
      mwrdApp.Quit
      Set mwrdApp = Nothing
      mstrErrorMessage = "File already exists."

    Case 2  'Add Sequential number to file
      mstrFileName = mobjParent.GetSequentialNumberedFile(mstrFileName)
      GetNewDocument

    Case 3  'Append to existing file
      'Start at the bottom of the document

      If Not IsFileCompatibleWithWordVersion(mstrFileName) Then
        mstrErrorMessage = "This definition is set to append to a file which is not compatible with your version of Microsoft Office."
        GetFile = False
        Exit Function
      End If

      Set mwrdDoc = mwrdApp.Documents.Open(mstrFileName)
      mwrdApp.Selection.EndKey Unit:=wdStory
      mwrdApp.Selection.InsertBreak Type:=wdPageBreak

    Case 4  'Create new worksheet within existing workbook...
      'N/A (EXCEL ONLY)

    End Select

  Else
    GetNewDocument

  End If

  GetFile = (mstrErrorMessage = vbNullString)

Exit Function

LocalErr:
  mstrErrorMessage = Err.Description
  GetFile = False

End Function


Private Sub GetNewDocument()
  
  Dim mwrdFooterTable As Word.Table
  Dim strTempFile As String
  Dim lngView As Long
  Dim lngCount As Long
  
  On Local Error GoTo LocalErr
  
  If mstrWrdTemplate <> "" And Dir(mstrWrdTemplate) <> "" Then
    'Set mwrdDoc = mwrdApp.Documents.Open(mstrWrdTemplate, ReadOnly:=True)

    If Not IsFileCompatibleWithWordVersion(mstrWrdTemplate) Then
      mstrErrorMessage = "Your User Configuration Output options are set to use a template file which is not compatible with your version of Microsoft Office."
      Exit Sub
    End If

    'MH20030905 Fault 6911
    'If Word 2000 then make a copy of the template
    If Val(mwrdApp.Version) >= 9 And Val(mwrdApp.Version) < 10 Then
      strTempFile = mobjParent.GetTempFileName(vbNullString)
      FileCopy mstrWrdTemplate, strTempFile
    Else
      strTempFile = mstrWrdTemplate
    End If

    Set mwrdDoc = mwrdApp.Documents.Add(strTempFile)
    mwrdApp.Selection.EndKey Unit:=wdStory
  Else
    Set mwrdDoc = mwrdApp.Documents.Add

    ''Insert heading (slightly bigger, bold, underline and centered)
    '.Font.Bold = True
    '.Font.Size = .Font.Size + 2
    '.Font.Underline = wdUnderlineSingle
    '.ParagraphFormat.Alignment = wdAlignParagraphCenter

    mwrdDoc.PageSetup.Orientation = IIf(mblnWrdLandscape, wdOrientLandscape, wdOrientPortrait)

    With mwrdApp
      lngView = .ActiveWindow.View.Type
      .ActiveDocument.PageSetup.OddAndEvenPagesHeaderFooter = False
      .ActiveDocument.PageSetup.DifferentFirstPageHeaderFooter = False
      .ActiveWindow.View.Type = wdPageView
      .ActiveWindow.ActivePane.View.SeekView = wdSeekPrimaryFooter

      Set mwrdFooterTable = mwrdDoc.Tables.Add(Range:=.Selection.Range, NumRows:=1, NumColumns:=2)

      .Selection.TypeText Text:="Created on "
      '.Selection.Fields.Add .Selection.Range, Type:=wdFieldDate,
      'Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:=
      'NHRD20102004 Fault 7771 Made the footer 24 clock 'HH' = 24 hours clock: 'hh' = 12 hour clock
      .Selection.Fields.Add Range:=.Selection.Range, Type:=wdFieldDate, _
          Text:="CREATEDATE \@ ""dd/MM/yyyy at HH:mm"" ", PreserveFormatting:=True
      .Selection.TypeText Text:=" by " & gsUserName
      .Selection.MoveRight Unit:=wdCell
  
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      .Selection.TypeText Text:="Page "
      .Selection.Fields.Add .Selection.Range, Type:=wdFieldPage

      With mwrdFooterTable
        .Borders.InsideLineStyle = wdLineStyleNone
        .Borders.OutsideLineStyle = wdLineStyleNone
      End With

      .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
      .ActiveWindow.View.Type = lngView

    End With

  End If

  Set mwrdFooterTable = Nothing

Exit Sub

LocalErr:
  Set mwrdFooterTable = Nothing
  mstrErrorMessage = Err.Description

End Sub


Public Sub AddPage(strDefTitle As String, mstrSheetName As String, colStyles As Collection)

  Const strBookMark As String = "ASRSysTitleStart"
  Dim objRange As Word.Range
  Dim objStyle As clsOutputStyle

  On Local Error GoTo LocalErr

  mwrdDoc.Bookmarks.Add strBookMark

  With mwrdDoc.ActiveWindow.Selection
    
    'Remember the current setting to restore them later...
    Set objStyle = GetCurrentStyle(mwrdDoc.ActiveWindow.Selection)
    
    
    'If its the first page then add the definition title...
    mlngPageCount = mlngPageCount + 1
    If mlngPageCount = 1 Then
      .ParagraphFormat.Alignment = wdAlignParagraphCenter
      .TypeText strDefTitle
      .TypeParagraph
      .ParagraphFormat.Alignment = wdAlignParagraphLeft
      .TypeParagraph
    Else
      mwrdApp.Selection.InsertBreak Type:=wdPageBreak
    End If

    If mstrSheetName <> vbNullString And mobjParent.PageTitles Then
      .TypeText mstrSheetName
      .TypeParagraph
      .TypeParagraph
    End If
    
    'Move to the begining of the title, then highlight to end of document (minus one line)
    .GoTo What:=wdGoToBookmark, Name:=strBookMark
    .EndKey Unit:=wdStory, Extend:=wdExtend
    '.MoveUp Unit:=wdLine, Count:=1, Extend:=wdExtend

    'Apply the "title" style to this range
    ApplyStylesToRange mwrdApp.Selection.Range, colStyles("Title"), True
    
    'Now move to the end of the document, store the formatting, ready for the data...
    .EndKey Unit:=wdStory
    ApplyStylesToRange mwrdApp.Selection.Range, objStyle, False

  End With

Exit Sub

LocalErr:
  mstrErrorMessage = Err.Description

End Sub

Private Function GetCurrentStyle(objRange As Object) As clsOutputStyle

  Dim objStyle As clsOutputStyle

  On Local Error Resume Next

  Set objStyle = New clsOutputStyle

  With objRange
    objStyle.Bold = .Font.Bold
    objStyle.Underline = .Font.Underline
    objStyle.CenterText = (.ParagraphFormat.Alignment = wdAlignParagraphCenter)

    If Val(mwrdApp.Version) < 9 Then
      objStyle.BackCol = .Shading.BackgroundPatternColorIndex
      objStyle.ForeCol = .Font.ForegroundPatternColorIndex
    Else
      objStyle.BackCol = .Shading.BackgroundPatternColor
      objStyle.ForeCol = .Font.ForegroundPatternColor
    End If

  End With

  Set GetCurrentStyle = objStyle
  Set objStyle = Nothing

End Function

Public Sub DataArray(strArray() As String, colColumns As Collection, colStyles As Collection, colMerges As Collection)

  Const strBookMark As String = "ASRSysTableStart"
  Dim objColumn As clsColumn
  Dim strOutput As String
  Dim lngGridCol As Long
  Dim lngGridRow As Long
  Dim lngBorder As Long
  
  On Local Error GoTo LocalErr
  
  mwrdDoc.Bookmarks.Add strBookMark

  For lngGridRow = 0 To UBound(strArray, 2)
    strOutput = vbNullString
    For lngGridCol = 0 To UBound(strArray, 1)
      strOutput = _
        IIf(lngGridCol > 0, strOutput & vbTab, "") & _
        strArray(lngGridCol, lngGridRow)
    Next
    
    If gobjProgress.Visible And gobjProgress.Cancelled Then
      mstrErrorMessage = "Cancelled by user."
      Exit Sub
    End If
      
    mwrdDoc.ActiveWindow.Selection.TypeText strOutput
    mwrdDoc.ActiveWindow.Selection.TypeParagraph
  
    If mobjParent.UpdateProgressPerRow Then
      gobjProgress.UpdateProgress gblnBatchMode
    End If
  Next

  mwrdDoc.ActiveWindow.Selection.GoTo What:=wdGoToBookmark, Name:=strBookMark
  mwrdDoc.ActiveWindow.Selection.EndKey Unit:=wdStory, Extend:=wdExtend

  'convert selected text into a table
  Set mwrdTable = mwrdDoc.ActiveWindow.Selection.ConvertToTable _
    (Separator:=wdSeparateByTabs, _
    NumRows:=UBound(strArray, 2), _
    Format:=wdTableFormatNone, _
    ApplyFont:=False, _
    ApplyColor:=False, _
    AutoFit:=False)
  
  
  For lngGridCol = 0 To UBound(strArray, 1)
    mwrdTable.Columns(lngGridCol + 1).Select
    Select Case colColumns(lngGridCol + 1).DataType
    Case sqlNumeric, sqlInteger
      mwrdApp.Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
    Case sqlBoolean
      mwrdApp.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End Select
  Next


  With mwrdTable
    lngBorder = IIf(mblnWrdWordGridlines, wdLineStyleSingle, wdLineStyleNone)
    .Borders(wdBorderLeft).LineStyle = lngBorder
    .Borders(wdBorderRight).LineStyle = lngBorder
    .Borders(wdBorderTop).LineStyle = lngBorder
    .Borders(wdBorderBottom).LineStyle = lngBorder
    .Borders(wdBorderVertical).LineStyle = lngBorder
    .Borders(wdBorderHorizontal).LineStyle = lngBorder
  End With
  
  With colStyles("Title")
    .StartCol = -1
    .StartRow = -1
    .EndCol = -1
    .EndRow = -1
  End With
  
  
  With colStyles("Heading")
    .StartCol = 0
    .StartRow = 0
    .EndCol = mwrdTable.Columns.Count - 1
    .EndRow = mlngHeaderRows - 1
  End With
  
  
  If mlngHeaderCols > 0 Then
    With colStyles("HeadingCols")
      .StartCol = 0
      .StartRow = 0
      .EndCol = mlngHeaderCols - 1
      .EndRow = mwrdTable.Rows.Count - 1
    End With
  End If


  With colStyles("Data")
    .StartCol = mlngHeaderCols
    .StartRow = mlngHeaderRows
    .EndCol = mwrdTable.Columns.Count - 1
    .EndRow = mwrdTable.Rows.Count - 1
  End With
  
  
  'If mstrWrdTemplate = vbNullString Then
  If mblnApplyStyles Then
    ApplyStyle colStyles
    ApplyMerges colMerges
  End If


  'mwrdTable.Columns(0).Select
  
  'mwrdApp.Selection.Tables(1).AutoFitBehavior wdAutoFitWindow

  'With colStyles("Data")
  '  mwrdDoc.Range(mwrdTable.Cell(.StartRow + 1, .StartCol + 1).Range.Start, mwrdTable.Cell(.EndRow + 1, .EndCol + 1).Range.End).Select
  '  mwrdApp.Selection.Cells.DistributeWidth
  'End With

  mwrdApp.Selection.EndKey Unit:=wdStory
  mwrdDoc.ActiveWindow.Selection.TypeParagraph

Exit Sub

LocalErr:
  If Err.Number = 4608 Then
    mstrErrorMessage = "Cannot output more than 60 columns to Word"
  Else
    mstrErrorMessage = Err.Description
  End If

End Sub


Private Sub PrepareRows(lngStartRow As Long, lngRowCount As Long)

'  Dim objColumn As clsColumn
'  Dim lngStartCol As Long
'  Dim lngEndRow As Long
'  Dim lngCount As Long
'
'  With mwrdWorkSheet
'
'    If mstrWrdTemplate <> vbNullString Then
'      For lngCount = 1 To lngRowCount
'        .Rows(lngStartRow).Select
'        mwrdApp.Selection.Copy
'        mwrdApp.Selection.Insert Shift:=-4121 'xlDown
'      Next
'      mwrdApp.CutCopyMode = False
'    End If
'    .Range("A1").Select
'
'
'    lngStartCol = mwrdData.StartCol - 1
'    lngEndRow = lngStartRow + lngRowCount
'    For lngCount = 1 To mcolColumns.Count
'
'      Set objColumn = mcolColumns(lngCount)
'      With .Range(.Cells(lngStartRow, lngStartCol + lngCount), .Cells(lngEndRow, lngStartCol + lngCount))
'        Select Case objColumn.DataType
'        Case sqlNumeric, sqlInteger
'          .NumberFormat = "0" & IIf(objColumn.DecPlaces, "." & String(objColumn.DecPlaces, "0"), "")
'        Case Else
'          .NumberFormat = "@"   'Dates need to be exported as text as Word insists on changing some dates!
'        End Select
'      End With
'
'    Next
'
'  End With

End Sub


Private Sub ApplyStyle(colStyles As Collection)

  Dim objTemp As clsOutputStyle
  'Dim objRange As Object
  'Dim lngBorder As Long
  'Dim lngCol As Long
  'Dim lngRow As Long
  
  
  For Each objTemp In colStyles
    
    'Must do it row by row otherwise other cells get included in the formatting...
    'For lngRow = objTemp.StartRow To objTemp.EndRow
    '  Set objRange = mwrdDoc.Range(mwrdTable.Cell(lngRow + 1, objTemp.StartCol + 1).Range.Start, mwrdTable.Cell(lngRow + 1, objTemp.EndCol + 1).Range.End)
    
    'JPD 20030404 Changed the endCol/endRow conditions as we might
    'want to apply a style to the top-left (0,0) cell, as we do sometimes in Record Profile.
    'If objTemp.EndCol > 0 Or objTemp.EndRow > 0 Then
    If objTemp.EndCol >= 0 Or objTemp.EndRow >= 0 Then
      mwrdDoc.Range(mwrdTable.Cell(objTemp.StartRow + 1, objTemp.StartCol + 1).Range.Start, mwrdTable.Cell(objTemp.EndRow + 1, objTemp.EndCol + 1).Range.End).Select
      ApplyStylesToRange mwrdApp.Selection, objTemp, True
    End If
  
  Next

End Sub


Private Sub ApplyStylesToRange(objRange As Object, objStyle As clsOutputStyle, blnColourIndex As Boolean)

  On Local Error Resume Next
  
  With objRange
    .Font.Bold = objStyle.Bold
    .Font.Underline = objStyle.Underline

    If objStyle.CenterText Then
      .ParagraphFormat.Alignment = wdAlignParagraphCenter
    End If

    If Val(mwrdApp.Version) < 9 Then
      If blnColourIndex Then
        .Shading.BackgroundPatternColorIndex = GetWordColourIndex(objStyle.BackCol)
        .Font.ColorIndex = GetWordColourIndex(objStyle.ForeCol)
      Else
        .Shading.BackgroundPatternColorIndex = objStyle.BackCol
        .Font.ColorIndex = objStyle.ForeCol
      End If
    Else
      .Shading.BackgroundPatternColor = objStyle.BackCol
      .Font.Color = objStyle.ForeCol
    End If

    'MH20050907 Fault 10318
    If objStyle.Name = "Heading" Then
      .Rows.HeadingFormat = True
    End If
  
  End With

End Sub

Private Sub ApplyMerges(colMerges As Collection)

  Dim objTemp As clsOutputStyle
  Dim objRange As Object
  Dim lngBorder As Long
  Dim lngCol As Long
  Dim lngRow As Long
  Dim lngOffset As Long

  On Local Error GoTo LocalErr

  lngOffset = 1
  For Each objTemp In colMerges
    Set objRange = mwrdDoc.Range(mwrdTable.Cell(objTemp.StartRow + 1, objTemp.StartCol + lngOffset).Range.Start, _
      mwrdTable.Cell(objTemp.EndRow + 1, objTemp.EndCol + lngOffset).Range.End)
    objRange.Select
    mwrdApp.ActiveWindow.Selection.Cells.Merge
    lngOffset = lngOffset - (objTemp.EndCol - objTemp.StartCol)
  Next

Exit Sub

LocalErr:
  mstrErrorMessage = Err.Description

End Sub


Public Sub Complete()

  Dim objDefPrinter As cSetDfltPrinter
  Dim strDefaultPrinter As String
  Dim strFormat As String
  
  On Local Error GoTo LocalErr
  
  If mstrErrorMessage <> vbNullString Then
    Exit Sub
  End If
  
  
  'SAVE
  If mblnSave Then
    mstrErrorMessage = "Error saving file <" & mstrFileName & ">"
    
    strFormat = GetOfficeSaveAsFormat(mstrFileName, GetOfficeWordVersion)
    If strFormat = "" Then
      mstrErrorMessage = "This definition is set to save in a file format which is not compatible with your version of Microsoft Office."
      GoTo TidyAndExit
    End If
    mwrdDoc.SaveAs mstrFileName, FileFormat:=Val(strFormat)

  End If
  
  
  'EMAIL
  If mblnEmail Then
    mstrErrorMessage = "Error sending email"
    
    If Not mblnSave Then
      strFormat = GetOfficeSaveAsFormat(mobjParent.EmailAttachAs, GetOfficeWordVersion)
      If strFormat = "" Then
        mstrErrorMessage = "This definition is set to email an attachment in a file format which is not compatible with your version of Microsoft Office."
        GoTo TidyAndExit
      End If
      
      mstrFileName = mobjParent.GetTempFileName(mstrFileName)
      mwrdDoc.SaveAs mstrFileName, FileFormat:=Val(strFormat)
    End If
    
    mobjParent.SendEmail mstrFileName
  End If
  'If Not mblnSave And mstrFileName <> vbNullString Then
  '  Kill mstrFileName
  'End If


  'PRINTER
  If mblnPrinter Then
    Dim strCurrentPrinter As String
    mstrErrorMessage = "Error printing"
    
    strCurrentPrinter = mwrdApp.ActivePrinter
    mwrdApp.ActivePrinter = mstrPrinterName
    mwrdDoc.PrintOut
    mwrdApp.ActivePrinter = strCurrentPrinter
  
  End If
  
  
  'SCREEN
  If mblnScreen Then
    mstrErrorMessage = "Error displaying Word"
    mwrdApp.Selection.HomeKey Unit:=wdStory
    mwrdApp.Visible = True
    mwrdApp.WindowState = wdWindowStateMaximize
    mwrdApp.Activate
    Set mwrdApp = Nothing     'Stops word quitting!
  Else
    Do While mwrdApp.BackgroundPrintingStatus > 0 Or mwrdApp.BackgroundSavingStatus > 0
      DoEvents
    Loop
    
    mwrdDoc.Close False
'    mwrdApp.Quit
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


'Private Sub GetWidthFirstCol(wrdDoc As Word.Document)
'
'  Const strBookMark As String = "TableStart"
'  Dim lngCount As Long
'
'  With wrdDoc.ActiveWindow.Selection.End
'
'    '.Font.Size = 8
'    .Bookmarks.Add Range:=.Range, Name:=strBookMark
'
'    For lngCount = 0 To UBound(mvarHeadings(VER))
'      .TypeText Trim(mvarHeadings(VER)(lngCount))
'      .TypeParagraph
'    Next
'
'    'go to start of table highlight to end of document
'    .GoTo What:=wdGoToBookmark, Name:=strBookMark
'    .Bookmarks.Item(strBookMark).Delete
'    .EndKey Unit:=wdStory, Extend:=wdExtend
'
'    .ConvertToTable _
'      NumRows:=UBound(mvarHeadings(VER)) + 1, _
'      NumColumns:=1, _
'      Format:=wdTableFormatNone, _
'      ApplyFont:=False, _
'      ApplyColor:=False, _
'      AutoFit:=False
'
'    .Cells.AutoFit
'    mlngFirstColWidth = .Columns(1).Width + 2
'    .Tables(1).Delete
'
'  End With
'
'End Sub

Public Property Let HeaderRows(ByVal lngNewValue As Long)
  mlngHeaderRows = lngNewValue
End Property

Public Property Let HeaderCols(ByVal lngNewValue As Long)
  mlngHeaderCols = lngNewValue
End Property

Public Property Let ApplyStyles(ByVal blnNewValue As Boolean)
  mblnApplyStyles = blnNewValue
End Property

Public Property Let Parent(objNewValue As clsOutputRun)
  Set mobjParent = objNewValue
End Property

Public Property Get ErrorMessage() As String
  ErrorMessage = mstrErrorMessage
End Property

'Public Property Let SaveFileFormat(lngSaveFileFormat As Long)
'  mlngSaveFileFormat = lngSaveFileFormat
'End Property
'
'Public Property Get SaveFileFormat() As Long
'  SaveFileFormat = mlngSaveFileFormat
'End Property
'
'Public Property Let EmailFileFormat(lngEmailFileFormat As Long)
'  mlngEmailFileFormat = lngEmailFileFormat
'End Property
'
'Public Property Get EmailFileFormat() As Long
'  EmailFileFormat = mlngEmailFileFormat
'End Property
