VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOutputDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mfrmParent As Form
Attribute mfrmParent.VB_VarHelpID = -1
Private mlngFormat As OutputFormats
Private WithEvents cmdEmailGroup As CommandButton
Attribute cmdEmailGroup.VB_VarHelpID = -1
Private WithEvents cmdFileName As CommandButton
Attribute cmdFileName.VB_VarHelpID = -1
Private WithEvents cmdTitlePageTemplate As CommandButton
Attribute cmdTitlePageTemplate.VB_VarHelpID = -1
'Private mlngFileFormats As Collection
'Private mlngFileFormatsDefault As Long

Public Sub PopulateOutputControls(rsTemp As Recordset)

  With mfrmParent

    .chkDestination(desScreen).Value = IIf(rsTemp!OutputScreen, vbChecked, vbUnchecked)
    
    .chkDestination(desPrinter).Value = IIf(rsTemp!OutputPrinter, vbChecked, vbUnchecked)
    SetComboText .cboPrinterName, rsTemp!OutputPrinterName
    If rsTemp!OutputPrinterName <> vbNullString Then
      If .cboPrinterName.Text <> rsTemp!OutputPrinterName Then
        .cboPrinterName.AddItem rsTemp!OutputPrinterName
        .cboPrinterName.ListIndex = .cboPrinterName.NewIndex
        COAMsgBox "This definition is set to output to printer " & rsTemp!OutputPrinterName & _
               " which is not set up on your PC.", vbInformation, App.Title
      End If
    End If
    
    .chkDestination(desSave).Value = IIf(rsTemp!OutputSave, vbChecked, vbUnchecked)
    If .chkDestination(desSave).Value Then
      .txtFileName.Text = rsTemp!OutputFilename
      .txtFileName.SelLength = Len(.txtFileName.Text)
      '.txtFilename.Tag = rsTemp!OutputSaveFormat
      SetComboItem .cboSaveExisting, rsTemp!OutputSaveExisting
    End If
    
    .chkDestination(desEmail).Value = IIf(rsTemp!OutputEmail, vbChecked, vbUnchecked)
    If rsTemp!OutputEmail Then
      .txtEmailGroup.Text = datGeneral.GetEmailGroupName(rsTemp!OutputEmailAddr)
      .txtEmailGroup.Tag = rsTemp!OutputEmailAddr
      .txtEmailSubject.Text = rsTemp!OutputEmailSubject
      .txtEMailAttachAs.Text = IIf(IsNull(rsTemp!OutputEmailAttachAs), vbNullString, rsTemp!OutputEmailAttachAs)
      '.txtEmailAttachAs.Tag = rsTemp!OutputEmailFileFormat
    End If

  End With

End Sub

  
Private Function IsADestinationChecked() As Boolean

  Dim blnChecked As Boolean

  On Error Resume Next

  With mfrmParent.chkDestination
    blnChecked = False
    blnChecked = (blnChecked Or .Item(desScreen).Value = vbChecked)
    blnChecked = (blnChecked Or .Item(desPrinter).Value = vbChecked)
    blnChecked = (blnChecked Or .Item(desSave).Value = vbChecked)
    blnChecked = (blnChecked Or .Item(desEmail).Value = vbChecked)
  End With

  IsADestinationChecked = blnChecked

End Function


Public Property Let ParentForm(ByVal frmNewValue As Form)
  
  Set mfrmParent = frmNewValue
  Set cmdEmailGroup = mfrmParent.cmdEmailGroup
  Set cmdFileName = mfrmParent.cmdFileName
  If mfrmParent.Name = "frmBatchJob" Then Set cmdTitlePageTemplate = mfrmParent.cmdTitlePageTemplate
  
  If Not IsADestinationChecked Then
    CheckFirstAvailableDestination
  End If

End Property


Public Sub FormatClick(Index As Integer, Optional blnDataToScreen As Boolean = True)

  mlngFormat = Index
  CheckIfPreviewAvailable Index
  
  'Set mlngFileFormats = New Collection
  
  Select Case mlngFormat
  'DATA ONLY
  Case fmtDataOnly
    EnableDestinations blnDataToScreen, True, False, False

  'CSV / CMG
  Case fmtCSV, fmtCMGFile
    EnableDestinations False, False, True, True
    'JPD 20041202 Fault 9577
    'mfrmParent.CommonDialog1.Filter = "Comma Separated Values (*.csv)|*.csv|Text (*.txt)|*.txt"
    'mfrmParent.CommonDialog1.Filter = "Comma Separated Values (*.csv)|*.csv|Text (*.txt)|*.txt|All Files|*.*"
    frmMain.CommonDialog1.Filter = "Comma Separated Values (*.csv)|*.csv|Text (*.txt)|*.txt|All Files|*.*"


  'FIXEDLENGTH
  Case fmtFixedLengthFile
    EnableDestinations False, False, True, True
    'JPD 20041202 Fault 9577
    'mfrmParent.CommonDialog1.Filter = "Text (*.txt)|*.txt"
    'mfrmParent.CommonDialog1.Filter = "Text (*.txt)|*.txt|All Files|*.*"
    frmMain.CommonDialog1.Filter = "Text (*.txt)|*.txt|All Files|*.*"


  'HTML
  Case fmtHTML
    EnableDestinations True, False, True, True
    'mfrmParent.CommonDialog1.Filter = "HTML Document (*.htm)|*.htm"
    frmMain.CommonDialog1.Filter = "HTML Document (*.htm)|*.htm"

  'WORD
  Case fmtWordDoc
    EnableDestinations True, True, True, True
    'mfrmParent.CommonDialog1.Filter = gsOfficeFileFilter_Word
    'mfrmParent.CommonDialog1.Filter = GetCommonDialogFormats("Word", GetOfficeWordVersion)
    'InitialiseCommonDialogFormats mfrmParent.CommonDialog1, "Word", GetOfficeWordVersion
    InitialiseCommonDialogFormats frmMain.CommonDialog1, "Word", GetOfficeWordVersion

  'EXCEL
  Case fmtExcelWorksheet, fmtExcelChart, fmtExcelPivotTable
    EnableDestinations True, True, True, True
    'mfrmParent.CommonDialog1.Filter = gsOfficeFileFilter_Excel
    'mfrmParent.CommonDialog1.Filter = GetCommonDialogFormats("Excel", GetOfficeExcelVersion)
    'InitialiseCommonDialogFormats mfrmParent.CommonDialog1, "Excel", GetOfficeExcelVersion
    'InitialiseCommonDialogFormats mfrmParent.CommonDialog1, "Excel", GetOfficeExcelVersion
    InitialiseCommonDialogFormats frmMain.CommonDialog1, "Excel", GetOfficeExcelVersion

  'SQL TABLE
  Case fmtSQLTable
    EnableDestinations False, False, False, False

  End Select


  'In case the output format has changed, ensure
  'the extension of the output file name is correct
  If mfrmParent.txtFileName.Text <> vbNullString Then
    'mfrmParent.txtFilename.Text = Left(mfrmParent.txtFilename.Text, InStrRev(mfrmParent.txtFilename.Text, ".")) & _
      Right(Split(mfrmParent.CommonDialog1.Filter, ".")(1), 3)
    mfrmParent.txtFileName.Text = Left(mfrmParent.txtFileName.Text, InStrRev(mfrmParent.txtFileName.Text, ".")) & _
      Right(Split(frmMain.CommonDialog1.Filter, ".")(1), 3)
  End If

  'Need to check if we can allow "Add new Worksheet" option
  PopulateCombos False, True, False

End Sub


Private Sub CheckIfPreviewAvailable(Index As Integer)
  
  On Local Error Resume Next

  mfrmParent.chkPreview.Enabled = (Index <> fmtDataOnly)
  mfrmParent.chkPreview.Value = IIf(Index = fmtDataOnly, vbUnchecked, vbChecked)

End Sub


Private Sub EnableDestinations(blnScreen As Boolean, blnPrinter As Boolean, blnSave As Boolean, blnEmail As Boolean)
  
  Dim blnSelected As Boolean
  
  On Local Error Resume Next
  
  blnEmail = (blnEmail And mfrmParent.cboEmailAddr.ListCount > 0)
  
  
  With mfrmParent.chkDestination
  
    'SCREEN
    .Item(desScreen).Enabled = blnScreen
    .Item(desScreen).Value = vbUnchecked
    
    'PRINTER
    .Item(desPrinter).Enabled = blnPrinter
    .Item(desPrinter).Value = vbUnchecked
    
    'SAVE
    .Item(desSave).Enabled = blnSave
    .Item(desSave).Value = vbUnchecked
    
    'EMAIL
    .Item(desEmail).Enabled = blnEmail
    .Item(desEmail).Value = vbUnchecked

    If Not IsADestinationChecked Then
      CheckFirstAvailableDestination
    End If

  End With

End Sub


Private Function IsAvailable(lngIndex As Long) As Boolean

  Dim blnIsAvailable As Boolean

  On Local Error Resume Next
  blnIsAvailable = False
  With mfrmParent.chkDestination.Item(lngIndex)
    blnIsAvailable = (.Enabled Or Not .Visible)
  End With

  IsAvailable = blnIsAvailable

End Function


Private Function CheckFirstAvailableDestination()

  Dim blnChecked As Boolean

  With mfrmParent.chkDestination
  
    If Not blnChecked Then
      If IsAvailable(desScreen) Then
        .Item(desScreen).Value = vbChecked
      ElseIf IsAvailable(desPrinter) Then
        .Item(desPrinter).Value = vbChecked
      ElseIf IsAvailable(desSave) Then
        .Item(desSave).Value = vbChecked
      ElseIf IsAvailable(desEmail) Then
        .Item(desEmail).Value = vbChecked
      End If
    End If

  End With

End Function


Public Sub DestinationClick(Index As Integer)

  Dim blnPrinter As Boolean
  Dim blnSave As Boolean
  Dim blnEmail As Boolean

  On Local Error Resume Next

  blnPrinter = (mfrmParent.chkDestination(desPrinter).Value = vbChecked)
  blnSave = (mfrmParent.chkDestination(desSave).Value = vbChecked)
  blnEmail = (mfrmParent.chkDestination(desEmail).Value = vbChecked)
  
  With mfrmParent
  
    Select Case Index
    'PRINTER
    Case desPrinter
      .lblPrinter.Enabled = blnPrinter
      EnableCombo .cboPrinterName, blnPrinter
  
    'SAVE
    Case desSave
      .lblSave.Enabled = blnSave
      .lblSaveAs.Enabled = blnSave
      .lblFileName.Enabled = blnSave
      .cmdFileName.Enabled = blnSave
      If Not blnSave Then
        .txtFileName.Text = vbNullString
      End If
      EnableCombo .cboSaveExisting, blnSave
      EnableCombo .cboSaveAs, blnSave
  
    'EMAIL
    Case desEmail
      .lblEmail(0).Enabled = blnEmail
      .lblEmail(1).Enabled = blnEmail
      .lblEmail(2).Enabled = blnEmail
      .lblEmail(3).Enabled = blnEmail
      .cmdEmailGroup.Enabled = blnEmail
      .txtEmailSubject.Enabled = blnEmail
      .txtEmailSubject.BackColor = IIf(blnEmail, vbWindowBackground, vbButtonFace)
      .txtEMailAttachAs.Enabled = blnEmail
      .txtEMailAttachAs.BackColor = IIf(blnEmail, vbWindowBackground, vbButtonFace)
      'EnableCombo .cboEmailFileFormat, blnEmail
      
      If blnEmail Then
        .txtEmailGroup.Text = "<None>"
        .txtEmailSubject.Text = Trim(.txtName.Text)
        If (.txtFileName.Text <> vbNullString) Then
          .txtEMailAttachAs.Text = FileFromFullPath(.txtFileName.Text)
        ElseIf Trim(.txtName.Text) <> vbNullString Then
          .txtEMailAttachAs.Text = Trim(.txtName.Text) & "." & Choose(mlngFormat, "csv", "htm", "doc", "xls", "xls", "xls")
        End If
      Else
        .txtEmailGroup.Text = vbNullString
        .txtEmailGroup.Tag = 0
        .txtEmailSubject.Text = vbNullString
        .txtEMailAttachAs.Text = vbNullString
        .txtEMailAttachAs.Tag = 0
      End If

    End Select
    
  End With

End Sub


Public Sub PopulateCombos(blnPrinter As Boolean, blnSave As Boolean, blnEmail As Boolean)

  Dim objPrinter As Printer
  Dim rsEmailDefs As Recordset
  Dim strSQL As String
  Dim lngListIndex As Long

  'PRINTER
  If blnPrinter Then
    With mfrmParent.cboPrinterName
      .Clear
      .AddItem "<Default Printer>"
      For Each objPrinter In Printers
        .AddItem objPrinter.DeviceName
      Next
    End With
  End If
  
 
  'SAVE
  If blnSave Then
    With mfrmParent.cboSaveExisting
      lngListIndex = .ListIndex
      .Clear
      .AddItem "Overwrite"
      .ItemData(.NewIndex) = 0
      .AddItem "Do not overwrite"
      .ItemData(.NewIndex) = 1
      .AddItem "Add sequential number to name"
      .ItemData(.NewIndex) = 2
      .AddItem "Append to file"
      .ItemData(.NewIndex) = 3

      If IsExcel Then
        .AddItem "Create new sheet in workbook"
        .ItemData(.NewIndex) = 4
      End If
       
      If mfrmParent.chkDestination(desSave).Value = vbChecked Then
        .ListIndex = IIf(lngListIndex < .ListCount - 1, lngListIndex, 0)
      End If
    End With
  End If
  
  'PopulateFileFormatCombo mfrmParent.cboSaveAs, IIf(IsExcel, "Excel", "Word")
  'PopulateFileFormatCombo mfrmParent.cboEmailFileFormat, IIf(IsExcel, "Excel", "Word")
  
  
End Sub


Private Function IsExcel() As Boolean

  Dim blnIsExcel As Boolean

  On Local Error Resume Next

  blnIsExcel = False
  blnIsExcel = (blnIsExcel Or mfrmParent.optOutputFormat(fmtExcelWorksheet).Value)
  blnIsExcel = (blnIsExcel Or mfrmParent.optOutputFormat(fmtExcelChart).Value)
  blnIsExcel = (blnIsExcel Or mfrmParent.optOutputFormat(fmtExcelPivotTable).Value)
  IsExcel = blnIsExcel

End Function


Public Function GetSelectedFormatIndex() As Long

  Dim optTemp As OptionButton

  GetSelectedFormatIndex = 0  'Default
  For Each optTemp In mfrmParent.optOutputFormat
    If optTemp.Value = True Then
      GetSelectedFormatIndex = optTemp.Index
      Exit For
    End If
  Next

End Function


Public Sub ShowFormats(blnData As Boolean, blnCSV As Boolean, blnHTML As Boolean, _
  blnWord As Boolean, blnExcel As Boolean, blnChart As Boolean, blnPivot As Boolean)

  Dim lngTop As Long

  lngTop = 400
  lngTop = ShowFormat(fmtDataOnly, blnData, lngTop)
  lngTop = ShowFormat(fmtCSV, blnCSV, lngTop)
  lngTop = ShowFormat(fmtHTML, blnHTML, lngTop)
  lngTop = ShowFormat(fmtWordDoc, blnWord, lngTop)
  lngTop = ShowFormat(fmtExcelWorksheet, blnExcel, lngTop)
  lngTop = ShowFormat(fmtExcelChart, blnChart, lngTop)
  lngTop = ShowFormat(fmtExcelPivotTable, blnPivot, lngTop)

End Sub


Private Function ShowFormat(lngIndex As Long, blnVisible As Boolean, lngTop As Long) As Long
  
  mfrmParent.optOutputFormat.Item(lngIndex).Visible = blnVisible
  If blnVisible Then
    mfrmParent.optOutputFormat.Item(lngIndex).Top = lngTop
    lngTop = lngTop + 400
  End If
  
  ShowFormat = lngTop

End Function


Public Sub ReadDefFromRecset(rsDef As Recordset)

  With mfrmParent
    .optOutputFormat(rsDef!OutputFormat).Value = True
    .chkPreview.Value = IIf(rsDef!OutputPreview, vbChecked, vbUnchecked)
    .chkDestination(desScreen).Value = IIf(rsDef!OutputScreen, vbChecked, vbUnchecked)
  
    .chkDestination(desPrinter).Value = IIf(rsDef!OutputPrinter, vbChecked, vbUnchecked)
    If rsDef!OutputPrinterName <> vbNullString Then
      SetComboText .cboPrinterName, rsDef!OutputPrinterName
      If .cboPrinterName.Text <> rsDef!OutputPrinterName Then
        .cboPrinterName.AddItem rsDef!OutputPrinterName
        .cboPrinterName.ListIndex = .cboPrinterName.NewIndex
        COAMsgBox "This definition is set to output to printer " & rsDef!OutputPrinterName & _
               " which is not set up on your PC.", vbInformation, .Caption
      End If
    End If
  
    .chkDestination(desSave).Value = IIf(rsDef!OutputSave, vbChecked, vbUnchecked)
    If rsDef!OutputSave Then
      .txtFileName.Text = rsDef!OutputFilename
      '.txtFilename.Tag = rsDef!OutputSaveFormat
      SetComboItem .cboSaveExisting, rsDef!OutputSaveExisting
    End If

    .chkDestination(desEmail).Value = IIf(rsDef!OutputEmail, vbChecked, vbUnchecked)
    If rsDef!OutputEmail Then
      .txtEmailGroup.Text = datGeneral.GetEmailGroupName(rsDef!OutputEmailAddr)
      .txtEmailGroup.Tag = rsDef!OutputEmailAddr
      .txtEmailSubject.Text = rsDef!OutputEmailSubject
      .txtEMailAttachAs.Text = IIf(IsNull(rsDef!OutputEmailAttachAs), vbNullString, rsDef!OutputEmailAttachAs)
      '.txtEmailAttachAs.Tag = IIf(IsNull(rsDef!OutputEmailFileFormat), vbNullString, rsDef!OutputEmailFileFormat)
    End If
  End With

End Sub


Private Sub optOutputFormat_Click(Index As Integer)
  FormatClick Index
End Sub

Private Sub chkDestination_Click(Index As Integer)
  DestinationClick Index
End Sub

Private Sub cmdFileName_Click()

  Dim cd1 As CommonDialog
  Dim txtFileName As TextBox
  Dim txtEMailAttachAs As TextBox
  Dim blnEmail As Boolean
  Dim lngIndex As Long
  Dim lngFormat As Long

  On Local Error GoTo LocalErr
  
  'Set cd1 = mfrmParent.CommonDialog1
  Set cd1 = frmMain.CommonDialog1
  Set txtFileName = mfrmParent.txtFileName
  Set txtEMailAttachAs = mfrmParent.txtEMailAttachAs
  blnEmail = (mfrmParent.chkDestination(desEmail).Value = vbChecked)
  
  
  With cd1

    .FileName = txtFileName.Text
    If txtFileName.Text = vbNullString Then
      .InitDir = gsDocumentsPath
    End If

    .CancelError = True
    .DialogTitle = "Output file name"
    .Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNLongNames
      
      
'    lngFormat = IIf(txtFilename.Text <> "", Val(txtFilename.Tag), mlngFileFormatsDefault)
'    For lngIndex = 1 To mlngFileFormats.Count
'      If Val(mlngFileFormats(lngIndex)) = lngFormat Then
'        .FilterIndex = lngIndex
'        Exit For
'      End If
'    Next
    
    .ShowSave

    If .FileName <> vbNullString Then
      If Len(.FileName) > 255 Then
        COAMsgBox "Path and file name must not exceed 255 characters in length", vbExclamation, mfrmParent.Caption
      Else
        txtFileName.Text = vbNullString   'activates the change event
        txtFileName.Text = .FileName
        txtFileName.SelStart = Len(txtFileName.Text)
        'txtFilename.Tag = mlngFileFormats(.FilterIndex)
        
        If blnEmail Then
          If Trim(txtEMailAttachAs.Text) = vbNullString Then
            txtEMailAttachAs.Text = FileFromFullPath(.FileName)
          End If
        End If
      
      End If
    End If

  End With

Exit Sub

LocalErr:
  If Err.Number <> 32755 Then   '32755 = Cancel was selected.
    COAMsgBox Err.Description, vbCritical
  End If

End Sub


Public Function ValidDestination() As Boolean

  Dim blnDestination As Boolean
  Dim blnPreview As Boolean

  ValidDestination = False

  If mfrmParent.chkDestination(desSave).Value = vbChecked Then
    If mfrmParent.txtFileName.Text = vbNullString Then
      COAMsgBox "You must enter a file name.", vbExclamation, mfrmParent.Caption
      Exit Function
    End If
  End If

  If mfrmParent.chkDestination(desEmail).Value = vbChecked Then
    If Val(mfrmParent.txtEmailGroup.Tag) = 0 Then
      COAMsgBox "You must select an email group.", vbExclamation, mfrmParent.Caption
      Exit Function
    End If

    If datGeneral.GetEmailGroupName(Val(mfrmParent.txtEmailGroup.Tag)) = vbNullString Then
      COAMsgBox "The email group has been deleted by another user.", vbExclamation, mfrmParent.Caption
      mfrmParent.txtEmailGroup.Text = vbNullString
      mfrmParent.txtEmailGroup.Tag = 0
      Exit Function
    End If

    If mfrmParent.txtEMailAttachAs.Text = vbNullString Then
      COAMsgBox "You must enter an email attachment file name.", vbExclamation, mfrmParent.Caption
      Exit Function
    End If
    
    If InStr(mfrmParent.txtEMailAttachAs.Text, "/") Or _
       InStr(mfrmParent.txtEMailAttachAs.Text, ":") Or _
       InStr(mfrmParent.txtEMailAttachAs.Text, "?") Or _
       InStr(mfrmParent.txtEMailAttachAs.Text, Chr(34)) Or _
       InStr(mfrmParent.txtEMailAttachAs.Text, "<") Or _
       InStr(mfrmParent.txtEMailAttachAs.Text, ">") Or _
       InStr(mfrmParent.txtEMailAttachAs.Text, "|") Or _
       InStr(mfrmParent.txtEMailAttachAs.Text, "\") Or _
       InStr(mfrmParent.txtEMailAttachAs.Text, "*") Then
          COAMsgBox "The email attachment file name cannot contain any of the following characters:" & vbCrLf & _
                 "/  :  ?  " & Chr(34) & "  <  >  |  \  *", vbExclamation, mfrmParent.Caption
          Exit Function
    End If
  
  
  End If
  
  On Local Error Resume Next
  blnPreview = (mfrmParent.chkPreview.Value = vbChecked)

  With mfrmParent.chkDestination
    blnDestination = False
    blnDestination = (blnDestination Or .Item(desScreen).Value = vbChecked)
    blnDestination = (blnDestination Or .Item(desPrinter).Value = vbChecked)
    blnDestination = (blnDestination Or .Item(desSave).Value = vbChecked)
    blnDestination = (blnDestination Or .Item(desEmail).Value = vbChecked)
  End With

  If Not blnDestination Then
    COAMsgBox "You must select a destination" & _
           IIf(blnPreview, " in addition to preview on screen.", "."), _
        vbExclamation, mfrmParent.Caption
    Exit Function
  End If

  ValidDestination = True

End Function


Private Sub cmdEmailGroup_Click()

  Dim frmDefinition As frmEmailDefGroup
  Dim frmSelection As frmDefSel
  Dim lForms As Long
  Dim blnExit As Boolean
  Dim blnOK As Boolean

  Set frmSelection = New frmDefSel
  blnExit = False

  Set frmDefinition = New frmEmailDefGroup
  
  With frmSelection

    .Options = edtAdd + edtDelete + edtEdit + edtCopy + edtPrint + edtProperties + edtSelect + edtDeselect
    .EnableRun = False
    .TableComboVisible = False
    .SelectedID = Val(mfrmParent.txtEmailGroup.Tag)

    Do While Not blnExit
      
      If .ShowList(utlEmailGroup) Then

        .Show vbModal
        Select Case .Action
        Case edtAdd
          Set frmDefinition = New frmEmailDefGroup
          frmDefinition.Initialise True, .FromCopy
          frmDefinition.Show vbModal
          .SelectedID = frmDefinition.SelectedID
          Unload frmDefinition
          Set frmDefinition = Nothing

        'TM20010808 Fault 2656 - Must validate the definition before allowing the edit/copy.
        Case edtEdit
          Set frmDefinition = New frmEmailDefGroup
          frmDefinition.Initialise False, .FromCopy, .SelectedID
          If Not frmDefinition.Cancelled Then
            frmDefinition.Show vbModal
            If .FromCopy And frmDefinition.SelectedID > 0 Then
              .SelectedID = frmDefinition.SelectedID
            End If
          End If
          Unload frmDefinition
          Set frmDefinition = Nothing

        Case edtPrint
          Set frmDefinition = New frmEmailDefGroup
          frmDefinition.Initialise False, .FromCopy, .SelectedID
          If Not frmDefinition.Cancelled Then
            frmDefinition.PrintDef .SelectedID
          End If
          Unload frmDefinition
          Set frmDefinition = Nothing

        Case edtSelect
          mfrmParent.txtEmailGroup.Text = .SelectedText
          mfrmParent.txtEmailGroup.Tag = .SelectedID
          mfrmParent.Changed = True
          blnExit = True

        Case edtDeselect
          mfrmParent.txtEmailGroup.Text = "<None>"
          mfrmParent.txtEmailGroup.Tag = 0
          mfrmParent.Changed = True
          blnExit = True

        Case 0
          blnExit = True  'cancel

        End Select

      End If

    Loop
  End With

  Unload frmSelection
  Set frmSelection = Nothing

End Sub

Private Sub cmdTitlePageTemplate_Click()

  Dim cd1 As CommonDialog
  Dim txtFileName As TextBox
  Dim txtReportTitle As TextBox
  
  Dim lngIndex As Long
  Dim lngFormat As Long

  On Local Error GoTo LocalErr
  
  'Set cd1 = mfrmParent.CommonDialog1
  Set cd1 = frmMain.CommonDialog1
  Set txtFileName = mfrmParent.txtTitlePage
  Set txtReportTitle = mfrmParent.txtReportPackTitle
  'blnEmail = (mfrmParent.chkDestination(desEmail).Value = vbChecked)
  
  
  With cd1

    .FileName = txtFileName.Text
    If txtFileName.Text = vbNullString Then
      .InitDir = gsDocumentsPath
    End If

    .CancelError = True
    .DialogTitle = "Output Title Page Filename"
    .Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNLongNames
 
    .ShowSave

    If .FileName <> vbNullString Then
      If Len(.FileName) > 255 Then
        COAMsgBox "Path and file name must not exceed 255 characters in length", vbExclamation, mfrmParent.Caption
      Else
        txtFileName.Text = vbNullString   'activates the change event
        txtFileName.Text = .FileName
        txtFileName.SelStart = Len(txtFileName.Text)
        txtReportTitle.Text = FileFromFullPath(.FileName)
      End If
    End If

  End With

Exit Sub

LocalErr:
  If Err.Number <> 32755 Then   '32755 = Cancel was selected.
    COAMsgBox Err.Description, vbCritical
  End If

End Sub

Private Sub mfrmParent_Activate()
  If Not IsADestinationChecked Then
    CheckFirstAvailableDestination
  End If
End Sub

Private Function FileFromFullPath(strFullFileName As String) As String
  FileFromFullPath = Mid(strFullFileName, InStrRev(strFullFileName, "\") + 1)
End Function


'Private Function PopulateFileFormatCombo(cbo As ComboBox, strDestin As String)
'
'  Dim rsTemp As Recordset
'  Dim strSQL As String
'
'  On Local Error GoTo LocalErr
'
'  strSQL = "SELECT * FROM ASRSysFileFormats " & _
'           "WHERE Destination = '" & Replace(strDestin, "'", "''") & "'"
'  Set rsTemp = datGeneral.GetRecords(strSQL)
'
'  With cbo
'    .Clear
'    Do While Not rsTemp.EOF
'      .AddItem rsTemp.Fields("Description").Value
'      .ItemData(.NewIndex) = rsTemp.Fields("Value").Value
'
'      rsTemp.MoveNext
'    Loop
'  End With
'
'LocalErr:
'  If Not rsTemp Is Nothing Then
'    If rsTemp.State <> adStateClosed Then
'      rsTemp.Close
'    End If
'    Set rsTemp = Nothing
'  End If
'
'End Function
