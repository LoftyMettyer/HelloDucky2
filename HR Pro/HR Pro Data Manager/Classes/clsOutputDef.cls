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
Private WithEvents cmdFilename As CommandButton
Attribute cmdFilename.VB_VarHelpID = -1
Private WithEvents cmdTitlePageTemplate As CommandButton
Attribute cmdTitlePageTemplate.VB_VarHelpID = -1
'Private mlngFileFormats As Collection
'Private mlngFileFormatsDefault As Long

Public Property Get Format() As OutputFormats
  Format = mlngFormat
End Property

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
               " which is not set up on your PC.", vbInformation, app.Title
      End If
    End If
    
    .chkDestination(desSave).Value = IIf(rsTemp!OutputSave, vbChecked, vbUnchecked)
    If .chkDestination(desSave).Value Then
      .txtFilename.Text = rsTemp!OutputFilename
      .txtFilename.SelLength = Len(.txtFilename.Text)
      '.txtFilename.Tag = rsTemp!OutputSaveFormat
      SetComboItem .cboSaveExisting, rsTemp!OutputSaveExisting
    End If
    
    .chkDestination(desEmail).Value = IIf(rsTemp!OutputEmail, vbChecked, vbUnchecked)
    If rsTemp!OutputEmail Then
      .txtEmailGroup.Text = datGeneral.GetEmailGroupName(rsTemp!OutputEmailAddr)
      .txtEmailGroup.Tag = rsTemp!OutputEmailAddr
      .txtEmailSubject.Text = rsTemp!OutputEmailSubject
      .txtEmailAttachAs.Text = IIf(IsNull(rsTemp!OutputEmailAttachAs), vbNullString, rsTemp!OutputEmailAttachAs)
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
  Set cmdFilename = mfrmParent.cmdFilename
  
  If mfrmParent.Name = "frmBatchJob" Then Set cmdTitlePageTemplate = mfrmParent.cmdTitlePageTemplate
  
  If Not IsADestinationChecked Then
    CheckFirstAvailableDestination
  End If

End Property


Public Sub FormatClick(Index As Integer, Optional blnDataToScreen As Boolean = True)
  mlngFormat = Index
  CheckIfPreviewAvailable Index
  
  Select Case mlngFormat
    Case fmtDataOnly
      EnableDestinations blnDataToScreen, True, False, False
      EnableTitlePageTemplate False
      
    Case fmtCSV, fmtCMGFile
      EnableDestinations False, False, True, True
      EnableTitlePageTemplate False
    
    Case fmtFixedLengthFile
      EnableDestinations False, False, True, True
      EnableTitlePageTemplate False
    
    Case fmtHTML
      EnableDestinations True, False, True, True
      EnableTitlePageTemplate False
      EnableReportPackTitle True
      If gblnReportPackMode Then EnableReportPackOutputs True, False, False, True
    
    Case fmtWordDoc
      EnableDestinations True, True, True, True
      EnableTitlePageTemplate True
      EnableReportPackTitle True
      If gblnReportPackMode Then EnableReportPackOutputs True, False, True, True
    
    Case fmtExcelWorksheet, fmtExcelchart, fmtExcelPivotTable
      EnableDestinations True, True, True, True
      EnableTitlePageTemplate True
      EnableReportPackTitle False
      If gblnReportPackMode Then EnableReportPackOutputs False, True, False, True
    
    Case fmtSQLTable
      EnableDestinations False, False, False, False
  End Select

  'Need to check if we can allow "Add new Worksheet" option
  PopulateCombos False, True, False
End Sub

Private Sub EnableTitlePageTemplate(enable As Boolean)

  If Not cmdTitlePageTemplate Is Nothing Then
    cmdTitlePageTemplate.Enabled = enable
  End If
  
End Sub

Private Sub EnableReportPackTitle(enable As Boolean)
  On Error Resume Next
  EnableControl mfrmParent.txtReportPackTitle, enable
  Select Case mlngFormat
  Case fmtExcelWorksheet, fmtExcelchart, fmtExcelPivotTable
    mfrmParent.txtReportPackTitle.Text = ""
  End Select
End Sub

Private Sub EnableReportPackOutputs(blnTOC As Boolean, blnRetainPivot As Boolean, blnForceCover As Boolean, blnBlankTitle As Boolean)
  Dim blnSelected As Boolean
  On Local Error Resume Next
   
  With mfrmParent
    'TOC
    .chkTOC.Enabled = blnTOC
    .chkTOC.Value = vbUnchecked
    'Retain Pivot
    .chkRetainPivot.Enabled = blnRetainPivot
    .chkRetainPivot.Value = vbUnchecked
    '.ChkRetainCharts.Enabled = blnRetainPivot
    '.ChkRetainCharts.Value = vbUnchecked
    'Force Cover
    .chkForceCoverSheet.Enabled = blnForceCover
    .chkForceCoverSheet.Value = vbUnchecked
  End With
End Sub

Private Sub CheckIfPreviewAvailable(Index As Integer)
  On Local Error Resume Next
    
  If gblnReportPackMode Then
    mfrmParent.chkPreview.Value = False
  Else
    mfrmParent.chkPreview.Enabled = IIf(Index <> fmtDataOnly, True, False)
    mfrmParent.chkPreview.Value = IIf(Index = fmtDataOnly, vbUnchecked, vbChecked)
  End If

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
      .cmdFilename.Enabled = blnSave
      If Not blnSave Then
        .txtFilename.Text = vbNullString
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
      .txtEmailAttachAs.Enabled = blnEmail
      .txtEmailAttachAs.BackColor = IIf(blnEmail, vbWindowBackground, vbButtonFace)
      'EnableCombo .cboEmailFileFormat, blnEmail
      
      If blnEmail Then
        .txtEmailGroup.Text = "<None>"
        .txtEmailSubject.Text = Trim(.txtName.Text)
        If (.txtFilename.Text <> vbNullString) Then
          .txtEmailAttachAs.Text = FileFromFullPath(.txtFilename.Text)
        ElseIf Trim(.txtName.Text) <> vbNullString Then
          .txtEmailAttachAs.Text = Trim(.txtName.Text) & "." & Choose(mlngFormat, "csv", "htm", "doc", "xls", "xls", "xls")
        End If
      Else
        .txtEmailGroup.Text = vbNullString
        .txtEmailGroup.Tag = 0
        .txtEmailSubject.Text = vbNullString
        .txtEmailAttachAs.Text = vbNullString
        .txtEmailAttachAs.Tag = 0
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
      If gblnReportPackMode Then
        .AddItem "Overwrite"
        .ItemData(.NewIndex) = 0
        .AddItem "Add sequential number to name"
        .ItemData(.NewIndex) = 2
      Else
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
      End If
      
      If mfrmParent.chkDestination(desSave).Value = vbChecked Then
        .ListIndex = IIf(lngListIndex < .ListCount - 1, lngListIndex, 0)
      End If
    End With
  End If
  
'  PopulateFileFormatCombo mfrmParent.cboSaveAs, IIf(IsExcel, "Excel", "Word")
'  PopulateFileFormatCombo mfrmParent.cboEmailFileFormat, IIf(IsExcel, "Excel", "Word")
  
  
End Sub


Private Function IsExcel() As Boolean

  Dim blnIsExcel As Boolean

  On Local Error Resume Next

  blnIsExcel = False
  blnIsExcel = (blnIsExcel Or mfrmParent.optOutputFormat(fmtExcelWorksheet).Value)
  blnIsExcel = (blnIsExcel Or mfrmParent.optOutputFormat(fmtExcelchart).Value)
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
  lngTop = ShowFormat(fmtExcelchart, blnChart, lngTop)
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
      .txtFilename.Text = rsDef!OutputFilename
      '.txtFilename.Tag = rsDef!OutputSaveFormat
      SetComboItem .cboSaveExisting, rsDef!OutputSaveExisting
    End If

    .chkDestination(desEmail).Value = IIf(rsDef!OutputEmail, vbChecked, vbUnchecked)
    If rsDef!OutputEmail Then
      .txtEmailGroup.Text = datGeneral.GetEmailGroupName(rsDef!OutputEmailAddr)
      .txtEmailGroup.Tag = rsDef!OutputEmailAddr
      .txtEmailSubject.Text = rsDef!OutputEmailSubject
      .txtEmailAttachAs.Text = IIf(IsNull(rsDef!OutputEmailAttachAs), vbNullString, rsDef!OutputEmailAttachAs)
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
  Dim txtFilename As TextBox
  Dim txtEmailAttachAs As TextBox
  Dim blnEmail As Boolean
  Dim lngIndex As Long
  Dim blnExtensionsDiffer As Boolean
  
  On Local Error GoTo LocalErr
  
  Set cd1 = frmMain.CommonDialog1
  Set txtFilename = mfrmParent.txtFilename
  Set txtEmailAttachAs = mfrmParent.txtEmailAttachAs
  blnEmail = (mfrmParent.chkDestination(desEmail).Value = vbChecked)
    
  With cd1
  
    Select Case mlngFormat
      Case fmtCSV, fmtCMGFile
        .Filter = "Comma Separated Values (*.csv)|*.csv|Text (*.txt)|*.txt|All Files|*.*"
      Case fmtFixedLengthFile
        .Filter = "Text (*.txt)|*.txt|All Files|*.*"
      Case fmtHTML
        .Filter = "HTML Document (*.htm)|*.htm"
      Case fmtWordDoc
        InitialiseCommonDialogFormats cd1, "Word", GetOfficeWordVersion, DirectionOutput
      Case fmtExcelWorksheet, fmtExcelchart, fmtExcelPivotTable
        InitialiseCommonDialogFormats cd1, "Excel", GetOfficeExcelVersion, DirectionOutput
      Case Else
        Debug.Assert False
    End Select
  
  
    .FileName = txtFilename.Text
    If txtFilename.Text = vbNullString Then
      .InitDir = gsDocumentsPath
    End If

    .CancelError = True
    .DialogTitle = "Output file name"
    .Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNLongNames
    
    ' Save the save as dropdown to a correct format
    SetSaveAsType
    
    .ShowSave

    If .FileName <> vbNullString Then
      If Len(.FileName) > 255 Then
        COAMsgBox "Path and file name must not exceed 255 characters in length", vbExclamation, mfrmParent.Caption
      Else
        If gblnReportPackMode Then
          ' Interrogate and remove wrong template file from Title Page textbox
          blnExtensionsDiffer = (Mid(mfrmParent.txtFilename, InStrRev(mfrmParent.txtFilename, ".") + 1)) <> (Mid(mfrmParent.txtTitlePage, InStrRev(mfrmParent.txtTitlePage, ".") + 1))
          If blnExtensionsDiffer Then mfrmParent.txtTitlePage = ""
        End If
        
        txtFilename.Text = vbNullString   'activates the change event
        txtFilename.Text = .FileName
        txtFilename.SelStart = Len(txtFilename.Text)
        
        If blnEmail Then
          If Trim(txtEmailAttachAs.Text) = vbNullString Then
            txtEmailAttachAs.Text = FileFromFullPath(.FileName)
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

Private Sub SetSaveAsType()

  Dim aryFileBits() As String
  Dim strExtension As String
  Dim iCount As Integer
 
  If mfrmParent.txtFilename.Text <> vbNullString Then

    aryFileBits = Split(mfrmParent.txtFilename.Text, ".")
    strExtension = UCase(aryFileBits(UBound(aryFileBits)))

    aryFileBits = Split(frmMain.CommonDialog1.Filter, "(*.")
    
    For iCount = LBound(aryFileBits) To UBound(aryFileBits)
      If InStr(1, UCase(aryFileBits(iCount)), strExtension + ")|*.", vbTextCompare) > 0 Then
        frmMain.CommonDialog1.FilterIndex = iCount
        Exit For
      End If
    Next iCount
    
  Else
    frmMain.CommonDialog1.FilterIndex = 1
  End If
End Sub


Public Function ValidDestination() As Boolean

  Dim blnDestination As Boolean
  Dim blnPreview As Boolean

  ValidDestination = False

  If mfrmParent.chkDestination(desSave).Value = vbChecked Then
    If mfrmParent.txtFilename.Text = vbNullString Then
      COAMsgBox "You must enter a file name.", vbExclamation, mfrmParent.Caption
      Exit Function
    End If
  End If

  If mfrmParent.chkDestination(desEmail).Value = vbChecked Then
  
    Select Case mlngFormat
    Case fmtWordDoc, fmtExcelWorksheet, fmtExcelchart, fmtExcelPivotTable
      If mfrmParent.txtEmailAttachAs.Text Like "*.html" Then
        COAMsgBox "You cannot email html output from word or excel.", vbExclamation, mfrmParent.Caption
        Exit Function
      End If
    End Select
  
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

    If mfrmParent.txtEmailAttachAs.Text = vbNullString Then
      COAMsgBox "You must enter an email attachment file name.", vbExclamation, mfrmParent.Caption
      Exit Function
    End If
    
    If InStr(mfrmParent.txtEmailAttachAs.Text, "/") Or _
       InStr(mfrmParent.txtEmailAttachAs.Text, ":") Or _
       InStr(mfrmParent.txtEmailAttachAs.Text, "?") Or _
       InStr(mfrmParent.txtEmailAttachAs.Text, Chr(34)) Or _
       InStr(mfrmParent.txtEmailAttachAs.Text, "<") Or _
       InStr(mfrmParent.txtEmailAttachAs.Text, ">") Or _
       InStr(mfrmParent.txtEmailAttachAs.Text, "|") Or _
       InStr(mfrmParent.txtEmailAttachAs.Text, "\") Or _
       InStr(mfrmParent.txtEmailAttachAs.Text, "*") Then
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

    .SelectedUtilityType = utlEmailGroup
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
          'mfrmParent.Changed = True
          blnExit = True

        Case edtDeselect
          mfrmParent.txtEmailGroup.Text = "<None>"
          mfrmParent.txtEmailGroup.Tag = 0
          'mfrmParent.Changed = True
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
  Dim txtFilename As TextBox
  Dim txtReportTitle As TextBox
  
  Dim lngIndex As Long
  Dim lngFormat As Long

  On Local Error GoTo LocalErr
  
  Set cd1 = frmMain.CommonDialog1
  Set txtFilename = mfrmParent.txtTitlePage
  Set txtReportTitle = mfrmParent.txtReportPackTitle
  'lngIndex = fmtExcelWorksheet 'default to Excel
 
  With cd1
    Select Case mlngFormat
      Case fmtWordDoc
        .Filter = "Word Template (*.dot;*.dotx;*.doc;*.docx)|*.dot;*.dotx;*.doc;*.docx"
      Case fmtExcelWorksheet
        .Filter = "Excel Template (*.xlt;*.xltx;*.xls;*.xlsx)|*.xlt;*.xltx;*.xls;*.xlsx"
      Case Else
        Debug.Assert False
    End Select

    .FileName = vbNullString ' txtFilename.Text
    'If txtFilename.Text = vbNullString Then
      .InitDir = gsDocumentsPath
    'End If

    .CancelError = True
    .DialogTitle = "Report Pack Output Template"
    .Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNFileMustExist
    
    .ShowOpen

    If .FileName <> vbNullString Then
      If Len(.FileName) > 255 Then
        COAMsgBox "Path and file name must not exceed 255 characters in length", vbExclamation, mfrmParent.Caption
      Else
        txtFilename.Text = vbNullString   'activates the change event
        txtFilename.Text = .FileName
        txtFilename.SelStart = Len(txtFilename.Text)
        mfrmParent.cmdTitlePageClear.Enabled = txtFilename.SelStart <> 0
        
        If txtReportTitle.Text = vbNullString Then
          txtReportTitle.Text = FileFromFullPath(.FileName)
          txtReportTitle.Text = Left(txtReportTitle.Text, InStrRev(txtReportTitle.Text, ".") - 1)
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

Private Sub mfrmParent_Activate()
  If Not IsADestinationChecked Then
    CheckFirstAvailableDestination
  End If
End Sub

Private Function FileFromFullPath(strFullFileName As String) As String
  FileFromFullPath = Mid(strFullFileName, InStrRev(strFullFileName, "\") + 1)
End Function

