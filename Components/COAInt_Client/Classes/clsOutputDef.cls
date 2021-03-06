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

Private mfrmParent As Form
Private mlngFormat As Long
Private WithEvents cmdEmailGroup As CommandButton
Attribute cmdEmailGroup.VB_VarHelpID = -1
Private WithEvents cmdFilename As CommandButton
Attribute cmdFilename.VB_VarHelpID = -1

Public Property Let ParentForm(ByVal frmNewValue As Form)
  Set mfrmParent = frmNewValue
  Set cmdEmailGroup = mfrmParent.cmdEmailGroup
  Set cmdFilename = mfrmParent.cmdFilename
  EnsureADestinationIsChecked
End Property


Public Sub FormatClick(Index As Integer, Optional blnDataToScreen As Boolean = True)

  mlngFormat = Index
  CheckIfPreviewAvailable Index
  
  Select Case mlngFormat
  'DATA ONLY
  Case fmtDataOnly
    EnableDestinations blnDataToScreen, True, False, False

  'CSV / CMG
  Case fmtCSV, fmtCMGFile
    EnableDestinations False, False, True, True
    mfrmParent.CommonDialog1.Filter = "Comma Separated Values (*.csv)|*.csv|Text (*.txt)|*.txt"

  'FIXEDLENGTH
  Case fmtFixedLengthFile
    EnableDestinations False, False, True, True
    mfrmParent.CommonDialog1.Filter = "Text (*.txt)|*.txt"

  'HTML
  Case fmtHTML
    EnableDestinations True, False, True, True
    mfrmParent.CommonDialog1.Filter = "HTML Document (*.htm)|*.htm"

  'WORD
  Case fmtWordDoc
    EnableDestinations True, True, True, True
    mfrmParent.CommonDialog1.Filter = "Word Document (*.doc)|*.doc"

  'EXCEL
  Case fmtExcelWorksheet, fmtExcelChart, fmtExcelPivotTable
    EnableDestinations True, True, True, True
    mfrmParent.CommonDialog1.Filter = "Excel Workbook (*.xls)|*.xls"

  'SQL TABLE
  Case fmtSQLTable
    EnableDestinations False, False, False, False

  End Select


  'In case the output format has changed, ensure
  'the extension of the output file name is correct
  If mfrmParent.txtFilename.Text <> vbNullString Then
    mfrmParent.txtFilename.Text = Left(mfrmParent.txtFilename.Text, InStrRev(mfrmParent.txtFilename.Text, ".")) & _
      Right(Split(mfrmParent.CommonDialog1.Filter, ".")(1), 3)
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
    'If Not blnScreen Then
      .Item(desScreen).Value = vbUnchecked
    'End If
    
    'PRINTER
    .Item(desPrinter).Enabled = blnPrinter
    'If Not blnPrinter Then
      .Item(desPrinter).Value = vbUnchecked
    'End If
    
    'SAVE
    .Item(desSave).Enabled = blnSave
    'If Not blnSave Then
      .Item(desSave).Value = vbUnchecked
    'End If
    
    'EMAIL
    .Item(desEmail).Enabled = blnEmail
    'If Not blnEmail Then
      .Item(desEmail).Value = vbUnchecked
    'End If

    EnsureADestinationIsChecked

  End With

End Sub


Private Function IsEnabled(lngIndex As Long) As Boolean

  Dim blnIsEnabled As Boolean

  On Local Error Resume Next
  blnIsEnabled = False
  blnIsEnabled = mfrmParent.chkDestination.Item(lngIndex).Enabled

  IsEnabled = blnIsEnabled

End Function


Private Function EnsureADestinationIsChecked()

  Dim blnChecked As Boolean

  With mfrmParent.chkDestination
  
    blnChecked = False
    If Not blnChecked Then
      If IsEnabled(desScreen) Then
        blnChecked = (.Item(desScreen).Value = vbChecked)
      End If
    End If
  
    If Not blnChecked Then
      If IsEnabled(desPrinter) Then
        blnChecked = (.Item(desPrinter).Value = vbChecked)
      End If
    End If
    
    If Not blnChecked Then
      If IsEnabled(desSave) Then
        blnChecked = (.Item(desSave).Value = vbChecked)
      End If
    End If
    
    If Not blnChecked Then
      If IsEnabled(desEmail) Then
        blnChecked = (.Item(desEmail).Value = vbChecked)
      End If
    End If
  
  
    If Not blnChecked Then
      If IsEnabled(desScreen) Then
        .Item(desScreen).Value = vbChecked
      ElseIf IsEnabled(desPrinter) Then
        .Item(desPrinter).Value = vbChecked
      ElseIf IsEnabled(desSave) Then
        .Item(desSave).Value = vbChecked
      ElseIf IsEnabled(desEmail) Then
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
      EnableCombo .cboSaveExisting, blnSave
  
    'EMAIL
    Case desEmail
      .lblEmail(0).Enabled = blnEmail
      .lblEmail(1).Enabled = blnEmail
      .cmdEmailGroup.Enabled = blnEmail
      .txtEmailSubject.Enabled = blnEmail
      .txtEmailSubject.BackColor = IIf(blnEmail, vbWindowBackground, vbButtonFace)
      If blnEmail Then
        .txtEmailGroup.Text = "<None>"
      Else
        .txtEmailGroup.Text = vbNullString
        .txtEmailGroup.Tag = 0
        .txtEmailSubject.Text = vbNullString
      End If
  
    End Select
    
    .lblFileName.Enabled = (blnSave Or blnEmail)
    .cmdFilename.Enabled = (blnSave Or blnEmail)
    If Not (blnSave Or blnEmail) Then
      .txtFilename.Text = vbNullString
    End If

  End With

End Sub


Public Sub PopulateCombos(blnPrinter As Boolean, blnSave As Boolean, blnEmail As Boolean)

  Dim objPrinter As Printer
  'Dim rsEmailDefs As Recordset
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
      .AddItem "Add sequential number to filename"
      .ItemData(.NewIndex) = 2
      .AddItem "Append to file"
      .ItemData(.NewIndex) = 3

      If IsExcel Then
        .AddItem "Create new worksheet in workbook"
        .ItemData(.NewIndex) = 4
      End If
       
      If mfrmParent.chkDestination(desSave).Value = vbChecked Then
        .ListIndex = IIf(lngListIndex < .ListCount - 1, lngListIndex, 0)
      End If
    End With
  End If
  
  'EMAIL
  'If blnEmail Then
  '  strSQL = "SELECT DISTINCT EmailID, Name, Fixed " & _
  '           "FROM ASRSYSEmailAddress " & _
  '           "WHERE Type = 0 ORDER BY Name"
  '  Set rsEmailDefs = datGeneral.GetReadOnlyRecords(strSQL)
  '
  '  With mfrmParent.cboEmailAddr
  '    .Clear
  '    Do While Not rsEmailDefs.EOF
  '      .AddItem rsEmailDefs!Name
  '      .ItemData(.NewIndex) = rsEmailDefs!EmailID
  '      rsEmailDefs.MoveNext
  '    Loop
  '  End With
  'End If

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

  'Dim lngIndex As Long
  Dim optTemp As OptionButton

  GetSelectedFormatIndex = 0  'Default
  'For lngIndex = 0 To mfrmParent.optOutputFormat.Count - 1
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


'Public Sub ReadDefFromRecset(rsDef As Recordset)
'
'  With mfrmParent
'    .optOutputFormat(rsDef!OutputFormat).Value = True
'    .chkPreview.Value = IIf(rsDef!OutputPreview, vbChecked, vbUnchecked)
'    .chkDestination(desScreen).Value = IIf(rsDef!OutputScreen, vbChecked, vbUnchecked)
'
'    .chkDestination(desPrinter).Value = IIf(rsDef!OutputPrinter, vbChecked, vbUnchecked)
'    If rsDef!OutputPrinterName <> vbNullString Then
'      SetComboText .cboPrinterName, rsDef!OutputPrinterName
'      If .cboPrinterName.Text <> rsDef!OutputPrinterName Then
'        .cboPrinterName.AddItem rsDef!OutputPrinterName
'        .cboPrinterName.ListIndex = .cboPrinterName.NewIndex
'        MsgBox "This definition is set to output to printer " & rsDef!OutputPrinterName & _
'               " which is not set up on your PC.", vbInformation, .Caption
'      End If
'    End If
'
'    .chkDestination(desSave).Value = IIf(rsDef!OutputSave, vbChecked, vbUnchecked)
'    If rsDef!OutputSave Then
'      SetComboItem .cboSaveExisting, rsDef!OutputSaveExisting
'    End If
'
'    .chkDestination(desEmail).Value = IIf(rsDef!OutputEmail, vbChecked, vbUnchecked)
'    If rsDef!OutputEmail Then
'      .txtEmailGroup.Text = datGeneral.GetEmailGroupName(rsDef!OutputEmailAddr)
'      .txtEmailGroup.Tag = rsDef!OutputEmailAddr
'      .txtEmailSubject.Text = rsDef!OutputEmailSubject
'    End If
'    .txtFilename.Text = rsDef!OutputFilename
'  End With
'
'End Sub


Private Sub optOutputFormat_Click(Index As Integer)
  FormatClick Index
End Sub

Private Sub chkDestination_Click(Index As Integer)
  DestinationClick Index
End Sub

Private Sub cmdFilename_Click()

  On Local Error GoTo LocalErr
  
  With mfrmParent.CommonDialog1
    If mfrmParent.txtFilename.Text = vbNullString Then
      .InitDir = gsDocumentsPath
    Else
      .FileName = mfrmParent.txtFilename.Text
    End If

    .CancelError = True
    .DialogTitle = "Output Filename"
    .Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNLongNames
    .ShowSave

    If .FileName <> vbNullString Then
      If Len(.FileName) > 255 Then
        MsgBox "Path and filename must not exceed 255 characters in length", vbExclamation, mfrmParent.Caption
      Else
        mfrmParent.txtFilename = .FileName
      End If
    End If

  End With

Exit Sub

LocalErr:
  If Err.Number <> 32755 Then   '32755 = Cancel was selected.
    MsgBox Err.Description, vbCritical
  End If

End Sub


Public Function ValidDestination() As Boolean

  Dim blnValidDestination As Boolean

  blnValidDestination = False

  If (mfrmParent.chkDestination(desSave).Value = vbChecked Or _
     mfrmParent.chkDestination(desEmail).Value = vbChecked) Then
      If mfrmParent.txtFilename.Text = vbNullString Then
        MsgBox "You must enter a filename", vbExclamation, mfrmParent.Caption
        Exit Function
      End If
  End If

  If mfrmParent.chkDestination(desEmail).Value = vbChecked Then
    If Val(mfrmParent.txtEmailGroup.Tag) = 0 Then
      MsgBox "You must select an email group", vbExclamation, mfrmParent.Caption
      Exit Function
    End If
  End If

  blnValidDestination = _
    (mfrmParent.chkDestination(desScreen).Value = vbChecked Or _
     mfrmParent.chkDestination(desPrinter).Value = vbChecked Or _
     mfrmParent.chkDestination(desSave).Value = vbChecked Or _
     mfrmParent.chkDestination(desEmail).Value = vbChecked)

  If Not blnValidDestination Then
    MsgBox "You must select a destination", vbExclamation, mfrmParent.Caption
    Exit Function
  End If

  blnValidDestination = True
  ValidDestination = blnValidDestination

End Function


Private Sub cmdEmailGroup_Click()

'  Dim frmDefinition As frmEmailDefGroup
'  Dim frmSelection As frmDefSel
'  Dim lForms As Long
'  Dim blnExit As Boolean
'  Dim blnOK As Boolean
'
'  Set frmSelection = New frmDefSel
'  blnExit = False
'
'  Set frmDefinition = New frmEmailDefGroup
'
'  With frmSelection
'
'    .Options = edtAdd + edtDelete + edtEdit + edtCopy + edtPrint + edtProperties + edtSelect + edtDeselect
'    .EnableRun = False
'    .TableComboVisible = False
'    .SelectedID = Val(mfrmParent.txtEmailGroup.Tag)
'
'    Do While Not blnExit
'
'      If .ShowList("EMAILGROUPS") Then
'
'        .Show vbModal
'        Select Case .Action
'        Case edtAdd
'          Set frmDefinition = New frmEmailDefGroup
'          frmDefinition.Initialise True, .FromCopy
'          frmDefinition.Show vbModal
'          .SelectedID = frmDefinition.SelectedID
'          Unload frmDefinition
'          Set frmDefinition = Nothing
'
'        'TM20010808 Fault 2656 - Must validate the definition before allowing the edit/copy.
'        Case edtEdit
'          Set frmDefinition = New frmEmailDefGroup
'          frmDefinition.Initialise False, .FromCopy, .SelectedID
'          If Not frmDefinition.Cancelled Then
'            frmDefinition.Show vbModal
'            If .FromCopy And frmDefinition.SelectedID > 0 Then
'              .SelectedID = frmDefinition.SelectedID
'            End If
'          End If
'          Unload frmDefinition
'          Set frmDefinition = Nothing
'
'        Case edtPrint
'          Set frmDefinition = New frmEmailDefGroup
'          frmDefinition.Initialise False, .FromCopy, .SelectedID
'          If Not frmDefinition.Cancelled Then
'            frmDefinition.PrintDef .SelectedID
'          End If
'          Unload frmDefinition
'          Set frmDefinition = Nothing
'
'        Case edtSelect
'          mfrmParent.txtEmailGroup.Text = .SelectedText
'          mfrmParent.txtEmailGroup.Tag = .SelectedID
'          blnExit = True
'
'        Case edtDeselect
'          mfrmParent.txtEmailGroup.Text = "<None>"
'          mfrmParent.txtEmailGroup.Tag = 0
'          blnExit = True
'
'        Case 0
'          blnExit = True  'cancel
'
'        End Select
'
'      End If
'
'    Loop
'  End With
'
'  Unload frmSelection
'  Set frmSelection = Nothing
'
End Sub
