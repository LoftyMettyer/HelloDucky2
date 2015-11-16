VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMailMergeRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private Const sTYPECODE_HEADING = "H"
Private Const sTYPECODE_SEPARATOR = "S"
Private Const sTYPECODE_COLUMN = "C"
Private Const sTYPECODE_EXPRESSION = "E"

Private mdatData As clsDataAccess
Private mrsMailMergeColumns As Recordset
Private mrsMergeData As Recordset
Private mlngMailMergeID As Long
'Private gblnBatchMode As Boolean
Private mblnNoRecords As Boolean

Private mlngSuccessCount As Long
Private mlngFailCount As Long

Private fOK As Boolean
Private mstrStatusMessage As String
Private mblnUserCancelled As Boolean
'Private mlngSingleRecordID As Long


'Merge Definition Variables
Private mstrDefName As String
Private mstrDefBaseTable As String
Private mlngRecordDescExprID As Long
Private mlngDefBaseTableID As Long
Private mlngDefOrderID As Long
Private mintDefSelection As Integer
Private mlngDefPickListID As Long
Private mlngDefFilterID As Long
Private mstrDefTemplateFile As String
Private mstrDefTemplateFormat As String
Private mblnDefPauseBeforeMerge As Boolean
Private mblnDefSuppressBlankLines As Boolean
Private mbIsLabel As Boolean

'Private mintDefOutput As Integer
'Private mblnDefDocSave As Boolean
'Private mstrDefDocFile As String
'Private mstrDefDocFileFormat As String
'Private mblnDefCloseDoc As Boolean
'Private mstrOutputPrinterName As String

Private mstrDefEmailCol As String
Private mstrDefEMailSubject As String
Private mblnDefEMailAttachment As Boolean
Private mstrDefAttachmentName As String
Private mstrDefAttachmentFormat As String
Private mstrEmailAddr As String
Private mlngEmailAddrCalc As Long

Private mintOutputFormat As Integer
Private mblnOutputScreen As Boolean
Private mblnOutputPrinter As Boolean
Private mstrOutputPrinterName As String
Private mblnOutputSave As Boolean
Private mstrOutputFileName As String
Private mstrOutputFileFormat As String

Private mlngDocumentMapID As Long
Private mbManualHeader As Boolean



Private mintType() As Integer
Private mlngSize() As Long
Private mintDecimals() As Integer
'Private ambStartOnNewLine() As Boolean
Private amColumnType() As String
Private amHeadingText() As String
Private amColumnInformation() As Variant

'SQL Variables
Private mstrSQLSelect As String
Private mstrSQLFrom As String
Private mstrSQLJoin As String
Private mstrSQLWhere As String
Private mstrSQLOrder As String

Private mlngTableViews() As Long
Private mstrWhereIDs As String

'Word Variables
Private wrdApp As New Word.Application
Private wrdDocTemplate As New Word.Document
Private wrdDocDataSource As New Word.Document
Private wrdDocOutput As New Word.Document
Private mstrDataSourceName As String
Private mstrDataSourceFormat As String
Private mblnAppVisible As Boolean

Private mbDefinitionOwner As Boolean

' Label definition variables
Private mlngLabelDefinitionID As Long
Private msngLabelTopMargin As Single
Private msngLabelSideMargin As Single
Private msngLabelHorizontalPitch As Long
Private msngLabelVerticalPitch As Long
Private msngLabelHeight As Long
Private msngLabelWidth As Long
Private miLabelNumberDown As Integer
Private miLabelNumberAcross As Integer

Private msngLabelPageWidth As Single
Private msngLabelPageHeight As Single
Private mintLabelPageOrientation As Integer
Private mbIsEnvelope As Boolean

Private msngEnvelopeFromTop As Single
Private msngEnvelopeFromLeft As Single
Private mbEnvelopeFromTopAuto As Boolean
Private mbEnvelopeFromLeftAuto As Boolean
    
Private miLabelStartColumn As Integer
Private miLabelStartRow As Integer

' Array holding the User Defined functions that are needed for this report
Private mastrUDFsRequired() As String

Private mobjFormatDetails(1) As DataMgr.clsOutputStyle

Private mstrPicklistFilterIDs As String

Private Sub Class_Initialize()
  Set mdatData = New DataMgr.clsDataAccess
End Sub

Private Sub Class_Terminate()
  
  Set mdatData = Nothing
  Set mrsMailMergeColumns = Nothing
  Set mrsMergeData = Nothing

  Set wrdApp = Nothing
  Set wrdDocTemplate = Nothing
  Set wrdDocDataSource = Nothing
  Set wrdDocOutput = Nothing

End Sub

Public Property Get ErrorString() As String
  ErrorString = mstrStatusMessage
End Property

Public Property Get UserCancelled() As Boolean
  UserCancelled = mblnUserCancelled
End Property

Private Function Records(lngRec As Long) As String
  Records = CStr(lngRec) & IIf(lngRec <> 1, " records", " record")
End Function

Private Function Progress() As Boolean
  
  'This needs to be here, otherwise the progress bar will continue to the end
  'rather than cancelling immediately
  If fOK = False Then
    Progress = False
    Exit Function
  End If
  
  If gobjProgress.Cancelled Then
    mblnUserCancelled = True
    fOK = False
  End If
  
  gobjProgress.UpdateProgress gblnBatchMode
  
  Progress = fOK
  
End Function


Private Function MergeFieldExists() As Boolean

  Dim strTemplateFieldName As String
  Dim intCount As Integer
  
  intCount = 1
  Do While intCount <= wrdDocTemplate.Fields.Count

    strTemplateFieldName = Trim$(wrdDocTemplate.Fields(intCount).Code)
    If Left$(strTemplateFieldName, 10) = "MERGEFIELD" Then
      MergeFieldExists = True
      Exit Function
    End If
  
    intCount = intCount + 1
  Loop

  MergeFieldExists = False

End Function

'Public Function ExecuteMailMerge(lngMailMergeID As Long, Optional lngRecordID As Long) As Boolean
Public Function ExecuteMailMerge(lngMailMergeID As Long, Optional strRecordIDs As String) As Boolean

  Dim strProcedureName As String
  Dim plngEventLogID As Long

  mlngMailMergeID = lngMailMergeID
  'mlngSingleRecordID = lngRecordID
  mstrPicklistFilterIDs = strRecordIDs
  fOK = True

  ReDim mastrUDFsRequired(0)

  Set mobjFormatDetails(0) = New DataMgr.clsOutputStyle
  Set mobjFormatDetails(1) = New DataMgr.clsOutputStyle

  'SQL Stuff
  If fOK Then Call SQLGetMergeDefinition
  If fOK Then Call InitialiseProgressBar
  
  If mbIsLabel Then
    gobjEventLog.AddHeader eltLabel, mstrDefName
  Else
    gobjEventLog.AddHeader eltMailMerge, mstrDefName
  End If
  
  If mbIsLabel Then
    If Progress Then Call RetrieveLabelDefinition
    'If Progress And Not gblnBatchMode And mlngSingleRecordID > 0 And Not mbIsEnvelope Then
    If Progress And Not gblnBatchMode And Not mbIsEnvelope Then
      If mstrPicklistFilterIDs <> "" And InStr(mstrPicklistFilterIDs, ",") = 0 Then
        Call GetLabelStartPosition
      End If
    End If
  End If
  
  If Progress Then Call SQLGetMergeColumns
  If Progress Then Call SQLCodeReset

  If Progress Then Call SQLCodeCreate
  If Progress Then Call UDFFunctions(mastrUDFsRequired, True)
  If Progress Then Call SQLGetMergeData
  
  If Progress Then Call WrdOpenApp
  
  ' Create the Labels template
  If mbIsLabel Then
    If mbIsEnvelope Then
      If Progress Then Call WrdCreateEnvelopeTemplate(False)
    Else
      If Progress Then Call WrdCreateLabelTemplate(False)
    End If
  End If
  
  If Progress Then Call WrdCreateDataSource
  If Progress Then Call WrdPopulateDataSource
  If Progress Then Call WrdOpenTemplate(False)
  
  If fOK Then
    If gblnBatchMode Or wrdDocTemplate.ReadOnly = True Then
      'if running in batch mode then attempt to run
      'the merge as is, and report any error which
      'may occur.
      If Progress Then Call WrdValidateTemplateMergeFields(True)
      If Progress Then Call WrdSetMergeOptions
      'If Progress Then Call WrdSetV1HeaderText
    
    Else
      'If not running in batch mode then
      'check for pause before merge etc.
      'and report any errors
      If Progress Then Call WrdValidateTemplateMergeFields(False)
      If Progress Then Call WrdSetMergeOptions
      If Progress Then Call WrdPauseBeforeMerge
      If Progress Then Call WrdValidateTemplateMergeFields(False)
      If Progress Then Call WrdCheckIfSaveTemplate
      'If Progress Then Call WrdSetV1HeaderText
    
    End If
  End If

  If Progress Then Call WrdRunMailMerge
  'If Progress Then Call WrdPrintDocument(mstrOutputPrinterName)
 
  
  
  'MH20041112 Fault 9001 & 9003
  'If Progress Then Call UDFFunctions(mastrUDFsRequired, False)
  Call UDFFunctions(mastrUDFsRequired, False)

  Call TidyUpAndExit
  Call OutputJobStatus
  
  ExecuteMailMerge = fOK
  
End Function


Private Sub OutputJobStatus()

  If mbIsLabel Then
    Call UtilUpdateLastRun(utlLabel, mlngMailMergeID)
  Else
    Call UtilUpdateLastRun(utlMailMerge, mlngMailMergeID)
  End If
  Dim strCaption As String

  If fOK = False Then
    mlngFailCount = mlngFailCount + mlngSuccessCount
    mlngSuccessCount = 0
  End If
  
  If mlngFailCount > 0 And Not mblnUserCancelled Then
    fOK = False
  End If

  With gobjEventLog
  
    If fOK Or mblnNoRecords Then
      .ChangeHeaderStatus elsSuccessful, mlngSuccessCount, mlngFailCount
      mstrStatusMessage = "Completed successfully."
    
      If mblnNoRecords Then
        mstrStatusMessage = mstrStatusMessage & vbCrLf & _
            "No records meet selection criteria."
        .AddDetailEntry mstrStatusMessage
        fOK = True
      End If

    ElseIf mblnUserCancelled Then
      .ChangeHeaderStatus elsCancelled, mlngSuccessCount, mlngFailCount
      mstrStatusMessage = "Cancelled by user."
    Else
      If mstrStatusMessage <> vbNullString Then
        mstrStatusMessage = "Failed." & vbCrLf & vbCrLf & mstrStatusMessage
        'Only details records for failures !
        .AddDetailEntry mstrStatusMessage
      Else
        mstrStatusMessage = "failed" & vbCrLf & vbCrLf
      End If
      .ChangeHeaderStatus elsFailed, mlngSuccessCount, mlngFailCount
    End If
  
  End With
  
  If mbIsLabel Then
    mstrStatusMessage = "Envelopes & Labels : '" & mstrDefName & "' " & mstrStatusMessage
    strCaption = "Envelopes & Labels"
  Else
    mstrStatusMessage = "Mail Merge : '" & mstrDefName & "' " & mstrStatusMessage
    strCaption = "Mail Merge"
  End If
  
  If Not mblnNoRecords And Not mblnUserCancelled Then
    mstrStatusMessage = mstrStatusMessage & vbCrLf & vbCrLf & _
      Records(mlngSuccessCount) & " successful"
  
    If mlngFailCount > 0 Then
      mstrStatusMessage = mstrStatusMessage & vbCrLf & _
        Records(mlngFailCount) & " failed"
    End If
  End If
  
  If Not gblnBatchMode Then
    COAMsgBox mstrStatusMessage, IIf(fOK, vbInformation, vbExclamation), strCaption
  End If

End Sub


Private Sub InitialiseProgressBar()
  
  With gobjProgress
    '.AviFile = App.Path & "\videos\word.avi"
    .AVI = dbText
    .MainCaption = IIf(mbIsLabel, "Envelopes & Labels", "Mail Merge")
    If Not gblnBatchMode Then
      .NumberOfBars = 1
      .Caption = IIf(mbIsLabel, "Envelopes & Labels", "Mail Merge")
      .Time = False
      .Cancel = True
      .Bar1Value = 0
      .Bar1MaxValue = 19
      .Bar1Caption = IIf(mbIsLabel, "Envelopes & Labels", "Mail Merge") & " : " & mstrDefName
      .OpenProgress
    Else
      .ResetBar2
      .Bar2MaxValue = IIf(mbIsLabel, 16, 13)
      .Bar2Caption = IIf(mbIsLabel, "Envelopes & Labels", "Mail Merge") & " : " & mstrDefName
    End If
  End With
  
End Sub


Private Sub SQLCodeReset()

  mstrSQLSelect = vbNullString
  mstrSQLFrom = vbNullString
  mstrSQLJoin = vbNullString
  mstrSQLWhere = vbNullString
  mstrSQLOrder = vbNullString

End Sub


Private Sub SQLGetMergeData()

  Dim strSQL As String
  
  On Local Error GoTo LocalErr
  strSQL = "SELECT " & mstrSQLSelect & vbCrLf & _
           " FROM " & mstrSQLFrom & mstrSQLJoin & vbCrLf & _
                      mstrSQLWhere & vbCrLf & _
                      mstrSQLOrder & vbCrLf
  
  'JPD 20031120 Fault 7677
  Set mrsMergeData = mdatData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly, adUseServer)

  If mrsMergeData.EOF Then
    mblnNoRecords = True
    fOK = False
  Else
    mblnNoRecords = False
  End If

Exit Sub

LocalErr:
  mstrStatusMessage = "Error retrieving merge data"
  fOK = False
  Exit Sub

End Sub


Private Sub SQLGetMergeDefinition()

  On Error GoTo LocalErr

  Dim rsMailMergeDefinition As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT ASRSysMailMergeName.*, " & _
           "ASRSysTables.TableName as TableName, " & _
           "ASRSysTables.RecordDescExprID as RecordDescExprID " & _
           "FROM ASRSysMailMergeName " & _
           "JOIN ASRSYSTables ON (ASRSysTables.TableID = ASRSysMailMergeName.TableID) " & vbCrLf & _
           "WHERE MailMergeID = " & mlngMailMergeID & " "
  Set rsMailMergeDefinition = mdatData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
  If rsMailMergeDefinition.BOF And rsMailMergeDefinition.EOF Then
    Set rsMailMergeDefinition = Nothing
    mstrStatusMessage = "This definition has been deleted by another user."
    fOK = False
    Exit Sub
  End If

  With rsMailMergeDefinition

    mstrDefName = !Name
    mlngDefBaseTableID = !TableID
    mstrDefBaseTable = !TableName
    mlngRecordDescExprID = !RecordDescExprID
    'mlngDefOrderID = !OrderID
    mintDefSelection = !Selection
    mlngDefPickListID = !PicklistID
    mlngDefFilterID = !FilterID
    mstrDefTemplateFile = !TemplateFileName
    mstrDefTemplateFormat = GetOfficeSaveAsFormat(mstrDefTemplateFile, GetOfficeWordVersion, oaWord)
    mblnDefPauseBeforeMerge = !PauseBeforeMerge
    mblnDefSuppressBlankLines = !SuppressBlanks
    mbIsLabel = !IsLabel
    mlngLabelDefinitionID = IIf(IsNull(!LabelTypeID), 0, !LabelTypeID)
    mintOutputFormat = !OutputFormat
    'mintDefOutput = !Output
    'mblnDefCloseDoc = !CloseDoc
    
    mlngEmailAddrCalc = 0
    mstrEmailAddr = vbNullString
        
    If Not mbIsLabel And mstrDefTemplateFormat = vbNullString Then
      mstrStatusMessage = "This definition has been defined using OpenHR Web and has not had a template file assigned for use in Data Manager. Please edit the definition and select the required template file."
      fOK = False
      GoTo TidyAndExit
    End If

    Select Case mintOutputFormat
      Case 0    'Word Document
        'mblnDefDocSave = !DocSave
        'mstrDefDocFile = !DocFileName
        'mstrDefDocFileFormat = GetOfficeSaveAsFormat(mstrDefDocFile, GetOfficeWordVersion)
        mblnOutputScreen = !OutputScreen
        mblnOutputPrinter = !OutputPrinter
        mstrOutputPrinterName = IIf(IsNull(!OutputPrinterName), "", !OutputPrinterName)
        mblnOutputSave = !OutputSave
        mstrOutputFileName = IIf(IsNull(!OutputFilename), "", !OutputFilename)
        mstrOutputFileFormat = GetOfficeSaveAsFormat(mstrOutputFileName, GetOfficeWordVersion, oaWord)
        
        If Trim(mstrOutputFileName) <> "" And Trim(mstrOutputFileFormat) = "" Then
          mstrStatusMessage = "This definition is set to save in a file format which is not compatible with your version of Microsoft Office."
          fOK = False
          GoTo TidyAndExit
        End If
      
      Case 1    'Individual Emails
        'MH20040122 Fault 5665
        If datGeneral.SystemPermission("EMAILADDRESSES", "VIEW") = False Then
          mstrStatusMessage = "You do not have permission to use email addresses."
          Set rsMailMergeDefinition = Nothing
          fOK = False
          Exit Sub
        End If
  
  
        If IIf(IsNull(!EmailAddrID), 0, !EmailAddrID) = 0 Then
          mstrStatusMessage = "No destination email address selected."
          Set rsMailMergeDefinition = Nothing
          fOK = False
          Exit Sub
        End If
        
        'mstrDefEmailCol = !EmailAddrName  'GetTableAndColumnName(!EmailColumnID)
        mblnDefEMailAttachment = !EMailAsAttachment
        mstrDefAttachmentName = IIf(IsNull(!EmailAttachmentName), "", !EmailAttachmentName)
        mstrDefAttachmentFormat = GetOfficeSaveAsFormat(mstrDefAttachmentName, GetOfficeWordVersion, oaWord)
        mstrDefEMailSubject = !EmailSubject
        mlngEmailAddrCalc = !EmailAddrID
    
        If Trim(mstrDefAttachmentName) <> vbNullString And Trim(mstrDefAttachmentFormat) = vbNullString Then
          mstrStatusMessage = "This definition is set to email an attachment in a file format which is not compatible with your version of Microsoft Office."
          fOK = False
          GoTo TidyAndExit
        End If
      
      Case 2    'Document Management
        mblnOutputScreen = !OutputScreen
        mblnOutputPrinter = True
        mstrOutputPrinterName = !OutputPrinterName
        mlngDocumentMapID = IIf(IsNull(!DocumentMapID), 0, !DocumentMapID)
        mbManualHeader = IIf(IsNull(!ManualDocManHeader), 0, !ManualDocManHeader)
    
    End Select

    mbDefinitionOwner = (LCase(Trim(gsUserName)) = LCase(Trim(!userName)))


  End With

  fOK = IsRecordSelectionValid


  If fOK Then
    If Not ValidPrinter(mstrOutputPrinterName) Then
      mstrStatusMessage = _
          "This definition is set to output to printer " & mstrOutputPrinterName & _
          " which is not set up on your PC."
      fOK = False
    End If
  End If


TidyAndExit:
  Set rsMailMergeDefinition = Nothing

Exit Sub

LocalErr:
  mstrStatusMessage = "Error reading Mail Merge definition."
  fOK = False
  Resume TidyAndExit

End Sub

Private Function ValidPrinter(strName As String) As Boolean

  Dim objPrinter As Printer
  Dim blnFound As Boolean

  If strName <> vbNullString And strName <> "<Default Printer>" Then
    blnFound = False
    For Each objPrinter In Printers
      If objPrinter.DeviceName = strName Then
        blnFound = True
        Exit For
      End If
    Next
  Else
    blnFound = True
  End If

  ValidPrinter = blnFound

End Function

Private Function IsRecordSelectionValid() As Boolean

  Dim sSQL As String
  Dim lCount As Long
  Dim rsTemp As Recordset
  Dim iResult As RecordSelectionValidityCodes
  
  ' Filter
  If mlngDefFilterID > 0 Then
    iResult = ValidateRecordSelection(REC_SEL_FILTER, mlngDefFilterID)
    Select Case iResult
      Case REC_SEL_VALID_DELETED
        mstrStatusMessage = "The base table filter used in this definition has been deleted by another user."
      Case REC_SEL_VALID_INVALID
        mstrStatusMessage = "The base table filter used in this definition is invalid."
      Case REC_SEL_VALID_HIDDENBYOTHER
        If Not gfCurrentUserIsSysSecMgr Then
          mstrStatusMessage = "The base table filter used in this definition has been made hidden by another user."
        End If
    End Select
  ElseIf mlngDefPickListID > 0 Then
    iResult = ValidateRecordSelection(REC_SEL_PICKLIST, mlngDefPickListID)
    Select Case iResult
      Case REC_SEL_VALID_DELETED
        mstrStatusMessage = "The base table picklist used in this definition has been deleted by another user."
      Case REC_SEL_VALID_INVALID
        mstrStatusMessage = "The base table picklist used in this definition is invalid."
      Case REC_SEL_VALID_HIDDENBYOTHER
        If Not gfCurrentUserIsSysSecMgr Then
          mstrStatusMessage = "The base table picklist used in this definition has been made hidden by another user."
        End If
    End Select
  End If
 
'******* Check calculations for hidden/deleted elements *******
  If Len(mstrStatusMessage) = 0 Then

    sSQL = "SELECT * FROM ASRSysMailMergeColumns " & _
             "WHERE MailMergeID = " & mlngMailMergeID & _
             " AND LOWER(Type) = 'e' "
  
    Set rsTemp = datGeneral.GetRecords(sSQL)
    With rsTemp
      If Not (.EOF And .BOF) Then
        .MoveFirst
        Do Until .EOF
          iResult = ValidateCalculation(!ColumnID)
          Select Case iResult
            Case REC_SEL_VALID_DELETED
              mstrStatusMessage = "A calculation used in this definition has been deleted by another user."
            Case REC_SEL_VALID_INVALID
              mstrStatusMessage = "A calculation used in this definition is invalid."
            Case REC_SEL_VALID_HIDDENBYOTHER
              If Not gfCurrentUserIsSysSecMgr Then
                mstrStatusMessage = "A calculation used in this definition has been made hidden by another user."
              End If
          End Select
          
          If Len(mstrStatusMessage) > 0 Then
            Exit Do
          End If
          
          .MoveNext
        Loop
      End If
    End With
    
    Set rsTemp = Nothing
  End If
  
  IsRecordSelectionValid = (Len(mstrStatusMessage) = 0)
  
End Function

Private Sub WrdOpenApp()
  
  '                  Word '97     Word 2000
  '---------------------------------------
  'wrdApp.Version    8.0a         9.0
  'wrdApp.Build      8.0.4412     9.0.2717
  
  'wrdapp.MailSystem = wdNoMailSystem
  'wrdapp.MAPIAvailable
  
  On Error GoTo LocalErr

  Set wrdApp = CreateObject("Word.Application")
  'If ASRDEVELOPMENT Then
  '  wrdApp.Visible = True
  'End If
  mblnAppVisible = (GetSystemSetting("Output", "WordVisible", "0") = "1")
  If mblnAppVisible Then
    wrdApp.Visible = True
    wrdApp.DisplayAlerts = wdAlertsAll
  Else
    wrdApp.DisplayAlerts = wdAlertsNone
  End If

  If Val(wrdApp.Version) < 8 Then
    mstrStatusMessage = "You are running a version of Word " & _
                      "which is not supported by HR-Pro"
    fOK = False
    Exit Sub
  End If

Exit Sub

LocalErr:
  mstrStatusMessage = "Error opening Word application"
  fOK = False

End Sub

Private Sub WrdOpenTemplate(blnSuppressPrompt As Boolean)

  Dim strHiddenMenuItems(6) As String
  Dim intCount As Integer
  
  strHiddenMenuItems(0) = "Mail Me&rge..."
  strHiddenMenuItems(1) = "&Error Check Merge"
  strHiddenMenuItems(2) = "&Merge to Document"
  strHiddenMenuItems(3) = "&Merge to Printer"
  strHiddenMenuItems(4) = "&Merge Options"
  strHiddenMenuItems(5) = "&Find in Field"
  strHiddenMenuItems(6) = "&Data Form"
    
  On Error GoTo LocalErr
  
  'It seems to be a good idea to show the word app when
  'opening the template file in case there is a problem
  'with the data source document (in which case, OpenHR
  'would appear to hang!)
  'wrdApp.Visible = True
  

  
  
  'As the data source is no longer referenced by the template
  'I don't think that we need to see the app when opening
  'the template file
  Set wrdDocTemplate = wrdApp.Documents.Open( _
      FileName:=mstrDefTemplateFile, _
      ConfirmConversions:=False, _
      ReadOnly:=False, _
      AddToRecentFiles:=False, _
      Revert:=False)
  
  'MH20001020
  ' Check if the template already has a data source
  If wrdDocTemplate.MailMerge.MainDocumentType <> wdNotAMergeDocument Then
    'mstrStatusMessage = "Error with Word template file <" & mstrDefTemplateFile & ">" & vbCrLf & _
        "This has been saved as merge document"
    mstrStatusMessage = _
        "The Word template file <" & mstrDefTemplateFile & "> " & _
        "has been saved referencing a different set of data" & vbCrLf & vbCrLf & _
        "You need to remove this reference in order to proceed with this mail merge."
        
    If blnSuppressPrompt = False Then
      If COAMsgBox(mstrStatusMessage & vbCrLf & "Proceed ?", vbYesNo + vbExclamation, "Mail Merge") <> vbYes Then
        fOK = False
        Exit Sub
      End If
    End If
  
    wrdDocTemplate.MailMerge.MainDocumentType = wdNotAMergeDocument
    mstrStatusMessage = vbNullString

  End If
    
  
  
  'Set all menu change to be stored in the template file
  wrdApp.CustomizationContext = wrdDocTemplate

  On Error Resume Next
  With wrdDocTemplate.CommandBars("Mail Merge")
    
    For intCount = 0 To UBound(strHiddenMenuItems)
      With .Controls(strHiddenMenuItems(intCount))
        If .Visible = True Then
          .Visible = False
        End If
      End With
    Next
  
  End With
  On Error GoTo LocalErr
  
  'Save menu changes (if applicable)
  If wrdDocTemplate.Saved = False Then
    wrdDocTemplate.Saved = True
  End If

Exit Sub

LocalErr:
  mstrStatusMessage = "Error opening Word template file <" & mstrDefTemplateFile & ">"
  fOK = False

End Sub


Private Sub WrdCreateDataSource()

  Dim intCount As Integer

  On Error GoTo LocalErr

  mstrDataSourceName = GetTmpFName
  mstrDataSourceFormat = GetOfficeSaveAsFormat(".doc", GetOfficeWordVersion, oaWord)
  
  Set wrdDocDataSource = wrdApp.Documents.Add
  wrdDocDataSource.SaveAs FileName:=mstrDataSourceName, FileFormat:=Val(mstrDataSourceFormat), AddToRecentFiles:=False

Exit Sub

LocalErr:
  mstrStatusMessage = "Error creating datasource file <" & mstrDataSourceName & ">" & vbCrLf & _
                      Err.Description
  fOK = False

End Sub


Private Sub WrdPopulateDataSource()

  'This sub converts the data to a table as it was causing
  'a problem when opening a data source with a single field
  
  Const strBookMark As String = "TableStart"
  Dim strOutput As String
  Dim intCount As Integer
  Dim strSQL As String
  Dim strEmailAddr As String
  Dim blnRecordOkay As Boolean
  Dim lngRecordCount As Long

  On Error GoTo LocalErr

  With wrdDocDataSource.ActiveWindow.Selection
  
    .Bookmarks.Add Range:=.Range, Name:=strBookMark
    
    'Add Column headers
    strOutput = vbNullString
    For intCount = 2 To mrsMergeData.Fields.Count
      strOutput = strOutput & _
        IIf(intCount > 2, vbTab, vbNullString) & _
        Trim(Replace(mrsMergeData.Fields(intCount - 1).Name, vbTab, " "))
    Next

    If mlngEmailAddrCalc > 0 Then
      strOutput = strOutput & vbTab & "Email_Address"
    End If

    lngRecordCount = 0
    mrsMergeData.MoveFirst
    .TypeText strOutput
    'If Not mrsMergeData.EOF Then
    '  .TypeParagraph
    'End If

    While Not mrsMergeData.EOF And fOK

      'And record's fields
      lngRecordCount = lngRecordCount + 1
      strOutput = vbNullString
      For intCount = 2 To mrsMergeData.Fields.Count
        'strOutput = strOutput & _
          IIf(intCount > 2, vbTab, vbNullString) & _
          Trim(mrsMergeData.Fields(intCount - 1).Value)
        strOutput = strOutput & _
          IIf(intCount > 2, vbTab, vbNullString) & _
          FormatData(mrsMergeData.Fields(intCount - 1).value, intCount - 1)
      Next

      blnRecordOkay = True
      If mlngEmailAddrCalc > 0 Then
        strEmailAddr = GetEmailAddress(mlngEmailAddrCalc, mrsMergeData.Fields(0).value)
        If Trim(strEmailAddr) = vbNullString Then
          'email is blank
          gobjEventLog.AddDetailEntry _
              GetRecordDesc(mrsMergeData.Fields(0).value) & vbCrLf & vbCrLf & _
              "Email Address is blank"
          blnRecordOkay = False
        
'MH20000314
'Removed valid email address check as per PC
        'ElseIf ValidEmailAddress(strEmailAddr) = False Then
        '  'error with email address
        '  gobjEventLog.AddDetailEntry _
        '      GetRecordDesc(mrsMergeData.Fields(0).Value) & vbCrLf & vbCrLf & _
        '      "Email Address is invalid <" & strEmailAddr & ">"
        '  blnRecordOkay = False
        Else
          strOutput = strOutput & vbTab & strEmailAddr
        End If
      End If

      mrsMergeData.MoveNext
      
      'MH20000620
      'Need to check for null output in case you are merging a single
      'record, single column and the data in that column is blank!
      '(This used to result in the data source table containing a
      'single cell which contains the column header, which in turn,
      'will result in an error opening the data source document).
      
      If blnRecordOkay Then
        mlngSuccessCount = mlngSuccessCount + 1
        .TypeParagraph
        .TypeText IIf(strOutput <> vbNullString, strOutput, " ")
      Else
        mlngFailCount = mlngFailCount + 1
      End If

      If lngRecordCount Mod 1000 = 0 Then
        If gobjProgress.Cancelled Then
          mblnUserCancelled = True
          fOK = False
        End If
      End If

    Wend
  
    If mlngSuccessCount = 0 Then
      mstrStatusMessage = "No records have a valid email address."
      fOK = False
    End If
  
  
    'If only one column is selected, the system will prompt
    'for a deliminator at the time of opening the data source.
    'This can be avoided by putting the data into a table.
    
    'need to do this regardless now as error occured
    'when the data contained a comma (also now use tabs as delim)!

    'MH20000620
    'Okay, so now we get an error converting to table if more than
    'about sixty columns/calculations are selected. So only convert
    'to a table if one column selected and use tabs as delim to avoid
    'problems when the data contains commas.

    If fOK Then
      If mrsMergeData.Fields.Count < 3 Then
        
        'go to start of table highlight to end of document
        .GoTo What:=wdGoToBookmark, Name:=strBookMark
        .EndKey Unit:=wdStory, Extend:=wdExtend
      
        'convert selected text into a table
        .ConvertToTable _
            Separator:=wdSeparateByTabs, _
            Format:=wdTableFormatNone, _
            ApplyFont:=False, _
            ApplyColor:=False, _
            AutoFit:=False
      
      End If
    End If

  End With

  wrdDocDataSource.SaveAs FileFormat:=Val(mstrDataSourceFormat)
  wrdDocDataSource.Close SaveChanges:=False

Exit Sub

LocalErr:
  mstrStatusMessage = "Error populating data source document" & vbCrLf & "(" & Err.Description & ")"
  fOK = False

End Sub


Private Sub WrdPauseBeforeMerge()

  Dim strMBText As String
  Dim intMBButtons As Long
  Dim strMBTitle As String
  Dim intMBResponse As Integer
  
  On Error GoTo LocalErr

  Screen.MousePointer = vbDefault
  wrdDocTemplate.Saved = True

  'This is the section which controls the "pause before merge"
  If mblnDefPauseBeforeMerge = True Then
  
    gobjProgress.Visible = False
    
    '1. Ask the user if they want to amend the document and say "DON'T CLOSE WORD !"
    strMBText = "Would you like to amend the document and/or merge fields?" '& vbCrLf & vbCrLf & _
                "IMPORTANT NOTE: If amending the document, do not close Microsoft Word"
    intMBButtons = vbQuestion + vbYesNoCancel
    strMBTitle = IIf(mbIsLabel, "Envelopes & Labels", "Mail Merge")
    
    intMBResponse = COAMsgBox(strMBText, intMBButtons, strMBTitle)

    If intMBResponse = vbCancel Then
      mblnUserCancelled = True
      fOK = False
      Exit Sub
    End If


    If intMBResponse = vbYes Then

      '2. Activate Word
      wrdApp.Visible = True
      DoEvents
      wrdApp.WindowState = wdWindowStateNormal
      wrdApp.Activate


      '3. Wait for the user to click OK or CANCEL
      strMBText = "Click OK to proceed with the " & IIf(mbIsLabel, "label", "mail") & " merge."
      intMBButtons = vbInformation + vbOKCancel
      strMBTitle = IIf(mbIsLabel, "Envelopes & Labels", "Mail Merge")
      intMBResponse = COAMsgBox(strMBText, intMBButtons, strMBTitle)

      If intMBResponse = vbCancel Then
        mblnUserCancelled = True
        fOK = False
        Exit Sub
      End If
    
    End If
    
    
    gobjProgress.Visible = True

  End If
  
  
CheckMergeFieldExists:

  'If there are no merge fields then force "pause before merge"
  'If its a batch job then this will produce an error and be trapped
  If Not MergeFieldExists Then
    gobjProgress.Visible = False
    
    strMBText = "There are no merge fields specified in the template document." & vbCrLf & vbCrLf & _
                "Would you like to switch to word to insert merge fields now?"
    intMBButtons = vbQuestion + vbOKCancel
    strMBTitle = "Mail Merge"
    intMBResponse = COAMsgBox(strMBText, intMBButtons, strMBTitle)
    
    If intMBResponse = vbCancel Then
      mblnUserCancelled = True
      fOK = False
      Exit Sub
    End If
    
    wrdApp.Visible = True
    Do While Not MergeFieldExists
      
      wrdApp.Activate
      
      strMBText = "Please insert merge fields on the template document then click OK to proceed"
      intMBButtons = vbInformation + vbOKCancel
      strMBTitle = "Mail Merge"
      intMBResponse = COAMsgBox(strMBText, intMBButtons, strMBTitle)
    
      If intMBResponse = vbCancel Then
        mblnUserCancelled = True
        fOK = False
        Exit Sub
      End If
      
    Loop
      
    gobjProgress.Visible = True
  End If
  
Exit Sub

LocalErr:
  Select Case Err.Number
  Case 462, -2147417848
    '462
    'The remote server machine does not exist or is unavailable
    'mstrStatusMessage = "Microsoft Word has been closed by the user"
    
    '-2147417848
    'Automation Error
    'The object invoked has disconnected from its clients.

    If Err.Number = 462 Then Call WrdOpenApp   'Re-open word so that we can sort out the data source
    If fOK Then Call WrdOpenTemplate(True)
    If fOK Then Call WrdSetMergeOptions

    wrdDocTemplate.Saved = True
    Resume CheckMergeFieldExists

  Case Else
    If Trim(Err.Description) <> vbNullString Then
      mstrStatusMessage = Err.Description
    Else
      mstrStatusMessage = "Error during pause prior merge"
    End If
    fOK = False
  
  End Select

End Sub


Private Sub WrdCheckIfSaveTemplate()
    
  Dim strMBText As String
  Dim intMBButtons As Long
  Dim strMBTitle As String
  Dim intMBResponse As Integer
  
  Dim blnTemplateReadOnly As Boolean
  
  On Error GoTo LocalErr

  wrdApp.DisplayAlerts = wdAlertsAll
  
  blnTemplateReadOnly = False
  intMBResponse = vbYes
  Do While wrdDocTemplate.Saved = False And intMBResponse = vbYes

    gobjProgress.Visible = False

    strMBText = "You have not saved changes to the template document, would you like to save changes now?"
    intMBButtons = vbQuestion + vbYesNo
    strMBTitle = "Mail Merge"
    intMBResponse = COAMsgBox(strMBText, intMBButtons, strMBTitle)

    If intMBResponse = vbYes Then
      wrdApp.Visible = True
      wrdApp.DisplayAlerts = wdAlertsAll
      wrdDocTemplate.SaveAs FileFormat:=Val(mstrDefTemplateFormat)
      
      If blnTemplateReadOnly Then
        strMBText = "This template file is currently in use by another user.  Please save as a new file."
        intMBButtons = vbInformation
        strMBTitle = "Mail Merge"
        intMBResponse = COAMsgBox(strMBText, intMBButtons, strMBTitle)
        wrdDocTemplate.SaveAs FileFormat:=Val(mstrDefTemplateFormat)
        blnTemplateReadOnly = False
        intMBResponse = vbYes
      End If

      wrdApp.Visible = mblnAppVisible
      wrdApp.DisplayAlerts = wdAlertsNone
    End If

  Loop

  gobjProgress.Visible = True
  
  wrdApp.Visible = mblnAppVisible
  
  ' JDM - 25/06/03 - Fault 6092 - Commented out line below - was messing up the defsel screen
  'Screen.MousePointer = vbHourglass

Exit Sub

LocalErr:
  Select Case Err.Number
  Case 462, -2147417848
    '462
    'The remote server machine does not exist or is unavailable
    'mstrStatusMessage = "Microsoft Word has been closed by the user"
    
    '-2147417848
    'Automation Error
    'The object invoked has disconnected from its clients.

    If Err.Number = 462 Then Call WrdOpenApp   'Re-open word so that we can sort out the data source
    If fOK Then Call WrdOpenTemplate(True)
    If fOK Then Call WrdSetMergeOptions
    wrdDocTemplate.Saved = True

  Case 4198, 5155
    '4198 = Command failed
    '5155 = This file is read-only.
    'Both of these are when the user clicked cancel on "save as" dialog so ask again if want to save
    Err.Clear
    On Error GoTo LocalErr
    Resume Next

  Case 5356
    'Word cannot save this file because it is already open elsewhere.
    blnTemplateReadOnly = True
    Err.Clear
    On Error GoTo LocalErr
    Resume Next

  Case Else
    If Trim(Err.Description) <> vbNullString Then
      mstrStatusMessage = Err.Description
    Else
      mstrStatusMessage = "Error checking if template has been saved"
    End If
    fOK = False
  
  End Select
  
End Sub
  

Private Sub WrdValidateTemplateMergeFields(blnAbortOnError As Boolean)

  'Loop through all of the fields on the template
  'try to find them in the data source
  'if not found then prompt user and
  'delete field from template is required

  Dim strMBText As String
  Dim intMBButtons As Long
  Dim strMBTitle As String
  Dim intMBResponse As Integer
  
  Dim strTemplateFieldName As String
  Dim intCount As Integer
  Dim intRecordsetField As Integer
  Dim strTemp As String
  Dim blnFieldOkay As Boolean

  On Error GoTo LocalErr
  
  With mrsMergeData
    
    intCount = 1
    Do While intCount <= wrdDocTemplate.Fields.Count

      strTemplateFieldName = Trim$(wrdDocTemplate.Fields(intCount).Code)
      If Left$(strTemplateFieldName, 10) = "MERGEFIELD" Then
          
        strTemplateFieldName = Mid$(strTemplateFieldName, 12)
        
''        'MH20011004 Fault 2918
''        If Left(strTemplateFieldName, 1) & Right(strTemplateFieldName, 1) = String(2, 34) Then
''          'enclosed in quotes on office xp...
''          strTemplateFieldName = Mid(strTemplateFieldName, 2, Len(strTemplateFieldName) - 2)
''        End If
''
''
''        blnFieldOkay = False
''
''        'Now that we've got the name of the merge field
''        'try and find it in the data soure
''        For intRecordsetField = 0 To mrsMergeData.Fields.Count - 1
''
''          'MH20011121 Changed to 39 for Office XP compatibility...
''          'If UCase$(Left(Trim$(mrsMergeData.Fields(intRecordsetField).Name), 40)) = UCase$(strTemplateFieldName) Then
''          If UCase$(Left(Trim$(mrsMergeData.Fields(intRecordsetField).Name), 39)) = UCase$(Left(strTemplateFieldName, 39)) Then
''            blnFieldOkay = True
''            Exit For
''          End If
''        Next

        'MH20030205 Faults 4374 & 4578
        strTemplateFieldName = Replace(strTemplateFieldName, Chr(34), vbNullString)
        strTemplateFieldName = Replace(strTemplateFieldName, " ", "_")
        strTemplateFieldName = UCase(Left(strTemplateFieldName, 39))

        blnFieldOkay = False
        
        'Now that we've got the name of the merge field
        'try and find it in the data soure
        For intRecordsetField = 0 To mrsMergeData.Fields.Count - 1

          strTemp = mrsMergeData.Fields(intRecordsetField).Name
          strTemp = Replace(strTemp, Chr(34), vbNullString)
          strTemp = Replace(strTemp, " ", "_")
          strTemp = UCase(Left(strTemp, Len(strTemplateFieldName)))

          If strTemp = strTemplateFieldName Then
            blnFieldOkay = True
            Exit For
          End If
        Next


        If mlngEmailAddrCalc > 0 And UCase(strTemplateFieldName) = UCase("Email_Address") Then
          blnFieldOkay = True
        End If
        
        If blnFieldOkay = False Then
          strTemp = "The Template Document '" & mstrDefTemplateFile & "' " & _
                    "contains the merge field '" & strTemplateFieldName & "' " & _
                    "which was not selected in the mail merge definition." & vbCrLf

          If blnAbortOnError Then
            mstrStatusMessage = strTemp
            fOK = False
            Exit Sub
          End If
          

          'If not abort on error (i.e. not batch mode) then
          'prompt the user as to what to do next

          strMBText = strTemp & vbCrLf & _
                      "Select OK to proceed with the merge and remove this field from the template or " & _
                      "select Cancel to abort this Mail Merge."
          intMBButtons = vbExclamation + vbOKCancel
          strMBTitle = "Mail Merge"
          
          gobjProgress.Visible = False
          intMBResponse = COAMsgBox(strMBText, intMBButtons, strMBTitle)
          gobjProgress.Visible = True

          If intMBResponse = vbOK Then
            wrdDocTemplate.Fields(intCount).Delete
            wrdDocTemplate.Saved = False
            'intCount not increment because all field indexes following
            'this field will be reduced by one following this delete
            'therefore need to take one off the loop count so that we
            'don't skip the next field !
            intCount = intCount - 1
          Else
            mstrStatusMessage = strTemp
            fOK = False
            Exit Sub
          End If
          
        End If
      End If

      intCount = intCount + 1
    Loop
  End With
    
Exit Sub

LocalErr:
  mstrStatusMessage = "Error reading template merge fields"
  fOK = False

End Sub


Private Sub WrdSetMergeOptions()
  
  Dim objTemp As Object
  
  On Error GoTo LocalErr
  
  With wrdDocTemplate.MailMerge

    .OpenDataSource Name:=mstrDataSourceName, AddToRecentFiles:=False
    .SuppressBlankLines = mblnDefSuppressBlankLines
    
    If mintOutputFormat = 1 Then     'Email
      .Destination = wdSendToEmail
      .MailAddressFieldName = "Email_Address"
      .MailSubject = mstrDefEMailSubject
      .MailAsAttachment = mblnDefEMailAttachment Or mbIsLabel
      
      If Val(wrdApp.Version) >= 12 Then
        .DataSource.FirstRecord = wdDefaultFirstRecord
        .DataSource.LastRecord = wdDefaultLastRecord
        
        'MH20100726 Email from Microsoft (Bill Goldade):
        '
        'The Word 2010 help topic for MailFormat (below) states that
        '"The MailFormat property is ignored if the MailAsAttachment property is set to True."
        'However, it appears that Word 2010 does not ignore the property as per the help topic.
        'Please try to run your macro in 2010 where you comment out or remove the Mail.Merge.MailFormat property from the code.
        
        If Not mblnDefEMailAttachment And Not mbIsLabel Then
          Set objTemp = wrdDocTemplate.MailMerge
          objTemp.MailFormat = GetSystemSetting("Mail Merge", "Email Format", 1)
          Set objTemp = Nothing
        End If
      End If
    Else
      'If destination = printer then send to new document
      'and print.  This avoids printer selection !
      .Destination = wdSendToNewDocument
    End If

    '.ViewMailMergeFieldCodes = (mlngSingleRecordID = 0)

  End With

Exit Sub

LocalErr:
  mstrStatusMessage = "Error setting merge destination/options"
  fOK = False

End Sub
  
  
Private Sub WrdRunMailMerge()

  Dim strTempFileName As String
  Dim strOriginalDefaultPrinter As String

  On Error GoTo LocalErr

  If Not MergeFieldExists Then
    mstrStatusMessage = "No merge fields specified in template document <" & mstrDefTemplateFile & ">"
    fOK = False
    Exit Sub
  End If


  wrdApp.DisplayAlerts = wdAlertsAll
  
  
  'Rename attachment
  If Trim(mstrDefAttachmentName) <> vbNullString Then
    
    'Get temp path
    strTempFileName = Space(1024)
    Call GetTempPath(1024, strTempFileName)
    strTempFileName = Left(strTempFileName, InStr(strTempFileName, Chr(0)) - 1)

    'temp path + "\" + attachment file name
    mstrDefAttachmentName = strTempFileName & _
      IIf(Right(strTempFileName, 1) <> "\", "\", "") & _
      mstrDefAttachmentName

    If Dir(mstrDefAttachmentName) <> vbNullString Then
      Kill mstrDefAttachmentName
    End If
    
    wrdDocTemplate.SaveAs FileName:=mstrDefAttachmentName, FileFormat:=Val(mstrDefAttachmentFormat)

  End If



  'Pause:=False to report errors in a new document.
  '(Pause:=True to have Word pause and display a troubleshooting dialog box if a mail merge error is found)
  If mintOutputFormat = 1 Then    'Email
    gobjProgress.Visible = False
  End If
  wrdDocTemplate.MailMerge.Execute Pause:=False
  gobjProgress.Visible = True
  
  
  Set wrdDocOutput = wrdApp.ActiveDocument
  wrdDocOutput.ActiveWindow.View.Type = wdNormalView


  If mblnOutputSave Then
    If Dir$(mstrOutputFileName) <> vbNullString Then
      Kill mstrOutputFileName
    End If
    wrdDocOutput.SaveAs FileName:=mstrOutputFileName, FileFormat:=Val(mstrOutputFileFormat)
    'Wait for the save to finish !
    Do While wrdApp.BackgroundSavingStatus > 0
      DoEvents
    Loop
  End If


  If mblnOutputPrinter Then
    If Trim(mstrOutputPrinterName) <> "" Then
      strOriginalDefaultPrinter = wrdApp.ActivePrinter
      wrdApp.ActivePrinter = mstrOutputPrinterName
    End If

    gobjProgress.Visible = False
    wrdApp.Visible = True
    wrdDocOutput.PrintOut

    'wrdDocOutput.PrintOut Range:=wdPrintAllDocument, Item:= _
        wdPrintDocumentContent, Copies:=1, Pages:="", PageType:=wdPrintAllPages, _
        ManualDuplexPrint:=False, Collate:=True, Background:=True, PrintToFile:= _
        False, PrintZoomColumn:=0, PrintZoomRow:=0, PrintZoomPaperWidth:=0, _
        PrintZoomPaperHeight:=0
    
    'Wait for the print job to finish !
    Do While wrdApp.BackgroundPrintingStatus > 0
      DoEvents
    Loop
    
    
    wrdApp.Visible = mblnAppVisible
    gobjProgress.Visible = True
  
    If Trim(strOriginalDefaultPrinter) <> "" Then
      wrdApp.ActivePrinter = strOriginalDefaultPrinter
    End If
  
  End If

Exit Sub

LocalErr:
  Select Case Err.Number
  Case 462
    'The remote server machine does not exist or is unavailable
    mstrStatusMessage = "Microsoft Word has been closed by the user"
    Call WrdOpenApp   'Re-open word so that we can sort out the data source
  
  Case -2147417848
    'Automation Error
    'The object invoked has disconnected from its clients.
    mstrStatusMessage = "Microsoft Word document has been closed by the user"

  Case 4605
    'This method or property is not available because the current mail merge main document has no merge fields.
    mstrStatusMessage = "No merge fields specified in template document <" & mstrDefTemplateFile & ">"

  Case 5152
    'This is not a valid file name.
    'Try one or more of the following:
    '* Check the path to make sure it was typed correctly.
    '* Select a file from the list of files and folders.
    'NHRD18082004 Fault 7994 Removed Mail Merge as this message is used for Env & Labs too.
    mstrStatusMessage = "Error Saving to file <" & mstrOutputFileName & ">" & vbCrLf & _
                      "Please ensure that the output document path is correct within this definition"

  Case 5630
    'Word cannot merge documents that can be distributed by mail or fax without a valid mail address. Choose the Setup button to select a mail address data field.
    mstrStatusMessage = "Email field <" & mstrDefEmailCol & "> is not specified in merge data"
  
  Case 70
    ' Permission denied - file is either open, or we do not have security to it.
    mstrStatusMessage = "OpenHR cannot get access to the output file." & vbCrLf & "Ensure that no other applications have this file open."
  
  Case Else
    If Trim(Err.Description) <> vbNullString Then
      mstrStatusMessage = Err.Description
    Else
      mstrStatusMessage = "Error sending merge to output destination"
    End If
  
  End Select
  
  'Mail not open/installed ??????
  fOK = False
  
End Sub


Private Sub TidyUpAndExit()

  Dim wrdNewTable As Word.Table
  Dim strSQL As String
  Dim blnCloseWord As Boolean
  
  On Error Resume Next
  
  Err.Clear
  
  If Not gblnBatchMode Then
    gobjProgress.CloseProgress
    DoEvents
  Else
    gobjProgress.ResetBar2
  End If

  Err.Clear
  wrdApp.DisplayAlerts = wdAlertsNone

  'MH20010709
  If Err.Number <> 0 Then
    'Problem with wrdapp object, could be closed so recreate it.
    Set wrdApp = CreateObject("Word.Application")
    wrdApp.DisplayAlerts = wdAlertsNone
  End If

  'Make sure that the template does not still reference the data source
  '(close it and open it incase there have been changes to it or that
  ' the template is now the temporary email attachment name thingy !)
  If Not mbIsLabel And mstrDefTemplateFormat <> vbNullString Then
    wrdDocTemplate.Close SaveChanges:=False
    Set wrdDocTemplate = wrdApp.Documents.Open( _
        FileName:=mstrDefTemplateFile, _
        ConfirmConversions:=False, _
        ReadOnly:=False, _
        AddToRecentFiles:=False, _
        Revert:=False)
  
    wrdDocTemplate.MailMerge.MainDocumentType = wdNotAMergeDocument
    wrdDocTemplate.Saved = False    'MH20070124 Fault 11824 - Force save for Office 2007
    wrdDocTemplate.SaveAs FileFormat:=Val(mstrDefTemplateFormat)
  End If
  
  wrdDocTemplate.Close SaveChanges:=False
  Set wrdDocTemplate = Nothing


  'If the temporary email attachment name exists then kill it
  If mstrDefAttachmentName <> vbNullString And Dir(mstrDefAttachmentName) <> vbNullString Then
    Kill mstrDefAttachmentName
  End If

'''''''  ' We are not using the original template - kill this fake
'''''''  If mlngDocumentMapID > 0 Then
'''''''    Kill mstrDefTemplateFile
'''''''  End If

  'Kill the data source
  wrdDocDataSource.Close SaveChanges:=False
  Kill mstrDataSourceName
  Set wrdDocDataSource = Nothing


  'blnCloseWord = (mblnDefCloseDoc Or _
                 (mintDefOutput <> wdSendToNewDocument) Or _
                 (fOK = False))
  blnCloseWord = (mblnOutputScreen = False Or fOK = False)


  If blnCloseWord Then
    wrdDocOutput.Close SaveChanges:=False
    wrdApp.Quit
  Else
    wrdApp.Visible = True
    If Not gblnBatchMode Then
      DoEvents
      wrdApp.WindowState = wdWindowStateNormal
      wrdApp.Activate
      
      'MH20010307 Fault 1681
      'If word 2000 then also need to activate the document !
      wrdDocOutput.Activate
    End If
  End If

  mrsMailMergeColumns.Close
  Set mrsMailMergeColumns = Nothing

  mrsMergeData.Close
  Set mrsMergeData = Nothing
  
  Set wrdDocOutput = Nothing
  Set wrdApp = Nothing
  
  On Error GoTo 0

End Sub


Private Function GetTableAndColumnName(lngColumnID As Long) As String

  Dim strSQL As String
  Dim rsTemp As Recordset

  strSQL = "SELECT ColumnName, TableName " & _
           "FROM ASRSYSColumns " & _
           "JOIN ASRSYSTables ON (ASRSYSColumns.TableID = ASRSYSTables.TableID)" & _
           "WHERE ColumnID = " & CStr(lngColumnID) & " "
  Set rsTemp = mdatData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)

  fOK = Not (rsTemp.BOF And rsTemp.EOF)
  If fOK Then
    GetTableAndColumnName = rsTemp!TableName & "_" & rsTemp!ColumnName
  End If

  rsTemp.Close
  Set rsTemp = Nothing

End Function


Private Function GetPicklistFilterSelect() As String

  Dim rsTemp As Recordset

  On Error GoTo LocalErr
  
  
  'If mlngSingleRecordID > 0 Then
  '  GetPicklistFilterSelect = CStr(mlngSingleRecordID)

  If mstrPicklistFilterIDs <> "" Then
    GetPicklistFilterSelect = mstrPicklistFilterIDs
  
  
  ElseIf mlngDefPickListID > 0 Then
    
    mstrStatusMessage = IsPicklistValid(mlngDefPickListID)
    If mstrStatusMessage <> vbNullString Then
      fOK = False
      Exit Function
    End If
    
    
    'Get List of IDs from Picklist
    Set rsTemp = mdatData.OpenRecordset("EXEC sp_ASRGetPickListRecords " & mlngDefPickListID, adOpenForwardOnly, adLockReadOnly)
    fOK = Not (rsTemp.BOF And rsTemp.EOF)

    If Not fOK Then
      mstrStatusMessage = "The base table picklist contains no records."
    Else
      GetPicklistFilterSelect = vbNullString
      Do While Not rsTemp.EOF
        GetPicklistFilterSelect = GetPicklistFilterSelect & _
            IIf(Len(GetPicklistFilterSelect) > 0, ", ", "") & rsTemp.Fields(0)
        rsTemp.MoveNext
      Loop
    End If

    rsTemp.Close
    Set rsTemp = Nothing

  ElseIf mlngDefFilterID > 0 Then
    
    mstrStatusMessage = IsFilterValid(mlngDefFilterID)
    If mstrStatusMessage <> vbNullString Then
      fOK = False
      Exit Function
    End If
    
    'Get list of IDs from Filter
    fOK = datGeneral.FilteredIDs(mlngDefFilterID, GetPicklistFilterSelect)
    
    ' Generate any UDFs that are used in this filter
    If fOK Then
      datGeneral.FilterUDFs mlngDefFilterID, mastrUDFsRequired()
    End If

    If Not fOK Then
      ' Permission denied on something in the filter.
      mstrStatusMessage = "You do not have permission to use the '" & datGeneral.GetFilterName(mlngDefFilterID) & "' filter."
    End If

  End If

Exit Function

LocalErr:
  mstrStatusMessage = "Error processing picklist"
  fOK = False

End Function


Private Sub SQLGetMergeColumns()
  
  Dim strSQL As String
  Dim strSQL1 As String
  Dim strSQL2 As String
  Dim strSQL3 As String
  Dim strOrderBy As String

  On Error GoTo LocalErr

  'Merge Column Types
  '
  ' "C" is a column which has been selected by the user
  ' "E" is an express which has been selected by the user
  ' "X" is a system column required by the merge
  '     (currently only used for the email field where required)
  
  strSQL1 = "SELECT 'ColExp'          = 'Col',                                  " & vbCrLf & _
            "       'ColExpId'        = ASRSysColumns.ColumnID,                 " & vbCrLf & _
            "       'TableID'         = ASRSysTables.TableID,                   " & vbCrLf & _
            "       'Table'           = ASRSysTables.Tablename,                 " & vbCrLf & _
            "       'Name'            = ASRSysColumns.ColumnName,               " & vbCrLf & _
            "       'Type'            = ASRSysColumns.DataType,                 " & vbCrLf & _
            "       'Size'            = ASRSysMailMergeColumns.Size,            " & vbCrLf & _
            "       'Decimals'        = ASRSysMailMergeColumns.Decimals,        " & vbCrLf & _
            "       'StartOnNewLine'  = ASRSysMailMergeColumns.StartOnNewLine,  " & vbCrLf & _
            "       'HeadingText'     = ASRSysMailMergeColumns.HeadingText,     " & vbCrLf & _
            "       'ColumnType'      = ASRSysMailMergeColumns.Type,            " & vbCrLf & _
            "       'ColumnOrder'     = ASRSysMailMergeColumns.ColumnOrder      " & vbCrLf & _
            "FROM ASRSysMailMergeColumns " & vbCrLf & _
            "JOIN ASRSysColumns ON (ASRSysColumns.ColumnID = ASRSysMailMergeColumns.ColumnID) " & vbCrLf & _
            "JOIN ASRSysTables ON (ASRSysTables.TableID = ASRSysColumns.TableID) " & vbCrLf & _
            "WHERE ASRSysMailMergeColumns.Type = 'C' " & _
            "  AND ASRSysMailMergeColumns.MailMergeID = " & CStr(mlngMailMergeID) & " "

            '"WHERE ASRSysMailMergeColumns.Type <> 'E' " & vbCrLf &

  'MH20000906 Added the words "LEFT OUTER" so that we can pick up invalid calcs.
  strSQL2 = "SELECT 'ColExp'          = 'Exp',                                  " & vbCrLf & _
            "       'ColExpId'        = ASRSysExpressions.ExprID,               " & vbCrLf & _
            "       'TableID'         = 0,                                      " & vbCrLf & _
            "       'Table'           = 'Calculation_',                         " & vbCrLf & _
            "       'Name'            = ASRSysExpressions.Name,                 " & vbCrLf & _
            "       'Type'            = ASRSysExpressions.ReturnType,           " & vbCrLf & _
            "       'Size'            = ASRSysMailMergeColumns.Size,            " & vbCrLf & _
            "       'Decimals'        = ASRSysMailMergeColumns.Decimals,        " & vbCrLf & _
            "       'StartOnNewLine'  = ASRSysMailMergeColumns.StartOnNewLine,  " & vbCrLf & _
            "       'HeadingText'     = ASRSysMailMergeColumns.HeadingText,     " & vbCrLf & _
            "       'ColumnType'      = ASRSysMailMergeColumns.Type,            " & vbCrLf & _
            "       'ColumnOrder'     = ASRSysMailMergeColumns.ColumnOrder      " & vbCrLf & _
            "FROM ASRSysMailMergeColumns " & vbCrLf & _
            "LEFT OUTER JOIN ASRSysExpressions ON (ASRSysExpressions.ExprID = ASRSysMailMergeColumns.ColumnID) " & vbCrLf & _
            "WHERE ASRSysMailMergeColumns.Type = 'E' " & vbCrLf & _
            "  AND ASRSysMailMergeColumns.MailMergeID = " & CStr(mlngMailMergeID) & " "
 
  ' Headings & Seperators
  strSQL3 = "SELECT 'ColExp'          = 'Col',                                  " & vbCrLf & _
            "       'ColExpId'        = 0,                                      " & vbCrLf & _
            "       'TableID'         = 0,                                      " & vbCrLf & _
            "       'Table'           = '',                                     " & vbCrLf & _
            "       'Name'            = '',                                     " & vbCrLf & _
            "       'Type'            = 0,                                      " & vbCrLf & _
            "       'Size'            = ASRSysMailMergeColumns.Size,            " & vbCrLf & _
            "       'Decimals'        = ASRSysMailMergeColumns.Decimals,        " & vbCrLf & _
            "       'StartOnNewLine'  = ASRSysMailMergeColumns.StartOnNewLine,  " & vbCrLf & _
            "       'HeadingText'     = ASRSysMailMergeColumns.HeadingText,     " & vbCrLf & _
            "       'ColumnType'      = ASRSysMailMergeColumns.Type,            " & vbCrLf & _
            "       'ColumnOrder'     = ASRSysMailMergeColumns.ColumnOrder      " & vbCrLf & _
            "FROM ASRSysMailMergeColumns " & vbCrLf & _
            "WHERE (ASRSysMailMergeColumns.Type = 'H' " & _
            "       OR ASRSysMailMergeColumns.Type = 'S') " & vbCrLf & _
            "  AND ASRSysMailMergeColumns.MailMergeID = " & CStr(mlngMailMergeID) & " "
  
  'strSQL3 = "SELECT 'ColExp'   = 'Exp',                             " & vbCrLf & _
            "       'ColExpId' = ASRSysMailMergeName.EmailAddrID,   " & vbCrLf & _
            "       'TableID'  = 0,                                 " & vbCrLf & _
            "       'Table'    = 'EmailAddr_',                      " & vbCrLf & _
            "       'Name'     = ASRSYSEmailAddress.Name,           " & vbCrLf & _
            "       'Type'     = 0,                                 " & vbCrLf & _
            "       'Size'     = 8000                               " & vbCrLf & _
            "FROM ASRSYSMailMergeName " & vbCrLf & _
            "JOIN ASRSYSEmailAddress ON (ASRSYSEmailAddress.EmailID = ASRSysMailMergeName.EmailAddrID) " & vbCrLf & _
            "WHERE ASRSysMailMergeName.MailMergeID = " & CStr(mlngMailMergeID) & _
            " AND ASRSysMailMergeName.EmailAddrID > 0"
  
  If mbIsLabel Then
    strOrderBy = "ORDER BY 'ColumnOrder'"
  Else
    strOrderBy = "ORDER BY 'ColExp', 'Table', 'Name'"
  End If
  
  strSQL = strSQL1 & vbCrLf & vbCrLf & _
           "UNION" & vbCrLf & vbCrLf & _
           strSQL2 & vbCrLf & vbCrLf & _
           "UNION" & vbCrLf & vbCrLf & _
           strSQL3 & vbCrLf & vbCrLf & _
           strOrderBy
  
  Set mrsMailMergeColumns = mdatData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)

Exit Sub

LocalErr:
  mstrStatusMessage = "Error reading calculation/column definitions."
  fOK = False

End Sub
  

Private Sub SQLCodeCreate()

  Dim strPicklistFilterSelect As String
  Dim objExpr As DataMgr.clsExprExpression
  Dim intIndex As Integer
  Dim objTableView As CTablePrivilege
  Dim lngRecordCount As Integer
  
  fOK = True

  On Error GoTo LocalErr

  ReDim mlngTableViews(2, 0)
  ReDim asViews(0)
  
  intIndex = 0
  'ReDim mintType(intIndex) As Integer
  'ReDim mlngSize(intIndex) As Integer
  'ReDim ambStartOnNewLine(intIndex) As Boolean
  'ReDim amColumnType(intIndex) As String
  'ReDim amHeadingText(intIndex) As String
  'ReDim amColumnInformation(intIndex, 1) As Variant      '0 - Start on New Line
                                                         '1 - Use Thousand Seperator
  ' JPD20030219 Fault 5069
  ' Check the user has permission to read the base table.
  fOK = False
  For Each objTableView In gcoTablePrivileges.Collection
    If (objTableView.TableID = mlngDefBaseTableID) And _
      (objTableView.AllowSelect) Then
      fOK = True
      Exit For
    End If
  Next objTableView
  Set objTableView = Nothing
    
  If Not fOK Then
    mstrStatusMessage = "You do not have permission to read the base table" & vbCrLf & "either directly or through any views."
    Exit Sub
  End If

  mstrSQLFrom = gcoTablePrivileges.item(mstrDefBaseTable).RealSource
  mstrSQLSelect = mstrSQLFrom & ".ID"

  With mrsMailMergeColumns
    
    fOK = Not (.BOF And .EOF)
    If fOK Then
            
      ' Do a record count (.recordcount does not work)
      Do While Not .EOF
        lngRecordCount = lngRecordCount + 1
        .MoveNext
      Loop
      .Requery
  
      ReDim Preserve mintType(lngRecordCount)
      ReDim Preserve mlngSize(lngRecordCount)
      ReDim Preserve mintDecimals(lngRecordCount)
      'ReDim Preserve ambStartOnNewLine(intIndex)
      ReDim Preserve amColumnType(lngRecordCount)
      ReDim Preserve amHeadingText(lngRecordCount)
      ReDim Preserve amColumnInformation(lngRecordCount, 1)
  
      Do While Not .EOF


        '01/08/2001 MH Fault 2125
        intIndex = intIndex + 1
        
        mlngSize(intIndex) = IIf(IsNull(!Size), 0, !Size)
        mintDecimals(intIndex) = IIf(IsNull(!Decimals), 0, !Decimals)
        'ambStartOnNewLine(intIndex) = IIf(IsNull(!StartOnNewLine), True, !StartOnNewLine)
        amColumnType(intIndex) = !ColumnType
        amHeadingText(intIndex) = IIf(IsNull(!HeadingText), "", !HeadingText)
        amColumnInformation(intIndex, 0) = IIf(IsNull(!StartOnNewLine), True, !StartOnNewLine)
        amColumnInformation(intIndex, 1) = datGeneral.DoesColumnUseSeparators(!ColExpId)
        
        Select Case !ColumnType
        Case sTYPECODE_COLUMN
          Call SQLAddColumn(mstrSQLSelect, !TableID, !Table, !Name, !Table & "_" & !Name)
          mintType(intIndex) = !Type
              
        Case sTYPECODE_HEADING
          mstrSQLSelect = mstrSQLSelect & ", '" & IIf(IsNull(!HeadingText), "", Replace(!HeadingText, "'", "''")) & "' AS " & sTYPECODE_HEADING & Trim(Str(intIndex))
          mintType(intIndex) = !Type
      
        Case sTYPECODE_SEPARATOR
          mstrSQLSelect = mstrSQLSelect & ", '" & IIf(IsNull(!HeadingText), "", !HeadingText) & "' AS " & sTYPECODE_HEADING & Trim(Str(intIndex))
          mintType(intIndex) = !Type

        Case sTYPECODE_EXPRESSION
          If IsNull(!Name) Then
            mstrStatusMessage = _
              "This definition contains one or more calculation(s) which" & vbCrLf & _
              "have been deleted by another user."
              fOK = False
              Exit Sub
              
          'TM20010807 Fault 2656 - Check if calc has hidden components.
          ElseIf IsCalcValid(!ColExpId) <> vbNullString Then
              mstrStatusMessage = "You cannot run this Mail Merge definition as it contains one or more calculation(s) which have been deleted or made hidden by another user." & vbCrLf & _
                     "Please re-visit your definition to remove the hidden calculation(s)." & vbCrLf
              fOK = False
              Exit Sub
          
          Else
            
            ' JDM - Fault 8763 - 24/06/2004 - Problem with spaces in expression names
            Call SQLAddCalculation(!ColExpId, !Table & Replace(!Name, " ", "_"))

            Set objExpr = New clsExprExpression
            objExpr.ExpressionID = !ColExpId
            objExpr.ConstructExpression
            
            'JPD 20031211 Fault 7679
            'objExpr.ValidateExpression True 'MH20010508
            objExpr.ValidateExpression True, True 'MH20010508

            Select Case objExpr.ReturnType
            Case giEXPRVALUE_DATE, giEXPRVALUE_BYREF_DATE
              mintType(intIndex) = sqlDate
            Case giEXPRVALUE_NUMERIC, giEXPRVALUE_BYREF_NUMERIC
              mintType(intIndex) = sqlNumeric
            Case giEXPRVALUE_LOGIC, giEXPRVALUE_BYREF_LOGIC
              mintType(intIndex) = sqlBoolean
            Case Else
              mintType(intIndex) = 0
            End Select

            Set objExpr = Nothing

          End If

        End Select
      
        If fOK = False Then
          Exit Sub
        End If


        .MoveNext
      Loop
    
    End If
  End With


  'If Email address is same for all records then add to select
  'If mstrEmailAddr <> vbNullString Then
  '  mstrSQLSelect = mstrSQLSelect & ", " & _
  '        mstrEmailAddr & " as 'Email Address'"
  'End If


  If mstrWhereIDs <> vbNullString Then
    mstrSQLWhere = _
      "(" & mstrWhereIDs & ")" & _
      IIf(mstrSQLWhere <> vbNullString, " OR ", vbNullString) & _
      mstrSQLWhere
  End If

  
  strPicklistFilterSelect = GetPicklistFilterSelect
  If fOK = False Then
    Exit Sub
  End If
  If strPicklistFilterSelect <> vbNullString Then
    mstrSQLWhere = mstrSQLWhere & _
      IIf(mstrSQLWhere <> vbNullString, " AND ", vbNullString) & _
      mstrSQLFrom & ".ID IN (" & strPicklistFilterSelect & ")"
  End If
    
    
  If mstrSQLWhere <> vbNullString Then
    mstrSQLWhere = " WHERE " & mstrSQLWhere
  End If
  
  
  mrsMailMergeColumns.Close
  Set mrsMailMergeColumns = Nothing


  Call SQLOrderByClause

Exit Sub

LocalErr:
  mstrStatusMessage = "Error processing calculation/column definitions."
  fOK = False

End Sub


Private Sub SQLAddColumn(sColumnList As String, lngTableID As Long, sTableName As String, sColumnName As String, strColCode As String)

  Dim objTableView As CTablePrivilege
  Dim objColumnPrivileges As CColumnPrivileges
  Dim fColumnOK As Boolean
  Dim sSource As String
  Dim fFound As Boolean
  Dim iNextIndex As Integer
  
  Dim sRealSource As String
  Dim sCaseStatement As String
  Dim sWhereColumn As String
  Dim sBaseIDColumn As String

  Dim asViews() As String
  
  On Error GoTo LocalErr
  
  
  Set objColumnPrivileges = GetColumnPrivileges(sTableName)
  fColumnOK = objColumnPrivileges.IsValid(sColumnName)
  If fColumnOK Then
    fColumnOK = objColumnPrivileges.item(sColumnName).AllowSelect
  End If
    
  Set objColumnPrivileges = Nothing
    
  If fColumnOK Then
    ' The column can be read from the base table/view, or directly from a parent table.
    ' Add the column to the column list.
      
    sRealSource = gcoTablePrivileges.item(sTableName).RealSource
      
    sColumnList = sColumnList & _
        IIf(sColumnList <> vbNullString, ", ", "") & _
        sRealSource & "." & sColumnName
        
    If strColCode <> vbNullString Then
      sColumnList = sColumnList & _
        " AS " & "'" & strColCode & "'"
    End If
    
    If sTableName <> mstrDefBaseTable Then
        
      fFound = False
      For iNextIndex = 1 To UBound(mlngTableViews, 2)
        If mlngTableViews(1, iNextIndex) = 0 And _
          mlngTableViews(2, iNextIndex) = lngTableID Then
          fFound = True
          Exit For
        End If
      Next iNextIndex
          
      ' if this column is not from the base table then it must be from a parent
      ' table, therefore include it in the join code
      If Not fFound Then
        iNextIndex = UBound(mlngTableViews, 2) + 1
        ReDim Preserve mlngTableViews(2, iNextIndex)
        mlngTableViews(1, iNextIndex) = 0
        mlngTableViews(2, iNextIndex) = lngTableID

        
        ' The table has not yet been added to the join code, and it is
        ' not the base table so add it to the array and the join code.
        mstrSQLJoin = mstrSQLJoin & _
          " LEFT OUTER JOIN " & sRealSource & _
          " ON " & mstrSQLFrom & ".ID_" & CStr(lngTableID) & _
          " = " & sRealSource & ".ID"
      End If

    End If
  
  Else
        
    ReDim asViews(0)
    For Each objTableView In gcoTablePrivileges.Collection
            
      'Loop thru all of the views for this table where the user has select access
      If (Not objTableView.IsTable) And _
        (objTableView.TableID = lngTableID) And _
        (objTableView.AllowSelect) Then
              
        sSource = objTableView.ViewName

        ' Get the column permission for the view.
        Set objColumnPrivileges = GetColumnPrivileges(sSource)

        If objColumnPrivileges.IsValid(sColumnName) Then
          If objColumnPrivileges.item(sColumnName).AllowSelect Then
            ' Add the view info to an array to be put into the column list or order code below.
            iNextIndex = UBound(asViews) + 1
            ReDim Preserve asViews(iNextIndex)
            asViews(iNextIndex) = sSource

            '=== This is the join code section ===
            ' Add the view to the Join code.
            ' Check if the view has already been added to the join code.
            fFound = False
            For iNextIndex = 1 To UBound(mlngTableViews, 2)
              If mlngTableViews(2, iNextIndex) = objTableView.ViewID Then
                fFound = True
                Exit For
              End If
            Next iNextIndex

            If Not fFound Then
              ' The view has not yet been added to the join code, so add it to the array and the join code.

              iNextIndex = UBound(mlngTableViews, 2) + 1
              ReDim Preserve mlngTableViews(2, iNextIndex)
              mlngTableViews(1, iNextIndex) = 1
              mlngTableViews(2, iNextIndex) = objTableView.ViewID

              
              'MH20000725 Fault 638
              'A problem was occuring for a self service user.
              'Base = view of child and column from view of parent included
              'caused a problem with the following join command
              'Need to check view on same table as base otherwise
              'join slightly differently
              
              'mstrSQLJoin = mstrSQLJoin & vbCrLf & _
                  " LEFT OUTER JOIN " & sSource & _
                  " ON " & mstrSQLFrom & ".ID = " & sSource & ".ID"

              'mstrWhereIDs = mstrWhereIDs & _
                IIf(mstrWhereIDs <> vbNullString, " OR ", vbNullString) & _
                mstrSQLFrom & ".ID IN (SELECT ID FROM " & sSource & ")"
              
              
              If objTableView.TableID = mlngDefBaseTableID Then
                sBaseIDColumn = mstrSQLFrom & ".ID"
              Else
                sBaseIDColumn = mstrSQLFrom & ".ID_" & CStr(objTableView.TableID)
              End If

              mstrSQLJoin = mstrSQLJoin & vbCrLf & _
                " LEFT OUTER JOIN " & sSource & _
                " ON " & sBaseIDColumn & " = " & sSource & ".ID"

              mstrWhereIDs = mstrWhereIDs & _
                IIf(mstrWhereIDs <> vbNullString, " OR ", vbNullString) & _
                sBaseIDColumn & " IN (SELECT ID FROM " & sSource & ")" & _
                " OR (ISNULL(" & sBaseIDColumn & ", 0) = 0)"

            End If
          End If
          '=== End of Join Code ===
            
            
          Set objColumnPrivileges = Nothing
        End If

      End If
    Next objTableView
    Set objTableView = Nothing

    ' The current user does have permission to 'read' the column through a/some view(s) on the
    ' table.
    If UBound(asViews) = 0 Then
      mstrStatusMessage = "You do not have permission to see the column '" & sColumnName & "'" & vbCrLf & _
                        "either directly or through any views." & vbCrLf
      fOK = False
      Exit Sub
                  
    Else
      ' Add the column to the column list.
      sCaseStatement = "CASE"
      sWhereColumn = vbNullString
      For iNextIndex = 1 To UBound(asViews)
        sCaseStatement = sCaseStatement & _
          " WHEN NOT " & asViews(iNextIndex) & "." & sColumnName & " IS NULL THEN " & asViews(iNextIndex) & "." & sColumnName & vbCrLf
      Next iNextIndex

      If Len(sCaseStatement) > 0 Then
        sCaseStatement = sCaseStatement & _
          " ELSE NULL END"

        If strColCode <> vbNullString Then
          sCaseStatement = sCaseStatement & _
            " AS " & "'" & strColCode & "'"
        End If

        sColumnList = sColumnList & _
          IIf(Len(sColumnList) > 0, ", ", "") & vbCrLf & _
          sCaseStatement

        If sWhereColumn <> vbNullString Then
          mstrSQLWhere = mstrSQLWhere & _
            IIf(Len(mstrSQLWhere) > 0, " AND ", vbNullString) & _
            "((" & sWhereColumn & "))"
        End If

      End If
    End If
  End If

Exit Sub

LocalErr:
  mstrStatusMessage = "Error building SQL Statement"
  fOK = False

End Sub


Private Sub SQLOrderByClause()

  Dim rsTemp As Recordset
  Dim strSQL As String
  Dim intCount As Integer

  On Error GoTo LocalErr
  
  'strSQL = "SELECT ASRSysTables.TableID, " & _
           "       ASRSysTables.TableName, " & _
           "       ASRSysColumns.ColumnID, " & _
           "       ASRSysColumns.ColumnName, " & _
           "       ASRSysOrderItems.Ascending " & _
           " FROM ASRSysOrderItems" & _
           " JOIN ASRSysColumns ON (ASRSysOrderItems.ColumnID = ASRSysColumns.ColumnID)" & _
           " JOIN ASRSysTables ON (ASRSysColumns.TableID = ASRSysTables.TableID)" & _
           " WHERE Type = 'O' AND OrderID = " & CStr(mlngDefOrderID) & _
           " ORDER BY Sequence"

  strSQL = "SELECT ASRSysTables.TableID, " & _
           "       ASRSysTables.TableName, " & _
           "       ASRSysColumns.ColumnID, " & _
           "       ASRSysColumns.ColumnName, " & _
           "       ASRSysMailMergeColumns.SortOrder " & _
           "FROM ASRSysMailMergeColumns " & _
           "JOIN ASRSysColumns ON (ASRSysMailMergeColumns.ColumnID = ASRSysColumns.ColumnID) " & _
           "JOIN ASRSysTables ON (ASRSysColumns.TableID = ASRSysTables.TableID) " & _
           "WHERE ASRSysMailMergeColumns.MailMergeID = " & CStr(mlngMailMergeID) & " " & _
           "  AND SortOrderSequence > 0 " & _
           "  AND (ASRSysMailMergeColumns.Type = '" & sTYPECODE_COLUMN & "' OR ASRSysMailMergeColumns.Type = '" & sTYPECODE_EXPRESSION & "') " & _
           "ORDER BY SortOrderSequence"
  
  Set rsTemp = mdatData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)

  
  With rsTemp
    Do While Not .EOF
      Call SQLAddColumn(mstrSQLOrder, !TableID, !TableName, !ColumnName, vbNullString)

      mstrSQLOrder = mstrSQLOrder & IIf(Left$(!SortOrder, 1) = "A", " ASC", " DESC")

      If fOK = False Then
        Exit Sub
      End If
      .MoveNext
    Loop
  End With

  If mstrSQLOrder <> vbNullString Then
    mstrSQLOrder = " ORDER BY " & mstrSQLOrder
  End If
  
  Set rsTemp = Nothing

Exit Sub

LocalErr:
  mstrStatusMessage = "Error building 'Order By' clause"
  fOK = False

End Sub


Private Sub SQLAddCalculation(lngExpID As Long, strColCode As String)
       
  Dim lngCalcViews() As Long
  Dim objCalcExpr As clsExprExpression
  Dim intCount As Integer
  Dim blnFound As Boolean
  Dim intNextIndex As Integer
  Dim sCalcCode As String
  Dim sSource As String
  Dim lngTestTableID As Long
  Dim objTableView As CTablePrivilege
  Dim sBaseIDColumn As String
    
  ReDim lngCalcViews(2, 0)
  Set objCalcExpr = New clsExprExpression
  fOK = objCalcExpr.Initialise(mlngDefBaseTableID, lngExpID, giEXPR_RUNTIMECALCULATION, giEXPRVALUE_UNDEFINED)
  If fOK Then
    fOK = objCalcExpr.RuntimeCalculationCode(lngCalcViews, sCalcCode, True)
    
    If fOK And gbEnableUDFFunctions Then
      fOK = objCalcExpr.UDFCalculationCode(lngCalcViews, mastrUDFsRequired(), True)
    End If
  End If
        
  If fOK = False Then
    'mstrStatusMessage = "You do not have permission to run a calculation contained in this mail merge."
    mstrStatusMessage = "You do not have permission to use the '" & Trim(objCalcExpr.Name) & "' calculation."
    Set objCalcExpr = Nothing
    Exit Sub
  End If
  Set objCalcExpr = Nothing
  
  
  mstrSQLSelect = mstrSQLSelect & _
    IIf(mstrSQLSelect <> vbNullString, ", ", vbNullString) & _
    sCalcCode & " AS '" & strColCode & "'"

  
  ' Add the required views to the JOIN code.
  For intCount = 1 To UBound(lngCalcViews, 2)
    If lngCalcViews(1, intCount) = 1 Then
      ' Check if view has already been added to the array
      blnFound = False
      For intNextIndex = 1 To UBound(mlngTableViews, 2)
        If mlngTableViews(1, intNextIndex) = 1 And _
          mlngTableViews(2, intNextIndex) = lngCalcViews(2, intCount) Then
          blnFound = True
          Exit For
        End If
      Next intNextIndex

      If Not blnFound Then
        ' View hasnt yet been added, so add it !
        intNextIndex = UBound(mlngTableViews, 2) + 1
        ReDim Preserve mlngTableViews(2, intNextIndex)
        mlngTableViews(1, intNextIndex) = 1
        mlngTableViews(2, intNextIndex) = lngCalcViews(2, intCount)
                 
        lngTestTableID = lngCalcViews(2, intCount)
        
        Set objTableView = gcoTablePrivileges.FindViewID(lngCalcViews(2, intCount))
        'sSource = gcoTablePrivileges.FindViewID(lngCalcViews(2, intCount)).RealSource
        sSource = objTableView.RealSource
        
        'TM20020904 Fault 4358 - depending on whether the table that is about to
        '                        joined is a Parent or Child denotes which ID
        '                        columns are used to establish the join.
        If datGeneral.IsAParentOf(objTableView.TableID, mlngDefBaseTableID) Then
          'Table/View is parent of Base Table.
          mstrSQLJoin = mstrSQLJoin & vbCrLf & _
            " LEFT OUTER JOIN " & sSource & _
            " ON " & mstrSQLFrom & ".ID_" & CStr(objTableView.TableID) & " = " & sSource & ".ID"
            
        ElseIf datGeneral.IsAChildOf(objTableView.TableID, mlngDefBaseTableID) Then
          'Table/View is child of Base Table.
          ' JPD20021028 Fault 4661
          'mstrSQLJoin = mstrSQLJoin & vbCrLf & _
            " LEFT OUTER JOIN " & sSource & _
            " ON " & mstrSQLFrom & ".ID = " & sSource & ".ID_" & CStr(mlngDefBaseTableID)
            
'        Else
'          mstrSQLJoin = mstrSQLJoin & vbCrLf & _
'            " LEFT OUTER JOIN " & sSource & _
'            " ON " & mstrSQLFrom & ".ID = " & sSource & ".ID"
'
        
        'JDM - 01/10/03 - Fault 7093 - Problem with calcs when going via views.
        ElseIf objTableView.IsTable = False Then
          
          If objTableView.TableID = mlngDefBaseTableID Then
            sBaseIDColumn = mstrSQLFrom & ".ID"
          Else
            sBaseIDColumn = mstrSQLFrom & ".ID_" & CStr(objTableView.TableID)
          End If
          
          mstrWhereIDs = mstrWhereIDs & _
            IIf(mstrWhereIDs <> vbNullString, " OR ", vbNullString) & _
            sBaseIDColumn & " IN (SELECT ID FROM " & sSource & ")"
          
          mstrSQLJoin = mstrSQLJoin & vbCrLf & _
            " LEFT OUTER JOIN " & sSource & _
            " ON " & sBaseIDColumn & " = " & sSource & ".ID"

        End If

       

        'mstrWhereIDs = mstrWhereIDs & _
          IIf(mstrWhereIDs <> vbNullString, " OR ", vbNullString) & _
          mstrSQLFrom & ".ID IN (SELECT ID FROM " & sSource & ")"
      End If

'TM20020320 Fault 3683
'********************************************************************************
    ElseIf lngCalcViews(1, intCount) = 0 Then
      ' Check if table has already been added to the array
      blnFound = False
      For intNextIndex = 1 To UBound(mlngTableViews, 2)
        'TM20020827 Fault 4338
        If mlngTableViews(1, intNextIndex) = 0 And _
          mlngTableViews(2, intNextIndex) = lngCalcViews(2, intCount) Then
          blnFound = True
          Exit For
        End If
      Next intNextIndex

      'TM20020827 Fault 4338
      'Don't add the table id to the array if it is the base table id.
      If Not blnFound Then
        blnFound = (lngCalcViews(2, intCount) = mlngDefBaseTableID)
      End If
      
      If Not blnFound Then
        ' Table hasnt yet been added, so add it !
        intNextIndex = UBound(mlngTableViews, 2) + 1
        ReDim Preserve mlngTableViews(2, intNextIndex)
        mlngTableViews(1, intNextIndex) = 0
        mlngTableViews(2, intNextIndex) = lngCalcViews(2, intCount)
           
        lngTestTableID = lngCalcViews(2, intCount)
        
        Set objTableView = gcoTablePrivileges.FindTableID(lngCalcViews(2, intCount))
        'sSource = gcoTablePrivileges.FindTableID(lngCalcViews(2, intCount)).RealSource
        sSource = objTableView.RealSource
        
        'TM20020904 Fault 4358 - depending on whether the table that is about to
        '                        joined is a Parent or Child denotes which ID
        '                        columns are used to establish the join.
        If datGeneral.IsAParentOf(objTableView.TableID, mlngDefBaseTableID) Then
          'Table/View is parent of Base Table.
          mstrSQLJoin = mstrSQLJoin & vbCrLf & _
            " LEFT OUTER JOIN " & sSource & _
            " ON " & mstrSQLFrom & ".ID_" & lngCalcViews(2, intCount) & " = " & sSource & ".ID"
            
        ElseIf datGeneral.IsAChildOf(objTableView.TableID, mlngDefBaseTableID) Then
          'Table/View is child of Base Table.
          ' JPD20021028 Fault 4661
          'mstrSQLJoin = mstrSQLJoin & vbCrLf & _
            " LEFT OUTER JOIN " & sSource & _
            " ON " & mstrSQLFrom & ".ID = " & sSource & ".ID_" & CStr(mlngDefBaseTableID)
            
        End If

      End If
'********************************************************************************

    End If
  Next
        
End Sub


Private Function ValidEmailAddress(strEmailAddress) As Boolean
  'Must only contain one @ sign and must be something in front of @ and after @
  
  Dim varTemp As Variant
  
  ValidEmailAddress = False
  
  varTemp = Split(Trim(strEmailAddress), "@")
  
  If UBound(varTemp) = 1 Then
    If Left(strEmailAddress, "1") <> "@" And Right(strEmailAddress, "1") <> "@" Then
      'Check for full stop after @ sign
      ValidEmailAddress = (InStr(varTemp(1), ".") > 0)
    End If
  End If

End Function



Private Sub AddEventLogFail(lngRecordID As Long, strErrDesc As String)

End Sub


Private Function GetRecordDesc(lngRecordID As Long)

  ' Return TRUE if the user has been granted the given permission.
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter

  On Error GoTo LocalErr
  
  If mlngRecordDescExprID < 1 Then
    GetRecordDesc = "Record Description Undefined"
    Exit Function
  End If
  
  
  ' Check if the user can create New instances of the given category.
  Set cmADO = New ADODB.Command
  With cmADO
    .CommandText = "dbo.sp_ASRExpr_" & mlngRecordDescExprID
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon

    Set pmADO = .CreateParameter("Result", adVarChar, adParamOutput, VARCHAR_MAX_Size)
    .Parameters.Append pmADO

    Set pmADO = .CreateParameter("RecordID", adInteger, adParamInput)
    .Parameters.Append pmADO
    pmADO.value = lngRecordID

    cmADO.Execute

    GetRecordDesc = .Parameters(0).value
  End With
  Set cmADO = Nothing

  
  If Trim$(GetRecordDesc) = vbNullString Then
    GetRecordDesc = "Record Description Undefined"
  End If

Exit Function

LocalErr:
  mstrStatusMessage = "Error reading record description" & vbCr & _
                      "(ID = " & CStr(lngRecordID) & ", Record Description = " & CStr(mlngRecordDescExprID)
  fOK = False

End Function


Private Function FormatData(varInput As Variant, intColIndex As Integer) As String
  
  Dim vData As Variant
  
'  'Change logic columns to 'Yes' or 'No'
'  If Not IsNull(varInput) Then
'    If mintType(intColIndex) = sqlBoolean Then
'      If CBool(varInput) Then
'        varInput = "Y"
'      Else
'        varInput = "N"
'      End If
'    End If
'  End If
  
  'MH20010807 Check for Nulls...
  vData = IIf(IsNull(varInput), vbNullString, varInput)
  
  'MH20021018 Remove invalid characters...
  vData = Replace(vData, vbCr, " ")
  vData = Replace(vData, vbLf, "")
  vData = Replace(vData, vbTab, " ")
  vData = Replace(vData, Chr(34), "'")
  
  
  Select Case mintType(intColIndex)
  Case SQLDataType.sqlNumeric
    If mintDecimals(intColIndex) <> 0 Then
      vData = Format(vData, "0." & String(mintDecimals(intColIndex), "0"))
    Else
      If mlngSize(intColIndex) > 0 Then
        If vData = "0" Then
          vData = Format(vData, "0")
        Else
          vData = Format(vData, "#")
        End If
      End If
    End If
    
  Case SQLDataType.sqlDate
    vData = Format(vData, DateFormat)
      
  Case SQLDataType.sqlBoolean
    vData = IIf(CBool(vData), "Y", "N")

  Case Else
    vData = Trim(vData)

  End Select


  'Check if has decimal places
  If mlngSize(intColIndex) > 0 Then   'Size restriction
    If mintDecimals(intColIndex) > 0 Then
      If InStr(vData, ".") > mlngSize(intColIndex) Then
        vData = Left(vData, mlngSize(intColIndex)) & _
                     Mid(vData, InStr(vData, "."))
      End If
    
    Else
      If Len(vData) > mlngSize(intColIndex) Then
        vData = Left(vData, mlngSize(intColIndex))
      End If
  
    End If
  End If

  ' Do we have the digit seperator
  If amColumnInformation(intColIndex, 1) = True Then
    If mintDecimals(intColIndex) <> 0 Then
      vData = Format(vData, "#,0." & String(mintDecimals(intColIndex), "0"))
    Else
      vData = Format(vData, "#,0")
    End If
  End If

  FormatData = vData

End Function

' Creates a labels template as a Word Document
Private Sub WrdCreateLabelTemplate(blnSuppressPrompt As Boolean)

  On Error GoTo CannotCreateTemplate

  Dim objRange As Object
  Dim objWordRange As Word.Range
  Dim objWordSelection As Word.Selection
 
  Dim intFields As Integer
  Dim intNumberOfLabels As Integer
  Dim intColumns As Integer
  Dim intRows As Integer
  Dim objCustomLabel As Word.CustomLabel
  Dim intStepHorizontal As Integer
  Dim intStepVertical As Integer
  Dim iTempRow As Integer
  Dim iTempColumn As Integer
  Dim bOnlyOnePerPage As Boolean

  ' Create a custom label based on the defined label type
  Set objCustomLabel = wrdApp.MailingLabel.CustomLabels.Add("hrprolabel")
  objCustomLabel.PageSize = IIf(mintLabelPageOrientation, wdCustomLabelA4LS, wdCustomLabelA4)
  objCustomLabel.HorizontalPitch = Maximum(msngLabelHorizontalPitch, msngLabelWidth)
  objCustomLabel.VerticalPitch = Maximum(msngLabelVerticalPitch, msngLabelHeight)
  objCustomLabel.Height = msngLabelHeight
  objCustomLabel.Width = msngLabelWidth
  objCustomLabel.TopMargin = msngLabelTopMargin
  objCustomLabel.SideMargin = msngLabelSideMargin
  objCustomLabel.NumberAcross = miLabelNumberAcross
  objCustomLabel.NumberDown = miLabelNumberDown

  ' Are we starting in a location other than lop left
  bOnlyOnePerPage = (miLabelStartColumn > 0 Or miLabelStartRow > 0)
       
  ' Create mail merge document
  mstrDefTemplateFile = GetTmpFName
  wrdApp.Documents.Add
       
  Set wrdDocTemplate = wrdApp.Documents.Add
  ' Set the page size
  'wrdDocTemplate.PageSetup.Orientation = mintLabelPageOrientation
  If mintLabelPageOrientation Then
    wrdDocTemplate.PageSetup.PageWidth = msngLabelPageHeight
    wrdDocTemplate.PageSetup.PageHeight = msngLabelPageWidth
  Else
    wrdDocTemplate.PageSetup.PageWidth = msngLabelPageWidth
    wrdDocTemplate.PageSetup.PageHeight = msngLabelPageHeight
  End If
  
  If objCustomLabel.Valid = False Then
    mstrStatusMessage = "You are running a label template " & _
                    "which is not supported by the installed version of Microsoft Word."
    fOK = False
    GoTo TidyUpAndExit
  Else
    
    Set wrdDocTemplate = wrdApp.MailingLabel.CreateNewDocument("hrprolabel")
    wrdDocTemplate.PageSetup.FirstPageTray = wdPrinterAutomaticSheetFeed
    wrdDocTemplate.PageSetup.OtherPagesTray = wdPrinterAutomaticSheetFeed
    
    ' Are there filler columns in our table
    intStepHorizontal = IIf(miLabelNumberAcross = wrdDocTemplate.Tables(1).Columns.Count, 1, 2)
    intStepVertical = IIf(miLabelNumberDown = wrdDocTemplate.Tables(1).Rows.Count, 1, 2)
     
    ' Record amount of rows and columns we've entered values into
    iTempColumn = 0
     
    ' Populate the template with mail merge fields
    For intColumns = 1 To wrdDocTemplate.Tables(1).Columns.Count Step intStepHorizontal
      
      iTempColumn = iTempColumn + 1
      iTempRow = 0
      
      For intRows = 1 To wrdDocTemplate.Tables(1).Rows.Count Step intStepVertical
    
        iTempRow = iTempRow + 1
    
        ' Only output one location if running singally
        If Not bOnlyOnePerPage Or (miLabelStartColumn = iTempColumn And miLabelStartRow = iTempRow) Then
    
          Set objRange = wrdApp.Selection
    
          ' Select current cell
          objRange.Tables(1).Cell(intRows, intColumns).Select
          objRange.Collapse
      
          ' Add next mail merge field (has to be at top of label, not at end...)
          If (intColumns > 1 Or intRows > 1) And Not bOnlyOnePerPage Then
            wrdApp.ActiveDocument.Fields.Add objRange.Range, wdFieldNext, "", False
          End If
      
          ' Add the fields
          For intFields = 1 To mrsMergeData.Fields.Count - 1
                                    
            ' Start column on new line?
            If intFields > 1 Then
              If amColumnInformation(intFields, 0) Then
                objRange.TypeParagraph
              Else
                objRange.TypeText " "
              End If
            End If
                   
            Select Case amColumnType(intFields)
              
              ' Add a header
              Case sTYPECODE_HEADING
                objRange.Font.Name = mobjFormatDetails(0).FontName
                objRange.Font.Size = mobjFormatDetails(0).FontSize
                If GetOfficeWordVersion >= 9 Then
                  objRange.Font.Color = mobjFormatDetails(0).ForeCol
                Else
                  objRange.Font.ColorIndex = GetWordColourIndex(mobjFormatDetails(0).ForeCol)
                End If
                objRange.Font.Underline = IIf(mobjFormatDetails(0).Underline = True, wdUnderlineSingle, wdUnderlineNone)
                objRange.Font.Bold = mobjFormatDetails(0).Bold
                objRange.Font.Italic = mobjFormatDetails(0).Italic
                wrdApp.ActiveDocument.Fields.Add objRange.Range, wdFieldMergeField, mrsMergeData.Fields(intFields).Name, False
                objRange.Font.Underline = wdUnderlineNone
                objRange.Font.Bold = False
         
              ' Add a separator
              Case sTYPECODE_SEPARATOR
                objRange.Font.Name = mobjFormatDetails(1).FontName
                objRange.Font.Size = mobjFormatDetails(1).FontSize
                If GetOfficeWordVersion >= 9 Then
                  objRange.Font.Color = mobjFormatDetails(1).ForeCol
                Else
                  objRange.Font.ColorIndex = GetWordColourIndex(mobjFormatDetails(1).ForeCol)
                End If
                objRange.Font.Underline = IIf(mobjFormatDetails(1).Underline = True, wdUnderlineSingle, wdUnderlineNone)
                objRange.Font.Bold = mobjFormatDetails(1).Bold
                objRange.Font.Italic = mobjFormatDetails(1).Italic
                objRange.TypeText ""
              
              ' Add the fields
              Case sTYPECODE_COLUMN
                objRange.Font.Name = mobjFormatDetails(1).FontName
                objRange.Font.Size = mobjFormatDetails(1).FontSize
                If GetOfficeWordVersion >= 9 Then
                  objRange.Font.Color = mobjFormatDetails(1).ForeCol
                Else
                  objRange.Font.ColorIndex = GetWordColourIndex(mobjFormatDetails(1).ForeCol)
                End If
                objRange.Font.Underline = IIf(mobjFormatDetails(1).Underline = True, wdUnderlineSingle, wdUnderlineNone)
                objRange.Font.Bold = mobjFormatDetails(1).Bold
                objRange.Font.Italic = mobjFormatDetails(1).Italic
                wrdApp.ActiveDocument.Fields.Add objRange.Range, wdFieldMergeField, mrsMergeData.Fields(intFields).Name, False
            
              ' Add exprsssions
              Case sTYPECODE_EXPRESSION
                objRange.Font.Name = mobjFormatDetails(1).FontName
                objRange.Font.Size = mobjFormatDetails(1).FontSize
                If GetOfficeWordVersion >= 9 Then
                  objRange.Font.Color = mobjFormatDetails(1).ForeCol
                Else
                  objRange.Font.ColorIndex = GetWordColourIndex(mobjFormatDetails(1).ForeCol)
                End If
                objRange.Font.Underline = IIf(mobjFormatDetails(1).Underline = True, wdUnderlineSingle, wdUnderlineNone)
                objRange.Font.Bold = mobjFormatDetails(1).Bold
                objRange.Font.Italic = mobjFormatDetails(1).Italic
                wrdApp.ActiveDocument.Fields.Add objRange.Range, wdFieldMergeField, mrsMergeData.Fields(intFields).Name, False
            
            End Select
                   
        
          Next intFields
        
        End If
                              
      Next intRows
    Next intColumns
    
    wrdDocTemplate.SaveAs FileName:=mstrDefTemplateFile, FileFormat:=Val(mstrDefTemplateFormat), AddToRecentFiles:=False
  End If
  
TidyUpAndExit:
  wrdDocTemplate.Saved = True
  wrdApp.Documents.Close False
  wrdApp.Visible = mblnAppVisible

  ' Delete the temporary custom label
  objCustomLabel.Delete
  Set objCustomLabel = Nothing
  Exit Sub

CannotCreateTemplate:
  Select Case Err.Number
    Case 5138
      mstrStatusMessage = "You are running a label template " & _
                      "which is not supported by the installed version of Microsoft Word."
      fOK = False
      GoTo TidyUpAndExit

    Case Else
      Resume Next
  End Select

End Sub

Private Sub RetrieveLabelDefinition()

  On Error GoTo LocalErr

  Dim rsLabelDefinition As Recordset
  Dim strSQL As String
  
  fOK = True
  
  strSQL = "SELECT ASRSysLabelTypes.*, " & _
           "CONVERT(integer,ASRSysLabelTypes.TimeStamp) AS intTimeStamp " & _
           "FROM ASRSysLabelTypes " & _
           "WHERE ASRSysLabelTypes.LabelTypeID = " & CStr(mlngLabelDefinitionID)
  Set rsLabelDefinition = mdatData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
  
  If rsLabelDefinition.BOF And rsLabelDefinition.EOF Then
    Set rsLabelDefinition = Nothing
    mstrStatusMessage = "This definition has been deleted by another user."
    fOK = False
    Exit Sub
  End If

  With rsLabelDefinition
    msngLabelTopMargin = wrdApp.CentimetersToPoints(.Fields("TopMargin").value)
    msngLabelSideMargin = wrdApp.CentimetersToPoints(.Fields("SideMargin").value)
    msngLabelHorizontalPitch = wrdApp.CentimetersToPoints(.Fields("HorizontalPitch").value)
    msngLabelVerticalPitch = wrdApp.CentimetersToPoints(.Fields("VerticalPitch").value)
    msngLabelHeight = wrdApp.CentimetersToPoints(.Fields("LabelHeight").value)
    msngLabelWidth = wrdApp.CentimetersToPoints(.Fields("LabelWidth").value)
    miLabelNumberDown = .Fields("NumberDown").value
    miLabelNumberAcross = .Fields("NumberAcross").value
    
    mintLabelPageOrientation = IIf(.Fields("PageOrientation").value, 1, 0)
    msngLabelPageHeight = wrdApp.CentimetersToPoints(.Fields("PageHeight").value)
    msngLabelPageWidth = wrdApp.CentimetersToPoints(.Fields("PageWidth").value)
    mbIsEnvelope = IIf(IsNull(.Fields("IsEnvelope").value), False, .Fields("IsEnvelope").value)
    
    msngEnvelopeFromTop = wrdApp.CentimetersToPoints(.Fields("FromTop").value)
    msngEnvelopeFromLeft = wrdApp.CentimetersToPoints(.Fields("FromLeft").value)
    mbEnvelopeFromTopAuto = .Fields("FromTopAuto").value
    mbEnvelopeFromLeftAuto = .Fields("FromLeftAuto").value
    
    ' Centimeterstopoints fires off extraenous winword.exes for some reason
    wrdApp.Quit
    Set wrdApp = Nothing
    
    ' Style settings
    mobjFormatDetails(0).FontName = .Fields("HeadingFontName")
    mobjFormatDetails(0).FontSize = .Fields("HeadingFontSize")
    mobjFormatDetails(0).ForeCol = .Fields("HeadingFontColour")
    mobjFormatDetails(0).Bold = .Fields("HeadingFontBold")
    mobjFormatDetails(0).Italic = .Fields("HeadingFontItalic")
    mobjFormatDetails(0).Underline = .Fields("HeadingFontUnderline")
          
    ' Standard Format
    mobjFormatDetails(1).FontName = .Fields("StandardFontName")
    mobjFormatDetails(1).FontSize = .Fields("StandardFontSize")
    mobjFormatDetails(1).ForeCol = .Fields("StandardFontColour")
    mobjFormatDetails(1).Bold = .Fields("StandardFontBold")
    mobjFormatDetails(1).Italic = .Fields("StandardFontItalic")
    mobjFormatDetails(1).Underline = .Fields("StandardFontUnderline")
  
  End With

TidyAndExit:
  Set rsLabelDefinition = Nothing
  Exit Sub

LocalErr:
  mstrStatusMessage = "Error reading Label definition"
  fOK = False
  Resume TidyAndExit

End Sub

' Prompt for the start row & column
Private Sub GetLabelStartPosition()

  Dim frmLabelsPrompt As frmLabelsPrompt
  
  Set frmLabelsPrompt = New frmLabelsPrompt

  ' Temporarily hide the progress bar
  gobjProgress.Visible = False

  With frmLabelsPrompt
    .MaximumRows = miLabelNumberDown
    .MaximumColumns = miLabelNumberAcross
    .Show vbModal

    If .Cancelled Then
      mblnUserCancelled = True
      fOK = False
    End If

    miLabelStartColumn = .StartColumn
    miLabelStartRow = .StartRow

  End With

  ' Tidy up
  Unload frmLabelsPrompt
  Set frmLabelsPrompt = Nothing

  ' Show the progress bar
  gobjProgress.Visible = True

End Sub

' Creates a envelope template as a Word Document
Private Sub WrdCreateEnvelopeTemplate(blnSuppressPrompt As Boolean)
  
  On Error GoTo CannotCreateTemplate
  
  Dim intFields As Integer
  Dim intColumns As Integer
  Dim intRows As Integer
  Dim objRange As Object

  ' Create mail merge document
  mstrDefTemplateFile = GetTmpFName
              
  Set wrdDocTemplate = wrdApp.Documents.Add
  wrdDocTemplate.PageSetup.PageWidth = msngLabelPageHeight            'This looks wrong - but it isn't...
  wrdDocTemplate.PageSetup.PageHeight = msngLabelPageWidth
  wrdApp.ActiveDocument.PageSetup.Orientation = IIf(mintLabelPageOrientation = 1, wdOrientLandscape, wdOrientPortrait)
    
'  wrdDocTemplate.Envelope.DefaultHeight = msngLabelPageHeight
'  wrdDocTemplate.Envelope.DefaultWidth = msngLabelPageWidth
  wrdDocTemplate.Envelope.Insert
  wrdApp.ActiveDocument.Envelope.address.Select
  
  'Clone the envelope section into first page (we have to delete second page later
  wrdApp.ActiveDocument.Sections(1).PageSetup.TopMargin = wrdApp.ActiveDocument.Sections(2).PageSetup.TopMargin
  wrdApp.ActiveDocument.Sections(1).PageSetup.BottomMargin = wrdApp.ActiveDocument.Sections(2).PageSetup.BottomMargin

  wrdApp.ActiveDocument.Envelope.AddressFromTop = IIf(mbEnvelopeFromTopAuto, 0, msngEnvelopeFromTop)
  wrdApp.ActiveDocument.Envelope.AddressFromLeft = IIf(mbEnvelopeFromLeftAuto, 0, msngEnvelopeFromLeft)

  ' Add the fields
  For intFields = 1 To mrsMergeData.Fields.Count - 1
           
    ' Set current range
    Set objRange = wrdApp.Selection
           
    ' Start column on new line?
    If intFields > 1 Then
      If amColumnInformation(intFields, 0) Then
        objRange.TypeParagraph
      Else
        objRange.TypeText " "
      End If
    End If
                 
    Select Case amColumnType(intFields)
      ' Add a header
      Case sTYPECODE_HEADING
'        If intFields > 1 Then
'          objRange.TypeParagraph
'        End If
        objRange.Font.Name = mobjFormatDetails(0).FontName
        objRange.Font.Size = mobjFormatDetails(0).FontSize
        If GetOfficeWordVersion >= 9 Then
          objRange.Font.Color = mobjFormatDetails(0).ForeCol
        Else
          objRange.Font.ColorIndex = GetWordColourIndex(mobjFormatDetails(0).ForeCol)
        End If
        objRange.Font.Underline = IIf(mobjFormatDetails(0).Underline = True, wdUnderlineSingle, wdUnderlineNone)
        objRange.Font.Bold = mobjFormatDetails(0).Bold
        objRange.Font.Italic = mobjFormatDetails(0).Italic
        wrdApp.ActiveDocument.Envelope.address.Fields.Add objRange.Range, wdFieldMergeField, mrsMergeData.Fields(intFields).Name, False

    
      ' Add a separator
      Case sTYPECODE_SEPARATOR
        objRange.Font.Name = mobjFormatDetails(1).FontName
        objRange.Font.Size = mobjFormatDetails(1).FontSize
        If GetOfficeWordVersion >= 9 Then
          objRange.Font.Color = mobjFormatDetails(1).ForeCol
        Else
          objRange.Font.ColorIndex = GetWordColourIndex(mobjFormatDetails(1).ForeCol)
        End If
        objRange.Font.Underline = IIf(mobjFormatDetails(1).Underline = True, wdUnderlineSingle, wdUnderlineNone)
        objRange.Font.Bold = mobjFormatDetails(1).Bold
        objRange.Font.Italic = mobjFormatDetails(1).Italic
        objRange.TypeText ""
              
      ' Add the fields
      Case sTYPECODE_COLUMN
        objRange.Font.Name = mobjFormatDetails(1).FontName
        objRange.Font.Size = mobjFormatDetails(1).FontSize
        If GetOfficeWordVersion >= 9 Then
          objRange.Font.Color = mobjFormatDetails(1).ForeCol
        Else
          objRange.Font.ColorIndex = GetWordColourIndex(mobjFormatDetails(1).ForeCol)
        End If
        objRange.Font.Underline = IIf(mobjFormatDetails(1).Underline = True, wdUnderlineSingle, wdUnderlineNone)
        objRange.Font.Bold = mobjFormatDetails(1).Bold
        objRange.Font.Italic = mobjFormatDetails(1).Italic
        wrdApp.ActiveDocument.Envelope.address.Fields.Add objRange.Range, wdFieldMergeField, mrsMergeData.Fields(intFields).Name, False

      ' Add an expression
      Case sTYPECODE_EXPRESSION
        objRange.Font.Name = mobjFormatDetails(1).FontName
        objRange.Font.Size = mobjFormatDetails(1).FontSize
        If GetOfficeWordVersion >= 9 Then
          objRange.Font.Color = mobjFormatDetails(1).ForeCol
        Else
          objRange.Font.ColorIndex = GetWordColourIndex(mobjFormatDetails(1).ForeCol)
        End If
        objRange.Font.Underline = IIf(mobjFormatDetails(1).Underline = True, wdUnderlineSingle, wdUnderlineNone)
        objRange.Font.Bold = mobjFormatDetails(1).Bold
        objRange.Font.Italic = mobjFormatDetails(1).Italic
        wrdApp.ActiveDocument.Envelope.address.Fields.Add objRange.Range, wdFieldMergeField, mrsMergeData.Fields(intFields).Name, False

    End Select
  Next intFields

  ' Get rid of the section break
  objRange.MoveDown wdParagraph, 1
  objRange.Range.Delete
  
  wrdDocTemplate.SaveAs FileName:=mstrDefTemplateFile, FileFormat:=Val(mstrDefTemplateFormat), AddToRecentFiles:=False
  
TidyUpAndExit:
  wrdDocTemplate.Saved = True
  wrdApp.Documents.Close False
  wrdApp.Visible = mblnAppVisible
  Exit Sub

CannotCreateTemplate:
  Select Case Err.Number
    Case 5138
      mstrStatusMessage = "You are running an envelope template " & _
                      "which is not supported by the installed version of Microsoft Word."
      fOK = False
      GoTo TidyUpAndExit

    Case Else
      Resume Next
  End Select

End Sub

'Private Sub WrdPrintDocument(ByVal psPrinterName As String)
'
'  If Len(psPrinterName) > 0 Then
'    wrdApp.ActivePrinter = psPrinterName
'    wrdDocOutput.PrintOut Range:=wdPrintAllDocument, Item:= _
'        wdPrintDocumentContent, Copies:=1, Pages:="", PageType:=wdPrintAllPages, _
'        ManualDuplexPrint:=False, Collate:=True, Background:=True, PrintToFile:= _
'        False, PrintZoomColumn:=0, PrintZoomRow:=0, PrintZoomPaperWidth:=0, _
'        PrintZoomPaperHeight:=0
'
'  End If
'
'End Sub

Private Sub WrdSetV1HeaderText()

  On Error GoTo LocalErr

  Dim rsDocumentMap As ADODB.Recordset
  Dim rsData As ADODB.Recordset
  Dim objWordTemplate As Word.Document
  Dim objWordSelection As Word.Selection
  
  Dim sTempTemplateFile As String
  Dim sHeader As String
  Dim sSQL As String
  
  Dim sTableName As String
  Dim sColumnName As String
  Dim sDocDateColumn As String
  
  Dim sCategory As String
  Dim sType As String
  Dim sColumn As String
  Dim sKeyField1 As String
  Dim sKeyField2 As String
  Dim sKeyField3 As String
  
  Dim iParent1TableID As String
  Dim iParent2TableID As String

  ' If Version1 and we have auto add header then insert the header text at the top of a copy of the source template
  If mlngDocumentMapID > 0 And Not mbManualHeader Then
    
    sSQL = "SELECT * FROM dbo.[ASRSysDocumentManagementTypes] WHERE [DocumentMapID] = " & mlngDocumentMapID
    Set rsDocumentMap = datGeneral.GetRecords(sSQL)
    If Not (rsDocumentMap.EOF And rsDocumentMap.BOF) Then
      
      sTableName = datGeneral.GetTableName(Val(GetModuleParameter(MODULEKEY_DOCMANAGEMENT, PARAMETERKEY_DOCMAN_CATEGORYTABLE)))
      sColumnName = datGeneral.GetColumnName(Val(GetModuleParameter(MODULEKEY_DOCMANAGEMENT, PARAMETERKEY_DOCMAN_CATEGORYCOLUMN)))
      Set rsData = datGeneral.GetRecords("SELECT DISTINCT [id], [" & sColumnName & "] FROM dbo.[" & sTableName & "] " & _
                    "WHERE [id] = " & rsDocumentMap!CategoryRecordID)
      If Not (rsData.EOF And rsData.BOF) Then
        sCategory = rsData.Fields(1).value
      End If
      rsData.Close

      sTableName = datGeneral.GetTableName(Val(GetModuleParameter(MODULEKEY_DOCMANAGEMENT, PARAMETERKEY_DOCMAN_TYPETABLE)))
      sColumnName = datGeneral.GetColumnName(Val(GetModuleParameter(MODULEKEY_DOCMANAGEMENT, PARAMETERKEY_DOCMAN_TYPECOLUMN)))
      Set rsData = datGeneral.GetRecords("SELECT [id], [" & sColumnName & "] FROM dbo.[" & sTableName & "] " & _
                    "WHERE [id] = " & rsDocumentMap!TypeRecordID)
      If Not (rsData.EOF And rsData.BOF) Then
        sType = rsData.Fields(1).value
      End If
      rsData.Close
      
      sKeyField1 = datGeneral.GetColumnName(rsDocumentMap!Parent1KeyfieldColumnID, True)
      sKeyField3 = 0
      
      wrdDocTemplate.ActiveWindow.Selection.Font.Size = 10
      wrdDocTemplate.ActiveWindow.Selection.Font.Color = vbRed
    
      Set objWordSelection = wrdApp.Selection

      With wrdDocTemplate.Fields

        objWordSelection.TypeText "~~!:~~@[$TABLE:COMPLETE_FILES]"
        objWordSelection.TypeText "~~@[DOC_SECTION:" & sCategory & "]"
        objWordSelection.TypeText "~~@[DOC_TYPE:" & sType & "]"

        ' Keyfields based on the parent table, loaded from module setup in the system manager
        Set rsData = datGeneral.GetRecords("SELECT [heading], [columnid], [value], [type] FROM dbo.[ASRSysDocumentManagementHeaderInfo]" & _
                    "WHERE NOT [type] = 0 AND [tableid] = " & rsDocumentMap!Parent1TableID)
        Do While Not (rsData.EOF Or rsData.BOF)
          objWordSelection.TypeText "~~@[" & rsData.Fields("Heading").value & ":"
          
          Select Case rsData.Fields("type").value
            Case 1
              sColumn = datGeneral.GetColumnName(rsData!ColumnID, True)
              .Add objWordSelection.Range, wdFieldMergeField, Replace(sColumn, ".", "_", 1), False
          
            Case 2
              objWordSelection.TypeText rsData.Fields("Value").value
              
            
          End Select
          
          objWordSelection.TypeText "]"
          rsData.MoveNext
        Loop
        rsData.Close
        
        ' The [DOC_DATE] comes from the keyfield in the document management type
        sDocDateColumn = datGeneral.GetColumnName(rsDocumentMap!TargetKeyfieldColumnID, True)
        objWordSelection.TypeText "~~@[DOC_DATE:"
        .Add objWordSelection.Range, wdFieldMergeField, Replace(sDocDateColumn, ".", "_", 1), False
        objWordSelection.TypeText "]~~"
        
      End With
    End If
  
    ' Tidy up
    rsDocumentMap.Close
  
  End If
  
  Set rsData = Nothing
  Set rsDocumentMap = Nothing

Exit Sub

LocalErr:
  mstrStatusMessage = "Error adding header to the Word template file <" & mstrDefTemplateFile & ">"
  fOK = False

End Sub



