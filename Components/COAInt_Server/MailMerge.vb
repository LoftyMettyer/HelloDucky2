Option Strict Off
Option Explicit On

Public Class MailMerge

  Private mrsMailMergeColumns As ADODB.Recordset
  Private mrsMergeData As ADODB.Recordset
  Private mlngMailMergeID As Integer
  Private mblnBatchMode As Boolean
  Private mblnNoRecords As Boolean

  Private mlngSuccessCount As Integer
  Private mlngFailCount As Integer

  Private fOK As Boolean
  Private mstrStatusMessage As String
  Private mblnUserCancelled As Boolean

  'Merge Definition Variables
  Private mstrDefName As String
  Private mstrDefBaseTable As String
  Private mlngRecordDescExprID As Integer
  Private mlngDefBaseTableID As Integer
  Private mlngDefOrderID As Integer
  Private mintDefSelection As Short
  Private mlngDefPickListID As Integer
  Private mlngDefFilterID As Integer

  'Private mintDefOutput As Integer
  'Private mblnDefDocSave As Boolean
  'Private mstrDefDocFile As String
  'Private mstrDefEmailCol As String
  'Private mblnDefCloseDoc As Boolean
  'Private mstrEmailAddr As String

  Private mstrDefTemplateFile As String
  Private mblnDefPauseBeforeMerge As Boolean
  Private mblnDefSuppressBlankLines As Boolean

  Private mstrDefEMailSubject As String
  Private mlngDefEmailAddrCalc As Integer
  Private mblnDefEMailAttachment As Boolean
  Private mstrDefAttachmentName As String

  Private mintDefOutputFormat As Short
  Private mblnDefOutputScreen As Boolean
  Private mblnDefOutputPrinter As Boolean
  Private mstrDefOutputPrinterName As String
  Private mblnDefOutputSave As Boolean
  Private mstrDefOutputFileName As String

  Private mlngDocManMapID As Integer
  Private mblnDocManManualHeader As Boolean


  Private mbDefinitionOwner As Boolean

  Private mintType() As Short
  Private mlngSize() As Integer
  Private mintDecimals() As Short


  'SQL Variables
  Private mstrSQLSelect As String
  Private mstrSQLFrom As String
  Private mstrSQLJoin As String
  Private mstrSQLWhere As String
  Private mstrSQLOrder As String

  Private mlngTableViews(,) As Integer
  Private mstrWhereIDs As String

  'Word Variables
  'Private wrdApp As New Word.Application
  'Private wrdDocTemplate As New Word.Document
  'Private wrdDocDataSource As New Word.Document
  'Private wrdDocOutput As New Word.Document
  'Private mstrDataSourceName As String

  ' Classes
  Private mclsData As clsDataAccess
  Private mclsGeneral As clsGeneral
  Private mobjEventLog As clsEventLog

  Private mstrOutputArray_Data() As Object
  Private mvarPrompts(,) As Object
  Private mstrClientDateFormat As String
  Private mstrLocalDecimalSeparator As String

  ' Array holding the User Defined functions that are needed for this report
  Private mastrUDFsRequired() As String

  Private mlngSingleRecordID As Integer

  Public ReadOnly Property DefOutput() As Short
    Get
      'DefOutput = mintDefOutput
    End Get
  End Property

  Public ReadOnly Property DefDocSave() As Boolean
    Get
      'DefDocSave = mblnDefDocSave
    End Get
  End Property

  Public ReadOnly Property DefDocFile() As String
    Get
      'DefDocFile = mstrDefDocFile
    End Get
  End Property

  Public ReadOnly Property DefCloseDoc() As Boolean
    Get
      'DefCloseDoc = mblnDefCloseDoc
    End Get
  End Property

  Public ReadOnly Property DefName() As String
    Get
      DefName = mstrDefName
    End Get
  End Property

  Public ReadOnly Property DefTemplateFile() As String
    Get
      DefTemplateFile = mstrDefTemplateFile
    End Get
  End Property

  Public ReadOnly Property DefPauseBeforeMerge() As Boolean
    Get
      DefPauseBeforeMerge = mblnDefPauseBeforeMerge
    End Get
  End Property

  Public ReadOnly Property DefSuppressBlankLines() As String
    Get
      DefSuppressBlankLines = CStr(mblnDefSuppressBlankLines)
    End Get
  End Property


  Public ReadOnly Property DefEMailSubject() As String
    Get
      DefEMailSubject = mstrDefEMailSubject
    End Get
  End Property

  Public ReadOnly Property DefEmailAddrCalc() As Integer
    Get
      DefEmailAddrCalc = mlngDefEmailAddrCalc
    End Get
  End Property

  Public ReadOnly Property DefEMailAttachment() As Boolean
    Get
      DefEMailAttachment = mblnDefEMailAttachment
    End Get
  End Property

  Public ReadOnly Property DefAttachmentName() As String
    Get
      DefAttachmentName = mstrDefAttachmentName
    End Get
  End Property


  Public ReadOnly Property DefOutputFormat() As Short
    Get
      DefOutputFormat = mintDefOutputFormat
    End Get
  End Property

  Public ReadOnly Property DefOutputScreen() As Boolean
    Get
      DefOutputScreen = mblnDefOutputScreen
    End Get
  End Property

  Public ReadOnly Property DefOutputPrinter() As Boolean
    Get
      DefOutputPrinter = mblnDefOutputPrinter
    End Get
  End Property

  Public ReadOnly Property DefOutputPrinterName() As String
    Get
      DefOutputPrinterName = mstrDefOutputPrinterName
    End Get
  End Property

  Public ReadOnly Property DefOutputSave() As Boolean
    Get
      DefOutputSave = mblnDefOutputSave
    End Get
  End Property

  Public ReadOnly Property DefOutputFileName() As String
    Get
      DefOutputFileName = mstrDefOutputFileName
    End Get
  End Property


  Public ReadOnly Property DefDocManMapID() As Integer
    Get
      DefDocManMapID = mlngDocManMapID
    End Get
  End Property

  Public ReadOnly Property DefDocManManualHeader() As Boolean
    Get
      DefDocManManualHeader = mblnDocManManualHeader
    End Get
  End Property

  Public WriteOnly Property Connection() As Object
    Set(ByVal Value As Object)

      ' JDM - Create connection object differently if we are in development mode (i.e. debug mode)
      If ASRDEVELOPMENT Then
        gADOCon = New ADODB.Connection
        'UPGRADE_WARNING: Couldn't resolve default property of object vConnection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        gADOCon.Open(Value)
      Else
        gADOCon = Value
      End If

    End Set
  End Property

  Public WriteOnly Property CustomReportID() As Integer
    Set(ByVal Value As Integer)
      mlngMailMergeID = Value
    End Set
  End Property

  'Public Property Let Failed(BValue As Boolean)
  '  If BValue = True Then
  '    mobjEventLog.ChangeHeaderStatus elsFailed
  '  End If
  'End Property

  Public WriteOnly Property FailedMessage() As String
    Set(ByVal Value As String)
      mobjEventLog.AddDetailEntry(Value)
    End Set
  End Property

  'Public Property Let Cancelled(BValue As Boolean)
  '  If BValue = True Then
  '    mobjEventLog.ChangeHeaderStatus elsCancelled
  '  Else
  '    mobjEventLog.ChangeHeaderStatus elsSuccessful
  '  End If
  'End Property

  Public WriteOnly Property ClientDateFormat() As String
    Set(ByVal Value As String)
      mstrClientDateFormat = Value
    End Set
  End Property

  Public WriteOnly Property LocalDecimalSeparator() As String
    Set(ByVal Value As String)
      mstrLocalDecimalSeparator = Value
    End Set
  End Property

  Public ReadOnly Property NoRecords() As Boolean
    Get
      NoRecords = mblnNoRecords
    End Get
  End Property

  'Public Function OutputArrayUBound() As Long
  '  OutputArrayUBound = UBound(mstrOutputArray_Data)
  'End Function

  Public WriteOnly Property MailMergeID() As Integer
    Set(ByVal Value As Integer)
      mlngMailMergeID = Value
    End Set
  End Property


  Public WriteOnly Property Username() As String
    Set(ByVal Value As String)
      gsUsername = Value
    End Set
  End Property

  Public ReadOnly Property ErrorString() As String
    Get
      ErrorString = mstrStatusMessage
    End Get
  End Property

  Public ReadOnly Property UserCancelled() As Boolean
    Get
      UserCancelled = mblnUserCancelled
    End Get
  End Property

  Public WriteOnly Property SingleRecordID() As Integer
    Set(ByVal Value As Integer)
      mlngSingleRecordID = Value
    End Set
  End Property

  Public WriteOnly Property EventLogID() As Integer
    Set(ByVal Value As Integer)
      mobjEventLog.EventLogID = Value
    End Set
  End Property


  Public Property SuccessCount() As Integer
    Get
      SuccessCount = mlngSuccessCount
    End Get
    Set(ByVal Value As Integer)
      mlngSuccessCount = Value
    End Set
  End Property


  Public Property FailCount() As Integer
    Get
      FailCount = mlngFailCount
    End Get
    Set(ByVal Value As Integer)
      mlngFailCount = Value
    End Set
  End Property



  'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
  Private Sub Class_Initialize_Renamed()

    ' Initialise the the classes/arrays to be used
    mclsData = New clsDataAccess
    mclsGeneral = New clsGeneral
    mobjEventLog = New clsEventLog
    Dim mvarSortOrder(2, 0) As Object
    Dim mvarColDetails(18, 0) As Object
    ReDim mlngTableViews(2, 0)
    Dim mstrViews(0) As Object
    Dim mlngColWidth(0) As Object
    Dim mvarOutputArray_Definition(0) As Object
    Dim mvarOutputArray_Columns(0) As Object
    ReDim mstrOutputArray_Data(0)

    fOK = True

  End Sub
  Public Sub New()
    MyBase.New()
    Class_Initialize_Renamed()
  End Sub

  'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
  Private Sub Class_Terminate_Renamed()

    ' Clear references to classes and clear collection objects
    'UPGRADE_NOTE: Object mclsData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    mclsData = Nothing
    'UPGRADE_NOTE: Object mclsGeneral may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    mclsGeneral = Nothing
    'UPGRADE_NOTE: Object mobjEventLog may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    mobjEventLog = Nothing
    ' JPD20030313 Do not drop the tables & columns collections as they can be reused.
    'Set gcoTablePrivileges = Nothing
    'Set gcolColumnPrivilegesCollection = Nothing

  End Sub
  Protected Overrides Sub Finalize()
    Class_Terminate_Renamed()
    MyBase.Finalize()
  End Sub

  'Public Function OutputArrayData(Index As Long) As Variant
  Public Function OutputArrayData() As Object

    'Open "D:\mike.txt" For Output As #99
    'Print #99, CStr(Index)
    'Close

    'Open "D:\mike.txt" For Append As #99
    'Print #99, mstrOutputArray_Data(Index)
    'Close

    'UPGRADE_WARNING: Couldn't resolve default property of object OutputArrayData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    OutputArrayData = VB6.CopyArray(mstrOutputArray_Data) '(Index)

  End Function

  Public Function MergeFieldsData() As Object

    Dim strArray() As Object
    Dim lngIndex As Integer

    ReDim strArray(mrsMergeData.Fields.Count - 2)

    'Open "d:\mike.txt" For Output As #99
    For lngIndex = 2 To mrsMergeData.Fields.Count
      'Print #99, CStr(lngIndex)
      'Print #99, mrsMergeData.Fields(lngIndex - 1).Name
      'Print #99, ""
      'UPGRADE_WARNING: Couldn't resolve default property of object strArray(lngIndex - 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      strArray(lngIndex - 2) = mrsMergeData.Fields(lngIndex - 1).Name
    Next
    'Close #99

    'UPGRADE_WARNING: Couldn't resolve default property of object MergeFieldsData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    MergeFieldsData = VB6.CopyArray(strArray)

  End Function

  Public Function MergeFieldsUBound() As Integer
    MergeFieldsUBound = mrsMergeData.Fields.Count - 1
  End Function

  Private Function Records(ByRef lngRec As Integer) As String
    Records = CStr(lngRec) & IIf(lngRec <> 1, " records", " record")
  End Function

  'Private Function Progress() As Boolean
  '
  '  'This needs to be here, otherwise the progress bar will continue to the end
  '  'rather than cancelling immediately
  '  If fOK = False Then
  '    Progress = False
  '    Exit Function
  '  End If
  '
  '  If gobjProgress.Cancelled Then
  '    mblnUserCancelled = True
  '    fOK = False
  '  End If
  '
  '  gobjProgress.UpdateProgress mblnBatchMode
  '
  '  Progress = fOK
  '
  'End Function


  'Private Function MergeFieldExists() As Boolean
  '
  '  Dim strTemplateFieldName As String
  '  Dim intCount As Integer
  '
  '  intCount = 1
  '  Do While intCount <= wrdDocTemplate.Fields.Count
  '
  '    strTemplateFieldName = Trim$(wrdDocTemplate.Fields(intCount).Code)
  '    If Left$(strTemplateFieldName, 10) = "MERGEFIELD" Then
  '      MergeFieldExists = True
  '      Exit Function
  '    End If
  '
  '    intCount = intCount + 1
  '  Loop
  '
  '  MergeFieldExists = False
  '
  'End Function

  Private Function CheckHiddenElements() As Boolean

    'Sub created as part of fix for Fault 2656.

    Dim sSQL As String
    Dim bShowMSG As Boolean
    Dim rsReport As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim sMessage As String
    Dim sText As String

    On Error GoTo ErrorTrap

    bShowMSG = False

    sSQL = "SELECT * FROM ASRSysMailMergeName WHERE MailMergeID = " & mlngMailMergeID
    rsTemp = mclsGeneral.GetRecords(sSQL)

    'Check for hidden picklists.
    If rsTemp.Fields("PickListID").Value Then
      sText = IsPicklistValid(rsTemp.Fields("PickListID"))
      If sText <> vbNullString Then
        'NO MSGBOX ON THE SERVER ! - MsgBox "You cannot run this Mail Merge definition as it contains a hidden picklist which has been deleted or made hidden by another user.\n" & _
        '"Please re-visit your definition to remove the hidden picklist.", vbExclamation, App.Title
        bShowMSG = True
        CheckHiddenElements = False
        GoTo TidyUpAndExit
      End If
    End If

    'Check if Primary filter is hidden.
    If rsTemp.Fields("FilterID").Value Then
      sText = IsFilterValid(rsTemp.Fields("FilterID"))
      If sText <> vbNullString Then
        'NO MSGBOX ON THE SERVER ! - MsgBox "You cannot run this Mail Merge definition as it contains a hidden filter which has been deleted or made hidden by another user.\n" & _
        '"Please re-visit your definition to remove the hidden filter.", vbExclamation, App.Title
        bShowMSG = True
        CheckHiddenElements = False
        GoTo TidyUpAndExit
      End If
    End If

    sSQL = "SELECT * FROM ASRSysMailMergeColumns WHERE MailMergeID = " & mlngMailMergeID
    rsReport = mclsGeneral.GetRecords(sSQL)

    'Check if any calculations are hidden.
    With rsReport
      If .RecordCount > 0 Then
        .MoveFirst()
        Do Until .EOF
          'If the the column type in the mail merge is an expression then check the expression
          'for hidden components / deleted components.
          If .Fields("Type").Value = "E" Then
            'If the expression has hidden components and is owned by another user or has been deleted then notify the user.
            sMessage = IsCalcValid(rsReport.Fields("ColumnID"))
            If sMessage <> vbNullString Then
              If Not bShowMSG Then
                'NO MSGBOX ON THE SERVER ! - MsgBox "You cannot run this Report definition as it contains one or more calculation(s) which have been deleted or made hidden by another user.\n" & _
                '"Please re-visit your definition to remove the hidden calculations.", vbExclamation, App.Title
                bShowMSG = True
              End If
            End If
          End If
          .MoveNext()
        Loop
      End If
      .Close()
    End With

    CheckHiddenElements = Not bShowMSG

TidyUpAndExit:
    'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsTemp = Nothing
    'UPGRADE_NOTE: Object rsReport may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsReport = Nothing
    Exit Function

ErrorTrap:
    CheckHiddenElements = False
    'NO MSGBOX ON THE SERVER ! - MsgBox "Error validating the Mail Merge definition.", vbOKOnly + vbExclamation, App.Title
    Resume TidyUpAndExit

  End Function

  Public Function ExecuteMailMerge(ByRef lngMailMergeID As Integer, ByRef pblnBatchMode As Boolean, Optional ByRef lngRecordID As Integer = 0) As Boolean

    Dim strProcedureName As String
    Dim plngEventLogID As Integer

    mlngMailMergeID = lngMailMergeID
    mblnBatchMode = pblnBatchMode
    mlngSingleRecordID = lngRecordID
    fOK = True

    'SQL Stuff
    If fOK Then fOK = CheckHiddenElements()
    If fOK Then Call SQLGetMergeDefinition()

    mobjEventLog.AddHeader(clsEventLog.EventLog_Type.eltMailMerge, mstrDefName)

    If fOK Then Call SQLGetMergeColumns()
    If fOK Then Call SQLCodeCreate()
    If fOK Then Call SQLGetMergeData()

    'If fOK Then Call WrdOpenApp
    'If fOK Then Call WrdCreateDataSource
    If fOK Then Call BuildOutputArray()

    'Call TidyUpAndExit(pblnBatchMode)
    'Call OutputJobStatus

    ExecuteMailMerge = fOK

  End Function

  Public Function EventLogAddHeader() As Integer
    mobjEventLog.AddHeader(clsEventLog.EventLog_Type.eltMailMerge, mstrDefName)
    EventLogAddHeader = mobjEventLog.EventLogID
  End Function


  Public Function SQLGetMergeData() As Boolean

    Dim strSQL As String

    On Error GoTo LocalErr
    strSQL = "SELECT " & mstrSQLSelect & vbNewLine & " FROM " & mstrSQLFrom & mstrSQLJoin & vbNewLine & mstrSQLWhere & vbNewLine & mstrSQLOrder & vbNewLine
    mrsMergeData = mclsData.OpenRecordset(strSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

    If mrsMergeData.EOF Then
      mstrStatusMessage = "No records meet selection criteria"
      mblnNoRecords = True
      fOK = False
    Else
      mblnNoRecords = False
    End If

    SQLGetMergeData = fOK

    Exit Function

LocalErr:
    fOK = False
    mstrStatusMessage = "Error retrieving merge data"
    SQLGetMergeData = fOK

  End Function


  Public Function SQLGetMergeDefinition() As Boolean

    On Error GoTo LocalErr

    Dim rsMailMergeDefinition As ADODB.Recordset
    Dim strSQL As String

    SetupTablesCollection()

    strSQL = "SELECT ASRSysMailMergeName.*, " & "ASRSysTables.TableName as TableName, " & "ASRSysTables.RecordDescExprID as RecordDescExprID " & "FROM ASRSysMailMergeName " & "JOIN ASRSYSTables ON (ASRSysTables.TableID = ASRSysMailMergeName.TableID) " & vbNewLine & "WHERE MailMergeID = " & mlngMailMergeID & " "
    rsMailMergeDefinition = mclsData.OpenRecordset(strSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

    If rsMailMergeDefinition.BOF And rsMailMergeDefinition.EOF Then
      'UPGRADE_NOTE: Object rsMailMergeDefinition may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
      rsMailMergeDefinition = Nothing
      mstrStatusMessage = "This definition has been deleted by another user."
      fOK = False
      GoTo TidyAndExit
    End If

    With rsMailMergeDefinition

      If LCase(.Fields("Username").Value) <> LCase(gsUsername) And CurrentUserAccess(modUtilAccessLog.UtilityType.utlMailMerge, mlngMailMergeID) = ACCESS_HIDDEN Then
        'UPGRADE_NOTE: Object rsMailMergeDefinition may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rsMailMergeDefinition = Nothing
        mstrStatusMessage = "This definition has been made hidden by another user."
        fOK = False
        GoTo TidyAndExit
      End If

      mstrDefName = .Fields("Name").Value
      mlngDefBaseTableID = .Fields("TableID").Value
      mstrDefBaseTable = .Fields("TableName").Value
      mlngRecordDescExprID = .Fields("RecordDescExprID").Value
      mintDefSelection = .Fields("Selection").Value
      mlngDefPickListID = .Fields("PickListID").Value
      mlngDefFilterID = .Fields("FilterID").Value
      'mstrDefBaseTable = mclsGeneral.GetTableName(mlngDefBaseTableID)
      'mlngDefOrderID = !OrderID
      'mintDefOutput = !Output
      'mblnDefCloseDoc = !CloseDoc

      mstrDefTemplateFile = .Fields("TemplateFileName").Value
      mblnDefSuppressBlankLines = .Fields("SuppressBlanks").Value
      mblnDefPauseBeforeMerge = .Fields("PauseBeforeMerge").Value

      mlngDefEmailAddrCalc = 0
      'mstrEmailAddr = vbNullString

      mintDefOutputFormat = .Fields("OutputFormat").Value
      Select Case mintDefOutputFormat
        Case 0 'Word Document
          mblnDefOutputScreen = .Fields("OutputScreen").Value
          mblnDefOutputPrinter = .Fields("OutputPrinter").Value
          mstrDefOutputPrinterName = .Fields("OutputPrinterName").Value
          mblnDefOutputSave = .Fields("OutputSave").Value
          mstrDefOutputFileName = .Fields("OutputFilename").Value

        Case 1 'Individual Emails
          'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
          If IIf(IsDBNull(.Fields("EmailAddrID").Value), 0, .Fields("EmailAddrID").Value) = 0 Then
            mstrStatusMessage = "No destination email address selected"
            'UPGRADE_NOTE: Object rsMailMergeDefinition may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            rsMailMergeDefinition = Nothing
            fOK = False
            GoTo TidyAndExit
          End If

          mstrDefEMailSubject = .Fields("EmailSubject").Value
          mlngDefEmailAddrCalc = .Fields("EmailAddrID").Value
          mblnDefEMailAttachment = .Fields("EMailAsAttachment").Value
          'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
          mstrDefAttachmentName = IIf(IsDBNull(.Fields("EmailAttachmentName").Value), "", .Fields("EmailAttachmentName").Value)

        Case 2 'Document Management
          mblnDefOutputPrinter = True
          mblnDefOutputScreen = .Fields("OutputScreen").Value
          mstrDefOutputPrinterName = .Fields("OutputPrinterName").Value
          'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
          mlngDocManMapID = IIf(IsDBNull(.Fields("DocumentMapID").Value), 0, .Fields("DocumentMapID").Value)
          'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
          mblnDocManManualHeader = IIf(IsDBNull(.Fields("ManualDocManHeader").Value), 0, .Fields("ManualDocManHeader").Value)

      End Select


      mbDefinitionOwner = (LCase(Trim(gsUsername)) = LCase(Trim(.Fields("Username").Value)))

    End With

    fOK = IsRecordSelectionValid()


    '  If fOK Then
    '    If Not ValidPrinter(mstrDefOutputPrinterName) Then
    '      mstrStatusMessage = _
    ''          "This definition is set to output to printer " & mstrDefOutputPrinterName & _
    ''          " which is not set up on your PC."
    '      fOK = False
    '    End If
    '  End If
    '

    If fOK Then
      Call UtilUpdateLastRun(modUtilAccessLog.UtilityType.utlMailMerge, mlngMailMergeID)
    End If

TidyAndExit:
    'UPGRADE_NOTE: Object rsMailMergeDefinition may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsMailMergeDefinition = Nothing
    SQLGetMergeDefinition = fOK

    Exit Function

LocalErr:
    mstrStatusMessage = "Error reading Mail Merge definition"
    'mstrStatusMessage = mstrStatusMessage & "  (" & Err.Description & ")"
    fOK = False
    Resume TidyAndExit

  End Function

  '
  'Private Function ValidPrinter(strName As String) As Boolean
  '
  '  Dim objPrinter As Printer
  '  Dim blnFound As Boolean
  '
  '  If strName <> vbNullString And strName <> "<Default Printer>" Then
  '    blnFound = False
  '    For Each objPrinter In Printers
  '      If objPrinter.DeviceName = strName Then
  '        blnFound = True
  '        Exit For
  '      End If
  '    Next
  '  Else
  '    blnFound = True
  '  End If
  '
  '  ValidPrinter = blnFound
  '
  'End Function


  Public Function BuildOutputArray() As Boolean

    'This sub converts the data to a table as it was causing
    'a problem when opening a data source with a single field

    'Const strBookMark As String = "TableStart"
    Dim strOutput As String
    Dim intCount As Short
    Dim strSQL As String
    Dim strEmailAddr As String
    Dim blnRecordOkay As Boolean

    Dim lngIndex As Integer


    On Error GoTo LocalErr

    'With wrdDocDataSource.ActiveWindow.Selection

    '.Bookmarks.Add Range:=.Range, Name:=strBookMark

    'Add Column headers
    strOutput = vbNullString
    For intCount = 2 To mrsMergeData.Fields.Count
      strOutput = strOutput & IIf(intCount > 2, vbTab, vbNullString) & Left(Trim(Replace(mrsMergeData.Fields(intCount - 1).Name, vbTab, " ")), 40)
    Next

    If mlngDefEmailAddrCalc > 0 Then
      strOutput = strOutput & vbTab & "Email_Address"
    End If

    mrsMergeData.MoveFirst()

    lngIndex = 0
    ReDim mstrOutputArray_Data(0)
    'UPGRADE_WARNING: Couldn't resolve default property of object mstrOutputArray_Data(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    mstrOutputArray_Data(0) = strOutput


    'Open "d:\mike.txt" For Output As #99
    'Print #99, strOutput

    While Not mrsMergeData.EOF And fOK

      'And record's fields
      strOutput = vbNullString
      For intCount = 2 To mrsMergeData.Fields.Count
        'strOutput = strOutput & _
        'IIf(intCount > 2, vbTab, vbNullString) & _
        'Trim(mrsMergeData.Fields(intCount - 1).Value)
        strOutput = strOutput & IIf(intCount > 2, vbTab, vbNullString) & FormatData((mrsMergeData.Fields(intCount - 1).Value), intCount - 1)
      Next

      blnRecordOkay = True
      If mlngDefEmailAddrCalc > 0 Then
        strEmailAddr = GetEmailAddress((mrsMergeData.Fields(0).Value))
        If Trim(strEmailAddr) = vbNullString Then
          'email is blank
          'UPGRADE_WARNING: Couldn't resolve default property of object GetRecordDesc(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          mobjEventLog.AddDetailEntry(GetRecordDesc((mrsMergeData.Fields(0).Value)) & vbNewLine & vbNewLine & "Email Address is blank")
          blnRecordOkay = False
        Else
          strOutput = strOutput & vbTab & strEmailAddr
        End If
      End If

      mrsMergeData.MoveNext()

      'MH20000620
      'Need to check for null output in case you are merging a single
      'record, single column and the data in that column is blank!
      '(This used to result in the data source table containing a
      'single cell which contains the column header, which in turn,
      'will result in an error opening the data source document).

      If blnRecordOkay Then
        mlngSuccessCount = mlngSuccessCount + 1
        '.TypeParagraph
        '.TypeText IIf(strOutput <> vbNullString, strOutput, " ")

        strOutput = Replace(strOutput, "\", "\\")
        strOutput = Replace(strOutput, Chr(10), "")
        strOutput = Replace(strOutput, Chr(13), " ")

        lngIndex = lngIndex + 1
        ReDim Preserve mstrOutputArray_Data(lngIndex)
        ' JPD 14/02/02 Fault 3490
        'mstrOutputArray_Data(lngIndex) = strOutput
        'UPGRADE_WARNING: Couldn't resolve default property of object mstrOutputArray_Data(lngIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        mstrOutputArray_Data(lngIndex) = IIf(strOutput = vbNullString, " ", strOutput)

      Else
        mlngFailCount = mlngFailCount + 1
      End If
      'Print #99, strOutput

    End While

    'Close #99

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

    '    If fOK Then
    '      If mrsMergeData.Fields.Count < 3 Then
    '
    '        'go to start of table highlight to end of document
    '        .GoTo What:=wdGoToBookmark, Name:=strBookMark
    '        .EndKey Unit:=wdStory, Extend:=wdExtend
    '
    '        'convert selected text into a table
    '        .ConvertToTable _
    ''            Separator:=wdSeparateByTabs, _
    ''            Format:=wdTableFormatNone, _
    ''            ApplyFont:=False, _
    ''            ApplyColor:=False, _
    ''            AutoFit:=False
    '
    '      End If
    '    End If

    'End With

    'wrdDocDataSource.Save
    'wrdDocDataSource.Close SaveChanges:=False

    BuildOutputArray = fOK

    Exit Function

LocalErr:
    mstrStatusMessage = "Error populating data source document"
    fOK = False
    BuildOutputArray = fOK

  End Function


  Public Function TidyUpAndExit() As Boolean

    '  Dim wrdNewTable As Word.Table
    '  Dim strSQL As String
    '  Dim blnCloseWord As Boolean

    On Error Resume Next

    '  Err.Clear
    '
    '  If mblnBatchMode = False Then
    '    gobjProgress.CloseProgress
    '    DoEvents
    '  Else
    '    gobjProgress.ResetBar2
    '  End If
    '
    '  Err.Clear
    '  wrdApp.DisplayAlerts = wdAlertsNone
    '
    '  'MH20010709
    '  If Err.Number <> 0 Then
    '    'Problem with wrdapp object, could be closed so recreate it.
    '    Set wrdApp = CreateObject("Word.Application")
    '    wrdApp.DisplayAlerts = wdAlertsNone
    '  End If
    '
    '  'Make sure that the template does not still reference the data source
    '  '(close it and open it incase there have been changes to it or that
    '  ' the template is now the temporary email attachment name thingy !)
    '  wrdDocTemplate.Close SaveChanges:=False
    '  Set wrdDocTemplate = wrdApp.Documents.Open( _
    ''      FileName:=mstrDefTemplateFile, _
    ''      ConfirmConversions:=False, _
    ''      ReadOnly:=False, _
    ''      AddToRecentFiles:=False, _
    ''      Revert:=False)
    '
    '  wrdDocTemplate.MailMerge.MainDocumentType = wdNotAMergeDocument
    ''  If wrdDocTemplate.Saved = False Then
    '    wrdDocTemplate.Save
    ''  End If
    '
    '  wrdDocTemplate.Close SaveChanges:=False
    '  Set wrdDocTemplate = Nothing
    '
    '
    '  'If the temporary email attachment name exists then kill it
    '  If Dir(mstrDefAttachmentName) <> vbNullString Then
    '    Kill mstrDefAttachmentName
    '  End If
    '
    '
    '  'Kill the data source
    '  wrdDocDataSource.Close SaveChanges:=False
    '  Kill mstrDataSourceName
    '  Set wrdDocDataSource = Nothing
    '
    '
    '  blnCloseWord = (mblnDefCloseDoc Or _
    ''                 (mintDefOutput <> wdSendToNewDocument) Or _
    ''                 (fOK = False))
    '
    '  If blnCloseWord Then
    '    wrdDocOutput.Close SaveChanges:=False
    '    wrdApp.Quit
    '  Else
    '    wrdApp.Visible = True
    '    If Not blnBatchMode Then
    '      DoEvents
    '      wrdApp.WindowState = wdWindowStateNormal
    '      wrdApp.Activate
    '
    '      'MH20010307 Fault 1681
    '      'If word 2000 then also need to activate the document !
    '      wrdDocOutput.Activate
    '    End If
    '  End If

    mrsMailMergeColumns.Close()
    'UPGRADE_NOTE: Object mrsMailMergeColumns may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    mrsMailMergeColumns = Nothing

    mrsMergeData.Close()
    'UPGRADE_NOTE: Object mrsMergeData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    mrsMergeData = Nothing

    '  Set wrdDocOutput = Nothing
    '  Set wrdApp = Nothing

    On Error GoTo 0
    TidyUpAndExit = True

  End Function


  Private Function GetTableAndColumnName(ByRef lngColumnID As Integer) As String

    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset

    strSQL = "SELECT ColumnName, TableName " & "FROM ASRSYSColumns " & "JOIN ASRSYSTables ON (ASRSYSColumns.TableID = ASRSYSTables.TableID)" & "WHERE ColumnID = " & CStr(lngColumnID) & " "
    rsTemp = mclsData.OpenRecordset(strSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

    fOK = Not (rsTemp.BOF And rsTemp.EOF)
    If fOK Then
      GetTableAndColumnName = rsTemp.Fields("TableName").Value & "_" & rsTemp.Fields("ColumnName").Value
    End If

    rsTemp.Close()
    'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsTemp = Nothing

  End Function


  Private Function GetPicklistFilterSelect() As String

    Dim rsTemp As ADODB.Recordset

    On Error GoTo LocalErr


    If mlngSingleRecordID > 0 Then
      GetPicklistFilterSelect = CStr(mlngSingleRecordID)

    ElseIf mlngDefPickListID > 0 Then

      mstrStatusMessage = IsPicklistValid(mlngDefPickListID)
      If mstrStatusMessage <> vbNullString Then
        fOK = False
        Exit Function
      End If


      'Get List of IDs from Picklist
      rsTemp = mclsData.OpenRecordset("EXEC sp_ASRGetPickListRecords " & mlngDefPickListID, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
      fOK = Not (rsTemp.BOF And rsTemp.EOF)

      If Not fOK Then
        mstrStatusMessage = "The base table picklist contains no records."
      Else
        GetPicklistFilterSelect = vbNullString
        Do While Not rsTemp.EOF
          GetPicklistFilterSelect = GetPicklistFilterSelect & IIf(Len(GetPicklistFilterSelect) > 0, ", ", "") & rsTemp.Fields(0).Value
          rsTemp.MoveNext()
        Loop
      End If

      rsTemp.Close()
      'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
      rsTemp = Nothing

    ElseIf mlngDefFilterID > 0 Then

      mstrStatusMessage = IsFilterValid(mlngDefFilterID)
      If mstrStatusMessage <> vbNullString Then
        fOK = False
        Exit Function
      End If

      'Get list of IDs from Filter
      fOK = mclsGeneral.FilteredIDs(mlngDefFilterID, GetPicklistFilterSelect, mvarPrompts)

      ' Generate any UDFs that are used in this filter
      If fOK Then
        mclsGeneral.FilterUDFs(mlngDefFilterID, mastrUDFsRequired)
      End If

      If Not fOK Then
        ' Permission denied on something in the filter.
        mstrStatusMessage = "You do not have permission to use the '" & mclsGeneral.GetFilterName(mlngDefFilterID) & "' filter."
      End If

    End If

    Exit Function

LocalErr:
    mstrStatusMessage = "Error processing picklist"
    fOK = False

  End Function


  Public Function SQLGetMergeColumns() As Boolean

    Dim strSQL As String
    Dim strSQL1 As String
    Dim strSQL2 As String
    Dim strSQL3 As String

    On Error GoTo LocalErr

    'Merge Column Types
    '
    ' "C" is a column which has been selected by the user
    ' "E" is an express which has been selected by the user
    ' "X" is a system column required by the merge
    '     (currently only used for the email field where required)

    strSQL1 = "SELECT 'ColExp'   = 'Col',                             " & vbNewLine & "       'ColExpId' = ASRSysColumns.ColumnID,            " & vbNewLine & "       'TableID'  = ASRSysTables.TableID,              " & vbNewLine & "       'Table'    = ASRSysTables.Tablename,            " & vbNewLine & "       'Name'     = ASRSysColumns.ColumnName,          " & vbNewLine & "       'Type'     = ASRSysColumns.DataType,            " & vbNewLine & "       'Size'     = ASRSysMailMergeColumns.Size,       " & vbNewLine & "       'Decimals' = ASRSysMailMergeColumns.Decimals    " & vbNewLine & "FROM ASRSysMailMergeColumns " & vbNewLine & "JOIN ASRSysColumns ON (ASRSysColumns.ColumnID = ASRSysMailMergeColumns.ColumnID) " & vbNewLine & "JOIN ASRSysTables ON (ASRSysTables.TableID = ASRSysColumns.TableID) " & vbNewLine & "WHERE ASRSysMailMergeColumns.Type = 'C' " & vbNewLine & "  AND ASRSysMailMergeColumns.MailMergeID = " & CStr(mlngMailMergeID) & " "

    '"WHERE ASRSysMailMergeColumns.Type <> 'E' " & vbNewLine &

    'MH20000906 Added the words "LEFT OUTER" so that we can pick up invalid calcs.
    strSQL2 = "SELECT 'ColExp'   = 'Exp',                             " & vbNewLine & "       'ColExpId' = ASRSysExpressions.ExprID,          " & vbNewLine & "       'TableID'  = 0,                                 " & vbNewLine & "       'Table'    = 'Calculation_',                    " & vbNewLine & "       'Name'     = ASRSysExpressions.Name,            " & vbNewLine & "       'Type'     = ASRSysExpressions.ReturnType,      " & vbNewLine & "       'Size'     = ASRSysMailMergeColumns.Size,       " & vbNewLine & "       'Decimals' = ASRSysMailMergeColumns.Decimals    " & vbNewLine & "FROM ASRSysMailMergeColumns " & vbNewLine & "LEFT OUTER JOIN ASRSysExpressions ON (ASRSysExpressions.ExprID = ASRSysMailMergeColumns.ColumnID) " & vbNewLine & "WHERE ASRSysMailMergeColumns.Type = 'E' " & vbNewLine & "  AND ASRSysMailMergeColumns.MailMergeID = " & CStr(mlngMailMergeID) & " "

    strSQL = strSQL1 & vbNewLine & vbNewLine & "UNION" & vbNewLine & vbNewLine & strSQL2 & vbNewLine & vbNewLine & "ORDER BY 'ColExp', 'Table', 'Name'"

    mrsMailMergeColumns = mclsData.OpenRecordset(strSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

    SQLGetMergeColumns = fOK

    Exit Function

LocalErr:
    mstrStatusMessage = "Error reading calculation/column definitions."
    fOK = False
    SQLGetMergeColumns = fOK

  End Function


  Public Function SQLCodeCreate() As Boolean

    Dim strPicklistFilterSelect As String
    Dim objExpr As clsExprExpression
    Dim intIndex As Short
    Dim objTableView As CTablePrivilege

    fOK = True

    On Error GoTo LocalErr

    mstrSQLSelect = vbNullString
    mstrSQLFrom = vbNullString
    mstrSQLJoin = vbNullString
    mstrSQLWhere = vbNullString
    mstrSQLOrder = vbNullString

    ReDim mastrUDFsRequired(0)

    ReDim mlngTableViews(2, 0)
    Dim asViews(0) As Object

    intIndex = 0
    ReDim mintType(intIndex)
    ReDim mlngSize(intIndex)

    ' JPD20030219 Fault 5070
    ' Check the user has permission to read the base table.
    fOK = False
    For Each objTableView In gcoTablePrivileges.Collection
      If (objTableView.TableID = mlngDefBaseTableID) And (objTableView.AllowSelect) Then
        fOK = True
        Exit For
      End If
    Next objTableView
    'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    objTableView = Nothing

    If Not fOK Then
      mstrStatusMessage = "You do not have permission to read the base table either directly or through any views."
      Exit Function
    End If

    mstrSQLFrom = gcoTablePrivileges.Item(mstrDefBaseTable).RealSource
    mstrSQLSelect = mstrSQLFrom & ".ID"

    With mrsMailMergeColumns
      fOK = Not (.BOF And .EOF)
      If fOK Then

        Do While Not .EOF


          '01/08/2001 MH Fault 2125
          intIndex = intIndex + 1
          ReDim Preserve mintType(intIndex)
          ReDim Preserve mlngSize(intIndex)
          ReDim Preserve mintDecimals(intIndex)

          'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
          mlngSize(intIndex) = IIf(IsDBNull(.Fields("Size").Value), 0, .Fields("Size").Value)
          'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
          mintDecimals(intIndex) = IIf(IsDBNull(.Fields("Decimals").Value), 0, .Fields("Decimals").Value)

          Select Case .Fields("ColExp").Value
            Case "Col"
              Call SQLAddColumn(mstrSQLSelect, .Fields("TableID").Value, .Fields("Table").Value, .Fields("Name").Value, .Fields("Table").Value & "_" & .Fields("Name").Value)
              mintType(intIndex) = .Fields("Type").Value

            Case "Exp"
              'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
              If IsDBNull(.Fields("Name").Value) Then
                'MH20011127
                'mstrStatusMessage = _
                '"This definition contains one or more calculation(s) which" & vbNewLine & _
                '"have been deleted by another user."
                mstrStatusMessage = "This definition contains one or more calculation(s) which " & "have been deleted by another user."
                fOK = False
                Exit Function

              ElseIf IsCalcValid(.Fields("ColExpID")) <> vbNullString Then
                'MH20011127
                'mstrStatusMessage = "You cannot run this Global definition as it contains one or more calculation(s) which have been deleted or made hidden by another user." & vbNewLine & _
                '"Please re-visit your definition to remove the hidden calculations." & vbNewLine
                mstrStatusMessage = "You cannot run this Mail Merge definition as it contains one or more calculation(s) which have been deleted or made hidden by another user. " & "Please re-visit your definition to remove the hidden calculations."
                fOK = False
                Exit Function

              Else
                Call SQLAddCalculation(.Fields("ColExpID").Value, .Fields("Table").Value & .Fields("Name").Value)

                objExpr = New clsExprExpression
                objExpr.ExpressionID = .Fields("ColExpID").Value
                objExpr.ConstructExpression()
                objExpr.ValidateExpression(True) 'MH20010508

                Select Case objExpr.ReturnType
                  Case modExpression.ExpressionValueTypes.giEXPRVALUE_DATE, modExpression.ExpressionValueTypes.giEXPRVALUE_BYREF_DATE
                    mintType(intIndex) = Declarations.SQLDataType.sqlDate
                  Case modExpression.ExpressionValueTypes.giEXPRVALUE_NUMERIC, modExpression.ExpressionValueTypes.giEXPRVALUE_BYREF_NUMERIC
                    mintType(intIndex) = Declarations.SQLDataType.sqlNumeric
                  Case modExpression.ExpressionValueTypes.giEXPRVALUE_LOGIC, modExpression.ExpressionValueTypes.giEXPRVALUE_BYREF_LOGIC
                    mintType(intIndex) = Declarations.SQLDataType.sqlBoolean
                  Case Else
                    mintType(intIndex) = 0
                End Select

                'UPGRADE_NOTE: Object objExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                objExpr = Nothing

              End If

          End Select

          If fOK = False Then
            Exit Function
          End If


          .MoveNext()
        Loop

      End If
    End With


    'If Email address is same for all records then add to select
    'If mstrEmailAddr <> vbNullString Then
    '  mstrSQLSelect = mstrSQLSelect & ", " & _
    ''        mstrEmailAddr & " as 'Email Address'"
    'End If


    If mstrWhereIDs <> vbNullString Then
      mstrSQLWhere = "(" & mstrWhereIDs & ")" & IIf(mstrSQLWhere <> vbNullString, " OR ", vbNullString) & mstrSQLWhere
    End If


    strPicklistFilterSelect = GetPicklistFilterSelect()
    If fOK = False Then
      Exit Function
    End If
    If strPicklistFilterSelect <> vbNullString Then
      mstrSQLWhere = mstrSQLWhere & IIf(mstrSQLWhere <> vbNullString, " AND ", vbNullString) & mstrSQLFrom & ".ID IN (" & strPicklistFilterSelect & ")"
    End If


    If mstrSQLWhere <> vbNullString Then
      mstrSQLWhere = " WHERE " & mstrSQLWhere
    End If


    mrsMailMergeColumns.Close()
    'UPGRADE_NOTE: Object mrsMailMergeColumns may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    mrsMailMergeColumns = Nothing


    Call SQLOrderByClause()
    SQLCodeCreate = fOK

    Exit Function

LocalErr:
    mstrStatusMessage = "Error processing calculation/column definitions."
    fOK = False
    SQLCodeCreate = fOK

  End Function


  Private Sub SQLAddColumn(ByRef sColumnList As String, ByRef lngTableID As Integer, ByRef sTableName As String, ByRef sColumnName As String, ByRef strColCode As String)

    Dim objTableView As CTablePrivilege
    Dim objColumnPrivileges As CColumnPrivileges
    Dim fColumnOK As Boolean
    Dim sSource As String
    Dim fFound As Boolean
    Dim iNextIndex As Short

    Dim sRealSource As String
    Dim sCaseStatement As String
    Dim sWhereColumn As String
    Dim sBaseIDColumn As String

    Dim asViews() As String

    On Error GoTo LocalErr


    objColumnPrivileges = GetColumnPrivileges(sTableName)
    fColumnOK = objColumnPrivileges.IsValid(sColumnName)
    If fColumnOK Then
      fColumnOK = objColumnPrivileges.Item(sColumnName).AllowSelect
    End If

    'UPGRADE_NOTE: Object objColumnPrivileges may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    objColumnPrivileges = Nothing

    If fColumnOK Then
      ' The column can be read from the base table/view, or directly from a parent table.
      ' Add the column to the column list.

      sRealSource = gcoTablePrivileges.Item(sTableName).RealSource

      sColumnList = sColumnList & IIf(sColumnList <> vbNullString, ", ", "") & sRealSource & "." & sColumnName

      If strColCode <> vbNullString Then
        sColumnList = sColumnList & " AS " & "'" & strColCode & "'"
      End If

      If sTableName <> mstrDefBaseTable Then

        fFound = False
        For iNextIndex = 1 To UBound(mlngTableViews, 2)
          If mlngTableViews(1, iNextIndex) = 0 And mlngTableViews(2, iNextIndex) = lngTableID Then
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
          mstrSQLJoin = mstrSQLJoin & " LEFT OUTER JOIN " & sRealSource & " ON " & mstrSQLFrom & ".ID_" & CStr(lngTableID) & " = " & sRealSource & ".ID"
        End If

      End If

    Else

      ReDim asViews(0)
      For Each objTableView In gcoTablePrivileges.Collection

        'Loop thru all of the views for this table where the user has select access
        If (Not objTableView.IsTable) And (objTableView.TableID = lngTableID) And (objTableView.AllowSelect) Then

          sSource = objTableView.ViewName

          ' Get the column permission for the view.
          objColumnPrivileges = GetColumnPrivileges(sSource)

          If objColumnPrivileges.IsValid(sColumnName) Then
            If objColumnPrivileges.Item(sColumnName).AllowSelect Then
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

                'mstrSQLJoin = mstrSQLJoin & vbNewLine & _
                '" LEFT OUTER JOIN " & sSource & _
                '" ON " & mstrSQLFrom & ".ID = " & sSource & ".ID"

                'mstrWhereIDs = mstrWhereIDs & _
                'IIf(mstrWhereIDs <> vbNullString, " OR ", vbNullString) & _
                'mstrSQLFrom & ".ID IN (SELECT ID FROM " & sSource & ")"


                If objTableView.TableID = mlngDefBaseTableID Then
                  sBaseIDColumn = mstrSQLFrom & ".ID"
                Else
                  sBaseIDColumn = mstrSQLFrom & ".ID_" & CStr(objTableView.TableID)
                End If

                mstrSQLJoin = mstrSQLJoin & vbNewLine & " LEFT OUTER JOIN " & sSource & " ON " & sBaseIDColumn & " = " & sSource & ".ID"

                mstrWhereIDs = mstrWhereIDs & IIf(mstrWhereIDs <> vbNullString, " OR ", vbNullString) & sBaseIDColumn & " IN (SELECT ID FROM " & sSource & ")" & " OR (ISNULL(" & sBaseIDColumn & ", 0) = 0)"

              End If
            End If
            '=== End of Join Code ===


            'UPGRADE_NOTE: Object objColumnPrivileges may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            objColumnPrivileges = Nothing
          End If

        End If
      Next objTableView
      'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
      objTableView = Nothing

      ' The current user does have permission to 'read' the column through a/some view(s) on the
      ' table.
      If UBound(asViews) = 0 Then
        mstrStatusMessage = "You do not have permission to see the column '" & sColumnName & "' " & "either directly or through any views."
        fOK = False
        Exit Sub

      Else
        ' Add the column to the column list.
        sCaseStatement = "CASE"
        sWhereColumn = vbNullString
        For iNextIndex = 1 To UBound(asViews)
          sCaseStatement = sCaseStatement & " WHEN NOT " & asViews(iNextIndex) & "." & sColumnName & " IS NULL THEN " & asViews(iNextIndex) & "." & sColumnName & vbNewLine
        Next iNextIndex

        If Len(sCaseStatement) > 0 Then
          sCaseStatement = sCaseStatement & " ELSE NULL END"

          If strColCode <> vbNullString Then
            sCaseStatement = sCaseStatement & " AS " & "'" & strColCode & "'"
          End If

          sColumnList = sColumnList & IIf(Len(sColumnList) > 0, ", ", "") & vbNewLine & sCaseStatement

          If sWhereColumn <> vbNullString Then
            mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", vbNullString) & "((" & sWhereColumn & "))"
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

    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim intCount As Short

    On Error GoTo LocalErr

    'strSQL = "SELECT ASRSysTables.TableID, " & _
    '"       ASRSysTables.TableName, " & _
    '"       ASRSysColumns.ColumnID, " & _
    '"       ASRSysColumns.ColumnName, " & _
    '"       ASRSysOrderItems.Ascending " & _
    '" FROM ASRSysOrderItems" & _
    '" JOIN ASRSysColumns ON (ASRSysOrderItems.ColumnID = ASRSysColumns.ColumnID)" & _
    '" JOIN ASRSysTables ON (ASRSysColumns.TableID = ASRSysTables.TableID)" & _
    '" WHERE Type = 'O' AND OrderID = " & CStr(mlngDefOrderID) & _
    '" ORDER BY Sequence"

    strSQL = "SELECT ASRSysTables.TableID, " & "       ASRSysTables.TableName, " & "       ASRSysColumns.ColumnID, " & "       ASRSysColumns.ColumnName, " & "       ASRSysMailMergeColumns.SortOrder " & "FROM ASRSysMailMergeColumns " & "JOIN ASRSysColumns ON (ASRSysMailMergeColumns.ColumnID = ASRSysColumns.ColumnID) " & "JOIN ASRSysTables ON (ASRSysColumns.TableID = ASRSysTables.TableID) " & "WHERE ASRSysMailMergeColumns.MailMergeID = " & CStr(mlngMailMergeID) & " " & "  AND SortOrderSequence > 0 " & "ORDER BY SortOrderSequence"

    rsTemp = mclsData.OpenRecordset(strSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)


    With rsTemp
      Do While Not .EOF
        Call SQLAddColumn(mstrSQLOrder, .Fields("TableID").Value, .Fields("TableName").Value, .Fields("ColumnName").Value, vbNullString)

        mstrSQLOrder = mstrSQLOrder & IIf(Left(.Fields("SortOrder").Value, 1) = "A", " ASC", " DESC")

        If fOK = False Then
          Exit Sub
        End If
        .MoveNext()
      Loop
    End With

    If mstrSQLOrder <> vbNullString Then
      mstrSQLOrder = " ORDER BY " & mstrSQLOrder
    End If

    'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsTemp = Nothing

    Exit Sub

LocalErr:
    mstrStatusMessage = "Error building 'Order By' clause"
    fOK = False

  End Sub


  Private Sub SQLAddCalculation(ByRef lngExpID As Integer, ByRef strColCode As String)

    Dim lngCalcViews(,) As Integer
    Dim objCalcExpr As clsExprExpression
    Dim intCount As Short
    Dim blnFound As Boolean
    Dim intNextIndex As Short
    Dim sCalcCode As String
    Dim sSource As String
    Dim lngTestTableID As Integer
    Dim objTableView As CTablePrivilege

    ReDim lngCalcViews(2, 0)
    objCalcExpr = New clsExprExpression
    fOK = objCalcExpr.Initialise(mlngDefBaseTableID, lngExpID, modExpression.ExpressionTypes.giEXPR_RUNTIMECALCULATION, modExpression.ExpressionValueTypes.giEXPRVALUE_UNDEFINED)
    If fOK Then
      fOK = objCalcExpr.RuntimeCalculationCode(lngCalcViews, sCalcCode, True, False, mvarPrompts)

      If fOK And gbEnableUDFFunctions Then
        fOK = objCalcExpr.UDFCalculationCode(lngCalcViews, mastrUDFsRequired, True)
      End If

    End If

    If fOK = False Then
      'mstrStatusMessage = "You do not have permission to run a calculation contained in this mail merge."
      mstrStatusMessage = "You do not have permission to use the '" & Trim(objCalcExpr.Name) & "' calculation."
      'UPGRADE_NOTE: Object objCalcExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
      objCalcExpr = Nothing
      Exit Sub
    End If
    'UPGRADE_NOTE: Object objCalcExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    objCalcExpr = Nothing


    mstrSQLSelect = mstrSQLSelect & IIf(mstrSQLSelect <> vbNullString, ", ", vbNullString) & sCalcCode & " AS '" & strColCode & "'"


    ' Add the required views to the JOIN code.
    For intCount = 1 To UBound(lngCalcViews, 2)
      If lngCalcViews(1, intCount) = 1 Then
        ' Check if view has already been added to the array
        blnFound = False
        For intNextIndex = 1 To UBound(mlngTableViews, 2)
          If mlngTableViews(1, intNextIndex) = 1 And mlngTableViews(2, intNextIndex) = lngCalcViews(2, intCount) Then
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

          objTableView = gcoTablePrivileges.FindViewID(lngCalcViews(2, intCount))
          'sSource = gcoTablePrivileges.FindViewID(lngCalcViews(2, intCount)).RealSource
          sSource = objTableView.RealSource

          'TM20020904 Fault 4364 - depending on whether the table that is about to
          '                        joined is a Parent or Child denotes which ID
          '                        columns are used to establish the join.
          If datGeneral.IsAParentOf((objTableView.TableID), mlngDefBaseTableID) Then
            'Table/View is parent of Base Table.
            mstrSQLJoin = mstrSQLJoin & vbNewLine & " LEFT OUTER JOIN " & sSource & " ON " & mstrSQLFrom & ".ID_" & CStr(objTableView.TableID) & " = " & sSource & ".ID"

          ElseIf datGeneral.IsAChildOf((objTableView.TableID), mlngDefBaseTableID) Then
            'Table/View is child of Base Table.
            ' JPD20021028 Fault  4661
            'mstrSQLJoin = mstrSQLJoin & vbNewLine & _
            '" LEFT OUTER JOIN " & sSource & _
            '" ON " & mstrSQLFrom & ".ID = " & sSource & ".ID_" & CStr(mlngDefBaseTableID)

            '        Else
            '          mstrSQLJoin = mstrSQLJoin & vbNewLine & _
            ''            " LEFT OUTER JOIN " & sSource & _
            ''            " ON " & mstrSQLFrom & ".ID = " & sSource & ".ID"
            '
          End If

          'mstrWhereIDs = mstrWhereIDs & _
          'IIf(mstrWhereIDs <> vbNullString, " OR ", vbNullString) & _
          'mstrSQLFrom & ".ID IN (SELECT ID FROM " & sSource & ")"
        End If

      ElseIf lngCalcViews(1, intCount) = 0 Then
        ' Check if table has already been added to the array
        blnFound = False
        For intNextIndex = 1 To UBound(mlngTableViews, 2)
          'TM20020827 Fault 4339
          If mlngTableViews(1, intNextIndex) = 0 And mlngTableViews(2, intNextIndex) = lngCalcViews(2, intCount) Then
            blnFound = True
            Exit For
          End If
        Next intNextIndex

        'TM20020827 Fault 4339
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

          objTableView = gcoTablePrivileges.FindTableID(lngCalcViews(2, intCount))
          'sSource = gcoTablePrivileges.FindTableID(lngCalcViews(2, intCount)).RealSource
          sSource = objTableView.RealSource

          'TM20020904 Fault 4364 - depending on whether the table that is about to
          '                        joined is a Parent or Child denotes which ID
          '                        columns are used to establish the join.
          If datGeneral.IsAParentOf((objTableView.TableID), mlngDefBaseTableID) Then
            'Table/View is parent of Base Table.
            mstrSQLJoin = mstrSQLJoin & vbNewLine & " LEFT OUTER JOIN " & sSource & " ON " & mstrSQLFrom & ".ID_" & lngCalcViews(2, intCount) & " = " & sSource & ".ID"

          ElseIf datGeneral.IsAChildOf((objTableView.TableID), mlngDefBaseTableID) Then
            'Table/View is child of Base Table.
            ' JPD20021028 Fault 4661
            'mstrSQLJoin = mstrSQLJoin & vbNewLine & _
            '" LEFT OUTER JOIN " & sSource & _
            '" ON " & mstrSQLFrom & ".ID = " & sSource & ".ID_" & CStr(mlngDefBaseTableID)

          End If

        End If

      End If
    Next

  End Sub


  Private Function GetEmailAddress(ByRef lngRecordID As Integer) As String

    ' Return TRUE if the user has been granted the given permission.
    Dim cmADO As ADODB.Command
    Dim pmADO As ADODB.Parameter

    On Error GoTo LocalErr

    ' Check if the user can create New instances of the given category.
    cmADO = New ADODB.Command
    With cmADO
      .CommandText = "dbo.spASRSysEmailAddr"
      .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
      .CommandTimeout = 0
      .ActiveConnection = gADOCon

      pmADO = .CreateParameter("Result", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 8000)
      .Parameters.Append(pmADO)

      pmADO = .CreateParameter("EmailID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
      .Parameters.Append(pmADO)
      pmADO.Value = mlngDefEmailAddrCalc

      pmADO = .CreateParameter("RecordID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
      .Parameters.Append(pmADO)
      pmADO.Value = lngRecordID

      cmADO.Execute()

      'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
      GetEmailAddress = IIf(IsDBNull(.Parameters(0).Value), vbNullString, .Parameters(0).Value)
    End With
    'UPGRADE_NOTE: Object cmADO may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    cmADO = Nothing

    Exit Function

LocalErr:
    'MsgBox "Error reading email details" & vbCr & "(" & Err.Description & ")", vbExclamation
    'UPGRADE_NOTE: Object cmADO may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    cmADO = Nothing

  End Function


  Private Function ValidEmailAddress(ByRef strEmailAddress As Object) As Boolean
    'Must only contain one @ sign and must be something in front of @ and after @

    Dim varTemp As Object

    ValidEmailAddress = False

    'UPGRADE_WARNING: Couldn't resolve default property of object strEmailAddress. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    'UPGRADE_WARNING: Couldn't resolve default property of object varTemp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    varTemp = Split(Trim(strEmailAddress), "@")

    If UBound(varTemp) = 1 Then
      'UPGRADE_WARNING: Couldn't resolve default property of object strEmailAddress. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      If Left(strEmailAddress, CInt("1")) <> "@" And Right(strEmailAddress, CInt("1")) <> "@" Then
        'Check for full stop after @ sign
        'UPGRADE_WARNING: Couldn't resolve default property of object varTemp(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ValidEmailAddress = (InStr(varTemp(1), ".") > 0)
      End If
    End If

  End Function


  Private Function GetRecordDesc(ByRef lngRecordID As Integer) As Object

    ' Return TRUE if the user has been granted the given permission.
    Dim cmADO As ADODB.Command
    Dim pmADO As ADODB.Parameter

    On Error GoTo LocalErr

    If mlngRecordDescExprID < 1 Then
      'UPGRADE_WARNING: Couldn't resolve default property of object GetRecordDesc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      GetRecordDesc = "Record Description Undefined"
      Exit Function
    End If


    ' Check if the user can create New instances of the given category.
    cmADO = New ADODB.Command
    With cmADO
      .CommandText = "dbo.sp_ASRExpr_" & mlngRecordDescExprID
      .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
      .CommandTimeout = 0
      .ActiveConnection = gADOCon

      pmADO = .CreateParameter("Result", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, VARCHAR_MAX_Size)
      .Parameters.Append(pmADO)

      pmADO = .CreateParameter("RecordID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
      .Parameters.Append(pmADO)
      pmADO.Value = lngRecordID

      cmADO.Execute()

      'UPGRADE_WARNING: Couldn't resolve default property of object GetRecordDesc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      GetRecordDesc = .Parameters(0).Value
    End With
    'UPGRADE_NOTE: Object cmADO may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    cmADO = Nothing


    'UPGRADE_WARNING: Couldn't resolve default property of object GetRecordDesc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    If Trim(GetRecordDesc) = vbNullString Then
      'UPGRADE_WARNING: Couldn't resolve default property of object GetRecordDesc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      GetRecordDesc = "Record Description Undefined"
    End If

    Exit Function

LocalErr:
    mstrStatusMessage = "Error reading record description " & "(ID = " & CStr(lngRecordID) & ", Record Description = " & CStr(mlngRecordDescExprID)
    fOK = False

  End Function


  Private Function FormatData(ByRef strInput As Object, ByRef intColIndex As Short) As String


    '07/08/2001 MH Check for Nulls...
    'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
    FormatData = IIf(IsDBNull(strInput), vbNullString, strInput)

    FormatData = Replace(FormatData, vbCr, " ")
    FormatData = Replace(FormatData, vbLf, "")
    FormatData = Replace(FormatData, vbTab, " ")
    FormatData = Replace(FormatData, Chr(34), "'")

    Select Case mintType(intColIndex)
      Case Declarations.SQLDataType.sqlNumeric
        If mintDecimals(intColIndex) <> 0 Then
          FormatData = VB6.Format(FormatData, "0." & New String("0", mintDecimals(intColIndex)))
        Else
          If mlngSize(intColIndex) > 0 Then
            If FormatData = "0" Then
              FormatData = VB6.Format(FormatData, "0")
            Else
              FormatData = VB6.Format(FormatData, "#")
            End If
          End If
        End If

      Case Declarations.SQLDataType.sqlDate
        FormatData = VB6.Format(FormatData, mstrClientDateFormat)

      Case Declarations.SQLDataType.sqlBoolean
        FormatData = IIf(CBool(FormatData), "Y", "N")

      Case Else
        FormatData = Trim(Replace(FormatData, vbTab, ""))

    End Select


    'Check if has decimal places
    If mlngSize(intColIndex) > 0 Then 'Size restriction
      If mintDecimals(intColIndex) > 0 Then
        If InStr(FormatData, ".") > mlngSize(intColIndex) Then
          FormatData = Left(FormatData, mlngSize(intColIndex)) & Mid(FormatData, InStr(FormatData, "."))
        End If

      Else
        If Len(FormatData) > mlngSize(intColIndex) Then
          FormatData = Left(FormatData, mlngSize(intColIndex))
        End If

      End If
    End If

  End Function

  Public Sub EventLogChangeHeaderStatus(ByRef lngStatus As Integer)
    mobjEventLog.ChangeHeaderStatus(lngStatus, mlngSuccessCount, mlngFailCount)
  End Sub

  Public Function SetPromptedValues(ByRef pavPromptedValues As Object) As Boolean

    ' Purpose : This function calls the individual functions that
    '           generate the components of the main SQL string.
    On Error GoTo ErrorTrap

    Dim fOK As Boolean
    Dim iLoop As Short
    Dim iDataType As Short
    Dim lngComponentID As Integer

    fOK = True

    ReDim mvarPrompts(1, 0)

    If IsArray(pavPromptedValues) Then
      ReDim mvarPrompts(1, UBound(pavPromptedValues, 2))

      For iLoop = 0 To UBound(pavPromptedValues, 2)
        ' Get the prompt data type.
        'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Len(Trim(Mid(pavPromptedValues(0, iLoop), 10))) > 0 Then
          'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          lngComponentID = CInt(Mid(pavPromptedValues(0, iLoop), 10))
          'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          iDataType = CShort(Mid(pavPromptedValues(0, iLoop), 8, 1))

          'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(0, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          mvarPrompts(0, iLoop) = lngComponentID

          ' NB. Locale to server conversions are done on the client.
          Select Case iDataType
            Case 2
              ' Numeric.
              'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              mvarPrompts(1, iLoop) = CDbl(pavPromptedValues(1, iLoop))
            Case 3
              ' Logic.
              'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              mvarPrompts(1, iLoop) = (UCase(CStr(pavPromptedValues(1, iLoop))) = "TRUE")
            Case 4
              ' Date.
              ' JPD 20040212 Fault 8082 - DO NOT CONVERT DATE PROMPTED VALUES
              ' THEY ARE PASSED IN FROM THE ASPs AS STRING VALUES IN THE CORRECT
              ' FORMAT (mm/dd/yyyy) AND DOING ANY KIND OF CONVERSION JUST SCREWS
              ' THINGS UP.
              'mvarPrompts(1, iLoop) = CDate(pavPromptedValues(1, iLoop))
              'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              mvarPrompts(1, iLoop) = pavPromptedValues(1, iLoop)
            Case Else
              'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              mvarPrompts(1, iLoop) = CStr(pavPromptedValues(1, iLoop))
          End Select
        Else
          'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(0, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          mvarPrompts(0, iLoop) = 0
        End If
      Next iLoop
    End If

    SetPromptedValues = fOK

    Exit Function

ErrorTrap:
    mstrStatusMessage = "Error whilst setting prompted values. " & Err.Description
    fOK = False
    SetPromptedValues = False

  End Function


  Private Function IsRecordSelectionValid() As Boolean

    Dim sSQL As String
    Dim lCount As Integer
    Dim rsTemp As ADODB.Recordset
    Dim iResult As modUtilityAccess.RecordSelectionValidityCodes
    Dim fCurrentUserIsSysSecMgr As Boolean

    fCurrentUserIsSysSecMgr = CurrentUserIsSysSecMgr()

    ' Filter
    If mlngDefFilterID > 0 Then
      iResult = ValidateRecordSelection(modUtilityAccess.RecordSelectionTypes.REC_SEL_FILTER, mlngDefFilterID)
      Select Case iResult
        Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
          mstrStatusMessage = "The base table filter used in this definition has been deleted by another user."
        Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
          mstrStatusMessage = "The base table filter used in this definition is invalid."
        Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
          If Not fCurrentUserIsSysSecMgr Then
            mstrStatusMessage = "The base table filter used in this definition has been made hidden by another user."
          End If
      End Select
    ElseIf mlngDefPickListID > 0 Then
      iResult = ValidateRecordSelection(modUtilityAccess.RecordSelectionTypes.REC_SEL_PICKLIST, mlngDefPickListID)
      Select Case iResult
        Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
          mstrStatusMessage = "The base table picklist used in this definition has been deleted by another user."
        Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
          mstrStatusMessage = "The base table picklist used in this definition is invalid."
        Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
          If Not fCurrentUserIsSysSecMgr Then
            mstrStatusMessage = "The base table picklist used in this definition has been made hidden by another user."
          End If
      End Select
    End If

    '******* Check calculations for hidden/deleted elements *******
    If Len(mstrStatusMessage) = 0 Then
      sSQL = "SELECT * FROM ASRSysMailMergeColumns " & "WHERE MailMergeID = " & mlngMailMergeID & " AND LOWER(Type) = 'e' "

      rsTemp = mclsGeneral.GetRecords(sSQL)
      With rsTemp
        If Not (.EOF And .BOF) Then
          .MoveFirst()
          Do Until .EOF
            iResult = ValidateCalculation(.Fields("ColumnID").Value)
            Select Case iResult
              Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
                mstrStatusMessage = "A calculation used in this definition has been deleted by another user."
              Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
                mstrStatusMessage = "A calculation used in this definition is invalid."
              Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
                If Not fCurrentUserIsSysSecMgr Then
                  mstrStatusMessage = "A calculation used in this definition has been made hidden by another user."
                End If
            End Select

            If Len(mstrStatusMessage) > 0 Then
              Exit Do
            End If

            .MoveNext()
          Loop
        End If
      End With

      'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
      rsTemp = Nothing
    End If

    IsRecordSelectionValid = (Len(mstrStatusMessage) = 0)

  End Function


  Public Function UDFFunctions(ByRef pbCreate As Object) As Object

    On Error GoTo UDFFunctions_ERROR

    Dim iCount As Short
    Dim strDropCode As String
    Dim strFunctionName As String
    Dim sUDFCode As String
    Dim datData As clsDataAccess
    Dim iStart As Short
    Dim iEnd As Short
    Dim strFunctionNumber As String

    Const FUNCTIONPREFIX As String = "udf_ASRSys_"

    If gbEnableUDFFunctions Then

      For iCount = 1 To UBound(mastrUDFsRequired)

        'JPD 20060110 Fault 10509
        'iStart = Len("CREATE FUNCTION udf_ASRSys_") + 1
        iStart = InStr(mastrUDFsRequired(iCount), FUNCTIONPREFIX) + Len(FUNCTIONPREFIX)
        iEnd = InStr(1, Mid(mastrUDFsRequired(iCount), 1, 1000), "(@Per")
        strFunctionNumber = Mid(mastrUDFsRequired(iCount), iStart, iEnd - iStart)
        strFunctionName = FUNCTIONPREFIX & strFunctionNumber

        'Drop existing function (could exist if the expression is used more than once in a report)
        strDropCode = "IF EXISTS" & " (SELECT *" & "   FROM sysobjects" & "   WHERE id = object_id('[" & Replace(gsUsername, "'", "''") & "]." & strFunctionName & "')" & "     AND sysstat & 0xf = 0)" & " DROP FUNCTION [" & gsUsername & "]." & strFunctionName
        mclsData.ExecuteSql(strDropCode)

        ' Create the new function
        'UPGRADE_WARNING: Couldn't resolve default property of object pbCreate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If pbCreate Then
          sUDFCode = mastrUDFsRequired(iCount)
          mclsData.ExecuteSql(sUDFCode)
        End If

      Next iCount
    End If

    'UPGRADE_WARNING: Couldn't resolve default property of object UDFFunctions. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    UDFFunctions = True
    Exit Function

UDFFunctions_ERROR:
    mstrStatusMessage = "Error within filters/calculation UDFs (" & Err.Description & ")"
    'UPGRADE_WARNING: Couldn't resolve default property of object UDFFunctions. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    UDFFunctions = False

  End Function
End Class