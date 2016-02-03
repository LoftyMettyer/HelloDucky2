VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExportRUN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' To hold Properties
Private mlngExportID As Long
Private mstrErrorString As String
Private mblnEnableSQLTable As Boolean
Private mbNoRecords As Boolean

' String to hold the temp table name
Private mstrTempTableName As String
Private mstrArrayHeader() As String
Private mstrArrayData() As String
Private mstrArrayFooter() As String

' Variables to store definition
Private mstrExportName As String
Private mstrExportDescription As String
Private mlngExportBaseTable As Long
Private mstrExportBaseTableName As String
Private mlngExportAllRecords As Long
Private mlngExportPickListID As Long
Private mlngExportFilterID As Long
Private mlngExportParent1Table As Long
Private mstrExportParent1TableName As String
Private mlngExportParent1FilterID As Long
Private mlngExportParent2Table As Long
Private mstrExportParent2TableName As String
Private mlngExportParent2FilterID As Long
Private mlngExportChildTable As Long
Private mstrExportChildTableName As String
Private mlngExportChildFilterID As Long
Private mlngExportChildMaxRecords As Long
'Private mstrExportOutputType As String
'Private mstrOutputFilename As String
Private mblnExportQuotes As Boolean
Private mblnStripDelimiter As Boolean
'Private mblnExportHeader As Boolean
Private mintExportHeader As Integer
Private mstrExportHeaderText As String
Private mintExportFooter As Integer
Private mstrExportFooterText As String
Private mstrExportDelimiter As String
Private mstrExportOtherDelimiter As String
Private mstrExportActualDelimiter As String
Private mstrExportDateFormat As String
Private mstrExportDateSeparator As String
'Private mstrExportDateYearDigits As String
Private mlngExportRecordCount As Long

'Private mbAppendToFile As Boolean                   ' Append export to existing file
Private mblnHeader As Boolean
Private mblnFooter As Boolean
Private mbForceHeader As Boolean                    ' Always force header if no records found
Private mbOmitHeader As Boolean                     ' Omit header if appending to file
Private mbAuditChangesOnly As Boolean

Private mbSplitFile As Boolean
Private mlngSplitFileSize As Long

Private mlngOutputFormat As Long
Private mblnOutputSave As Boolean
Private mlngOutputSaveExisting As Long
'Private mlngOutputSaveFormat As Long
Private mblnOutputEmail As Boolean
Private mlngOutputEmailAddr As Long
Private mstrOutputEmailSubject As String
Private mstrOutputEmailAttachAs As String
'Private mlngOutputEmailFileFormat As Long
Private mstrOutputFileName As String

Private mlngExportParent1AllRecords As Long
Private mlngExportParent1PickListID As Long
Private mlngExportParent2AllRecords As Long
Private mlngExportParent2PickListID As Long

Private mstrLinkedServer As String
Private mstrLinkedCatalog As String
Private mstrLinkedTable As String

' Recordsets to store the definition/column information and the final output data
Private mrstExportDetails As New Recordset
Private mrstExportOutput As New Recordset

' Strings to hold the SQL statement
Private mstrSQLSelect As String
Private mstrSQLFrom As String
Private mstrSQLJoin As String
Private mstrSQLWhere As String
Private mstrSQLOrderBy As String
Private mstrSQL As String
Private mstrOnlyChangesFilter As String

Private mstrTransformFile As String
Private mstrXMLDataNodeName As String
Private mstrXSDFilename As String
Private mbPreserveTransformPath As Boolean
Private mbPreserveXSDPath As Boolean
Private mbSplitXMLIntoFiles As Boolean

' Data access classes
Private mclsData As clsDataAccess
Private mclsGeneral As clsGeneral

' Array holding the columns to sort the report by
Private mvarSortOrder() As Variant

' Array to hold the columns used in the export
Dim mvarColDetails() As Variant

' TableViewsGuff
Private mstrRealSource As String
Private mstrBaseTableRealSource As String
Private mlngTableViews() As Long
Private mstrViews() As String
Private mobjTableView As CTablePrivilege
Private mobjColumnPrivileges As CColumnPrivileges

' Batch mode ?
'Private gblnBatchMode As Boolean

Private mblnUserCancelled As Boolean
Private mblnNoRecords As Boolean

' CMG File Code
Private mstrExportFileCode As String
Private mlngCMGRecordIdentifier As Long
Private mbUpdateAuditLog As Boolean
Private mdLastSuccessfulOutput As Date

Private mdExportCreateDate As Date

Private Enum CMGFields
  NewValue = 0    ' "NewValue"
  DateTime = 1    ' "DateTimeStamp"
  ColumnID = 2    ' "ColumnID"
End Enum

Private mbDefinitionOwner As Boolean
Private mbLoggingExportSuccess As Boolean

Private mlSuccessfulRecords As Long

' Array holding the User Defined functions that are needed for this report
Private mastrUDFsRequired() As String

  
Public Property Get UserCancelled() As Boolean
  UserCancelled = mblnUserCancelled
End Property

Public Property Get ExportID() As Long
  ExportID = mlngExportID
End Property

Public Property Let ExportID(ByVal lNewID As Long)
  mlngExportID = lNewID
End Property

Public Property Get ErrorString() As String
  ErrorString = mstrErrorString
End Property

Public Property Let ErrorString(ByVal sErrorString As String)
  mstrErrorString = sErrorString
End Property

Private Sub Class_Initialize()
  
  ' Purpose : Sets references to other classes and redimensions arrays
  '           used for table usage information
  ' Input   : None
  ' Output  : None
  
  Set mclsData = New DataMgr.clsDataAccess
  Set mclsGeneral = New DataMgr.clsGeneral
  ReDim mvarSortOrder(2, 0)
  'ReDim mvarColDetails(6, 0)
'  ReDim mvarColDetails(10, 0)
'  ReDim mvarColDetails(15, 0)
'NPG20071217 Fault 12867
  ' ReDim mvarColDetails(16, 0)
  'NPG20080617 Suggestion S000816
  ReDim mvarColDetails(17, 0)
  ReDim mlngTableViews(2, 0)
  ReDim mstrViews(0)

  mblnEnableSQLTable = (GetSystemSetting("Output", "ExportToSQLTable", 0) = 1)

End Sub

Private Sub Class_Terminate()
  
  ' Purpose : Clears references to other classes.
  ' Input   : None
  ' Output  : None
  
  Set mclsData = Nothing
  Set mclsGeneral = Nothing
  
End Sub

Private Function IsRecordSelectionValid() As Boolean

  Dim sSQL As String
  Dim lCount As Long
  Dim rsTemp As Recordset
  Dim iResult As RecordSelectionValidityCodes
  
  ' Filter
  If mlngExportFilterID > 0 Then
    iResult = ValidateRecordSelection(REC_SEL_FILTER, mlngExportFilterID)
    Select Case iResult
      Case REC_SEL_VALID_DELETED
        mstrErrorString = "The base table filter used in this definition has been deleted by another user."
      Case REC_SEL_VALID_INVALID
        mstrErrorString = "The base table filter used in this definition is invalid."
      Case REC_SEL_VALID_HIDDENBYOTHER
        If Not gfCurrentUserIsSysSecMgr Then
          mstrErrorString = "The base table filter used in this definition has been made hidden by another user."
        End If
    End Select
  ElseIf mlngExportPickListID > 0 Then
    iResult = ValidateRecordSelection(REC_SEL_PICKLIST, mlngExportPickListID)
    Select Case iResult
      Case REC_SEL_VALID_DELETED
        mstrErrorString = "The base table picklist used in this definition has been deleted by another user."
      Case REC_SEL_VALID_INVALID
        mstrErrorString = "The base table picklist used in this definition is invalid."
      Case REC_SEL_VALID_HIDDENBYOTHER
        If Not gfCurrentUserIsSysSecMgr Then
          mstrErrorString = "The base table picklist used in this definition has been made hidden by another user."
        End If
      End Select
  End If
 
  If Len(mstrErrorString) = 0 Then
    If mlngExportParent1FilterID > 0 Then
      iResult = ValidateRecordSelection(REC_SEL_FILTER, mlngExportParent1FilterID)
      Select Case iResult
        Case REC_SEL_VALID_DELETED
          mstrErrorString = "The first parent table filter used in this definition has been deleted by another user."
        Case REC_SEL_VALID_INVALID
          mstrErrorString = "The first parent table filter used in this definition is invalid."
        Case REC_SEL_VALID_HIDDENBYOTHER
          If Not gfCurrentUserIsSysSecMgr Then
            mstrErrorString = "The first parent table filter used in this definition has been made hidden by another user."
          End If
      End Select
    ElseIf mlngExportParent1PickListID > 0 Then
      iResult = ValidateRecordSelection(REC_SEL_PICKLIST, mlngExportParent1PickListID)
      Select Case iResult
        Case REC_SEL_VALID_DELETED
          mstrErrorString = "The first parent table picklist used in this definition has been deleted by another user."
        Case REC_SEL_VALID_INVALID
          mstrErrorString = "The first parent table picklist used in this definition is invalid."
        Case REC_SEL_VALID_HIDDENBYOTHER
          If Not gfCurrentUserIsSysSecMgr Then
            mstrErrorString = "The first parent table picklist used in this definition has been made hidden by another user."
          End If
        End Select
    End If
  End If
  
  If Len(mstrErrorString) = 0 Then
    If mlngExportParent2FilterID > 0 Then
      iResult = ValidateRecordSelection(REC_SEL_FILTER, mlngExportParent2FilterID)
      Select Case iResult
        Case REC_SEL_VALID_DELETED
          mstrErrorString = "The second parent table filter used in this definition has been deleted by another user."
        Case REC_SEL_VALID_INVALID
          mstrErrorString = "The second parent table filter used in this definition is invalid."
        Case REC_SEL_VALID_HIDDENBYOTHER
          If Not gfCurrentUserIsSysSecMgr Then
            mstrErrorString = "The second parent table filter used in this definition has been made hidden by another user."
          End If
      End Select
    ElseIf mlngExportParent2PickListID > 0 Then
      iResult = ValidateRecordSelection(REC_SEL_PICKLIST, mlngExportParent2PickListID)
      Select Case iResult
        Case REC_SEL_VALID_DELETED
          mstrErrorString = "The second parent table picklist used in this definition has been deleted by another user."
        Case REC_SEL_VALID_INVALID
          mstrErrorString = "The second parent table picklist used in this definition is invalid."
        Case REC_SEL_VALID_HIDDENBYOTHER
          If Not gfCurrentUserIsSysSecMgr Then
            mstrErrorString = "The second parent table picklist used in this definition has been made hidden by another user."
          End If
        End Select
    End If
  End If
  
  If Len(mstrErrorString) = 0 Then
    If mlngExportChildFilterID > 0 Then
      iResult = ValidateRecordSelection(REC_SEL_FILTER, mlngExportChildFilterID)
      Select Case iResult
        Case REC_SEL_VALID_DELETED
          mstrErrorString = "The child table filter used in this definition has been deleted by another user."
        Case REC_SEL_VALID_INVALID
          mstrErrorString = "The child table filter used in this definition is invalid."
        Case REC_SEL_VALID_HIDDENBYOTHER
          If Not gfCurrentUserIsSysSecMgr Then
            mstrErrorString = "The child table filter used in this definition has been made hidden by another user."
          End If
      End Select
    End If
  End If
  
  
'******* Check calculations for hidden/deleted elements *******

  If Len(mstrErrorString) = 0 Then
    sSQL = "SELECT * FROM ASRSYSExportDetails " & _
             "WHERE ExportID = " & mlngExportID & _
             " AND LOWER(Type) = 'x' "
  
    Set rsTemp = datGeneral.GetRecords(sSQL)
    With rsTemp
      If Not (.EOF And .BOF) Then
        .MoveFirst
        Do Until .EOF
          iResult = ValidateCalculation(!ColExprID)
          Select Case iResult
            Case REC_SEL_VALID_DELETED
              mstrErrorString = "A calculation used in this definition has been deleted by another user."
            Case REC_SEL_VALID_INVALID
              mstrErrorString = "A calculation used in this definition is invalid."
            Case REC_SEL_VALID_HIDDENBYOTHER
              If Not gfCurrentUserIsSysSecMgr Then
                mstrErrorString = "A calculation used in this definition has been made hidden by another user."
              End If
          End Select
          
          If Len(mstrErrorString) > 0 Then
            Exit Do
          End If
          
          .MoveNext
        Loop
      End If
    End With
    
    Set rsTemp = Nothing
  End If
  
  IsRecordSelectionValid = (Len(mstrErrorString) = 0)
  
End Function


Public Function RunExport() As Boolean

  ' Purpose : This function is called from frmDefsel and Batch Jobs.
  '           It is the main function which runs the export.
  ' Input   : bSilent, Boolean, Suppress COAMsgBoxs ? (ie, for batch jobs)
  ' Output  : True/False Success

  On Error GoTo RunExport_ERROR
  
  Dim fOK As Boolean
  Dim sToday As String
  
  mdExportCreateDate = Now
  mstrErrorString = vbNullString
  mblnNoRecords = False
  
  fOK = True
  Screen.MousePointer = vbHourglass
 
  ' JDM - 07/10/01 - Fault 2644 - Cannot run CMG export if security is removed
  If fOK Then fOK = GetExportDefinition

  mbLoggingExportSuccess = CBool(GetUserSetting("LogEvents", "Export_Success", False))
  
  If fOK Then
    With gobjProgress
      '.AviFile = App.Path & "\videos\export.avi"
      .AVI = dbText
      .MainCaption = "Export"
      If gblnBatchMode = False Then
        .NumberOfBars = 1
        .Caption = "Export"
        .Time = False
        .Cancel = True
        '.Bar1MaxValue = 9
        .OpenProgress
      Else
        .ResetBar2
        '.Bar2MaxValue = 9
      End If
    End With
    
    
    'If mstrExportOutputType = "C" And Not datGeneral.SystemPermission("CMG", "CMGRUN") Then
    If mlngOutputFormat = fmtCMGFile And Not datGeneral.SystemPermission("CMG", "CMGRUN") Then
      mstrErrorString = vbCrLf & "You do not have permission to run a CMG export" & vbCrLf _
        & "Please contact your system administrator"
        fOK = False
    End If
  
    If gblnBatchMode Then
      gobjProgress.Bar2Caption = "Export : " & mstrExportName
    Else
      gobjProgress.Bar1Caption = "Export : " & mstrExportName
    End If
    
    gobjEventLog.AddHeader eltExport, mstrExportName
  End If
  
  If fOK Then fOK = GetDetailsRecordsets
  If fOK Then fOK = GenerateSQL
  If fOK Then fOK = AddTempTableToSQL
  If fOK Then fOK = MergeSQLStrings
  If fOK Then fOK = UDFFunctions(mastrUDFsRequired, True)
  If fOK Then fOK = ExecuteSql
  If fOK Then fOK = UDFFunctions(mastrUDFsRequired, False)
  If fOK Then fOK = CheckRecordSet
  If fOK Then fOK = ExportData
  
  'If gobjProgress.Visible Then gobjProgress.UpdateProgress gblnBatchMode
  
  If Not gblnBatchMode And gobjProgress.Visible Then gobjProgress.CloseProgress
  
  Call UtilUpdateLastRun(utlExport, mlngExportID)
  
  
  'MH20000705 Fault 519
  'No records is reported as a failure !

'mlngExportRecordCount
  
  If mblnNoRecords Then
    gobjEventLog.ChangeHeaderStatus elsSuccessful, mlngExportRecordCount, 0
    gobjEventLog.AddDetailEntry mstrErrorString
    mstrErrorString = "Completed successfully." & vbCrLf & mlngExportRecordCount & " record(s) exported." & vbCrLf & mstrErrorString
    fOK = True

  ElseIf fOK Then
    gobjEventLog.ChangeHeaderStatus elsSuccessful, mlngExportRecordCount, 0
    mstrErrorString = "Completed successfully." & vbCrLf & mlngExportRecordCount & " record(s) exported."
  
    sToday = "convert(datetime, '" & Replace(Format(Now, "MM/dd/yyyy hh:mm:ss"), UI.GetSystemDateSeparator, "/") & "')"
    mstrSQL = "UPDATE ASRSysExportName SET LastSuccessfulOutput = " & sToday & " WHERE ID = " & mlngExportID
    mclsData.ExecuteSql mstrSQL
  
  ElseIf mblnUserCancelled Then
    gobjEventLog.ChangeHeaderStatus elsCancelled, mlSuccessfulRecords, 0
    mstrErrorString = "Cancelled by user." & vbCrLf & mlSuccessfulRecords & " record(s) exported."
  Else
    'Only details records for failures !
    gobjEventLog.AddDetailEntry mstrErrorString
    gobjEventLog.ChangeHeaderStatus elsFailed, mlSuccessfulRecords, 0
    mstrErrorString = "Failed." & vbCrLf & vbCrLf & mstrErrorString '& vbCrLf & vbCrLf & mlSuccessfulRecords & " record(s) exported."
  End If
    
  mstrErrorString = "Export : '" & mstrExportName & "' " & mstrErrorString
    
  If Not gblnBatchMode Then
    COAMsgBox mstrErrorString, IIf(fOK, vbInformation, vbExclamation) + vbOKOnly, "Export"
  End If
  
  If fOK = True Then fOK = ClearUp Else ClearUp
  
  Screen.MousePointer = vbDefault
  
  RunExport = fOK
  
  Exit Function
  
RunExport_ERROR:

  fOK = False
  mstrErrorString = "Error Whilst Running Export." & vbCrLf & "(" & Err.Description & ")"
  Resume Next
  
End Function

Private Function AddTempTableToSQL() As Boolean

  ' Purpose : This function retrieves a unique temp table name and
  '           inserts it into the SQL Select statement
  ' Input   : None
  ' Output  : True/False Success

  On Error GoTo AddTempTableToSQL_ERROR
  
  'gobjProgress.UpdateProgress gblnBatchMode
  
  ' If user cancels the export, abort
  If gobjProgress.Cancelled Then
    mblnUserCancelled = True
    AddTempTableToSQL = False
    Exit Function
  End If
  
  mstrTempTableName = datGeneral.UniqueSQLObjectName("ASRSysTempExport", 3)
  
  mstrSQLSelect = mstrSQLSelect & " INTO " & "[" & mstrTempTableName & "]"
  
  AddTempTableToSQL = True
  Exit Function
  
AddTempTableToSQL_ERROR:
  
  mstrErrorString = "Error whilst retrieving unique temp table name." & vbCrLf & "(" & Err.Description & ")"
  AddTempTableToSQL = False
  
End Function

Private Function MergeSQLStrings() As Boolean
    
  ' Purpose : This function merges all the SQL string variables
  '           into one long string
  ' Input   : None
  ' Output  : True/False Success

  On Error GoTo MergeSQLStrings_ERROR
  
  'gobjProgress.UpdateProgress gblnBatchMode

  ' If user cancels the export, abort
  If gobjProgress.Cancelled Then
    mblnUserCancelled = True
    MergeSQLStrings = False
    Exit Function
  End If

  mstrSQL = mstrSQLSelect & _
            " FROM " & _
            mstrSQLFrom & _
            IIf(Len(mstrSQLJoin) = 0, "", " " & mstrSQLJoin) & _
            IIf(Len(mstrSQLWhere) = 0, "", " " & mstrSQLWhere) & _
            " " & mstrSQLOrderBy

  MergeSQLStrings = True
  Exit Function
  
MergeSQLStrings_ERROR:

  mstrErrorString = "Error whilst merging SQL string components." & vbCrLf & "(" & Err.Description & ")"
  MergeSQLStrings = False
  
End Function

Private Function ExecuteSql() As Boolean

  ' Purpose : This function executes the SQL string
  ' Input   : None
  ' Output  : True/False Success

  On Error GoTo ExecuteSQL_ERROR
  
  'gobjProgress.UpdateProgress gblnBatchMode
  
  ' If user cancels the export, abort
  If gobjProgress.Cancelled Then
    mblnUserCancelled = True
    ExecuteSql = False
    Exit Function
  End If
  
  'COAMsgBox "Recordset will be generated from : " & vbCrLf & mstrSQL
  mclsData.ExecuteSql mstrSQL

  ExecuteSql = True
  Exit Function
  
ExecuteSQL_ERROR:
  
  mstrErrorString = "Error whilst executing SQL statement." & vbCrLf & "(" & Err.Description & ")"
  ExecuteSql = False
  
End Function

Private Function GetExportDefinition() As Boolean
  
  ' Purpose : This function retrieves the basic definition details
  '           and stores it in module level variables
  ' Input   : None
  ' Output  : True/False Success
  
  On Error GoTo GetExportDefinition_ERROR

  Dim prstTemp_Definition As Recordset
  Dim pstrSQL As String
  Dim rsTemp As Recordset
  
'  gobjProgress.UpdateProgress gblnBatchMode
'
'  ' If user cancels the export, abort
'  If gobjProgress.Cancelled Then
'    mblnUserCancelled = True
'    GetExportDefinition = False
'    Exit Function
'  End If
  
  pstrSQL = "SELECT * FROM AsrSysExportName " & _
           "WHERE ID = " & mlngExportID & " "
  
  Set prstTemp_Definition = mclsData.OpenRecordset(pstrSQL, adOpenForwardOnly, adLockReadOnly)
  
  With prstTemp_Definition
  
    If .BOF And .EOF Then
      GetExportDefinition = False
      mstrErrorString = "Could not find specified Export definition !"
      prstTemp_Definition.Close
      Set prstTemp_Definition = Nothing
      Exit Function
    End If
    
    
    mstrExportName = !Name
    mstrExportDescription = !Description
    mlngExportBaseTable = !BaseTable
    mstrExportBaseTableName = datGeneral.GetTableName(mlngExportBaseTable)
    mlngExportAllRecords = !AllRecords
    mlngExportPickListID = !picklist
    mlngExportFilterID = !Filter
    mlngExportParent1Table = !parent1table
    mstrExportParent1TableName = datGeneral.GetTableName(mlngExportParent1Table)
    mlngExportParent1FilterID = !parent1filter
    mlngExportParent2Table = !parent2table
    mstrExportParent2TableName = datGeneral.GetTableName(mlngExportParent2Table)
    mlngExportParent2FilterID = !parent2filter
    mlngExportChildTable = !ChildTable
    mstrExportChildTableName = datGeneral.GetTableName(mlngExportChildTable)
    mlngExportChildFilterID = !childFilter
    mlngExportChildMaxRecords = !ChildMaxRecords
    'mstrExportOutputType = !outputtype
    'mstrExportOutputName = !outputname
    mblnExportQuotes = !Quotes
    mblnStripDelimiter = True
    
    mlngSplitFileSize = IIf(IsNull(!SplitFileSize), 0, !SplitFileSize)
    mbSplitFile = IIf(IsNull(!SplitFile), False, !SplitFile And mlngSplitFileSize > 0)
    
    mstrExportFileCode = IIf(IsNull(!CMGExportFileCode), "", !CMGExportFileCode)
    mlngCMGRecordIdentifier = IIf(IsNull(!CMGExportRecordID), 0, !CMGExportRecordID)
    mbUpdateAuditLog = IIf(IsNull(!CMGExportUpdateAudit), False, !CMGExportUpdateAudit)

    mlngExportParent1AllRecords = !parent1AllRecords
    mlngExportParent1PickListID = !parent1picklist
    mlngExportParent2AllRecords = !parent2AllRecords
    mlngExportParent2PickListID = !parent2picklist
    
    mstrExportDelimiter = IIf(IsNull(!delimiter), "", !delimiter)
    mstrExportOtherDelimiter = IIf(IsNull(!otherdelimiter), "", !otherdelimiter)
   
    Select Case UCase(mstrExportDelimiter)
      Case ",": mstrExportActualDelimiter = mstrExportDelimiter
      Case "<TAB>": mstrExportActualDelimiter = vbTab
      Case "<OTHER>": mstrExportActualDelimiter = mstrExportOtherDelimiter
      Case Else: mstrExportActualDelimiter = mstrExportDelimiter
    End Select
   
    'mbAppendToFile = IIf(IsNull(!AppendToFile), False, !AppendToFile)
    mbForceHeader = IIf(IsNull(!ForceHeader), False, !ForceHeader)
    mbOmitHeader = IIf(IsNull(!OmitHeader), False, !OmitHeader)
   
    mlngOutputFormat = !OutputFormat
    mblnOutputSave = !OutputSave
    mlngOutputSaveExisting = !OutputSaveExisting
    'mlngOutputSaveFormat = !OutputSaveFormat
    mblnOutputEmail = !OutputEmail
    mlngOutputEmailAddr = !OutputEmailAddr
    mstrOutputEmailSubject = !OutputEmailSubject
    mstrOutputEmailAttachAs = IIf(IsNull(!OutputEmailAttachAs), vbNullString, !OutputEmailAttachAs)
    'mlngOutputEmailFileFormat = !OutputEmailFileFormat
    mstrOutputFileName = !OutputFilename
    mdLastSuccessfulOutput = IIf(IsNull(!LastSuccessfulOutput), "00:00:00", !LastSuccessfulOutput)
   
    mintExportHeader = !Header
    mstrExportHeaderText = IIf(IsNull(!HeaderText), vbNullString, !HeaderText)
    mintExportFooter = !Footer
    mstrExportFooterText = IIf(IsNull(!FooterText), vbNullString, !FooterText)
    mstrXMLDataNodeName = IIf(IsNull(!XMLDataNodeName), vbNullString, !XMLDataNodeName)
    mbSplitXMLIntoFiles = IIf(IsNull(!SplitXMLNodesFile), False, !SplitXMLNodesFile)
    
    If mlngOutputFormat = fmtXML Then
      mstrExportHeaderText = IIf(Len(mstrExportHeaderText) = 0, "root", mstrExportHeaderText)
      mstrXMLDataNodeName = IIf(Len(mstrXMLDataNodeName) = 0, mstrExportBaseTableName, mstrXMLDataNodeName)
    End If
   
    mstrTransformFile = IIf(IsNull(!TransformFile), "", !TransformFile)
    mstrXSDFilename = IIf(IsNull(!XSDFilename), "", !XSDFilename)
    mbPreserveTransformPath = IIf(IsNull(!PreserveTransformPath), False, !PreserveTransformPath)
    mbPreserveXSDPath = IIf(IsNull(!PreserveXSDPath), False, !PreserveXSDPath)

    mstrLinkedServer = IIf(IsNull(!LinkedServer), "", !LinkedServer)
    mstrLinkedCatalog = IIf(IsNull(!LinkedCatalog), "", !LinkedCatalog)
    mstrLinkedTable = IIf(IsNull(!LinkedTable), "", !LinkedTable)

    mbDefinitionOwner = (LCase(Trim(gsUserName)) = LCase(Trim(!userName)))
    mbAuditChangesOnly = IIf(IsNull(!AuditChangesOnly), False, !AuditChangesOnly)

    If Not IsRecordSelectionValid Then
      GetExportDefinition = False
      Exit Function
    End If
    
    If LCase(!dateseparator) = "<none>" Then
      mstrExportDateSeparator = ""
    Else
      mstrExportDateSeparator = !dateseparator
    End If
    
    Select Case !DateFormat
    Case "dmy"
      mstrExportDateFormat = "dd" & mstrExportDateSeparator & _
                             "mm" & mstrExportDateSeparator & _
                             IIf(!Dateyeardigits = "2", "yy", "yyyy")
    Case "mdy"
      mstrExportDateFormat = "mm" & mstrExportDateSeparator & _
                             "dd" & mstrExportDateSeparator & _
                             IIf(!Dateyeardigits = "2", "yy", "yyyy")
    Case "ymd"
      mstrExportDateFormat = IIf(!Dateyeardigits = "2", "yy", "yyyy") & _
                             mstrExportDateSeparator & "mm" & _
                             mstrExportDateSeparator & "dd"
    Case "ydm"
      mstrExportDateFormat = IIf(!Dateyeardigits = "2", "yy", "yyyy") & _
                             mstrExportDateSeparator & "dd" & _
                             mstrExportDateSeparator & "mm"
    End Select
  
  End With
  
  
  If mlngOutputFormat = fmtExcelWorksheet And Not gblnBatchMode Then
    gobjProgress.AVI = dbExcel
  End If
  
  
  GetExportDefinition = True
  
TidyAndExit:
  
  Set prstTemp_Definition = Nothing

Exit Function

GetExportDefinition_ERROR:

  GetExportDefinition = False
  mstrErrorString = "Error whilst reading the Export definition !" & vbCrLf & "(" & Err.Description & ")"
  Resume TidyAndExit

End Function

Private Function GetDetailsRecordsets() As Boolean

  ' Purpose : This function loads export details and sort details into
  '           arrays and leaves the details recordset reference there
  '           (dont remove it...used for summary info !) NB CAN REMOVE FOR EXPORT !?!
  ' Input   : None
  ' Output  : True/False Success

  On Error GoTo GetDetailsRecordsets_ERROR
  
  Dim pstrTempSQL As String
  Dim pintTemp As Integer
  Dim prstExportSortOrder As Recordset
  
  Dim lngTextCount As Long
  Dim lngFillerCount As Long
  Dim lngRecNumCount As Long
  Dim lngCalculationCount As Long
  
  'gobjProgress.UpdateProgress gblnBatchMode
  
  ' If user cancels the export, abort
  If gobjProgress.Cancelled Then
    mblnUserCancelled = True
    GetDetailsRecordsets = False
    Exit Function
  End If
  
  ' Get the column information from the Details table, in order
  
  pstrTempSQL = "SELECT * FROM AsrSysExportDetails WHERE " & _
               "ExportID = " & mlngExportID & " " & _
               "ORDER BY [ID]"
             
  Set mrstExportDetails = mclsData.OpenRecordset(pstrTempSQL, adOpenKeyset, adLockReadOnly)

  With mrstExportDetails
    
    If .BOF And .EOF Then
      GetDetailsRecordsets = False
      mstrErrorString = "No columns found in the specified Export definition." & vbCrLf & "Please remove this definition and create a new one."
      Exit Function
    End If
    
    Do Until .EOF
      pintTemp = UBound(mvarColDetails, 2) + 1
      ReDim Preserve mvarColDetails(UBound(mvarColDetails, 1), pintTemp)
      mvarColDetails(0, pintTemp) = !Type
      mvarColDetails(1, pintTemp) = Trim(!TableID)
      mvarColDetails(2, pintTemp) = datGeneral.GetTableName(CLng(!TableID))
      mvarColDetails(3, pintTemp) = !ColExprID
      If !ColExprID = 0 Or !Type = "X" Then
        mvarColDetails(4, pintTemp) = ""
      Else
        mvarColDetails(4, pintTemp) = datGeneral.GetColumnName(CLng(!ColExprID))
      End If
      mvarColDetails(5, pintTemp) = !Data
      'mvarColDetails(6, pintTemp) = !fillerlength
      mvarColDetails(6, pintTemp) = IIf(!fillerlength > 999999, 999999, !fillerlength)
      mvarColDetails(7, pintTemp) = datGeneral.DateColumn(!Type, !TableID, !ColExprID)

      mvarColDetails(8, pintTemp) = IIf(IsNull(!CMGColumnCode), "", !CMGColumnCode) ' The CMG Code for this column
      mvarColDetails(9, pintTemp) = datGeneral.IsColumnAudited(!ColExprID)          ' Is this column audited
      mvarColDetails(10, pintTemp) = True                                           ' Do we export this column
      
      'TM20011010 Fault 2197
      mvarColDetails(11, pintTemp) = !Decimals
      
      mvarColDetails(12, pintTemp) = datGeneral.NumericColumn(!Type, !TableID, !ColExprID)

      
      'Default to old heading
      '(still need to increase counter even if heading is overwritten!)
      Select Case !Type
      Case "F", "R"   'Filler or Carriage Return
        lngFillerCount = lngFillerCount + 1
        mvarColDetails(13, pintTemp) = "Filler" & CStr(lngFillerCount)
      Case "T"        'Text
        lngTextCount = lngTextCount + 1
        mvarColDetails(13, pintTemp) = "Text" & CStr(lngTextCount)
      Case "N"        'Record Number
        lngRecNumCount = lngRecNumCount + 1
        mvarColDetails(13, pintTemp) = "Record Number" & CStr(lngRecNumCount)
      Case "X"        'Calc
        lngCalculationCount = lngCalculationCount + 1
        mvarColDetails(13, pintTemp) = "Calculation" & CStr(lngCalculationCount)
      End Select

      'MH20030120
      If Not IsNull(!Heading) Then
        mvarColDetails(13, pintTemp) = !Heading   'might be null but thats okay!
      End If
      mvarColDetails(14, pintTemp) = Replace(GetUniqueHeading(pintTemp), "'", "''")

      ' Does this column use digit seperators
      mvarColDetails(15, pintTemp) = datGeneral.DoesColumnUseSeparators(!ColExprID)
      
      'NPG20071217 Fault 12867
      mvarColDetails(16, pintTemp) = !ConvertCase
      
      'NPG20080617 Suggestion S000816
      mvarColDetails(17, pintTemp) = !SuppressNulls
            
      .MoveNext
    Loop
  .MoveFirst
  End With
  
  'Add in the ID of the record if we are a cmg output
  'If mstrExportOutputType = "C" Then
  If mlngOutputFormat = fmtCMGFile Or mlngOutputFormat = fmtXML Then
    pintTemp = UBound(mvarColDetails, 2) + 1
    ReDim Preserve mvarColDetails(UBound(mvarColDetails, 1), pintTemp)
    mvarColDetails(0, pintTemp) = "C"
    mvarColDetails(1, pintTemp) = mlngExportBaseTable
    mvarColDetails(2, pintTemp) = mstrExportBaseTableName
    mvarColDetails(3, pintTemp) = 0
    mvarColDetails(4, pintTemp) = "ID"
    mvarColDetails(5, pintTemp) = "ID"
    mvarColDetails(6, pintTemp) = 0
    mvarColDetails(7, pintTemp) = ""
    mvarColDetails(8, pintTemp) = ""    ' The CMG Code for this column
    mvarColDetails(9, pintTemp) = False ' Is this column audited
    mvarColDetails(10, pintTemp) = False ' Do we export this column
    mvarColDetails(11, pintTemp) = 0
    mvarColDetails(12, pintTemp) = False
    mvarColDetails(13, pintTemp) = "ID"
    mvarColDetails(14, pintTemp) = "ID"
  End If
    
  If mlngOutputFormat = fmtCMGFile Then
    pintTemp = UBound(mvarColDetails, 2) + 1
    ReDim Preserve mvarColDetails(UBound(mvarColDetails, 1), pintTemp)
    mvarColDetails(0, pintTemp) = "C"
    mvarColDetails(1, pintTemp) = mlngExportBaseTable
    mvarColDetails(2, pintTemp) = mstrExportBaseTableName
    mvarColDetails(3, pintTemp) = mlngCMGRecordIdentifier
    mvarColDetails(4, pintTemp) = datGeneral.GetColumnName(mlngCMGRecordIdentifier)
    mvarColDetails(5, pintTemp) = "Identifier"
    mvarColDetails(6, pintTemp) = 0
    mvarColDetails(7, pintTemp) = ""
    mvarColDetails(8, pintTemp) = ""      ' The CMG Code for this column
    mvarColDetails(9, pintTemp) = False   ' Is this column audited
    mvarColDetails(10, pintTemp) = False   ' Do we export this column
    mvarColDetails(11, pintTemp) = 0
    mvarColDetails(12, pintTemp) = False
    mvarColDetails(13, pintTemp) = "Identifier"
    mvarColDetails(14, pintTemp) = "Identifier"

  End If
 
  ' Get those columns defined as a SortOrder and load into array
  
  pstrTempSQL = "SELECT * FROM AsrSysExportDetails WHERE " & _
               "ExportID = " & mlngExportID & " " & _
               "AND SortOrderSequence > 0 AND Type = 'C' " & _
               "ORDER BY [SortOrderSequence]"
               
  Set prstExportSortOrder = datGeneral.GetReadOnlyRecords(pstrTempSQL)
  
  With prstExportSortOrder
    If .BOF And .EOF Then
      GetDetailsRecordsets = False
      mstrErrorString = "No columns have been defined as a sort order for the specified Export definition." & vbCrLf & "Please remove this definition and create a new one."
      Exit Function
    End If
    Do Until .EOF
      pintTemp = UBound(mvarSortOrder, 2) + 1
      ReDim Preserve mvarSortOrder(2, pintTemp)
      mvarSortOrder(1, pintTemp) = !ColExprID
      mvarSortOrder(2, pintTemp) = !SortOrder
      .MoveNext
    Loop
  End With
  
  Set prstExportSortOrder = Nothing
  Set mrstExportDetails = Nothing

  GetDetailsRecordsets = True
  Exit Function
  
GetDetailsRecordsets_ERROR:
  
  GetDetailsRecordsets = False
  mstrErrorString = "Error whilst retrieving the details recordsets." & vbCrLf & "(" & Err.Description & ")"
  
End Function

Private Function GenerateSQL() As Boolean

  ' Purpose : This function calls the individual functions that
  '           general the components of the main SQL string.
  ' Input   : None
  ' Output  : True/False Success
  
  Dim fOK As Boolean
  
  'gobjProgress.UpdateProgress gblnBatchMode
  
  ' If user cancels the export, abort
  If gobjProgress.Cancelled Then
    mblnUserCancelled = True
    GenerateSQL = False
    Exit Function
  End If
  
  fOK = True
  
  If fOK Then fOK = GenerateSQLSelect
  If fOK Then fOK = GenerateSQLFrom
  If fOK Then fOK = GenerateSQLJoin
  If fOK Then fOK = GenerateSQLWhere
  If fOK Then fOK = GenerateSQLOrderBy

  If fOK Then GenerateSQL = True Else GenerateSQL = False
  
End Function

Private Function GenerateSQLSelect() As Boolean
  
  ' Purpose : This function compiles the SQLSelect string looping
  '           thru the column details recordset.
  '           NB. NEEDS TO BE CHANGED FOR EXPRESSIONS !!!
  ' Input   : None
  ' Output  : True/False Success
  
  On Error GoTo GenerateSQLSelect_ERROR
  
  Dim plngTempTableID As Long
  Dim pstrTempTableName As String
  Dim pstrTempColumnName As String
  
  Dim pblnOK As Boolean
  Dim pblnColumnOK As Boolean
  
  Dim pblnNoSelect As Boolean
  Dim pblnFound As Boolean
  
  Dim pintLoop As Integer
  Dim pstrColumnList As String
  Dim pstrColumnCode As String
  Dim pstrSource As String
  Dim pintNextIndex As Integer
  
  Dim alngSourceTables() As Long
  Dim sCalcCode As String
  Dim blnOK As Boolean
  Dim objCalcExpr As New clsExprExpression
  
  Dim iTextCount As Integer
  Dim iFillerCount As Integer
  Dim iRecNumCount As Integer
  Dim iCalculationCount As Integer
  Dim iLoop1 As Integer
  
  Dim sCalcName As String
  Dim objTableView As CTablePrivilege

  ' Set flags with their starting values
  pblnOK = True
  pblnNoSelect = False
  pstrColumnList = ""
  
  ReDim mastrUDFsRequired(0)
  
  ' JPD20030219 Fault 5066
  ' Check the user has permission to read the base table.
  pblnOK = False
  For Each objTableView In gcoTablePrivileges.Collection
    If (objTableView.TableID = mlngExportBaseTable) And _
      (objTableView.AllowSelect) Then
      pblnOK = True
      Exit For
    End If
  Next objTableView
  Set objTableView = Nothing
    
  If Not pblnOK Then
    GenerateSQLSelect = False
    mstrErrorString = "You do not have permission to read the base table" & vbCrLf & "either directly or through any views."
    Exit Function
  End If
  
  ' COWBOY ALERT !!!!!!! (Forgive me)
  ' JDM - 05/09/2005 - Fault 1030X - SQL 2005 ORDER BY clause does not work if used in conjunction with the INTO clause.
  '                                  This looks like its by design and not just a beta fault, but we can re-invetiagte when
  '                                  Micr*s*ft release the full product.
  ' Start off the select statement
  
  ' COWBOY ALERT REVISITED !!!!!!!
  ' JDM - 20/7/2012 - JIRA xxxx - Congratulations to Microsoft - you have now reintroduced this with SQL 2012. Exactly the same issue
  '                               as with SQL 2005, for some reason when using the SELECT... INTO it ignores the sort.
  '                               Run the query analyser with display execution plan to find out for yourself.
  '                               This doesn't affect 2008. I wonder what the next version of SQL will do?
  
  ' COWBOY ALERT RE-REVISTITED !!!!!!
  ' JDM - 17/09/2014 - TFS-9973 - Well, well, well, here we go again.
  If glngSQLVersion = 9 Or glngSQLVersion >= 10.5 Then
    mstrSQLSelect = "SELECT TOP 1000000000000 "
  Else
    mstrSQLSelect = "SELECT "
  End If
  
  ' Dimension an array of tables/views joined to the base table/view
  ' Column 1 = 0 if this row is for a table, 1 if it is for a view
  ' Column 2 = table/view ID
  ' (should contain everything which needs to be joined to the base tbl/view)
  ReDim mlngTableViews(2, 0)
  
  ' Loop thru the columns collection creating the SELECT and JOIN code
  iRecNumCount = 0
  For pintLoop = 1 To UBound(mvarColDetails, 2)
    
    ' Clear temp vars
    plngTempTableID = 0
    pstrTempTableName = vbNullString
    pstrTempColumnName = vbNullString

    'C - Column
    'T - Text
    'F - Filler
    'X - Expression
    'Y - Carriage return
    
    
    ' If its a COLUMN then...
    If mvarColDetails(0, pintLoop) = "C" Then
        
      ' Load the temp variables
      plngTempTableID = mvarColDetails(1, pintLoop)
      pstrTempTableName = mvarColDetails(2, pintLoop)
      pstrTempColumnName = mvarColDetails(4, pintLoop)
      
      ' Check permission on that column
      Set mobjColumnPrivileges = GetColumnPrivileges(pstrTempTableName)
      mstrRealSource = gcoTablePrivileges.Item(pstrTempTableName).RealSource
      
      pblnColumnOK = mobjColumnPrivileges.IsValid(pstrTempColumnName)
        
      If pblnColumnOK Then
        pblnColumnOK = mobjColumnPrivileges.Item(pstrTempColumnName).AllowSelect
      End If
      
      If pblnColumnOK Then
        
        ' this column can be read direct from the tbl/view or from a parent table
        
'        pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ",", "") & _
'        mstrRealSource & "." & Trim(pstrTempColumnName) & _
'        " AS '" & mvarColDetails(14, pintLoop) & "'"
'        '" AS '" & mvarColDetails(5, pintLoop) & "'"


        pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ",", "") & _
        mstrRealSource & "." & Trim(pstrTempColumnName) & _
        " AS '" & mvarColDetails(14, pintLoop) & "'"
        
        
        ' If the table isnt the base table (or its realsource) then
        ' Check if it has already been added to the array. If not, add it.
        If plngTempTableID <> mlngExportBaseTable Then
          pblnFound = False
          For pintNextIndex = 1 To UBound(mlngTableViews, 2)
            If mlngTableViews(1, pintNextIndex) = 0 And _
            mlngTableViews(2, pintNextIndex) = plngTempTableID Then
              pblnFound = True
              Exit For
            End If
          Next pintNextIndex
          
          If Not pblnFound Then
            pintNextIndex = UBound(mlngTableViews, 2) + 1
            ReDim Preserve mlngTableViews(2, pintNextIndex)
            mlngTableViews(1, pintNextIndex) = 0
            mlngTableViews(2, pintNextIndex) = plngTempTableID
          End If
        End If
                
        ' Optional filter for only auditted changes
        If mdLastSuccessfulOutput <> "00:00:00" Then
          mstrOnlyChangesFilter = mstrOnlyChangesFilter + IIf(Len(mstrOnlyChangesFilter) > 0, " OR ", "") _
            & " dbo.udfsys_FieldChangedSinceLastExport(" & mvarColDetails(3, pintLoop) & ", '" & Format(mdLastSuccessfulOutput, "MM/dd/yyyy hh:mm:ss") & "', " & mstrRealSource & ".[id]) = 1"
        End If
        
      Else
      
        ' this column cannot be read direct. If its from a parent, try parent views
        ' Loop thru the views on the table, seeing if any have read permis for the column
        
        ReDim mstrViews(0)
        For Each mobjTableView In gcoTablePrivileges.Collection
          If (Not mobjTableView.IsTable) And _
          (mobjTableView.TableID = plngTempTableID) And _
          (mobjTableView.AllowSelect) Then
          
            pstrSource = mobjTableView.ViewName
            mstrRealSource = gcoTablePrivileges.Item(pstrSource).RealSource
            
            ' Get the column permission for the view
            Set mobjColumnPrivileges = GetColumnPrivileges(pstrSource)
            
            ' If we can see the column from this view
            If mobjColumnPrivileges.IsValid(pstrTempColumnName) Then
              If mobjColumnPrivileges.Item(pstrTempColumnName).AllowSelect Then
                
                ReDim Preserve mstrViews(UBound(mstrViews) + 1)
                mstrViews(UBound(mstrViews)) = mobjTableView.ViewName
                
                ' Check if view has already been added to the array
                pblnFound = False
                For pintNextIndex = 1 To UBound(mlngTableViews, 2)
                  If mlngTableViews(1, pintNextIndex) = 1 And _
                  mlngTableViews(2, pintNextIndex) = mobjTableView.ViewID Then
                    pblnFound = True
                    Exit For
                  End If
                Next pintNextIndex
                
                If Not pblnFound Then
                
                  ' View hasnt yet been added, so add it !
                  pintNextIndex = UBound(mlngTableViews, 2) + 1
                  ReDim Preserve mlngTableViews(2, pintNextIndex)
                  mlngTableViews(1, pintNextIndex) = 1
                  mlngTableViews(2, pintNextIndex) = mobjTableView.ViewID
                  
                End If
              End If
            End If
          End If
        
        Next mobjTableView
      
        Set mobjTableView = Nothing
        
        ' Does the user have select permission thru ANY views ?
        If UBound(mstrViews()) = 0 Then
          pblnNoSelect = True
        Else
        
          ' Add the column to the column list
          pstrColumnCode = ""
          For pintNextIndex = 1 To UBound(mstrViews)
            If pintNextIndex = 1 Then
              pstrColumnCode = "CASE"
            End If
            
            pstrColumnCode = pstrColumnCode & _
            " WHEN NOT " & mstrViews(pintNextIndex) & "." & pstrTempColumnName & " IS NULL THEN " & mstrViews(pintNextIndex) & "." & pstrTempColumnName
            
          Next pintNextIndex
          
          If Len(pstrColumnCode) > 0 Then
            pstrColumnCode = pstrColumnCode & _
            " ELSE NULL" & _
            " END AS '" & mvarColDetails(14, pintLoop) & "'"
            '" END AS '" & mvarColDetails(5, pintLoop) & "'"

        
            pstrColumnList = pstrColumnList & _
            IIf(Len(pstrColumnList) > 0, ",", "") & _
            pstrColumnCode
          End If
        
        End If
      
        ' If we cant see a column, then get outta here
        If pblnNoSelect Then
          GenerateSQLSelect = False
          mstrErrorString = "You do not have permission to see the column '" & mvarColDetails(4, pintLoop) & "'" & vbCrLf & "either directly or through any views." & vbCrLf & vbCrLf & "You may not run this Export."
          Exit Function
        End If
        

        If Not pblnOK Then
          GenerateSQLSelect = False
          Exit Function
        End If
        
      End If
      
    ElseIf mvarColDetails(0, pintLoop) = "T" Then

      ' Text, increase the text counter...
      iTextCount = iTextCount + 1
      'pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ", ", "") & _
        "'" & Replace(mvarColDetails(5, pintLoop), "'", "''") & "' AS Text" & iTextCount
      pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ", ", "") & _
        "'" & Replace(mvarColDetails(5, pintLoop), "'", "''") & "' AS '" & mvarColDetails(14, pintLoop) & "'"
              
    ElseIf mvarColDetails(0, pintLoop) = "R" Then
              
      'Need to allow two characters for carriage return and line feed
      mvarColDetails(6, pintLoop) = 2

      ' Do the same as filler, increase the filer counter.
      iFillerCount = iFillerCount + 1
      'pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ", ", "") & _
        "'" & vbCrLf & "' AS Filler" & iFillerCount
      pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ", ", "") & _
        "'" & vbCrLf & "' AS '" & mvarColDetails(14, pintLoop) & "'"

        '"0 AS Filler" & iFillerCount

    ElseIf mvarColDetails(0, pintLoop) = "N" Then
              
      ' Do the same as filler, increase the filer counter.
      iRecNumCount = iRecNumCount + 1
      'pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ", ", "") & _
        "'" & vbCrLf & "' AS 'Record Number" & IIf(iRecNumCount > 1, CStr(iRecNumCount), "") & "'"
      pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ", ", "") & _
        "'" & vbCrLf & "' AS '" & mvarColDetails(14, pintLoop) & "'"

    ElseIf mvarColDetails(0, pintLoop) = "F" Then
    
      ' Filler, increase the filer counter.
      iFillerCount = iFillerCount + 1

      ' If export is fixed length, use filler length as defined else use 1.
      'If mstrExportOutputType <> "F" Then
      If mlngOutputFormat <> fmtFixedLengthFile Then
        'pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ", ", "") & _
          "' ' AS Filler" & iFillerCount
        pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ", ", "") & _
          "' ' AS '" & mvarColDetails(14, pintLoop) & "'"
      Else
        'pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ", ", "") & _
          "'" & Space(mvarColDetails(6, pintLoop)) & "' AS Filler" & iFillerCount
        pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ", ", "") & _
          "'" & Space(mvarColDetails(6, pintLoop)) & "' AS '" & mvarColDetails(14, pintLoop) & "'"
      End If
      
    ElseIf mvarColDetails(0, pintLoop) = "X" Then
    
      ' Its an expression !!!
      
      iCalculationCount = iCalculationCount + 1
      
'      pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ", ", "") & _
        "'" & Replace(mvarColDetails(5, pintLoop), "'", "''") & "' AS Calculation" & iCalculationCount
      
      
      '################ NICKED FROM CUSTOM REPORT CODE....
'
'      ' Get the calculation SQL, and the array of tables/views that are used to create it.
'      ' Column 1 = 0 if this row is for a table, 1 if it is for a view.
'      ' Column 2 = table/view ID.
      ReDim alngSourceTables(2, 0)
      Set objCalcExpr = New clsExprExpression
      blnOK = objCalcExpr.Initialise(mlngExportBaseTable, CLng(mvarColDetails(3, pintLoop)), giEXPR_RUNTIMECALCULATION, giEXPRVALUE_UNDEFINED)
      sCalcName = objCalcExpr.Name
      If blnOK Then
        blnOK = objCalcExpr.RuntimeCalculationCode(alngSourceTables, sCalcCode, True)
      
        If blnOK And gbEnableUDFFunctions Then
          blnOK = objCalcExpr.UDFCalculationCode(alngSourceTables, mastrUDFsRequired(), True)
        End If
        
      End If
      
      'TM20030422 Fault 5242 - The "SELECT ... INTO..." statement errors when it trys to create a column for
      'and empty string. Therefore wrap this empty sting in a CONVERT(varchar... clause.
      'TM20030521 Fault 5702 - Compare the empty string with the calc code value converted to varchar
      sCalcCode = "CASE WHEN CONVERT(varchar," & sCalcCode & ") = '' " & _
                  "THEN CONVERT(varchar," & sCalcCode & ") " & _
                  "ELSE " & sCalcCode & " END"

      Set objCalcExpr = Nothing

      If blnOK Then
      
        'pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ",", "") & _
          sCalcCode & " AS 'Calculation" & iCalculationCount & "' "
        pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ",", "") & _
          sCalcCode & " AS '" & mvarColDetails(14, pintLoop) & "' "

        ' Add the required views to the JOIN code.
         For iLoop1 = 1 To UBound(alngSourceTables, 2)
           If alngSourceTables(1, iLoop1) = 1 Then
             ' Check if view has already been added to the array
             pblnFound = False
             For pintNextIndex = 1 To UBound(mlngTableViews, 2)
               If mlngTableViews(1, pintNextIndex) = 1 And _
                 mlngTableViews(2, pintNextIndex) = alngSourceTables(2, iLoop1) Then
                 pblnFound = True
                 Exit For
               End If
             Next pintNextIndex

             If Not pblnFound Then

               ' View hasnt yet been added, so add it !
               pintNextIndex = UBound(mlngTableViews, 2) + 1
               ReDim Preserve mlngTableViews(2, pintNextIndex)
               mlngTableViews(1, pintNextIndex) = 1
               mlngTableViews(2, pintNextIndex) = alngSourceTables(2, iLoop1)

             End If
             
'TM20020121 Fault 2277
'********************************************************************************
          ElseIf alngSourceTables(1, iLoop1) = 0 Then
             ' Check if table has already been added to the array
             pblnFound = False
             For pintNextIndex = 1 To UBound(mlngTableViews, 2)
               If mlngTableViews(1, pintNextIndex) = 0 And _
                 mlngTableViews(2, pintNextIndex) = alngSourceTables(2, iLoop1) Then
                 pblnFound = True
                 Exit For
               End If
             Next pintNextIndex
                     
             If Not pblnFound Then
               ' table hasnt yet been added, so add it !
               pintNextIndex = UBound(mlngTableViews, 2) + 1
               ReDim Preserve mlngTableViews(2, pintNextIndex)
               mlngTableViews(1, pintNextIndex) = 0
               mlngTableViews(2, pintNextIndex) = alngSourceTables(2, iLoop1)
             End If
'********************************************************************************
          End If
        Next iLoop1
      Else
        ' Permission denied on something in the calculation.
        mstrErrorString = "You do not have permission to use the '" & sCalcName & "' calculation."
        GenerateSQLSelect = False
        Exit Function
      End If

      '################
    End If
      
  Next pintLoop

  mstrSQLSelect = mstrSQLSelect & pstrColumnList
  
  If mlngOutputFormat = fmtXML Then
    mstrSQLSelect = mstrSQLSelect & ", NEWID() AS [_XMLSplitID]"
  End If
  
  GenerateSQLSelect = True
  
  Exit Function
  
GenerateSQLSelect_ERROR:
  
  GenerateSQLSelect = False
  mstrErrorString = "Error whilst generating SQL Select statement." & vbCrLf & "(" & Err.Description & ")"
  
End Function

Private Function GenerateSQLFrom() As Boolean

  ' Purpose : It doesnt take Einstein to work out that this function
  '           adds the base table name to the from clause of the SQL string.
  ' Input   : None
  ' Output  : True/False Success
 
'  Dim iLoop As Integer
  Dim pobjTableView As CTablePrivilege
  
  Set pobjTableView = New CTablePrivilege
  
    mstrSQLFrom = gcoTablePrivileges.Item(mstrExportBaseTableName).RealSource
  
'  Else
'
'    ' need some way of determining which view out of all these
'    ' is based on the base table, use it in the from statement
'    ' and mark it so that i know what tables to join to it
'
'    ' NB 13/03/00 IS THIS NEEDED NOW AS ALL USERS WILL HAVE SELECT PERMISSION
'    '             ON THE ID COLUMN OF ALL TABLES ???
'
'    For Each pobjTableView In gcoTablePrivileges.Collection
'
'      If (Not pobjTableView.IsTable) And _
'      (pobjTableView.TableID = mlngExportBaseTable) And _
'      (pobjTableView.AllowSelect) Then
'
'        mstrSQLFrom = pobjTableView.ViewName
'        Exit For
'
'    End If
'
'    Next pobjTableView
'
'  End If
 
  Set pobjTableView = Nothing
  
  GenerateSQLFrom = True
  Exit Function
  
GenerateSQLFrom_ERROR:
  
  GenerateSQLFrom = False
  mstrErrorString = "Error in GenerateSQLFrom." & vbCrLf & "(" & Err.Description & ""
  
End Function

Private Function GenerateSQLJoin() As Boolean

  ' Purpose : Add the join strings for parent/child/views.
  '           Also adds filter clauses to the joins if used
  ' Input   : None
  ' Output  : True/False Success
  
  On Error GoTo GenerateSQLJoin_ERROR

  Dim blnOK As Boolean
  Dim pobjTableView As CTablePrivilege
  Dim objChildTable As CTablePrivilege
  Dim pintLoop As Integer
  Dim sChildJoinCode As String
  Dim sReuseJoinCode As String
  Dim sChildOrderString As String
  Dim rsTemp As Recordset
  Dim strFilterIDs As String
  Dim pblnChildUsed As Boolean
  Dim sChildJoin As String
  Dim objExpr As clsExprExpression
  
  ' Get the base table real source
  mstrBaseTableRealSource = mstrSQLFrom
  
  ' First, do the joins for all the views etc...
  For pintLoop = 1 To UBound(mlngTableViews, 2)
  
    ' Get the table/view object from the id stored in the array
    If mlngTableViews(1, pintLoop) = 0 Then
      Set pobjTableView = gcoTablePrivileges.FindTableID(mlngTableViews(2, pintLoop))
    Else
      Set pobjTableView = gcoTablePrivileges.FindViewID(mlngTableViews(2, pintLoop))
    End If
    
    ' Dont add a join here if its the child or parent table...do that later
    If pobjTableView.TableID <> mlngExportChildTable Then
      If pobjTableView.TableID <> mlngExportParent1Table Then
        If pobjTableView.TableID <> mlngExportParent2Table Then
           
          If (pobjTableView.ViewName <> mstrBaseTableRealSource) Then
            mstrSQLJoin = mstrSQLJoin & _
            " LEFT OUTER JOIN " & pobjTableView.RealSource & _
            " ON " & mstrBaseTableRealSource & ".ID = " & pobjTableView.RealSource & ".ID"
          End If
          
        End If
      End If
    End If
  
    ' OK, parent table joins...
    If (pobjTableView.TableID = mlngExportParent1Table) Or _
       (pobjTableView.TableID = mlngExportParent2Table) Then
      
      mstrSQLJoin = mstrSQLJoin & _
           " LEFT OUTER JOIN " & pobjTableView.RealSource & _
           " ON " & mstrBaseTableRealSource & ".ID_" & pobjTableView.TableID & " = " & pobjTableView.RealSource & ".ID"
    End If
  
  Next pintLoop
   
  ' Second, do the childview bit, if required

  If mlngExportChildTable > 0 Then
  
    ' are any child fields in the export ? # 12/06/00 RH - FAULT 419
    For pintLoop = 1 To UBound(mvarColDetails, 2)
      If mvarColDetails(1, pintLoop) = mlngExportChildTable Then pblnChildUsed = True
    Next pintLoop
      
    If pblnChildUsed = True Then

      Set objChildTable = gcoTablePrivileges.FindTableID(mlngExportChildTable)
      
      If objChildTable.AllowSelect Then
        sChildJoinCode = " LEFT OUTER JOIN " & objChildTable.RealSource & _
                         " ON " & mstrBaseTableRealSource & ".ID = " & _
                         objChildTable.RealSource & ".ID_" & mlngExportBaseTable
          
        sChildJoinCode = sChildJoinCode & " AND " & objChildTable.RealSource & ".ID IN"
        
        sChildJoinCode = sChildJoinCode & _
        " (SELECT TOP" & IIf(mlngExportChildMaxRecords = 0, " 100 PERCENT", " " & mlngExportChildMaxRecords) & _
        " " & objChildTable.RealSource & ".ID FROM " & objChildTable.RealSource
        
        ' Now the child order by bit - done here in case tables need to be joined.
        Set rsTemp = datGeneral.GetOrderDefinition(datGeneral.GetDefaultOrder(mlngExportChildTable))
        sChildOrderString = DoChildOrderString(rsTemp, sChildJoin)
        Set rsTemp = Nothing
      
        sChildJoinCode = sChildJoinCode & sChildJoin
        
        sChildJoinCode = sChildJoinCode & _
          " WHERE (" & objChildTable.RealSource & ".ID_" & mlngExportBaseTable & _
          " = " & mstrBaseTableRealSource & ".ID)"
        
        ' is the child filtered ?
        
        If mlngExportChildFilterID > 0 Then
          blnOK = datGeneral.FilteredIDs(mlngExportChildFilterID, strFilterIDs)
    
          ' Generate any UDFs that are used in this filter
          If blnOK Then
            datGeneral.FilterUDFs mlngExportChildFilterID, mastrUDFsRequired()
          End If
    
          If blnOK Then
            sChildJoinCode = sChildJoinCode & " AND " & _
              objChildTable.RealSource & ".ID IN (" & strFilterIDs & ")"
          Else
            ' Permission denied on something in the filter.
            mstrErrorString = "You do not have permission to use the '" & datGeneral.GetFilterName(mlngExportChildFilterID) & "' filter."
            GenerateSQLJoin = False
            Exit Function
          End If
        End If
                   
      End If
    
    End If

  End If
  
  mstrSQLJoin = mstrSQLJoin & sChildJoinCode & IIf(Len(sChildOrderString) > 0, " ORDER BY " & sChildOrderString & ")", "")
  
  GenerateSQLJoin = True
  Exit Function
  
GenerateSQLJoin_ERROR:
  
  GenerateSQLJoin = False
  mstrErrorString = "Error in GenerateSQLJoin." & vbCrLf & Err.Description
  
End Function

Private Function DoChildOrderString(rsTemp As Recordset, ByRef psJoinCode As String) As String

  ' This function loops through the child tables default order
  ' checking if the user has privileges. If they do, add to the order string
  ' if not, leave it out.
  
  On Error GoTo DoChildOrderString_ERROR
  
  Dim fColumnOK As Boolean
  Dim fFound As Boolean
  Dim iNextIndex As Integer
  Dim sSource As String
  Dim sRealSource As String
  Dim sColumnCode As String
  Dim sCurrentTableViewName As String
  Dim objColumnPrivileges As CColumnPrivileges
  Dim pobjOrderCol As CTablePrivilege
  Dim objTableView As CTablePrivilege
  Dim alngTableViews() As Long
  Dim asViews() As String
  Dim iTempCounter As Integer
  
  ' Dimension an array of tables/views joined to the base table/view.
  ' Column 1 = 0 if this row is for a table, 1 if it is for a view.
  ' Column 2 = table/view ID.
  ReDim alngTableViews(2, 0)

  Set pobjOrderCol = gcoTablePrivileges.FindTableID(mlngExportChildTable)
  sCurrentTableViewName = pobjOrderCol.RealSource
  Set pobjOrderCol = Nothing
    
  Do Until rsTemp.EOF
    If rsTemp!Type = "O" Then
      ' Check if the user can read the column.
      Set pobjOrderCol = gcoTablePrivileges.FindTableID(rsTemp!TableID)
      Set objColumnPrivileges = GetColumnPrivileges(pobjOrderCol.TableName)
      fColumnOK = objColumnPrivileges.Item(rsTemp!ColumnName).AllowSelect
      Set objColumnPrivileges = Nothing
      
      If fColumnOK Then
        If rsTemp!TableID = mlngExportChildTable Then
          DoChildOrderString = DoChildOrderString & _
            IIf(Len(DoChildOrderString) > 0, ",", "") & _
            pobjOrderCol.RealSource & "." & rsTemp!ColumnName & _
            IIf(rsTemp!Ascending, "", " DESC")
        Else
          ' If the column comes from a parent table, then add the table to the Join code.
          ' Check if the table has already been added to the join code.
          fFound = False
          iTempCounter = 0
          For iNextIndex = 1 To UBound(alngTableViews, 2)
            If alngTableViews(1, iNextIndex) = 0 And _
              alngTableViews(2, iNextIndex) = rsTemp!TableID Then
              iTempCounter = iNextIndex
              fFound = True
              Exit For
            End If
          Next iNextIndex

          If Not fFound Then
            ' The table has not yet been added to the join code, so add it to the array and the join code.
            iNextIndex = UBound(alngTableViews, 2) + 1
            ReDim Preserve alngTableViews(2, iNextIndex)
            alngTableViews(1, iNextIndex) = 0
            alngTableViews(2, iNextIndex) = rsTemp!TableID

            iTempCounter = iNextIndex
            
            psJoinCode = psJoinCode & _
              " LEFT OUTER JOIN " & pobjOrderCol.RealSource & " ASRSysTemp_" & Trim(Str(iTempCounter)) & _
              " ON " & sCurrentTableViewName & ".ID_" & Trim(Str(rsTemp!TableID)) & _
              " = ASRSysTemp_" & Trim(Str(iTempCounter)) & ".ID"
          End If
        
          DoChildOrderString = DoChildOrderString & _
            IIf(Len(DoChildOrderString) > 0, ",", "") & _
            "ASRSysTemp_" & Trim(Str(iTempCounter)) & "." & rsTemp!ColumnName & _
            IIf(rsTemp!Ascending, "", " DESC")
        End If
      Else
        ' The column cannot be read from the base table/view, or directly from a parent table.
        ' If it is a column from a prent table, then try to read it from the views on the parent table.
        If rsTemp!TableID <> mlngExportChildTable Then
          ' Loop through the views on the column's table, seeing if any have 'read' permission granted on them.
          ReDim asViews(0)
          For Each objTableView In gcoTablePrivileges.Collection
            If (Not objTableView.IsTable) And _
              (objTableView.TableID = rsTemp!TableID) And _
              (objTableView.AllowSelect) Then

              sSource = objTableView.ViewName
              sRealSource = gcoTablePrivileges.Item(sSource).RealSource

              ' Get the column permission for the view.
              Set objColumnPrivileges = GetColumnPrivileges(sSource)

              If objColumnPrivileges.IsValid(rsTemp!ColumnName) Then
                If objColumnPrivileges.Item(rsTemp!ColumnName).AllowSelect Then
                  ' Add the view info to an array to be put into the column list or order code below.
                  iNextIndex = UBound(asViews) + 1
                  ReDim Preserve asViews(iNextIndex)
                  asViews(iNextIndex) = objTableView.ViewName

                  ' Add the view to the Join code.
                  ' Check if the view has already been added to the join code.
                  fFound = False
                  iTempCounter = 0
                  For iNextIndex = 1 To UBound(alngTableViews, 2)
                    If alngTableViews(1, iNextIndex) = 1 And _
                      alngTableViews(2, iNextIndex) = objTableView.ViewID Then
                      fFound = True
                      iTempCounter = iNextIndex
                      Exit For
                    End If
                  Next iNextIndex

                  If Not fFound Then
                    ' The view has not yet been added to the join code, so add it to the array and the join code.
                    iNextIndex = UBound(alngTableViews, 2) + 1
                    ReDim Preserve alngTableViews(2, iNextIndex)
                    alngTableViews(1, iNextIndex) = 1
                    alngTableViews(2, iNextIndex) = objTableView.ViewID

                    iTempCounter = iNextIndex

                    psJoinCode = psJoinCode & _
                      " LEFT OUTER JOIN " & sRealSource & " ASRSysTemp_" & Trim(Str(iTempCounter)) & _
                      " ON " & sCurrentTableViewName & ".ID_" & Trim(Str(objTableView.TableID)) & _
                      " = ASRSysTemp_" & Trim(Str(iTempCounter)) & ".ID"
                  End If
                End If
              End If
              Set objColumnPrivileges = Nothing
            End If
          Next objTableView
          Set objTableView = Nothing

          ' The current user does have permission to 'read' the column through a/some view(s) on the
          ' table.
          If UBound(asViews) > 0 Then
            ' Add the column to the column list.
            sColumnCode = ""
            For iNextIndex = 1 To UBound(asViews)
              If iNextIndex = 1 Then
                sColumnCode = "CASE "
              End If

              sColumnCode = sColumnCode & _
                " WHEN NOT ASRSysTemp_" & Trim(Str(iNextIndex)) & "." & rsTemp!ColumnName & " IS NULL THEN ASRSysTemp_" & Trim(Str(iNextIndex)) & "." & rsTemp!ColumnName
            Next iNextIndex

            If Len(sColumnCode) > 0 Then
              sColumnCode = sColumnCode & _
                " ELSE NULL" & _
                " END"

              ' Add the column to the order string.
              DoChildOrderString = DoChildOrderString & _
                IIf(Len(DoChildOrderString) > 0, ", ", "") & _
                sColumnCode & _
                IIf(rsTemp!Ascending, "", " DESC")
            End If
          End If
        End If
      End If
  
      Set pobjOrderCol = Nothing
    End If
    
    rsTemp.MoveNext
  Loop

  Exit Function
  
DoChildOrderString_ERROR:
  
  Set pobjOrderCol = Nothing
  mstrErrorString = "Error while generating child order string." & vbCrLf & "(" & Err.Description & ")"
  DoChildOrderString = ""

End Function

Private Function GenerateSQLWhere() As Boolean

  ' Purpose : Generate the where clauses that cope with the joins
  '           NB Need to add the where clauses for filters/picklists etc
  ' Input   : None
  ' Output  : True/False Success

  On Error GoTo GenerateSQLWhere_ERROR
  
  Dim blnOK As Boolean
  Dim pintLoop As Integer
  Dim pobjTableView As CTablePrivilege
  Dim prstTemp As New Recordset
  Dim pstrPickListIDs As String
  Dim strFilterIDs As String
  Dim objExpr As clsExprExpression
  Dim pstrParent1PickListIDs As String
  Dim pstrParent2PickListIDs As String
  
  '# remove this if test dont work RH 18/05
  Set pobjTableView = gcoTablePrivileges.FindTableID(mlngExportBaseTable)
  If pobjTableView.AllowSelect = False Then
  '#
    
  ' First put the where clauses in for the joins...only if base table is a top level table
    If UCase(Left(mstrBaseTableRealSource, 6)) <> "ASRSYS" Then
      
      For pintLoop = 1 To UBound(mlngTableViews, 2)
        ' Get the table/view object from the id stored in the array
        If mlngTableViews(1, pintLoop) = 0 Then
          Set pobjTableView = gcoTablePrivileges.FindTableID(mlngTableViews(2, pintLoop))
        Else
          Set pobjTableView = gcoTablePrivileges.FindViewID(mlngTableViews(2, pintLoop))
        End If
  
        ' dont add where clause for the base/chil/p1/p2 TABLES...only add views here
        ' JPD20030207 Fault 5033
        'If (mlngTableViews(1, pintLoop) = 1) And _
          (mlngTableViews(2, pintLoop) <> mlngExportChildTable) _
          And (mlngTableViews(2, pintLoop) <> mlngExportParent1Table) _
          And (mlngTableViews(2, pintLoop) <> mlngExportParent2Table) _
          And (mlngTableViews(2, pintLoop) <> mlngExportBaseTable) Then
        If (mlngTableViews(1, pintLoop) = 1) Then
            mstrSQLWhere = mstrSQLWhere & _
            IIf(Len(mstrSQLWhere) > 0, " OR ", " WHERE (") & _
            mstrBaseTableRealSource & ".ID IN (SELECT ID FROM " & pobjTableView.RealSource & ")"
        End If
        
      Next pintLoop
    
      If Len(mstrSQLWhere) > 0 Then mstrSQLWhere = mstrSQLWhere & ")"
      
    End If
        
    '# remove this if test dont work RH 18/05
    End If
    '#
    
  ' Parent 1 filter and picklist
  If mlngExportParent1PickListID > 0 Then
    pstrParent1PickListIDs = ""
    Set prstTemp = mclsData.OpenRecordset("EXEC sp_ASRGetPickListRecords " & mlngExportParent1PickListID, adOpenForwardOnly, adLockReadOnly)
    
    If prstTemp.BOF And prstTemp.EOF Then
      mstrErrorString = "The first parent table picklist contains no records."
      GenerateSQLWhere = False
      Exit Function
    End If
      
    Do While Not prstTemp.EOF
      pstrParent1PickListIDs = pstrParent1PickListIDs & IIf(Len(pstrParent1PickListIDs) > 0, ", ", "") & prstTemp.Fields(0)
      prstTemp.MoveNext
    Loop
    
    prstTemp.Close
    Set prstTemp = Nothing

    mstrSQLWhere = mstrSQLWhere & _
      IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & _
      mstrBaseTableRealSource & ".ID_" & mlngExportParent1Table & " IN (" & pstrParent1PickListIDs & ") "
  ElseIf mlngExportParent1FilterID > 0 Then
      blnOK = True
      blnOK = datGeneral.FilteredIDs(mlngExportParent1FilterID, strFilterIDs)
           
      ' Generate any UDFs that are used in this filter
      If blnOK Then
        datGeneral.FilterUDFs mlngExportParent1FilterID, mastrUDFsRequired()
      End If
      
      If blnOK Then
        mstrSQLWhere = mstrSQLWhere & _
        IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & _
        mstrBaseTableRealSource & ".ID_" & mlngExportParent1Table & " IN (" & strFilterIDs & ") "
      Else
        mstrErrorString = "You do not have permission to use the '" & datGeneral.GetFilterName(mlngExportParent1FilterID) & "' filter."
        GenerateSQLWhere = False
        Exit Function
      End If
    End If

  ' Parent 2 filter and picklist
  If mlngExportParent2PickListID > 0 Then
    pstrParent2PickListIDs = ""
    Set prstTemp = mclsData.OpenRecordset("EXEC sp_ASRGetPickListRecords " & mlngExportParent2PickListID, adOpenForwardOnly, adLockReadOnly)
    
    If prstTemp.BOF And prstTemp.EOF Then
      mstrErrorString = "The second parent table picklist contains no records."
      GenerateSQLWhere = False
      Exit Function
    End If
      
    Do While Not prstTemp.EOF
      pstrParent2PickListIDs = pstrParent2PickListIDs & IIf(Len(pstrParent2PickListIDs) > 0, ", ", "") & prstTemp.Fields(0)
      prstTemp.MoveNext
    Loop
    
    prstTemp.Close
    Set prstTemp = Nothing

    mstrSQLWhere = mstrSQLWhere & _
      IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & _
      mstrBaseTableRealSource & ".ID_" & mlngExportParent2Table & " IN (" & pstrParent2PickListIDs & ") "
  ElseIf mlngExportParent2FilterID > 0 Then
      blnOK = True
      blnOK = datGeneral.FilteredIDs(mlngExportParent2FilterID, strFilterIDs)
      
      ' Generate any UDFs that are used in this filter
      If blnOK Then
        datGeneral.FilterUDFs mlngExportParent2FilterID, mastrUDFsRequired()
      End If
      
      If blnOK Then
        mstrSQLWhere = mstrSQLWhere & _
        IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & _
        mstrBaseTableRealSource & ".ID_" & mlngExportParent2Table & " IN (" & strFilterIDs & ") "
      Else
        mstrErrorString = "You do not have permission to use the '" & datGeneral.GetFilterName(mlngExportParent2FilterID) & "' filter."
        GenerateSQLWhere = False
        Exit Function
      End If
    End If
  
  ' Now if we are using a picklist, add a where clause for that
  'Get List of IDs from Picklist
  If mlngExportPickListID > 0 Then
    Set prstTemp = mclsData.OpenRecordset("EXEC sp_ASRGetPickListRecords " & mlngExportPickListID, adOpenForwardOnly, adLockReadOnly)
    
    If prstTemp.BOF And prstTemp.EOF Then
      mstrErrorString = "The selected picklist contains no records."
      GenerateSQLWhere = False
      Exit Function
    End If
      
    Do While Not prstTemp.EOF
      pstrPickListIDs = pstrPickListIDs & IIf(Len(pstrPickListIDs) > 0, ", ", "") & prstTemp.Fields(0)
      prstTemp.MoveNext
    Loop
    
    prstTemp.Close
    Set prstTemp = Nothing

    mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & _
    mstrSQLFrom & ".ID IN (" & pstrPickListIDs & ")"
  
  ElseIf mlngExportFilterID > 0 Then
    
    blnOK = datGeneral.FilteredIDs(mlngExportFilterID, strFilterIDs)
    
    ' Generate any UDFs that are used in this filter
    If blnOK Then
      datGeneral.FilterUDFs mlngExportFilterID, mastrUDFsRequired()
    End If
    
    If blnOK Then
      mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & _
        mstrSQLFrom & ".ID IN (" & strFilterIDs & ")"
    Else
      ' Permission denied on something in the filter.
      mstrErrorString = "You do not have permission to use the '" & datGeneral.GetFilterName(mlngExportFilterID) & "' filter."
      GenerateSQLWhere = False
      Exit Function
    End If
  End If
      
  Set prstTemp = Nothing

  ' Audit changes only
  If Len(mstrOnlyChangesFilter) > 0 And mbAuditChangesOnly And mlngOutputFormat = fmtXML Then
    mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & "(" & mstrOnlyChangesFilter & ")"
  End If

  GenerateSQLWhere = True
  Exit Function
  
GenerateSQLWhere_ERROR:
  
  GenerateSQLWhere = False
  mstrErrorString = "Error in GenerateSQLWhere." & vbCrLf & "(" & Err.Description & ")"
  
End Function

Private Function GenerateSQLOrderBy() As Boolean

  ' Purpose : Returns order by string from the sort order array
  ' Input   : None
  ' Output  : True/False Success

  On Error GoTo GenerateSQLOrderBy_ERROR
  
  If UBound(mvarSortOrder, 2) > 0 Then
    ' Columns have been defined, so use these for the base table/view
    mstrSQLOrderBy = DoDefinedOrderBy
  End If
  
  If Len(mstrSQLOrderBy) > 0 Then mstrSQLOrderBy = " ORDER BY " & mstrSQLOrderBy
    
   
  GenerateSQLOrderBy = True
  Exit Function
  
GenerateSQLOrderBy_ERROR:
  
  GenerateSQLOrderBy = False
  mstrErrorString = "Error in GenerateSQLOrderBy." & vbCrLf & "(" & Err.Description & ")"
  
End Function

Private Function DoDefinedOrderBy() As String

  ' This function creates the base ORDER BY statement by searching
  ' through the columns defined as the reports sort order, then
  ' uses the relevant alias name
  
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  
  For iLoop = 1 To UBound(mvarSortOrder, 2)
      
    For iLoop2 = 1 To UBound(mvarColDetails, 2)
      
      If mvarSortOrder(1, iLoop) = mvarColDetails(3, iLoop2) Then
        
        DoDefinedOrderBy = DoDefinedOrderBy & _
        IIf(Len(DoDefinedOrderBy) > 0, ",", "") & _
        "[" & mvarColDetails(14, iLoop2) & "] " & _
        mvarSortOrder(2, iLoop)

        '"[" & mvarColDetails(5, iLoop2) & "] " & _

        Exit For
        
      End If
    
    Next iLoop2
    
  Next iLoop
  
End Function

Private Function CheckRecordSet() As Boolean

  ' Purpose : To get recordset from temptable and show recordcount
  ' Input   : None
  ' Output  : True/False Success
  
  On Error GoTo CheckRecordSet_ERROR
  
  'gobjProgress.UpdateProgress gblnBatchMode

  ' If user cancels the export, abort
  If gobjProgress.Cancelled Then
    mblnUserCancelled = True
    CheckRecordSet = False
    Exit Function
  End If

  Set mrstExportOutput = mclsData.OpenRecordset("SELECT * FROM " & mstrTempTableName, adOpenStatic, adLockReadOnly)

  If mrstExportOutput.BOF And mrstExportOutput.EOF Then
    
    ' If we want to force the header trick export into thinking we have records
    If mbForceHeader Then
      CheckRecordSet = True
    Else
      CheckRecordSet = False
      mstrErrorString = "Export generated no records !"
    End If
    
    mblnNoRecords = True
    Exit Function
  End If
  
  mrstExportOutput.MoveLast
  mrstExportOutput.MoveFirst
  
  mlngExportRecordCount = mrstExportOutput.RecordCount
  
  If gblnBatchMode = False Then
    gobjProgress.Bar1MaxValue = mlngExportRecordCount
  Else
    gobjProgress.Bar2MaxValue = mlngExportRecordCount
  End If
  
  CheckRecordSet = True
  Exit Function
  
CheckRecordSet_ERROR:
  
  mstrErrorString = "Error while checking returned recordset." & vbCrLf & "(" & Err.Description & ")"
  CheckRecordSet = False
  
End Function

Private Function ClearUp() As Boolean

  ' Purpose : To clear all variables/recordsets/references and drops temptable
  ' Input   : None
  ' Output  : True/False success

  ' Definition variables
  
  On Error GoTo ClearUp_ERROR
  
  mlngExportID = 0
  mstrExportName = vbNullString
  mstrExportDescription = vbNullString
  mlngExportBaseTable = 0
  mstrExportBaseTableName = vbNullString
  mlngExportAllRecords = 1
  mlngExportPickListID = 0
  mlngExportFilterID = 0
  mlngExportParent1Table = 0
  mstrExportParent1TableName = vbNullString
  mlngExportParent1FilterID = 0
  mlngExportParent2Table = 0
  mstrExportParent2TableName = vbNullString
  mlngExportParent2FilterID = 0
  mlngExportChildTable = 0
  mstrExportChildTableName = vbNullString
  mlngExportChildFilterID = 0
  mlngExportChildMaxRecords = 0
  'mstrExportOutputType = vbNullString
  'mstrExportOutputName = vbNullString
  mblnExportQuotes = False
  'mblnExportHeader = False
  mintExportHeader = 0
  mstrExportHeaderText = vbNullString
  mintExportFooter = 0
  mstrExportFooterText = vbNullString
  
  mstrExportDateFormat = vbNullString
  'mstrExportDateSeparator = vbNullString
  'mstrExportDateYearDigits = vbNullString
  mlngExportRecordCount = 0
  'mstrErrorString = vbNullString

  mlngExportParent1AllRecords = 1
  mlngExportParent1PickListID = 0
  mlngExportParent2AllRecords = 1
  mlngExportParent2PickListID = 0
  
  'mstrErrorString = vbNullString
  mblnNoRecords = False
  mlSuccessfulRecords = 0
  mlngExportRecordCount = 0
  
  ' Recordsets
  
  Set mrstExportOutput = Nothing
  
  'TM20020531 Fault 3756
  ' Delete the temptable if exists, and then clear the variable
'  If Len(mstrTempTableName) > 0 Then
'
'    mclsData.ExecuteSql ("IF EXISTS(SELECT * FROM sysobjects WHERE name = '" & mstrTempTableName & "') " & _
'                         "DROP TABLE " & mstrTempTableName)
'  End If
  datGeneral.DropUniqueSQLObject mstrTempTableName, 3
  mstrTempTableName = vbNullString
  
  ' SQL strings
  
  mstrSQLSelect = vbNullString
  mstrSQLFrom = vbNullString
  mstrSQLWhere = vbNullString
  mstrSQLJoin = vbNullString
  mstrSQLOrderBy = vbNullString
  mstrSQL = vbNullString
  
  ' Class references
  
  'Set mclsData = Nothing
  'Set mclsGeneral = Nothing
  
  ' Arrays
  
  ReDim mvarColDetails(UBound(mvarColDetails, 1), 0)
  ReDim mvarSortOrder(2, 0)
  
  ' Column Privilege arrays / collections / variables
  
  mstrBaseTableRealSource = vbNullString
  mstrRealSource = vbNullString
  'Set mobjTableView = Nothing
  'Set mobjColumnPrivileges = Nothing
  ReDim mlngTableViews(2, 0)
  ReDim mstrViews(0)
  
  gobjProgress.ResetBar2
  mblnUserCancelled = False

  ClearUp = True
  
  Exit Function
  
ClearUp_ERROR:
  
  mstrErrorString = "Error whilst clearing data." & vbCrLf & "(" & Err.Description & ")"
  ClearUp = False
  
End Function

Private Function ExportData() As Boolean

  On Error GoTo ExportData_ERROR
  
  'gobjProgress.UpdateProgress gblnBatchMode
  
  mlSuccessfulRecords = 0
  
  ' If user cancels the export, abort
  If gobjProgress.Cancelled Then
    mblnUserCancelled = True
    ExportData = False
    Exit Function
  End If

  Select Case mlngOutputFormat
    Case fmtXML
      ExportData = ExportData_XML

    Case fmtCMGFile
      ExportData = ExportData_CMGfile

    Case fmtSQLTable
      ExportData = ExportData_SQLTable

    Case Else
      If ReadDataIntoArray Then
        ExportData = SendArrayToOutputOptions
      End If

  End Select

  Exit Function
  
ExportData_ERROR:
  
  mstrErrorString = "Error whilst exporting data." & vbCrLf & "(" & Err.Description & ")"
  ExportData = False
  
End Function

'Private Function ExportData_Delimited() As Boolean
'
'  ' The export is to an delimited file.
'
'  Dim pintFileNo As Integer
'  Dim pstrExportString As String
'  'Dim pstrDateString As String
'  Dim prstReadyToExport As Recordset
'  Dim pintLoop As Integer
'  Dim lngRecordNumber As Long
'  Dim bAppended As Boolean
'  Dim bForceHeader As Boolean
'  Dim bNoRecords As Boolean
'
'  Dim tmpDec As Long
'  Dim tmpLen As Long
'
'  Dim objExpr As clsExprExpression
'
'  On Error GoTo ExportData_Delimited_ERROR
'
'  ' If filename specified already exists then delete it first.
'  If Len(Dir(mstrOutputFilename)) > 0 Then
'    If mlngOutputSaveExisting = 4 Then
'      bAppended = True
'    Else
'      bAppended = False
'      Kill mstrOutputFilename
'    End If
'  Else
'    bAppended = False
'  End If
'
'  ' Open file for output.
'  pintFileNo = FreeFile
'
'  If bAppended Then
'    Open mstrOutputFilename For Append As pintFileNo
'  Else
'    Open mstrOutputFilename For Output As pintFileNo
'  End If
'
'  ' Open the export table as a recordset.
'  Set prstReadyToExport = mclsGeneral.GetRecords("SELECT * FROM [" & mstrTempTableName & "]")
'
'  If (prstReadyToExport.BOF And prstReadyToExport.EOF) Then
'    mstrErrorString = "No records to export."
'    mclsData.ExecuteSql ("DROP TABLE [" & mstrTempTableName & "]")
'    ExportData_Delimited = False
'    bForceHeader = mbForceHeader
'    bNoRecords = True
'  Else
'    bForceHeader = False
'    bNoRecords = False
'  End If
'
'  ' Omit the header if we are appending to file (if specified)
'  If Not (bAppended And mbOmitHeader) Or bForceHeader Then
'
'    Select Case mintExportHeader
'    Case 1  'Column Names
'      'MH20030120
'      'pstrExportString = vbNullString
'      'For pintLoop = 0 To prstReadyToExport.Fields.Count - 1
'      '  pstrExportString = pstrExportString & IIf(Len(pstrExportString) > 0, ",", "") & prstReadyToExport.Fields(pintLoop).Name
'      'Next pintLoop
'      pstrExportString = vbNullString
'      For pintLoop = 0 To UBound(mvarColDetails, 2)
'        pstrExportString = pstrExportString & _
'          IIf(pstrExportString <> vbNullString, ",", vbNullString) & _
'          mvarColDetails(13, pintLoop)
'      Next
'      Print #pintFileNo, pstrExportString
'    Case 2  'Custom Heading
'      Print #pintFileNo, mstrExportHeaderText
'    End Select
'
'  End If
'
'  ' If we only need to output the header bomb out here
'  If bNoRecords Then
'    Close #pintFileNo
'    Exit Function
'  End If
'
'  ' Loop through the export table and print stuff to file.
'  lngRecordNumber = 0
'  Do While Not prstReadyToExport.EOF
'
'    ' If user cancels the export, abort
'    If gobjProgress.Cancelled Then
'      mblnUserCancelled = True
'      Close #pintFileNo
'      ExportData_Delimited = False
'      Exit Function
'    End If
'
'    pstrExportString = vbNullString
'
'    ' Loop through the fields in each record, adding them and the delimiter to the export string.
'    lngRecordNumber = lngRecordNumber + 1
'    For pintLoop = 0 To prstReadyToExport.Fields.Count - 1
'
'
'      'MH20000705 Fault 541
'      'When exporting a sort code it was being converted to a date !!??!??!
'      'so check if its a column, if yes then check if it is a date column
'
'      '' IsDate isnt perfect, so check for a ':' char too...if its there, it is a date
'      'If IsDate(prstReadyToExport.Fields(pintLoop)) And InStr(prstReadyToExport.Fields(pintLoop), ":") = 0 Then
'
'      If mvarColDetails(0, pintLoop + 1) = "N" Then
'        'Record Number
'        pstrExportString = pstrExportString & IIf(mblnExportQuotes, Chr(34), "") & CStr(lngRecordNumber) & _
'        IIf(mblnExportQuotes, Chr(34), "") & mstrExportDelimiter
'
'      ElseIf mvarColDetails(7, pintLoop + 1) = True Then
'
'
'        'pstrExportString = pstrExportString & Format(prstReadyToExport.Fields(pintLoop), mstrExportDateFormat) & _
'                           IIf(mstrExportDelimiter <> ",", vbTab, ",")
'
'
'              If Len(Format(prstReadyToExport.Fields(pintLoop), mstrExportDateFormat)) > mvarColDetails(6, pintLoop + 1) Then
''                pstrExportString = pstrExportString & Left(Format(prstReadyToExport.Fields(pintLoop), mstrExportDateFormat), mvarColDetails(6, pintLoop + 1)) & _
'                IIf(mstrExportDelimiter <> ",", vbTab, ",")
'                pstrExportString = pstrExportString & IIf(mblnExportQuotes, Chr(34), "") & Left(Format(prstReadyToExport.Fields(pintLoop), mstrExportDateFormat), mvarColDetails(6, pintLoop + 1)) & _
'                IIf(mblnExportQuotes, Chr(34), "") & mstrExportDelimiter
'              Else
''                pstrExportString = pstrExportString & Format(prstReadyToExport.Fields(pintLoop), mstrExportDateFormat) & _
'                IIf(mstrExportDelimiter <> ",", vbTab, ",")
'                pstrExportString = pstrExportString & IIf(mblnExportQuotes, Chr(34), "") & Format(prstReadyToExport.Fields(pintLoop), mstrExportDateFormat) & _
'                IIf(mblnExportQuotes, Chr(34), "") & mstrExportDelimiter
'              End If
'
'
'      ElseIf mvarColDetails(12, pintLoop + 1) Then
'
'        '#RH 04/11
'        ' Request from m edwynn - trim the data when exporting to a csv
'
''        pstrExportString = pstrExportString & Trim(prstReadyToExport.Fields(pintLoop)) & _
'                           IIf(mstrExportDelimiter <> ",", vbTab, ",")
'
'                'TM20011010 Fault 2197
'                'When creating the string, use the FormatNumeric() function so the
'                'Size and Decimals of the numeric can be formatted as required.
'                tmpLen = mvarColDetails(6, pintLoop + 1)
'                tmpDec = mvarColDetails(11, pintLoop + 1)
'
'                'TM20020507 Fault 3840 - Stupid NULLs mucked it up.
'                pstrExportString = pstrExportString & IIf(mblnExportQuotes, Chr(34), "") & FormatNumeric(IIf(IsNull(prstReadyToExport.Fields(pintLoop)), 0, prstReadyToExport.Fields(pintLoop)), tmpLen, tmpDec) & _
'                  IIf(mblnExportQuotes, Chr(34), "") & mstrExportDelimiter
'
''                If Len(prstReadyToExport.Fields(pintLoop)) > mvarColDetails(6, pintLoop + 1) Then
'''                  pstrExportString = pstrExportString & Left(prstReadyToExport.Fields(pintLoop), mvarColDetails(6, pintLoop + 1)) & _
''                  IIf(mstrExportDelimiter <> ",", vbTab, ",")
''                  pstrExportString = pstrExportString & IIf(mblnExportQuotes, Chr(34), "") & FormatNumeric(prstReadyToExport.Fields(pintLoop), tmpLen, tmpDec) & _
''                  IIf(mblnExportQuotes, Chr(34), "") & IIf(mstrExportDelimiter <> ",", vbTab, ",")
''                Else
'''                  pstrExportString = pstrExportString & prstReadyToExport.Fields(pintLoop) & _
''                  IIf(mstrExportDelimiter <> ",", vbTab, ",")
''                  pstrExportString = pstrExportString & IIf(mblnExportQuotes, Chr(34), "") & FormatNumeric(prstReadyToExport.Fields(pintLoop), tmpLen, tmpDec) & _
''                  IIf(mblnExportQuotes, Chr(34), "") & IIf(mstrExportDelimiter <> ",", vbTab, ",")
''                End If
'
'
'      Else
'
'         'RH - BUG 614
'
'        If Len(prstReadyToExport.Fields(pintLoop)) > mvarColDetails(6, pintLoop + 1) Then
'
'          ' Need to trim the data, its longer than the length required to export
'
'          ' Trim is here
''          pstrExportString = pstrExportString & IIf(mblnExportQuotes, Chr(34), "") & Left(Trim(prstReadyToExport.Fields(pintLoop)), mvarColDetails(6, pintLoop + 1)) & _
'           IIf(mblnExportQuotes, Chr(34), "") & IIf(mstrExportDelimiter <> ",", vbTab, ",")
'
'          ' Take away the trim - ruins working pattern field !
'          pstrExportString = pstrExportString & IIf(mblnExportQuotes, Chr(34), "") & Left(prstReadyToExport.Fields(pintLoop), mvarColDetails(6, pintLoop + 1)) & _
'           IIf(mblnExportQuotes, Chr(34), "") & mstrExportDelimiter
'
'        Else
'
'          ' Data is fine as it is (dont pad with spaces for delimited files!)
'
'          ' Trim is here
'
''          pstrExportString = pstrExportString & IIf(mblnExportQuotes, Chr(34), "") & Left(Trim(prstReadyToExport.Fields(pintLoop)), mvarColDetails(6, pintLoop + 1)) & _
'           IIf(mblnExportQuotes, Chr(34), "") & IIf(mstrExportDelimiter <> ",", vbTab, ",")
'
'          ' Take away the trim - ruins working pattern field !!!
'          pstrExportString = pstrExportString & IIf(mblnExportQuotes, Chr(34), "") & Left(prstReadyToExport.Fields(pintLoop), mvarColDetails(6, pintLoop + 1)) & _
'           IIf(mblnExportQuotes, Chr(34), "") & mstrExportDelimiter
'
'        End If
'
'      End If
'
'    Next pintLoop
'
'    ' If the delimiter is "," then remove the last character.
'    If Not mstrExportDelimiter = vbTab Then
'      pstrExportString = Left(pstrExportString, Len(pstrExportString) - 1)
'    End If
'
'    ' Print the stuff to the file.
'    Print #pintFileNo, pstrExportString
'
'    'JDM - 13/12/01 - Fault 3280 - Log successful records
'    If mbLoggingExportSuccess Then
'      gobjEventLog.AddDetailEntry pstrExportString & " Exported successfully"
'    End If
'
'    mlSuccessfulRecords = mlSuccessfulRecords + 1
'
'    prstReadyToExport.MoveNext
'
'  Loop
'
'
'  Select Case mintExportFooter
'  Case 1  'Column Names
'    pstrExportString = vbNullString
'    For pintLoop = 0 To prstReadyToExport.Fields.Count - 1
'      pstrExportString = pstrExportString & IIf(Len(pstrExportString) > 0, ",", "") & prstReadyToExport.Fields(pintLoop).Name
'    Next pintLoop
'    Print #pintFileNo, pstrExportString
'  Case 2  'Custom Heading
'    Print #pintFileNo, mstrExportFooterText
'  End Select
'
'
'  ' Close the final output file
'  Close #pintFileNo
'
'  'Drop the export table as we dont need it now.
'  mclsData.ExecuteSql ("DROP TABLE [" & mstrTempTableName & "]")
'
'  Set prstReadyToExport = Nothing
'  ExportData_Delimited = True
'  Exit Function
'
'ExportData_Delimited_ERROR:
'
'  Select Case Err.Number
'    Case 70 ' file creation error / sharing violation
'      mstrErrorString = "Error creating file '" & mstrOutputFilename & "'. File may already be in use." & vbCrLf & "(" & Err.Description & ")"
'    Case 76 ' path not found error
'      mlngExportRecordCount = 0
'      mstrErrorString = "Error whilst exporting to delimited file." & vbCrLf & "(" & Err.Description & ")"
'    Case Else
'    mstrErrorString = "Error whilst exporting to delimited file." & vbCrLf & "(" & Err.Description & ")"
'    Close #pintFileNo
'  End Select
'
'  ExportData_Delimited = False
'
'End Function

Private Function FormatNumeric(dValue As Double, iLen As Long, iDec As Long) As Variant

  Const sDecimalChar = "."

  Dim iCount As Integer
  Dim iDecimalCount As Integer
  Dim iLengthCount As Integer
  Dim sCurrDig As String
  Dim bHasDec As Boolean
  Dim sDecimalFormat As String
  Dim i As Integer
  
  iCount = 0
  iDecimalCount = 0
  iLengthCount = 0
  sCurrDig = vbNullString
  bHasDec = False
  sDecimalFormat = vbNullString
  
  If iLen > 0 Then
    Do While iCount <= iLen
      iCount = iCount + 1
      sCurrDig = Mid(CStr(dValue), iCount, 1)
      
      If sCurrDig = sDecimalChar Then
        bHasDec = True
      Else
        If bHasDec And iDecimalCount < iDec Then
          iDecimalCount = iDecimalCount + 1
          iLengthCount = iLengthCount + 1
        ElseIf Not bHasDec Then
          iLengthCount = iLengthCount + 1
        End If
      End If
    Loop

    'TM20020417 Fault 3772
    If iDec > 0 Then
      For i = 1 To iDec Step 1
        sDecimalFormat = sDecimalFormat & "0"
      Next i
      If LenB(sDecimalFormat) <> 0 Then
        sDecimalFormat = "###0." & sDecimalFormat
      End If
      FormatNumeric = Format(Round(dValue, iDecimalCount), sDecimalFormat)
    
    Else
      FormatNumeric = Round(dValue, iDecimalCount)
    End If
    
  Else
    If iDecimalCount > 0 Then
      FormatNumeric = Round(dValue, iDecimalCount)
    Else
      FormatNumeric = CStr(dValue)
    End If
  End If

End Function

Public Function PrettyPrintXml(xmlDoc)
  Dim rdr, wrt
  Set rdr = CreateObject("Msxml2.SAXXMLReader.3.0")
  Set wrt = CreateObject("Msxml2.MXXMLWriter.3.0")
  wrt.Encoding = "UTF-8"
  wrt.standalone = False
  wrt.byteOrderMark = True
  wrt.omitXMLDeclaration = False
  wrt.Indent = True

  Set rdr.contentHandler = wrt
  Set rdr.dtdHandler = wrt
  Set rdr.errorHandler = wrt
  rdr.PutProperty "http://xml.org/sax/properties/lexical-handler", wrt
  rdr.PutProperty "http://xml.org/sax/properties/declaration-handler", wrt
  rdr.Parse xmlDoc
  
  Dim sRemoveHeader As String
  sRemoveHeader = "standalone=""no"""
  PrettyPrintXml = Replace(wrt.Output, sRemoveHeader, "")
  PrettyPrintXml = Replace(PrettyPrintXml, "encoding=""UTF-16""", "encoding=""UTF-8""")
  
End Function

Private Function ExportData_XML() As Boolean

  Dim pintFileNo As Integer
  Dim bAppended As Boolean
  Dim sExportText As String
  Dim sTransFormText As String
  Dim sSQL As String
  Dim sSQLGetData As String
  Dim prstReadyToExport As Recordset
  Dim rstConverted As Recordset
  Dim iFileCount As Integer
  Dim strOutputFilename As String
  Dim sLimitedColumns As String
  Dim iCount As Integer
  
  Dim sXSDName As String
  Dim sXSDTransformFile As String
  Dim sWhereClause As String
  
  strOutputFilename = mstrOutputFileName
  sXSDName = mstrXSDFilename
  sXSDTransformFile = mstrTransformFile
  
  If Not mbPreserveTransformPath And Len(mstrTransformFile) > 0 Then
    sXSDTransformFile = GetFileNameOnly(mstrTransformFile)
  End If
  
  If Not mbPreserveXSDPath And Len(mstrXSDFilename) > 0 Then
    sXSDName = GetFileNameOnly(mstrXSDFilename)
  End If
      
  ' Limit the column selection to actual columns the user has selected (XML has nto include the ID for the splitter to work)
  sLimitedColumns = ""
  For iCount = 1 To UBound(mvarColDetails, 2) - 1
    sLimitedColumns = sLimitedColumns & IIf(Len(sLimitedColumns) > 0, ", ", "") & "[" & mvarColDetails(14, iCount) & "]"
  Next
      
  If mbSplitXMLIntoFiles Then
    sWhereClause = " WHERE basetable._XMLSplitID = _XMLSplitID "
  End If
   
  sSQLGetData = "SELECT CAST((SELECT " & sLimitedColumns & " FROM [" & mstrTempTableName & "]" & sWhereClause & " FOR XML PATH('" & mstrXMLDataNodeName & "'), ROOT('" & mstrExportHeaderText & "')" & _
    ", ELEMENTS XSINIL) AS VARCHAR(MAX)) AS XmlData"
  
  If mbSplitXMLIntoFiles Then
    sSQLGetData = "SELECT XMLData =(" & sSQLGetData & ") FROM [" & mstrTempTableName & "] basetable"
  End If
   
  If Len(sXSDName) > 0 Then
    sSQLGetData = "WITH XMLNAMESPACES ('" & sXSDName & "' as noNamespaceSchemaLocation)" & sSQLGetData
  End If
   
  Set prstReadyToExport = mclsData.OpenRecordset(sSQLGetData, adOpenForwardOnly, adLockReadOnly)

  If (prstReadyToExport.BOF And prstReadyToExport.EOF) Then
    mstrErrorString = "No records to export."
    mclsData.ExecuteSql ("IF EXISTS(SELECT Name FROM sysobjects WHERE name = '" & mstrTempTableName & "') " & _
                         "DROP TABLE " & mstrTempTableName)
    ExportData_XML = False
    Exit Function
  Else
    prstReadyToExport.MoveFirst

    Do While Not prstReadyToExport.EOF

      iFileCount = iFileCount + 1

      If Len(sXSDTransformFile) > 0 Then
        sExportText = prstReadyToExport(0).Value
        sTransFormText = GetFileText(sXSDTransformFile)
        sSQL = "SELECT convert(nvarchar(MAX), dbo.udfASRNetApplyXsltTransform(convert(xml,'" & Replace(sExportText, "'", "''") & "'), convert(xml,'" & Replace(sTransFormText, "'", "''") & "')))"
        Set rstConverted = mclsData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
        sExportText = PrettyPrintXml(rstConverted(0).Value)
      Else
        If Not IsNull(prstReadyToExport(0).Value) Then
          sExportText = prstReadyToExport(0).Value
          sExportText = PrettyPrintXml(sExportText)
          
          If Len(sXSDName) > 0 Then
            sExportText = Replace(sExportText, "xmlns:noNamespaceSchemaLocation", "xsi:noNamespaceSchemaLocation")
          End If
          
        End If
      End If

      If mbSplitXMLIntoFiles Then
        strOutputFilename = InsertNumberIntoFilename(mstrOutputFileName, iFileCount)
      End If

      ' Process the file name
      strOutputFilename = ReplaceFormatExpressions(strOutputFilename, 1, mlngExportRecordCount, True)

      Select Case mlngOutputSaveExisting
      Case 0  ' Overwrite
        bAppended = False
       ' Kill mstrOutputFileName
        
      Case 1  ' Do not overwrite (fail)
        mstrErrorString = "File already exists."
        ExportData_XML = False
        Exit Function
    
      Case 2  ' Add sequential number to filename
        mstrOutputFileName = GetSequentialNumberedFile(strOutputFilename)
        bAppended = False
      
      Case 3  ' Append to Existing File
        bAppended = True
        
      End Select
  
      ' Open file for output.
      pintFileNo = FreeFile
    
      If bAppended Then
        Open strOutputFilename For Append As pintFileNo
      Else
        Open strOutputFilename For Output As pintFileNo
      End If
       
      Print #pintFileNo, sExportText
       
      Close #pintFileNo
      
    prstReadyToExport.MoveNext
  Loop
  
  End If
  
  ExportData_XML = True

End Function

Private Function ExportData_SQLTable() As Boolean

  On Error GoTo ErrorTrap

  Dim rstMatchingColumns As Recordset
  Dim sSQL As String
  Dim bOK As Boolean
  Dim sInsertString As String
  
  bOK = True
  sInsertString = ""

  If Not CheckLinkedServerIsValid(mstrLinkedServer, mstrLinkedCatalog, mstrLinkedTable) Then
    mstrErrorString = "Error establishing connection to the linked server." & vbCrLf
    ExportData_SQLTable = False
    Exit Function
  End If

  sSQL = "SELECT t.Column_Name" & _
    " FROM [" & mstrLinkedServer & "].[" & mstrLinkedCatalog & "].INFORMATION_SCHEMA.COLUMNS t" & _
    " INNER JOIN ASRSysExportDetails s ON s.Heading = t.COLUMN_NAME COLLATE DATABASE_DEFAULT" & _
    " WHERE t.TABLE_NAME = '" & mstrLinkedTable & "' AND s.ExportID = " & mlngExportID
  Set rstMatchingColumns = mclsData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  With rstMatchingColumns
    
    If .BOF And .EOF Then
      ExportData_SQLTable = False
      mstrErrorString = "No matching columns found in the specified Export definition." & vbCrLf & "Please ensure your column names are correct."
      Exit Function
    End If
    
    Do Until .EOF
      sInsertString = sInsertString & IIf(Len(sInsertString) > 0, ", ", "") & "[" & .Fields("Column_Name").Value & "]"
      .MoveNext
    Loop
  End With

  ' Open the export table as a recordset.
  sSQL = "INSERT [" & mstrLinkedServer & "].[" & mstrLinkedCatalog & "].dbo.[" & mstrLinkedTable & "] (" & _
    sInsertString & ") SELECT " & sInsertString & " FROM [" & mstrTempTableName & "]"
  mclsData.ExecuteSql sSQL
  
  
TidyUpAndExit:
  ExportData_SQLTable = bOK
  Exit Function

ErrorTrap:
  bOK = False
  GoTo TidyUpAndExit

End Function

Private Function ExportData_CMGfile() As Boolean

  'JDM - 21/08/01 - Fault 2706 - Not to overwrite file if no records found

  ' The export is to a CMG file.
    
  Dim iLoop As Integer
  Dim pintFileNo As Integer
  Dim pstrExportString As DataMgr.clsStringBuilder
  Dim strExportColumn As String
  Dim prstReadyToExport As Recordset
  Dim pintLoop As Integer
  Dim bOutputColumn As Boolean
  Dim dLastChangeDate As Date
  Dim strRecordIdentifier As String
  Dim bCMGExportFileCode As Boolean
  Dim bCMGExportFieldCode As Boolean
  Dim bCMGExportLastChangeDate As Boolean
  Dim iCMGExportFileCodeSize As Integer
  Dim iCMGEXportRecordIDSize As Integer
  Dim iCMGExportFieldCodeSize As Integer
  Dim iCMGExportOutputColumnSize As Integer
  Dim iCMGExportLastChangeDateSize As Integer
  'NPG20090313 Fault 13595
  Dim iCMGExportFileCodeOrderID As Integer
  Dim iCMGEXportRecordIDOrderID As Integer
  Dim iCMGExportFieldCodeOrderID As Integer
  Dim iCMGExportOutputColumnOrderID As Integer
  Dim iCMGExportLastChangeDateOrderID As Integer
  
  Dim bExportFileOpened As Boolean
  Dim bUseCSV As Boolean
  Dim bCMGIgnoreBlanks As Boolean
  Dim bCMGReverseOutput As Boolean
  Dim iCount As Integer
  Dim bAppended As Boolean
  
  Dim lngRecordCount As Long
  Dim astrColumnIDs() As String
  Dim astrRecordIDs() As String
  Dim astrBulkRecordIDs() As String
  'NPG20071218 Fault 12867
  Dim astrConvertCase() As String
  'NPG20080617 Suggestion S000816
  Dim astrSuppressNulls() As String
  Dim strRecordIDs As String
  Dim strColumnIDs As String
  'NPG20071218 Fault 12867
  Dim strConvertCase As String
  'NPG20080617 Suggestion S000816
  Dim strSuppressNulls As String
  Dim sSQL As String
  Dim rsAffectedRecords As ADODB.Recordset
  Dim varNewValue As String
  Dim astrColumnCodes() As String
  Dim aiDataTypes() As Integer
  Dim pbNullSuppressed As Boolean
  'Dim lngLastColumnID As Long
  Dim strCMGCode As String
  Dim iDataType As Integer
  Dim fNewLine As Boolean
    
  Dim iRow As Integer
  Dim fNewRow As Boolean
  
  On Error GoTo ExportData_CMGfile_ERROR
  
  Set pstrExportString = New DataMgr.clsStringBuilder
  Set rsAffectedRecords = New ADODB.Recordset
  
  ' Which fields to export and their size
  bUseCSV = GetSystemSetting("CMGExport", "UseCSV", False)
  bCMGIgnoreBlanks = GetSystemSetting("CMGExport", "IgnoreBlanks", False)
  bCMGReverseOutput = GetSystemSetting("CMGExport", "ReverseOutput", False)
  bCMGExportFileCode = GetSystemSetting("CMGExport", "FileCode", True)
  bCMGExportFieldCode = GetSystemSetting("CMGExport", "FieldCode", True)
  bCMGExportLastChangeDate = GetSystemSetting("CMGExport", "LastChange", True)
  iCMGExportFileCodeSize = GetSystemSetting("CMGExport", "FileCodeSize", 6)
  iCMGEXportRecordIDSize = GetSystemSetting("CMGExport", "RecordIdentifierSize", 11)
  iCMGExportFieldCodeSize = GetSystemSetting("CMGExport", "FieldCodeSize", 10)
  iCMGExportOutputColumnSize = GetSystemSetting("CMGExport", "OutputColumnSize", 53)
  iCMGExportLastChangeDateSize = GetSystemSetting("CMGExport", "LastChangeSize", 8)
  'NPG20090313 Fault 13595
  iCMGExportFileCodeOrderID = GetSystemSetting("CMGExport", "FileCodeOrderID", 0)
  iCMGEXportRecordIDOrderID = GetSystemSetting("CMGExport", "RecordIdentifierOrderID", 0)
  iCMGExportFieldCodeOrderID = GetSystemSetting("CMGExport", "FieldCodeOrderID", 0)
  iCMGExportOutputColumnOrderID = GetSystemSetting("CMGExport", "OutputColumnOrderID", 0)
  iCMGExportLastChangeDateOrderID = GetSystemSetting("CMGExport", "LastChangeOrderID", 0)
  
  bExportFileOpened = False
  
  ' Constantly exported strings
  If bUseCSV Then
    'NPG20090403 Fault 13636
     'mstrExportFileCode = mstrExportFileCode
  Else
    mstrExportFileCode = SetStringLength(mstrExportFileCode, iCMGExportFileCodeSize)
  End If
  
  ' Build a list of all the columnids we'll be searching on
  ' Use 2 arrays, because we'll be using the join statement later on in the code
  ReDim astrColumnIDs(UBound(mvarColDetails, 2))
  ReDim astrColumnCodes(UBound(mvarColDetails, 2))
  ReDim aiDataTypes(UBound(mvarColDetails, 2))
  'NPG20071218 Fault 12867
  ReDim astrConvertCase(UBound(mvarColDetails, 2))
  'NPG20080617 Suggestion S000816
  ReDim astrSuppressNulls(UBound(mvarColDetails, 2))
  
  iCount = 0
  For pintLoop = 0 To UBound(mvarColDetails, 2)
    If mvarColDetails(9, pintLoop) Then
      astrColumnIDs(iCount) = mvarColDetails(3, pintLoop)    ' Column ID
      aiDataTypes(iCount) = datGeneral.GetColumnDataType(CLng(astrColumnIDs(iCount)))
      'NPG20071218 Fault 12867
      astrConvertCase(iCount) = mvarColDetails(16, pintLoop)
      'NPG20080617 Suggestion S000816
      astrSuppressNulls(iCount) = mvarColDetails(17, pintLoop)
      
      If bUseCSV Then
        astrColumnCodes(iCount) = mvarColDetails(8, pintLoop)
      Else
        astrColumnCodes(iCount) = SetStringLength(mvarColDetails(8, pintLoop), iCMGExportFieldCodeSize)
      End If
      
      iCount = iCount + 1
    End If
  Next pintLoop


  If iCount = 0 Then
    mstrErrorString = "The export definition contains no columns which are marked for audit."
    ExportData_CMGfile = False
    Exit Function
  End If

  ReDim Preserve astrColumnCodes(iCount - 1)
  ReDim Preserve astrColumnIDs(iCount - 1)
  ReDim Preserve aiDataTypes(iCount - 1)
  'NPG20071218 Fault 12867
  ReDim Preserve astrConvertCase(iCount - 1)
  'NPG20080617 Suggestion S000816
  ReDim Preserve astrSuppressNulls(iCount - 1)
  
  ' Open the export table as a recordset.
  Set prstReadyToExport = mclsData.OpenRecordset("SELECT ID,Identifier FROM [" & mstrTempTableName & "]", adOpenForwardOnly, adLockReadOnly)

  If (prstReadyToExport.BOF And prstReadyToExport.EOF) Then
    mstrErrorString = "No records to export."
    mclsData.ExecuteSql ("IF EXISTS(SELECT Name FROM sysobjects WHERE name = '" & mstrTempTableName & "') " & _
                         "DROP TABLE " & mstrTempTableName)
    ExportData_CMGfile = False
    Exit Function
  End If

  ReDim astrRecordIDs(0)
  ReDim astrBulkRecordIDs(0)
  lngRecordCount = 0
  mlngExportRecordCount = 0
  strColumnIDs = Join(astrColumnIDs, ",")
 
  ' Loop through the export table and print stuff to file.
  Do While Not prstReadyToExport.EOF
    
    If lngRecordCount > UBound(astrRecordIDs) Then ReDim Preserve astrRecordIDs(lngRecordCount + 100)
    astrRecordIDs(lngRecordCount) = prstReadyToExport.Fields(0).Value
    
    If ((lngRecordCount + 1) Mod 1000) = 0 Then
      ReDim Preserve astrBulkRecordIDs(UBound(astrBulkRecordIDs) + 1)
    Else
      astrBulkRecordIDs(UBound(astrBulkRecordIDs)) = astrBulkRecordIDs(UBound(astrBulkRecordIDs)) & _
        IIf(lngRecordCount = 0, "", ",")
    End If
    
    astrBulkRecordIDs(UBound(astrBulkRecordIDs)) = astrBulkRecordIDs(UBound(astrBulkRecordIDs)) & _
      CStr(prstReadyToExport.Fields(0).Value)
    
    ' If user cancels the export, abort
    If gobjProgress.Cancelled Then
      mblnUserCancelled = True
      Close #pintFileNo
      ExportData_CMGfile = False
      Exit Function
    Else
      If lngRecordCount Mod 50 = 0 Then
        gobjProgress.Bar1Caption = "Calculating Audit Records ( " & Trim$(Str(lngRecordCount)) & " of " & Trim(Str(prstReadyToExport.RecordCount)) & " )"
      End If
    End If
      
    ' Get audit transaction for this record
    sSQL = "SELECT ISNULL(NewValue, ''), DateTimeStamp, ColumnID " & _
            "FROM ASRSysAuditTrail WHERE ID IN(SELECT MAX(ID) FROM ASRSysAuditTrail " & _
            " WHERE ColumnID IN (" & strColumnIDs & ") And RecordID = " & astrRecordIDs(lngRecordCount) & _
            " AND CMGCommitDate IS null AND ISNULL(OldValue,'') <> ISNULL(NewValue,'') AND NOT (OldValue = '* New Record *' AND NewValue = '')" & _
            " GROUP BY ColumnID)"
    rsAffectedRecords.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly

    ' Loop through each of the returned audit records
    pstrExportString.TheString = vbNullString

    If Not rsAffectedRecords.EOF Then
      strRecordIdentifier = IIf(Not IsNull(prstReadyToExport.Fields(1).Value), prstReadyToExport.Fields(1).Value, vbNullString)

      If bUseCSV Then
        ' strRecordIdentifier = strRecordIdentifier & ","
      Else
        strRecordIdentifier = SetStringLength(strRecordIdentifier, iCMGEXportRecordIDSize)
      End If
    End If

    Do While Not rsAffectedRecords.EOF
      pbNullSuppressed = False
      'NPG20080617 Suggestion S000816
      ' This for/next block has been moved above the [File Export Code] and [Record Identifier] sections
      ' so that we can exclude any null values completely - otherwise the identifier goes into the export file.

      For iCount = 0 To UBound(astrColumnIDs, 1)
        If astrColumnIDs(iCount) = rsAffectedRecords.Fields(CMGFields.ColumnID).Value Then
          strCMGCode = astrColumnCodes(iCount)
          iDataType = aiDataTypes(iCount)
          'NPG20071218 Fault 12867
          strConvertCase = astrConvertCase(iCount)
          'NPG20080617 Suggestion S000816
          strSuppressNulls = astrSuppressNulls(iCount)
          Exit For
        End If
      Next iCount

      'NPG20080617 Suggestion S000816
      If strSuppressNulls _
        And (rsAffectedRecords.Fields(CMGFields.NewValue).Value = vbNullString _
              Or (Left(rsAffectedRecords.Fields(CMGFields.NewValue).Value, 2) = "0." And Val(rsAffectedRecords.Fields(CMGFields.NewValue).Value) = "0") _
              Or rsAffectedRecords.Fields(CMGFields.NewValue).Value = "0" _
              Or rsAffectedRecords.Fields(CMGFields.NewValue).Value = "00:00" _
              Or rsAffectedRecords.Fields(CMGFields.NewValue).Value = "False") Then
        pbNullSuppressed = True
      End If
      
      'NPG20080731 Fault 13305
      If fNewLine And Not pbNullSuppressed Then pstrExportString.Append vbNewLine
        
      fNewLine = False
      
      'NPG20090313 Fault 13595
      If Not pbNullSuppressed Then
      
      'NPG20090403 Fault 13636
      fNewRow = True
      
      For iRow = 0 To 4
        If iCMGExportFileCodeOrderID + iCMGEXportRecordIDOrderID + iCMGExportFieldCodeOrderID + _
              iCMGExportOutputColumnOrderID + iCMGExportLastChangeDateOrderID = 0 _
              And iRow > 0 Then Exit For
                                            
        ' File Export Code (padded to specified length)
        'NPG20090313 Fault 13595
        ' If bCMGExportFileCode Then pstrExportString.Append mstrExportFileCode
        If iCMGExportFileCodeOrderID = iRow And bCMGExportFileCode Then
          ' Insert the delimeter if required.
          If bUseCSV And Not fNewRow Then pstrExportString.Append "," ' may need to set fnewrow to false here...
          pstrExportString.Append mstrExportFileCode
          fNewRow = False
        End If

        ' Record Identifier (usually staff number)
        'NPG20090313 Fault 13595
        ' pstrExportString.Append strRecordIdentifier
        If iCMGEXportRecordIDOrderID = iRow Then
          ' Insert the delimeter if required.
          If bUseCSV And Not fNewRow Then pstrExportString.Append "," ' may need to set fnewrow to false here...
          pstrExportString.Append strRecordIdentifier
          fNewRow = False
        End If

        ' Code for the column (padded to specified length)
        'NPG20090313 Fault 13595
        ' If bCMGExportFieldCode Then pstrExportString.Append strCMGCode
        If iCMGExportFieldCodeOrderID = iRow And bCMGExportFieldCode Then
          ' Insert the delimeter if required.
          If bUseCSV And Not fNewRow Then pstrExportString.Append "," ' may need to set fnewrow to false here...
          pstrExportString.Append strCMGCode
          fNewRow = False
        End If

        ' New value
        ' varNewValue = IIf(IsNull(rsAffectedRecords.Fields(CMGFields.NewValue).Value), vbNullString, rsAffectedRecords.Fields(CMGFields.NewValue).Value)

        'NPG20071218 Fault 12867
        Select Case strConvertCase
          Case "1"
              varNewValue = IIf(IsNull(rsAffectedRecords.Fields(CMGFields.NewValue).Value), vbNullString, UCase(rsAffectedRecords.Fields(CMGFields.NewValue).Value))
          Case "2"
              varNewValue = IIf(IsNull(rsAffectedRecords.Fields(CMGFields.NewValue).Value), vbNullString, LCase(rsAffectedRecords.Fields(CMGFields.NewValue).Value))
          Case Else
              varNewValue = IIf(IsNull(rsAffectedRecords.Fields(CMGFields.NewValue).Value), vbNullString, rsAffectedRecords.Fields(CMGFields.NewValue).Value)
        End Select

        ' Format the new value
        If iDataType = 11 Then
          strExportColumn = Format(varNewValue, mstrExportDateFormat)
        Else
          strExportColumn = varNewValue
        End If



        'NPG20090313 Fault 13595
        
        If iCMGExportFileCodeOrderID + iCMGEXportRecordIDOrderID + iCMGExportFieldCodeOrderID + iCMGExportOutputColumnOrderID + iCMGExportLastChangeDateOrderID = 0 Then
        ' The CMG layout hasn't been saved in v3.7 format, use existing values (including reverseoutput option)
          
          ' ePayFact or normal CMG output
          If bCMGReverseOutput Then
  
            ' Last changed date
            If bCMGExportLastChangeDate Then
              dLastChangeDate = IIf(IsNull(rsAffectedRecords.Fields(CMGFields.DateTime).Value), #12/31/9999#, rsAffectedRecords.Fields(CMGFields.DateTime).Value)
              pstrExportString.Append IIf(bUseCSV, "," & Format(dLastChangeDate, mstrExportDateFormat), SetStringLength(Format(dLastChangeDate, mstrExportDateFormat), iCMGExportLastChangeDateSize))
            End If
  
            ' Add the output column (padded/trimmed to specified width)
            If bUseCSV Then
              pstrExportString.Append IIf(bUseCSV, ",", vbNullString) & Trim$(strExportColumn)
            Else
              pstrExportString.Append SetStringLength(strExportColumn, iCMGExportOutputColumnSize)
            End If
  
          Else
  
            ' Add the output column (padded/trimmed to specified width)
            If bUseCSV Then
              pstrExportString.Append "," & Trim$(strExportColumn)
            Else
              pstrExportString.Append SetStringLength(strExportColumn, iCMGExportOutputColumnSize)
            End If
  
            ' Last changed date
            If bCMGExportLastChangeDate Then
              dLastChangeDate = IIf(IsNull(rsAffectedRecords.Fields(CMGFields.DateTime).Value), #12/31/9999#, rsAffectedRecords.Fields(CMGFields.DateTime).Value)
              pstrExportString.Append IIf(bUseCSV, "," & Format(dLastChangeDate, mstrExportDateFormat), SetStringLength(Format(dLastChangeDate, mstrExportDateFormat), iCMGExportLastChangeDateSize))
            End If
          End If
        Else
        ' Using new CMG layout ordering...
          If iCMGExportOutputColumnOrderID = iRow Then
            ' If iRow < 4 And bUseCSV Then pstrExportString.Append ","
            If bUseCSV Then
              ' Insert the delimeter if required.
              If Not fNewRow Then pstrExportString.Append "," ' may need to set fnewrow to false here...
              pstrExportString.Append Trim$(strExportColumn)
              fNewRow = False
            Else
              pstrExportString.Append SetStringLength(strExportColumn, iCMGExportOutputColumnSize)
              fNewRow = False
            End If
          End If
          
          If iCMGExportLastChangeDateOrderID = iRow And bCMGExportLastChangeDate Then
            ' If iRow < 4 And bUseCSV Then pstrExportString.Append ","
            dLastChangeDate = IIf(IsNull(rsAffectedRecords.Fields(CMGFields.DateTime).Value), #12/31/9999#, rsAffectedRecords.Fields(CMGFields.DateTime).Value)
            ' Insert the delimeter if required.
            If bUseCSV And Not fNewRow Then pstrExportString.Append "," ' may need to set fnewrow to false here...
            pstrExportString.Append IIf(bUseCSV, Format(dLastChangeDate, mstrExportDateFormat), SetStringLength(Format(dLastChangeDate, mstrExportDateFormat), iCMGExportLastChangeDateSize))
            fNewRow = False
          End If
          
        End If
        
      Next iRow
      
      mlngExportRecordCount = mlngExportRecordCount + 1

      End If

      ' Scroll to next record
      rsAffectedRecords.MoveNext

      ' CMG Special - Line feed after after string
      If Not rsAffectedRecords.EOF And pstrExportString.Length <> 0 Then
        'NPG20080731 Fault 13305
        'pstrExportString.Append vbNewLine
        fNewLine = True
        fNewRow = True
      End If
       
    Loop

    ' Process the file name
    mstrOutputFileName = ReplaceFormatExpressions(mstrOutputFileName, 1, mlngExportRecordCount, True)

    ' Output all the columns for this record ID
    If pstrExportString.Length <> 0 Then
      ' Open the file if not already opened
      If Not bExportFileOpened Then

        ' If filename specified already exists then delete it first.
        If Len(Dir(mstrOutputFileName)) > 0 Then

'NPG20080421 Fault 13113
'          If mlngOutputSaveExisting = 3 Then  ' Append To File
'            bAppended = True
'          Else
'            ' NPG20080218 Fault 12778
'            If mlngOutputSaveExisting = 2 Then    ' Add Sequential Number to Filename
'              mstrOutputFilename = GetSequentialNumberedFile(mstrOutputFilename)
'              bAppended = False
'            Else
'              bAppended = False
'              Kill mstrOutputFilename
'            End If
'          End If
        
          Select Case mlngOutputSaveExisting
          Case 0  ' Overwrite
            bAppended = False
            Kill mstrOutputFileName
            
          Case 1  ' Do not overwrite (fail)
            mstrErrorString = "File already exists."
            ExportData_CMGfile = False
            Exit Function
        
          Case 2  ' Add sequential number to filename
            mstrOutputFileName = GetSequentialNumberedFile(mstrOutputFileName)
            bAppended = False
          
          Case 3  ' Append to Existing File
            bAppended = True
            
          End Select
          
        Else
          bAppended = False
        End If
      
        ' Open file for output.
        pintFileNo = FreeFile
      
        If bAppended Then
          Open mstrOutputFileName For Append As pintFileNo
        Else
          Open mstrOutputFileName For Output As pintFileNo
        End If

        bExportFileOpened = True
      End If

      Print #pintFileNo, pstrExportString.ToString
    End If
  
    ' Move to next record within the selected filter
    prstReadyToExport.MoveNext
    lngRecordCount = lngRecordCount + 1
    
    rsAffectedRecords.Close
  
  Loop

  prstReadyToExport.Close

  ' Close the final output file
  Close #pintFileNo

  ' If user cancels the export, abort
  If gobjProgress.Cancelled Then
    mblnUserCancelled = True
    ExportData_CMGfile = False
    Exit Function
  Else
    gobjProgress.Bar1Caption = "Updating Audit Records..."
  End If
 
  ' Update which records have been exported
  ReDim Preserve astrRecordIDs(lngRecordCount - 1)
  
  strColumnIDs = Join(astrColumnIDs, ",")
  'strRecordIDs = Join(astrRecordIDs, ",")

  For iLoop = 0 To UBound(astrBulkRecordIDs)
    strRecordIDs = astrBulkRecordIDs(iLoop)
  
    sSQL = "Update asrsysAuditTrail Set CMGExportDate = Convert(smalldatetime,getdate()) " _
      & "Where RecordID IN (" & strRecordIDs & ") AND ColumnID IN (" & strColumnIDs & ")"
    gADOCon.Execute sSQL, , adExecuteNoRecords
  Next iLoop
  
  ' If we are to auto commit, do it here
  If mbUpdateAuditLog = True Then
    datGeneral.CMGCommit
  End If

  'Drop the export table as we dont need it now.
  mclsData.ExecuteSql ("IF EXISTS(SELECT * FROM sysobjects WHERE name = '" & mstrTempTableName & "') " & _
                       "DROP TABLE " & mstrTempTableName)

  Set prstReadyToExport = Nothing
  ExportData_CMGfile = True
  
  ' Clear up the connection
  Set rsAffectedRecords = Nothing
  
  Exit Function

ExportData_CMGfile_ERROR:

  Select Case Err.Number
    Case 70 ' file creation error / sharing violation
      mstrErrorString = "Error creating file '" & mstrOutputFileName & "'. File may already be in use." & vbCrLf & "(" & Err.Description & ")"
    Case 76 ' path not found error
      mlngExportRecordCount = 0
      mstrErrorString = "Error whilst exporting to CMG file." & vbCrLf & "(" & Err.Description & ")"
    Case Else
    mstrErrorString = "Error whilst exporting to CMG file." & vbCrLf & "(" & Err.Description & ")"
    Close #pintFileNo
  End Select
  
  ExportData_CMGfile = False

End Function

Private Function GetUniqueHeading(intIndex As Integer)

  Dim strTempHeading As String
  Dim lngCount As Long
  Dim lngLoop As Long
  Dim blnFound As Boolean

  lngCount = 1
  strTempHeading = mvarColDetails(13, intIndex)
  If strTempHeading = vbNullString Then
    strTempHeading = mvarColDetails(4, intIndex)
  End If

  Do

    blnFound = False
    For lngLoop = 1 To intIndex - 1
      If LCase(mvarColDetails(14, lngLoop)) = LCase(strTempHeading) Then
        lngCount = lngCount + 1
        blnFound = True
      End If
    Next

    If blnFound Then
      strTempHeading = mvarColDetails(13, intIndex) & CStr(lngCount)
    End If

  Loop While blnFound

  GetUniqueHeading = strTempHeading

End Function

Public Function GetSequentialNumberedFile(ByVal strFileName As String) As String

  Dim lngFound As Long
  Dim lngCount As Long

  lngCount = 2
  lngFound = InStrRev(strFileName, ".")
  Do While Dir(Left(strFileName, lngFound - 1) & "(" & CStr(lngCount) & ")" & Mid(strFileName, lngFound)) <> vbNullString
    lngCount = lngCount + 1
  Loop
  GetSequentialNumberedFile = Left(strFileName, lngFound - 1) & "(" & CStr(lngCount) & ")" & Mid(strFileName, lngFound)

End Function

Public Function InsertNumberIntoFilename(ByVal strFileName As String, intNumber As Integer) As String

  Dim lngFound As Long
  Dim lngCount As Long

  lngCount = 2
  lngFound = InStrRev(strFileName, ".")

  InsertNumberIntoFilename = Left(strFileName, lngFound - 1) & "(" & CStr(intNumber) & ")" & Mid(strFileName, lngFound)

End Function


Private Function ReadDataIntoArray() As Boolean

  'Dim pintFileNo As Integer
  Dim pstrExportString As New DataMgr.clsStringBuilder
  'Dim pstrDateString As String
  Dim prstReadyToExport As ADODB.Recordset
  Dim pintLoop As Integer
  Dim lngRecordNumber As Long
  Dim bAppended As Boolean
  Dim bForceHeader As Boolean
  Dim tmpDec As Long
  Dim tmpLen As Long
  Dim objExpr As clsExprExpression
  Dim strTemp As String
  
  Dim lngCol As Long
  Dim lngRow As Long
  Dim lngMaxCols As Long
  Dim lngCount As Long
  Dim lngExportRows As Long
  Dim iStartRow As Integer

  If Len(Dir(mstrOutputFileName)) > 0 Then
    bAppended = (mlngOutputSaveExisting = 3)  'Check if we are appending to file
  Else
    bAppended = False
  End If

  On Error GoTo ReadDataIntoArray_ERROR


  ' Open the export table as a recordset.
  'Set prstReadyToExport = mclsGeneral.GetRecords("SELECT * FROM [" & mstrTempTableName & "]")
  'Set prstReadyToExport = mclsData.OpenRecordset("SELECT * FROM [" & mstrTempTableName & "]", adOpenForwardOnly, adLockReadOnly)
  Set prstReadyToExport = mclsData.OpenTableDirect("[" & mstrTempTableName & "]", adOpenForwardOnly, adLockReadOnly)

  If (prstReadyToExport.BOF And prstReadyToExport.EOF) Then
    mstrErrorString = "No records to export."
    mclsData.ExecuteSql ("IF EXISTS(SELECT * FROM sysobjects WHERE name = '" & mstrTempTableName & "') " & _
                         "DROP TABLE " & mstrTempTableName)
    ReadDataIntoArray = False
    bForceHeader = mbForceHeader
    mbNoRecords = True
  Else
    bForceHeader = False
    mbNoRecords = False
  End If


  'Work out how many columns the array requires...
  lngCount = -1
  lngMaxCols = -1
  For pintLoop = 0 To prstReadyToExport.Fields.Count - 1
    If mvarColDetails(0, pintLoop + 1) = "R" Then   'Carriage Return
      lngCount = 0
    Else
      lngCount = lngCount + 1
    End If
    If lngMaxCols < lngCount Then
      lngMaxCols = lngCount
    End If
  Next
  
  
  
  mblnHeader = False
  ' Omit the header if we are appending to file (if specified)
  If Not (bAppended And mbOmitHeader) Or bForceHeader Then
  
    If mintExportHeader = 2 Or mintExportHeader = 3 Then   'Custom Heading
      mblnHeader = True
      ReDim mstrArrayHeader(lngMaxCols, 0)
      mstrArrayHeader(0, 0) = ReplaceFormatExpressions(mstrExportHeaderText, 1, mlngExportRecordCount, False)
      iStartRow = 1
    End If
  
    ' Column Names
    If mintExportHeader = 1 Or mintExportHeader = 3 Then
      mblnHeader = True
      lngCol = 0
      lngRow = iStartRow
      ReDim Preserve mstrArrayHeader(lngMaxCols, iStartRow) As String
      For pintLoop = 1 To prstReadyToExport.Fields.Count
        If mvarColDetails(0, pintLoop) = "R" Then   'Carriage Return
          lngCol = -1
          lngRow = lngRow + 1
          ReDim Preserve mstrArrayHeader(lngMaxCols, lngRow) As String
        Else
          mstrArrayHeader(lngCol, lngRow) = mvarColDetails(13, pintLoop)
          If mlngOutputFormat = fmtFixedLengthFile Then
            SetLength mstrArrayHeader(lngCol, lngRow), mvarColDetails(6, pintLoop)
          End If
        End If
        lngCol = lngCol + 1
      Next
    End If
        
  End If
  
  ' If we only need to output the header bomb out here
  lngRow = -1
  mblnFooter = False
  If Not mbNoRecords Then
    'Close #pintFileNo
    'Exit Function

    ReDim mstrArrayData(lngMaxCols, 0) As String

    ' Loop through the export table and print stuff to file.
    lngRecordNumber = 0
    Do While Not prstReadyToExport.EOF
      
      ' If user cancels the export, abort
      If (lngRow Mod 100) = 0 Then
        If gobjProgress.Cancelled Then
          mblnUserCancelled = True
          'Close #pintFileNo
          ReadDataIntoArray = False
          Exit Function
        End If
      End If
        
      lngCol = -1
      lngRow = lngRow + 1
      If lngRow > UBound(mstrArrayData, 2) Then ReDim Preserve mstrArrayData(lngMaxCols, lngRow + 1000) As String
      
      If mbLoggingExportSuccess Then
        pstrExportString.TheString = vbNullString
      End If
      
      ' Loop through the fields in each record, adding them and the delimiter to the export string.
      lngRecordNumber = lngRecordNumber + 1
      For pintLoop = 0 To prstReadyToExport.Fields.Count - 1

        lngCol = lngCol + 1
        If mvarColDetails(0, pintLoop + 1) = "N" Then
          'Record Number
          mstrArrayData(lngCol, lngRow) = CStr(lngRecordNumber)
  
        ElseIf mvarColDetails(0, pintLoop + 1) = "R" Then
          'Carriage Return
          lngCol = -1
          lngRow = lngRow + 1
          If lngRow > UBound(mstrArrayData, 2) Then ReDim Preserve mstrArrayData(lngMaxCols, lngRow + 1000) As String

        'MH20040402 Fault 8434
        ElseIf mvarColDetails(0, pintLoop + 1) = "F" Then
          mstrArrayData(lngCol, lngRow) = Space$(mvarColDetails(6, pintLoop + 1))

        ElseIf mvarColDetails(7, pintLoop + 1) = True Then
          mstrArrayData(lngCol, lngRow) = Format(prstReadyToExport.Fields(pintLoop).Value, mstrExportDateFormat)
  
        ElseIf mvarColDetails(12, pintLoop + 1) Then
          tmpLen = mvarColDetails(6, pintLoop + 1)
          tmpDec = mvarColDetails(11, pintLoop + 1)
          'TM19032004 Fault 8052 - DO NOT format Null values to zero. Set to empty string instead.
          If IsNull(prstReadyToExport.Fields(pintLoop).Value) Then
            strTemp = vbNullString
          Else
            strTemp = FormatNumeric(prstReadyToExport.Fields(pintLoop).Value, tmpLen, tmpDec)
          End If
          
          mstrArrayData(lngCol, lngRow) = strTemp
  
        Else
          'MH20030401 Fault 5225
          'mstrArrayData(lngCol, lngRow) = Left(prstReadyToExport.Fields(pintLoop), mvarColDetails(6, pintLoop + 1))
          If IsNull(prstReadyToExport.Fields(pintLoop).Value) Then
            mstrArrayData(lngCol, lngRow) = vbNullString
          Else
            mstrArrayData(lngCol, lngRow) = prstReadyToExport.Fields(pintLoop).Value
          End If
        End If

        
        'NPG20080718 Fault 13276
        Select Case mvarColDetails(16, lngCol + 1)
          Case 1    ' Upper Case
            mstrArrayData(lngCol, lngRow) = UCase(mstrArrayData(lngCol, lngRow))
          Case 2    ' Lower Case
            mstrArrayData(lngCol, lngRow) = LCase(mstrArrayData(lngCol, lngRow))
        End Select
                
        If mblnStripDelimiter Then
          mstrArrayData(lngCol, lngRow) = Replace(mstrArrayData(lngCol, lngRow), mstrExportActualDelimiter, "")
        End If
        
        If lngCol >= 0 Then
          'Add spaces for fixed length or trim is data is too long...
          SetLength mstrArrayData(lngCol, lngRow), mvarColDetails(6, pintLoop + 1)
  
          If mbLoggingExportSuccess Then
            pstrExportString.Append IIf((pstrExportString.Length <> 0), mstrExportDelimiter, vbNullString) & _
              mstrArrayData(lngCol, lngRow)
          End If
        End If

      Next pintLoop
  
      ' Print the stuff to the file.
      'Print #pintFileNo, pstrExportString
      
      
      'JDM - 13/12/01 - Fault 3280 - Log successful records
      If mbLoggingExportSuccess Then
        gobjEventLog.AddDetailEntry pstrExportString.ToString & vbCrLf & "Exported successfully"
      End If
    
      mlSuccessfulRecords = mlSuccessfulRecords + 1
      
      prstReadyToExport.MoveNext
    
    Loop
  
    ReDim Preserve mstrArrayData(lngMaxCols, lngRow)
     
    If mintExportFooter = 2 Then  'Custom Footing
      mblnFooter = True
      ReDim mstrArrayFooter(lngMaxCols, 0)
      mstrArrayFooter(0, 0) = ReplaceFormatExpressions(mstrExportFooterText, 1, mlngExportRecordCount, False)
    End If

    If mintExportFooter = 1 Then  'Column Names
      mblnFooter = True
      lngCol = 0
      lngRow = 0
      ReDim mstrArrayFooter(lngMaxCols, 0) As String
      For pintLoop = 1 To prstReadyToExport.Fields.Count
        If mvarColDetails(0, pintLoop) = "R" Then   'Carriage Return
          lngCol = -1
          lngRow = lngRow + 1
          ReDim Preserve mstrArrayFooter(lngMaxCols, lngRow) As String
        Else
          mstrArrayFooter(lngCol, lngRow) = mvarColDetails(13, pintLoop)
          If mlngOutputFormat = fmtFixedLengthFile Then
            SetLength mstrArrayFooter(lngCol, lngRow), mvarColDetails(6, pintLoop)
          End If
        End If
        lngCol = lngCol + 1
      Next
    End If
      
    ' Close the final output file
    'Close #pintFileNo

  End If

'Debug.Print "end" & Now

  'Drop the export table as we dont need it now.
  Set prstReadyToExport = Nothing
  mclsData.ExecuteSql ("IF EXISTS(SELECT * FROM sysobjects WHERE name = '" & mstrTempTableName & "') " & _
                       "DROP TABLE " & mstrTempTableName)
  ReadDataIntoArray = True
  Exit Function

ReadDataIntoArray_ERROR:

  'Select Case Err.Number
  'Case 70 ' file creation error / sharing violation
  '  mstrErrorString = "Error creating file '" & mstrOutputFilename & "'. File may already be in use." & vbCrLf & "(" & Err.Description & ")"
  'Case 76 ' path not found error
  '  mlngExportRecordCount = 0
  '  mstrErrorString = "Error whilst exporting to delimited file." & vbCrLf & "(" & Err.Description & ")"
  'Case Else
    mstrErrorString = "Error whilst exporting data." & vbCrLf & "(" & Err.Description & ")"
  '  Close #pintFileNo
  'End Select
  
  ReadDataIntoArray = False

End Function

Private Function SendArrayToOutputOptions()

  Dim objOutput As clsOutputRun
  Dim objColumn As clsColumn
  Dim lngLoop As Long
  Dim intBlockLoop As Integer
  Dim intBlockTotal As Integer
  Dim sExtension As String
  Dim sFileName As String
  Dim aryOutputData() As String

  Set objOutput = New clsOutputRun

  ' How many files required for this export?
  If mbSplitFile Then
    intBlockTotal = Floor(UBound(mstrArrayData, 2) / mlngSplitFileSize) + 1
    
    ' If split into blocks ensure that there is some logic for the filename
    If InStr(mstrOutputFileName, "{BLOCKNUMBER}") = 0 Then
      sExtension = "." & Mid(mstrOutputFileName, InStrRev(mstrOutputFileName, ".") + 1, Len(mstrOutputFileName))
      mstrOutputFileName = Replace(mstrOutputFileName, sExtension, "_{BLOCKNUMBER}" + sExtension)
    End If
    
  Else
    intBlockTotal = 1
  End If

  For intBlockLoop = 1 To intBlockTotal

    ' Get a block of exporting data
    If mbSplitFile And mlngSplitFileSize > 0 Then
      aryOutputData = OutputArrayBlock(intBlockLoop)
      
      If mblnHeader Then
        mstrArrayHeader(0, 0) = ReplaceFormatExpressions(mstrExportHeaderText, intBlockLoop, UBound(aryOutputData, 2) + 1, False)
      End If
      
      If mblnFooter Then
        mstrArrayFooter(0, 0) = ReplaceFormatExpressions(mstrExportFooterText, intBlockLoop, UBound(aryOutputData, 2) + 1, False)
      End If
      
    Else
      aryOutputData = mstrArrayData
    End If

    ' Process the file name
    sFileName = ReplaceFormatExpressions(mstrOutputFileName, intBlockLoop, UBound(aryOutputData, 2) + 1, True)
  
    If objOutput.SetOptions _
        (False, mlngOutputFormat, False, _
        False, vbNullString, _
        mblnOutputSave, mlngOutputSaveExisting, _
        mblnOutputEmail, mlngOutputEmailAddr, mstrOutputEmailSubject, _
        mstrOutputEmailAttachAs, sFileName) Then
  
      objOutput.FileDelimiter = mstrExportActualDelimiter
      objOutput.ApplyStyles = False
      objOutput.SizeColumnsIndependently = True
    
      If objOutput.GetFile Then
    
        For lngLoop = 1 To UBound(mvarColDetails, 2)
          objOutput.AddColumn _
            CStr(mvarColDetails(13, lngLoop)), _
            IIf(mvarColDetails(12, lngLoop), sqlNumeric, sqlVarChar), _
            CLng(mvarColDetails(11, lngLoop))
        Next
  
        objOutput.HeaderRows = 0
        objOutput.AddPage mstrExportName, mstrExportName
  
        If mblnHeader Then
          objOutput.DisableDelimiterCheck = True
          objOutput.DisableSeparatorInHeaderAndFooter = (mintExportHeader <> 1)
          objOutput.DataArray mstrArrayHeader
        End If
    
        objOutput.EncloseInQuotes = mblnExportQuotes
        If Not mblnNoRecords Then
          objOutput.DisableDelimiterCheck = False
          objOutput.DisableSeparatorInHeaderAndFooter = (mintExportHeader <> 1)
          objOutput.UpdateProgressPerRow = True
          objOutput.DataArray aryOutputData
          objOutput.UpdateProgressPerRow = False
        End If
        objOutput.EncloseInQuotes = False
  
        If mblnFooter Then
          objOutput.DisableDelimiterCheck = True
          objOutput.DataArray mstrArrayFooter
        End If
  
        If Not gblnBatchMode Then
          gobjProgress.CloseProgress
        End If
        objOutput.Complete
     
      End If
      
    End If
    
  Next intBlockLoop

  objOutput.ClearUp
  mstrErrorString = objOutput.ErrorMessage

  Set objOutput = Nothing
  SendArrayToOutputOptions = (mstrErrorString = vbNullString)

End Function


Private Sub SetLength(ByRef strInput As String, lngLength)
  
  Dim blnFixedLength As Boolean
  
  blnFixedLength = (mlngOutputFormat = fmtFixedLengthFile)
  
  If blnFixedLength Or lngLength > 0 Then
    If blnFixedLength Then
      strInput = strInput & Space$(lngLength)
    End If
    
    ' If they've banged in a very big size ignore it as Left$ can only trim so much...
    If lngLength < 999999999 Then
      strInput = Left$(strInput, lngLength)
    End If
    
  End If

End Sub

Private Function GetFileText(ByVal FileName As String) As String
  Dim Handle As Integer
  Handle = FreeFile
  Open FileName$ For Input As #Handle
  GetFileText = Input$(LOF(Handle), Handle)
  Close #Handle
End Function

' Extracts just the filename from a path
Function GetFileNameOnly(pstrFilePath As String) As String
  Dim astrPath() As String
  astrPath = Split(pstrFilePath, "\")
  GetFileNameOnly = astrPath(UBound(astrPath))
End Function

Private Function ReplaceFormatExpressions(ByVal Value As String, ByVal blockNumber As Long, ByVal blockRecordCount As Long, ByVal isFileName As Boolean) As String

  Dim sDateFormat As String
  sDateFormat = IIf(isFileName, "yyyymmdd", mstrExportDateFormat)

  ' Replace variables
  Value = Replace(Value, "{BLOCKNUMBER}", blockNumber)
  Value = Replace(Value, "{BLOCKCOUNT}", blockRecordCount)
  Value = Replace(Value, "{TOTALCOUNT}", mlngExportRecordCount)
  Value = Replace(Value, "{DATETIME}", Format(mdExportCreateDate, sDateFormat + "hhnnss"))
  Value = Replace(Value, "{DATE}", Format(DateValue(mdExportCreateDate), sDateFormat))

  ReplaceFormatExpressions = Value

End Function

Private Function Floor(ByVal dblValue As Double) As Double
  Dim myDec As Long
  myDec = InStr(1, CStr(dblValue), ".", vbTextCompare)
  If myDec > 0 Then
    Floor = CDbl(Left(CStr(dblValue), myDec))
  Else
    Floor = dblValue
  End If
End Function

Private Function OutputArrayBlock(ByVal intBlockNumber As Integer) As String()

  On Error GoTo ArrayExceeded:
   
  Dim iStartBlock As Integer
  Dim iRowsInBlock As Integer
  Dim lngTotalRows As Long
  Dim iColumnLoop As Integer
  Dim J As Integer
        
  lngTotalRows = UBound(mstrArrayData, 2)
  iStartBlock = (intBlockNumber - 1) * mlngSplitFileSize
  iRowsInBlock = Minimum(mlngSplitFileSize, CLng(lngTotalRows - iStartBlock) + 1)
  ReDim aryOutput(UBound(mstrArrayData, 1), iRowsInBlock - 1) As String
   
  For iColumnLoop = 0 To UBound(mstrArrayData, 1)
    For J = 0 To iRowsInBlock - 1
      aryOutput(iColumnLoop, J) = mstrArrayData(iColumnLoop, J + iStartBlock)
    Next J
  Next iColumnLoop

ArrayExceeded:
  OutputArrayBlock = aryOutput

End Function
