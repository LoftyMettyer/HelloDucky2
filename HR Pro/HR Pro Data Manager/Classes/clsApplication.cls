VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'For batch job checking
Private datData As HRProDataMgr.clsDataAccess

' Screen Types
Enum ScreenType
  screenParentTable = 1
  screenParentView = 2
  screenHistoryTable = 4
  screenHistoryView = 8
  screenLookup = 16
  screenFind = 32
  screenHistorySummary = 64
  screenQuickEntry = 128
  screenPickList = 256
End Enum

'DefSel Screen Enum
Public Enum DefSelScreen
    screenNone = 0
    screenDataTransfer = 1
    screenGlobalAdd
    screenGlobalUpdate
    screenGlobalDelete
End Enum

Public Enum GlobalType
    glAdd = 1
    glUpdate = 2
    glDelete = 3
End Enum

'Edit option constants
Public Enum EditOptions
  edtCancel = 0
  edtAdd = 2 ^ 10
  edtDelete = 2 ^ 11
  edtEdit = 2 ^ 12
  edtCopy = 2 ^ 13
  edtSelect = 2 ^ 14
  edtDeselect = 2 ^ 15
  edtPrint = 2 ^ 16
  edtProperties = 2 ^ 17
  edtRefresh = 2 ^ 18
End Enum

'Standard report constants
Public Enum ReportOptions
  rptCancel = 0
  rptOK = 1
  rptRun = 2
End Enum

'Table type constants
Public Enum TableTypes
  tabTopLevel = 1
  tabChild = 2
  tabLookup = 3
End Enum

'Control type constants
Public Enum ControlTypes
  ctlCheck = 1
  ctlCombo = 2
  ctlImage = 4
  ctlOle = 8
  ctlRadio = 16
  ctlSpin = 32
  ctlText = 64
  ctlTab = 128
  ctlLabel = 256
  ctlFrame = 512
  ctlPhoto = 1024
  ctlCommand = 2048
  ctlWorkingPattern = 4096
  ctlLine = 2 ^ 13
  ctlBrowser = 2 ^ 14
End Enum
  
'Column type constants
Public Enum ColumnTypes
  ColData = 0
  colLookup = 1
  colCalc = 2
  colSystem = 3
  colLink = 4
'  colWorkingPattern = 5
End Enum

'SQL DatType
Public Enum SQLDataType
  sqlUnknown = 0      ' ?
  sqlOle = -4         ' OLE columns
  sqlBoolean = -7     ' Logic columns
  sqlNumeric = 2      ' Numeric columns
  sqlInteger = 4      ' Integer columns
  sqlDate = 11        ' Date columns
  sqlVarChar = 12     ' Character columns
  sqlVarBinary = -3   ' Photo columns
  sqlLongVarChar = -1 ' Working Pattern columns
End Enum

'Case Conversion Types
Public Enum CaseConvert
    convNone = 0
    convUpper = 1
    convLower = 2
    convProper = 3
End Enum
  
'Instantiate internal classes
Private ODBC As New ODBC

'Local variables to hold property values
Private mvarLoggedIn As Boolean
Private mvarlCurrentFormId As Long  ' A unique number that stores the formid to give out to
                                   ' the next form

Public Enum test
  globfuncvaltyp_STRAIGHTVALUE = 1
  globfuncvaltyp_LOOKUPTABLE = 2
  globfuncvaltyp_FIELD = 3
  globfuncvaltyp_CALCULATION = 4
End Enum

Public Enum FilterOperators
  giFILTEROP_UNDEFINED = 0
  giFILTEROP_EQUALS = 1
  giFILTEROP_NOTEQUALTO = 2
  giFILTEROP_ISATMOST = 3
  giFILTEROP_ISATLEAST = 4
  giFILTEROP_ISMORETHAN = 5
  giFILTEROP_ISLESSTHAN = 6
  giFILTEROP_ON = 7
  giFILTEROP_NOTON = 8
  giFILTEROP_AFTER = 9
  giFILTEROP_BEFORE = 10
  giFILTEROP_ONORAFTER = 11
  giFILTEROP_ONORBEFORE = 12
  giFILTEROP_CONTAINS = 13
  giFILTEROP_IS = 14
  giFILTEROP_DOESNOTCONTAIN = 15
  giFILTEROP_ISNOT = 16
End Enum

' Record profile orientation constants
Public Enum OrientationTypes
  giHORIZONTAL = 0
  giVERTICAL = 1
End Enum


Public Property Get NextFormID() As Long
  ' Increment the formid
  mvarlCurrentFormId = mvarlCurrentFormId + 1
  NextFormID = mvarlCurrentFormId
End Property

Public Property Get LoggedIn() As Boolean
  LoggedIn = mvarLoggedIn
End Property

Public Property Let LoggedIn(ByVal vData As Boolean)
  mvarLoggedIn = vData
End Property

Public Property Get Name() As String
  Name = App.ProductName
End Property

Public Function Activate() As Boolean
  On Error GoTo ErrorTrap
  
  Dim blnAccessToDiary As Boolean
    
  If Not LoggedIn Then
    Login
  End If
  
  Dim frmStyle As New frmHiddenStyle
  Load frmStyle
  
  ' Setup the collection of Tables.
  SetupTablesCollection

  ' Get the module specifics.
  InitialiseModules

  'If LoggedIn Then
  If LoggedIn And Not gblnBatchJobsOnly Then
    
    CheckOvernightJob
    
    'Unload the styling form BEFORE the Main MDI is loaded
    'otherwise Unloading the style form will remove the styles
    'from the MDI too...?!?!?
    Unload frmStyle
    
    frmMain.SetBackground (True)
    frmMain.Show
    Activate = True
  End If

  'Does the user have permission to run batch jobs ?
  If datGeneral.SystemPermission("BATCHJOBS", "RUN") = True Then
    frmSplash.Visible = False
    CheckScheduledBatchJobs
  End If

  ' Are any workflow steps pending for the current user to action?
  If gbWorkflowEnabled Then
    frmSplash.Visible = False
    CheckWorkflowOutOfOffice
    CheckPendingWorkflowSteps False
  End If
  
  blnAccessToDiary = datGeneral.SystemPermission("DIARY", "MANUALEVENTS")
  
  With frmMain.abMain
    .Tools("Diary").Enabled = blnAccessToDiary
    .Tools("DiaryRebuild").Enabled = datGeneral.SystemPermission("DIARY", "REBUILDPURGE")
    .Tools("DiaryDelete").Enabled = datGeneral.SystemPermission("DIARY", "REBUILDPURGE")
    '.Tools("MailMerge").Enabled = datGeneral.SystemPermission("MAILMERGE", "VIEW")
  
    .Tools("EmailQueue").Enabled = datGeneral.SystemPermission("EMAIL", "VIEW")
    
    'MH20040825
    .Tools("OutlookQueue").Enabled = datGeneral.SystemPermission("OUTLOOKQUEUE", "VIEW")
  
  End With

  If gblnDiaryStartUpCheck = True And blnAccessToDiary Then
    'gobjDiary.GetPassedAlarms
    '1=Past Events,3=List View
    gobjDiary.ShowAlarmedEvents 1, 3
    Set gobjDiary = Nothing
  End If

  'Diary and mail merge stuff
  frmMain.tmrDiary.Enabled = (gblnDiaryConstCheck And blnAccessToDiary)
  
'  If gbAllowOutput_Word Then
'    'WdSaveFormat.wdFormatDocumentDefault
'    'WdSaveFormat.wdFormatDocument97
'    giOfficeSaveVersion_Word = GetSystemSetting("output", "save version word", WdSaveFormat.wdFormatDocument)
'    gsOfficeFileFilter_Word = GetSystemSetting("output", "file filter word", "Word Document (*.doc)|*.doc")
'    gsOfficeTemplateFilter_Word = GetSystemSetting("output", "template filter word", "Word Template (*.doc;*.dot)|*.doc;*.dot")
'  End If
'
'  If gbAllowOutput_Excel Then
'    'XlFileFormat.xlWorkbookDefault
'    'XlFileFormat.xlExcel8
'    giOfficeSaveVersion_Excel = GetSystemSetting("output", "save version excel", 56)  'XlFileFormat.xlExcel8)
'    gsOfficeFileFilter_Excel = GetSystemSetting("output", "file filter excel", "Excel Workbook (*.xls)|*.xls")
'    gsOfficeTemplateFilter_Excel = GetSystemSetting("output", "template filter excel", "Excel Template (*.xls;*.xlt)|*.xls;*.xlt")
'  End If
  
  'Unload the styling form BEFORE the Main MDI is loaded
  'otherwise Unloading the style form will remove the styles
  'from the MDI too...?!?!?
  If IsFormLoaded("frmHiddenStyle") Then
    Unload frmStyle
  End If
  
  Exit Function

ErrorTrap:
  If ASRDEVELOPMENT Then
    MsgBox Err.Description, vbCritical, "ASRDEVELOPMENT"
    Stop
  End If
  Activate = False
  Err = False
  
End Function

Public Function IsFormLoaded(ByVal FormName As String) As Boolean
    Dim frm As Form


    For Each frm In Forms
        If LCase(frm.Name) = LCase(FormName) Then IsFormLoaded = True
    Next frm
End Function

Public Function Login() As Boolean
  On Error GoTo ErrorTrap
  
  If LoggedIn Then
    Logout
  End If
  
  If Not LoggedIn Then
    
    Load frmLogin
    LoggedIn = frmLogin.OK

    If Not LoggedIn And Not gblnBatchJobsOnly Then
      frmLogin.Show vbModal
      LoggedIn = frmLogin.OK
      
      If Not LoggedIn Then
        Logout
      End If
    End If

    Unload frmLogin
    Set frmLogin = Nothing

  End If
  
  Login = LoggedIn
  If Login Then Call AuditAccess(iLOGIN, "Data")
  Exit Function
  
ErrorTrap:
  Login = False
  Err = False
  
End Function

Public Function Logout() As Boolean
  On Error GoTo ErrorTrap
  
  If LoggedIn Then
'    Database.ClearConnection
    
    LoggedIn = False
  End If
  
  Logout = (Not LoggedIn)
  
  Exit Function
  
ErrorTrap:
  Logout = False
  Err = False
  
End Function

Private Sub Class_Initialize()
  mvarLoggedIn = False
  mvarlCurrentFormId = 0
End Sub

Private Sub Class_Terminate()
  'Release internal classes
  Set ODBC = Nothing
End Sub

'Public Function GetBatchLock() As Integer
'
'  Dim iStatus As Integer
'  Dim cmADO As ADODB.Command
'  Dim pmADO As ADODB.Parameter
'
'  Set cmADO = New ADODB.Command
'  With cmADO
'    .CommandText = "sp_ASRGetBatchLock"
'    .CommandType = adCmdStoredProc
'    Set .ActiveConnection = gADOCon
'
'    Set pmADO = .CreateParameter("Ret", adInteger, adParamOutput)
'    .Parameters.Append pmADO
'
'    Set pmADO = Nothing
'
'    cmADO.Execute
'
'    iStatus = .Parameters(0).Value
'  End With
'
'  Set cmADO = Nothing
'
'  GetBatchLock = iStatus
'
'End Function
'
'

Public Sub CheckScheduledBatchJobs()

  'Dim pobjBatchJobRUN As clsBatchJobRUN
  Dim frmPrompt As frmDefSel
  Dim strEventLogIDs As String
  Dim fExit As Boolean
  Dim dblTimeOut As Double

  Set frmPrompt = New frmDefSel
  Screen.MousePointer = vbDefault
  
  Do

    frmPrompt.Options = edtProperties + edtSelect
    frmPrompt.EnableRun = True
    frmPrompt.BatchPrompt = True
    frmPrompt.EventLogIDs = strEventLogIDs
    frmPrompt.ShowList utlBatchJob, , True
    fExit = True
    If frmPrompt.ListCount > 0 Then
      If gblnBatchJobsOnly Then
        frmPrompt.RunSelectedJobs
        fExit = True
      Else
        frmPrompt.Show vbModal
        strEventLogIDs = frmPrompt.EventLogIDs
        fExit = (frmPrompt.Action = edtCancel)
      End If
    End If

  Loop While Not fExit

  Unload frmPrompt
  Set frmPrompt = Nothing

  'MH20010515
  If gblnBatchJobsOnly Then

    dblTimeOut = Timer + 5
    If dblTimeOut > 86395 Then  'Just in case it gets stuck in a loop at midnight
      dblTimeOut = 0
    End If

    'Delay to prevent RunTime 440 and finish sending any emails, if required
    Do
      DoEvents
    Loop While objEmail.StillExecuting And dblTimeOut > Timer

    End
  End If

End Sub


Private Sub CheckOvernightJob()

  Dim datData As clsDataAccess
  Dim rsTemp As Recordset
  Dim strSQL As String

  On Local Error GoTo LocalErr

  If GetSystemSetting("Overnight", "Warning", "1") <> "1" Then
    Exit Sub
  End If


  Set datData = New clsDataAccess

  strSQL = "SELECT datediff(hh,convert(datetime,SettingValue,103),getdate()) FROM ASRSysSystemSettings " & _
           " WHERE Section = 'overnight'" & _
           " AND SettingKey = 'last completed'"
  Set rsTemp = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)

  With rsTemp
    If Not .BOF And Not .EOF Then
      If rsTemp.Fields(0).Value > 24 Then
        MsgBox "The overnight job has not completed within the last 24 hours.  This may affect date dependant areas of HR Pro." & vbCrLf & "Please inform your SQL administrator.", vbExclamation, "HR Pro Overnight Job"
      End If
    End If
  End With

  rsTemp.Close

  Set rsTemp = Nothing
  Set datData = Nothing

LocalErr:

End Sub
