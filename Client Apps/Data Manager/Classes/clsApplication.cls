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

'Instantiate internal classes
Private ODBC As New ODBC

'Local variables to hold property values
Private mvarLoggedIn As Boolean
Private mvarlCurrentFormId As Long  ' A unique number that stores the formid to give out to
                                   ' the next form


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
  Name = app.ProductName
End Property

Public Function Activate() As Boolean
  On Error GoTo ErrorTrap
  
  Dim blnAccessToDiary As Boolean
  
  
  DebugOutput "Application.Activate", "Check LoggedIn"
  
  If Not LoggedIn Then
    Login
  End If
  
  DebugOutput "Application.Activate", "Load frmStyle"
  
  Dim frmStyle As New frmHiddenStyle
  Load frmStyle
  
  DebugOutput "Application.Activate", "SetupTablesCollection"
  
  ' Setup the collection of Tables.
  SetupTablesCollection

  DebugOutput "Application.Activate", "InitialiseModules"
  
  ' Get the module specifics.
  InitialiseModules

  DebugOutput "Application.Activate", "CheckOvernightJob"
  
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

  DebugOutput "Application.Activate", "CheckScheduledBatchJobs"
  
  'Does the user have permission to run batch jobs ?
  If datGeneral.SystemPermission("BATCHJOBS", "RUN") = True Then
    frmSplash.Visible = False
    CheckScheduledBatchJobs
  End If

  DebugOutput "Application.Activate", "CheckScheduledReportPacks"
  
  'Does the user have permission to run report packs ?
  If datGeneral.SystemPermission("REPORTPACKS", "RUN") = True Then
    gblnReportPackMode = True
      frmSplash.Visible = False
      CheckScheduledReportPacks
    gblnReportPackMode = False
  End If

  If gblnBatchJobsOnly Then
    End
  End If

  ' Are any workflow steps pending for the current user to action?
  If gbWorkflowEnabled Then
    frmSplash.Visible = False
    DebugOutput "Application.Activate", "CheckWorkflowOutOfOffice"
    CheckWorkflowOutOfOffice
    DebugOutput "Application.Activate", "CheckPendingWorkflowSteps"
    CheckPendingWorkflowSteps False
  End If
  
  DebugOutput "Application.Activate", "CheckDiary"
  
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
  
  DebugOutput "Application.Activate", "Unload frmStyle"
  
  'Unload the styling form BEFORE the Main MDI is loaded
  'otherwise Unloading the style form will remove the styles
  'from the MDI too...?!?!?
  If IsFormLoaded("frmHiddenStyle") Then
    Unload frmStyle
  End If
  
  Exit Function

ErrorTrap:
  If ASRDEVELOPMENT Then
    COAMsgBox Err.Description, vbCritical, "ASRDEVELOPMENT"
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
  
  DebugOutput "Application.Login", "LoggedIn"
  
  If LoggedIn Then
    Logout
  End If
  
  DebugOutput "Application.Login", "Load frmLogin"
  
  If Not LoggedIn Then
    
    Load frmLogin
    LoggedIn = frmLogin.OK

    DebugOutput "Application.Login", "Show frmLogin"
    
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
  
  DebugOutput "Application.Login", "Audit Login"
  
  Login = LoggedIn
  If Login Then Call AuditAccess(iLOGIN, "Data")
  Exit Function
  
ErrorTrap:
  DebugOutput "Application.Login", "Error"

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

  End If

End Sub


Public Sub CheckScheduledReportPacks()
  Dim frmPrompt As frmDefSel
  Dim strEventLogIDs As String
  Dim fExit As Boolean
  Dim dblTimeOut As Double

  Set frmPrompt = New frmDefSel
  'gblnReportPackMode = True
  Screen.MousePointer = vbDefault
  
  Do

    frmPrompt.Options = edtProperties + edtSelect
    frmPrompt.EnableRun = True
    frmPrompt.BatchPrompt = True
    frmPrompt.EventLogIDs = strEventLogIDs
    frmPrompt.ShowList utlReportPack, , True
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
        COAMsgBox "The overnight job has not completed within the last 24 hours.  This may affect date dependant areas of OpenHR." & vbCrLf & "Please inform your SQL administrator.", vbExclamation, "OpenHR Overnight Job"
      End If
    End If
  End With

  rsTemp.Close

  Set rsTemp = Nothing
  Set datData = Nothing

LocalErr:

End Sub
