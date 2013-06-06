VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBatchJobRUN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Data access class
Private mclsData As clsDataAccess

Private mstrEventLogIDs As String
Private mblnJobCancelled As Boolean
Private mblnBatchCancelled As Boolean
Private fOK As Boolean

'New Default Output Variables
Private mlngOutputFormat As Long
Private mblnOutputScreen As Boolean
Private mblnOutputPrinter As Boolean
Private mstrOutputPrinterName As String
Private mblnOutputSave As Boolean
Private mlngOutputSaveExisting As Long
Private mblnOutputEmail As Boolean
Private mlngOutputEmailAddr As Long
Private mstrOutputEmailSubject As String
Private mstrOutputEmailAttachAs As String
Private mstrOutputFileName As String
Private mblnPreviewOnScreen As Boolean
Private mblnPrintFilterHeader As Boolean

Private Sub Class_Initialize()
  Set mclsData = New clsDataAccess
End Sub

Private Sub Class_Terminate()
  Set mclsData = Nothing
End Sub

Public Function DoesUserHavePermissionForAllJobs(plngID As Long) As Boolean

  On Error GoTo DoesUserHavePermissionForAllJobs_ERROR
  
  Dim prstBatchDetails As Recordset
  Dim pstrCatKey As String
    
  'Retrieve the individual job information for the specified batch job id
  Set prstBatchDetails = mclsData.OpenRecordset("SELECT * FROM AsrSysBatchJobDetails WHERE BatchJobNameID = " & plngID & " ORDER BY JobOrder", adOpenKeyset, adLockReadOnly)
  
  'Loop thu recordset, setting the Category Key so the system permission can be checked
  With prstBatchDetails
  
    Do Until .EOF
    
      Select Case .Fields("JobType")
        Case "Career Progression"
          pstrCatKey = "CAREER"
        Case "Calendar Report"
          pstrCatKey = "CALENDARREPORTS"
        Case "Cross Tab"
          pstrCatKey = "CROSSTABS"
        Case "Custom Report"
          pstrCatKey = "CUSTOMREPORTS"
        Case "Data Transfer"
          pstrCatKey = "DATATRANSFER"
        Case "Export"
          pstrCatKey = "EXPORT"
        Case "Global Add"
          pstrCatKey = "GLOBALADD"
        Case "Global Delete"
          pstrCatKey = "GLOBALDELETE"
        Case "Global Update"
          pstrCatKey = "GLOBALUPDATE"
        Case "Import"
          pstrCatKey = "IMPORT"
        Case "Mail Merge"
          pstrCatKey = "MAILMERGE"
        Case "Match Report"
          pstrCatKey = "MATCHREPORTS"
        Case "Record Profile"
          pstrCatKey = "RECORDPROFILE"
        Case "Succession Planning"
          pstrCatKey = "SUCCESSION"
        Case Else
          pstrCatKey = ""
      End Select
    
      'Can the user run the individual job ? If not, exit
      If pstrCatKey <> "" Then
        If datGeneral.SystemPermission(pstrCatKey, "RUN") = False Then
          DoesUserHavePermissionForAllJobs = False
          Set prstBatchDetails = Nothing
          Exit Function
        End If
      End If
    
      .MoveNext
    
    Loop
  
  End With
  
  'User does have permission to run all jobs contained in the batch
  DoesUserHavePermissionForAllJobs = True
  Set prstBatchDetails = Nothing
  Exit Function
  
DoesUserHavePermissionForAllJobs_ERROR:
  
  Set prstBatchDetails = Nothing
  DoesUserHavePermissionForAllJobs = False
  COAMsgBox "Error while checking individual job permissions." & vbCrLf & "Contact Asr Support stating : " & Err.Description, vbCritical + vbOKOnly, IIf(gblnReportPackMode, "Report Pack", "Batch Job")
  
End Function

Public Function RunBatchJob(lBatchJobID As Long, sBatchJobName As String, lEventLogID As Long) As String
  Dim rsBatchDetails As Recordset
  Dim rsBatchOutputOptions As Recordset
  Dim strNotes As String
  Dim strOutputOptions As String
  
'New Default Output Variables
Dim lngOutputFormat As Long
Dim blnOutputScreen As Boolean
Dim blnOutputPrinter As Boolean
Dim strOutputPrinterName As String
Dim blnOutputSave As Boolean
Dim lngOutputSaveExisting As Long
Dim blnOutputEmail As Boolean
Dim lngOutputEmailAddr As Long
Dim strOutputEmailSubject As String
Dim strOutputEmailAttachAs As String
Dim strOutputFilename As String
Dim blnPreviewOnScreen As Boolean
Dim blnPrintFilterHeader As Boolean
Dim strOutputTitlePage As String
Dim strOutputReportPackTitle As String
Dim strOutputOverrideFilter As String
Dim blnOutputTOC As Boolean
Dim blnOutputCoverSheet As Boolean
  
  gblnBatchMode = True

  'Retrive the individual job information for the specified batch job id
  Set rsBatchDetails = mclsData.OpenPersistentRecordset("SELECT * FROM AsrSysBatchJobDetails WHERE BatchJobNameID = " & lBatchJobID & " ORDER BY JobOrder", adOpenKeyset, adLockReadOnly)
  
  If gblnReportPackMode Then
    'GetBatchJobDefaultOutputParams
    strOutputOptions = "outputPreview, outputFormat, outputScreen, outputPrinter,OutputPrinterName, outputSave, outputSaveExisting, outputEmail, outputEmailAddr, outputEmailSubject, outputFilename, outputEmailAttachAs, outputTitlePage, outputReportPackTitle, outputOverrideFilter, outputTOC, outputCoverSheet, OutputTitlePage, OutputReportPackTitle, OutputOverrideFilter, OutputTOC, OutputCoverSheet"
    Set rsBatchOutputOptions = mclsData.OpenPersistentRecordset("SELECT " & strOutputOptions & " FROM ASRSysBatchJobName WHERE ID = " & lBatchJobID, adOpenKeyset, adLockReadOnly)
      
    With rsBatchOutputOptions
      rsBatchOutputOptions.MoveLast
      rsBatchOutputOptions.MoveFirst
      'Populate local vars
      blnPreviewOnScreen = False '!OutputPreview 'its always going to be false so why read it in
      blnOutputScreen = False '!OutputScreen 'its always going to be false so why read it in
      lngOutputFormat = !OutputFormat
      blnOutputPrinter = !OutputPrinter
      strOutputPrinterName = !OutputPrinterName
      
      blnOutputSave = !OutputSave
      lngOutputSaveExisting = !OutputSaveExisting
      strOutputFilename = !OutputFilename
      
      blnOutputEmail = !OutputEmail
      lngOutputEmailAddr = !OutputEmailAddr
      strOutputEmailSubject = !OutputEmailSubject
      strOutputEmailAttachAs = IIf(IsNull(!OutputEmailAttachAs), vbNullString, !OutputEmailAttachAs)
      
      strOutputTitlePage = !OutputTitlePage
      strOutputReportPackTitle = !OutputReportPackTitle
      strOutputOverrideFilter = !OutputOverrideFilter
      blnOutputTOC = !OutputTOC
      blnOutputCoverSheet = !OutputCoverSheet
    End With
End If

  'This should never happen, but just in case...
  If rsBatchDetails.BOF And rsBatchDetails.EOF Then
    COAMsgBox IIf(gblnReportPackMode, "report pack", "batch job") & " has been deleted by another user.", vbExclamation + vbOKOnly, IIf(gblnReportPackMode, "Report Pack", "Batch Job")
    RunBatchJob = IIf(gblnReportPackMode, "Report pack", "Batch job") & " Failed : " & IIf(gblnReportPackMode, "Report pack", "Batch job") & " has been deleted by another user."
    Exit Function
  End If
    
  'Get record counts and initialise progress bar
  rsBatchDetails.MoveLast
  rsBatchDetails.MoveFirst
 
    With gobjProgress
      .Bar1MaxValue = rsBatchDetails.RecordCount
      .Caption = IIf(gblnReportPackMode, "Report Pack", "Batch Job") & " : " & sBatchJobName
      .AVI = dbTransfer
      .Time = False
      .Cancel = True
      .NumberOfBars = 2
      .OpenProgress
    End With

  gobjEventLog.GetBatchRunID
  gobjEventLog.BatchJobID = lBatchJobID
  gobjEventLog.BatchJobName = sBatchJobName
  
  mstrEventLogIDs = vbNullString
    
  'Loop thru the individual jobs, running them in turn
  fOK = True
  Do Until rsBatchDetails.EOF
    'Update the job pintCounter of pintCounter progress bar caption
    gobjProgress.Bar1Caption = "Job " & rsBatchDetails.AbsolutePosition & " Of " & rsBatchDetails.RecordCount
    
    Select Case rsBatchDetails.Fields("JobType")
      
      Case "-- Pause --"
        'NHRD29102004 Fault 8240 added gblnBatchJobsOnly variable
        If fOK And (Not mblnBatchCancelled) And (Not gblnBatchJobsOnly) Then
          ' No previous jobs in the batch have been cancelled so run this
          gobjProgress.AVI = dbTransfer
          gobjProgress.ResetBar2
          gobjProgress.Bar2Caption = "Pause Command Encountered"
          gobjProgress.Visible = False
          COAMsgBox rsBatchDetails.Fields("Parameter"), vbInformation + vbOKOnly, "Pause Command Encountered"
          gobjProgress.Visible = True
          fOK = True
        End If
      
      Case "Absence Breakdown"
        Dim pobjAbsenceBreakdown As frmCrossTabRun
        
        If fOK And Not mblnBatchCancelled Then
          If datGeneral.SystemPermission("STANDARDREPORTS", "RUN_AB") Then
            If gfAbsenceEnabled Then
              Set pobjAbsenceBreakdown = New frmCrossTabRun
              fOK = pobjAbsenceBreakdown.AbsenceBreakdownExecuteReport(0)
              strNotes = pobjAbsenceBreakdown.ErrorString
              mblnJobCancelled = pobjAbsenceBreakdown.UserCancelled
            Else
              fOK = False
              strNotes = "The system is not configured to run the Absence Breakdown Report"
              gobjEventLog.AddHeader eltStandardReport, "Absence Breakdown"
              gobjEventLog.ChangeHeaderStatus elsFailed
              gobjEventLog.AddDetailEntry strNotes
            End If
          Else
            fOK = False
            strNotes = "You do not have permission to run the Absence Breakdown Report"
            gobjEventLog.AddHeader eltStandardReport, "Absence Breakdown"
            gobjEventLog.ChangeHeaderStatus elsFailed
            gobjEventLog.AddDetailEntry strNotes
          End If
        Else
          ' A job previous in the batch has been cancelled,
          ' so populate event log with skipped
          gobjEventLog.AddHeader eltStandardReport, "Absence Breakdown Report"
          gobjEventLog.ChangeHeaderStatus elsSkipped
        End If
        Set pobjAbsenceBreakdown = Nothing


      Case "Bradford Factor"
        Dim pobjBradfordFactor As clsCustomReportsRUN
        
        If fOK And Not mblnBatchCancelled Then
          If datGeneral.SystemPermission("STANDARDREPORTS", "RUN_BF") Then
            If gfAbsenceEnabled Then
              Set pobjBradfordFactor = New clsCustomReportsRUN
              fOK = pobjBradfordFactor.RunBradfordReport("")
              strNotes = pobjBradfordFactor.ErrorString
              mblnJobCancelled = pobjBradfordFactor.UserCancelled
            Else
              fOK = False
              strNotes = "The system is not configured to run the Bradford Factor Report"
              gobjEventLog.AddHeader eltStandardReport, "Bradford Factor Report"
              gobjEventLog.ChangeHeaderStatus elsFailed
              gobjEventLog.AddDetailEntry strNotes
            End If
          Else
            fOK = False
            strNotes = "You do not have permission to run the Bradford Factor Report"
            gobjEventLog.AddHeader eltStandardReport, "Bradford Factor Report"
            gobjEventLog.ChangeHeaderStatus elsFailed
            gobjEventLog.AddDetailEntry strNotes
          End If
        Else
          ' A job previous in the batch has been cancelled,
          ' so populate event log with skipped
          gobjEventLog.AddHeader eltStandardReport, "Bradford Factor Report"
          gobjEventLog.ChangeHeaderStatus elsSkipped
        End If
        Set pobjBradfordFactor = Nothing

      Case "Calendar Report"
      
        Dim pobjCalendarReportRUN As clsCalendarReportsRUN
        Set pobjCalendarReportRUN = New clsCalendarReportsRUN
        pobjCalendarReportRUN.CalendarReportID = rsBatchDetails.Fields("JobID")
        If fOK And Not mblnBatchCancelled Then
          ' No previous jobs in the batch have been cancelled so run this
          If datGeneral.SystemPermission("CALENDARREPORTS", "RUN") Then
            fOK = pobjCalendarReportRUN.RunCalendarReport("")
            strNotes = pobjCalendarReportRUN.ErrorString
            mblnJobCancelled = pobjCalendarReportRUN.UserCancelled
          Else
            fOK = False
            strNotes = "You do not have permission to run calendar report definitions"
            gobjEventLog.AddHeader eltCalandarReport, GetUtilName("Calendar Report", rsBatchDetails.Fields("JobID"))
            gobjEventLog.ChangeHeaderStatus elsFailed
            gobjEventLog.AddDetailEntry strNotes
          End If
        Else
          ' A job previous in the batch has been cancelled,
          ' so populate event log with skipped
          gobjEventLog.AddHeader eltCalandarReport, GetUtilName("Calendar Report", rsBatchDetails.Fields("JobID"))
          gobjEventLog.ChangeHeaderStatus elsSkipped
        End If
        Set pobjCalendarReportRUN = Nothing
      
      Case "Cross Tab"
      
        Dim pfrmCrossTabRUN As New frmCrossTabRun
        Set pfrmCrossTabRUN = New frmCrossTabRun
        If fOK And Not mblnBatchCancelled Then
          ' No previous jobs in the batch have been cancelled so run this
          'Check permissions...
          If datGeneral.SystemPermission("CROSSTABS", "RUN") Then
            fOK = pfrmCrossTabRUN.ExecuteCrossTab(rsBatchDetails.Fields("JobID"))
            strNotes = pfrmCrossTabRUN.ErrorString
            mblnJobCancelled = pfrmCrossTabRUN.UserCancelled
          Else
            fOK = False
            strNotes = "You do not have permission to run cross tab definitions"
            gobjEventLog.AddHeader eltCrossTab, GetUtilName("Cross Tab", rsBatchDetails.Fields("JobID"))
            gobjEventLog.ChangeHeaderStatus elsFailed
            gobjEventLog.AddDetailEntry strNotes
          End If
        Else
          ' A job previous in the batch has been cancelled,
          ' so populate event log with skipped
          gobjEventLog.AddHeader eltCrossTab, GetUtilName("Cross Tab", rsBatchDetails.Fields("JobID"))
          gobjEventLog.ChangeHeaderStatus elsSkipped
        End If
        Set pfrmCrossTabRUN = Nothing
        
      Case "Custom Report"
      
        Dim pobjReportRUN As New clsCustomReportsRUN
        Set pobjReportRUN = New clsCustomReportsRUN
        pobjReportRUN.CustomReportID = rsBatchDetails.Fields("JobID")
        
        If gblnReportPackMode Then
          pobjReportRUN.SetOutputParameters _
            lngOutputFormat, _
            blnOutputScreen, _
            blnOutputPrinter, _
            strOutputPrinterName, _
            blnOutputSave, _
            lngOutputSaveExisting, _
            blnOutputEmail, _
            lngOutputEmailAddr, _
            strOutputEmailSubject, _
            strOutputEmailAttachAs, _
            strOutputFilename, _
            blnPreviewOnScreen, _
            blnPrintFilterHeader, _
            strOutputTitlePage, _
            strOutputReportPackTitle, _
            strOutputOverrideFilter, _
            blnOutputTOC, _
            blnOutputCoverSheet
        End If
          
        If fOK And Not mblnBatchCancelled Then
          ' No previous jobs in the batch have been cancelled so run this
          'Check permissions...
          If datGeneral.SystemPermission("CUSTOMREPORTS", "RUN") Then
            fOK = pobjReportRUN.RunCustomReport("")
            strNotes = pobjReportRUN.ErrorString
            mblnJobCancelled = pobjReportRUN.UserCancelled
          Else
            fOK = False
            strNotes = "You do not have permission to run custom report definitions"
            gobjEventLog.AddHeader eltCustomReport, GetUtilName("Custom Report", rsBatchDetails.Fields("JobID"))
            gobjEventLog.ChangeHeaderStatus elsFailed
            gobjEventLog.AddDetailEntry strNotes
          End If
        Else
          ' A job previous in the batch has been cancelled,
          ' so populate event log with skipped
          gobjEventLog.AddHeader eltCustomReport, GetUtilName("Custom Report", rsBatchDetails.Fields("JobID"))
          gobjEventLog.ChangeHeaderStatus elsSkipped
        End If
        Set pobjReportRUN = Nothing
      
      Case "Record Profile"
      
        Dim pobjRecordProfileRUN As New clsRecordProfileRUN
        Set pobjRecordProfileRUN = New clsRecordProfileRUN
        pobjRecordProfileRUN.RecordProfileID = rsBatchDetails.Fields("JobID")
        If fOK And Not mblnBatchCancelled Then
          ' No previous jobs in the batch have been cancelled so run this
          
          'MH20011107 Check permissions...
          If datGeneral.SystemPermission("RECORDPROFILE", "RUN") Then
            fOK = pobjRecordProfileRUN.RunRecordProfile
            strNotes = pobjRecordProfileRUN.ErrorString
            mblnJobCancelled = pobjRecordProfileRUN.UserCancelled
          Else
            fOK = False
            strNotes = "You do not have permission to run record profile definitions"
            gobjEventLog.AddHeader eltRecordProfile, GetUtilName("Record Profile", rsBatchDetails.Fields("JobID"))
            gobjEventLog.ChangeHeaderStatus elsFailed
            gobjEventLog.AddDetailEntry strNotes
          End If
        
        Else
          ' A job previous in the batch has been cancelled,
          ' so populate event log with skipped
          gobjEventLog.AddHeader eltRecordProfile, GetUtilName("Record Profile", rsBatchDetails.Fields("JobID"))
          gobjEventLog.ChangeHeaderStatus elsSkipped
        End If
        Set pobjRecordProfileRUN = Nothing
      
      Case "Data Transfer"
        
        Dim pobjDataTransfer As New clsDataTransferRun
        Set pobjDataTransfer = New clsDataTransferRun
        If fOK And Not mblnBatchCancelled Then
          ' No previous jobs in the batch have been cancelled so run this
          
          'MH20011107 Check permissions...
          If datGeneral.SystemPermission("DATATRANSFER", "RUN") Then
            fOK = pobjDataTransfer.ExecuteDataTransfer(rsBatchDetails.Fields("JobID"))
            strNotes = pobjDataTransfer.ErrorString
            mblnJobCancelled = pobjDataTransfer.UserCancelled
          Else
            fOK = False
            strNotes = "You do not have permission to run data transfer definitions"
            gobjEventLog.AddHeader eltDataTransfer, GetUtilName("Data Transfer", rsBatchDetails.Fields("JobID"))
            gobjEventLog.ChangeHeaderStatus elsFailed
            gobjEventLog.AddDetailEntry strNotes
          End If
        Else
          ' A job previous in the batch has been cancelled,
          ' so populate event log with skipped
          gobjEventLog.AddHeader eltDataTransfer, GetUtilName("Data Transfer", rsBatchDetails.Fields("JobID"))
          gobjEventLog.ChangeHeaderStatus elsSkipped
        End If
        Set pobjDataTransfer = Nothing
        
      Case "Export"
        
        Dim pobjExportRUN As New clsExportRUN
        Set pobjExportRUN = New clsExportRUN
        pobjExportRUN.ExportID = rsBatchDetails.Fields("JobID")
        If fOK And Not mblnBatchCancelled Then
          ' No previous jobs in the batch have been cancelled so run this
          
          'MH20011107 Check permissions...
          If datGeneral.SystemPermission("EXPORT", "RUN") Then
            fOK = pobjExportRUN.RunExport
            strNotes = pobjExportRUN.ErrorString
            mblnJobCancelled = pobjExportRUN.UserCancelled
          Else
            fOK = False
            strNotes = "You do not have permission to run export definitions"
            gobjEventLog.AddHeader eltExport, GetUtilName("Export", rsBatchDetails.Fields("JobID"))
            gobjEventLog.ChangeHeaderStatus elsFailed
            gobjEventLog.AddDetailEntry strNotes
          End If
        Else
          ' A job previous in the batch has been cancelled,
          ' so populate the event log with skipped
          gobjEventLog.AddHeader eltExport, GetUtilName("Export", rsBatchDetails.Fields("JobID"))
          gobjEventLog.ChangeHeaderStatus elsSkipped
        End If
        Set pobjExportRUN = Nothing
      
      Case "Global Add"
        
        'Dim pobjGlobalAdd As New clsGlobalAddUpdateRun
        'Set pobjGlobalAdd = New clsGlobalAddUpdateRun
        Dim pobjGlobalAdd As New clsGlobalRun
        Set pobjGlobalAdd = New clsGlobalRun
        If fOK And Not mblnBatchCancelled Then
          ' No previous jobs in the batch have been cancelled so run this
          
          'MH20011107 Check permissions...
          If datGeneral.SystemPermission("GLOBALADD", "RUN") Then
            'fOK = pobjGlobalAdd.RunGlobalAddUpdate(rsBatchDetails.Fields("JobID"), True, glAdd)
            fOK = pobjGlobalAdd.RunGlobal(rsBatchDetails.Fields("JobID"), glAdd, "")
            strNotes = pobjGlobalAdd.ErrorString
            mblnJobCancelled = pobjGlobalAdd.UserCancelled
          Else
            fOK = False
            strNotes = "You do not have permission to run global add definitions"
            gobjEventLog.AddHeader eltGlobalAdd, GetUtilName("Global Add", rsBatchDetails.Fields("JobID"))
            gobjEventLog.ChangeHeaderStatus elsFailed
            gobjEventLog.AddDetailEntry strNotes
          End If
        Else
          ' A job previous in the batch has been cancelled,
          ' so populate the event log with skipped
          gobjEventLog.AddHeader eltGlobalAdd, GetUtilName("Global Add", rsBatchDetails.Fields("JobID"))
          gobjEventLog.ChangeHeaderStatus elsSkipped
        End If
        Set pobjGlobalAdd = Nothing
     
      Case "Global Delete"
        
        'Dim pobjGlobalDelete As New clsGlobalDeleteRun
        'Set pobjGlobalDelete = New clsGlobalDeleteRun
        Dim pobjGlobalDelete As New clsGlobalRun
        Set pobjGlobalDelete = New clsGlobalRun
        If fOK And Not mblnBatchCancelled Then
          ' No previous jobs in the batch have been cancelled so run this
          
          'MH20011107 Check permissions...
          If datGeneral.SystemPermission("GLOBALDELETE", "RUN") Then
            'fOK = pobjGlobalDelete.RunGlobalDelete(rsBatchDetails.Fields("JobID"), True)
            fOK = pobjGlobalDelete.RunGlobal(rsBatchDetails.Fields("JobID"), glDelete, "")
            strNotes = pobjGlobalDelete.ErrorString
            mblnJobCancelled = pobjGlobalDelete.UserCancelled
          Else
            fOK = False
            strNotes = "You do not have permission to run global delete definitions"
            gobjEventLog.AddHeader eltGlobalDelete, GetUtilName("Global Delete", rsBatchDetails.Fields("JobID"))
            gobjEventLog.ChangeHeaderStatus elsFailed
            gobjEventLog.AddDetailEntry strNotes
          End If
        Else
          ' A job previous in the batch has been cancelled,
          ' so populate the event log with skipped
          gobjEventLog.AddHeader eltGlobalDelete, GetUtilName("Global Delete", rsBatchDetails.Fields("JobID"))
          gobjEventLog.ChangeHeaderStatus elsSkipped
        End If
        Set pobjGlobalDelete = Nothing
        
      Case "Global Update"
      
        'Dim pobjGlobalUpdate As New clsGlobalAddUpdateRun
        'Set pobjGlobalUpdate = New clsGlobalAddUpdateRun
        Dim pobjGlobalUpdate As New clsGlobalRun
        Set pobjGlobalUpdate = New clsGlobalRun
        If fOK And Not mblnBatchCancelled Then
          ' No previous jobs in the batch have been cancelled so run this
          
          'MH20011107 Check permissions...
          If datGeneral.SystemPermission("GLOBALUPDATE", "RUN") Then
            'fOK = pobjGlobalUpdate.RunGlobalAddUpdate(rsBatchDetails.Fields("JobID"), True, glUpdate)
            fOK = pobjGlobalUpdate.RunGlobal(rsBatchDetails.Fields("JobID"), glUpdate, "")
            strNotes = pobjGlobalUpdate.ErrorString
            mblnJobCancelled = pobjGlobalUpdate.UserCancelled
          Else
            fOK = False
            strNotes = "You do not have permission to run global update definitions"
            gobjEventLog.AddHeader eltGlobalUpdate, GetUtilName("Global Update", rsBatchDetails.Fields("JobID"))
            gobjEventLog.ChangeHeaderStatus elsFailed
            gobjEventLog.AddDetailEntry strNotes
          End If
        Else
          ' A job previous in the batch has been cancelled,
          ' so populate the event log with skipped
          gobjEventLog.AddHeader eltGlobalUpdate, GetUtilName("Global Update", rsBatchDetails.Fields("JobID"))
          gobjEventLog.ChangeHeaderStatus elsSkipped
        End If
        Set pobjGlobalUpdate = Nothing
      
      Case "Import"
        
        Dim pobjImport As New clsImportRUN
        Set pobjImport = New clsImportRUN
        If fOK And Not mblnBatchCancelled Then
          ' No previous jobs in the batch have been cancelled so run this
          
          'MH20011107 Check permissions...
          If datGeneral.SystemPermission("IMPORT", "RUN") Then
            pobjImport.ImportID = rsBatchDetails.Fields("JobID")
            fOK = pobjImport.RunImport
            strNotes = pobjImport.ErrorString
            mblnJobCancelled = pobjImport.UserCancelled
          Else
            fOK = False
            strNotes = "You do not have permission to run import definitions"
            gobjEventLog.AddHeader eltImport, GetUtilName("Import", rsBatchDetails.Fields("JobID"))
            gobjEventLog.ChangeHeaderStatus elsFailed
            gobjEventLog.AddDetailEntry strNotes
          End If
        Else
          ' A job previous in the batch has been cancelled,
          ' so populate the event log with skipped
          gobjEventLog.AddHeader eltImport, GetUtilName("Import", rsBatchDetails.Fields("JobID"))
          gobjEventLog.ChangeHeaderStatus elsSkipped
        End If
        Set pobjImport = Nothing
        
      Case "Mail Merge"
      
        Dim pobjMailMergeRUN As clsMailMergeRun
        Set pobjMailMergeRUN = New clsMailMergeRun
        If fOK And Not mblnBatchCancelled Then
          ' No previous jobs in the batch have been cancelled so run this
          
          'MH20011107 Check permissions...
          If datGeneral.SystemPermission("MAILMERGE", "RUN") Then
            fOK = pobjMailMergeRUN.ExecuteMailMerge(rsBatchDetails.Fields("JobID"))
            strNotes = pobjMailMergeRUN.ErrorString
            mblnJobCancelled = pobjMailMergeRUN.UserCancelled
          Else
            fOK = False
            strNotes = "You do not have permission to run mail merge definitions"
            gobjEventLog.AddHeader eltMailMerge, GetUtilName("Mail Merge", rsBatchDetails.Fields("JobID"))
            gobjEventLog.ChangeHeaderStatus elsFailed
            gobjEventLog.AddDetailEntry strNotes
          End If
        Else
          ' A job previous in the batch has been cancelled,
          ' so populate the event log with skipped
          gobjEventLog.AddHeader eltMailMerge, GetUtilName("Mail Merge", rsBatchDetails.Fields("JobID"))
          gobjEventLog.ChangeHeaderStatus elsSkipped
        End If
        Set pobjMailMergeRUN = Nothing
      
      Case "Envelopes & Labels"
      
        Dim pobjLabelRUN As clsMailMergeRun
        Set pobjLabelRUN = New clsMailMergeRun
        If fOK And Not mblnBatchCancelled Then
          ' No previous jobs in the batch have been cancelled so run this
          
          'MH20011107 Check permissions...
          If datGeneral.SystemPermission("LABELS", "RUN") Then
            fOK = pobjLabelRUN.ExecuteMailMerge(rsBatchDetails.Fields("JobID"))
            strNotes = pobjLabelRUN.ErrorString
            mblnJobCancelled = pobjLabelRUN.UserCancelled
          Else
            fOK = False
            strNotes = "You do not have permission to run envelope & Label definitions"
            gobjEventLog.AddHeader eltLabel, GetUtilName("Envelopes & Labels", rsBatchDetails.Fields("JobID"))
            gobjEventLog.ChangeHeaderStatus elsFailed
            gobjEventLog.AddDetailEntry strNotes
          End If
        Else
          ' A job previous in the batch has been cancelled,
          ' so populate the event log with skipped
          gobjEventLog.AddHeader eltLabel, GetUtilName("Envelopes & Labels", rsBatchDetails.Fields("JobID"))
          gobjEventLog.ChangeHeaderStatus elsSkipped
        End If
        Set pobjLabelRUN = Nothing
      
      Case "Match Report"

        Dim pfrmMatchRUN As frmMatchRun
        Set pfrmMatchRUN = New frmMatchRun
        pfrmMatchRUN.MatchReportID = rsBatchDetails.Fields("JobID")
        If fOK And Not mblnBatchCancelled Then
          ' No previous jobs in the batch have been cancelled so run this
          
          If datGeneral.SystemPermission("MATCHREPORTS", "RUN") Then
            fOK = pfrmMatchRUN.RunMatchReport
            strNotes = pfrmMatchRUN.ErrorString
            mblnJobCancelled = pfrmMatchRUN.UserCancelled
          Else
            fOK = False
            strNotes = "You do not have permission to run match report definitions"
            gobjEventLog.AddHeader eltMatchReport, GetUtilName("Match Report", rsBatchDetails.Fields("JobID"))
            gobjEventLog.ChangeHeaderStatus elsFailed
            gobjEventLog.AddDetailEntry strNotes
          End If

        Else
          ' A job previous in the batch has been cancelled,
          ' so populate event log with skipped
          gobjEventLog.AddHeader eltMatchReport, GetUtilName("Match Report", rsBatchDetails.Fields("JobID"))
          gobjEventLog.ChangeHeaderStatus elsSkipped
        End If
        
        Set pfrmMatchRUN = Nothing
              
      Case "Succession Planning"

        'Dim pfrmMatchRUN As frmMatchRun
        Set pfrmMatchRUN = New frmMatchRun
        pfrmMatchRUN.MatchReportID = rsBatchDetails.Fields("JobID")
        pfrmMatchRUN.MatchReportType = mrtSucession
        If fOK And Not mblnBatchCancelled Then
          ' No previous jobs in the batch have been cancelled so run this
          
          If datGeneral.SystemPermission("SUCCESSION", "RUN") Then
            fOK = pfrmMatchRUN.RunMatchReport
            strNotes = pfrmMatchRUN.ErrorString
            mblnJobCancelled = pfrmMatchRUN.UserCancelled
          Else
            fOK = False
            strNotes = "You do not have permission to run succession planning definitions"
            gobjEventLog.AddHeader eltSuccessionPlanning, GetUtilName("Succession Planning", rsBatchDetails.Fields("JobID"))
            gobjEventLog.ChangeHeaderStatus elsFailed
            gobjEventLog.AddDetailEntry strNotes
          End If

        Else
          ' A job previous in the batch has been cancelled,
          ' so populate event log with skipped
          gobjEventLog.AddHeader eltSuccessionPlanning, GetUtilName("Succession Planning", rsBatchDetails.Fields("JobID"))
          gobjEventLog.ChangeHeaderStatus elsSkipped
        End If
        
        Set pfrmMatchRUN = Nothing
              
      Case "Career Progression"

        'Dim pfrmMatchRUN As frmMatchRun
        Set pfrmMatchRUN = New frmMatchRun
        pfrmMatchRUN.MatchReportID = rsBatchDetails.Fields("JobID")
        pfrmMatchRUN.MatchReportType = mrtCareer
        If fOK And Not mblnBatchCancelled Then
          ' No previous jobs in the batch have been cancelled so run this
          
          If datGeneral.SystemPermission("CAREER", "RUN") Then
            fOK = pfrmMatchRUN.RunMatchReport
            strNotes = pfrmMatchRUN.ErrorString
            mblnJobCancelled = pfrmMatchRUN.UserCancelled
          Else
            fOK = False
            strNotes = "You do not have permission to run career progression definitions"
            gobjEventLog.AddHeader eltCareerProgression, GetUtilName("Career Progression", rsBatchDetails.Fields("JobID"))
            gobjEventLog.ChangeHeaderStatus elsFailed
            gobjEventLog.AddDetailEntry strNotes
          End If

        Else
          ' A job previous in the batch has been cancelled,
          ' so populate event log with skipped
          gobjEventLog.AddHeader eltCareerProgression, GetUtilName("Career Progression", rsBatchDetails.Fields("JobID"))
          gobjEventLog.ChangeHeaderStatus elsSkipped
        End If
        
        Set pfrmMatchRUN = Nothing
              
      Case "Stability Index Report"
        Dim pobjStabilityReport As frmCrossTabRun
        Set pobjStabilityReport = New frmCrossTabRun
        
        If fOK And Not mblnBatchCancelled Then
          If datGeneral.SystemPermission("STANDARDREPORTS", "RUN_SI") Then
            If gfPersonnelEnabled Then
              fOK = pobjStabilityReport.TurnoverStabilityReport
              strNotes = pobjStabilityReport.ErrorString
              mblnJobCancelled = pobjStabilityReport.UserCancelled
            Else
              'JPD 20040227 Fault 8156
              fOK = False
              strNotes = "The system is not configured to run the Stability Index Report"
              gobjEventLog.AddHeader eltStandardReport, "Stability Index Report"
              gobjEventLog.ChangeHeaderStatus elsFailed
              gobjEventLog.AddDetailEntry strNotes
            End If
          Else
            fOK = False
            strNotes = "You do not have permission to run the Stability Index Report"
            gobjEventLog.AddHeader eltStandardReport, "Stability Index Report"
            gobjEventLog.ChangeHeaderStatus elsFailed
            gobjEventLog.AddDetailEntry strNotes
          End If
        Else
          ' A job previous in the batch has been cancelled,
          ' so populate event log with skipped
          gobjEventLog.AddHeader eltStandardReport, "Stability Index Report"
          gobjEventLog.ChangeHeaderStatus elsSkipped
        End If
        Set pobjStabilityReport = Nothing
      
      
      Case "Turnover Report"
        Dim pobjTurnoverReport As frmCrossTabRun
        Set pobjTurnoverReport = New frmCrossTabRun

        If fOK And Not mblnBatchCancelled Then
          If datGeneral.SystemPermission("STANDARDREPORTS", "RUN_TR") Then
            If gfPersonnelEnabled Then
              fOK = pobjTurnoverReport.TurnoverExecuteReport
              strNotes = pobjTurnoverReport.ErrorString
              mblnJobCancelled = pobjTurnoverReport.UserCancelled
            Else
              'JPD 20040227 Fault 8156
              fOK = False
              strNotes = "The system is not configured to run the Turnover Report"
              gobjEventLog.AddHeader eltStandardReport, "Turnover Report"
              gobjEventLog.ChangeHeaderStatus elsFailed
              gobjEventLog.AddDetailEntry strNotes
            End If
          Else
            fOK = False
            strNotes = "You do not have permission to run the Turnover Report"
            gobjEventLog.AddHeader eltStandardReport, "Turnover Report"
            gobjEventLog.ChangeHeaderStatus elsFailed
            gobjEventLog.AddDetailEntry strNotes
          End If
        Else
          ' A job previous in the batch has been cancelled,
          ' so populate event log with skipped
          gobjEventLog.AddHeader eltStandardReport, "Turnover Report"
          gobjEventLog.ChangeHeaderStatus elsSkipped
        End If
        Set pobjTurnoverReport = Nothing


      Case Else
        
        gobjEventLog.AddHeader eltSystemError, "Unknown Job in Batch"
        gobjEventLog.ChangeHeaderStatus elsSkipped
        fOK = False
        
    End Select
    
    
    If Not mblnBatchCancelled Then
      If mblnJobCancelled Or gobjProgress.Cancelled Then
        gobjProgress.Visible = False
        mblnBatchCancelled = True
      End If
    End If

    mstrEventLogIDs = mstrEventLogIDs & _
      IIf(mstrEventLogIDs <> vbNullString, ", ", vbNullString) & _
      CStr(gobjEventLog.EventLogID)
    
    If gobjProgress.Visible Then
      gobjProgress.UpdateProgress2 True
    End If
    rsBatchDetails.MoveNext
  Loop


  CheckNotifications lBatchJobID

  Set rsBatchDetails = Nothing
  
  gobjEventLog.TidyUp
  
  
  Call UtilUpdateLastRun(IIf(gblnReportPackMode, utlReportPack, utlBatchJob), lBatchJobID)

  RunBatchJob = strNotes

  gblnBatchMode = False

  'Clean up and exit
  gobjProgress.CloseProgress
  Set rsBatchDetails = Nothing
  
End Function

Public Function SetLastCompleteDate(lBatchJobID As Long, dLastComp As Date) As Boolean

  On Error GoTo SetLastCompleteDate_ERROR
  
  Dim pstrSQL As String
  
  'Update the last completed field in ASRSysBatchJobName for the current job
  pstrSQL = "UPDATE AsrSysBatchJobName SET " & _
            "LastCompleted = '" & Replace(Format(dLastComp, "mm/dd/yy"), UI.GetSystemDateSeparator, "/") & "', " & _
            "Lockspid = null, " & _
            "LockLoginTime = null " & _
            "WHERE ID = " & CStr(lBatchJobID)
  mclsData.ExecuteSql (pstrSQL)

  SetLastCompleteDate = True
  Exit Function
  
SetLastCompleteDate_ERROR:
  
  COAMsgBox "Error setting the LastCompleted Date In AsrSysBatchJobName" & vbCrLf & "(" & Err.Description & ")"
  SetLastCompleteDate = False
      
End Function

Private Function GetUtilName(strUtilType As String, lngUtilID As Long) As String

  Dim pstrSQL As String
  Dim prstTemp As Recordset

  pstrSQL = vbNullString
  GetUtilName = "<Unknown>"
  
  Select Case strUtilType
  Case "Calendar Report":    pstrSQL = "ASRSysCalendarReports WHERE ID"
  Case "Cross Tab":          pstrSQL = "ASRSysCrossTab WHERE CrossTabID"
  Case "Custom Report":      pstrSQL = "ASRSysCustomReportsName WHERE ID"
  Case "Data Transfer":      pstrSQL = "ASRSysDataTransferName WHERE DataTransferID"
  Case "Export":             pstrSQL = "ASRSysExportName WHERE ID"
  Case "Global Add":         pstrSQL = "ASRSysGlobalFunctions WHERE Type = 'A' AND FunctionID"
  Case "Global Delete":      pstrSQL = "ASRSysGlobalFunctions WHERE Type = 'D' AND FunctionID"
  Case "Global Update":      pstrSQL = "ASRSysGlobalFunctions WHERE Type = 'U' AND FunctionID"
  Case "Import":             pstrSQL = "ASRSysImportName WHERE ID"
  Case "Mail Merge":         pstrSQL = "ASRSysMailMergeName WHERE MailMergeID"
  Case "Envelopes & Labels": pstrSQL = "ASRSysMailMergeName WHERE MailMergeID"
  Case "Record Profile":     pstrSQL = "ASRSysRecordProfileName WHERE RecordProfileID"

  Case "Match Report", "Succession Planning", "Career Progression"
                             pstrSQL = "ASRSysMatchReportName WHERE MatchReportID"
  End Select

  If pstrSQL <> vbNullString Then
    pstrSQL = "SELECT [Name] FROM " & pstrSQL & " = " & CStr(lngUtilID)
    Set prstTemp = mclsData.OpenRecordset(pstrSQL, adOpenForwardOnly, adLockReadOnly)
  
    If Not (prstTemp.BOF And prstTemp.EOF) Then
      GetUtilName = prstTemp.Fields(0)
    End If
  
    Set prstTemp = Nothing
  End If

End Function


Public Function CheckBatchNeedsRunning2(lBatchJobNameID As Long, sBatchJobName As String) As String

  On Error GoTo CheckBatchNeedsRunning_ERROR
  
  Dim rsTemp As Recordset
  Dim strSQL As String

  Dim dTestDate As String
  Dim pintCounter As Integer
  'Dim pstrNotes As String
  Dim pblnBeenRunOnce As Boolean
  'Dim pfrmPrompt As New frmBatchJobPrompt

  Dim iFrequency As Integer
  Dim sPeriod As String
  Dim bWeekends As Boolean
  Dim dStartDate As Date
  Dim sLastCompDate As String
  Dim bRunOnce As Boolean
  
  ' For the event log header record
  'Dim plngEventLogID As Long
  Dim blnRequiresRunning As Boolean
  
  blnRequiresRunning = False

  strSQL = "SELECT * FROM ASRSysBatchJobName WHERE ID = " & CStr(lBatchJobNameID)
  Set rsTemp = mclsData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
  
  
  If rsTemp.BOF And rsTemp.EOF Then
    CheckBatchNeedsRunning2 = "The job '" & sBatchJobName & "' has been deleted by another user."
    Exit Function
  End If
  
  
  With rsTemp
    sBatchJobName = .Fields("Name").Value
    iFrequency = .Fields("Frequency").Value
    sPeriod = .Fields("Period").Value
    bWeekends = .Fields("Weekends").Value
    dStartDate = .Fields("StartDate").Value
    sLastCompDate = IIf(IsNull(.Fields("LastCompleted")), "", .Fields("LastCompleted"))
    bRunOnce = .Fields("RunOnce")
  End With
  
  
  'Start checking dates from the start date
  dTestDate = dStartDate

  'do this until the test date is greater than todays date
  Do Until CDate(dTestDate) > Date

    'If the batch job hasnt been run before, set lastcompdate to startdate and skip the
    'next IF line.
    If Not IsDate(sLastCompDate) Then
      sLastCompDate = dStartDate
      GoTo wasntalastcompdate
    End If

    If CDate(dTestDate) > CDate(sLastCompDate) Then

    'RH: ok, so this bit is really naff, but it works and theres no time to improve it...

wasntalastcompdate:

      'Check if 31st of month has been selected and testdate month doesnt have 31 days
      If sPeriod = "M" Then
        Do Until Day(CDate(dTestDate)) <= Day(CDate(sLastCompDate)) And ((Month(CDate(dTestDate)) = (Month(CDate(sLastCompDate)) + pintCounter) Mod 12) Or (Month(CDate(sLastCompDate)) + pintCounter) Mod 12 = 0)
          'dTestDate = Format((CDate(dTestDate) - 1), "dd/mm/yyyy")
          dTestDate = DateAdd("d", -1, dTestDate)
        Loop
      End If

      'Check if 29/02 is selected, if so, if not leap year, use 28/02 instead
      If sPeriod = "Y" And Day(dStartDate) = 29 And Month(dStartDate) = 2 Then
        Do Until Day(CDate(dTestDate)) <= Day(CDate(sLastCompDate)) And (Month(CDate(dTestDate)) = 2)
          'dTestDate = Format((CDate(dTestDate) - 1), "dd/mm/yyyy")
          dTestDate = DateAdd("d", -1, dTestDate)
        Loop
      End If

      If bWeekends = False And (Weekday(CDate(dTestDate)) = vbSaturday Or Weekday(CDate(dTestDate)) = vbSunday) Then
        'dont run it
      Else
        If bRunOnce Then
          blnRequiresRunning = Not pblnBeenRunOnce
        Else
          blnRequiresRunning = True
        End If
      End If

    End If

    pintCounter = pintCounter + 1
    
    'Select Case sPeriod
    'Case "D": dTestDate = Format(DateSerial(Year(CDate(dStartDate)), Month(CDate(dStartDate)), Day(CDate(dStartDate)) + (iFrequency * pintCounter)), "dd/mm/yyyy")
    'Case "W": dTestDate = Format(DateSerial(Year(CDate(dStartDate)), Month(CDate(dStartDate)), Day(CDate(dStartDate)) + (iFrequency * 7 * pintCounter)), "dd/mm/yyyy")
    'Case "M": dTestDate = Format(DateSerial(Year(CDate(dStartDate)), Month(CDate(dStartDate)) + (iFrequency * pintCounter), Day(CDate(dStartDate))), "dd/mm/yyyy")
    'Case "Y": dTestDate = Format(DateSerial(Year(CDate(dStartDate)) + (iFrequency * pintCounter), Month(CDate(dStartDate)), Day(CDate(dStartDate))), "dd/mm/yyyy")
    'End Select
    Select Case sPeriod
    Case "D": dTestDate = DateSerial(Year(CDate(dStartDate)), Month(CDate(dStartDate)), Day(CDate(dStartDate)) + (iFrequency * pintCounter))
    Case "W": dTestDate = DateSerial(Year(CDate(dStartDate)), Month(CDate(dStartDate)), Day(CDate(dStartDate)) + (iFrequency * 7 * pintCounter))
    Case "M": dTestDate = DateSerial(Year(CDate(dStartDate)), Month(CDate(dStartDate)) + (iFrequency * pintCounter), Day(CDate(dStartDate)))
    Case "Y": dTestDate = DateSerial(Year(CDate(dStartDate)) + (iFrequency * pintCounter), Month(CDate(dStartDate)), Day(CDate(dStartDate)))
    End Select

  Loop

  
  If Not blnRequiresRunning Then
    CheckBatchNeedsRunning2 = "The job '" & sBatchJobName & "' will not be run as it has already been run by another user."
  Else
    CheckBatchNeedsRunning2 = vbNullString
  End If

  Exit Function

CheckBatchNeedsRunning_ERROR:
  CheckBatchNeedsRunning2 = "Error whilst calculating whether " & IIf(gblnReportPackMode, "report pack", "batch job") & " is due or not." & vbCrLf & "Contact Asr Support stating : " & Err.Description

End Function


Public Function LockJob(lngID As Long, blnClearLock As Boolean) As Boolean

  'Dim rsTemp As Recordset
  'Dim strSQL As String

  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  'Dim lngSpid As Long

  On Local Error GoTo LocalErr
  
  
  LockJob = False

  Set cmADO = New ADODB.Command
  With cmADO
    .CommandText = "spASRLockWriteBatch"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon

    Set pmADO = .CreateParameter("BatchJobID", adInteger, adParamInput)
    .Parameters.Append pmADO
    pmADO.Value = lngID

    Set pmADO = .CreateParameter("ClearLock", adBoolean, adParamInput)
    .Parameters.Append pmADO
    pmADO.Value = blnClearLock
    
    Set pmADO = .CreateParameter("LockedByOther", adInteger, adParamOutput)
    .Parameters.Append pmADO

    Set pmADO = Nothing


    .Execute
    LockJob = (IIf(IsNull(.Parameters(1).Value), 0, .Parameters(1).Value) = 0)

  End With

  Set cmADO = Nothing

Exit Function

LocalErr:
  COAMsgBox Err.Description, vbExclamation

End Function

Public Property Get EventLogIDs() As String
  EventLogIDs = mstrEventLogIDs
End Property

Public Property Let EventLogIDs(ByVal strNewValue As String)
  mstrEventLogIDs = strNewValue
End Property

Public Property Get JobStatus() As EventLog_Status
  
  If mblnJobCancelled Then
    JobStatus = elsCancelled
  ElseIf fOK = False Then
    JobStatus = elsFailed
  Else
    JobStatus = elsSuccessful
  End If

End Property


Public Sub CheckNotifications(lngBatchJobID As Long)

  Dim frmSendEmail As frmEmailSel
  Dim rsBatch As Recordset

  If datGeneral.SystemPermission("EMAILGROUPS", "VIEW") Then

    Set rsBatch = mclsData.OpenRecordset("SELECT Name, EmailFailed, EmailSuccess FROM AsrSysBatchJobName WHERE ID = " & CStr(lngBatchJobID), adOpenKeyset, adLockReadOnly)

    If Not mblnJobCancelled Then
      If fOK = False Then
        If rsBatch!EmailFailed > 0 Then
          Set frmSendEmail = New frmEmailSel
          frmSendEmail.SendBatchNotification "OpenHR " & IIf(gblnReportPackMode, "Report Pack", "Batch Job") & " Failed: '" & rsBatch!Name & "'", mstrEventLogIDs, rsBatch!EmailFailed
          Set frmSendEmail = Nothing
        End If
      Else
        If rsBatch!EmailSuccess > 0 Then
          Set frmSendEmail = New frmEmailSel
          frmSendEmail.SendBatchNotification "OpenHR " & IIf(gblnReportPackMode, "Report Pack", "Batch Job") & " Succeeded: '" & rsBatch!Name & "'", mstrEventLogIDs, rsBatch!EmailFailed
          Set frmSendEmail = Nothing
        End If
      End If
    End If

    rsBatch.Close
    Set rsBatch = Nothing

  End If

End Sub
