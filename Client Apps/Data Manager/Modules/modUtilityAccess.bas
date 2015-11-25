Attribute VB_Name = "modUtilityAccess"
Option Explicit

Public Enum RecordSelectionTypes
  REC_SEL_ALL = 0
  REC_SEL_PICKLIST = 1
  REC_SEL_FILTER = 2
End Enum

Public Const ACCESS_READWRITE = "RW"
Public Const ACCESS_READONLY = "RO"
Public Const ACCESS_HIDDEN = "HD"
Public Const ACCESS_UNKNOWN = ""

Public Const ACCESSDESC_READWRITE = "Read / Write"
Public Const ACCESSDESC_READONLY = "Read Only"
Public Const ACCESSDESC_HIDDEN = "Hidden"
Public Const ACCESSDESC_UNKNOWN = "Unknown"

Public Enum RecordSelectionValidityCodes
  REC_SEL_VALID_OK = 0
  REC_SEL_VALID_DELETED = 1
  REC_SEL_VALID_HIDDENBYUSER = 2
  REC_SEL_VALID_HIDDENBYOTHER = 3
  REC_SEL_VALID_INVALID = 4
End Enum



Public Sub CheckCanMakeHiddenInBatchJobs(piUtilityType As UtilityType, _
  psIDs As String, _
  pstrUser As String, _
  ByRef piOwnedJobCount As Integer, _
  ByRef psOwnedJobDetails As String, _
  ByRef psOwnedJobIDs As String, _
  ByRef psNonOwnedJobDetails As String, _
  ByRef pblnBatchJobsOK As Boolean, _
  ByRef psScheduledJobDetails As String, _
  ByRef psScheduledUserGroups As String, _
  Optional pvHiddenToGroups As Variant)
                                   
  ' Check for any Batch Jobs that contain the given utility/report.
  Dim sSQL As String
  Dim sTableName As String
  Dim sAccessTableName As String
  Dim sIDColumnName As String
  Dim rsTemp As ADODB.Recordset
  Dim sKey As String
  Dim fHiddenToAllGroups As Boolean
  Dim sHiddenToGroups As String
  Dim sHiddenToGroupsList As String
  
  fHiddenToAllGroups = IsMissing(pvHiddenToGroups)
  sHiddenToGroups = IIf(fHiddenToAllGroups, "", UCase(CStr(pvHiddenToGroups)))
  
  If Len(sHiddenToGroups) > 0 Then
    sHiddenToGroupsList = "'" & Replace(Mid(sHiddenToGroups, 2, Len(sHiddenToGroups) - 2), vbTab, "','") & "'"
  Else
    sHiddenToGroupsList = "''"
  End If
  
  pstrUser = LCase(pstrUser)
  
  Select Case piUtilityType
    Case utlCalendarReport
      sTableName = "ASRSysCalendarReports"
      sKey = "Calendar Report"
      sIDColumnName = "ID"
    
    Case utlCrossTab
      sTableName = "ASRSysCrossTab"
      sKey = "Cross Tab"
      sIDColumnName = "CrossTabID"
        
    Case utlCustomReport
      sTableName = "ASRSysCustomReportsName"
      sKey = "Custom Report"
      sIDColumnName = "ID"
  
    Case utlDataTransfer
      sTableName = "AsrSysDataTransferName"
      sKey = "Data Transfer"
      sIDColumnName = "DataTransferID"
  
    Case utlExport
      sTableName = "AsrSysExportName"
      sKey = "Export"
      sIDColumnName = "ID"
  
    Case UtlGlobalAdd
      sTableName = "AsrSysGlobalFunctions"
      sKey = "Global Add"
      sIDColumnName = "FunctionID"
  
    Case utlGlobalUpdate
      sTableName = "AsrSysGlobalFunctions"
      sKey = "Global Update"
      sIDColumnName = "FunctionID"
  
    Case utlGlobalDelete
      sTableName = "AsrSysGlobalFunctions"
      sKey = "Global Delete"
      sIDColumnName = "FunctionID"
  
    Case utlImport
      sTableName = "AsrSysImportName"
      sKey = "Import"
      sIDColumnName = "ID"
  
    Case utlMailMerge
      sTableName = "AsrSysMailMergeName"
      sKey = "Mail Merge"
      sIDColumnName = "MailMergeID"
  
    Case utlLabel
      sTableName = "AsrSysMailMergeName"
      sKey = "Envelopes & Labels"
      sIDColumnName = "MailMergeID"
  
    Case utlRecordProfile
      sTableName = "ASRSysRecordProfileName"
      sKey = "Record Profile"
      sIDColumnName = "recordProfileID"
      
    Case utlMatchReport
      sTableName = "ASRSysMatchReportName"
      sKey = "Match Report"
      sIDColumnName = "matchReportID"
  
    Case utlSuccession
      sTableName = "ASRSysMatchReportName"
      sKey = "Succession Planning"
      sIDColumnName = "matchReportID"
  
    Case utlCareer
      sTableName = "ASRSysMatchReportName"
      sKey = "Career Progression"
      sIDColumnName = "matchReportID"
  
  End Select
  
  If Len(sTableName) > 0 Then
    sSQL = "SELECT ASRSysBatchJobName.Name," & _
      " ASRSysBatchJobName.ID," & _
      " convert(integer, ASRSysBatchJobName.scheduled) AS [scheduled]," & _
      " ASRSysBatchJobName.roleToPrompt," & _
      " COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount]," & _
      " ASRSysBatchJobName.Username," & _
      " " & sTableName & ".Name AS 'JobName'" & _
      " FROM ASRSysBatchJobDetails" & _
      " INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID " & _
      " INNER JOIN " & sTableName & " ON " & sTableName & "." & sIDColumnName & " = ASRSysBatchJobDetails.jobID"
    sSQL = sSQL & _
      " LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID" & _
      "   AND ASRSysBatchJobAccess.access <> '" & ACCESS_HIDDEN & "'"
      
    If Not fHiddenToAllGroups Then
      sSQL = sSQL & _
        "   AND ASRSysBatchJobAccess.groupName IN (" & sHiddenToGroupsList & ")"
    End If
    
    sSQL = sSQL & _
      "   AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name" & _
      "     FROM sysusers" & _
      "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
      "       AND ASRSysGroupPermissions.permitted = 1" & _
      "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
      "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
      "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
      "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
      "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
      "     WHERE sysusers.uid = sysusers.gid" & _
      "       AND sysusers.uid <> 0)"
    sSQL = sSQL & _
      " WHERE ASRSysBatchJobDetails.JobType = '" & sKey & "' " & _
      "   AND ASRSysBatchJobDetails.JobID IN (" & psIDs & ")" & _
      " GROUP BY ASRSysBatchJobName.Name," & _
      "   ASRSysBatchJobName.ID," & _
      "   convert(integer, ASRSysBatchJobName.scheduled)," & _
      "   ASRSysBatchJobName.roleToPrompt," & _
      "   ASRSysBatchJobName.Username," & _
      "   " & sTableName & ".Name"
  
    Set rsTemp = datGeneral.GetReadOnlyRecords(sSQL)
        
    Do Until rsTemp.EOF
      If LCase(rsTemp!userName) = pstrUser Then
        ' Found a Batch Job whose owner is the same
        If (rsTemp!scheduled = 1) And _
          (Len(rsTemp!RoleToPrompt) > 0) And _
          (UCase(rsTemp!RoleToPrompt) <> UCase(gsUserGroup)) And _
          (fHiddenToAllGroups Or (InStr(sHiddenToGroups, vbTab & UCase(rsTemp!RoleToPrompt) & vbTab) > 0)) Then
          ' Found a Batch Job which is scheduled for another user group to run.
          pblnBatchJobsOK = False
      
          psScheduledUserGroups = psScheduledUserGroups & rsTemp!RoleToPrompt & vbCrLf
          
          If CurrentUserAccess(utlBatchJob, rsTemp!ID) = ACCESS_HIDDEN Then
            psScheduledJobDetails = psScheduledJobDetails & "Batch Job : <Hidden> by " & rsTemp!userName & vbCrLf
          Else
            psScheduledJobDetails = psScheduledJobDetails & "Batch Job : " & rsTemp!Name & vbCrLf
          End If
        ElseIf rsTemp!nonHiddenCount > 0 Then
          piOwnedJobCount = piOwnedJobCount + 1
          psOwnedJobDetails = psOwnedJobDetails & "Batch Job : " & rsTemp!Name & " (Contains " & sKey & " '" & rsTemp!jobname & "') " & vbCrLf
          psOwnedJobIDs = psOwnedJobIDs & IIf(Len(psOwnedJobIDs) > 0, ", ", "") & rsTemp!ID
        End If
      'JPD 20041124 Fault 9382
      'Else
      ElseIf rsTemp!nonHiddenCount > 0 Then
        ' Found a Batch Job whose owner is not the same
        pblnBatchJobsOK = False
    
        If CurrentUserAccess(utlBatchJob, rsTemp!ID) = ACCESS_HIDDEN Then
          psNonOwnedJobDetails = psNonOwnedJobDetails & "Batch Job : <Hidden> by " & rsTemp!userName & vbCrLf
        Else
          psNonOwnedJobDetails = psNonOwnedJobDetails & "Batch Job : " & rsTemp!Name & vbCrLf
        End If
      End If
      
      rsTemp.MoveNext
    Loop
                                   
    rsTemp.Close
    Set rsTemp = Nothing
  End If
  
End Sub



Public Sub CheckForPicklistsExpressions(piUtilityType As UtilityType, _
  psSQL As String, _
  pstrUser As String, _
  ByRef piOwnedCount As Integer, _
  ByRef psOwnedDetails As String, _
  ByRef psOwnedIDs As String, _
  ByRef piNonOwnedCount As Integer, _
  ByRef psNonOwnedDetails As String)
                                   
  ' Check for any of the given utility/report definitions that contain picklists/expressions (as defined in the SQL code thats passed in).
  Dim rsTemp As ADODB.Recordset
  Dim sKey As String
  Dim objComp As clsExprComponent
  
  sKey = ""
  pstrUser = LCase(pstrUser)
  
  Select Case piUtilityType
    Case utlCalculation
      sKey = "Calculation"
    
    Case utlCalendarReport
      sKey = "Calendar Report"
    
    Case utlCrossTab
      sKey = "Cross Tab"
    
    Case utlCustomReport
      sKey = "Custom Report"
    
    Case utlDataTransfer
      sKey = "Data Transfer"
    
    Case utlExport
      sKey = "Export"
    
    Case utlFilter
      sKey = "Filter"
    
    Case UtlGlobalAdd
      sKey = "Global Add"
  
    Case utlGlobalUpdate
      sKey = "Global Update"
  
    Case utlGlobalDelete
      sKey = "Global Delete"
    
    Case utlImport
      sKey = "Import"
    
    Case utlLabel
      sKey = "Envelopes & Labels"
    
    Case utlMailMerge
      sKey = "Mail Merge"
    
    Case utlRecordProfile
      sKey = "Record Profile"
    
    Case utlMatchReport
      sKey = "Match Report"
  
    Case utlSuccession
      sKey = "Succession Planning"
  
    Case utlCareer
      sKey = "Career Progression"
    
  End Select

  If Len(sKey) > 0 Then
    Set rsTemp = datGeneral.GetReadOnlyRecords(psSQL)
              
    Do Until rsTemp.EOF
      
      Select Case piUtilityType
        Case utlCalculation, utlFilter
          If LCase(rsTemp!userName) = pstrUser Then
            ' Found a tool whose owner is the same
            If rsTemp!Access <> ACCESS_HIDDEN Then
              piOwnedCount = piOwnedCount + 1
              psOwnedDetails = psOwnedDetails & sKey & " : " & rsTemp!Name & vbCrLf
              psOwnedIDs = psOwnedIDs & IIf(Len(psOwnedIDs) > 0, ", ", "") & rsTemp!ID
            End If
          Else
            ' Found a tool whose owner is not the same
              piNonOwnedCount = piNonOwnedCount + 1
      
            If rsTemp!Access = ACCESS_HIDDEN Then
              psNonOwnedDetails = psNonOwnedDetails & sKey & " : <Hidden> by " & rsTemp!userName & vbCrLf
            Else
              psNonOwnedDetails = psNonOwnedDetails & sKey & " : " & rsTemp!Name & vbCrLf
            End If
          End If
      
        Case Else
          If LCase(rsTemp!userName) = pstrUser Then
            ' Found a utility/report whose owner is the same
            If rsTemp!nonHiddenCount > 0 Then
                piOwnedCount = piOwnedCount + 1
                psOwnedDetails = psOwnedDetails & sKey & " : " & rsTemp!Name & vbCrLf
                psOwnedIDs = psOwnedIDs & IIf(Len(psOwnedIDs) > 0, ", ", "") & rsTemp!ID
            End If
          Else
            ' Found a utility/report whose owner is not the same
            piNonOwnedCount = piNonOwnedCount + 1
            
            If CurrentUserAccess(piUtilityType, rsTemp!ID) = ACCESS_HIDDEN Then
              psNonOwnedDetails = psNonOwnedDetails & sKey & " : <Hidden> by " & rsTemp!userName & vbCrLf
            Else
              psNonOwnedDetails = psNonOwnedDetails & sKey & " : " & rsTemp!Name & vbCrLf
            End If
          End If
      
      End Select

      rsTemp.MoveNext
    Loop
    
    rsTemp.Close
    Set rsTemp = Nothing
  End If
  
End Sub





Public Function GetAllExprRootIDs(plngID As Long) As String
  Dim rsTemp As Recordset
  Dim objComp As clsExprComponent
  Dim strSQL As String
  Dim sSuperExprs As String
  Dim lngExprID As Long
  
  GetAllExprRootIDs = vbNullString
  
  strSQL = "SELECT componentID" & _
    " FROM ASRSysExprComponents" & _
    " WHERE (calculationID = " & plngID & ")" & _
    "   OR (filterID = " & plngID & ")" & _
    "   OR ((fieldSelectionFilter = " & plngID & ")" & _
    "     AND (type = " & CStr(giCOMPONENT_FIELD) & "))"

  Set rsTemp = datGeneral.GetRecords(strSQL)
  With rsTemp
    Do While Not .EOF
      Set objComp = New clsExprComponent
      objComp.ComponentID = !ComponentID

      lngExprID = objComp.RootExpressionID
      
      GetAllExprRootIDs = GetAllExprRootIDs & _
        IIf(GetAllExprRootIDs <> vbNullString, ", ", vbNullString) & _
        CStr(lngExprID)
    
      sSuperExprs = GetAllExprRootIDs(lngExprID)
      If sSuperExprs <> vbNullString Then
        GetAllExprRootIDs = GetAllExprRootIDs & IIf(GetAllExprRootIDs <> vbNullString, ",", "") & sSuperExprs
      End If
      
      .MoveNext
    Loop
    Set objComp = Nothing
    .Close
  
  End With
  Set rsTemp = Nothing
        
End Function



Public Function CurrentUserIsSysSecMgr() As Boolean
  Dim sSQL As String
  Dim rsAccess As ADODB.Recordset
  Dim datData As DataMgr.clsDataAccess
  Dim fIsSysSecUser As Boolean
  
    sSQL = "SELECT count(*) AS [result]" & _
  " FROM ASRSysGroupPermissions" & _
  " INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
  "   AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
  "   OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
  " INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
  "   AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
  " INNER JOIN sysusers b ON b.name = ASRSysGroupPermissions.groupname" & _
  "   AND b.name = '" & gsUserGroup & "'" & _
  " WHERE ASRSysGroupPermissions.permitted = 1"

  '" INNER JOIN sysusers b ON b.name = ASRSysGroupPermissions.groupname" & _
  " INNER JOIN sysusers a ON b.uid = a.gid" & _
  "   AND a.Name = current_user" & _
  " WHERE ASRSysGroupPermissions.permitted = 1"

  Set datData = New clsDataAccess
  Set rsAccess = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  With rsAccess
    fIsSysSecUser = (!Result > 0)
    
    .Close
  End With
  Set rsAccess = Nothing
  
  Set datData = Nothing
  
  CurrentUserIsSysSecMgr = fIsSysSecUser

End Function


Public Function OldAccessUtility(piUtilityType As UtilityType) As Boolean
  ' Returns TRUE if the given utility type is still using the old access set-up.
  OldAccessUtility = (piUtilityType = utlAbsenceBreakdown) Or _
    (piUtilityType = utlBradfordFactor) Or _
    (piUtilityType = utlCalculation) Or _
    (piUtilityType = utlEmailAddress) Or _
    (piUtilityType = utlEmailGroup) Or _
    (piUtilityType = utlFilter) Or _
    (piUtilityType = utlLabelType) Or _
    (piUtilityType = utlDocumentMapping) Or _
    (piUtilityType = utlOrder) Or _
    (piUtilityType = utlPicklist)

End Function


Public Sub UtilityAmended(piUtilityType As UtilityType, _
  plngRecordID As Long, _
  plngTimestamp As Long, _
  blnContinueSave As Boolean, _
  blnSaveAsNew As Boolean, _
  Optional strType As String)
  
  Dim datData As clsDataAccess
  Dim sSQL As String
  Dim rsCheck As Recordset
  Dim rsTemp As Recordset
  Dim sTemp As String
  Dim strMBText As String
  Dim intMBResponse As Integer
  Dim blnTimeStampChanged As Boolean
  Dim blnDeletedDef As Boolean
  Dim blnReadOnly As Boolean
  Dim sCurrentUserAccess As String
  Dim sTable As String
  Dim sIDColumn As String
  Dim sTypeCode As String
  
  On Error GoTo Amended_ERROR
  
  If strType = vbNullString Then
    strType = "definition"
  End If

  blnSaveAsNew = False
  blnContinueSave = True
  
  If plngRecordID = 0 Then
    ' The record cannot have been modified by another user if it hs not yet been given a record ID.
    Exit Sub
  End If
  
  ' NB. See modHRPro.GetTypeCodeFromTable for the type codes.
  Select Case piUtilityType
    Case utlBatchJob
      sTable = "ASRSysBatchJobName"
      sIDColumn = "ID"
      sTypeCode = "0"
      
    Case utlCalendarReport
      sTable = "ASRSysCalendarReports"
      sIDColumn = "ID"
      sTypeCode = "15"
    
    Case utlCrossTab
      sTable = "ASRSysCrossTab"
      sIDColumn = "CrossTabID"
      sTypeCode = "1"
    
    Case utlCustomReport
      sTable = "ASRSysCustomReportsName"
      sIDColumn = "ID"
      sTypeCode = "2"
    
    Case utlDataTransfer
      sTable = "ASRSysDataTransferName"
      sIDColumn = "dataTransferID"
      sTypeCode = "3"
      
    Case utlExport
      sTable = "ASRSysExportName"
      sIDColumn = "ID"
      sTypeCode = "4"
      
    Case UtlGlobalAdd, utlGlobalDelete, utlGlobalUpdate
      sTable = "ASRSysGlobalFunctions"
      sIDColumn = "functionID"
      sTypeCode = "5,6,7"
      
    Case utlImport
      sTable = "ASRSysImportName"
      sIDColumn = "ID"
      sTypeCode = "8"
      
    Case utlLabel, utlMailMerge
      sTable = "ASRSysMailMergeName"
      sIDColumn = "mailMergeID"
      sTypeCode = "9"
    
    Case utlRecordProfile
      sTable = "ASRSysRecordProfileName"
      sIDColumn = "recordProfileID"
      sTypeCode = "20"
  
    Case utlMatchReport, utlSuccession, utlCareer
      sTable = "ASRSysMatchReportName"
      sIDColumn = "matchReportID"
      sTypeCode = "14"
  
  End Select
  
  Set datData = New clsDataAccess
  ' Compare the given Timestamp with the Timestamp in the given record on the server.
  sSQL = "SELECT convert(int, timestamp) AS TimeStamp, UserName " & _
         " FROM " & sTable & _
         " WHERE " & sIDColumn & " = " & Trim(Str(plngRecordID))
  Set rsCheck = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  blnTimeStampChanged = True
  blnReadOnly = False
  
  blnDeletedDef = (rsCheck.BOF And rsCheck.EOF)
  If Not blnDeletedDef Then
    blnTimeStampChanged = (plngTimestamp <> rsCheck!Timestamp)
    
    sCurrentUserAccess = CurrentUserAccess(piUtilityType, plngRecordID)
    blnReadOnly = (LCase$(rsCheck!userName) <> LCase$(gsUserName)) And _
      (sCurrentUserAccess <> ACCESS_READWRITE)
  End If
  
  rsCheck.Close
  Set rsCheck = Nothing
  Set datData = Nothing

  If blnTimeStampChanged = False Then
    Exit Sub
  End If
  
  If blnDeletedDef Or blnReadOnly Then
    'Unable to overwrite existing definition
    If blnDeletedDef Then
      strMBText = "The current " & strType & " has been deleted by another user."
    Else
      strMBText = "The current " & strType & " has been amended by another user and is now " & AccessDescription(sCurrentUserAccess) & "."
    End If
                  
    strMBText = strMBText & vbCrLf & _
                "Save as a new " & strType & "?"
    intMBResponse = COAMsgBox(strMBText, vbExclamation + vbOKCancel, app.ProductName)
      
    Select Case intMBResponse
      Case vbOK         'save as new (but this may cause duplicate name message)
        blnContinueSave = True
        blnSaveAsNew = True
      Case vbCancel     'Do not save
        blnContinueSave = False
    End Select
  Else
    ' Use this to get the host name, as W95/8 doesnt like UI.GetHostName
    Set rsTemp = datGeneral.GetReadOnlyRecords("SELECT HOST_NAME()")
    sTemp = rsTemp.Fields(0)
    
    Set rsTemp = datGeneral.GetReadOnlyRecords("SELECT SavedHost " & _
                                               "FROM   ASRSysUtilAccessLog " & _
                                               "WHERE  Type IN (" & sTypeCode & ") " & _
                                               "AND    UtilID = " & plngRecordID)
    
    If LCase(rsTemp.Fields("SavedHost")) <> LCase((sTemp)) Then
      ' If the definition was last changed by somebody else (rather than by
      ' automatically due to the access rights of a plist/filter/calc being
      ' changed, then prompt, otherwise, just overwrite it.
      
      'Prompt to see if user should overwrite definition
      strMBText = "The current " & strType & " has been amended by another user. " & vbCrLf & _
                  "Would you like to overwrite this " & strType & "?" & vbCrLf
      intMBResponse = COAMsgBox(strMBText, vbExclamation + vbYesNoCancel, app.ProductName)
      
      Select Case intMBResponse
        Case vbYes        'overwrite existing definition and any changes
          blnContinueSave = True
        Case vbNo         'save as new (but this may cause duplicate name message)
          blnContinueSave = True
          blnSaveAsNew = True
        Case vbCancel     'Do not save
          blnContinueSave = False
      End Select
    Else
      blnContinueSave = True
    End If
  End If

  Exit Sub
  
Amended_ERROR:
  
  COAMsgBox "Error whilst checking if utility definition has been amended." & vbCrLf & vbCrLf & "(" & Err.Number & " - " & Err.Description & ")", vbExclamation + vbOKOnly, app.title
  blnContinueSave = False
  
End Sub


Public Sub HideUtilities(piUtilityType As UtilityType, _
  psIDs As String, _
  Optional pvHiddenUserGroups As Variant)
  ' Set the access for the given utility to be HIDDEN for the given user groups.
  ' psIDs is a COMMA delimited string of the utility/report IDs to be hidden
  ' pvHiddenUserGroups is a TAB delimited string of the user groups
  '   to which these definitions are to be hidden.
  '   NB. this string starts with a TAB also.
  
  Dim sSQL As String
  Dim datData As DataMgr.clsDataAccess
  Dim sTableName As String
  Dim sAccessTableName As String
  Dim sIDColumnName As String
  Dim fHideFromAll As Boolean
  Dim sHiddenUserGroups As String
  
  fHideFromAll = IsMissing(pvHiddenUserGroups)
  sHiddenUserGroups = IIf(fHideFromAll, "", CStr(pvHiddenUserGroups))
  
  Set datData = New clsDataAccess

  Select Case piUtilityType
    Case utlBatchJob
      sTableName = "ASRSysBatchJobName"
      sAccessTableName = "ASRSysBatchJobAccess"
      sIDColumnName = "ID"

    Case utlCalculation, utlFilter
      sTableName = "ASRSysExpressions"
      sIDColumnName = "exprID"
    
    Case utlCalendarReport
      sTableName = "ASRSysCalendarReports"
      sAccessTableName = "ASRSysCalendarReportAccess"
      sIDColumnName = "ID"
    
    Case utlCrossTab
      sTableName = "ASRSysCrossTab"
      sAccessTableName = "ASRSysCrossTabAccess"
      sIDColumnName = "CrossTabID"
    
    Case utlCustomReport
      sTableName = "ASRSysCustomReportsName"
      sAccessTableName = "ASRSysCustomReportAccess"
      sIDColumnName = "ID"
    
    Case utlDataTransfer
      sTableName = "ASRSysDataTransferName"
      sAccessTableName = "ASRSysDataTransferAccess"
      sIDColumnName = "dataTransferID"

    Case utlExport
      sTableName = "ASRSysExportName"
      sAccessTableName = "ASRSysExportAccess"
      sIDColumnName = "ID"

    Case UtlGlobalAdd, utlGlobalDelete, utlGlobalUpdate
      sTableName = "ASRSysGlobalFunctions"
      sAccessTableName = "ASRSysGlobalAccess"
      sIDColumnName = "functionID"
    
    Case utlImport
      sTableName = "ASRSysImportName"
      sAccessTableName = "ASRSysImportAccess"
      sIDColumnName = "ID"

    Case utlLabel, utlMailMerge
      sTableName = "ASRSysMailMergeName"
      sAccessTableName = "ASRSysMailMergeAccess"
      sIDColumnName = "mailMergeID"
    
    Case utlRecordProfile
      sTableName = "ASRSysRecordProfileName"
      sAccessTableName = "ASRSysRecordProfileAccess"
      sIDColumnName = "recordProfileID"
  
    Case utlMatchReport, utlSuccession, utlCareer
      sTableName = "ASRSysMatchReportName"
      sAccessTableName = "ASRSysMatchReportAccess"
      sIDColumnName = "matchReportID"
  
  End Select
  
  Select Case piUtilityType
    Case utlCalculation, utlFilter
      sSQL = "UPDATE " & sTableName & _
        " SET access = '" & ACCESS_HIDDEN & "'" & _
        " WHERE " & sIDColumnName & " IN (" & psIDs & ")"
      datData.ExecuteSql sSQL
   
    Case Else
      If Len(sAccessTableName) > 0 Then
    
        If fHideFromAll Then
          sSQL = "DELETE FROM " & sAccessTableName & " WHERE ID IN (" & psIDs & ")"
          datData.ExecuteSql sSQL
        
          sSQL = "INSERT INTO " & sAccessTableName & _
            " (ID, groupName, access)" & _
            " (SELECT " & sTableName & "." & sIDColumnName & ", sysusers.name," & _
            " CASE" & _
            "   WHEN (SELECT count(*)" & _
            "     FROM ASRSysGroupPermissions" & _
            "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
            "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
            "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
            "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
            "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
            "     WHERE sysusers.Name = ASRSysGroupPermissions.groupname" & _
            "       AND ASRSysGroupPermissions.permitted = 1) > 0 THEN '" & ACCESS_READWRITE & "'" & _
            "   ELSE '" & ACCESS_HIDDEN & "'" & _
            " END" & _
            " FROM sysusers," & sTableName & _
            " WHERE sysusers.uid = sysusers.gid" & _
            " AND sysusers.uid <> 0" & _
            " AND sysusers.name <> 'ASRSysGroup'" & _
            " AND " & sTableName & "." & sIDColumnName & " IN (" & psIDs & "))"
          datData.ExecuteSql (sSQL)
        Else
          sHiddenUserGroups = "'" & Replace(Mid(Left(sHiddenUserGroups, Len(sHiddenUserGroups) - 1), 2), vbTab, "','") & "'"
          
          sSQL = "DELETE FROM " & sAccessTableName & " WHERE ID IN (" & psIDs & ") AND groupName IN (" & sHiddenUserGroups & ")"
          datData.ExecuteSql sSQL
          
          sSQL = "INSERT INTO " & sAccessTableName & _
            " (ID, groupName, access)" & _
            " (SELECT " & sTableName & "." & sIDColumnName & ", sysusers.name," & _
            " CASE" & _
            "   WHEN (SELECT count(*)" & _
            "     FROM ASRSysGroupPermissions" & _
            "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
            "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
            "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
            "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
            "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
            "     WHERE sysusers.Name = ASRSysGroupPermissions.groupname" & _
            "       AND ASRSysGroupPermissions.permitted = 1) > 0 THEN '" & ACCESS_READWRITE & "'" & _
            "   ELSE '" & ACCESS_HIDDEN & "'" & _
            " END" & _
            " FROM sysusers," & sTableName & _
            " WHERE sysusers.uid = sysusers.gid" & _
            " AND sysusers.uid <> 0" & _
            " AND sysusers.name IN (" & sHiddenUserGroups & ")" & _
            " AND " & sTableName & "." & sIDColumnName & " IN (" & psIDs & "))"
          datData.ExecuteSql (sSQL)
        End If
      End If
  End Select
  
  Set datData = Nothing
  
End Sub



Public Function GetUtilityAccessRecords(piUtilityType As UtilityType, _
  plngID As Long, _
  pfFromCopy As Boolean) As ADODB.Recordset
  ' Return a recordset of all user groups and their access setting for the given utility.
  ' First field in the recordset is 'name' - the user group name.
  ' Second field in the recordset is 'access' - RW/RO/HD
  ' Third field in the recordset is 'sysSecMgr' - 1 if the user group is a SysMgr or SecMgr user.
  On Error GoTo ErrorTrap
  
  Dim sSQL As String
  Dim sDefaultAccess As String
  Dim rsAccess As ADODB.Recordset
  Dim datData As DataMgr.clsDataAccess
  Dim sAccessTableName As String
  Dim sKey As String
  
  sDefaultAccess = ACCESS_HIDDEN
  
  ' Construct the SQL code to get the access settings for the given utility.
  ' NB. System and Security Manager users automatically have Read/Write access.
  Select Case piUtilityType
    Case utlBatchJob
      sAccessTableName = "ASRSysBatchJobAccess"
      sKey = "Batch Jobs"

    Case utlReportPack
      sAccessTableName = "ASRSysBatchJobAccess"
      sKey = "Report Packs"
      
    Case utlCalendarReport
      sAccessTableName = "ASRSysCalendarReportAccess"
      sKey = "Calendar Reports"
    
    Case utlCrossTab
      sAccessTableName = "ASRSysCrossTabAccess"
      sKey = "Cross Tabs"
    
    Case utlCustomReport
      sAccessTableName = "ASRSysCustomReportAccess"
      sKey = "Custom Reports"
    
    Case utlDataTransfer
      sAccessTableName = "ASRSysDataTransferAccess"
      sKey = "Data Transfer"
    
    Case utlExport
      sAccessTableName = "ASRSysExportAccess"
      sKey = "Export"
    
    Case UtlGlobalAdd
      sAccessTableName = "ASRSysGlobalAccess"
      sKey = "Global Add"
    
    Case utlGlobalDelete
      sAccessTableName = "ASRSysGlobalAccess"
      sKey = "Global Delete"
    
    Case utlGlobalUpdate
      sAccessTableName = "ASRSysGlobalAccess"
      sKey = "Global Update"
    
    Case utlImport
      sAccessTableName = "ASRSysImportAccess"
      sKey = "Import"
    
    Case utlLabel
      sAccessTableName = "ASRSysMailMergeAccess"
      sKey = "Labels"
      
    Case utlMailMerge
      sAccessTableName = "ASRSysMailMergeAccess"
      sKey = "Mail Merge"
      
    Case utlRecordProfile
      sAccessTableName = "ASRSysRecordProfileAccess"
      sKey = "Record Profile"
  
    Case utlMatchReport
      sAccessTableName = "ASRSysMatchReportAccess"
      sKey = "Match Reports"
  
    Case utlSuccession
      sAccessTableName = "ASRSysMatchReportAccess"
      sKey = "Succession Planning"
  
    Case utlCareer
      sAccessTableName = "ASRSysMatchReportAccess"
      sKey = "Career Progression"
  
  End Select

  If (plngID = 0) Or (pfFromCopy) Then
    ' Read from the user-configuration
    sDefaultAccess = GetUserSetting("utils&reports", "dfltaccess " & Replace(sKey, " ", ""), ACCESS_READWRITE)
  End If
  
  sSQL = "SELECT sysusers.name," & _
    "  CASE" & _
    "    WHEN (SELECT count(*)" & _
    "      FROM ASRSysGroupPermissions" & _
    "      INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
    "        AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
    "          OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
    "      INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
    "        AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
    "      WHERE sysusers.Name = ASRSysGroupPermissions.groupname" & _
    "        AND ASRSysGroupPermissions.permitted = 1) > 0 THEN '" & ACCESS_READWRITE & "'" & _
    "    ELSE"
    
  If (plngID = 0) Or (pfFromCopy) Then
    sSQL = sSQL & _
      "      '" & sDefaultAccess & "'" & _
      "  END As access,"
  Else
    sSQL = sSQL & _
      "      CASE" & _
      "        WHEN " & sAccessTableName & ".access IS null THEN '" & sDefaultAccess & "'" & _
      "        ELSE " & sAccessTableName & ".access" & _
      "      END" & _
      "  END As access,"
  End If
  
  sSQL = sSQL & _
    "  CASE" & _
    "    WHEN (SELECT count(*)" & _
    "      FROM ASRSysGroupPermissions" & _
    "      INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
    "        AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
    "          OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
    "      INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
    "        AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
    "      WHERE sysusers.Name = ASRSysGroupPermissions.groupName" & _
    "        AND ASRSysGroupPermissions.permitted = 1) > 0 THEN 1" & _
    "    ELSE" & _
    "      0" & _
    "  END AS sysSecMgr" & _
    " FROM sysusers" & _
    " LEFT OUTER JOIN " & sAccessTableName & " ON (sysusers.name = " & sAccessTableName & ".groupName" & _
    "  AND " & sAccessTableName & ".id = " & CStr(plngID) & ")" & _
    " WHERE sysusers.uid = sysusers.gid" & _
    " AND ISNULL(sysusers.uid, 0) <> 0" & _
    " AND NOT (sysusers.name LIKE 'ASRSys%') AND NOT (sysusers.name LIKE 'db[_]%')" & _
    " ORDER BY sysusers.name"

  'MH20031211 Added the line about ASRSys% above


  Set datData = New clsDataAccess
  Set rsAccess = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Set datData = Nothing
  
TidyUpAndExit:
  Set GetUtilityAccessRecords = rsAccess
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Function

Public Function GetSysSecMgrUserGroups() As ADODB.Recordset
  ' Return a recordset of all sysSecMgr user groups
  On Error GoTo ErrorTrap
  
  Dim sSQL As String
  Dim rsAccess As ADODB.Recordset
  Dim datData As DataMgr.clsDataAccess
  
  sSQL = "SELECT DISTINCT sysusers.name" & _
    " FROM sysusers" & _
    " INNER JOIN ASRSysGroupPermissions" & _
    "   ON sysusers.name = ASRSysGroupPermissions.groupName" & _
    "     AND ASRSysGroupPermissions.permitted = 1" & _
    " INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
    "   AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
    "     OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
    " INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
    "   AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
    " WHERE sysusers.uid = sysusers.gid" & _
    "   AND sysusers.uid <> 0" & _
    "   AND NOT (sysusers.name LIKE 'ASRSys%')" & _
    " ORDER BY sysusers.name"

  Set datData = New clsDataAccess
  Set rsAccess = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Set datData = Nothing
  
TidyUpAndExit:
  Set GetSysSecMgrUserGroups = rsAccess
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Function


Public Function GetUtilityAccessRecordsIgnoreSysSecUsers(piUtilityType As UtilityType, _
  plngID As Long, _
  pfFromCopy As Boolean) As ADODB.Recordset
  ' Return a recordset of all user groups and their access setting for the given utility.
  ' First field in the recordset is 'name' - the user group name.
  ' Second field in the recordset is 'access' - RW/RO/HD
  ' Third field in the recordset is 'sysSecMgr' - 1 if the user group is a SysMgr or SecMgr user.
  On Error GoTo ErrorTrap
  
  Dim sSQL As String
  Dim sDefaultAccess As String
  Dim rsAccess As ADODB.Recordset
  Dim datData As DataMgr.clsDataAccess
  Dim sAccessTableName As String
  Dim sKey As String
  
  sDefaultAccess = ACCESS_HIDDEN
  
  ' Construct the SQL code to get the access settings for the given utility.
  ' NB. System and Security Manager users automatically have Read/Write access.
  Select Case piUtilityType
    Case utlBatchJob
      sAccessTableName = "ASRSysBatchJobAccess"
      sKey = "Batch Jobs"

    Case utlCalendarReport
      sAccessTableName = "ASRSysCalendarReportAccess"
      sKey = "Calendar Reports"
    
    Case utlCrossTab
      sAccessTableName = "ASRSysCrossTabAccess"
      sKey = "Cross Tabs"
    
    Case utlCustomReport
      sAccessTableName = "ASRSysCustomReportAccess"
      sKey = "Custom Reports"
    
    Case utlDataTransfer
      sAccessTableName = "ASRSysDataTransferAccess"
      sKey = "Data Transfer"
    
    Case utlExport
      sAccessTableName = "ASRSysExportAccess"
      sKey = "Export"
    
    Case UtlGlobalAdd
      sAccessTableName = "ASRSysGlobalAccess"
      sKey = "Global Add"
    
    Case utlGlobalDelete
      sAccessTableName = "ASRSysGlobalAccess"
      sKey = "Global Delete"
    
    Case utlGlobalUpdate
      sAccessTableName = "ASRSysGlobalAccess"
      sKey = "Global Update"
    
    Case utlImport
      sAccessTableName = "ASRSysImportAccess"
      sKey = "Import"
    
    Case utlLabel
      sAccessTableName = "ASRSysMailMergeAccess"
      sKey = "Labels"
      
    Case utlMailMerge
      sAccessTableName = "ASRSysMailMergeAccess"
      sKey = "Mail Merge"
      
    Case utlRecordProfile
      sAccessTableName = "ASRSysRecordProfileAccess"
      sKey = "Record Profile"
  
    Case utlMatchReport
      sAccessTableName = "ASRSysMatchReportAccess"
      sKey = "Match Reports"
  
    Case utlSuccession
      sAccessTableName = "ASRSysMatchReportAccess"
      sKey = "Succession Planning"
  
    Case utlCareer
      sAccessTableName = "ASRSysMatchReportAccess"
      sKey = "Career Progression"
  
  End Select

  If (plngID = 0) Or (pfFromCopy) Then
    ' Read from the user-configuration
    sDefaultAccess = GetUserSetting("utils&reports", "dfltaccess " & Replace(sKey, " ", ""), ACCESS_READWRITE)
  End If
  
  sSQL = "SELECT sysusers.name,"
    
  If (plngID = 0) Or (pfFromCopy) Then
    sSQL = sSQL & _
      "      '" & sDefaultAccess & "'" & _
      "  As access"
  Else
    sSQL = sSQL & _
      "      CASE" & _
      "        WHEN " & sAccessTableName & ".access IS null THEN '" & sDefaultAccess & "'" & _
      "        ELSE " & sAccessTableName & ".access" & _
      "      END" & _
      "  As access"
  End If
  
  sSQL = sSQL & _
    " FROM sysusers" & _
    " LEFT OUTER JOIN " & sAccessTableName & " ON (sysusers.name = " & sAccessTableName & ".groupName" & _
    "  AND " & sAccessTableName & ".id = " & CStr(plngID) & ")" & _
    " WHERE sysusers.uid = sysusers.gid" & _
    " AND sysusers.uid <> 0" & _
    " AND NOT (sysusers.name LIKE 'ASRSys%') " & _
    " ORDER BY sysusers.name"

  'MH20031211 Added the line about ASRSys% above


  Set datData = New clsDataAccess
  Set rsAccess = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Set datData = Nothing
  
TidyUpAndExit:
  Set GetUtilityAccessRecordsIgnoreSysSecUsers = rsAccess
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Function


Public Function GetUtilityOwner(piUtilityType As UtilityType, _
  plngID As Long) As String
  ' Return a the owner of the given utility.
  ' Empty string returned indicates that the utility no longer exists.
  On Error GoTo ErrorTrap
  
  Dim sSQL As String
  Dim rsOwner As ADODB.Recordset
  Dim sTableName As String
  Dim sIDColumnName As String
  Dim sOwner As String
  
  sOwner = ""
  sTableName = ""
  
  ' Construct the SQL code to get the owner of the given utility.
  Select Case piUtilityType
    Case utlBatchJob
      sTableName = "ASRSysBatchJobName"
      sIDColumnName = "ID"
      
    Case utlCalendarReport
      sTableName = "ASRSysCalendarReports"
      sIDColumnName = "ID"
    
    Case utlCrossTab
      sTableName = "ASRSysCrossTab"
      sIDColumnName = "CrossTabID"
    
    Case utlCustomReport
      sTableName = "ASRSysCustomReportsName"
      sIDColumnName = "ID"
    
    Case utlDataTransfer
      sTableName = "ASRSysDataTransferName"
      sIDColumnName = "dataTransferID"
      
    Case utlExport
      sTableName = "ASRSysExportName"
      sIDColumnName = "ID"
      
    Case utlFilter, utlCalculation
      sTableName = "ASRSysExpressions"
      sIDColumnName = "exprID"

    Case UtlGlobalAdd, utlGlobalDelete, utlGlobalUpdate
      sTableName = "ASRSysGlobalFunctions"
      sIDColumnName = "functionID"

    Case utlImport
      sTableName = "ASRSysImportName"
      sIDColumnName = "ID"
      
    Case utlLabel, utlMailMerge
      sTableName = "ASRSysMailMergeName"
      sIDColumnName = "mailMergeID"
    
    Case utlPicklist
      sTableName = "ASRSysPicklistName"
      sIDColumnName = "picklistID"
  
    Case utlRecordProfile
      sTableName = "ASRSysRecordProfileName"
      sIDColumnName = "recordProfileID"
  
    Case utlMatchReport, utlSuccession, utlCareer
      sTableName = "ASRSysMatchReportName"
      sIDColumnName = "matchReportID"
  
  End Select
  
  If Len(sTableName) > 0 Then
    sSQL = "SELECT userName" & _
      " FROM " & sTableName & _
      " WHERE " & sIDColumnName & " = " & CStr(plngID)
  
    Set rsOwner = datGeneral.GetReadOnlyRecords(sSQL)
        
    If Not (rsOwner.BOF And rsOwner.EOF) Then
      sOwner = IIf(IsNull(rsOwner!userName), "", rsOwner!userName)
    End If
      
    rsOwner.Close
    Set rsOwner = Nothing
  End If
  
TidyUpAndExit:
  GetUtilityOwner = sOwner
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Function


Public Function CurrentUserAccess(piUtilityType As UtilityType, _
  plngID As Long) As String
  ' Return the access code (RW/RO/HD) of the current user's access
  ' on the given utility.
  On Error GoTo ErrorTrap
  
  Dim sAccessCode As String
  Dim sSQL As String
  Dim sDefaultAccess As String
  Dim rsAccess As ADODB.Recordset
  Dim datData As DataMgr.clsDataAccess
  Dim sTableName As String
  Dim sAccessTableName As String
  Dim sIDColumnName As String
  
  sTableName = ""
  sAccessTableName = ""
  
  If plngID > 0 Then
    sDefaultAccess = ACCESS_HIDDEN
  Else
    sDefaultAccess = ACCESS_HIDDEN
  End If
  
  ' Construct the SQL code to get the current user's access settings for the given utility.
  ' NB. System and Security Manager users automatically have Read/Write access.
  Select Case piUtilityType
    Case utlBatchJob
      sTableName = "ASRSysBatchJobName"
      sAccessTableName = "ASRSysBatchJobAccess"
      sIDColumnName = "ID"
    
    Case utlReportPack
      sTableName = "ASRSysBatchJobName"
      sAccessTableName = "ASRSysBatchJobAccess"
      sIDColumnName = "ID"
      
    Case utlCalendarReport
      sTableName = "ASRSysCalendarReports"
      sAccessTableName = "ASRSysCalendarReportAccess"
      sIDColumnName = "ID"
    
    Case utlCrossTab
      sTableName = "ASRSysCrossTab"
      sAccessTableName = "ASRSysCrossTabAccess"
      sIDColumnName = "CrossTabID"
    
    Case utlCustomReport
      sTableName = "ASRSysCustomReportsName"
      sAccessTableName = "ASRSysCustomReportAccess"
      sIDColumnName = "ID"
    
    Case utlDataTransfer
      sTableName = "ASRSysDataTransferName"
      sAccessTableName = "ASRSysDataTransferAccess"
      sIDColumnName = "DataTransferID"
      
    Case utlExport
      sTableName = "ASRSysExportName"
      sAccessTableName = "ASRSysExportAccess"
      sIDColumnName = "ID"
      
    Case UtlGlobalAdd, utlGlobalDelete, utlGlobalUpdate
      sTableName = "ASRSysGlobalFunctions"
      sAccessTableName = "ASRSysGlobalAccess"
      sIDColumnName = "functionID"

    Case utlImport
      sTableName = "ASRSysImportName"
      sAccessTableName = "ASRSysImportAccess"
      sIDColumnName = "ID"

    Case utlLabel, utlMailMerge
      sTableName = "ASRSysMailMergeName"
      sAccessTableName = "ASRSysMailMergeAccess"
      sIDColumnName = "mailMergeID"
    
    Case utlRecordProfile
      sTableName = "ASRSysRecordProfileName"
      sAccessTableName = "ASRSysRecordProfileAccess"
      sIDColumnName = "recordProfileID"
  
    Case utlMatchReport, utlSuccession, utlCareer
      sTableName = "ASRSysMatchReportName"
      sAccessTableName = "ASRSysMatchReportAccess"
      sIDColumnName = "matchReportID"
  
  End Select
  
  If Len(sAccessTableName) > 0 Then
    'sSQL = "SELECT" & _
      "  CASE" & _
      "    WHEN (SELECT count(*)" & _
      "      FROM ASRSysGroupPermissions" & _
      "      INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
      "        AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
      "        OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
      "      INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
      "        AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
      "      WHERE b.Name = ASRSysGroupPermissions.groupname" & _
      "        AND ASRSysGroupPermissions.permitted = 1) > 0 THEN '" & ACCESS_READWRITE & "'" & _
      "    WHEN " & sTableName & ".userName = system_user THEN '" & ACCESS_READWRITE & "'" & _
      "    ELSE" & _
      "      CASE" & _
      "        WHEN " & sAccessTableName & ".access IS null THEN '" & sDefaultAccess & "'" & _
      "        ELSE " & sAccessTableName & ".access" & _
      "      END" & _
      "  END AS Access" & _
      " FROM sysusers b" & _
      " INNER JOIN sysusers a ON b.uid = a.gid" & _
      " LEFT OUTER JOIN " & sAccessTableName & " ON (b.name = " & sAccessTableName & ".groupName" & _
      "   AND " & sAccessTableName & ".id = " & CStr(plngID) & ")" & _
      " INNER JOIN " & sTableName & " ON " & sAccessTableName & ".ID = " & sTableName & "." & sIDColumnName & _
      " WHERE a.Name = current_user"
    sSQL = "SELECT" & _
      "  CASE" & _
      "    WHEN (SELECT count(*)" & _
      "      FROM ASRSysGroupPermissions" & _
      "      INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
      "        AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
      "        OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
      "      INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
      "        AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
      "      WHERE b.Name = ASRSysGroupPermissions.groupname" & _
      "        AND ASRSysGroupPermissions.permitted = 1) > 0 THEN '" & ACCESS_READWRITE & "'" & _
      "    WHEN " & sTableName & ".userName = system_user THEN '" & ACCESS_READWRITE & "'" & _
      "    ELSE" & _
      "      CASE" & _
      "        WHEN " & sAccessTableName & ".access IS null THEN '" & sDefaultAccess & "'" & _
      "        ELSE " & sAccessTableName & ".access" & _
      "      END" & _
      "  END AS Access" & _
      " FROM sysusers b" & _
      " LEFT OUTER JOIN " & sAccessTableName & " ON (b.name = " & sAccessTableName & ".groupName" & _
      "   AND " & sAccessTableName & ".id = " & CStr(plngID) & ")" & _
      " INNER JOIN " & sTableName & " ON " & sAccessTableName & ".ID = " & sTableName & "." & sIDColumnName & _
      " WHERE b.Name = '" & gsUserGroup & "'"
    
    Set datData = New clsDataAccess
    
    Set rsAccess = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    With rsAccess
      If .BOF And .EOF Then
        sAccessCode = sDefaultAccess
      Else
        sAccessCode = !Access
      End If
      
      .Close
    End With
    Set rsAccess = Nothing
    
    Set datData = Nothing
  Else
    sAccessCode = ACCESS_UNKNOWN
  End If
  
TidyUpAndExit:
  CurrentUserAccess = sAccessCode
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Function



Public Function UtilityIsHiddenToAnyone(piUtilityType As UtilityType, _
  plngUtilityID As Long) As Boolean
  ' Return TRUE if the given utility is hidden to anyone.
  On Error GoTo ErrorTrap
  
  Dim fIsHidden As Boolean
  Dim sSQL As String
  Dim rsAccess As ADODB.Recordset
  Dim datData As DataMgr.clsDataAccess
  Dim sAccessTableName As String
  
  fIsHidden = True
  
  Select Case piUtilityType
    Case utlBatchJob
      sAccessTableName = "ASRSysBatchJobAccess"
    
    Case utlCalendarReport
      sAccessTableName = "ASRSysCalendarReportAccess"
    
    Case utlCrossTab
      sAccessTableName = "ASRSysCrossTabAccess"
    
    Case utlCustomReport
      sAccessTableName = "ASRSysCustomReportAccess"
    
    Case utlDataTransfer
      sAccessTableName = "ASRSysDataTransferAccess"
    
    Case utlExport
      sAccessTableName = "ASRSysExportAccess"
    
    Case UtlGlobalAdd, utlGlobalDelete, utlGlobalUpdate
      sAccessTableName = "ASRSysGlobalAccess"
    
    Case utlImport
      sAccessTableName = "ASRSysImportAccess"
    
    Case utlLabel, utlMailMerge
      sAccessTableName = "ASRSysMailMergeAccess"
    
    Case utlRecordProfile
      sAccessTableName = "ASRSysRecordProfileAccess"
  
    Case utlMatchReport, utlSuccession, utlCareer
      sAccessTableName = "ASRSysMatchReportAccess"
      
  End Select
  
  sSQL = "SELECT COUNT(sysusers.name) AS result" & _
    " FROM sysusers" & _
    " LEFT OUTER JOIN " & sAccessTableName & " ON (sysusers.name = " & sAccessTableName & ".groupName" & _
    "  AND " & sAccessTableName & ".id = " & CStr(plngUtilityID) & ")" & _
    " WHERE sysusers.uid = sysusers.gid" & _
    "  AND sysusers.name <> 'ASRSysGroup'" & _
    "  AND sysusers.uid <> 0" & _
    "  AND (CASE" & _
    "    WHEN (SELECT count(*)" & _
    "      FROM ASRSysGroupPermissions" & _
    "      INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
    "        AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
    "          OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
    "      INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
    "        AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
    "      WHERE sysusers.Name = ASRSysGroupPermissions.groupname" & _
    "        AND ASRSysGroupPermissions.permitted = 1) > 0 THEN '" & ACCESS_READWRITE & "'" & _
    "    ELSE" & _
    "      CASE" & _
    "        WHEN " & sAccessTableName & ".access IS null THEN '" & ACCESS_HIDDEN & "'" & _
    "        ELSE " & sAccessTableName & ".access" & _
    "      END" & _
    "  END = '" & ACCESS_HIDDEN & "')"

  Set datData = New clsDataAccess
  Set rsAccess = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  fIsHidden = (rsAccess!Result > 0)
  rsAccess.Close
  Set rsAccess = Nothing
  Set datData = Nothing
  
  Set datData = Nothing
  
TidyUpAndExit:
  UtilityIsHiddenToAnyone = fIsHidden
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit

End Function

Public Function ValidateRecordSelection(piType As RecordSelectionTypes, _
  plngID As Long) As RecordSelectionValidityCodes
  ' Return an integer code representing the validity of the record selection (picklist or filter).
  ' Return 0 if the record selection is OK.
  ' Return 1 if the record selection has been deleted by another user.
  ' Return 2 if the record selection is hidden, and is owned by the current user.
  ' Return 3 if the record selection is hidden, and is NOT owned by the current user.
  ' Return 4 if the record selection is no longer valid.
  On Error GoTo ErrorTrap
  
  Dim iResult As RecordSelectionValidityCodes
  
  iResult = REC_SEL_VALID_OK
  
  Select Case piType
    Case REC_SEL_PICKLIST
      iResult = ValidatePicklist(plngID)
  
    Case REC_SEL_FILTER
      iResult = ValidateFilter(plngID)
  End Select
    
TidyUpAndExit:
  ValidateRecordSelection = iResult
  Exit Function

ErrorTrap:
  iResult = REC_SEL_VALID_INVALID
  Resume TidyUpAndExit
  
End Function
Public Function ValidateFilter(plngID As Long) As RecordSelectionValidityCodes
  ' Return an integer code representing the validity of the filter.
  ' Return 0 if the filter is OK.
  ' Return 1 if the filter has been deleted by another user.
  ' Return 2 if the filter is hidden, and is owned by the current user.
  ' Return 3 if the filter is hidden, and is NOT owned by the current user.
  ' Return 4 if the filter is no longer valid.
  On Error GoTo ErrorTrap
  
  Dim iResult As RecordSelectionValidityCodes
  Dim rsTemp As ADODB.Recordset
  Dim sSQL As String
  Dim objExpr As clsExprExpression
  
  sSQL = ""
  iResult = REC_SEL_VALID_OK
  
  If plngID > 0 Then
    sSQL = "SELECT access, userName" & _
      " FROM ASRSysExpressions" & _
      " WHERE exprID = " & CStr(plngID)
      
    Set rsTemp = datGeneral.GetReadOnlyRecords(sSQL)
        
    If rsTemp.BOF And rsTemp.EOF Then
      ' Filter no longer exists
      iResult = REC_SEL_VALID_DELETED
    Else
      If (rsTemp!Access = ACCESS_HIDDEN) Or HasHiddenComponents(CLng(plngID)) Then
        If (LCase(Trim(rsTemp!userName)) = LCase(Trim(gsUserName))) Then
          ' Filter is hidden by the current user.
          iResult = REC_SEL_VALID_HIDDENBYUSER
        Else
          ' Filter is hidden by another user.
          iResult = REC_SEL_VALID_HIDDENBYOTHER
        End If
      Else
        'JPD 20031211 Fault 7679 - Agreed with JED that we do not
        ' really need to actually validate the calc/filter definition every time
        ' we load one into a report/utility. We can assume it is valid if
        ' it saved away okay in the first place.
        'Set objExpr = New clsExprExpression
        'With objExpr
        '  .ExpressionID = CLng(plngID)
        '  .ConstructExpression
        '  If (.ValidateExpression(True) <> giEXPRVALIDATION_NOERRORS) Then
        '    iResult = REC_SEL_VALID_INVALID
        '  End If
        'End With
        'Set objExpr = Nothing
      End If
    End If
      
    rsTemp.Close
    Set rsTemp = Nothing
  End If
  
TidyUpAndExit:
  ValidateFilter = iResult
  Exit Function
  
ErrorTrap:
  iResult = REC_SEL_VALID_INVALID
  Resume TidyUpAndExit
  
End Function

Public Function ValidateCalculation(plngID As Long) As RecordSelectionValidityCodes
  ' Return an integer code representing the validity of the Calculation.
  ' Return 0 if the Calculation is OK.
  ' Return 1 if the Calculation has been deleted by another user.
  ' Return 2 if the Calculation is hidden, and is owned by the current user.
  ' Return 3 if the Calculation is hidden, and is NOT owned by the current user.
  ' Return 4 if the Calculation is no longer valid.
  On Error GoTo ErrorTrap
  
  Dim iResult As RecordSelectionValidityCodes
  Dim rsTemp As ADODB.Recordset
  Dim sSQL As String
  Dim objExpr As clsExprExpression
  
  sSQL = ""
  iResult = REC_SEL_VALID_OK
  
  If plngID > 0 Then
    sSQL = "SELECT access, userName" & _
      " FROM ASRSysExpressions" & _
      " WHERE exprID = " & CStr(plngID)
      
    Set rsTemp = datGeneral.GetReadOnlyRecords(sSQL)
        
    If rsTemp.BOF And rsTemp.EOF Then
      ' Filter no longer exists
      iResult = REC_SEL_VALID_DELETED
    Else
      If (rsTemp!Access = ACCESS_HIDDEN) Or HasHiddenComponents(CLng(plngID)) Then
        If (LCase(Trim(rsTemp!userName)) = LCase(Trim(gsUserName))) Then
          ' Calculation is hidden by the current user.
          iResult = REC_SEL_VALID_HIDDENBYUSER
        Else
          ' Calculation is hidden by another user.
          iResult = REC_SEL_VALID_HIDDENBYOTHER
        End If
      Else
        'JPD 20031211 Fault 7679 - Agreed with JED that we do not
        ' really need to actually validate the calc/filter definition every time
        ' we load one into a report/utility. We can assume it is valid if
        ' it saved away okay in the first place.
        'Set objExpr = New clsExprExpression
        'With objExpr
        '  .ExpressionID = CLng(plngID)
        '  .ConstructExpression
        '  If (.ValidateExpression(True) <> giEXPRVALIDATION_NOERRORS) Then
        '    iResult = REC_SEL_VALID_INVALID
        '  End If
        'End With
        'Set objExpr = Nothing
      End If
    End If
      
    rsTemp.Close
    Set rsTemp = Nothing
  End If
  
TidyUpAndExit:
  ValidateCalculation = iResult
  Exit Function
  
ErrorTrap:
  iResult = REC_SEL_VALID_INVALID
  Resume TidyUpAndExit
  
End Function


Public Function ValidatePicklist(plngID As Long) As RecordSelectionValidityCodes
  ' Return an integer code representing the validity of the picklist.
  ' Return 0 if the picklist is OK.
  ' Return 1 if the picklist has been deleted by another user.
  ' Return 2 if the picklist is hidden, and is owned by the current user.
  ' Return 3 if the picklist is hidden, and is NOT owned by the current user.
  ' Return 4 if the picklist is no longer valid.
  On Error GoTo ErrorTrap
  
  Dim iResult As RecordSelectionValidityCodes
  Dim rsTemp As ADODB.Recordset
  Dim sSQL As String
  
  sSQL = ""
  iResult = REC_SEL_VALID_OK
  
  If plngID > 0 Then
    sSQL = "SELECT access, userName" & _
      " FROM ASRSysPickListName" & _
      " WHERE picklistID = " & CStr(plngID)
      
    Set rsTemp = datGeneral.GetReadOnlyRecords(sSQL)
        
    If rsTemp.BOF And rsTemp.EOF Then
      ' Picklist no longer exists
      iResult = REC_SEL_VALID_DELETED
    Else
      If (rsTemp!Access = ACCESS_HIDDEN) Then
        If (LCase(Trim(rsTemp!userName)) = LCase(Trim(gsUserName))) Then
          ' Picklist is hidden by the current user.
          iResult = REC_SEL_VALID_HIDDENBYUSER
        Else
          ' Picklist is hidden by another user.
          iResult = REC_SEL_VALID_HIDDENBYOTHER
        End If
      End If
    End If
      
    rsTemp.Close
    Set rsTemp = Nothing
  End If
  
TidyUpAndExit:
  ValidatePicklist = iResult
  Exit Function
  
ErrorTrap:
  iResult = REC_SEL_VALID_INVALID
  Resume TidyUpAndExit
  
End Function



Public Function AccessDescription(psCode As String) As String
  ' Return the descriptive string associated with the given Access code.
  Select Case psCode
    Case ACCESS_READWRITE
      AccessDescription = ACCESSDESC_READWRITE
    Case ACCESS_READONLY
      AccessDescription = ACCESSDESC_READONLY
    Case ACCESS_HIDDEN
      AccessDescription = ACCESSDESC_HIDDEN
    Case Else
      AccessDescription = ACCESSDESC_UNKNOWN
  End Select
  
End Function

Public Function AccessCode(psDescription As String) As String
  ' Return the descriptive string associated with the given Access code.
  Select Case psDescription
    Case ACCESSDESC_READWRITE
      AccessCode = ACCESS_READWRITE
    Case ACCESSDESC_READONLY
      AccessCode = ACCESS_READONLY
    Case ACCESSDESC_HIDDEN
      AccessCode = ACCESS_HIDDEN
    Case Else
      AccessCode = ACCESS_UNKNOWN
  End Select
  
End Function


