Attribute VB_Name = "modUtilityAccess"
Option Explicit

Public Const ACCESS_READWRITE = "RW"
Public Const ACCESS_READONLY = "RO"
Public Const ACCESS_HIDDEN = "HD"
Public Const ACCESS_UNKNOWN = ""

Public Const ACCESSDESC_READWRITE = "Read / Write"
Public Const ACCESSDESC_READONLY = "Read Only"
Public Const ACCESSDESC_HIDDEN = "Hidden"
Public Const ACCESSDESC_UNKNOWN = "Unknown"

Public gfCurrentUserIsSysSecMgr As Boolean
Public gbUserCanManageLogins As Boolean

Public Function CheckCanMakeHidden(pstrType As String, _
                                   plngID As Long, _
                                   pstrUser As String, _
                                   pstrCaption As String) As Boolean
                                   
  ' Recordset and its source string
  Dim sSQL As String
  
  ' Count and util names of utils which will be made hidden if allowed
  Dim iCount_Owner As Integer
  Dim sDetails_Owner As String
  
  ' count and util names of utils which prevent the change if applicable
  Dim iCount_NotOwner As Integer
  Dim sDetails_NotOwner As String
  
  ' comma separated list of the utility IDs
  Dim sCrossTabIDs As String
  Dim sCustomReportsIDs As String
  Dim sCalendarReportsIDs As String
  Dim sDataTransferIDs As String
  Dim sExportIDs As String
  Dim sGlobalAddIDs As String
  Dim sGlobalUpdateIDs As String
  Dim sGlobalDeleteIDs As String
  Dim sMailMergeIDs As String
  Dim sLabelIDs As String
  Dim sRecordProfileIDs As String
  Dim sMatchReportIDs As String
  Dim sSuccessionPlanningIDs As String
  Dim sCareerProgressionIDs As String
  
  ' batch job info in which utils which require changing are contained
  Dim sBatchJobDetails_Owner As String
  Dim sBatchJobDetails_NotOwner As String
  Dim sBatchJobDetails_ScheduledForOtherUsers As String
  Dim sBatchJobIDs As String
  Dim fBatchJobsOK As Boolean
  Dim sBatchJobScheduledUserGroups As String
  
  Dim sExprIDs As String
  Dim sCalculationIDs As String
  Dim sFilterIDs As String
  Dim sSuperExprIDs As String
  
  fBatchJobsOK = True
  
  'force the username to lowercase for comparisons in the function.
  pstrUser = LCase(pstrUser)
  
  Select Case UCase(pstrType)
  
    '*****************************************************
    ' Calculations/Filters
    '*****************************************************
    Case "E", "F"
      '---------------------------------------------------
      ' Check Calculations/Filters For This Expression
      ' NB. This check must be made before checking the reports/utilities
      '---------------------------------------------------
      sExprIDs = CStr(plngID)
      sSuperExprIDs = GetAllExprRootIDs(plngID)
      
      If sSuperExprIDs <> vbNullString Then
        sExprIDs = sExprIDs & "," & sSuperExprIDs
        
        sSQL = "SELECT ASRSysExpressions.Name," & _
          "   ASRSysExpressions.exprID AS [ID]," & _
          "   ASRSysExpressions.Username," & _
          "   ASRSysExpressions.Access" & _
          " FROM ASRSysExpressions" & _
          " WHERE ASRSysExpressions.exprID IN (" & sSuperExprIDs & ")" & _
          "   AND ASRSysExpressions.type = " & giEXPR_RUNTIMECALCULATION
      
        CheckForPicklistsExpressions utlCalculation, _
          sSQL, _
          pstrUser, _
          iCount_Owner, _
          sDetails_Owner, _
          sCalculationIDs, _
          iCount_NotOwner, _
          sDetails_NotOwner
      
        sSQL = "SELECT ASRSysExpressions.Name," & _
          "   ASRSysExpressions.exprID AS [ID]," & _
          "   ASRSysExpressions.Username," & _
          "   ASRSysExpressions.Access" & _
          " FROM ASRSysExpressions" & _
          " WHERE ASRSysExpressions.exprID IN (" & sSuperExprIDs & ")" & _
          "   AND ASRSysExpressions.type = " & giEXPR_RUNTIMEFILTER
      
        CheckForPicklistsExpressions utlFilter, _
          sSQL, _
          pstrUser, _
          iCount_Owner, _
          sDetails_Owner, _
          sFilterIDs, _
          iCount_NotOwner, _
          sDetails_NotOwner
      End If
  
      '---------------------------------------------------
      ' Check Calendar Reports For This Expression
      '---------------------------------------------------
      sSQL = "SELECT AsrSysCalendarReports.Name," & _
        "   AsrSysCalendarReports.ID," & _
        "   AsrSysCalendarReports.Username," & _
        "   COUNT (ASRSYSCalendarReportAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysCalendarReports" & _
        " LEFT OUTER JOIN ASRSYSCalendarReportEvents ON ASRSysCalendarReports.ID = ASRSYSCalendarReportEvents.calendarReportID " & _
        " LEFT OUTER JOIN ASRSYSCalendarReportAccess ON AsrSysCalendarReports.ID = ASRSYSCalendarReportAccess.ID" & _
        "   AND ASRSYSCalendarReportAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSCalendarReportAccess.groupName NOT IN (SELECT sysusers.name" & _
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
        " WHERE AsrSysCalendarReports.DescriptionExpr IN (" & sExprIDs & ")" & _
        "   OR AsrSysCalendarReports.StartDateExpr IN (" & sExprIDs & ")" & _
        "   OR AsrSysCalendarReports.EndDateExpr IN (" & sExprIDs & ")" & _
        "   OR ASRSysCalendarReports.Filter IN (" & sExprIDs & ")" & _
        "   OR ASRSYSCalendarReportEvents.FilterID IN (" & sExprIDs & ")" & _
        " GROUP BY AsrSysCalendarReports.Name," & _
        "   AsrSysCalendarReports.ID," & _
        "   AsrSysCalendarReports.Username"
      CheckForPicklistsExpressions utlCalendarReport, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sCalendarReportsIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
      
      ' Now check that any of these Calendar Reports are contained within a batch job
      If Len(Trim(sCalendarReportsIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlCalendarReport, _
          sCalendarReportsIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Check Career Progression For This Filter
      '---------------------------------------------------
      sSQL = "SELECT ASRSysMatchReportName.Name," & _
        "   ASRSysMatchReportName.MatchReportID AS [ID]," & _
        "   ASRSysMatchReportName.Username," & _
        "   COUNT (ASRSYSMatchReportAccess.Access) AS [nonHiddenCount]" & _
        " FROM ASRSysMatchReportName" & _
        " LEFT OUTER JOIN ASRSYSMatchReportAccess ON ASRSysMatchReportName.MatchReportID = ASRSYSMatchReportAccess.ID" & _
        "   AND ASRSYSMatchReportAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSMatchReportAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE ASRSysMatchReportName.matchReportType = 2 " & _
        "  AND (ASRSysMatchReportName.table1Filter IN (" & sExprIDs & ")" & _
        "  OR ASRSysMatchReportName.table2Filter IN (" & sExprIDs & "))" & _
        " GROUP BY ASRSysMatchReportName.Name," & _
        "   ASRSysMatchReportName.MatchReportID," & _
        "   ASRSysMatchReportName.Username"
      CheckForPicklistsExpressions utlCareer, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sCareerProgressionIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
      
      ' Now check if any of these Match Reports are contained within a batch job
      If Len(Trim(sCareerProgressionIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlCareer, _
          sCareerProgressionIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Check Cross Tabs For This Expression
      '---------------------------------------------------
      sSQL = "SELECT AsrSysCrossTab.Name," & _
        "   AsrSysCrossTab.[CrossTabID] AS [ID]," & _
        "   AsrSysCrossTab.Username," & _
        "   COUNT (ASRSYSCrossTabAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysCrossTab" & _
        " LEFT OUTER JOIN ASRSYSCrossTabAccess ON AsrSysCrossTab.crossTabID = ASRSYSCrossTabAccess.ID" & _
        "   AND ASRSYSCrossTabAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSCrossTabAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE AsrSysCrossTab.FilterID IN (" & sExprIDs & ")" & _
        " GROUP BY AsrSysCrossTab.Name," & _
        "   AsrSysCrossTab.crossTabID," & _
        "   AsrSysCrossTab.Username"
      CheckForPicklistsExpressions utlCrossTab, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sCrossTabIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
    
      ' Now check that any of these CrossTabs are contained within a batch job
      If Len(Trim(sCrossTabIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlCrossTab, _
          sCrossTabIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Check Custom Reports For This Expression
      '---------------------------------------------------
      sSQL = "SELECT ASRSysCustomReportsName.Name," & _
        "   ASRSysCustomReportsName.ID," & _
        "   ASRSysCustomReportsName.Username," & _
        "   COUNT (ASRSYSCustomReportAccess.Access) AS [nonHiddenCount]" & _
        " FROM ASRSysCustomReportsName" & _
        " LEFT OUTER JOIN ASRSysCustomReportsDetails ON ASRSysCustomReportsName.ID = AsrSysCustomReportsDetails.CustomReportID" & _
        " LEFT OUTER JOIN ASRSYSCustomReportsChildDetails ON ASRSysCustomReportsName.ID = ASRSYSCustomReportsChildDetails.customReportID" & _
        " LEFT OUTER JOIN ASRSYSCustomReportAccess ON ASRSysCustomReportsName.ID = ASRSYSCustomReportAccess.ID" & _
        "   AND ASRSYSCustomReportAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSCustomReportAccess.groupName NOT IN (SELECT sysusers.name" & _
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
        " WHERE ASRSysCustomReportsName.Filter IN (" & sExprIDs & ")" & _
        "   OR ASRSysCustomReportsName.Parent1Filter IN (" & sExprIDs & ")" & _
        "   OR ASRSysCustomReportsName.Parent2Filter IN (" & sExprIDs & ")" & _
        "   OR ASRSYSCustomReportsChildDetails.ChildFilter IN (" & sExprIDs & ")" & _
        "   OR(AsrSysCustomReportsDetails.Type = 'E' " & _
        "     AND AsrSysCustomReportsDetails.ColExprID IN (" & sExprIDs & "))" & _
        " GROUP BY ASRSysCustomReportsName.Name," & _
        "   ASRSysCustomReportsName.ID," & _
        "   ASRSysCustomReportsName.Username"
      
      CheckForPicklistsExpressions utlCustomReport, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sCustomReportsIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
          
      ' Now check that any of these Custom Reports are contained within a batch job
      If Len(Trim(sCustomReportsIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlCustomReport, _
          sCustomReportsIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Check Data Transfer For This Filter
      '---------------------------------------------------
      sSQL = "SELECT AsrSysDataTransferName.Name," & _
        "   AsrSysDataTransferName.DataTransferID AS [ID]," & _
        "   AsrSysDataTransferName.Username," & _
        "   COUNT (ASRSYSDataTransferAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysDataTransferName" & _
        " LEFT OUTER JOIN ASRSYSDataTransferAccess ON AsrSysDataTransferName.DataTransferID = ASRSYSDataTransferAccess.ID" & _
        "   AND ASRSYSDataTransferAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSDataTransferAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE AsrSysDataTransferName.FilterID IN (" & sExprIDs & ")" & _
        " GROUP BY AsrSysDataTransferName.Name," & _
        "   AsrSysDataTransferName.DataTransferID," & _
        "   AsrSysDataTransferName.Username"
      CheckForPicklistsExpressions utlDataTransfer, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sDataTransferIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
                    
      ' Now check that any of these DataTransfers are contained within a batch job
      If Len(Trim(sDataTransferIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlDataTransfer, _
          sDataTransferIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Check Envelopes & Labels For This Expression
      '---------------------------------------------------
      sSQL = "SELECT AsrSysMailMergeName.Name," & _
        "   AsrSysMailMergeName.MailMergeID AS [ID]," & _
        "   AsrSysMailMergeName.Username," & _
        "   COUNT (ASRSYSMailMergeAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysMailMergeName" & _
        " LEFT OUTER JOIN AsrSysMailMergeColumns ON AsrSysMailMergeName.mailMergeID = AsrSysMailMergeColumns.mailMergeID" & _
        " LEFT OUTER JOIN ASRSYSMailMergeAccess ON AsrSysMailMergeName.mailMergeID = ASRSYSMailMergeAccess.ID" & _
        "   AND ASRSYSMailMergeAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSMailMergeAccess.groupName NOT IN (SELECT sysusers.name" & _
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
        " WHERE AsrSysMailMergeName.isLabel = 1" & _
        " AND ((AsrSysMailMergeName.FilterID IN (" & sExprIDs & "))" & _
        "   OR (AsrSysMailMergeColumns.Type = 'E' " & _
        "     AND AsrSysMailMergeColumns.ColumnID IN (" & sExprIDs & ")))" & _
        " GROUP BY AsrSysMailMergeName.Name," & _
        "   AsrSysMailMergeName.MailMergeID," & _
        "   AsrSysMailMergeName.Username"
      CheckForPicklistsExpressions utlLabel, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sLabelIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
        
      ' Now check if any of these Envelopes & Labels are contained within a batch job
      If Len(Trim(sLabelIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlLabel, _
          sLabelIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Check Export For This Expression
      '---------------------------------------------------
      sSQL = "SELECT AsrSysExportName.Name," & _
        "   AsrSysExportName.ID," & _
        "   AsrSysExportName.Username," & _
        "   COUNT (ASRSYSExportAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysExportName" & _
        " LEFT OUTER JOIN AsrSysExportDetails ON AsrSysExportName.ID = AsrSysExportDetails.exportID" & _
        " LEFT OUTER JOIN ASRSYSExportAccess ON AsrSysExportName.ID = ASRSYSExportAccess.ID" & _
        "   AND ASRSYSExportAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSExportAccess.groupName NOT IN (SELECT sysusers.name" & _
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
        " WHERE AsrSysExportName.Filter IN (" & sExprIDs & ")" & _
        "   OR AsrSysExportName.Parent1Filter IN (" & sExprIDs & ")" & _
        "   OR AsrSysExportName.Parent2Filter IN (" & sExprIDs & ")" & _
        "   OR AsrSysExportName.ChildFilter IN (" & sExprIDs & ")" & _
        "   OR (AsrSysExportDetails.Type = 'X' " & _
        "     AND AsrSysExportDetails.ColExprID IN (" & sExprIDs & "))" & _
        " GROUP BY AsrSysExportName.Name," & _
        "   AsrSysExportName.ID," & _
        "   AsrSysExportName.Username"
      CheckForPicklistsExpressions utlExport, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sExportIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
            
      ' Now check that any of these Exports are contained within a batch job
      If Len(Trim(sExportIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlExport, _
          sExportIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Check Global Add For This Expression
      '---------------------------------------------------
      sSQL = "SELECT AsrSysGlobalFunctions.Name," & _
        "   AsrSysGlobalFunctions.functionID AS [ID]," & _
        "   AsrSysGlobalFunctions.Username," & _
        "   COUNT (ASRSYSGlobalAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysGlobalFunctions" & _
        " LEFT OUTER JOIN AsrSysGlobalItems ON AsrSysGlobalFunctions.functionID = AsrSysGlobalItems.FunctionID" & _
        " LEFT OUTER JOIN ASRSYSGlobalAccess ON AsrSysGlobalFunctions.functionID = ASRSYSGlobalAccess.ID" & _
        "   AND ASRSYSGlobalAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSGlobalAccess.groupName NOT IN (SELECT sysusers.name" & _
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
        " WHERE AsrSysGlobalFunctions.Type = 'A' " & _
        "   AND ((AsrSysGlobalFunctions.FilterID IN (" & sExprIDs & "))" & _
        "     OR (AsrSysGlobalItems.ValueType = 4 " & _
        "     AND AsrSysGlobalItems.ExprID IN (" & sExprIDs & ")))" & _
        " GROUP BY AsrSysGlobalFunctions.Name," & _
        "   AsrSysGlobalFunctions.functionID," & _
        "   AsrSysGlobalFunctions.Username"
      CheckForPicklistsExpressions UtlGlobalAdd, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sGlobalAddIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
      
      ' Now check that any of these Global Adds are contained within a batch job
      If Len(Trim(sGlobalAddIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs UtlGlobalAdd, _
          sGlobalAddIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Check Global Update For This Expression
      '---------------------------------------------------
      sSQL = "SELECT AsrSysGlobalFunctions.Name," & _
        "   AsrSysGlobalFunctions.functionID AS [ID]," & _
        "   AsrSysGlobalFunctions.Username," & _
        "   COUNT (ASRSYSGlobalAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysGlobalFunctions" & _
        " LEFT OUTER JOIN AsrSysGlobalItems ON AsrSysGlobalFunctions.functionID = AsrSysGlobalItems.FunctionID" & _
        " LEFT OUTER JOIN ASRSYSGlobalAccess ON AsrSysGlobalFunctions.functionID = ASRSYSGlobalAccess.ID" & _
        "   AND ASRSYSGlobalAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSGlobalAccess.groupName NOT IN (SELECT sysusers.name" & _
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
        " WHERE AsrSysGlobalFunctions.Type = 'U' " & _
        "  AND ((AsrSysGlobalFunctions.FilterID IN (" & sExprIDs & "))" & _
        "    OR (AsrSysGlobalItems.ValueType = 4 " & _
        "      AND AsrSysGlobalItems.ExprID IN (" & sExprIDs & ")))" & _
        " GROUP BY AsrSysGlobalFunctions.Name," & _
        "   AsrSysGlobalFunctions.functionID," & _
        "   AsrSysGlobalFunctions.Username"
      CheckForPicklistsExpressions utlGlobalUpdate, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sGlobalUpdateIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
            
      ' Now check that any of these Global Updates are contained within a batch job
      If Len(Trim(sGlobalUpdateIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlGlobalUpdate, _
          sGlobalUpdateIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If

      '---------------------------------------------------
      ' Check Global Delete For This Filter
      '---------------------------------------------------
      sSQL = "SELECT AsrSysGlobalFunctions.Name," & _
        "   AsrSysGlobalFunctions.functionID AS [ID]," & _
        "   AsrSysGlobalFunctions.Username," & _
        "   COUNT (ASRSYSGlobalAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysGlobalFunctions" & _
        " LEFT OUTER JOIN ASRSYSGlobalAccess ON AsrSysGlobalFunctions.functionID = ASRSYSGlobalAccess.ID" & _
        "   AND ASRSYSGlobalAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSGlobalAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE AsrSysGlobalFunctions.Type = 'D' " & _
        "  AND AsrSysGlobalFunctions.FilterID IN (" & sExprIDs & ")" & _
        " GROUP BY AsrSysGlobalFunctions.Name," & _
        "   AsrSysGlobalFunctions.functionID," & _
        "   AsrSysGlobalFunctions.Username"
      CheckForPicklistsExpressions utlGlobalDelete, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sGlobalDeleteIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
      
      ' Now check that any of these Global Deletes are contained within a batch job
      If Len(Trim(sGlobalDeleteIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlGlobalDelete, _
          sGlobalDeleteIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Check Mail Merge For This Expression
      '---------------------------------------------------
      sSQL = "SELECT AsrSysMailMergeName.Name," & _
        "   AsrSysMailMergeName.MailMergeID AS [ID]," & _
        "   AsrSysMailMergeName.Username," & _
        "   COUNT (ASRSYSMailMergeAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysMailMergeName" & _
        " LEFT OUTER JOIN AsrSysMailMergeColumns ON AsrSysMailMergeName.mailMergeID = AsrSysMailMergeColumns.mailMergeID" & _
        " LEFT OUTER JOIN ASRSYSMailMergeAccess ON AsrSysMailMergeName.mailMergeID = ASRSYSMailMergeAccess.ID" & _
        "   AND ASRSYSMailMergeAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSMailMergeAccess.groupName NOT IN (SELECT sysusers.name" & _
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
        " WHERE AsrSysMailMergeName.isLabel = 0" & _
        "   AND ((AsrSysMailMergeName.FilterID IN (" & sExprIDs & "))" & _
        "     OR (AsrSysMailMergeColumns.Type = 'E' " & _
        "       AND AsrSysMailMergeColumns.ColumnID IN (" & sExprIDs & ")))" & _
        " GROUP BY AsrSysMailMergeName.Name," & _
        "   AsrSysMailMergeName.MailMergeID," & _
        "   AsrSysMailMergeName.Username"
      CheckForPicklistsExpressions utlMailMerge, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sMailMergeIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
        
      ' Now check if any of these Merges are contained within a batch job
      If Len(Trim(sMailMergeIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlMailMerge, _
          sMailMergeIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Check Match Report For This Filter
      '---------------------------------------------------
      sSQL = "SELECT ASRSysMatchReportName.Name," & _
        "   ASRSysMatchReportName.MatchReportID AS [ID]," & _
        "   ASRSysMatchReportName.Username," & _
        "   COUNT (ASRSYSMatchReportAccess.Access) AS [nonHiddenCount]" & _
        " FROM ASRSysMatchReportName" & _
        " LEFT OUTER JOIN ASRSYSMatchReportAccess ON ASRSysMatchReportName.MatchReportID = ASRSYSMatchReportAccess.ID" & _
        "   AND ASRSYSMatchReportAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSMatchReportAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE ASRSysMatchReportName.matchReportType = 0 " & _
        "  AND (ASRSysMatchReportName.table1Filter IN (" & sExprIDs & ")" & _
        "  OR ASRSysMatchReportName.table2Filter IN (" & sExprIDs & "))" & _
        " GROUP BY ASRSysMatchReportName.Name," & _
        "   ASRSysMatchReportName.MatchReportID," & _
        "   ASRSysMatchReportName.Username"
      CheckForPicklistsExpressions utlMatchReport, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sMatchReportIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
      
      ' Now check if any of these Match Reports are contained within a batch job
      If Len(Trim(sMatchReportIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlMatchReport, _
          sMatchReportIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Check Record Profiles For This Filter
      '---------------------------------------------------
      sSQL = "SELECT ASRSysRecordProfileName.Name," & _
        "   ASRSysRecordProfileName.recordProfileID AS [ID]," & _
        "   ASRSysRecordProfileName.Username," & _
        "   COUNT (ASRSYSRecordProfileAccess.Access) AS [nonHiddenCount]" & _
        " FROM ASRSysRecordProfileName" & _
        " LEFT OUTER JOIN ASRSYSRecordProfileTables ON ASRSysRecordProfileName.recordProfileID = ASRSYSRecordProfileTables.recordProfileID" & _
        " LEFT OUTER JOIN ASRSYSRecordProfileAccess ON ASRSysRecordProfileName.recordProfileID = ASRSYSRecordProfileAccess.ID" & _
        "   AND ASRSYSRecordProfileAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSRecordProfileAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE ASRSysRecordProfileName.FilterID IN (" & sExprIDs & ")" & _
        "   OR ASRSYSRecordProfileTables.FilterID IN (" & sExprIDs & ")" & _
        " GROUP BY ASRSysRecordProfileName.Name," & _
        "   ASRSysRecordProfileName.recordProfileID," & _
        "   ASRSysRecordProfileName.Username"
      CheckForPicklistsExpressions utlRecordProfile, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sRecordProfileIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner

      ' Now check that any of these Record Profiles are contained within a batch job
      If Len(Trim(sRecordProfileIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlRecordProfile, _
          sRecordProfileIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Check Succession Planning For This Filter
      '---------------------------------------------------
      sSQL = "SELECT ASRSysMatchReportName.Name," & _
        "   ASRSysMatchReportName.MatchReportID AS [ID]," & _
        "   ASRSysMatchReportName.Username," & _
        "   COUNT (ASRSYSMatchReportAccess.Access) AS [nonHiddenCount]" & _
        " FROM ASRSysMatchReportName" & _
        " LEFT OUTER JOIN ASRSYSMatchReportAccess ON ASRSysMatchReportName.MatchReportID = ASRSYSMatchReportAccess.ID" & _
        "   AND ASRSYSMatchReportAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSMatchReportAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE ASRSysMatchReportName.matchReportType = 1 " & _
        "  AND (ASRSysMatchReportName.table1Filter IN (" & sExprIDs & ")" & _
        "  OR ASRSysMatchReportName.table2Filter IN (" & sExprIDs & "))" & _
        " GROUP BY ASRSysMatchReportName.Name," & _
        "   ASRSysMatchReportName.MatchReportID," & _
        "   ASRSysMatchReportName.Username"
      
      CheckForPicklistsExpressions utlSuccession, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sSuccessionPlanningIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
      
      ' Now check if any of these Match Reports are contained within a batch job
      If Len(Trim(sSuccessionPlanningIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlSuccession, _
          sSuccessionPlanningIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Ok, all relevant utility definitions have now been checked, so check
      ' the counts and act accordingly
      '---------------------------------------------------
      If (iCount_Owner = 0) And _
        (iCount_NotOwner = 0) And _
        fBatchJobsOK And _
        (Len(sBatchJobDetails_Owner) = 0) Then
          
        CheckCanMakeHidden = True
        Exit Function
      
      ElseIf (iCount_Owner > 0) And _
        (iCount_NotOwner = 0) And _
        fBatchJobsOK Then
        ' Can change utils and no utils
        ' are contained within batch jobs
        ' that cant be changed
        If MsgBox("Changing the selected " & IIf(UCase(pstrType) = "F", "filter", "calculation") & " to hidden will automatically" & vbCrLf & _
                  "make the following definition(s), of which you are the" & vbCrLf & _
                  "owner, hidden also:" & vbCrLf & vbCrLf & _
                  sDetails_Owner & sBatchJobDetails_Owner & vbCrLf & _
                  "Do you wish to continue ?", vbQuestion + vbYesNo, IIf(Len(pstrCaption) = 0, "HR Pro - Data Manager", pstrCaption)) _
                  = vbNo Then
          Screen.MousePointer = vbNormal
          CheckCanMakeHidden = False
          Exit Function
        
        Else
          ' Ok, we are continuing, so lets update all those utils to hidden !
          
          ' Calculations
          If Len(Trim(sCalculationIDs)) > 0 Then
            HideUtilities utlCalculation, sCalculationIDs
            Call UtilUpdateLastSavedMultiple(utlCalculation, sCalculationIDs)
          End If
                    
          ' Filters
          If Len(Trim(sFilterIDs)) > 0 Then
            HideUtilities utlFilter, sFilterIDs
            Call UtilUpdateLastSavedMultiple(utlFilter, sFilterIDs)
          End If
          
          ' Batch Jobs
          If Len(Trim(sBatchJobIDs)) > 0 Then
            HideUtilities utlBatchJob, sBatchJobIDs
            Call UtilUpdateLastSavedMultiple(utlBatchJob, sBatchJobIDs)
          End If
          
          ' Calendar Reports
          If Len(Trim(sCalendarReportsIDs)) > 0 Then
            HideUtilities utlCalendarReport, sCalendarReportsIDs
            Call UtilUpdateLastSavedMultiple(utlCalendarReport, sCalendarReportsIDs)
          End If
          
          ' Career Progression
          If Len(Trim(sCareerProgressionIDs)) > 0 Then
            HideUtilities utlCareer, sCareerProgressionIDs
            Call UtilUpdateLastSavedMultiple(utlCareer, sCareerProgressionIDs)
          End If
          
          ' Cross Tabs
          If Len(Trim(sCrossTabIDs)) > 0 Then
            HideUtilities utlCrossTab, sCrossTabIDs
            Call UtilUpdateLastSavedMultiple(utlCrossTab, sCrossTabIDs)
          End If
          
          ' Custom Reports
          If Len(Trim(sCustomReportsIDs)) > 0 Then
            HideUtilities utlCustomReport, sCustomReportsIDs
            Call UtilUpdateLastSavedMultiple(utlCustomReport, sCustomReportsIDs)
          End If
                   
          ' Data Transfer
          If Len(Trim(sDataTransferIDs)) > 0 Then
            HideUtilities utlDataTransfer, sDataTransferIDs
            Call UtilUpdateLastSavedMultiple(utlDataTransfer, sDataTransferIDs)
          End If
          
          ' Envelopes & Labels
          If Len(Trim(sLabelIDs)) > 0 Then
            HideUtilities utlLabel, sLabelIDs
            Call UtilUpdateLastSavedMultiple(utlLabel, sLabelIDs)
          End If
          
          ' Export
          If Len(Trim(sExportIDs)) > 0 Then
            HideUtilities utlExport, sExportIDs
            Call UtilUpdateLastSavedMultiple(utlExport, sExportIDs)
          End If
                             
          ' Global Add
          If Len(Trim(sGlobalAddIDs)) > 0 Then
            HideUtilities UtlGlobalAdd, sGlobalAddIDs
            Call UtilUpdateLastSavedMultiple(UtlGlobalAdd, sGlobalAddIDs)
          End If
          
          ' Global Update
          If Len(Trim(sGlobalUpdateIDs)) > 0 Then
            HideUtilities utlGlobalUpdate, sGlobalUpdateIDs
            Call UtilUpdateLastSavedMultiple(utlGlobalUpdate, sGlobalUpdateIDs)
          End If
          
          ' Global Delete
          If Len(Trim(sGlobalDeleteIDs)) > 0 Then
            HideUtilities utlGlobalDelete, sGlobalDeleteIDs
            Call UtilUpdateLastSavedMultiple(utlGlobalDelete, sGlobalDeleteIDs)
          End If
          
          ' Mail Merge
          If Len(Trim(sMailMergeIDs)) > 0 Then
            HideUtilities utlMailMerge, sMailMergeIDs
            Call UtilUpdateLastSavedMultiple(utlMailMerge, sMailMergeIDs)
          End If
          
          ' Match Reports
          If Len(Trim(sMatchReportIDs)) > 0 Then
            HideUtilities utlMatchReport, sMatchReportIDs
            Call UtilUpdateLastSavedMultiple(utlMatchReport, sMatchReportIDs)
          End If
          
          ' Record Profile
          If Len(Trim(sRecordProfileIDs)) > 0 Then
            HideUtilities utlRecordProfile, sRecordProfileIDs
            Call UtilUpdateLastSavedMultiple(utlRecordProfile, sRecordProfileIDs)
          End If
          
          ' Succession Planning
          If Len(Trim(sSuccessionPlanningIDs)) > 0 Then
            HideUtilities utlSuccession, sSuccessionPlanningIDs
            Call UtilUpdateLastSavedMultiple(utlSuccession, sSuccessionPlanningIDs)
          End If
          
          ' Ok, all done, so exit now
          CheckCanMakeHidden = True
          Exit Function
        
        End If
      
      ElseIf (iCount_Owner > 0) And _
        (iCount_NotOwner = 0) And _
        (Not fBatchJobsOK) Then
        ' Can change utils but abort cos those
        ' utils are in batch jobs which cannot
        ' be changed
        If Len(sBatchJobDetails_ScheduledForOtherUsers) > 0 Then
          MsgBox "This " & IIf(UCase(pstrType) = "F", "filter", "calculation") & " cannot be made hidden as it is used in " & vbCrLf & _
                 "definition(s) which are included in the following" & vbCrLf & _
                 "batch jobs which are scheduled to be run by other user groups :" & vbCrLf & vbCrLf & sBatchJobDetails_ScheduledForOtherUsers, vbExclamation + vbOKOnly _
                 , IIf(Len(pstrCaption) = 0, "HR Pro - Data Manager", pstrCaption)
        Else
          MsgBox "This " & IIf(UCase(pstrType) = "F", "filter", "calculation") & " cannot be made hidden as it is used in " & vbCrLf & _
                 "definition(s) which are included in the following" & vbCrLf & _
                 "batch jobs of which you are not the owner :" & vbCrLf & vbCrLf & sBatchJobDetails_NotOwner, vbExclamation + vbOKOnly _
                 , IIf(Len(pstrCaption) = 0, "HR Pro - Data Manager", pstrCaption)
        End If

        Screen.MousePointer = vbNormal
        CheckCanMakeHidden = False
        Exit Function

      ElseIf (iCount_NotOwner > 0) Then            ' Cannot change utils
      
        MsgBox "This " & IIf(UCase(pstrType) = "F", "filter", "calculation") & " cannot be made hidden as it is used in the" & vbCrLf & _
               "following definition(s), of which you are not the" & vbCrLf & _
               "owner :" & vbCrLf & vbCrLf & sDetails_NotOwner, _
               vbExclamation + vbOKOnly, IIf(Len(pstrCaption) = 0, "HR Pro - Data Manager", pstrCaption)
        Screen.MousePointer = vbNormal
        CheckCanMakeHidden = False
        Exit Function
      
      End If
    
    '*****************************************************
    Case "P"
    '*****************************************************
      '---------------------------------------------------
      ' Check Cross Tabs For This Picklist
      '---------------------------------------------------
      sSQL = "SELECT AsrSysCrossTab.Name," & _
        "   AsrSysCrossTab.[CrossTabID] AS [ID]," & _
        "   AsrSysCrossTab.Username," & _
        "   COUNT (ASRSYSCrossTabAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysCrossTab" & _
        " LEFT OUTER JOIN ASRSYSCrossTabAccess ON AsrSysCrossTab.crossTabID = ASRSYSCrossTabAccess.ID" & _
        "   AND ASRSYSCrossTabAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSCrossTabAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE AsrSysCrossTab.PicklistID = " & CStr(plngID) & _
        " GROUP BY AsrSysCrossTab.Name," & _
        "   AsrSysCrossTab.crossTabID," & _
        "   AsrSysCrossTab.Username"
      CheckForPicklistsExpressions utlCrossTab, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sCrossTabIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
      
      ' Now check that any of these CrossTabs are contained within a batch job
      If Len(Trim(sCrossTabIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlCrossTab, _
          sCrossTabIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If

      '---------------------------------------------------
      ' Check Custom Reports For This Picklist
      '---------------------------------------------------
      sSQL = "SELECT ASRSysCustomReportsName.Name," & _
        "   ASRSysCustomReportsName.ID," & _
        "   ASRSysCustomReportsName.Username," & _
        "   COUNT (ASRSYSCustomReportAccess.Access) AS [nonHiddenCount]" & _
        " FROM ASRSysCustomReportsName" & _
        " LEFT OUTER JOIN ASRSYSCustomReportAccess ON ASRSysCustomReportsName.ID = ASRSYSCustomReportAccess.ID" & _
        "   AND ASRSYSCustomReportAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSCustomReportAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE ASRSysCustomReportsName.Picklist = " & CStr(plngID) & _
        "   OR ASRSysCustomReportsName.Parent1Picklist = " & CStr(plngID) & _
        "   OR ASRSysCustomReportsName.Parent2Picklist = " & CStr(plngID) & _
        " GROUP BY ASRSysCustomReportsName.Name," & _
        "   ASRSysCustomReportsName.ID," & _
        "   ASRSysCustomReportsName.Username"
      
      CheckForPicklistsExpressions utlCustomReport, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sCustomReportsIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner

      ' Now check that any of these Custom Reports are contained within a batch job
      If Len(Trim(sCustomReportsIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlCustomReport, _
          sCustomReportsIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If

      '---------------------------------------------------
      ' Check Calendar Reports For This Picklist
      '---------------------------------------------------
      sSQL = "SELECT ASRSysCalendarReports.Name," & _
        "   ASRSysCalendarReports.ID," & _
        "   ASRSysCalendarReports.Username," & _
        "   COUNT (ASRSYSCalendarReportAccess.Access) AS [nonHiddenCount]" & _
        " FROM ASRSysCalendarReports" & _
        " LEFT OUTER JOIN ASRSYSCalendarReportAccess ON ASRSysCalendarReports.ID = ASRSYSCalendarReportAccess.ID" & _
        "   AND ASRSYSCalendarReportAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSCalendarReportAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE ASRSysCalendarReports.Picklist = " & CStr(plngID) & _
        " GROUP BY ASRSysCalendarReports.Name," & _
        "   ASRSysCalendarReports.ID," & _
        "   ASRSysCalendarReports.Username"
                
      CheckForPicklistsExpressions utlCalendarReport, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sCalendarReportsIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner

      ' Now check that any of these Custom Reports are contained within a batch job
      If Len(Trim(sCalendarReportsIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlCalendarReport, _
          sCalendarReportsIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If

      '---------------------------------------------------
      ' Check Record Profile For This Picklist
      '---------------------------------------------------
      sSQL = "SELECT ASRSysRecordProfileName.Name," & _
        "   ASRSysRecordProfileName.recordProfileID AS [ID]," & _
        "   ASRSysRecordProfileName.Username," & _
        "   COUNT (ASRSYSRecordProfileAccess.Access) AS [nonHiddenCount]" & _
        " FROM ASRSysRecordProfileName" & _
        " LEFT OUTER JOIN ASRSYSRecordProfileAccess ON ASRSysRecordProfileName.recordProfileID = ASRSYSRecordProfileAccess.ID" & _
        "   AND ASRSYSRecordProfileAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSRecordProfileAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE ASRSysRecordProfileName.PicklistID = " & CStr(plngID) & _
        " GROUP BY ASRSysRecordProfileName.Name," & _
        "   ASRSysRecordProfileName.recordProfileID," & _
        "   ASRSysRecordProfileName.Username"
      CheckForPicklistsExpressions utlRecordProfile, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sRecordProfileIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
    
      ' Now check that any of these Record Profiles are contained within a batch job
      If Len(Trim(sRecordProfileIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlRecordProfile, _
          sRecordProfileIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If

      '---------------------------------------------------
      ' Check Data Transfer For This Picklist
      '---------------------------------------------------
      sSQL = "SELECT AsrSysDataTransferName.Name," & _
        "   AsrSysDataTransferName.DataTransferID AS [ID]," & _
        "   AsrSysDataTransferName.Username," & _
        "   COUNT (ASRSYSDataTransferAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysDataTransferName" & _
        " LEFT OUTER JOIN ASRSYSDataTransferAccess ON AsrSysDataTransferName.DataTransferID = ASRSYSDataTransferAccess.ID" & _
        "   AND ASRSYSDataTransferAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSDataTransferAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE AsrSysDataTransferName.PicklistID = " & CStr(plngID) & _
        " GROUP BY AsrSysDataTransferName.Name," & _
        "   AsrSysDataTransferName.DataTransferID," & _
        "   AsrSysDataTransferName.Username"
      
      
      CheckForPicklistsExpressions utlDataTransfer, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sDataTransferIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner

      ' Now check that any of these DataTransfers are contained within a batch job
      If Len(Trim(sDataTransferIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlDataTransfer, _
          sDataTransferIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If

      '---------------------------------------------------
      ' Check Export For This Picklist
      '---------------------------------------------------
      sSQL = "SELECT AsrSysExportName.Name," & _
        "   AsrSysExportName.ID," & _
        "   AsrSysExportName.Username," & _
        "   COUNT (ASRSYSExportAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysExportName" & _
        " LEFT OUTER JOIN ASRSYSExportAccess ON AsrSysExportName.ID = ASRSYSExportAccess.ID" & _
        "   AND ASRSYSExportAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSExportAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE AsrSysExportName.Picklist = " & CStr(plngID) & _
        "   OR AsrSysExportName.Parent1Picklist = " & CStr(plngID) & _
        "   OR AsrSysExportName.Parent2Picklist = " & CStr(plngID) & _
        " GROUP BY AsrSysExportName.Name," & _
        "   AsrSysExportName.ID," & _
        "   AsrSysExportName.Username"
      
      CheckForPicklistsExpressions utlExport, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sExportIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner

      ' Now check that any of these Exports are contained within a batch job
      If Len(Trim(sExportIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlExport, _
          sExportIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If

      '---------------------------------------------------
      ' Check Global Add For This Picklist
      '---------------------------------------------------
      sSQL = "SELECT AsrSysGlobalFunctions.Name," & _
        "   AsrSysGlobalFunctions.functionID AS [ID]," & _
        "   AsrSysGlobalFunctions.Username," & _
        "   COUNT (ASRSYSGlobalAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysGlobalFunctions" & _
        " LEFT OUTER JOIN ASRSYSGlobalAccess ON AsrSysGlobalFunctions.functionID = ASRSYSGlobalAccess.ID" & _
        "   AND ASRSYSGlobalAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSGlobalAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE AsrSysGlobalFunctions.Type = 'A' " & _
        "  AND AsrSysGlobalFunctions.PicklistID = " & CStr(plngID) & _
        " GROUP BY AsrSysGlobalFunctions.Name," & _
        "   AsrSysGlobalFunctions.functionID," & _
        "   AsrSysGlobalFunctions.Username"
      
      CheckForPicklistsExpressions UtlGlobalAdd, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sGlobalAddIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner

      ' Now check that any of these Global Adds are contained within a batch job
      If Len(Trim(sGlobalAddIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs UtlGlobalAdd, _
          sGlobalAddIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If

      '---------------------------------------------------
      ' Check Global Update For This Picklist
      '---------------------------------------------------
      sSQL = "SELECT AsrSysGlobalFunctions.Name," & _
        "   AsrSysGlobalFunctions.functionID AS [ID]," & _
        "   AsrSysGlobalFunctions.Username," & _
        "   COUNT (ASRSYSGlobalAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysGlobalFunctions" & _
        " LEFT OUTER JOIN ASRSYSGlobalAccess ON AsrSysGlobalFunctions.functionID = ASRSYSGlobalAccess.ID" & _
        "   AND ASRSYSGlobalAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSGlobalAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE AsrSysGlobalFunctions.Type = 'U' " & _
        "  AND AsrSysGlobalFunctions.PicklistID = " & CStr(plngID) & _
        " GROUP BY AsrSysGlobalFunctions.Name," & _
        "   AsrSysGlobalFunctions.functionID," & _
        "   AsrSysGlobalFunctions.Username"
      
      CheckForPicklistsExpressions utlGlobalUpdate, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sGlobalUpdateIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner

      ' Now check that any of these Global Updates are contained within a batch job
      If Len(Trim(sGlobalUpdateIDs)) > 0 Then
        ' JPD20011219 Fault 3303
        CheckCanMakeHiddenInBatchJobs utlGlobalUpdate, _
          sGlobalUpdateIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If

      '---------------------------------------------------
      ' Check Global Delete For This Picklist
      '---------------------------------------------------
      sSQL = "SELECT AsrSysGlobalFunctions.Name," & _
        "   AsrSysGlobalFunctions.functionID AS [ID]," & _
        "   AsrSysGlobalFunctions.Username," & _
        "   COUNT (ASRSYSGlobalAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysGlobalFunctions" & _
        " LEFT OUTER JOIN ASRSYSGlobalAccess ON AsrSysGlobalFunctions.functionID = ASRSYSGlobalAccess.ID" & _
        "   AND ASRSYSGlobalAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSGlobalAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE AsrSysGlobalFunctions.Type = 'D' " & _
        "  AND AsrSysGlobalFunctions.PicklistID = " & CStr(plngID) & _
        " GROUP BY AsrSysGlobalFunctions.Name," & _
        "   AsrSysGlobalFunctions.functionID," & _
        "   AsrSysGlobalFunctions.Username"
      
      CheckForPicklistsExpressions utlGlobalDelete, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sGlobalDeleteIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
      
      ' Now check that any of these Global Deletes are contained within a batch job
      If Len(Trim(sGlobalDeleteIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlGlobalDelete, _
          sGlobalDeleteIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If

      '---------------------------------------------------
      ' Check Mail Merge For This Picklist
      '---------------------------------------------------
      sSQL = "SELECT AsrSysMailMergeName.Name," & _
        "   AsrSysMailMergeName.MailMergeID AS [ID]," & _
        "   AsrSysMailMergeName.Username," & _
        "   COUNT (ASRSYSMailMergeAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysMailMergeName" & _
        " LEFT OUTER JOIN ASRSYSMailMergeAccess ON AsrSysMailMergeName.mailMergeID = ASRSYSMailMergeAccess.ID" & _
        "   AND ASRSYSMailMergeAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSMailMergeAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE AsrSysMailMergeName.PicklistID = " & plngID & _
        "   AND AsrSysMailMergeName.isLabel = 0" & _
        " GROUP BY AsrSysMailMergeName.Name," & _
        "   AsrSysMailMergeName.MailMergeID," & _
        "   AsrSysMailMergeName.Username"
      
      CheckForPicklistsExpressions utlMailMerge, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sMailMergeIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner

      ' Now check if any of these Merges are contained within a batch job
      If Len(Trim(sMailMergeIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlMailMerge, _
          sMailMergeIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If

      '---------------------------------------------------
      ' Check Envelopes & Labels For This Picklist
      '---------------------------------------------------
      sSQL = "SELECT AsrSysMailMergeName.Name," & _
        "   AsrSysMailMergeName.MailMergeID AS [ID]," & _
        "   AsrSysMailMergeName.Username," & _
        "   COUNT (ASRSYSMailMergeAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysMailMergeName" & _
        " LEFT OUTER JOIN ASRSYSMailMergeAccess ON AsrSysMailMergeName.mailMergeID = ASRSYSMailMergeAccess.ID" & _
        "   AND ASRSYSMailMergeAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSMailMergeAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE AsrSysMailMergeName.PicklistID = " & plngID & _
        "   AND AsrSysMailMergeName.isLabel = 1" & _
        " GROUP BY AsrSysMailMergeName.Name," & _
        "   AsrSysMailMergeName.MailMergeID," & _
        "   AsrSysMailMergeName.Username"
      
      CheckForPicklistsExpressions utlLabel, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sLabelIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner

      ' Now check if any of these Merges are contained within a batch job
      If Len(Trim(sLabelIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlLabel, _
          sLabelIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If

      '---------------------------------------------------
      ' Check Match Report For This Picklist
      '---------------------------------------------------
      sSQL = "SELECT ASRSysMatchReportName.Name," & _
        "   ASRSysMatchReportName.MatchReportID AS [ID]," & _
        "   ASRSysMatchReportName.Username," & _
        "   COUNT (ASRSYSMatchReportAccess.Access) AS [nonHiddenCount]" & _
        " FROM ASRSysMatchReportName" & _
        " LEFT OUTER JOIN ASRSYSMatchReportAccess ON ASRSysMatchReportName.MatchReportID = ASRSYSMatchReportAccess.ID" & _
        "   AND ASRSYSMatchReportAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSMatchReportAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE ASRSysMatchReportName.matchReportType = 0 " & _
        "  AND (ASRSysMatchReportName.table1Picklist = " & CStr(plngID) & _
        "  OR ASRSysMatchReportName.table2Picklist = " & CStr(plngID) & ")" & _
        " GROUP BY ASRSysMatchReportName.Name," & _
        "   ASRSysMatchReportName.MatchReportID," & _
        "   ASRSysMatchReportName.Username"
      CheckForPicklistsExpressions utlMatchReport, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sMatchReportIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
      
      ' Now check if any of these Match Reports are contained within a batch job
      If Len(Trim(sMatchReportIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlMatchReport, _
          sMatchReportIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Check Succession Planning For This Filter
      '---------------------------------------------------
      sSQL = "SELECT ASRSysMatchReportName.Name," & _
        "   ASRSysMatchReportName.MatchReportID AS [ID]," & _
        "   ASRSysMatchReportName.Username," & _
        "   COUNT (ASRSYSMatchReportAccess.Access) AS [nonHiddenCount]" & _
        " FROM ASRSysMatchReportName" & _
        " LEFT OUTER JOIN ASRSYSMatchReportAccess ON ASRSysMatchReportName.MatchReportID = ASRSYSMatchReportAccess.ID" & _
        "   AND ASRSYSMatchReportAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSMatchReportAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE ASRSysMatchReportName.matchReportType = 1 " & _
        "  AND (ASRSysMatchReportName.table1Picklist = " & CStr(plngID) & _
        "  OR ASRSysMatchReportName.table2Picklist = " & CStr(plngID) & ")" & _
        " GROUP BY ASRSysMatchReportName.Name," & _
        "   ASRSysMatchReportName.MatchReportID," & _
        "   ASRSysMatchReportName.Username"
      CheckForPicklistsExpressions utlSuccession, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sSuccessionPlanningIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
      
      ' Now check if any of these Match Reports are contained within a batch job
      If Len(Trim(sSuccessionPlanningIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlSuccession, _
          sSuccessionPlanningIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Check Career Progression For This Filter
      '---------------------------------------------------
      sSQL = "SELECT ASRSysMatchReportName.Name," & _
        "   ASRSysMatchReportName.MatchReportID AS [ID]," & _
        "   ASRSysMatchReportName.Username," & _
        "   COUNT (ASRSYSMatchReportAccess.Access) AS [nonHiddenCount]" & _
        " FROM ASRSysMatchReportName" & _
        " LEFT OUTER JOIN ASRSYSMatchReportAccess ON ASRSysMatchReportName.MatchReportID = ASRSYSMatchReportAccess.ID" & _
        "   AND ASRSYSMatchReportAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSMatchReportAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE ASRSysMatchReportName.matchReportType = 2 " & _
        "  AND (ASRSysMatchReportName.table1Picklist = " & CStr(plngID) & _
        "  OR ASRSysMatchReportName.table2Picklist = " & CStr(plngID) & ")" & _
        " GROUP BY ASRSysMatchReportName.Name," & _
        "   ASRSysMatchReportName.MatchReportID," & _
        "   ASRSysMatchReportName.Username"
      CheckForPicklistsExpressions utlCareer, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sCareerProgressionIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
      
      ' Now check if any of these Match Reports are contained within a batch job
      If Len(Trim(sCareerProgressionIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlCareer, _
          sCareerProgressionIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Ok, all relevant utility definitions have now been checked, so check
      ' the counts and act accordingly
      '---------------------------------------------------
      If (iCount_Owner = 0) And _
        (iCount_NotOwner = 0) And _
        fBatchJobsOK And _
        (Len(sBatchJobDetails_Owner) = 0) Then
          
        CheckCanMakeHidden = True
        Exit Function
      
      ElseIf (iCount_Owner > 0) And _
        (iCount_NotOwner = 0) And _
        fBatchJobsOK Then
        ' Can change utils and no utils
        ' are contained within batch jobs
        ' that cant be changed
        If MsgBox("Changing the selected picklist to hidden will automatically" & vbCrLf & _
                  "make the following definition(s), of which you are the" & vbCrLf & _
                  "owner, hidden also:" & vbCrLf & vbCrLf & _
                  sDetails_Owner & sBatchJobDetails_Owner & vbCrLf & _
                  "Do you wish to continue ?", vbQuestion + vbYesNo, IIf(Len(pstrCaption) = 0, "HR Pro - Data Manager", pstrCaption)) _
                  = vbNo Then
          Screen.MousePointer = vbNormal
          CheckCanMakeHidden = False
          Exit Function
        
        Else
          ' Ok, we are continuing, so lets update all those utils to hidden !
          
          ' Cross Tabs
          If Len(Trim(sCrossTabIDs)) > 0 Then
            HideUtilities utlCrossTab, sCrossTabIDs
            Call UtilUpdateLastSavedMultiple(utlCrossTab, sCrossTabIDs)
          End If

          ' Custom Reports
          If Len(Trim(sCustomReportsIDs)) > 0 Then
            HideUtilities utlCustomReport, sCustomReportsIDs
            Call UtilUpdateLastSavedMultiple(utlCustomReport, sCustomReportsIDs)
          End If

          ' Calendar Reports
          If Len(Trim(sCalendarReportsIDs)) > 0 Then
            HideUtilities utlCalendarReport, sCalendarReportsIDs
            Call UtilUpdateLastSavedMultiple(utlCalendarReport, sCalendarReportsIDs)
          End If

          ' Record Profile
          If Len(Trim(sRecordProfileIDs)) > 0 Then
            HideUtilities utlRecordProfile, sRecordProfileIDs
            Call UtilUpdateLastSavedMultiple(utlRecordProfile, sRecordProfileIDs)
          End If

          ' Data Transfer
          If Len(Trim(sDataTransferIDs)) > 0 Then
            HideUtilities utlDataTransfer, sDataTransferIDs
            Call UtilUpdateLastSavedMultiple(utlDataTransfer, sDataTransferIDs)
          End If

          ' Export
          If Len(Trim(sExportIDs)) > 0 Then
            HideUtilities utlExport, sExportIDs
            Call UtilUpdateLastSavedMultiple(utlExport, sExportIDs)
          End If

          ' Global Add
          If Len(Trim(sGlobalAddIDs)) > 0 Then
            HideUtilities UtlGlobalAdd, sGlobalAddIDs
            Call UtilUpdateLastSavedMultiple(UtlGlobalAdd, sGlobalAddIDs)
          End If

          ' Global Update
          If Len(Trim(sGlobalUpdateIDs)) > 0 Then
            HideUtilities utlGlobalUpdate, sGlobalUpdateIDs
            Call UtilUpdateLastSavedMultiple(utlGlobalUpdate, sGlobalUpdateIDs)
          End If

          ' Global Delete
          If Len(Trim(sGlobalDeleteIDs)) > 0 Then
            HideUtilities utlGlobalDelete, sGlobalDeleteIDs
            Call UtilUpdateLastSavedMultiple(utlGlobalDelete, sGlobalDeleteIDs)
          End If

          ' Mail Merge
          If Len(Trim(sMailMergeIDs)) > 0 Then
            HideUtilities utlMailMerge, sMailMergeIDs
            Call UtilUpdateLastSavedMultiple(utlMailMerge, sMailMergeIDs)
          End If
          
          ' Envelopes & Labels
          If Len(Trim(sLabelIDs)) > 0 Then
            HideUtilities utlLabel, sLabelIDs
            Call UtilUpdateLastSavedMultiple(utlLabel, sLabelIDs)
          End If
          
          ' Match Reports
          If Len(Trim(sMatchReportIDs)) > 0 Then
            HideUtilities utlMatchReport, sMatchReportIDs
            Call UtilUpdateLastSavedMultiple(utlMatchReport, sMatchReportIDs)
          End If
          
          ' Succession Planning
          If Len(Trim(sSuccessionPlanningIDs)) > 0 Then
            HideUtilities utlSuccession, sSuccessionPlanningIDs
            Call UtilUpdateLastSavedMultiple(utlSuccession, sSuccessionPlanningIDs)
          End If
          
          ' Career Progression
          If Len(Trim(sCareerProgressionIDs)) > 0 Then
            HideUtilities utlCareer, sCareerProgressionIDs
            Call UtilUpdateLastSavedMultiple(utlCareer, sCareerProgressionIDs)
          End If
          
          ' Batch Jobs
          If Len(Trim(sBatchJobIDs)) > 0 Then
            HideUtilities utlBatchJob, sBatchJobIDs
            Call UtilUpdateLastSavedMultiple(utlBatchJob, sBatchJobIDs)
          End If
          
          ' Ok, all done, so exit now
          CheckCanMakeHidden = True
          Exit Function
        End If
      
      ElseIf (iCount_Owner > 0) And _
        (iCount_NotOwner = 0) And _
        (Not fBatchJobsOK) Then
        ' Can change utils but abort cos those
        ' utils are in batch jobs which cannot
        ' be changed
        If Len(sBatchJobDetails_ScheduledForOtherUsers) > 0 Then
          MsgBox "This expression cannot be made hidden as it is used in " & vbCrLf & _
                 "definition(s) which are included in the following" & vbCrLf & _
                 "batch jobs which are scheduled to be run by other user groups :" & vbCrLf & vbCrLf & sBatchJobDetails_ScheduledForOtherUsers, vbExclamation + vbOKOnly _
                 , IIf(Len(pstrCaption) = 0, "HR Pro - Data Manager", pstrCaption)
        Else
          MsgBox "This picklist cannot be made hidden as it is used in " & vbCrLf & _
                 "definition(s) which are included in the following" & vbCrLf & _
                 "batch jobs of which you are not the owner :" & vbCrLf & vbCrLf & sBatchJobDetails_NotOwner, vbExclamation + vbOKOnly _
                 , IIf(Len(pstrCaption) = 0, "HR Pro - Data Manager", pstrCaption)
        End If
        Screen.MousePointer = vbNormal
        CheckCanMakeHidden = False
        Exit Function

      ElseIf (iCount_NotOwner > 0) Then
        ' Cannot change utils
        MsgBox "This picklist cannot be made hidden as it is used in the" & vbCrLf & _
               "following definition(s), of which you are not the" & vbCrLf & _
               "owner :" & vbCrLf & vbCrLf & sDetails_NotOwner, _
               vbExclamation + vbOKOnly, IIf(Len(pstrCaption) = 0, "HR Pro - Data Manager", pstrCaption)
        Screen.MousePointer = vbNormal
        CheckCanMakeHidden = False
        Exit Function
      
      End If
    
  End Select
  
End Function



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
  Dim rsTemp As New ADODB.Recordset
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
  
    rsTemp.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
        
    Do Until rsTemp.EOF
      If LCase(rsTemp!UserName) = pstrUser Then
        ' Found a Batch Job whose owner is the same
        If (rsTemp!scheduled = 1) And _
          (Len(rsTemp!RoleToPrompt) > 0) And _
          (UCase(rsTemp!RoleToPrompt) <> UCase(gsUserGroup)) And _
          (fHiddenToAllGroups Or (InStr(sHiddenToGroups, vbTab & UCase(rsTemp!RoleToPrompt) & vbTab) > 0)) Then
          ' Found a Batch Job which is scheduled for another user group to run.
          pblnBatchJobsOK = False
      
          psScheduledUserGroups = psScheduledUserGroups & rsTemp!RoleToPrompt & vbCrLf
          
          If CurrentUserAccess(utlBatchJob, rsTemp!ID) = ACCESS_HIDDEN Then
            psScheduledJobDetails = psScheduledJobDetails & "Batch Job : <Hidden> by " & rsTemp!UserName & vbCrLf
          Else
            psScheduledJobDetails = psScheduledJobDetails & "Batch Job : " & rsTemp!Name & vbCrLf
          End If
        ElseIf rsTemp!nonHiddenCount > 0 Then
          piOwnedJobCount = piOwnedJobCount + 1
          psOwnedJobDetails = psOwnedJobDetails & "Batch Job : " & rsTemp!Name & " (Contains " & sKey & " '" & rsTemp!jobname & "') " & vbCrLf
          psOwnedJobIDs = psOwnedJobIDs & IIf(Len(psOwnedJobIDs) > 0, ", ", "") & rsTemp!ID
        End If
      Else
        ' Found a Batch Job whose owner is not the same
        pblnBatchJobsOK = False
    
        If CurrentUserAccess(utlBatchJob, rsTemp!ID) = ACCESS_HIDDEN Then
          psNonOwnedJobDetails = psNonOwnedJobDetails & "Batch Job : <Hidden> by " & rsTemp!UserName & vbCrLf
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
  Dim rsTemp As New ADODB.Recordset
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
    rsTemp.Open psSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
              
    Do Until rsTemp.EOF
      
      Select Case piUtilityType
        Case utlCalculation, utlFilter
          If LCase(rsTemp!UserName) = pstrUser Then
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
              psNonOwnedDetails = psNonOwnedDetails & sKey & " : <Hidden> by " & rsTemp!UserName & vbCrLf
            Else
              psNonOwnedDetails = psNonOwnedDetails & sKey & " : " & rsTemp!Name & vbCrLf
            End If
          End If
      
        Case Else
          If LCase(rsTemp!UserName) = pstrUser Then
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
              psNonOwnedDetails = psNonOwnedDetails & sKey & " : <Hidden> by " & rsTemp!UserName & vbCrLf
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


Public Sub HideUtilities(piUtilityType As UtilityType, _
  psIDs As String, _
  Optional pvHiddenUserGroups As Variant)
  ' Set the access for the given utility to be HIDDEN for the given user groups.
  ' psIDs is a COMMA delimited string of the utility/report IDs to be hidden
  ' pvHiddenUserGroups is a TAB delimited string of the user groups
  '   to which these definitions are to be hidden.
  '   NB. this string starts with a TAB also.
  
  Dim sSQL As String
  Dim sTableName As String
  Dim sAccessTableName As String
  Dim sIDColumnName As String
  Dim fHideFromAll As Boolean
  Dim sHiddenUserGroups As String

  fHideFromAll = IsMissing(pvHiddenUserGroups)
  sHiddenUserGroups = IIf(fHideFromAll, "", CStr(pvHiddenUserGroups))
  
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
      gADOCon.Execute sSQL, , adExecuteNoRecords
   
    Case Else
      If Len(sAccessTableName) > 0 Then
    
        If fHideFromAll Then
          sSQL = "DELETE FROM " & sAccessTableName & " WHERE ID IN (" & psIDs & ")"
          gADOCon.Execute sSQL, , adExecuteNoRecords
        
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
          gADOCon.Execute sSQL, , adExecuteNoRecords
        Else
          sHiddenUserGroups = "'" & Replace(Mid(Left(sHiddenUserGroups, Len(sHiddenUserGroups) - 1), 2), vbTab, "','") & "'"
          
          sSQL = "DELETE FROM " & sAccessTableName & " WHERE ID IN (" & psIDs & ") AND groupName IN (" & sHiddenUserGroups & ")"
          gADOCon.Execute sSQL, , adExecuteNoRecords
          
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
          gADOCon.Execute sSQL, , adExecuteNoRecords
        End If
      End If
  End Select

End Sub




Public Function CurrentUserAccess(piUtilityType As UtilityType, _
  plngID As Long) As String
  ' Return the access code (RW/RO/HD) of the current user's access
  ' on the given utility.
  On Error GoTo ErrorTrap
  
  Dim sAccessCode As String
  Dim sSQL As String
  Dim sDefaultAccess As String
  Dim rsAccess As New ADODB.Recordset
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
      " INNER JOIN sysusers a ON b.uid = a.gid" & _
      " LEFT OUTER JOIN " & sAccessTableName & " ON (b.name = " & sAccessTableName & ".groupName" & _
      "   AND " & sAccessTableName & ".id = " & CStr(plngID) & ")" & _
      " INNER JOIN " & sTableName & " ON " & sAccessTableName & ".ID = " & sTableName & "." & sIDColumnName & _
      " WHERE a.Name = system_user"

    rsAccess.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    With rsAccess
      If .BOF And .EOF Then
        sAccessCode = sDefaultAccess
      Else
        sAccessCode = !Access
      End If
      
      .Close
    End With
    Set rsAccess = Nothing
  Else
    sAccessCode = ACCESS_UNKNOWN
  End If
  
TidyUpAndExit:
  CurrentUserAccess = sAccessCode
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Function




Public Function GetAllExprRootIDs(plngID As Long) As String
  Dim rsTemp As New ADODB.Recordset
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

  rsTemp.Open strSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
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
  Dim rsAccess As New ADODB.Recordset
  Dim fIsSysSecUser As Boolean
  
  sSQL = "SELECT count(*) AS [result]" & _
    " FROM ASRSysGroupPermissions" & _
    " INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
    "   AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
    "   OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
    " INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
    "   AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
    " INNER JOIN sysusers b ON b.name = ASRSysGroupPermissions.groupname" & _
    " INNER JOIN sysusers a ON b.uid = a.gid" & _
    " AND IS_MEMBER(a.Name) = 1" & _
    " WHERE ASRSysGroupPermissions.permitted = 1"

  rsAccess.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsAccess
    fIsSysSecUser = (!Result > 0)
    
    .Close
  End With
  Set rsAccess = Nothing
  
  CurrentUserIsSysSecMgr = fIsSysSecUser

End Function

Public Function CurrentUserGroup() As String
  Dim sSQL As String
  Dim rsAccess As New ADODB.Recordset
  Dim sGroupName As String
  
  sSQL = "SELECT TOP 1 [GroupName]" & _
    " FROM ASRSysGroupPermissions" & _
    " INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
    "   AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
    "   OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
    " INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
    "   AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
    " INNER JOIN sysusers b ON b.name = ASRSysGroupPermissions.groupname" & _
    " INNER JOIN sysusers a ON b.uid = a.gid" & _
    "   AND IS_MEMBER(a.Name) = 1" & _
    " WHERE ASRSysGroupPermissions.permitted = 1" & _
    " ORDER BY [GroupName]"

  rsAccess.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsAccess
    If Not (.BOF And .EOF) Then
      sGroupName = .Fields(0).Value
    End If
    .Close
  End With
  Set rsAccess = Nothing
  
  CurrentUserGroup = sGroupName

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


