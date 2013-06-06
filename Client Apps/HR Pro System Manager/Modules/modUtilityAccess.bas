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
  
  If UCase(gsUserName) = "SA" Then
    CurrentUserAccess = ACCESS_READWRITE
    Exit Function
  End If
  
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


