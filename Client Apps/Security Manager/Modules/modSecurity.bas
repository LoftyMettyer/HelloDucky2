Attribute VB_Name = "modSecurity"
Option Explicit

'TM20010920 Fault 2530
'Boolean to check if the application is being exited.
Public blnIsExiting As Boolean

Public gObjGroups As New SecurityGroups

Public gasPrintOptions() As PrintOptions
Public asGroups() As String
Public gasPrintGroups() As String

Public gobjOperatorDefs As New clsOperatorDefs
Public gobjFunctionDefs As New clsFunctionDefs

Public Type PrintOptions
    PrintLPaneGROUPS As Boolean
    PrintLPaneGROUP As Boolean
    PrintLPaneUSERS As Boolean
    PrintLPaneTABLESVIEWS As Boolean
    PrintLPaneSYSTEM As Boolean
    PrintLPaneTABLE As Boolean
    PrintBlankVersion As Boolean
    '
    PrintRPaneGROUPS As Boolean
    PrintRPaneGROUP As Boolean
    PrintRPaneUSERS As Boolean
    PrintRPaneTABLESVIEWS As Boolean
    PrintRPaneSYSTEM As Boolean
    PrintRPaneTABLE As Boolean
End Type

' New Progress Bar - Global class to be used for progress bars
'Public gobjProgress As New HRProProgress.clsHRProProgress
'Public gobjProgress As New clsProgress

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hWnd As Long, ByVal lpoperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub RemoveGroupAccessRecords(psGroupName As String)
  ' Remove all utility/report access records for the given group.
  Dim sSQL As String
  
  sSQL = "DELETE FROM ASRSysBatchJobAccess" & _
    "  WHERE groupName = '" & psGroupName & "'"
  gADOCon.Execute sSQL, , adExecuteNoRecords

  sSQL = "DELETE FROM ASRSysCrossTabAccess" & _
    "  WHERE groupName = '" & psGroupName & "'"
  gADOCon.Execute sSQL, , adExecuteNoRecords

  sSQL = "DELETE FROM ASRSysCalendarReportAccess" & _
    "  WHERE groupName = '" & psGroupName & "'"
  gADOCon.Execute sSQL, , adExecuteNoRecords

  sSQL = "DELETE FROM ASRSysCustomReportAccess" & _
    "  WHERE groupName = '" & psGroupName & "'"
  gADOCon.Execute sSQL, , adExecuteNoRecords

  sSQL = "DELETE FROM ASRSysDataTransferAccess" & _
    "  WHERE groupName = '" & psGroupName & "'"
  gADOCon.Execute sSQL, , adExecuteNoRecords

  sSQL = "DELETE FROM ASRSysExportAccess" & _
    "  WHERE groupName = '" & psGroupName & "'"
  gADOCon.Execute sSQL, , adExecuteNoRecords

  sSQL = "DELETE FROM ASRSysGlobalAccess" & _
    "  WHERE groupName = '" & psGroupName & "'"
  gADOCon.Execute sSQL, , adExecuteNoRecords

  sSQL = "DELETE FROM ASRSysImportAccess" & _
    "  WHERE groupName = '" & psGroupName & "'"
  gADOCon.Execute sSQL, , adExecuteNoRecords

  sSQL = "DELETE FROM ASRSysMailMergeAccess" & _
    "  WHERE groupName = '" & psGroupName & "'"
  gADOCon.Execute sSQL, , adExecuteNoRecords

  sSQL = "DELETE FROM ASRSysRecordProfileAccess" & _
    "  WHERE groupName = '" & psGroupName & "'"
  gADOCon.Execute sSQL, , adExecuteNoRecords

  sSQL = "DELETE FROM ASRSysMatchReportAccess" & _
    "  WHERE groupName = '" & psGroupName & "'"
  gADOCon.Execute sSQL, , adExecuteNoRecords

End Sub

'Public Function UserLoggedIn(sUserID As String) As String
'
'  Dim sSQL As String
'  Dim rsUsers As New ADODB.Recordset
'
'  On Error GoTo Error_Trap
'
'  sSQL = "SELECT DISTINCT hostname, loginame, program_name, hostprocess " & _
'    "FROM master..sysprocesses " & _
'    "WHERE program_name like 'OpenHR%' " & _
'    "  AND program_name NOT LIKE 'OpenHR Workflow%' " & _
'    "  AND LOWER(loginame) = '" & LCase(Replace(sUserID, "'", "''")) & "' "
''                    "AND dbid in (" & _
''                        "SELECT dbid " & _
''                        "FROM master..sysdatabases " & _
''                        "WHERE name = '" & gsDatabaseName & "') "
'
'  rsUsers.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'  With rsUsers
'    If Not (.EOF And .BOF) Then
'      UserLoggedIn = Trim(.Fields("Program_name").Value)
'    Else
'      UserLoggedIn = vbNullString
'    End If
'  End With
'  rsUsers.Close
'
'TidyUpAndExit:
'  Set rsUsers = Nothing
'  Exit Function
'
'Error_Trap:
'  MsgBox "Error validating current users.", vbExclamation + vbOKOnly, App.Title
'  UserLoggedIn = True
'  GoTo TidyUpAndExit
'
'End Function


'Public Function UserSessions(sUserID As String) As Integer
'  Dim rsTemp As New ADODB.Recordset
'  Dim strSQL As String
'
'  On Local Error GoTo Error_Trap
'
'  'JPD 20050812 Fault 10166
'  strSQL = "SELECT COUNT(*) AS [count] " & _
'    "FROM master..sysprocesses " & _
'    "WHERE program_name like 'OpenHR%' " & _
'    "  AND program_name NOT LIKE 'OpenHR Workflow%' " & _
'    "  AND LOWER(loginame) = '" & LCase(Replace(sUserID, "'", "''")) & "' "
'  rsTemp.Open strSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
'  UserSessions = rsTemp!Count
'
'  rsTemp.Close
'  Set rsTemp = Nothing
'
'TidyUpAndExit:
'  Set rsTemp = Nothing
'  Exit Function
'
'Error_Trap:
'  MsgBox "Error validating current users.", vbExclamation + vbOKOnly, App.Title
'  UserSessions = 0
'  GoTo TidyUpAndExit
'
'End Function

Function InitialiseGroupsCollection(pobjGroups As SecurityGroups) As Boolean
  ' Initialise the security groups structure.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fGoodGroup As Boolean
  Dim iNextIndex As Integer
  Dim sGroupName As String
  Dim rsGroups As New ADODB.Recordset
  Dim rsSysRoles As New ADODB.Recordset
  Dim asFixedRoles() As String
  
  fOK = True
  ReDim asFixedRoles(0)
  
  ' Set the mouse pointer to an hour glass
  Screen.MousePointer = vbHourglass
  
  ' Get a list of groups/Roles from SQL Server
  rsSysRoles.Open "sp_helpdbfixedrole", gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsSysRoles
    Do While Not .EOF
      iNextIndex = UBound(asFixedRoles) + 1
      ReDim Preserve asFixedRoles(iNextIndex)
      asFixedRoles(iNextIndex) = .Fields(0).Value
      .MoveNext
    Loop
    
    .Close
  End With
  
  ' Create a security tables and security users collection for each group.
  rsGroups.Open "sp_helprole", gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsGroups
    If Not .EOF And Not .BOF Then
      While Not .EOF
        sGroupName = Trim(.Fields(0).Value)
      
        'fGoodGroup = True
        fGoodGroup = (LCase$(Left$(sGroupName, 6)) <> "asrsys")
        
        ' Check that the group is valid.
        If sGroupName = "public" Then
          fGoodGroup = False
        Else
          ' Check if the group is a 'fixed system role'.
          For iNextIndex = 1 To UBound(asFixedRoles)
            If asFixedRoles(iNextIndex) = sGroupName Then
              fGoodGroup = False
              Exit For
            End If
          Next iNextIndex
        End If

        
        If fGoodGroup Then
          ' Add the group to the groups collection
          AddGroup pobjGroups, .Fields(0).Value
          pobjGroups(.Fields(0).Value).NewGroup = False
          pobjGroups(.Fields(0).Value).Changed = False
        End If
        
        .MoveNext
      Wend
    End If
  
    .Close
  End With
    
TidyUpAndExit:
  Set rsSysRoles = Nothing
  Set rsGroups = Nothing
  ' Set the mouse pointer back to normal
  Screen.MousePointer = vbNormal
  InitialiseGroupsCollection = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Sub AddGroup(pobjGroups As SecurityGroups, psGroupName As String)
  ' Adds a security group structure.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objGroup As SecurityGroup
  
  ' If the group already exists but is marked as deleted then just mark it as undeleted.
  For Each objGroup In pobjGroups
    If UCase(objGroup.Name) = UCase(psGroupName) Then
      pobjGroups(psGroupName).DeleteGroup = False
      pobjGroups(psGroupName).RequireLogout = False   'MH20010410
      Set objGroup = Nothing
      Exit Sub
    End If
  Next objGroup
  Set objGroup = Nothing
  
  ' Add the new group to the groups collection now that we have instantiated the collection classes in it.
  Set objGroup = pobjGroups.Add(psGroupName, True, False, True, psGroupName)
  
TidyUpAndExit:
  ' Release the objects.
  Set objGroup = Nothing
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub




Function InitialiseGroup(pobjGroup As SecurityGroup, bShowProgress As Boolean) As Boolean
  ' Initialise the table, view and system permission collections for the given security group.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
    
  fOK = True
  
  Screen.MousePointer = vbHourglass
  
  If bShowProgress Then
    With gobjProgress
      '.AviFile = App.Path & "\videos\DB_Transfer.Avi"
      .AVI = dbTransfer
      .MainCaption = pobjGroup.Name & "..."
      .Caption = "User Group"
      .NumberOfBars = 0
      .Time = False
      .Cancel = False
      .OpenProgress
    End With
  End If
  
  ' Initialise the user views collection.
  fOK = SetupTablesCollection(pobjGroup)

  If fOK Then
    ' Initialise the system permissions collection.
    fOK = InitialiseSystemPermissionsCollection(pobjGroup)
    If bShowProgress Then gobjProgress.UpdateProgress False
  End If
  
  pobjGroup.Initialised = fOK
  
TidyUpAndExit:
  If Not fOK Then
    pobjGroup.Views.Clear
    pobjGroup.Tables.Clear
    pobjGroup.SystemPermissions.Clear
  End If
  
  ' Reset the mouse pointer
  If bShowProgress Then gobjProgress.CloseProgress
  Screen.MousePointer = vbNormal
  InitialiseGroup = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function



Public Function IsMinimumPasswordLength(sPassword As String) As Boolean

  Dim lMinimumLength As Long      ' The minimum length for passwords
  
  ' First store the config info in local variables
  If glngSQLVersion >= 9 Then
    lMinimumLength = glngDomainMinimumLength
  Else
    lMinimumLength = GetSystemSetting("Password", "Minimum Length", 0)
  End If

  If lMinimumLength = 0 Then
    IsMinimumPasswordLength = True
  ElseIf Len(sPassword) >= lMinimumLength Then
    IsMinimumPasswordLength = True
  Else
    IsMinimumPasswordLength = False
  End If
  
End Function


Private Function SetupTablesCollection(pobjGroup As SecurityGroup) As Boolean
  ' Read the list of tables the current user has permission to see.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim fSysSecManager As Boolean
  Dim fSelectAllPermission As Boolean
  Dim fSelectNonePermission As Boolean
  Dim fUpdateAllPermission As Boolean
  Dim fUpdateNonePermission As Boolean
  Dim iLoop As Integer
  Dim lngNextIndex As Long
  Dim lngRoleID As Long
  Dim lngChildViewID As Long
  Dim sSQL As String
  Dim sLastRealSource As String
  Dim sRealSourceList As New SecurityMgr.clsStringBuilder
  Dim sTableViewName As String
  Dim rsInfo As New ADODB.Recordset
  Dim rsTables As ADODB.Recordset
  Dim rsViews As ADODB.Recordset
  Dim rsPermissions As ADODB.Recordset
  Dim objColumn As SecurityColumn
  Dim objColumns As SecurityColumns
  Dim objTableView As SecurityTable
  Dim avChildViews() As Variant
'''  Dim adoCon As ADODB.Connection
  Dim iTemp As Integer
  Dim fChildView As Boolean
  Dim sConnectString As String

  Dim iAction As Integer
  Dim strViewName As String
  Dim strColumnName As String
  Dim strTableName As String

  fOK = True

  ' Check if the user is a 'system manager' or 'security manager'.
  ' If so then we can save time by applying all table permissions, instead of having to read them first.
  sSQL = "SELECT count(itemKey) AS recCount" & _
    " FROM ASRSysGroupPermissions" & _
    " INNER JOIN ASRSysPermissionItems ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID" & _
    " INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
    " INNER JOIN sysusers a ON ASRSysGroupPermissions.groupName = a.name" & _
    " INNER JOIN sysusers b ON a.uid = b.gid" & _
    " WHERE b.Name = '" & pobjGroup.Name & "'" & _
    " AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
    " OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER')" & _
    " AND ASRSysGroupPermissions.permitted = 1" & _
    " AND ASRSysPermissionCategories.categorykey = 'MODULEACCESS'"
  rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText

  fSysSecManager = (rsInfo.Fields(0).Value > 0)
  rsInfo.Close

  ' Create an array of child view IDs and their associated table names.
  ' Column 1 - child view ID
  ' Column 2 - associated table name
  ' Column 3 - 0=OR, 1=AND
  
  ' JPD20020809 New child views
  sSQL = "SELECT ASRSysChildViews2.childViewID, ASRSysTables.tableName, ASRSysChildViews2.type" & _
    " FROM ASRSysChildViews2" & _
    " INNER JOIN ASRSysTables ON ASRSysChildViews2.tableID = ASRSysTables.tableID" & _
    " WHERE ASRSysChildViews2.role = '" & pobjGroup.Name & "'"
  rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText

  ReDim avChildViews(3, 100)
  lngNextIndex = 0
  If Not rsInfo.EOF Then
    Do While Not rsInfo.EOF
      lngNextIndex = lngNextIndex + 1
      If lngNextIndex > UBound(avChildViews, 2) Then ReDim Preserve avChildViews(3, lngNextIndex + 100)
      avChildViews(1, lngNextIndex) = rsInfo(0).Value
      avChildViews(2, lngNextIndex) = rsInfo(1).Value
      avChildViews(3, lngNextIndex) = IIf(IsNull(rsInfo(2).Value), 0, rsInfo(2).Value)
      rsInfo.MoveNext
    Loop
    ReDim Preserve avChildViews(3, lngNextIndex)
  End If
  rsInfo.Close

  ' Get the collection with items for each TABLE in the system.
  sSQL = "SELECT ASRSysTables.tableID, ASRSysTables.tableName, ASRSysTables.tableType," & _
    " COUNT(ASRSysRelations.parentID) AS parentCount" & _
    " FROM ASRSysTables" & _
    " LEFT OUTER JOIN ASRSysRelations ON ASRSysTables.tableID= ASRSysRelations.childID" & _
    " GROUP BY ASRSysTables.tableID, ASRSysTables.tableName, ASRSysTables.tableType"

  Set rsTables = New Recordset
  rsTables.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
  With rsTables
    Do While Not .EOF

      strTableName = .Fields(1).Value

      Set objColumns = New SecurityColumns
      pobjGroup.Tables.Add objColumns, strTableName, .Fields(2).Value, strTableName
      Set objColumns = Nothing

      With pobjGroup.Tables(strTableName)
        .Columns_Initialised = True
        .SelectPrivilege = IIf(fSysSecManager, giPRIVILEGES_ALLGRANTED, IIf(rsTables.Fields(2).Value = tabLookup, giPRIVILEGES_ALLGRANTED, giPRIVILEGES_NONEGRANTED))
        .UpdatePrivilege = IIf(fSysSecManager, giPRIVILEGES_ALLGRANTED, giPRIVILEGES_NONEGRANTED)
        .InsertPrivilege = IIf(fSysSecManager, True, False)
        .DeletePrivilege = IIf(fSysSecManager, True, False)
        .ParentJoinType = 0
        .ParentCount = rsTables.Fields(3).Value
        .TableID = rsTables.Fields(0).Value
        .HideFromMenu = False
        .InsertOriginalPrivilege = .InsertPrivilege
        .DeleteOriginalPrivilege = .DeletePrivilege
      End With

      .MoveNext
    Loop
    .Close
  End With
  Set rsTables = Nothing

  ' Initialise the collection with items for each VIEW in the system.
  sSQL = "SELECT ASRSysViews.viewName, ASRSysViews.viewTableID " & _
    " FROM ASRSysViews"
  Set rsViews = New Recordset
  rsViews.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
  With rsViews
    Do While Not .EOF
      strViewName = .Fields(0).Value
      Set objColumns = New SecurityColumns
      pobjGroup.Views.Add objColumns, strViewName, 0, strViewName
      Set objColumns = Nothing

      With pobjGroup.Views(strViewName)
        .Columns_Initialised = True
        .SelectPrivilege = IIf(fSysSecManager, giPRIVILEGES_ALLGRANTED, giPRIVILEGES_NONEGRANTED)
        .UpdatePrivilege = IIf(fSysSecManager, giPRIVILEGES_ALLGRANTED, giPRIVILEGES_NONEGRANTED)
        .InsertPrivilege = IIf(fSysSecManager, True, False)
        .DeletePrivilege = IIf(fSysSecManager, True, False)
        .ParentJoinType = 0
        .ParentCount = 0
        .ViewTableID = rsViews.Fields(1).Value
        .InsertOriginalPrivilege = .InsertPrivilege
        .DeleteOriginalPrivilege = .DeletePrivilege
      End With

      .MoveNext
    Loop
    .Close
  End With
  Set rsViews = Nothing

  ' Get the permissions for each table or view.
  sRealSourceList.TheString = vbNullString
  sLastRealSource = vbNullString

  If Not fSysSecManager Then
    ' If the user is NOT a 'system manager' or 'security manager'
    ' read the table permissions from the server.
    sSQL = "exec sp_ASRAllTablePermissionsForGroup '" & pobjGroup.Name & "'"
    Set rsPermissions = New ADODB.Recordset
    rsPermissions.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText

    Do While Not rsPermissions.EOF
      Set objTableView = Nothing

      iAction = rsPermissions!Action
      strTableName = rsPermissions!Name

      If sLastRealSource <> strTableName Then
        sRealSourceList.Append IIf(sRealSourceList.Length <> 0, ", '", "'") & strTableName & "'"
        sLastRealSource = strTableName
      End If

      If (iAction = 195) Or (iAction = 196) Then
        fChildView = False

        If UCase(Left(strTableName, 8)) = "ASRSYSCV" Then
          fChildView = True
          ' Determine which table the child view is for.
          iTemp = InStr(strTableName, "#")
          lngChildViewID = Val(Mid(strTableName, 9, iTemp - 9))
        End If

        If fChildView Then
          For lngNextIndex = 1 To UBound(avChildViews, 2)
            If avChildViews(1, lngNextIndex) = lngChildViewID Then
              Set objTableView = pobjGroup.Tables(avChildViews(2, lngNextIndex))
              objTableView.ParentJoinType = avChildViews(3, lngNextIndex)
              Exit For
            End If
          Next lngNextIndex
        Else
          If UCase(Left(strTableName, 6)) <> "ASRSYS" Then
            If pobjGroup.Tables.IsValid(strTableName) Then
              Set objTableView = pobjGroup.Tables(strTableName)
            Else
              Set objTableView = pobjGroup.Views(strTableName)
            End If
          End If
        End If

        If Not objTableView Is Nothing Then
          Select Case iAction
            Case 195 ' Insert permission.
              objTableView.InsertPrivilege = True
              'MH20050106 Fault 9470, 9471, 9472, 9373
              objTableView.InsertOriginalPrivilege = True
            Case 196 ' Delete permission.
              objTableView.DeletePrivilege = True
              'MH20050106 Fault 9470, 9471, 9472, 9373
              objTableView.DeleteOriginalPrivilege = True
          End Select
        End If
      End If

      rsPermissions.MoveNext
    Loop
    rsPermissions.Close
    Set rsPermissions = Nothing

    ' Get the view menu permissions
    sSQL = "SELECT TableName, HideFromMenu FROM  ASRSysViewMenuPermissions WHERE GroupName = '" & pobjGroup.Name & "'"
    Set rsPermissions = New ADODB.Recordset
    rsPermissions.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText

    Do While Not rsPermissions.EOF
      'JPD 20031114 Fault 7624
      If pobjGroup.Tables.IsValid(rsPermissions.Fields(0).Value) Then
        Set objTableView = pobjGroup.Tables(rsPermissions.Fields(0).Value)
        objTableView.HideFromMenu = rsPermissions.Fields(1).Value
      End If

      rsPermissions.MoveNext
    Loop
    rsPermissions.Close
    Set rsPermissions = Nothing

  End If

  ' Get the list of all columns in all tables/views.
  sSQL = "SELECT ASRSysColumns.columnName," & _
    " ASRSysTables.tableName AS tableViewName," & _
    " ASRSysTables.tableType AS tableType" & _
    " FROM ASRSysColumns" & _
    " INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID" & _
    " WHERE ASRSysColumns.columnType <> " & Trim(Str(colSystem)) & _
    " AND ASRSysColumns.columnType <> " & Trim(Str(colLink)) & _
    " UNION" & _
    " SELECT ASRSysColumns.columnName," & _
    " ASRSysViews.viewName AS tableViewName," & _
    " 0 AS tableType" & _
    " FROM ASRSysColumns" & _
    " INNER JOIN ASRSysViews ON ASRSysColumns.tableID = ASRSysViews.viewTableID" & _
    " INNER JOIN ASRSysViewColumns ON (ASRSysColumns.columnID = ASRSysViewColumns.columnID AND ASRSysViews.viewID = ASRSysViewColumns.viewID)" & _
    " WHERE ASRSysViewColumns.inView = 1" & _
    " AND ASRSysColumns.columnType <> " & Trim(Str(colSystem)) & _
    " AND ASRSysColumns.columnType <> " & Trim(Str(colLink))
  rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText

  Do While Not rsInfo.EOF
    If pobjGroup.Tables.IsValid(rsInfo.Fields(1).Value) Then
      Set objColumns = pobjGroup.Tables(rsInfo.Fields(1).Value).Columns
    Else
      Set objColumns = pobjGroup.Views(rsInfo.Fields(1).Value).Columns
    End If

    ' Add the column object to the collection.
    Set objColumn = objColumns.Add(rsInfo.Fields(0).Value)

    ' Set the security column properties
    objColumn.Name = rsInfo.Fields(0).Value
    objColumn.Changed = False
    objColumn.SelectPrivilege = fSysSecManager Or (rsInfo.Fields(2).Value = tabLookup)
    objColumn.UpdatePrivilege = fSysSecManager

    ' Release the security column
    Set objColumn = Nothing
    Set objColumns = Nothing

    rsInfo.MoveNext
  Loop
  rsInfo.Close

  ' If the current user is not a system/security manager then read the column permissions from SQL.
  If (Not fSysSecManager) And (Not pobjGroup.NewGroup) And (sRealSourceList.Length <> 0) Then
    ' Get the SQL group id of the current user.
    sSQL = "SELECT gid" & _
      " FROM sysusers" & _
      " WHERE name = '" & pobjGroup.Name & "'"
    rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    lngRoleID = rsInfo.Fields(0).Value
    rsInfo.Close

    sSQL = "SELECT sysobjects.name AS tableViewName," & _
      " syscolumns.name AS columnName," & _
      " sysprotects.action," & _
      " CASE protectType" & _
      "   WHEN 205 THEN 1" & _
      "   WHEN 204 THEN 1" & _
      "   ELSE 0" & _
      " END AS permission" & _
      " FROM sysprotects" & _
      " INNER JOIN sysobjects ON sysprotects.id = sysobjects.id" & _
      " INNER JOIN syscolumns ON sysprotects.id = syscolumns.id" & _
      " WHERE sysprotects.uid = " & Trim(Str(lngRoleID)) & _
      " AND (sysprotects.action = 193 or sysprotects.action = 197)" & _
      " AND syscolumns.name <> 'timestamp'" & _
      " AND sysobjects.name in (" & sRealSourceList.ToString & ")" & _
      " AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0" & _
      " AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)" & _
      " OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0" & _
      " AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))"
    rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText

    Do While Not rsInfo.EOF
      ' Get the current column's table/view name.
      Set objTableView = Nothing

      fChildView = False

      strTableName = rsInfo.Fields(0).Value
      strColumnName = rsInfo.Fields(1).Value

      If UCase$(Left$(strTableName, 8)) = "ASRSYSCV" Then
        fChildView = True
        ' Determine which table the child view is for.
        iTemp = InStr(strTableName, "#")
        lngChildViewID = Val(Mid(strTableName, 9, iTemp - 9))
      End If

      If fChildView Then
        For lngNextIndex = 1 To UBound(avChildViews, 2)
          If avChildViews(1, lngNextIndex) = lngChildViewID Then
            Set objTableView = pobjGroup.Tables(avChildViews(2, lngNextIndex))
            objTableView.ParentJoinType = avChildViews(3, lngNextIndex)
            Exit For
          End If
        Next lngNextIndex
      Else
        If UCase$(Left$(strTableName, 6)) <> "ASRSYS" Then
          If pobjGroup.Tables.IsValid(strTableName) Then
            Set objTableView = pobjGroup.Tables(strTableName)
          Else
            Set objTableView = pobjGroup.Views(strTableName)
          End If
        End If
      End If

      If Not objTableView Is Nothing Then
        If rsInfo.Fields(2).Value = 193 Then
          If objTableView.Columns.IsValid(strColumnName) Then
            With objTableView.Columns(strColumnName)
              .SelectPrivilege = rsInfo.Fields(3).Value
              'MH20050106 Fault 9470, 9471, 9472, 9373
              .SelectOriginalPrivilege = .SelectPrivilege
            End With
          End If
        End If

        If rsInfo.Fields(2).Value = 197 Then
          If objTableView.Columns.IsValid(strColumnName) Then
            With objTableView.Columns(strColumnName)
              .UpdatePrivilege = rsInfo.Fields(3).Value
              'MH20050106 Fault 9470, 9471, 9472, 9373
              .UpdateOriginalPrivilege = .UpdatePrivilege
            End With
          End If
        End If
      End If

      rsInfo.MoveNext
    Loop
    rsInfo.Close

    ' Check if the table has SELECT/UPDATE ALL/SOME/NONE.
    For Each objTableView In pobjGroup.Tables
      fSelectAllPermission = True
      fSelectNonePermission = True
      fUpdateAllPermission = True
      fUpdateNonePermission = True

      For Each objColumn In objTableView.Columns
        If objColumn.SelectPrivilege Then
          fSelectNonePermission = False
        Else
          fSelectAllPermission = False
        End If

        If objColumn.UpdatePrivilege Then
          fUpdateNonePermission = False
        Else
          fUpdateAllPermission = False
        End If
      Next objColumn
      Set objColumn = Nothing

      objTableView.SelectPrivilege = IIf(fSelectAllPermission, giPRIVILEGES_ALLGRANTED, IIf(fSelectNonePermission, giPRIVILEGES_NONEGRANTED, giPRIVILEGES_SOMEGRANTED))
      objTableView.UpdatePrivilege = IIf(fUpdateAllPermission, giPRIVILEGES_ALLGRANTED, IIf(fUpdateNonePermission, giPRIVILEGES_NONEGRANTED, giPRIVILEGES_SOMEGRANTED))

    Next objTableView
    Set objTableView = Nothing
  End If

TidyUpAndExit:
  Set rsInfo = Nothing
  SetupTablesCollection = fOK
  Exit Function

ErrorTrap:
  fOK = False
  gobjProgress.Visible = False
  MsgBox Err.Description, vbCritical + vbOKOnly, App.Title
  Resume TidyUpAndExit

End Function


Function InitialiseUsersCollection(pobjGroup As SecurityGroup) As Boolean
  ' Initialise the security users structure.
  Dim fOK As Boolean
  Dim lngGroupID As Long
  Dim iNextIndex As Integer
  Dim sSQL As String
  Dim rsUserInfo As New ADODB.Recordset
  Dim rsGroupInfo As New ADODB.Recordset
  ' NPG20090204 Fault 11931
  Dim iLoginType As SecurityMgr.LoginType
  Dim strLoginName As String
  
  ' Set the mouse pointer to an hour glass
  Screen.MousePointer = vbHourglass
  fOK = True
  
  On Error GoTo err_InitialiseUsersCollection

  ' AE20080509 Fault #13163
  'If glngSQLVersion = 9 Then
  If glngSQLVersion >= 9 Then
  
    sSQL = "SELECT usu.IsNTGroup, usu.IsNTUser, usu.name, ISNULL(suser_sname(p.sid),N'') AS [LoginName] " & _
      " FROM sysusers usu " & _
      " INNER JOIN sys.database_principals p ON usu.name = p.name " & _
      " INNER JOIN (sysmembers mem INNER JOIN sysusers usg ON mem.groupuid = usg.uid) on usu.uid = mem.memberuid " & _
      " WHERE (usu.islogin = 1 And usu.isaliased = 0 And usu.hasdbaccess = 1) And (usg.issqlrole = 1 Or usg.uid Is Null) " & _
      " AND usg.name = '" & pobjGroup.Name & "' AND suser_sname(p.sid) IS NOT NULL"
  
  Else

    sSQL = "SELECT usu.IsNTGroup, usu.IsNTUser, usu.name, master.dbo.syslogins.loginname " & _
      " FROM sysusers usu" & _
      " INNER JOIN master.dbo.syslogins ON usu.sid = master.dbo.syslogins.sid" & _
      " LEFT OUTER join  (sysmembers mem INNER JOIN sysusers usg ON mem.groupuid = usg.uid) on usu.uid = mem.memberuid WHERE (usu.islogin = 1 and usu.isaliased = 0 and usu.hasdbaccess = 1) and (usg.issqlrole = 1 or usg.uid IS NULL)" & _
      " AND usg.name = '" & pobjGroup.Name & "'"
      
  End If
  
  rsUserInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
  With rsUserInfo
    Do While Not .EOF
    
      strLoginName = IIf(IsNull(!loginname), vbNullString, !loginname)
      pobjGroup.Users.Add !Name, False, False, False, vbNullString, vbNullString, strLoginName, vbNullString
      
      
      ' NPG20090204 Fault 11931
      ' Workout the login type
'      If !IsNTGroup Then
'        iLoginType = iUSERTYPE_TRUSTEDGROUP
'      Else
'        If !IsNTUser Then
'          iLoginType = iUSERTYPE_TRUSTEDUSER
'        Else
'          iLoginType = iUSERTYPE_SQLLOGIN
'        End If
'      End If
      
      ' Workout the login type
      If !IsNTGroup Or !ISNTUser Then
' -------------------------------------------------------------------------------------------------------
        ' NPG20090204 Fault 11931
        ' Does user exist in the Security Logins list?
        ' Decide if it is a new user login.
        If Not IsSQLLoginNameInUse(strLoginName) Then
          ' The login does not exist in the SQL Server database, mark as unlinked.
          iLoginType = IIf(!IsNTGroup, iUSERTYPE_ORPHANGROUP, iUSERTYPE_ORPHANUSER)
        Else
          ' The login does exist in the SQL Server database, mark as normal.
          iLoginType = IIf(!IsNTGroup, iUSERTYPE_TRUSTEDGROUP, iUSERTYPE_TRUSTEDUSER)
        End If
        
' -------------------------------------------------------------------------------------------------------
      Else
        iLoginType = iUSERTYPE_SQLLOGIN
      End If
     
      With pobjGroup.Users(!Name)
        .LoginType = iLoginType
      End With
      
      .MoveNext
    Loop
    
    .Close
  End With
  
  Set rsUserInfo = Nothing
  
  pobjGroup.Users_Initialised = True
    
  ' Indicate successfull procedure
  InitialiseUsersCollection = True
  
  Screen.MousePointer = vbNormal
  
  Exit Function
  
err_InitialiseUsersCollection:
  Dim lErrorNo As Long
  Dim sErrorMsg As String
  
  ' Clear the user collection.
  Set pobjGroup.Users = Nothing
  Set pobjGroup.Users = New SecurityUsers
  
  lErrorNo = Err.Number
  sErrorMsg = Err.Description
  
  Screen.MousePointer = vbNormal
  
  MsgBox Err.Description, vbCritical + vbOKOnly, App.Title
  
End Function

Function InitialiseSystemPermissionsCollection(pobjGroup As SecurityGroup) As Boolean
  ' Initialise the security system privileges structure for the given user group.
  On Error GoTo ErrorTrap
  
  Dim sSQL As String
  Dim rsPermissions As New ADODB.Recordset
  
  ' Get the list of System Permissions.
  sSQL = "SELECT ASRSysPermissionItems.itemID," & _
    " ASRSysPermissionCategories.categoryKey, " & _
    " ASRSysPermissionItems.itemKey," & _
    " ASRSysGroupPermissions.permitted" & _
    " FROM ASRSysPermissionItems" & _
    " INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
    " LEFT OUTER JOIN ASRSysGroupPermissions ON ASRSysPermissionItems.itemID = ASRSysGroupPermissions.itemID" & _
    "   AND ASRSysGroupPermissions.groupName = '" & pobjGroup.Name & "'" & _
    " ORDER BY ASRSysPermissionCategories.listOrder, ASRSysPermissionCategories.description, ASRSysPermissionItems.listOrder, ASRSysPermissionItems.description"
  
  rsPermissions.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
  
 ' Set rsPermissions = rdoCon.OpenResultset(sSQL, rdOpenForwardOnly)
  
  ' Iterate through the resultset adding a permission for each one.
  'While Not rsPermissions.EOF
  While Not rsPermissions.EOF
  
    ' Add the permission to the collection of system permissions if it is not already there.
    If Not pobjGroup.SystemPermissions.IsValid(rsPermissions.Fields("ItemKey").Value) Then
      pobjGroup.SystemPermissions.Add rsPermissions.Fields("ItemID").Value, _
        IIf(IsNull(rsPermissions.Fields("permitted").Value), 0, rsPermissions.Fields("permitted").Value), _
        rsPermissions.Fields("ItemKey").Value, _
        rsPermissions.Fields("CategoryKey").Value
    End If

    rsPermissions.MoveNext
  Wend
  
  ' Release the result set
  rsPermissions.Close
  Set rsPermissions = Nothing
        
  ' Indicate successfull procedure
  InitialiseSystemPermissionsCollection = True
  
  Exit Function

ErrorTrap:
  
  gobjProgress.Visible = False
  MsgBox Err.Description, vbCritical + vbOKOnly, App.Title
  
End Function



Function ApplyChanges() As Boolean
  ' Apply the security changes to SQL Server
  On Error GoTo ErrorTrap
  
  Const iProgressSteps = 14
  Dim fOK As Boolean
  Dim iCount As Integer
  
  fOK = True
  'Screen.MousePointer = vbHourglass
  
  'RH 02/10/00 - BUG 1044 - Disable the X button whilst saving
  EnableCloseButton frmMain.hWnd, False
  
  'RH 21/08/00 - BUG 788. Disable all forms while we save changes - this
  '              prevents users from col editing etc whilst saving ! a good idea
  '              methinks !
  For iCount = 0 To (Forms.Count - 1)
      Forms(iCount).Enabled = False
  Next iCount
    
  'MH20010410
  If fOK Then
    DoEvents
    fOK = ApplyChanges_LogoutCheck
    If fOK Then
      LockDatabase (lckSaving)
    Else
      MsgBox "Save process cancelled.", vbOKOnly + vbExclamation, Application.Name
    End If
  End If
   
  OutputCurrentProcess "Start of save process", True
  
  If fOK Then
    ' Initialise the progress bar.
    With gobjProgress
      '.AviFile = App.Path & "\videos\DB_Transfer.Avi"
      .AVI = dbSave
      .MainCaption = "Saving Changes"
      .Caption = Application.Name
      .NumberOfBars = 2
      .Bar1Value = 0
      .Bar1MaxValue = iProgressSteps
      .Bar1Caption = "Applying changes to the server database..."
      .Bar2Value = 0
      .Time = False
      .Cancel = False
      .OpenProgress
    End With
  End If

  ' Shift saving - Initialise every security group and mark as changed
  If fOK And gbShiftSave Then
    OutputCurrentProcess "Initialising User Groups"
    fOK = MarkAllGroupsAsChanged
    gobjProgress.UpdateProgress False
  'JPD 20060719 Fault 10841
  'Else
  ElseIf fOK Then
    OutputCurrentProcess "Initialising Standard User Groups"
    gobjProgress.UpdateProgress False
  End If

  ' Progess step 1 - Tidy up orphaned logins
  'JPD 20060227 Fault 10841
  If fOK Then
    If gbDeleteOrphanWindowsLogins And gbCanUseWindowsAuthentication Then
      ' Create the new User Groups (Roles).
      OutputCurrentProcess "Database preparation"
      fOK = ApplyChanges_TidyUpOrphans
      gobjProgress.UpdateProgress False
  
      ' Flag any errors.
      If Not fOK Then
        gobjProgress.Visible = False
        MsgBox "Error tidying up security logins.", vbOKOnly + vbExclamation, Application.Name
      End If
    Else
      gobjProgress.UpdateProgress False
    End If
  End If
  
  '
  ' Progress step 1A - Apply User deletions.
  '
  If fOK Then
  
    ' Delete Users.
    OutputCurrentProcess2 "Deletions"
    fOK = ApplyChanges_DeleteUsers
    gobjProgress.UpdateProgress2
  
    ' Flag any errors.
    If Not fOK Then
      gobjProgress.Visible = False
      MsgBox "Error deleting users.", vbOKOnly + vbExclamation, Application.Name
    End If
  End If

  '
  ' Progress step 1B - Add new User Groups (Roles).
  '
  If fOK Then
    
    ' Create the new User Groups (Roles).
    OutputCurrentProcess "New Groups"
    fOK = ApplyChanges_NewGroups
    gobjProgress.UpdateProgress False

    ' Flag any errors.
    If Not fOK Then
      gobjProgress.Visible = False
      MsgBox "Error creating new user groups.", vbOKOnly + vbExclamation, Application.Name
    End If
  End If
  
  '
  ' Progress step 2 - Apply Top-Level/Lookup Table and Table Column permissions.
  '
  If fOK Then
  
    ' Apply the Table and Table Column permissions.
    OutputCurrentProcess "Applying Permissions"
    OutputCurrentProcess2 "Non Child Tables", 6
    fOK = ApplyChanges_NonChildTablePermissions
    gobjProgress.UpdateProgress2 False
  
    ' Flag any errors.
    If Not fOK Then
       gobjProgress.Visible = False
      MsgBox "Error applying table permissions.", vbOKOnly + vbExclamation, Application.Name
    End If
  End If
  
  '
  ' Progress step 3 - Apply View and View Column permissions.
  '
  If fOK Then
  
    ' Apply the View and View Column permissions.
    OutputCurrentProcess2 "Views"
    fOK = ApplyChanges_UserViewPermissions
    gobjProgress.UpdateProgress2
  
    ' Flag any errors.
    If Not fOK Then
      gobjProgress.Visible = False
      MsgBox "Error applying view permissions.", vbOKOnly + vbExclamation, Application.Name
    End If
  End If

  '
  ' Progress step 4 - Apply Child Table (and child view) permissions.
  '
  If fOK Then
  
    ' Apply the Child Table (and child view) permissions.
    OutputCurrentProcess2 "Child Tables"
    fOK = ApplyChanges_ChildTablePermissions
    gobjProgress.UpdateProgress2
  
    ' Flag any errors.
    If Not fOK Then
      gobjProgress.Visible = False
      MsgBox "Error applying child view permutation permissions.", vbOKOnly + vbExclamation, Application.Name
    End If
  End If
  
  '
  ' Progress step 5 - Apply Child Table (and child view) permissions.
  '
  If fOK Then
  
    ' Apply the Child Table (and child view) permissions.
    OutputCurrentProcess2 "Child Tables"
    fOK = ApplyChanges_TidyChildTablePermissions
    gobjProgress.UpdateProgress2
  
    ' Flag any errors.
    If Not fOK Then
      gobjProgress.Visible = False
      MsgBox "Error tidying child view permissions.", vbOKOnly + vbExclamation, Application.Name
    End If
  End If
  
  '
  ' Progress step 5a - Apply System permissions.
  '
  If fOK Then
  
    ' Apply the System Permissions.
    OutputCurrentProcess2 "System Permissions"
    fOK = ApplyChanges_SystemPermissions
    gobjProgress.UpdateProgress2
  
    ' Flag any errors.
    If Not fOK Then
      gobjProgress.Visible = False
      MsgBox "Error applying system permissions.", vbOKOnly + vbExclamation, Application.Name
    End If
  End If

  '
  ' Progress step 5b - Apply Menu View permissions.
  '
  If fOK Then
  
    ' Apply the View menu Permissions.
    OutputCurrentProcess2 "Menus"
    fOK = ApplyChanges_ViewMenuPermissions
    gobjProgress.UpdateProgress
    
    ' Flag any errors.
    If Not fOK Then
      gobjProgress.Visible = False
      MsgBox "Error applying menu view permissions.", vbOKOnly + vbExclamation, Application.Name
    End If
  End If

  '
  ' Progress step 6 - Create new User Logins.
  '
  If fOK Then
    
    ' Create the new User Logins.
    OutputCurrentProcess "New Logins"
    fOK = ApplyChanges_NewUserLogins
    gobjProgress.UpdateProgress False
  
    ' Flag any errors.
    If Not fOK Then
      gobjProgress.Visible = False
      MsgBox "Error creating new user logins.", vbOKOnly + vbExclamation, Application.Name
    End If
  End If

  
  ' Apply fixed role security settings
  If fOK Then
  
    ' Create the new User Logins.
    OutputCurrentProcess "Applying Roles"
    fOK = ApplyChanges_ApplyRolesToLogins
    gobjProgress.UpdateProgress False
  
    ' Flag any errors.
    If Not fOK Then
      gobjProgress.Visible = False
      MsgBox "Error applying roles to logins.", vbOKOnly + vbExclamation, Application.Name
    End If
  End If


  '
  ' Progress step 6A - Apply Login field names to personnel records.
  '
  If fOK Then
  
    ' Apply Database Ownership.
    OutputCurrentProcess "Updating SS Logins"
    fOK = ApplyChanges_UpdatePersonnelRecords
    gobjProgress.UpdateProgress False
  
    ' Flag any errors.
    If Not fOK Then
      gobjProgress.Visible = False
      MsgBox "Error applying login names to Personnel records.", vbOKOnly + vbExclamation, Application.Name
    End If
  End If

  '
  ' Progress step 6B - Apply Login field names to personnel records.
  '
  If fOK Then
    
    ' Apply Force Password changes
    OutputCurrentProcess "Password Options"
    fOK = ApplyChanges_ForceChangePasswords
    gobjProgress.UpdateProgress False

    ' Flag any errors.
    If Not fOK Then
      gobjProgress.Visible = False
      MsgBox "Error applying forced password change Personnel records.", vbOKOnly + vbExclamation, Application.Name
    End If
  End If

  '
  ' Progress step 7 - Apply User moves.
  '
  If fOK Then
    ' Update the progress bar.
    OutputCurrentProcess "Users"
    gobjProgress.UpdateProgress False
    DoEvents
  
    ' Move Users.
    OutputCurrentProcess2 "Moves", 3
    fOK = ApplyChanges_MoveUsers
    gobjProgress.UpdateProgress2
  
    ' Flag any errors.
    If Not fOK Then
      gobjProgress.Visible = False
      MsgBox "Error moving users.", vbOKOnly + vbExclamation, Application.Name
    End If
  End If

  '
  ' Progress step 9 - Apply User additions.
  '
  If fOK Then
  
    ' Create new Users.
    OutputCurrentProcess2 "Additions"
    fOK = ApplyChanges_NewUsers
    gobjProgress.UpdateProgress2
    
    ' Flag any errors.
    If Not fOK Then
      gobjProgress.Visible = False
      MsgBox "Error creating new users.", vbOKOnly + vbExclamation, Application.Name
    End If
  End If

  '
  ' Progress step 10 - Apply User Group (Role) deletions.
  '
  If fOK Then
  
    ' Apply User Group (Role) deletions.
    OutputCurrentProcess "Group Removals"
    fOK = ApplyChanges_DeleteGroups
    gobjProgress.UpdateProgress False
  
    ' Flag any errors.
    If Not fOK Then
      gobjProgress.Visible = False
      MsgBox "Error deleting user groups.", vbOKOnly + vbExclamation, Application.Name
    End If
  End If

  '
  ' Progress step 11 - Clean up the 'public' User Group (Role).
  '
  If fOK Then
    
    ' Clean up the databases.
    OutputCurrentProcess "Surface Area Compression"
    fOK = ApplyChanges_CleanUp
    gobjProgress.UpdateProgress False
  
    ' Flag any errors.
    If Not fOK Then
      gobjProgress.Visible = False
      MsgBox "Error compressing surface area.", vbOKOnly + vbExclamation, Application.Name
    End If
  End If
  
  '
  ' Progress step 12 - Create the view validation stored procedures.
  '
  If fOK Then
  
    ' Apply the Table and Table Column permissions.
    OutputCurrentProcess "Creating Validation Procedures"
    fOK = ApplyChanges_CreateViewValidationStoredProcedures
    gobjProgress.UpdateProgress False
  
    ' Flag any errors.
    If Not fOK Then
      gobjProgress.Visible = False
      MsgBox "Error creating view validation stored procedures.", vbOKOnly + vbExclamation, Application.Name
    End If
  End If
  
  ' JPD20030206 Fault 5022
  '
  ' Progress step 14 - Apply Database Ownership.
  '
  If fOK Then
  
    ' Apply Database Ownership.
    OutputCurrentProcess "Applying Database Ownership"
    fOK = ApplyChanges_DatabaseOwnership
    gobjProgress.UpdateProgress False
  
    ' Flag any errors.
    If Not fOK Then
      gobjProgress.Visible = False
      MsgBox "Error applying database ownership.", vbOKOnly + vbExclamation, Application.Name
    End If
  End If
  
  
  ' Apply process admin right to logins if necessary
  If fOK Then
    OutputCurrentProcess "Granting process admin rights"
    fOK = ApplyPostSaveProcessing
    gobjProgress.UpdateProgress False
  
    ' Flag any errors.
    If Not fOK Then
      gobjProgress.Visible = False
      MsgBox "Error granting process admin rights.", vbOKOnly + vbExclamation, Application.Name
    End If
  End If
  
  
  'TM20010920 Fault 2530
  '
  ' Progress step 15 - If not logging out then refresh the groups collection.
  '
  'TM20011003 Fault 2903
  'Only re-initialise if the previous changes have successfully completed.
  OutputCurrentProcess "Reinitialising Interface"
  
  If fOK Then
    If Not blnIsExiting Then
      Set gObjGroups = Nothing
      'Initialise the collection of user groups (roles).
      fOK = InitialiseGroupsCollection(gObjGroups)
    End If
  End If
  
TidyUpAndExit:
  UnlockDatabase lckSaving
  
  If fOK Then
    If gbShiftSave Then
      frmMain.Caption = Replace(frmMain.Caption, " [Shift Save]", "")
      gbShiftSave = False   'MH20060801 Fault 11372
    End If
    AuditAccess "Save", "Security"
  End If

  ' Get rid of the progress bar.
  gobjProgress.CloseProgress
  
  'RH 21/08/00 - BUG 788. Re-enable all forms while we save changes - this
  '              prevents users from col editing etc whilst saving ! a good idea
  '              methinks !
  For iCount = 0 To Forms.Count - 1
    Forms(iCount).Enabled = True
  Next iCount
  
  'RH 02/10/00 - BUG 1044 - Enable the X button as finished/aborted saving
  EnableCloseButton frmMain.hWnd, True

  ApplyChanges = fOK
  Exit Function
  
ErrorTrap:
  ' Get rid of the progress bar.
  gobjProgress.Visible = False
  MsgBox Err.Description, vbCritical + vbOKOnly, App.Title
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Sub OutputCurrentProcess(strInput As String, Optional blnOverwriteExisting As Boolean)

  'Ignore any errors in here...
  On Local Error GoTo LocalErr
  
  
  Dim strFileName As String

  If Trim$(strInput) <> vbNullString Then
    gobjProgress.Bar1Caption = Trim$(strInput) & " ..."
    gobjProgress.Bar2MaxValue = 100000
    gobjProgress.Bar2Value = 0
    gobjProgress.ResetBar2
    gobjProgress.UpdateProgress2
    gobjProgress.Bar2Value = 0
  End If

  DoEvents

  strFileName = App.Path & "\savelog.txt"

  If blnOverwriteExisting Then
    If Dir(strFileName) <> vbNullString Then
      Kill strFileName
    End If
    
    Open strFileName For Append As #99
    Print #99, "Server    : " & gsServerName
    Print #99, "Database  : " & gsDatabaseName
    Print #99, "Username  : " & gsUserName
    Print #99, "Version   : " & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
    Print #99, vbNullString

  Else
    Open strFileName For Append As #99

  End If
  
  Print #99, Now & "  " & strInput
  Close #99

LocalErr:
  Err.Clear

End Sub

Private Sub OutputCurrentProcess2(strInput As String, Optional ByVal iMaxSteps As Integer)

  'Ignore any errors in here...
  On Local Error GoTo LocalErr
  
  Dim strFileName As String

  If iMaxSteps > 0 Then
    'gobjProgress.ResetBar2
    gobjProgress.Bar2MaxValue = iMaxSteps
  End If

  If Trim$(strInput) <> vbNullString Then
    gobjProgress.Bar2Caption = strInput & " ..."
  End If

  DoEvents

  strFileName = App.Path & "\savelog.txt"

  If Trim$(strInput) <> vbNullString Then
    Open strFileName For Append As #99
    Print #99, Now & "  " & strInput
    Close #99
  End If

LocalErr:
  Err.Clear

End Sub
Private Function ApplyChanges_TidyChildTablePermissions() As Boolean
  
  Dim objGroup As SecurityGroup
  Dim sGroupName As String
  Dim fSysSecManager As Boolean
  Dim sSQL As String
  Dim rsChildren As New ADODB.Recordset
  Dim asChildViews() As String
  Dim iLoop1 As Integer
  Dim fOK As Boolean
  Dim cmdIsSysSecMgr As New ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim sSysSecMgrRoleList As New SecurityMgr.clsStringBuilder
  Dim sRevokeRoleList As New SecurityMgr.clsStringBuilder
  Dim sSQLCommands As New SecurityMgr.clsStringBuilder
    
  fOK = True
  
  ' Column 1 = role name
  ' Column 2 = child view name
  ReDim asChildViews(2, 0)
  
  sSQL = "SELECT ASRSysChildViews2.Role, ASRSysChildViews2.ChildViewID, ASRSysTables.tableName" & _
    " FROM ASRSysChildViews2" & _
    " INNER JOIN ASRSysTables ON ASRSysChildViews2.tableID = ASRSysTables.tableID"

  rsChildren.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
  With rsChildren
    Do While (Not .EOF)
      ReDim Preserve asChildViews(2, UBound(asChildViews, 2) + 1)
      
      asChildViews(1, UBound(asChildViews, 2)) = UCase(!Role)
      asChildViews(2, UBound(asChildViews, 2)) = Left("ASRSysCV" & Trim(Str(!childViewID)) & "#" & Replace(!TableName, " ", "_") & "#" & Replace(!Role, " ", "_"), 255)
  
      .MoveNext
    Loop
  
    .Close
  End With
  Set rsChildren = Nothing
  
  sSysSecMgrRoleList.TheString = vbNullString
  sSQLCommands.TheString = vbNullString

  For iLoop1 = 1 To UBound(asChildViews, 2)
    sRevokeRoleList.TheString = vbNullString
  
    For Each objGroup In gObjGroups
      
     sGroupName = "[" & objGroup.Name & "]"
      
'      If objGroup.Initialised Then
    
        If iLoop1 = 1 Then
          If objGroup.Initialised Then
            ' Determine if the current group is granted access to the System or Security Managers.
            fSysSecManager = objGroup.SystemPermissions.Item("P_MODULEACCESS_SYSTEMMANAGER").Allowed Or _
              objGroup.SystemPermissions.Item("P_MODULEACCESS_SECURITYMANAGER").Allowed
          Else
          
             With cmdIsSysSecMgr
               .CommandText = "dbo.sp_ASRIsSysSecMgr"
               .CommandType = adCmdStoredProc
               .CommandTimeout = 0
               Set .ActiveConnection = gADOCon
           
               Set pmADO = .CreateParameter("GroupName", adVarChar, adParamInput, 256)
               .Parameters.Append pmADO
               pmADO.Value = objGroup.Name
           
               Set pmADO = .CreateParameter("Result", adBoolean, adParamOutput)
               .Parameters.Append pmADO
          
               .Execute
           
               fSysSecManager = .Parameters(1).Value
             End With
             Set cmdIsSysSecMgr = Nothing
          
          End If
          
          If fSysSecManager Then
            sSysSecMgrRoleList.Append IIf(sSysSecMgrRoleList.Length <> 0, ",", vbNullString) & sGroupName
          End If
        End If
        
        If asChildViews(1, iLoop1) <> UCase$(objGroup.Name) Then
          sRevokeRoleList.Append IIf(sRevokeRoleList.Length = 0, vbNullString, ",") & sGroupName
        End If
'      End If
    Next objGroup
  
    If sRevokeRoleList.Length <> 0 Then
      sSQLCommands.Append "REVOKE DELETE, INSERT, SELECT, UPDATE ON " & asChildViews(2, iLoop1) & " TO " & sRevokeRoleList.ToString & vbNewLine
    End If
    
    If sSysSecMgrRoleList.Length <> 0 Then
      sSQLCommands.Append "GRANT DELETE, INSERT, SELECT, UPDATE ON " & asChildViews(2, iLoop1) & " TO " & sSysSecMgrRoleList.ToString & vbNewLine
    End If
  Next iLoop1
  
  ' Fire off all the grants & revokes
  If sSQLCommands.Length <> 0 Then
    gADOCon.Execute sSQLCommands.ToString, , adExecuteNoRecords
  End If

TidyUpAndExit:
  Set objGroup = Nothing
  ApplyChanges_TidyChildTablePermissions = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Function IsSQLLoginNameInUse(psUserLogin As String) As Boolean

  On Error GoTo ErrorTrap
  
  Dim sSQL As String
  Dim rsRecords As New ADODB.Recordset
  Dim bIsSQLSystemAdmin As Boolean
  Dim rsUser As New ADODB.Recordset

  sSQL = "SELECT master.dbo.syslogins.loginname " & _
    " FROM master.dbo.syslogins" & _
    " WHERE master.dbo.syslogins.loginname = '" & Replace(psUserLogin, "'", "''") & "'"

  rsRecords.Open sSQL, gADOCon, adOpenKeyset, adLockReadOnly

  IsSQLLoginNameInUse = IIf(rsRecords.EOF And rsRecords.BOF, False, True)
  rsRecords.Close
   
  'NPG20090423 Fault 13672
  ' Is the user a system administrator on the server or is logged in as 'sa'
  sSQL = "SELECT IS_SRVROLEMEMBER('sysadmin') AS Permission"
  rsUser.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  bIsSQLSystemAdmin = IIf(rsUser!Permission = 1, True, False)
  rsUser.Close
  
  If Not bIsSQLSystemAdmin Then IsSQLLoginNameInUse = True

TidyUpAndExit:
  Set rsRecords = Nothing
  Exit Function
  
ErrorTrap:
  IsSQLLoginNameInUse = False
  Resume TidyUpAndExit
  
End Function

Public Function IsAlreadyNewUser(psUserLogin As String, pobjSecurityGroups As SecurityGroups, Optional ByRef pstrFoundInGroup As String) As Boolean
  
  ' Check the name passed has not been created as a new user already.
'  Dim sSql As String
  Dim objSecurityGroup As SecurityGroup
  Dim objSecurityUser As SecurityUser
  
  IsAlreadyNewUser = False

  ' Check if the given name is already used as a username or login in the users collections.
  If Not IsAlreadyNewUser Then
    For Each objSecurityGroup In pobjSecurityGroups
      
      ' See if the user collection has been initialised
      If Not objSecurityGroup.Users_Initialised Then
        InitialiseUsersCollection objSecurityGroup
      End If
      
      For Each objSecurityUser In objSecurityGroup.Users
        If ((UCase$(objSecurityUser.UserName) = UCase$(psUserLogin)) Or _
          (UCase$(objSecurityUser.Login) = UCase$(psUserLogin))) And _
           objSecurityUser.MovedUserTo = vbNullString Then
          
          IsAlreadyNewUser = objSecurityUser.NewUser
          pstrFoundInGroup = objSecurityGroup.Name
          Exit For
        End If
      Next
      Set objSecurityUser = Nothing
      
      If IsAlreadyNewUser Then Exit For
    Next
  
    Set objSecurityGroup = Nothing
  End If

End Function

Public Function IsUserNameInUse(psUserLogin As String, pobjSecurityGroups As SecurityGroups, Optional ByRef pstrFoundInGroup As String) As Boolean
  ' Check the name passed in is not currently in use as a group or user name.
'  Dim sSql As String
  Dim objSecurityGroup As SecurityGroup
  Dim objSecurityUser As SecurityUser
  
  IsUserNameInUse = False
  
  ' Check if the given name is already used by a group in the groups collection.
  For Each objSecurityGroup In pobjSecurityGroups
    If UCase$(objSecurityGroup.Name) = UCase$(psUserLogin) Then
      IsUserNameInUse = True 'Not objSecurityGroup.DeleteGroup
      Exit For
    End If
  Next
  Set objSecurityGroup = Nothing
  
  ' Check if the given name is already used as a username or login in the users collections.
  If Not IsUserNameInUse Then
    For Each objSecurityGroup In pobjSecurityGroups
      
      If Not objSecurityGroup.Users_Initialised Then
        InitialiseUsersCollection objSecurityGroup
      End If
      
      ' See if the user collection has been initialised
      For Each objSecurityUser In objSecurityGroup.Users
        If ((UCase$(objSecurityUser.UserName) = UCase$(psUserLogin)) Or _
          (UCase$(objSecurityUser.Login) = UCase$(psUserLogin))) And _
           objSecurityUser.MovedUserTo = vbNullString Then
          
          IsUserNameInUse = (Not objSecurityUser.DeleteUser) And (Not objSecurityUser.NewUser)
          pstrFoundInGroup = objSecurityGroup.Name
          Exit For
        End If
      Next
      Set objSecurityUser = Nothing
      
      If IsUserNameInUse Then Exit For
    Next
  
    Set objSecurityGroup = Nothing
  End If

End Function

Public Function CheckVersion() As Boolean
  ' Check that the database version is the right one for this application's version.
  ' If everything matches then return TRUE.
  ' If not, try to update the database.
  ' If the database can be updated return TRUE, else return FALSE.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fVersionOK As Boolean
  Dim iMajorAppVersion As Integer
  Dim iMinorAppVersion As Integer
  Dim iRevisionAppVersion As Integer
  Dim sDBVersion As String
  Dim blnNewStyleVersionNo As Boolean
  Dim fRefreshStoredProcedures As Boolean
  
  
  fOK = True
  fVersionOK = False
  
  sDBVersion = GetDBVersion
  'sDBVersion = GetSystemSetting("Database", "Version", vbNullString)
  
  
  If LenB(sDBVersion) = 0 Then
    fOK = False
    
    MsgBox "Error checking version compatibility." & vbNewLine & _
      "Version number not found.", _
      vbOKOnly + vbExclamation, Application.Name
  Else
    iMajorAppVersion = Val(Split(sDBVersion, ".")(0))
    iMinorAppVersion = Val(Split(sDBVersion, ".")(1))
    
    blnNewStyleVersionNo = (UBound(Split(sDBVersion, ".")) = 1)
    If Not blnNewStyleVersionNo Then
      iRevisionAppVersion = Val(Split(sDBVersion, ".")(2))
    End If
  End If
  
  
  If fOK Then
    ' Check the System Manager version against the one for the current database.
    If (App.Major = iMajorAppVersion) And _
      (App.Minor = iMinorAppVersion) And _
      (App.Revision = iRevisionAppVersion Or blnNewStyleVersionNo) Then
      ' Application and database versions match.
      fVersionOK = True
    End If
  End If
  
  
  If fOK Then
    ' Check the System Manager version against the one for the current database.
    ' Application is too old for the database.
    If (App.Major < iMajorAppVersion) Or _
      ((App.Major = iMajorAppVersion) And (App.Minor < iMinorAppVersion)) Or _
      ((App.Major = iMajorAppVersion) And (App.Minor = iMinorAppVersion) And (App.Revision < iRevisionAppVersion And Not blnNewStyleVersionNo)) Then
      fOK = False
    
      MsgBox "The application is out of date." & vbNewLine & _
        "Contact your administrator for a new version of the application." & vbNewLine & vbNewLine & _
        "Database Name : " & gsDatabaseName & vbNewLine & _
        "Database Version : " & sDBVersion & vbNewLine & vbNewLine & _
        "Application Version : " & CStr(App.Major) & "." & CStr(App.Minor), _
        vbExclamation + vbOKOnly, Application.Name
    End If
  End If
  
  
  If fOK Then
    ' Database is too old for the application. Try to update the database.
    If (App.Major > iMajorAppVersion) Or _
      ((App.Major = iMajorAppVersion) And (App.Minor > iMinorAppVersion)) Or _
      ((App.Major = iMajorAppVersion) And (App.Minor = iMinorAppVersion) And (App.Revision > iRevisionAppVersion And Not blnNewStyleVersionNo)) Then

      MsgBox "The database is out of date." & vbNewLine & _
       "Please ask the System Administrator to update the database in the System Manager." & vbNewLine & vbNewLine & _
       "Database Name : " & gsDatabaseName & vbNewLine & _
       "Database Version : " & sDBVersion & vbNewLine & vbNewLine & _
       "Application Version : " & CStr(App.Major) & "." & CStr(App.Minor), _
       vbExclamation + vbOKOnly, Application.Name
      fVersionOK = False
      fOK = fVersionOK
    
    End If
  End If
  
  
  If fOK Then
    ' Check if a new version of the application is required due to an Intranet update
    
    sDBVersion = GetSystemSetting("Database", "Minimum Version", vbNullString)
    If LenB(sDBVersion) <> 0 Then
      
      iMajorAppVersion = Val(Split(sDBVersion, ".")(0))
      iMinorAppVersion = Val(Split(sDBVersion, ".")(1))
      
      blnNewStyleVersionNo = (UBound(Split(sDBVersion, ".")) = 1)
      If Not blnNewStyleVersionNo Then
        iRevisionAppVersion = Val(Split(sDBVersion, ".")(2))
      End If
      
      If (App.Major < iMajorAppVersion) Or _
        ((App.Major = iMajorAppVersion) And (App.Minor < iMinorAppVersion)) Or _
        ((App.Major = iMajorAppVersion) And (App.Minor = iMinorAppVersion) And (App.Revision < iRevisionAppVersion And Not blnNewStyleVersionNo)) Then

        fVersionOK = False
        MsgBox "The application is now out of date due to an update to the intranet module." & vbNewLine & _
          "Please contact your administrator for a new version of the application.", _
          vbOKOnly + vbExclamation, Application.Name
        fOK = fVersionOK

      End If
    End If
  End If
  
  ' If the platform has changed tag the refresh stored procedures flag
  If fOK Then
    fOK = CheckPlatform
  End If
  
  If fOK Then
    fRefreshStoredProcedures = (GetSystemSetting("Database", "RefreshStoredProcedures", 0) = 1)
  
    If fRefreshStoredProcedures Then
      ' Tell the user that the System manager needs to be run, and changes saved
      ' before this application can run.
      fOK = False
  
      MsgBox "The database is out of date." & vbNewLine & _
        "Please ask the System Administrator to save the update in the System Manager.", _
        vbOKOnly + vbExclamation, Application.Name
    End If
  End If
  
  ' Do we enable UDF functions on this installation
  gbEnableUDFFunctions = EnableUDFFunctions
  
  ' If fOK and fVersionOK are true then the application and databases versions match.
TidyUpAndExit:
  If Not fOK Then
    fVersionOK = False
    Screen.MousePointer = vbDefault
  End If
  
  CheckVersion = fVersionOK
  Exit Function
  
ErrorTrap:
  If (Err.Number = 75) Or (Err.Number = 76) Then
    MsgBox "The database is out of date." & vbNewLine & _
      "Unable to update the database as the required update script cannot be found.", _
      vbOKOnly + vbExclamation, Application.Name
  Else
    MsgBox "Error checking database and application versions." & vbNewLine & _
      Err.Description, _
      vbOKOnly + vbExclamation, Application.Name
  End If
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function ApplyChanges_NewGroups() As Boolean
  ' Create the new user groups (roles) in the SQL Server database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim objGroup As SecurityGroup
  Dim avAccess As Variant
  Dim iLoop As Integer
  Dim fBatchJobsHidden As Boolean
  Dim sHiddenJobTypes As String
  Dim fChanged As Boolean
  Dim objTable As SecurityTable
  
  fOK = True
  fBatchJobsHidden = False
  sHiddenJobTypes = vbNullString
  
  ' Find any new user groups (roles).
  For Each objGroup In gObjGroups
  
    With objGroup
      fChanged = .Changed
      
      ' Check if the user group (role) is new.
      If fChanged And .NewGroup Then
          
        ' Create the user group (role) in the SQL Server database.
        sSQL = "sp_addrole '" & .Name & "', [dbo]"
        AuditGroup .Name, "User Group Added"
        gADOCon.Execute sSQL, , adExecuteNoRecords
        
        ' Initialize the System Permission records for the new user group.
        sSQL = "DELETE FROM ASRSysGroupPermissions" & _
          " WHERE groupName = '" & .Name & "'"
        gADOCon.Execute sSQL, , adExecuteNoRecords
        
        sSQL = "INSERT INTO ASRSysGroupPermissions" & _
          " (itemID, groupName, permitted)" & _
          " (SELECT ASRSysPermissionItems.itemID, '" & .Name & "', 0" & _
          " FROM ASRSysPermissionItems)"
        gADOCon.Execute sSQL, , adExecuteNoRecords

        ' JDM - 12/02/2004 - Fault 7652 - Force table permissions to be set on a new group
        For Each objTable In .Tables
          objTable.Changed = True
        Next objTable

        'MH20030205 Fault 4890
        If .OriginalName <> vbNullString Then
          sSQL = "UPDATE ASRSysBatchJobName " & _
                 "SET roletoprompt = '" & .Name & "' " & _
                 "WHERE roletoprompt = '" & .OriginalName & "'"
          gADOCon.Execute sSQL, , adExecuteNoRecords
        End If

        'JPD 20071203 Faults 12580, 12670, 12671
        If Len(.CopyGroup) > 0 Then
          sSQL = "INSERT INTO ASRSysSSIHiddenGroups" & _
            " (linkID, groupName)" & _
            " (SELECT SSIHG.linkID, '" & .Name & "'" & _
            " FROM ASRSysSSIHiddenGroups SSIHG" & _
            " WHERE SSIHG.groupName = '" & Replace(.CopyGroup, "'", "''") & "')"
          gADOCon.Execute sSQL, , adExecuteNoRecords
        Else
          sSQL = "INSERT INTO ASRSysSSIHiddenGroups" & _
            " (linkID, groupName)" & _
            " (SELECT SSIL.id, '" & .Name & "'" & _
            " FROM ASRSysSSIntranetLinks SSIL)"
          gADOCon.Execute sSQL, , adExecuteNoRecords
        End If

        ' Reset the 'new' and 'changed' flags on this user group (role).
        .NewGroup = False
        .Changed = False
      End If
      
      If fChanged And _
        ((LenB(.AccessCopyGroup) <> 0) Or (IsArray(.AccessConfiguration))) Then
        ' Update the utility/report access for the new group,
        ' as defined by its 'copy' group, or the defined configuration.
        
        RemoveGroupAccessRecords (.Name)
        
        If LenB(.AccessCopyGroup) <> 0 Then
          ' NEWACCESS - needs to be updated as each report/utility is updated for the new access.
          sSQL = "INSERT INTO ASRSysBatchJobAccess" & _
            " (groupName, access, id)" & _
            " (SELECT '" & .Name & "', access, id" & _
            "   FROM ASRSysBatchJobAccess" & _
            "   WHERE groupName = '" & .AccessCopyGroup & "')"
          gADOCon.Execute sSQL, , adExecuteNoRecords
          
          sSQL = "INSERT INTO ASRSysCrossTabAccess" & _
            " (groupName, access, id)" & _
            " (SELECT '" & .Name & "', access, id" & _
            "   FROM ASRSysCrossTabAccess" & _
            "   WHERE groupName = '" & .AccessCopyGroup & "')"
          gADOCon.Execute sSQL, , adExecuteNoRecords
    
          sSQL = "INSERT INTO ASRSysCalendarReportAccess" & _
            " (groupName, access, id)" & _
            " (SELECT '" & .Name & "', access, id" & _
            "   FROM ASRSysCalendarReportAccess" & _
            "   WHERE groupName = '" & .AccessCopyGroup & "')"
          gADOCon.Execute sSQL, , adExecuteNoRecords
          
          sSQL = "INSERT INTO ASRSysCustomReportAccess" & _
            " (groupName, access, id)" & _
            " (SELECT '" & .Name & "', access, id" & _
            "   FROM ASRSysCustomReportAccess" & _
            "   WHERE groupName = '" & .AccessCopyGroup & "')"

          gADOCon.Execute sSQL, , adExecuteNoRecords
          
          sSQL = "INSERT INTO ASRSysDataTransferAccess" & _
            " (groupName, access, id)" & _
            " (SELECT '" & .Name & "', access, id" & _
            "   FROM ASRSysDataTransferAccess" & _
            "   WHERE groupName = '" & .AccessCopyGroup & "')"
          gADOCon.Execute sSQL, , adExecuteNoRecords
          
          sSQL = "INSERT INTO ASRSysExportAccess" & _
            " (groupName, access, id)" & _
            " (SELECT '" & .Name & "', access, id" & _
            "   FROM ASRSysExportAccess" & _
            "   WHERE groupName = '" & .AccessCopyGroup & "')"
          gADOCon.Execute sSQL, , adExecuteNoRecords
          
          sSQL = "INSERT INTO ASRSysGlobalAccess" & _
            " (groupName, access, id)" & _
            " (SELECT '" & .Name & "', access, id" & _
            "   FROM ASRSysGlobalAccess" & _
            "   WHERE groupName = '" & .AccessCopyGroup & "')"
          gADOCon.Execute sSQL, , adExecuteNoRecords
          
          sSQL = "INSERT INTO ASRSysImportAccess" & _
            " (groupName, access, id)" & _
            " (SELECT '" & .Name & "', access, id" & _
            "   FROM ASRSysImportAccess" & _
            "   WHERE groupName = '" & .AccessCopyGroup & "')"
          gADOCon.Execute sSQL, , adExecuteNoRecords
          
          sSQL = "INSERT INTO ASRSysMailMergeAccess" & _
            " (groupName, access, id)" & _
            " (SELECT '" & .Name & "', access, ID" & _
            "   FROM ASRSysMailMergeAccess" & _
            "   WHERE groupName = '" & .AccessCopyGroup & "')"
          gADOCon.Execute sSQL, , adExecuteNoRecords
          
          sSQL = "INSERT INTO ASRSysRecordProfileAccess" & _
            " (groupName, access, id)" & _
            " (SELECT '" & .Name & "', access, id" & _
            "   FROM ASRSysRecordProfileAccess" & _
            "   WHERE groupName = '" & .AccessCopyGroup & "')"
          gADOCon.Execute sSQL, , adExecuteNoRecords
          
          sSQL = "INSERT INTO ASRSysMatchReportAccess" & _
            " (groupName, access, id)" & _
            " (SELECT '" & .Name & "', access, ID" & _
            "   FROM ASRSysMatchReportAccess" & _
            "   WHERE groupName = '" & .AccessCopyGroup & "')"
          gADOCon.Execute sSQL, , adExecuteNoRecords
          
          ' Reset the 'accessCopyGroup' property as the group's access
          ' is now saved and no longer needs to be copied.
          .AccessCopyGroup = vbNullString
        Else
          avAccess = .AccessConfiguration
          If IsArray(avAccess) Then
            If UBound(avAccess, 2) > 0 Then
              For iLoop = 1 To UBound(avAccess, 2)
                Select Case CInt(avAccess(1, iLoop))
                  Case utlAbsenceBreakdown
                  
                  Case utlBatchJob
                    sSQL = "INSERT INTO ASRSysBatchJobAccess" & _
                      " (groupName, access, id)" & _
                      " (SELECT '" & .Name & "', '" & CStr(avAccess(2, iLoop)) & "', id" & _
                      "   FROM ASRSysBatchJobName)"
                    gADOCon.Execute sSQL, , adExecuteNoRecords

                    fBatchJobsHidden = (CStr(avAccess(2, iLoop)) = ACCESS_HIDDEN)
                    
                  Case utlBradfordFactor
                  
                  Case utlCalendarReport
                    sSQL = "INSERT INTO ASRSysCalendarReportAccess" & _
                      " (groupName, access, id)" & _
                      " (SELECT '" & .Name & "', '" & CStr(avAccess(2, iLoop)) & "', ID" & _
                      "   FROM ASRSysCalendarReports)"
                    gADOCon.Execute sSQL, , adExecuteNoRecords

                    ' If the utility/report is being made hidden then any  batch jobs that
                    ' contain this utility/report must also be made hidden.
                    If CStr(avAccess(2, iLoop)) = ACCESS_HIDDEN Then
                      sHiddenJobTypes = sHiddenJobTypes & _
                        IIf(LenB(sHiddenJobTypes) <> 0, ",", vbNullString) & _
                        "'Calendar Report'"
                    End If
                  
                  Case utlCrossTab
                    sSQL = "INSERT INTO ASRSysCrossTabAccess" & _
                      " (groupName, access, id)" & _
                      " (SELECT '" & .Name & "', '" & CStr(avAccess(2, iLoop)) & "', crossTabID" & _
                      "   FROM ASRSysCrossTab)"
                    gADOCon.Execute sSQL, , adExecuteNoRecords

                    ' If the utility/report is being made hidden then any  batch jobs that
                    ' contain this utility/report must also be made hidden.
                    If CStr(avAccess(2, iLoop)) = ACCESS_HIDDEN Then
                      sHiddenJobTypes = sHiddenJobTypes & _
                        IIf(LenB(sHiddenJobTypes) <> 0, ",", vbNullString) & _
                        "'Cross Tab'"
                    End If
                                    
                  Case utlCustomReport
                    sSQL = "INSERT INTO ASRSysCustomReportAccess" & _
                      " (groupName, access, id)" & _
                      " (SELECT '" & .Name & "', '" & CStr(avAccess(2, iLoop)) & "', ID" & _
                      "   FROM ASRSysCustomReportsName)"
                    gADOCon.Execute sSQL, , adExecuteNoRecords

                    ' If the utility/report is being made hidden then any  batch jobs that
                    ' contain this utility/report must also be made hidden.
                    If CStr(avAccess(2, iLoop)) = ACCESS_HIDDEN Then
                      sHiddenJobTypes = sHiddenJobTypes & _
                        IIf(LenB(sHiddenJobTypes) <> 0, ",", vbNullString) & _
                        "'Custom Report'"
                    End If
                  
                  Case utlDataTransfer
                    sSQL = "INSERT INTO ASRSysDataTransferAccess" & _
                      " (groupName, access, id)" & _
                      " (SELECT '" & .Name & "', '" & CStr(avAccess(2, iLoop)) & "', dataTransferID" & _
                      "   FROM ASRSysDataTransferName)"
                    gADOCon.Execute sSQL, , adExecuteNoRecords

                    ' If the utility/report is being made hidden then any  batch jobs that
                    ' contain this utility/report must also be made hidden.
                    If CStr(avAccess(2, iLoop)) = ACCESS_HIDDEN Then
                      sHiddenJobTypes = sHiddenJobTypes & _
                        IIf(LenB(sHiddenJobTypes) <> 0, ",", vbNullString) & _
                        "'Data Transfer'"
                    End If
                  
                  Case utlLabel
                    sSQL = "INSERT INTO ASRSysMailMergeAccess" & _
                      " (groupName, access, id)" & _
                      " (SELECT '" & .Name & "', '" & CStr(avAccess(2, iLoop)) & "', mailMergeID" & _
                      "   FROM ASRSysMailMergeName" & _
                      "   WHERE isLabel = 1)"
                    gADOCon.Execute sSQL, , adExecuteNoRecords

                    ' If the utility/report is being made hidden then any  batch jobs that
                    ' contain this utility/report must also be made hidden.
                    If CStr(avAccess(2, iLoop)) = ACCESS_HIDDEN Then
                      sHiddenJobTypes = sHiddenJobTypes & _
                        IIf(LenB(sHiddenJobTypes) <> 0, ",", vbNullString) & _
                        "'Envelopes & Labels'"
                    End If
                  
                  Case utlExport
                    sSQL = "INSERT INTO ASRSysExportAccess" & _
                      " (groupName, access, id)" & _
                      " (SELECT '" & .Name & "', '" & CStr(avAccess(2, iLoop)) & "', ID" & _
                      "   FROM ASRSysExportName)"
                    gADOCon.Execute sSQL, , adExecuteNoRecords

                    ' If the utility/report is being made hidden then any  batch jobs that
                    ' contain this utility/report must also be made hidden.
                    If CStr(avAccess(2, iLoop)) = ACCESS_HIDDEN Then
                      sHiddenJobTypes = sHiddenJobTypes & _
                        IIf(LenB(sHiddenJobTypes) <> 0, ",", vbNullString) & _
                        "'Export'"
                    End If
                    
                  Case UtlGlobalAdd
                    sSQL = "INSERT INTO ASRSysGlobalAccess" & _
                      " (groupName, access, id)" & _
                      " (SELECT '" & .Name & "', '" & CStr(avAccess(2, iLoop)) & "', functionID" & _
                      "   FROM ASRSysGlobalFunctions" & _
                      "   WHERE type = 'A')"
                    gADOCon.Execute sSQL, , adExecuteNoRecords

                    ' If the utility/report is being made hidden then any  batch jobs that
                    ' contain this utility/report must also be made hidden.
                    If CStr(avAccess(2, iLoop)) = ACCESS_HIDDEN Then
                      sHiddenJobTypes = sHiddenJobTypes & _
                        IIf(LenB(sHiddenJobTypes) <> 0, ",", vbNullString) & _
                        "'Global Add'"
                    End If
                    
                  Case utlGlobalDelete
                    sSQL = "INSERT INTO ASRSysGlobalAccess" & _
                      " (groupName, access, id)" & _
                      " (SELECT '" & .Name & "', '" & CStr(avAccess(2, iLoop)) & "', functionID" & _
                      "   FROM ASRSysGlobalFunctions" & _
                      "   WHERE type = 'D')"
                    gADOCon.Execute sSQL, , adExecuteNoRecords

                    ' If the utility/report is being made hidden then any  batch jobs that
                    ' contain this utility/report must also be made hidden.
                    If CStr(avAccess(2, iLoop)) = ACCESS_HIDDEN Then
                      sHiddenJobTypes = sHiddenJobTypes & _
                        IIf(LenB(sHiddenJobTypes) <> 0, ",", vbNullString) & _
                        "'Global Delete'"
                    End If
                    
                  Case utlGlobalUpdate
                    sSQL = "INSERT INTO ASRSysGlobalAccess" & _
                      " (groupName, access, id)" & _
                      " (SELECT '" & .Name & "', '" & CStr(avAccess(2, iLoop)) & "', functionID" & _
                      "   FROM ASRSysGlobalFunctions" & _
                      "   WHERE type = 'U')"
                    gADOCon.Execute sSQL, , adExecuteNoRecords

                    ' If the utility/report is being made hidden then any  batch jobs that
                    ' contain this utility/report must also be made hidden.
                    If CStr(avAccess(2, iLoop)) = ACCESS_HIDDEN Then
                      sHiddenJobTypes = sHiddenJobTypes & _
                        IIf(LenB(sHiddenJobTypes) <> 0, ",", vbNullString) & _
                        "'Global Update'"
                    End If
                    
                  Case utlImport
                    sSQL = "INSERT INTO ASRSysImportAccess" & _
                      " (groupName, access, id)" & _
                      " (SELECT '" & .Name & "', '" & CStr(avAccess(2, iLoop)) & "', ID" & _
                      "   FROM ASRSysImportName)"
                    gADOCon.Execute sSQL, , adExecuteNoRecords

                    ' If the utility/report is being made hidden then any  batch jobs that
                    ' contain this utility/report must also be made hidden.
                    If CStr(avAccess(2, iLoop)) = ACCESS_HIDDEN Then
                      sHiddenJobTypes = sHiddenJobTypes & _
                        IIf(LenB(sHiddenJobTypes) <> 0, ",", vbNullString) & _
                        "'Import'"
                    End If
                  
                  Case utlLabelType
                  
                  Case utlMailMerge
                    sSQL = "INSERT INTO ASRSysMailMergeAccess" & _
                      " (groupName, access, id)" & _
                      " (SELECT '" & .Name & "', '" & CStr(avAccess(2, iLoop)) & "', mailMergeID" & _
                      "   FROM ASRSysMailMergeName" & _
                      "   WHERE isLabel = 0)"
                    gADOCon.Execute sSQL, , adExecuteNoRecords

                    ' If the utility/report is being made hidden then any  batch jobs that
                    ' contain this utility/report must also be made hidden.
                    If CStr(avAccess(2, iLoop)) = ACCESS_HIDDEN Then
                      sHiddenJobTypes = sHiddenJobTypes & _
                        IIf(LenB(sHiddenJobTypes) <> 0, ",", vbNullString) & _
                        "'Mail Merge'"
                    End If
                  
                  Case utlMatchReport
                    sSQL = "INSERT INTO ASRSysMatchReportAccess" & _
                      " (groupName, access, id)" & _
                      " (SELECT '" & .Name & "', '" & CStr(avAccess(2, iLoop)) & "', matchReportID" & _
                      "   FROM ASRSysMatchReportName" & _
                      "   WHERE matchReportType = 0)"
                    gADOCon.Execute sSQL, , adExecuteNoRecords

                    ' If the utility/report is being made hidden then any  batch jobs that
                    ' contain this utility/report must also be made hidden.
                    If CStr(avAccess(2, iLoop)) = ACCESS_HIDDEN Then
                      sHiddenJobTypes = sHiddenJobTypes & _
                        IIf(LenB(sHiddenJobTypes) <> 0, ",", vbNullString) & _
                        "'Match Report'"
                    End If
                  
                  Case utlSuccession
                    sSQL = "INSERT INTO ASRSysMatchReportAccess" & _
                      " (groupName, access, id)" & _
                      " (SELECT '" & .Name & "', '" & CStr(avAccess(2, iLoop)) & "', matchReportID" & _
                      "   FROM ASRSysMatchReportName" & _
                      "   WHERE matchReportType = 1)"
                    gADOCon.Execute sSQL, , adExecuteNoRecords

                    ' If the utility/report is being made hidden then any  batch jobs that
                    ' contain this utility/report must also be made hidden.
                    If CStr(avAccess(2, iLoop)) = ACCESS_HIDDEN Then
                      sHiddenJobTypes = sHiddenJobTypes & _
                        IIf(LenB(sHiddenJobTypes) <> 0, ",", vbNullString) & _
                        "'Succession Planning'"
                    End If
                  
                  Case utlCareer
                    sSQL = "INSERT INTO ASRSysMatchReportAccess" & _
                      " (groupName, access, id)" & _
                      " (SELECT '" & .Name & "', '" & CStr(avAccess(2, iLoop)) & "', matchReportID" & _
                      "   FROM ASRSysMatchReportName" & _
                      "   WHERE matchReportType = 2)"
                    gADOCon.Execute sSQL, , adExecuteNoRecords

                    ' If the utility/report is being made hidden then any  batch jobs that
                    ' contain this utility/report must also be made hidden.
                    If CStr(avAccess(2, iLoop)) = ACCESS_HIDDEN Then
                      sHiddenJobTypes = sHiddenJobTypes & _
                        IIf(LenB(sHiddenJobTypes) <> 0, ",", vbNullString) & _
                        "'Career Progression'"
                    End If
                  
                  Case utlRecordProfile
                    sSQL = "INSERT INTO ASRSysRecordProfileAccess" & _
                      " (groupName, access, id)" & _
                      " (SELECT '" & .Name & "', '" & CStr(avAccess(2, iLoop)) & "', recordProfileID" & _
                      "   FROM ASRSysRecordProfileName)"
                    gADOCon.Execute sSQL, , adExecuteNoRecords

                    ' If the utility/report is being made hidden then any  batch jobs that
                    ' contain this utility/report must also be made hidden.
                    If CStr(avAccess(2, iLoop)) = ACCESS_HIDDEN Then
                      sHiddenJobTypes = sHiddenJobTypes & _
                        IIf(LenB(sHiddenJobTypes) <> 0, ",", vbNullString) & _
                        "'Record Profile'"
                    End If
                End Select
              Next iLoop
            
              If (Not fBatchJobsHidden) And (LenB(sHiddenJobTypes) <> 0) Then
                sSQL = "UPDATE ASRSysBatchJobAccess" & _
                  " SET access = '" & ACCESS_HIDDEN & "'" & _
                  " WHERE groupName = '" & .Name & "'" & _
                  "   AND id IN (SELECT DISTINCT batchJobNameID" & _
                  "              FROM ASRSysBatchJobDetails" & _
                  "              WHERE jobType IN (" & sHiddenJobTypes & "))"
                gADOCon.Execute sSQL, , adExecuteNoRecords
              End If
              
              ' Reset the 'AccessConfiguration' property as the group's access
              ' is now saved and no longer needs to be copied.
              ReDim avAccess(2, 0)
              .AccessConfiguration = avAccess
            End If
          End If
        End If
      End If
    End With
  Next objGroup
  
TidyUpAndExit:
  Set objGroup = Nothing
  ApplyChanges_NewGroups = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function ApplyChanges_ChildTablePermissions() As Boolean
  ' Apply the Child View Permutation permissions to SQL Server database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fSysSecManager As Boolean
  Dim lngPermittedChildView As Long
  Dim sSQL As String
  
  Dim sGroupName As String
  Dim sTableName As String
  Dim sChildViewName As String
  Dim sColumnName As String
'  Dim sSelectDeny As String
'  Dim sUpdateDeny As String
  Dim objGroup As SecurityGroup
  Dim objTable As SecurityTable
  Dim avChildTables() As Variant
  Dim iMaxRouteLength As Integer
  Dim rsInfo As New ADODB.Recordset
  Dim rsTemp As New ADODB.Recordset
  Dim iNextIndex As Integer
  Dim sSysSecRoles As String
  Dim sNonSysSecRoles As String
  Dim iLoop1 As Integer
  Dim iLoop2 As Integer
  Dim iParentCount As Integer
  Dim avParents() As Variant
  Dim rsParents As New ADODB.Recordset
  Dim rsChildren As New ADODB.Recordset
  Dim rsViews As New ADODB.Recordset
  Dim rsChildViews As New ADODB.Recordset
  Dim rsColumns As New ADODB.Recordset
  Dim sTempName As String
  Dim sRelatedChildTables As String
  Dim fTableOK As Boolean
  Dim iOKViewCount As Integer
  Dim fViewOK As Boolean
  Dim lngParentViewID As Long
  Dim sTemp As String
  Dim iParentJoinType As Integer
  Dim cmdChildView As ADODB.Command
  Dim lngViewID As Long
  Dim sViewName As String
  Dim lngLastParentID As Long
  Dim iLoop As Integer
  Dim sExistingChildViews As String
  Dim sParentIDs As String
  Dim cmdIsSysSecMgr As New ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim sSelectGrant As New SecurityMgr.clsStringBuilder
  Dim sUpdateGrant As New SecurityMgr.clsStringBuilder
  Dim sSQLCommands As New SecurityMgr.clsStringBuilder
  Dim strColumnName As String
  Dim fSomeUpdatable As Boolean

  Dim blnRemoveFix1082 As Boolean
  
  blnRemoveFix1082 = (GetSystemSetting("Remove Fix", "1082", "0") = "1")
  
  
  
  sRelatedChildTables = "0"

  fOK = True
  
  ' Create an array of all child tables, and each child's longest route to the top-level.
  ' eg.
  '              Table A
  '               / |
  '              /  |
  '             /   |
  '       Table B   |
  '             \   |
  '              \  |
  '               \ |
  '              Table C
  '
  ' Table A is a top-level table.
  ' Table B has a longest route to the top-level of 1.
  ' Table C has a longest route to the top-level of 2.
  '
  ' We need to create views for the tables nearest the top-level first, as they might then
  ' need to be propogated down. So, even though Tables B and C are both children of table A,
  ' we need to create the views on Table B first.
    
  ' Create an array of child tables.
  ' Column 1 = table ID.
  ' Column 2 = longest route to the top-level.
  ' Column 3 = table name.
  ReDim avChildTables(3, 100)
  iMaxRouteLength = 0
  iNextIndex = 0
  sSQL = "SELECT DISTINCT ASRSysRelations.childID, ASRSysTables.tableName" & _
    " FROM ASRSysRelations" & _
    " INNER JOIN ASRSysTables ON ASRSysRelations.childID = ASRSysTables.tableID"
  rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
  
  If Not rsInfo.EOF Then
    Do While Not rsInfo.EOF
      iNextIndex = iNextIndex + 1
      If iNextIndex > UBound(avChildTables, 2) Then ReDim Preserve avChildTables(3, iNextIndex + 100)
      avChildTables(1, iNextIndex) = rsInfo.Fields(0).Value
      avChildTables(2, iNextIndex) = LongestRouteToTopLevel(rsInfo.Fields(0).Value)
      avChildTables(3, iNextIndex) = rsInfo.Fields(1).Value
           
      If iMaxRouteLength < avChildTables(2, iNextIndex) Then
        iMaxRouteLength = CInt(avChildTables(2, iNextIndex))
      End If
      
      rsInfo.MoveNext
    Loop
    ReDim Preserve avChildTables(3, iNextIndex)
  End If
  rsInfo.Close
  
  ' Get the list of SysSecMgr groups, and non-SysSecMgr groups.
  sSysSecRoles = vbNullString
  sNonSysSecRoles = vbNullString
  
  For Each objGroup In gObjGroups
    sGroupName = "[" & objGroup.Name & "]"
  
    If objGroup.Initialised Then
      ' Determine if the current group is granted access to the System or Security Managers.
      fSysSecManager = objGroup.SystemPermissions.Item("P_MODULEACCESS_SYSTEMMANAGER").Allowed Or _
        objGroup.SystemPermissions.Item("P_MODULEACCESS_SECURITYMANAGER").Allowed
    Else
       With cmdIsSysSecMgr
         .CommandText = "dbo.sp_ASRIsSysSecMgr"
         .CommandType = adCmdStoredProc
         .CommandTimeout = 0
         Set .ActiveConnection = gADOCon
     
         Set pmADO = .CreateParameter("GroupName", adVarChar, adParamInput, 256)
         .Parameters.Append pmADO
         pmADO.Value = objGroup.Name
     
         Set pmADO = .CreateParameter("Result", adBoolean, adParamOutput)
         .Parameters.Append pmADO
    
         .Execute
     
         fSysSecManager = .Parameters(1).Value
       End With
       Set cmdIsSysSecMgr = Nothing
    
    End If
    
    If fSysSecManager Then
      sSysSecRoles = sSysSecRoles & IIf(LenB(sSysSecRoles) <> 0, ",", vbNullString) & sGroupName
    Else
      sNonSysSecRoles = sNonSysSecRoles & IIf(LenB(sNonSysSecRoles) <> 0, ",", vbNullString) & sGroupName
    End If
  Next objGroup

  ' Deny non-SysMgr and non-SecMgr users access to orphaned child tables.
  sSQLCommands.TheString = vbNullString
  sSQL = "SELECT ASRSysTables.tableName" & _
    " FROM ASRSysTables" & _
    " WHERE ASRSysTables.tableType = " & Trim(Str(tabChild)) & _
    " AND ASRSysTables.tableID NOT IN (" & sRelatedChildTables & ")"
  rsChildren.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
  With rsChildren
    Do While (Not .EOF)
      If LenB(sSysSecRoles) <> 0 Then
        sSQLCommands.Append "GRANT DELETE, INSERT, SELECT, UPDATE ON " & .Fields(0).Value & " TO " & sSysSecRoles & vbNewLine
      End If
  
      If LenB(sNonSysSecRoles) <> 0 Then
        sSQLCommands.Append "REVOKE DELETE, INSERT, SELECT, UPDATE ON " & .Fields(0).Value & " TO " & sNonSysSecRoles & vbNewLine
      End If
  
      .MoveNext
    Loop
  
    .Close
  End With

  ' Runs the statements
  If sSQLCommands.Length <> 0 Then
    gADOCon.Execute sSQLCommands.ToString, , adExecuteNoRecords
  End If

  ' For each child table (do those nearest to the top-level first).
  sSQLCommands.TheString = vbNullString
  For iLoop1 = 1 To iMaxRouteLength
    ' For each table this distance from the top-level.
    For iLoop2 = 1 To UBound(avChildTables, 2)
      If CInt(avChildTables(2, iLoop2)) = iLoop1 Then
        ' JPD20020819 Fault 4308
        If LenB(sSysSecRoles) <> 0 Then
          sSQL = "GRANT DELETE, INSERT, SELECT, UPDATE ON " & avChildTables(3, iLoop2) & " TO " & sSysSecRoles
          gADOCon.Execute sSQL, , adExecuteNoRecords
        End If
        
        If LenB(sNonSysSecRoles) <> 0 Then
          sSQL = "REVOKE DELETE, INSERT, SELECT, UPDATE ON " & avChildTables(3, iLoop2) & " TO " & sNonSysSecRoles
          gADOCon.Execute sSQL, , adExecuteNoRecords
        End If
  
        For Each objGroup In gObjGroups
          If objGroup.Initialised Then
            sGroupName = "[" & objGroup.Name & "]"
            sTableName = avChildTables(3, iLoop2)

            If objGroup.Tables(sTableName).Changed Then
              ' Drop this table's existing child views for the current user group.
              sSQL = "SELECT childViewID FROM ASRSysChildViews2 WHERE role = '" & objGroup.Name & "' AND tableID = " & Trim(Str(avChildTables(1, iLoop2)))
              rsTemp.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
              Do While Not rsTemp.EOF
                sSQL = "SELECT name FROM sysobjects WHERE name LIKE 'ASRSysCV" & Trim(Str(rsTemp!childViewID)) & "#%#" & Replace(objGroup.Name, " ", "_") & "' AND xtype = 'V'"
                rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
                With rsInfo
                  Do While (Not .EOF)
                    sSQL = "DROP VIEW " & !Name
                    gADOCon.Execute sSQL, , adExecuteNoRecords
                
                    .MoveNext
                  Loop
                
                  .Close
                End With
    
                rsTemp.MoveNext
              Loop
            
              rsTemp.Close
                  
              If objGroup.Tables(sTableName).SelectPrivilege = giPRIVILEGES_NONEGRANTED Then
                
                'MH20050106 Fault 9470, 9471, 9472, 9373
                With objGroup.Tables(sTableName)
                  If .SelectPrivilegeChanged Then
                    AuditPermission sGroupName, sTableName, "Deny", "Read"
                  End If
                  If .UpdatePrivilegeChanged Then
                    AuditPermission sGroupName, sTableName, "Deny", "Edit"
                  End If
                  If .DeletePrivilege <> .DeleteOriginalPrivilege Then
                    AuditPermission sGroupName, sTableName, "Deny", "Delete"
                  End If
                  If .InsertPrivilege <> .InsertOriginalPrivilege Then
                    AuditPermission sGroupName, sTableName, "Deny", "New"
                  End If
                End With
              Else
                ' Check which parents of this child table, the current role can see.
                iParentCount = 0
                ' Create an array of the parents of the given table that are accessible by the given group.
                ' Column 1 = parent type (UT = top-level table
                '                         UV = view of a top-level table
                '                         SV = system view)
                ' Column 2 = parent ID
                ' Column 3 = parent table ID
                ' Column 4 = parent name
                ReDim avParents(4, 0)
    
                ' Get the parents of the current child table.
                sParentIDs = ""
                
                sSQL = "SELECT ASRSysTables.tableID, ASRSysTables.tableName, ASRSysTables.tableType" & _
                  " FROM ASRSysRelations" & _
                  " INNER JOIN ASRSysTables ON ASRSysRelations.parentID = ASRSysTables.tableID" & _
                  " WHERE ASRSysRelations.childID = " & Trim(Str(avChildTables(1, iLoop2)))
    
                rsParents.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
                Do While (Not rsParents.EOF)
                  sParentIDs = sParentIDs & _
                    IIf(Len(sParentIDs) > 0, ",", "") & _
                    "ID_" & CStr(rsParents!TableID)
                  
                  If rsParents!TableType = tabParent Then
                    ' Parent is a top-level table.
                    fTableOK = (objGroup.Tables(rsParents!TableName).SelectPrivilege <> giPRIVILEGES_NONEGRANTED)
    
                    If fTableOK Then
                      iParentCount = iParentCount + 1
    
                      ' The current group has permission to see all records in the parent table.
                      iNextIndex = UBound(avParents, 2) + 1
                      ReDim Preserve avParents(4, iNextIndex)
                      avParents(1, iNextIndex) = "UT"
                      avParents(2, iNextIndex) = rsParents!TableID
                      avParents(3, iNextIndex) = rsParents!TableID
                      avParents(4, iNextIndex) = rsParents!TableName
                    Else
                      ' The current group does NOT have permission to see all records in the parent table.
                      ' Get the permitted views on the table.
                      sSQL = "SELECT ASRSysViews.viewID, ASRSysViews.viewName" & _
                        " FROM ASRSysViews" & _
                        " WHERE ASRSysViews.viewTableID = " & Trim(Str(rsParents!TableID))
                      iOKViewCount = 0
    
                      rsViews.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
                      Do While Not rsViews.EOF
                        fViewOK = (objGroup.Views(rsViews!ViewName).SelectPrivilege <> giPRIVILEGES_NONEGRANTED)
    
                        If fViewOK Then
                          iOKViewCount = iOKViewCount + 1
    
                          iNextIndex = UBound(avParents, 2) + 1
                          ReDim Preserve avParents(4, iNextIndex)
                          avParents(1, iNextIndex) = "UV"
                          avParents(2, iNextIndex) = rsViews!ViewID
                          avParents(3, iNextIndex) = rsParents!TableID
                          avParents(4, iNextIndex) = rsViews!ViewName
                        End If
    
                        rsViews.MoveNext
                      Loop
    
                      rsViews.Close
    
                      If iOKViewCount > 0 Then
                        iParentCount = iParentCount + 1
                      End If
                    End If
                  ElseIf rsParents!TableType = tabChild Then
                    ' Parent is not a top-level table.
                    lngParentViewID = 0
                    sSQL = "SELECT childViewID FROM ASRSysChildViews2 WHERE role = '" & objGroup.Name & "' AND tableID = " & Trim(Str(rsParents!TableID))

                    rsChildViews.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
                    Do While (Not rsChildViews.EOF) And (lngParentViewID = 0)
                      lngParentViewID = rsChildViews!childViewID
    
                      ' Check if it really exists.
                      sTemp = Left("ASRSysCV" & Trim(Str(lngParentViewID)) & "#" & Replace(rsParents!TableName, " ", "_") & "#" & Replace(objGroup.Name, " ", "_"), 255)
                      sSQL = "SELECT COUNT(*) AS result FROM sysobjects WHERE name = '" & sTemp & "' AND xtype = 'V'"
                      rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
                      If rsInfo!Result = 0 Then
                        lngParentViewID = 0
                      End If
      
                      rsInfo.Close
                      rsChildViews.MoveNext
                    Loop
    
                    rsChildViews.Close
    
                    If lngParentViewID > 0 Then
                      iParentCount = iParentCount + 1
    
                      iNextIndex = UBound(avParents, 2) + 1
                      ReDim Preserve avParents(4, iNextIndex)
                      avParents(1, iNextIndex) = "SV"
                      avParents(2, iNextIndex) = lngParentViewID
                      avParents(3, iNextIndex) = rsParents!TableID
                      avParents(4, iNextIndex) = Left("ASRSysCV" & Trim(Str(lngParentViewID)) & "#" & Replace(rsParents!TableName, " ", "_") & "#" & Replace(objGroup.Name, " ", "_"), 255)
                    End If
                  End If
    
                  rsParents.MoveNext
                Loop
    
                rsParents.Close
    
                If iParentCount = 0 Then
                  'MH20050106 Fault 9470, 9471, 9472, 9373
                  With objGroup.Tables(sTableName)
                    If .SelectPrivilegeChanged Then
                      AuditPermission sGroupName, sTableName, "Deny", "Read"
                    End If
                    If .UpdatePrivilegeChanged Then
                      AuditPermission sGroupName, sTableName, "Deny", "Edit"
                    End If
                    If .DeletePrivilege <> .DeleteOriginalPrivilege Then
                      AuditPermission sGroupName, sTableName, "Deny", "Delete"
                    End If
                    If .InsertPrivilege <> .InsertOriginalPrivilege Then
                      AuditPermission sGroupName, sTableName, "Deny", "New"
                    End If
                  End With
                Else
                  ' More than 1 parent. Do we want the OR join child view, or the AND join child view ?
                  iParentJoinType = objGroup.Tables(avChildTables(3, iLoop2)).ParentJoinType
    
                  ' Enter the view definition in the ASRSysChildView2 table.
                  Set cmdChildView = New ADODB.Command
                  With cmdChildView
                    .CommandText = "dbo.sp_ASRInsertChildView2"
                    .CommandType = adCmdStoredProc
                    .CommandTimeout = 0
                    Set .ActiveConnection = gADOCon
                
                    Set pmADO = .CreateParameter("Result", adInteger, adParamOutput)
                    .Parameters.Append pmADO
                
                    Set pmADO = .CreateParameter("TableID", adInteger, adParamInput)
                    .Parameters.Append pmADO
                    pmADO.Value = avChildTables(1, iLoop2)
                
                    Set pmADO = .CreateParameter("JoinType", adInteger, adParamInput)
                    .Parameters.Append pmADO
                    pmADO.Value = iParentJoinType
                    
                    Set pmADO = .CreateParameter("Name", adVarChar, adParamInput, 256)
                    .Parameters.Append pmADO
                    pmADO.Value = objGroup.Name
                
                    .Execute
                
                    lngViewID = IIf(IsNull(.Parameters(0).Value), vbNullString, .Parameters(0).Value)
                  End With
                  Set cmdChildView = Nothing
                  
      
                  ' Delete the existing entries in the ASRSysChildViewParents2 table.
                  sSQL = "DELETE FROM ASRSysChildViewParents2 WHERE childViewID = " & Trim(Str(lngViewID))
                  gADOCon.Execute sSQL, , adExecuteNoRecords
    
                  For iNextIndex = 1 To UBound(avParents, 2)
                    sSQL = "INSERT INTO ASRSysChildViewParents2" & _
                      " (childViewID, parentType, parentID, parentTableID)" & _
                      " VALUES (" & Trim(Str(lngViewID)) & ", " & _
                      "'" & avParents(1, iNextIndex) & "', " & _
                      Trim(Str(avParents(2, iNextIndex))) & ", " & _
                      Trim(Str(avParents(3, iNextIndex))) & ")"
                    gADOCon.Execute sSQL, , adExecuteNoRecords
                  Next iNextIndex
      
                  ' Create the view name.
                  sViewName = Left("ASRSysCV" & Trim(Str(lngViewID)) & "#" & Replace(sTableName, " ", "_") & "#" & Replace(objGroup.Name, " ", "_"), 255)


                  'MH20110119 HRPRO-1082
                  If blnRemoveFix1082 Then
                    ' Create the view
                    sSQL = "CREATE VIEW dbo." & sViewName & vbNewLine & _
                      "AS" & vbNewLine & _
                      "        SELECT " & sTableName & ".*" & vbNewLine & _
                      "        FROM " & sTableName & vbNewLine & _
                      "        WHERE " & vbNewLine & _
                      "                (" & vbNewLine & _
                      "                        " & _
                      sTableName & ".ID_" & Trim(Str(avParents(3, 1))) & " IN (SELECT id FROM " & avParents(4, 1) & ")" & vbNewLine
  
                    lngLastParentID = avParents(3, 1)
  
                    For iLoop = 2 To UBound(avParents, 2)
                      If lngLastParentID <> avParents(3, iLoop) Then
                        lngLastParentID = avParents(3, iLoop)
  
                        sSQL = sSQL & _
                          "                )" & vbNewLine & _
                          "                " & IIf(iParentJoinType = 1, "AND", "OR") & vbNewLine & _
                          "                (" & vbNewLine
                      Else
                        sSQL = sSQL & _
                          "                        OR" & vbNewLine
                      End If
  
                      sSQL = sSQL & "                        " & _
                        sTableName & ".ID_" & Trim(Str(avParents(3, iLoop))) & " IN (SELECT id FROM " & avParents(4, iLoop) & ")" & vbNewLine
                    Next iLoop
  
                    sSQL = sSQL & "                )"
  
                    gADOCon.Execute sSQL, , adExecuteNoRecords

                  
                  Else
              
                    ' Create the view
                    sSQL = "CREATE VIEW dbo." & sViewName & vbNewLine & _
                      "AS" & vbNewLine & _
                      "        SELECT " & sTableName & ".*" & vbNewLine & _
                      "        FROM " & sTableName & vbNewLine & _
                      "        WHERE (" & vbNewLine & _
                      "                        " & _
                      sTableName & ".ID_" & Trim$(Str$(avParents(3, 1))) & " IN (SELECT id FROM " & avParents(4, 1) & vbNewLine
                    
                    lngLastParentID = avParents(3, 1)
                                 
                    For iLoop = 2 To UBound(avParents, 2)
                      
                      If lngLastParentID <> avParents(3, iLoop) Then
                        lngLastParentID = avParents(3, iLoop)
        
                        sSQL = sSQL & _
                          "                ))" & vbNewLine & _
                          "                " & IIf(iParentJoinType = 1, "AND", "OR") & vbNewLine & _
                          "                (" & vbNewLine & _
                          "                        " & _
                          sTableName & ".ID_" & Trim$(Str$(avParents(3, iLoop))) & " IN ("
                      Else
                        sSQL = sSQL & _
                          "                            UNION" & vbNewLine
                      End If
        
                      sSQL = sSQL & "                            SELECT id FROM " & avParents(4, iLoop) & vbNewLine
                    Next iLoop
        
                    sSQL = sSQL & "                ))"
                      
                    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
    
                  End If
                  
                  
                  
                  
                  If LenB(sSysSecRoles) <> 0 Then
                    sSQL = "GRANT DELETE, INSERT, SELECT, UPDATE ON " & sViewName & " TO " & sSysSecRoles
                    gADOCon.Execute sSQL, , adExecuteNoRecords
                  End If
    
                  If objGroup.SystemPermissions.Item("P_MODULEACCESS_SYSTEMMANAGER").Allowed Or _
                    objGroup.SystemPermissions.Item("P_MODULEACCESS_SECURITYMANAGER").Allowed Then
                    
                    'MH20050106 Fault 9470, 9471, 9472, 9373
                    With objGroup.Tables(sTableName)
                      If .SelectPrivilegeChanged Then
                        AuditPermission sGroupName, sTableName, "Grant", "Read"
                      End If
                      If .UpdatePrivilegeChanged Then
                        AuditPermission sGroupName, sTableName, "Grant", "Edit"
                      End If
                      If .DeletePrivilege <> .DeleteOriginalPrivilege Then
                        AuditPermission sGroupName, sTableName, "Grant", "Delete"
                      End If
                      If .InsertPrivilege <> .InsertOriginalPrivilege Then
                        AuditPermission sGroupName, sTableName, "Grant", "New"
                      End If
                    End With
                  Else
                    ' Apply the configured permissions to the child view permutation.
                    ' Initialise the Table Column permissions command strings.
                    
                    ' JDM - Needs these permissions for the history list to be generated properly (amongst other things!)
                    sUpdateGrant.TheString = "ID,TimeStamp," & sParentIDs
                    sSelectGrant.TheString = "ID,TimeStamp," & sParentIDs
      
                    fSomeUpdatable = False
                    
                    ' Get the set of non-system columns in the table.
                    sSQL = "SELECT columnName, columnID" & _
                      " FROM ASRSysColumns" & _
                      " WHERE tableID = " & Trim(Str(avChildTables(1, iLoop2))) & _
                      " AND columnType <> " & Trim(Str(colSystem)) & _
                      " AND columnType <> " & Trim(Str(colLink))
                    rsColumns.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
                    Do While Not rsColumns.EOF

'MH20050106 Fault 9470, 9471, 9472, 9373
'                      ' Existing column in an existing table.
'                      ' Get the column's permission from the collection read earlier.
'                      If Not objGroup.Tables.Item(avChildTables(3, iLoop2)).Columns.Item(UCase(Trim(rsColumns!ColumnName))).SelectPrivilege Then
'                        sSelectDeny = sSelectDeny & IIf(Len(sSelectDeny) > 0, ",", vbnullstring) & rsColumns!ColumnName
'                        AuditPermission sGroupName, sTableName, "Deny", "Read", rsColumns!ColumnName
'                      End If
'                      If Not objGroup.Tables.Item(avChildTables(3, iLoop2)).Columns.Item(UCase(Trim(rsColumns!ColumnName))).UpdatePrivilege Then
'                        sUpdateDeny = sUpdateDeny & IIf(Len(sUpdateDeny) > 0, ",", vbnullstring) & rsColumns!ColumnName
'                        AuditPermission sGroupName, sTableName, "Deny", "Edit", rsColumns!ColumnName
'                      End If

                      strColumnName = Trim(rsColumns!ColumnName)
                      
                      With objGroup.Tables.Item(avChildTables(3, iLoop2)).Columns.Item(UCase$(strColumnName))
                        If .SelectPrivilege Then
                          sSelectGrant.Append IIf(sSelectGrant.Length <> 0, ",", vbNullString) & strColumnName
                        End If
                          
                        If .UpdatePrivilege Then
                          sUpdateGrant.Append IIf(sUpdateGrant.Length <> 0, ",", vbNullString) & strColumnName
                          fSomeUpdatable = True
                        End If

                        If .SelectPrivilege <> .SelectOriginalPrivilege Then
                          AuditPermission sGroupName, sTableName, IIf(.SelectPrivilege, "Grant", "Deny"), "Read", strColumnName
                        End If
                        If .UpdatePrivilege <> .UpdateOriginalPrivilege Then
                          AuditPermission sGroupName, sTableName, IIf(.UpdatePrivilege, "Grant", "Deny"), "Edit", strColumnName
                        End If
                      End With

                      rsColumns.MoveNext
                    Loop
  
                    rsColumns.Close
  
                    If fSomeUpdatable Then
                      sSQL = "SELECT columnName, columnID" & _
                        " FROM ASRSysColumns" & _
                        " WHERE tableID = " & Trim(Str(avChildTables(1, iLoop2))) & _
                        " AND columnType = " & Trim(Str(colLink))
                      rsColumns.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
                      Do While Not rsColumns.EOF
                        strColumnName = Trim(rsColumns!ColumnName)
                        sUpdateGrant.Append IIf(sUpdateGrant.Length <> 0, ",", vbNullString) & strColumnName

                        rsColumns.MoveNext
                      Loop
    
                      rsColumns.Close
                    End If
    
                    With objGroup.Tables.Item(avChildTables(3, iLoop2))
                      
                      'MH20050106 Fault 9470, 9471, 9472, 9373
                      If .DeletePrivilege <> .DeleteOriginalPrivilege Then
                        AuditPermission sGroupName, sTableName, IIf(.DeletePrivilege, "Grant", "Deny"), "Delete"
                      End If
                      
                      'MH20050106 Fault 9470, 9471, 9472, 9373
                      If .InsertPrivilege <> .InsertOriginalPrivilege Then
                        AuditPermission sGroupName, sTableName, IIf(.InsertPrivilege, "Grant", "Deny"), "New"
                      End If
                      
                      'MH20050106 Fault 9470, 9471, 9472, 9373
                      If .SelectPrivilegeChanged Then
                        AuditPermission sGroupName, sTableName, IIf(.SelectPrivilege, "Grant", "Deny"), "Read"
                      End If
                      
                      'MH20050106 Fault 9470, 9471, 9472, 9373
                      If .UpdatePrivilegeChanged Then
                        AuditPermission sGroupName, sTableName, IIf(.UpdatePrivilege, "Grant", "Deny"), "Edit"
                      End If

                      ' INSERT permission.
                      If .InsertPrivilege Then
                        sSQL = "GRANT INSERT ON " & sViewName & " TO " & sGroupName
                        gADOCon.Execute sSQL, , adExecuteNoRecords
                      End If

                      ' SELECT permissions.
                      If .SelectPrivilege = giPRIVILEGES_ALLGRANTED Then
                        sSQL = "GRANT SELECT ON " & sViewName & " TO " & sGroupName
                        gADOCon.Execute sSQL, , adExecuteNoRecords
                      ElseIf .SelectPrivilege = giPRIVILEGES_SOMEGRANTED Then
                        If sSelectGrant.Length <> 0 Then
                          'MH20060620 Fault 11186
                          'sSQL = "GRANT SELECT(" & sSelectGrant.ToString & ") ON " & sViewName & " TO " & sGroupName
                          sSQL = "GRANT SELECT(ID,Timestamp," & sSelectGrant.ToString & ") ON " & sViewName & " TO " & sGroupName
                          gADOCon.Execute sSQL, , adExecuteNoRecords
                        End If
                      End If
                      
                      ' UPDATE permissions
                      If .UpdatePrivilege = giPRIVILEGES_ALLGRANTED Then
                        sSQL = "GRANT UPDATE ON " & sViewName & " TO " & sGroupName
                        gADOCon.Execute sSQL, , adExecuteNoRecords
                      ElseIf .UpdatePrivilege = giPRIVILEGES_SOMEGRANTED Then
                        If sUpdateGrant.Length <> 0 Then
                          'MH20060620 Fault 11186
                          'sSQL = "GRANT UPDATE(" & sUpdateGrant.ToString & ") ON " & sViewName & " TO " & sGroupName
                          sSQL = "GRANT UPDATE(ID,Timestamp," & sUpdateGrant.ToString & ") ON " & sViewName & " TO " & sGroupName
                          gADOCon.Execute sSQL, , adExecuteNoRecords
                        End If
                      End If
                      
                      ' DELETE permission.
                      If .DeletePrivilege Then
                        sSQL = "GRANT DELETE ON " & sViewName & " TO " & sGroupName
                        gADOCon.Execute sSQL, , adExecuteNoRecords
                      End If
                      
                    End With
                  End If
                End If
              End If
            
              objGroup.Tables(sTableName).Changed = False
            End If
          End If
        Next objGroup
        Set objGroup = Nothing
      End If
    Next iLoop2
  Next iLoop1

  ' Remove any redundant records from the ASRSysChildViews2 and ASRSysChildViewParents2 tables.
  sExistingChildViews = "0"
  sSQL = "SELECT name FROM sysobjects WHERE name LIKE 'ASRSysCV%' AND xtype = 'V'"
  rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
  With rsInfo
    Do While (Not .EOF)
      sTempName = .Fields(0).Value
      sExistingChildViews = sExistingChildViews & "," & Mid(sTempName, 9, InStr(sTempName, "#") - 9)
      
      .MoveNext
    Loop
  
    .Close
  End With
  
  ' Delete invalid records from ASRSysChildViews2 and ASRSysChildViewParents2
  sSQL = "DELETE FROM ASRSysChildViews2 WHERE childViewID NOT IN (" & sExistingChildViews & ")"
  gADOCon.Execute sSQL, , adExecuteNoRecords
  
  sSQL = "DELETE FROM ASRSysChildViewParents2 WHERE childViewID NOT IN (" & sExistingChildViews & ")"
  gADOCon.Execute sSQL, , adExecuteNoRecords
  
TidyUpAndExit:
  
  Set rsParents = Nothing
  Set rsChildren = Nothing
  Set rsViews = Nothing
  Set rsChildViews = Nothing
  Set rsColumns = Nothing
  Set rsInfo = Nothing
  Set rsTemp = Nothing
    
  Set objGroup = Nothing
  Set objTable = Nothing
  ApplyChanges_ChildTablePermissions = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function LongestRouteToTopLevel(plngTableID As Long) As Integer
  ' Return the given table's longest route to the top-level.
  ' This is used when creating child views.
  Dim iLongestRoute As Integer
  Dim iParentsLongestRoute As Integer
  Dim sSQL As String
  Dim rsParents As New ADODB.Recordset
  
  iLongestRoute = 0
  
  sSQL = "SELECT parentID" & _
    " FROM ASRSysRelations" & _
    " WHERE childID = " & Trim(Str(plngTableID))
  rsParents.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
  With rsParents
    Do While (Not .EOF)
      iParentsLongestRoute = LongestRouteToTopLevel(CLng(.Fields(0).Value))
      
      If (iParentsLongestRoute + 1) > iLongestRoute Then
        iLongestRoute = (iParentsLongestRoute + 1)
      End If
      
      .MoveNext
    Loop

    .Close
  End With
  Set rsParents = Nothing
  
  LongestRouteToTopLevel = iLongestRoute
End Function

Private Function ApplyChanges_CreateViewValidationStoredProcedures() As Boolean
  ' Create the view validation stored procedures.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim lngTableID As Long
  Dim sSPName As String
  Dim sSPCode As New SecurityMgr.clsStringBuilder
  Dim sSQL As String
  Dim objGroup As SecurityGroup
  Dim objTable As SecurityTable
  Dim rsTableInfo As New ADODB.Recordset
  Dim rsViewInfo As New ADODB.Recordset
    
  Const sVIEWVALIDATIONSPPREFIX = "sp_ASRValidateView_"
  
  fOK = True
  
  ' Create the stored proedures for each group thats changed.
  For Each objGroup In gObjGroups
    With objGroup
      If .Initialised Then
        For Each objTable In .Tables
          ' Only do top-level tables.
          If objTable.TableType = tabParent Then
            ' Get the table's id.
            sSQL = "SELECT tableID" & _
              " FROM ASRSysTables" & _
              " WHERE tableName = '" & objTable.Name & "'"
            rsTableInfo.Open sSQL, gADOCon, adOpenForwardOnly, adCmdText
            fOK = Not (rsTableInfo.EOF And rsTableInfo.BOF)
            
            If fOK Then
              lngTableID = rsTableInfo!TableID
            End If
            
            rsTableInfo.Close
            
            If fOK Then
              sSPName = "[" & sVIEWVALIDATIONSPPREFIX & Trim$(Str(lngTableID)) & "_" & .Name & "]"
                        
              ' Drop any existing stored procedure.
              sSQL = "IF EXISTS" & _
                " (SELECT Name" & _
                "   FROM sysobjects" & _
                "   WHERE id = object_id('" & sSPName & "')" & _
                "     AND sysstat & 0xf = 4)" & _
                " DROP PROCEDURE " & sSPName
              gADOCon.Execute sSQL, , adExecuteNoRecords
      
              '
              ' Create the stored procedure creation string if the table is a top level table.
              '
              sSPCode.TheString = "/* ---------------------------------------------------------- */" & vbNewLine & _
                "/* View validation stored procedure.                   */" & vbNewLine & _
                "/* Automatically generated by the System/Security Managers.   */" & vbNewLine & _
                "/* ---------------------------------------------------------- */" & vbNewLine & _
                "/* '" & objTable.Name & "' table */" & vbNewLine & _
                "/* ------------------------------------------------ */" & vbNewLine & _
                "CREATE PROCEDURE dbo." & sSPName & vbNewLine & _
                "(" & vbNewLine & _
                "    @pfResult bit OUTPUT," & vbNewLine & _
                "    @piRecordID integer" & vbNewLine & _
                ")" & vbNewLine & _
                "AS" & vbNewLine & _
                "BEGIN" & vbNewLine & _
                "    DECLARE @iRecCount integer" & vbNewLine
            
              ' If the current group has permission on the whole table return 1.
              If objTable.SelectPrivilege Then
                sSPCode.Append vbNewLine & _
                  "    SET @pfResult = 1" & vbNewLine
              Else
                sSPCode.Append vbNewLine & _
                  "    SET @pfResult = 0" & vbNewLine
          
                ' Loop through the views adding code for each permissable one.
                sSQL = "SELECT viewName" & _
                  " FROM ASRSysViews" & _
                  " WHERE viewTableID = " & Trim$(Str(lngTableID))
                rsViewInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText

                Do While Not rsViewInfo.EOF
                  If .Views.Item(rsViewInfo!ViewName).SelectPrivilege Then

                    sSPCode.Append vbNewLine & _
                      "    IF @pfResult = 0" & vbNewLine & _
                      "    BEGIN" & vbNewLine & _
                      "        /* Check if the current user can see the record in the '" & rsViewInfo!ViewName & "' view. */" & vbNewLine & _
                      "        SELECT @iRecCount = COUNT(id)" & vbNewLine & _
                      "        FROM " & rsViewInfo!ViewName & vbNewLine & _
                      "        WHERE id = @piRecordID" & vbNewLine & vbNewLine & _
                      "        IF @iRecCount > 0 SET @pfResult = 1" & vbNewLine & _
                      "    END" & vbNewLine
                  End If

                  rsViewInfo.MoveNext
                Loop
                
                rsViewInfo.Close
              End If
              
              sSPCode.Append vbNewLine & _
                "END"
        
              ' Create the stored procedure.
              gADOCon.Execute sSPCode.ToString, , adExecuteNoRecords
              
            End If
          End If
        
          If Not fOK Then
            Exit For
          End If
        Next objTable
        Set objTable = Nothing
      End If
    End With
    
    If Not fOK Then
      Exit For
    End If
  Next objGroup
  
TidyUpAndExit:
  Set rsTableInfo = Nothing
  Set rsViewInfo = Nothing
  Set objGroup = Nothing
  Set objTable = Nothing
  ApplyChanges_CreateViewValidationStoredProcedures = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function ApplyChanges_NonChildTablePermissions() As Boolean
  ' Apply the Table and Table Column permissions to SQL Server database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fSysSecManager As Boolean
  Dim sSQL As String
  Dim sGroupName As String
  Dim sTableName As String
  Dim sColumnName As String
'  Dim sSelectDeny As String
'  Dim sUpdateDeny As String
  Dim objGroup As SecurityGroup
  Dim objTable As SecurityTable
  Dim objColumn As SecurityColumn
  Dim sSelectGrant As New SecurityMgr.clsStringBuilder
  Dim sUpdateGrant As New SecurityMgr.clsStringBuilder
    
  fOK = True
  
  ' Apply the changes for each group..
  For Each objGroup In gObjGroups
    With objGroup
      sGroupName = "[" & .Name & "]"
      
      If .Initialised Then

        fSysSecManager = .SystemPermissions.Item("P_MODULEACCESS_SYSTEMMANAGER").Allowed Or _
          .SystemPermissions.Item("P_MODULEACCESS_SECURITYMANAGER").Allowed
        
        For Each objTable In .Tables
          ' Only do top-level, or lookup tables here. Child tables have permission applied in a different function.
          If (objTable.TableType <> tabChild) And (objTable.Changed) Then
            sTableName = objTable.Name
            
            If fSysSecManager Then
              sSQL = "GRANT DELETE, INSERT, SELECT, UPDATE ON [" & sTableName & "] TO " & sGroupName
              gADOCon.Execute sSQL, , adExecuteNoRecords
            Else
              ' Initialise the Table Column permissions command strings.
              sSelectGrant.TheString = vbNullString
              sUpdateGrant.TheString = vbNullString
            
              If ((objTable.SelectPrivilege = giPRIVILEGES_SOMEGRANTED) And (Not objTable.TableType = tabLookup)) Or _
                (objTable.UpdatePrivilege = giPRIVILEGES_SOMEGRANTED) Then
              
                For Each objColumn In objTable.Columns
                  sColumnName = objColumn.Name
                  
                  ' Create the SELECT and UPDATE grant/revoke/deny commands for each column.
                  ' Force SELECT permission on lookup tables.
                  ' Create the SELECT grant/revoke/deny command.
                  If objColumn.SelectPrivilege Then
                    sSelectGrant.Append IIf(sSelectGrant.Length <> 0, ",", vbNullString) & sColumnName
                    'MH20050106 Fault 9470, 9471, 9472, 9373
                    If objColumn.SelectPrivilege <> objColumn.SelectOriginalPrivilege Then
                      AuditPermission sGroupName, sTableName, IIf(objColumn.SelectPrivilege, "Grant", "Deny"), "Read", sColumnName
                    End If
                  End If
                
                  ' Create the UPDATE grant/revoke/deny command.
                  If objColumn.UpdatePrivilege Then
                    sUpdateGrant.Append IIf(sUpdateGrant.Length <> 0, ",", vbNullString) & sColumnName
                    'MH20050106 Fault 9470, 9471, 9472, 9373
                    If objColumn.UpdatePrivilege <> objColumn.UpdateOriginalPrivilege Then
                      AuditPermission sGroupName, sTableName, IIf(objColumn.UpdatePrivilege, "Grant", "Deny"), "Edit", sColumnName
                    End If
                  End If
                  
                  ' Reset the 'changed' flag on this column.
                  objColumn.Changed = False
                Next objColumn
                Set objColumn = Nothing
              End If
              
              ' Grant/revoke/deny the Table's DELETE permission.
              If objTable.DeletePrivilege Then
                sSQL = "GRANT DELETE ON " & sTableName & " TO " & sGroupName
                gADOCon.Execute sSQL, , adExecuteNoRecords
              Else
                sSQL = "REVOKE DELETE ON " & sTableName & " TO " & sGroupName
                gADOCon.Execute sSQL, , adExecuteNoRecords
              End If
              
              'MH20050106 Fault 9470, 9471, 9472, 9373
              If objTable.DeletePrivilege <> objTable.DeleteOriginalPrivilege Then
                AuditPermission sGroupName, sTableName, IIf(objTable.DeletePrivilege, "Grant", "Deny"), "Delete"
              End If
              
              ' Grant/revoke/deny the Table's INSERT permission.
              If objTable.InsertPrivilege Then
                sSQL = "GRANT INSERT ON " & sTableName & " TO " & sGroupName
                gADOCon.Execute sSQL, , adExecuteNoRecords
              Else
                sSQL = "REVOKE INSERT ON " & sTableName & " TO " & sGroupName
                gADOCon.Execute sSQL, , adExecuteNoRecords
              End If
              
              'MH20050106 Fault 9470, 9471, 9472, 9373
              If objTable.InsertPrivilege <> objTable.InsertOriginalPrivilege Then
                AuditPermission sGroupName, sTableName, IIf(objTable.InsertPrivilege, "Grant", "Deny"), "New"
              End If
              
              ' Select permissions.
              If objTable.SelectPrivilege = giPRIVILEGES_ALLGRANTED Or (objTable.TableType = tabLookup) Then
                sSQL = "GRANT SELECT ON " & sTableName & " TO " & sGroupName
                gADOCon.Execute sSQL, , adExecuteNoRecords
              ElseIf objTable.SelectPrivilege = giPRIVILEGES_SOMEGRANTED Then
                If sSelectGrant.Length <> 0 Then
                  
                  sSQL = "REVOKE SELECT ON " & sTableName & " TO " & sGroupName
                  gADOCon.Execute sSQL, , adExecuteNoRecords
                  
                  'MH20060620 Fault 11186
                  'sSQL = "GRANT SELECT(ID," & sSelectGrant.ToString & ") ON " & sTableName & " TO " & sGroupName
                  sSQL = "GRANT SELECT(ID,Timestamp," & sSelectGrant.ToString & ") ON " & sTableName & " TO " & sGroupName
                  gADOCon.Execute sSQL, , adExecuteNoRecords
                End If
              Else
                sSQL = "REVOKE SELECT ON " & sTableName & " TO " & sGroupName
                gADOCon.Execute sSQL, , adExecuteNoRecords
                gADOCon.Execute "GRANT SELECT(ID, TimeStamp) ON " & sTableName & " TO " & sGroupName
              End If

              If objTable.SelectPrivilegeChanged Then
                AuditPermission sGroupName, sTableName, IIf(objTable.SelectPrivilege, "Grant", "Deny"), "Read"
              End If
              
              ' Update permissions
              If objTable.UpdatePrivilege = giPRIVILEGES_ALLGRANTED Then
                sSQL = "GRANT UPDATE ON " & sTableName & " TO " & sGroupName
                gADOCon.Execute sSQL, , adExecuteNoRecords
              ElseIf objTable.UpdatePrivilege = giPRIVILEGES_SOMEGRANTED Then
                If sUpdateGrant.Length <> 0 Then
                  
                  sSQL = "REVOKE UPDATE ON " & sTableName & " TO " & sGroupName
                  gADOCon.Execute sSQL, , adExecuteNoRecords
                  
                  'MH20060620 Fault 11186
                  'sSQL = "GRANT UPDATE(" & sUpdateGrant.ToString & ") ON " & sTableName & " TO " & sGroupName
                  sSQL = "GRANT UPDATE(ID, Timestamp," & sUpdateGrant.ToString & ") ON " & sTableName & " TO " & sGroupName
                  gADOCon.Execute sSQL, , adExecuteNoRecords
                End If
              ElseIf objTable.UpdatePrivilege = giPRIVILEGES_NONEGRANTED Then
                sSQL = "REVOKE UPDATE ON " & sTableName & " TO " & sGroupName
                gADOCon.Execute sSQL, , adExecuteNoRecords
              End If
                            
              If objTable.UpdatePrivilegeChanged Then
                AuditPermission sGroupName, sTableName, IIf(objTable.UpdatePrivilege, "Grant", "Deny"), "Edit"
              End If
                            
              ' Reset the 'changed' flag on this table.
              objTable.Changed = False
            
            End If
          End If
        Next objTable
        Set objTable = Nothing
      End If
    End With
  Next objGroup
  
TidyUpAndExit:
  Set objGroup = Nothing
  Set objTable = Nothing
  Set objColumn = Nothing
  ApplyChanges_NonChildTablePermissions = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function



Private Function ApplyChanges_UserViewPermissions() As Boolean
  ' Apply the View and View Column permissions to SQL Server database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fSysSecManager As Boolean
  Dim sSQL As String
  Dim sGroupName As String
  Dim sViewName As String
  Dim sColumnName As String
  Dim objGroup As SecurityGroup
  Dim objView As SecurityTable
  Dim objColumn As SecurityColumn
  Dim sSelectGrant As New SecurityMgr.clsStringBuilder
  Dim sUpdateGrant As New SecurityMgr.clsStringBuilder
    
  fOK = True
  
  For Each objGroup In gObjGroups
    With objGroup
      sGroupName = "[" & .Name & "]"
      
      If .Initialised Then
        
        fSysSecManager = .SystemPermissions.Item("P_MODULEACCESS_SYSTEMMANAGER").Allowed Or _
          .SystemPermissions.Item("P_MODULEACCESS_SECURITYMANAGER").Allowed
        
        For Each objView In .Views
          If objView.Changed Or gbShiftSave Then
            sViewName = objView.Name
            
            If fSysSecManager Then
              sSQL = "GRANT DELETE, INSERT, SELECT, UPDATE ON " & sViewName & " TO " & sGroupName
              gADOCon.Execute sSQL, , adExecuteNoRecords
            Else
              ' Initialise the View Column permissions command strings.
              sSelectGrant.TheString = vbNullString
              sUpdateGrant.TheString = vbNullString
          
              If (objView.SelectPrivilege = giPRIVILEGES_SOMEGRANTED) Or _
                (objView.UpdatePrivilege = giPRIVILEGES_SOMEGRANTED) Then
  
                For Each objColumn In objView.Columns
                  sColumnName = objColumn.Name
                    
                  ' Create the SELECT revoke/deny command.
                  If objColumn.SelectPrivilege Then
                    sSelectGrant.Append IIf(sSelectGrant.Length > 0, ",", vbNullString) & sColumnName
                  End If
                    
                  ' Create the UPDATE revoke/deny command.
                  If objColumn.UpdatePrivilege Then
                    sUpdateGrant.Append IIf(sUpdateGrant.Length > 0, ",", vbNullString) & sColumnName
                  End If
                    
                  ' Reset the 'changed' flag on this column.
                  objColumn.Changed = False
                Next objColumn
                Set objColumn = Nothing
              End If
            
              ' Grant/revoke/deny the View's DELETE permission.
              If objView.DeletePrivilege Then
                sSQL = "GRANT DELETE ON " & sViewName & " TO " & sGroupName
                gADOCon.Execute sSQL, , adExecuteNoRecords
              Else
                sSQL = "REVOKE DELETE ON " & sViewName & " TO " & sGroupName
                gADOCon.Execute sSQL, , adExecuteNoRecords
              End If
            
              ' Grant/revoke/deny the View's INSERT permission.
              If objView.InsertPrivilege Then
                sSQL = "GRANT INSERT ON " & sViewName & " TO " & sGroupName
                gADOCon.Execute sSQL, , adExecuteNoRecords
              Else
                sSQL = "REVOKE INSERT ON " & sViewName & " TO " & sGroupName
                gADOCon.Execute sSQL, , adExecuteNoRecords
              End If
            
              ' Select permissions.
              If objView.SelectPrivilege = giPRIVILEGES_ALLGRANTED Then
                sSQL = "GRANT SELECT ON " & sViewName & " TO " & sGroupName
                gADOCon.Execute sSQL, , adExecuteNoRecords
              ElseIf objView.SelectPrivilege = giPRIVILEGES_SOMEGRANTED Then
                If sSelectGrant.Length <> 0 Then
                  
                  sSQL = "REVOKE SELECT ON " & sViewName & " TO " & sGroupName
                  gADOCon.Execute sSQL, , adExecuteNoRecords
                  
                  sSQL = "GRANT SELECT(ID,TimeStamp," & sSelectGrant.ToString & ") ON " & sViewName & " TO " & sGroupName
                  gADOCon.Execute sSQL, , adExecuteNoRecords
                End If
              ElseIf objView.SelectPrivilege = giPRIVILEGES_NONEGRANTED Then
                  sSQL = "REVOKE SELECT ON " & sViewName & " TO " & sGroupName
                  gADOCon.Execute sSQL, , adExecuteNoRecords
              End If

            
              ' Update permissions
              If objView.UpdatePrivilege = giPRIVILEGES_ALLGRANTED Then
                sSQL = "GRANT UPDATE ON " & sViewName & " TO " & sGroupName
                gADOCon.Execute sSQL, , adExecuteNoRecords
              ElseIf objView.UpdatePrivilege = giPRIVILEGES_SOMEGRANTED Then
                If sUpdateGrant.Length <> 0 Then
                  
                  sSQL = "REVOKE UPDATE ON " & sViewName & " TO " & sGroupName
                  gADOCon.Execute sSQL, , adExecuteNoRecords
                  
                  'MH20060620 Fault 11186
                  'sSQL = "GRANT UPDATE(" & sUpdateGrant.ToString & ") ON " & sViewName & " TO " & sGroupName
                  sSQL = "GRANT UPDATE(ID,Timestamp," & sUpdateGrant.ToString & ") ON " & sViewName & " TO " & sGroupName
                  gADOCon.Execute sSQL, , adExecuteNoRecords
                End If
              ElseIf objView.UpdatePrivilege = giPRIVILEGES_NONEGRANTED Then
                  sSQL = "REVOKE UPDATE ON " & sViewName & " TO " & sGroupName
                  gADOCon.Execute sSQL, , adExecuteNoRecords
              End If
            
              ' Reset the 'changed' flag on this view.
              objView.Changed = False
            
            End If
          End If
        Next objView
        Set objView = Nothing
      End If
    End With
  Next objGroup
  
TidyUpAndExit:
  Set objGroup = Nothing
  Set objView = Nothing
  Set objColumn = Nothing
  ApplyChanges_UserViewPermissions = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function ApplyChanges_SystemPermissions() As Boolean
  ' Apply the System permissions in the SQL Server database.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim sSQL As String
  Dim objGroup As SecurityGroup
  Dim objSystemPermission As clsSystemPermission

  Dim blnSecMgrRW As Boolean
  Dim blnSecMgrRO As Boolean
  Dim blnCMG As Boolean


  fOK = True

  For Each objGroup In gObjGroups
    With objGroup

      ' If the System permissions have been initialised then go through them.
      If .Initialised Then

        ' Delete all System Permission records for this user group.
        sSQL = "DELETE FROM ASRSysGroupPermissions" & _
          " WHERE groupName = '" & .Name & "'"
        gADOCon.Execute sSQL, , adExecuteNoRecords
        
        
        blnSecMgrRW = False
        blnSecMgrRO = False
        blnCMG = False
        
        For iLoop = 1 To .SystemPermissions.Count
          Set objSystemPermission = .SystemPermissions.Item(iLoop)

          ' Creating the System Permissions records for the current user group.
          sSQL = "INSERT INTO ASRSysGroupPermissions" & _
            " (itemID, groupName, permitted)" & _
            " VALUES(" & objSystemPermission.ItemID & ", '" & .Name & "', " & IIf(objSystemPermission.Allowed, "1", "0") & ")"
          gADOCon.Execute sSQL, , adExecuteNoRecords


          Select Case objSystemPermission.ItemKey
          Case "SECURITYMANAGER"
            blnSecMgrRW = objSystemPermission.Allowed
          
          Case "SECURITYMANAGERRO"
            blnSecMgrRO = objSystemPermission.Allowed
          
          Case "CMGRUN", "CMGCOMMIT", "CMGRECOVERY"
            If objSystemPermission.Allowed Then
              blnCMG = True
            End If
          
          End Select
          
          Set objSystemPermission = Nothing
        Next iLoop


        If blnSecMgrRW Then
          SetAuditPermissions "GRANT DELETE, INSERT, UPDATE, SELECT", .Name

        ElseIf blnSecMgrRO Then
          SetAuditPermissions "REVOKE DELETE, INSERT, UPDATE", .Name
          SetAuditPermissions "GRANT SELECT", .Name
          gADOCon.Execute "GRANT UPDATE ON ASRSysAuditTrail(CMGCommitDate, CMGExportDate) TO [" & .Name & "]", , adExecuteNoRecords

        ElseIf blnCMG Then
          SetAuditPermissions "REVOKE DELETE, INSERT, UPDATE, SELECT", .Name
          gADOCon.Execute "GRANT SELECT ON ASRSysAuditTrail TO [" & .Name & "]", , adExecuteNoRecords
          gADOCon.Execute "GRANT UPDATE ON ASRSysAuditTrail(CMGCommitDate, CMGExportDate) TO [" & .Name & "]", , adExecuteNoRecords

        Else
          SetAuditPermissions "REVOKE DELETE, INSERT, UPDATE, SELECT", .Name

        End If


      End If
    End With
  Next objGroup

TidyUpAndExit:
  Set objGroup = Nothing
  Set objSystemPermission = Nothing
  ApplyChanges_SystemPermissions = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Private Sub SetAuditPermissions(strPermission As String, strGroup As String)
        
  gADOCon.Execute strPermission & " ON ASRSysAuditAccess TO [" & strGroup & "]", , adExecuteNoRecords
  gADOCon.Execute strPermission & " ON ASRSysAuditCleardown TO [" & strGroup & "]", , adExecuteNoRecords
  gADOCon.Execute strPermission & " ON ASRSysAuditGroup TO [" & strGroup & "]", , adExecuteNoRecords
  gADOCon.Execute strPermission & " ON ASRSysAuditPermissions TO [" & strGroup & "]", , adExecuteNoRecords
  gADOCon.Execute strPermission & " ON ASRSysAuditTrail TO [" & strGroup & "]", , adExecuteNoRecords

End Sub



Private Function ApplyChanges_NewUserLogins() As Boolean
  ' Create the new User Logins in the SQL Server database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim rsLogins As New ADODB.Recordset
  Dim objGroup As SecurityGroup
  Dim objUser As SecurityUser
  Dim strSQLLoginName As String
  Dim strPassword As String
  'Dim bBypassPolicy As Boolean
  'bBypassPolicy = GetSystemSetting("Policy", "Sec Man Bypass", 0)   ' Default - Off
  
  fOK = True

  For Each objGroup In gObjGroups
    ' If the Users have been initialised then go through them.
    If objGroup.Users_Initialised Then
      For Each objUser In objGroup.Users
        With objUser
          
          strSQLLoginName = .Login
          
          ' Decide if it is a new user login.
          If .Changed And _
            .NewUser And _
            (Not .DeleteUser) Then
            
            ' See if the login already exists in the SQL Server database.
            'TM20011114 Fault - retrieve loginname not just name.
            'JDM - 26/11/01 - Fault 3211 - Strip the login names
            sSQL = "SELECT loginname " & _
              "FROM master.dbo.syslogins " & _
              "WHERE loginname = '" & Replace(strSQLLoginName, "'", "''") & "'"
            rsLogins.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
            
            If (rsLogins.EOF And rsLogins.BOF) Then
              ' The login does not exist in the SQL Server database, so create it.
              
              ' Type of login to create
              If .LoginType = iUSERTYPE_TRUSTEDUSER Or .LoginType = iUSERTYPE_TRUSTEDGROUP Then
                ' AE20080508 Fault #13154
                'sSQL = "sp_grantlogin N'" & strSQLLoginName & "'"
                sSQL = "sp_grantlogin N'" & Replace(strSQLLoginName, "'", "''") & "'"
              Else
                ' JDM - Fault 6709 - SQL 7 doesn't like blank passwords - needs to be passed as null.
                ' AE20080401 Fault #13057 - SQL2005/SQL2008 dont like password passed as null :o)
'                strPassword = IIf(LenB(.Password) = 0, "null", "'" & Replace(.Password, "'", "''") & "'")
                strPassword = Replace(.Password, "'", "''")
                
                If glngSQLVersion >= 9 Then
                   ' AE20080425 Fault #12827
'                  sSQL = "CREATE LOGIN [" & strSQLLoginName & "] WITH PASSWORD ='" & strPassword & "'" _
'                    & IIf(Not bBypassPolicy And .ForcePasswordChange, " MUST_CHANGE", "") _
'                    & ", DEFAULT_DATABASE = [" & gsDatabaseName & "]" _
'                    & IIf(bBypassPolicy, ", CHECK_POLICY = OFF", ", CHECK_POLICY = ON") _
'                    & IIf(bBypassPolicy, ", CHECK_EXPIRATION = OFF", ", CHECK_EXPIRATION = ON")

                  sSQL = "CREATE LOGIN [" & strSQLLoginName & "] WITH PASSWORD ='" & strPassword & "'" _
                    & IIf(.CheckPolicy And .ForcePasswordChange, " MUST_CHANGE", "") _
                    & ", DEFAULT_DATABASE = [" & gsDatabaseName & "]" _
                    & IIf(Not .CheckPolicy, ", CHECK_POLICY = OFF", ", CHECK_POLICY = ON") _
                    & IIf(Not .CheckPolicy, ", CHECK_EXPIRATION = OFF", ", CHECK_EXPIRATION = ON")
                Else
                  strPassword = IIf(LenB(.Password) = 0, "null", "'" & strPassword & "'")
                  sSQL = "sp_addlogin '" & Replace(strSQLLoginName, "'", "''") & "'," & strPassword & ", '" & gsDatabaseName & "', default"
                End If
              End If
              
              gADOCon.Execute sSQL, , adExecuteNoRecords

            End If
          
            rsLogins.Close
            
          End If
        End With
      Next
      Set objUser = Nothing
      
    End If
  Next objGroup

TidyUpAndExit:
  Set objGroup = Nothing
  Set objUser = Nothing
  Set rsLogins = Nothing
  ApplyChanges_NewUserLogins = fOK
  Exit Function

ErrorTrap:

  If gADOCon.Errors.Count > 0 Then
    
  
  End If

  fOK = False
  Resume TidyUpAndExit

End Function

Private Function ApplyChanges_MoveUsers() As Boolean
  ' Apply any User Moves in the SQL Server database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim objGroup As SecurityGroup
  Dim objUser As SecurityUser
  Dim sUserName As String
    
  fOK = True
  
  For Each objGroup In gObjGroups
    
    ' If the Users have been initialised then go through them.
    If objGroup.Users_Initialised Then
      
      For Each objUser In objGroup.Users
        
        With objUser
          ' Check if the User has moved from the current User Group (Role).
          If .Changed And _
            (Not .MovedUserTo = vbNullString) And (.NewUser = False) Then
            
            sUserName = Replace(.UserName, "'", "''")
            
            ' Move the User to the new User Group (Role).

            ' Drop the User from its original User Group (Role).
            sSQL = "sp_droprolemember '" & objGroup.Name & "','" & sUserName & "'"
            'AuditGroup objGroup.Name, "User Deleted", .UserName
            AuditGroup objGroup.Name, "User Moved", .UserName
            gADOCon.Execute sSQL, , adExecuteNoRecords
              
            ' Only add the User to the new User Group (Role) if the new User Group (Role) is not the 'public' User Group (Role).
            ' In SQL Server 7.0 all Users are always in the 'public' User Group (Role).
            If Not .MovedUserTo = "public" Then
              sSQL = "sp_addrolemember '" & .MovedUserTo & "', '" & sUserName & "'"
              AuditGroup .MovedUserTo, "User Added", .UserName
              gADOCon.Execute sSQL, , adExecuteNoRecords
            End If

            ' JDM - 11/04/06 - Fault 11067 - Problem when moving users from sys/sec access group to a normal lowlife group.
            If gObjGroups(.MovedUserTo).SystemPermissions.Item("P_MODULEACCESS_SYSTEMMANAGER").Allowed Or _
                gObjGroups(.MovedUserTo).SystemPermissions.Item("P_MODULEACCESS_SECURITYMANAGER").Allowed Then
              sSQL = "sp_addrolemember 'db_owner', '" & sUserName & "'"
            Else
              sSQL = "sp_droprolemember 'db_owner', '" & sUserName & "'"
            End If
            gADOCon.Execute sSQL, , adExecuteNoRecords

            ' Remove the User from the current User Group collection.
            objGroup.Users.Remove .UserName
          End If
          
          If (.MovedUserFrom <> vbNullString And Not .NewUser) Then
            .Changed = False
            .MovedUserFrom = vbNullString
          End If
          
        End With
      Next
      Set objUser = Nothing
      
    End If
  Next objGroup

TidyUpAndExit:
  Set objGroup = Nothing
  Set objUser = Nothing
  ApplyChanges_MoveUsers = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function ApplyChanges_DeleteUsers() As Boolean
  ' Apply any User deletions in the SQL Server database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim sGroupName As String
  Dim objGroup As SecurityGroup
  Dim objUser As SecurityUser
  Dim rsTempInfo As New ADODB.Recordset
  Dim strSQLUserName As String
  
  fOK = True
  
  For Each objGroup In gObjGroups
    sGroupName = objGroup.Name
 
    ' If the Users have been initialised then go through them.
    If objGroup.Users_Initialised Then
      For Each objUser In objGroup.Users
        With objUser
          ' Check if the user has been deleted.
          If .Changed And _
            .DeleteUser Then
            
'********************************************************************************
'TM20020429 Fault 3690 - Drop any 'ASRSysTemp*' tables the user may own.

            strSQLUserName = LCase$(.UserName)
            
            'drop all temp tables owned by this user...
            sSQL = "SELECT sysobjects.name, sysobjects.type FROM sysobjects " & _
                   "JOIN sysusers ON sysusers.uid = sysobjects.uid " & _
                   "WHERE sysobjects.name LIKE 'ASRSysTemp%' " & _
                   "AND sysusers.name = '" & Replace(strSQLUserName, "'", "''") & "'"
            rsTempInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
            
            Do While Not rsTempInfo.EOF
              
              Select Case Trim(UCase(rsTempInfo.Fields(1).Value))
              Case "U": sSQL = "DROP TABLE "
              Case Else: sSQL = vbNullString
              End Select

              If sSQL <> vbNullString Then
                sSQL = sSQL & "[" & strSQLUserName & "].[" & rsTempInfo.Fields(0).Value & "]"
                gADOCon.Execute sSQL, , adExecuteNoRecords
              End If

              rsTempInfo.MoveNext
            Loop

            rsTempInfo.Close
            sSQL = vbNullString

'********************************************************************************

            sSQL = "sp_revokedbaccess [" & strSQLUserName & "]"
            AuditGroup sGroupName, "User Deleted", .UserName
            gADOCon.Execute sSQL, , adExecuteNoRecords

            ' Remove this user from the current group.
            objGroup.Users.Remove .UserName
          End If
        End With
      Next objUser
      
      Set objUser = Nothing
    End If
  Next

TidyUpAndExit:
  Set rsTempInfo = Nothing
  Set objGroup = Nothing
  ApplyChanges_DeleteUsers = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Private Function ApplyChanges_NewUsers() As Boolean
  ' Apply any User Additions in the SQL Server database.
  On Error GoTo 0 'ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim sGroupName As String
  Dim sLastUserDropped As String
  Dim objGroup As SecurityGroup
  Dim objUser As SecurityUser
  Dim rsUserInfo As New ADODB.Recordset
  Dim rsTempInfo As New ADODB.Recordset
  Dim strSQLUserName As String
  Dim strSQLLoginName As String
  Dim asUsers() As String
  Dim iLoop As Integer

  fOK = True

  For Each objGroup In gObjGroups
    sGroupName = objGroup.Name
 
    ' If the Users have been initialised then go through them.
    If objGroup.Users_Initialised Then

      For Each objUser In objGroup.Users
        With objUser
          ' Check if the user has been deleted.
          If .Changed And _
            .NewUser Then
            
            strSQLUserName = LCase$(Replace(.UserName, "'", "''"))
            strSQLLoginName = LCase$(Replace(.Login, "'", "''"))
    
            ' Drop the user if it already exists in any other User Groups (Roles),
            ' but not the 'public' role.
            
            ' JPD20020429 Fault 3712 - Error occurred on some machines when we ran
            ' the 'SELECT sysobjects.name...' query whilst looping through the
            ' 'sp_helpuser' recordset. Got around this error by reading the 'sp_helpuser'
            ' recordset records into an array first, closing the recordset, and then calling the
            ' 'SELECT sysobjects.name...' whilst looping through the array. Not pretty but it works.
            ReDim asUsers(0)
            
            sLastUserDropped = vbNullString
            sSQL = "sp_helpuser"

            rsUserInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
            With rsUserInfo
              Do While Not .EOF
                ReDim Preserve asUsers(UBound(asUsers) + 1)
                asUsers(UBound(asUsers)) = !UserName
                
                .MoveNext
              Loop
              .Close
            End With
            
            ' JPD20020429 Fault 3712
            For iLoop = 1 To UBound(asUsers)
              'JPD 20030428 Fault 5460
              'If (LCase(asUsers(iLoop)) = LCase(objUser.UserName)) And _
                (asUsers(iLoop) <> sLastUserDropped) Then
              If (LCase$(asUsers(iLoop)) = LCase$(objUser.UserName)) And _
                (LCase$(asUsers(iLoop)) <> LCase$(sLastUserDropped)) Then
            
                'MH20011017
                'Drop all objects (only tables at the mo) which are owned by this user...
                sSQL = "SELECT sysobjects.name, sysobjects.type FROM sysobjects " & _
                       "JOIN sysusers ON sysusers.uid = sysobjects.uid " & _
                       "WHERE sysobjects.name LIKE 'ASRSysTemp%' " & _
                       "AND sysusers.name = '" & Replace(objUser.UserName, "'", "''") & "'"
                rsTempInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
            
                Do While Not rsTempInfo.EOF
            
                  'Only dropping tables at the moment but might need to add stored procedures ??
                  Select Case Trim(UCase(rsTempInfo.Fields(1).Value))
                  Case "U": sSQL = "DROP TABLE "
                  Case Else: sSQL = vbNullString
                  End Select
            
                  If sSQL <> vbNullString Then
                    ' JPD20020806 Fault 4269
                    'sSQL = sSQL & strSQLUserName & "." & rsTempInfo.rdoColumns(0).Value
                    sSQL = sSQL & "[" & strSQLUserName & "]." & rsTempInfo.Fields(0).Value
                    gADOCon.Execute sSQL, , adExecuteNoRecords
                    
                  End If
            
                  rsTempInfo.MoveNext
                Loop
            
                rsTempInfo.Close
            
                ' JDM - Fault 8938/8784 - dropuser doesn't seem to work when a user doesn't have a login
                'sSQL = "sp_dropuser '" & strSQLUserName & "'"
                sSQL = "sp_revokedbaccess '" & strSQLUserName & "'"
                gADOCon.Execute sSQL, , adExecuteNoRecords
            
                sLastUserDropped = objUser.UserName
              End If
            Next iLoop
            

            ' Grant the db access
            sSQL = "sp_grantdbaccess '" & strSQLLoginName & "','" & strSQLUserName & "'"
            On Error Resume Next
            gADOCon.Execute sSQL, , adExecuteNoRecords
            On Error GoTo ErrorTrap
            sSQL = "sp_addrolemember '" & sGroupName & "', '" & strSQLUserName & "'"
            AuditGroup sGroupName, "User Added", strSQLUserName
            gADOCon.Execute sSQL, , adExecuteNoRecords

            'MH20031106 Fault 5627
            sSQL = "sp_addrolemember 'ASRSysGroup', '" & strSQLUserName & "'"
            gADOCon.Execute sSQL, , adExecuteNoRecords

            ' Flag that the user has now been applied to the database.
            .Changed = False
            .NewUser = False
          End If
        End With
      Next
      Set objUser = Nothing
      
    End If
  Next

TidyUpAndExit:
  Set rsUserInfo = Nothing
  Set rsTempInfo = Nothing
  Set objGroup = Nothing
  Set objUser = Nothing
  ApplyChanges_NewUsers = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function




Private Function ApplyChanges_DatabaseOwnership() As Boolean
  ' Apply any Database Ownership in the SQL Server database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim objGroup As SecurityGroup
  Dim objUser As SecurityUser
  Dim rsTemp1 As New ADODB.Recordset
  Dim rsTemp2 As New ADODB.Recordset
  Dim sGroup As String
  Dim strType As String
  Dim strName As String
  Dim sSQLCommand As New SecurityMgr.clsStringBuilder
  Dim bAccordModuleEnabled As Boolean
  Dim bWorkflowModuleEnabled As Boolean
  Dim strAccordTables As String
  Dim strAuditTables As String
  Dim bGranted As Boolean
  Dim lngCount As Long
  
  bAccordModuleEnabled = IsModuleEnabled(modAccord)
  bWorkflowModuleEnabled = IsModuleEnabled(modWorkflow)
  
  ' Groups are only given write access to the tables in this array unless they have Accord View Permissions
  Dim astrAccordTables(2) As String
  astrAccordTables(0) = "ASRSysAccordTransactionData"
  astrAccordTables(1) = "ASRSysAccordTransactions"
  astrAccordTables(2) = "ASRSysAccordTransactionWarnings"


  Dim astrAuditTables(4) As String
  astrAuditTables(0) = "ASRSysAuditAccess"
  astrAuditTables(1) = "ASRSysAuditCleardown"
  astrAuditTables(2) = "ASRSysAuditGroup"
  astrAuditTables(3) = "ASRSysAuditPermissions"
  astrAuditTables(4) = "ASRSysAuditTrail"

  strAccordTables = Join(astrAccordTables, ",")
  strAuditTables = Join(astrAuditTables, ",")

  fOK = True

  ' Only apply Database Ownership for SQL Server 7.0 databases.
  ' Database Ownership in other versions is handled by aliasing the users to the 'dbo' user.
  
  OutputCurrentProcess2 vbNullString, gObjGroups.Count + 1
  For Each objGroup In gObjGroups
    
    OutputCurrentProcess2 objGroup.Name
    gobjProgress.UpdateProgress2
    
    ' JPD20030205 Fault 5022
    If objGroup.Initialised Then
      If (Not objGroup.SystemPermissions.Item("P_MODULEACCESS_SYSTEMMANAGER").Allowed) And _
        (Not objGroup.SystemPermissions.Item("P_MODULEACCESS_SECURITYMANAGER").Allowed) Then
        
        sSQLCommand.TheString = "GRANT CREATE PROCEDURE TO [" & objGroup.Name & "]" & vbNewLine
        'gADOCon.Execute sSQL, , adExecuteNoRecords
        
        sSQLCommand.Append "GRANT CREATE TABLE TO [" & objGroup.Name & "]" & vbNewLine
        'gADOCon.Execute sSQL, , adExecuteNoRecords
                
        If gbEnableUDFFunctions Then
          sSQLCommand.Append "GRANT CREATE FUNCTION TO [" & objGroup.Name & "]" & vbNewLine
'          gADOCon.Execute sSQL, , adExecuteNoRecords
        End If
                
        ' Get all dbo owned tables/views/procedures etc and grant permission on them
        sSQL = "SELECT id INTO #tmpProtects FROM sysprotects" _
          & " INNER JOIN sysusers ON sysprotects.uid = sysusers.uid AND sysusers.name = '" _
          & objGroup.Name & "'"
        gADOCon.Execute sSQL, , adExecuteNoRecords
        
        sSQL = "SELECT sysobjects.name, sysobjects.xtype" & _
          " FROM sysobjects" & _
          " INNER JOIN sysusers ON sysobjects.uid = sysusers.uid" & _
          " WHERE (((sysobjects.xtype = 'p') AND (sysobjects.name LIKE 'sp_asr%' OR sysobjects.name LIKE 'spasr%'))" & _
          "    OR ((sysobjects.xtype = 'fn') AND (sysobjects.name LIKE 'udf_ASRFn%')))" & _
          "    AND (sysusers.name = 'dbo')" & _
          "    AND sysobjects.id NOT IN (SELECT id FROM #tmpProtects)"
        
        rsTemp1.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
        Do While Not rsTemp1.EOF
        
          strType = UCase(Trim(rsTemp1.Fields(1).Value))
          strName = rsTemp1.Fields(0).Value
        
          If strType = "P" Or strType = "FN" Then
            sSQLCommand.Append "GRANT EXEC ON [" & strName & "] TO [" & objGroup.Name & "]" & vbNewLine
          Else
            
            'MH20071210 Fault 5141
            If InStr(strAuditTables, strName) = 0 Then
              sSQLCommand.Append "GRANT SELECT, INSERT, UPDATE, DELETE ON [" & strName & "] TO [" & objGroup.Name & "]" & vbNewLine
            End If
          
          End If
          
          rsTemp1.MoveNext
        Loop
        rsTemp1.Close
        Set rsTemp1 = Nothing
        
      
        ' Assign permissions on Accord system tables
        If bAccordModuleEnabled Then
          bGranted = objGroup.SystemPermissions.Item("P_ACCORD_VIEWTRANSFER").Allowed
        Else
          bGranted = False
        End If
        
        For lngCount = LBound(astrAccordTables) To UBound(astrAccordTables)
          If bGranted Then
            sSQLCommand.Append "GRANT SELECT, INSERT, UPDATE, DELETE ON [" & astrAccordTables(lngCount) & "] TO [" & objGroup.Name & "]" & vbNewLine
          Else
            sSQLCommand.Append "REVOKE SELECT, INSERT, UPDATE, DELETE ON [" & astrAccordTables(lngCount) & "] TO [" & objGroup.Name & "]" & vbNewLine
          End If
        Next lngCount
            
        ' Run all the statements
        gADOCon.Execute sSQLCommand.ToString, , adExecuteNoRecords
        
        sSQL = "DROP TABLE #tmpProtects"
        gADOCon.Execute sSQL, , adExecuteNoRecords
             

'MH20071108 Fault 5141
'      Else
'        sSQL = "GRANT SELECT, INSERT, UPDATE, DELETE ON "
'
'        sGroup = " TO [" & objGroup.Name & "]"
'        gADOCon.Execute sSQL & "ASRSysAuditAccess" & sGroup, , adExecuteNoRecords
'        gADOCon.Execute sSQL & "ASRSysAuditCleardown" & sGroup, , adExecuteNoRecords
'        gADOCon.Execute sSQL & "ASRSysAuditGroup" & sGroup, , adExecuteNoRecords
'        gADOCon.Execute sSQL & "ASRSysAuditPermissions" & sGroup, , adExecuteNoRecords
'        gADOCon.Execute sSQL & "ASRSysAuditTrail" & sGroup, , adExecuteNoRecords
      End If
          
      ' If the Users have been initialised then go through them.
      If objGroup.Users_Initialised Then
        For Each objUser In objGroup.Users
          With objUser
            ' JPD20020625 Fault 4059 - Changed way we link sysuser to syslogins as the suid
            ' does not exist in SQL2000. Oops, should've known that.
            sSQL = "SELECT sysusers.name AS result" & _
              " FROM sysusers" & _
              " INNER JOIN master..syslogins ON sysusers.sid = master..syslogins.sid" & _
              " WHERE master..syslogins.loginname = '" & Replace(.UserName, "'", "''") & "'"
            
            ' JPD20020430 Fault 3776 - sp_addrolemember needs the user name as the
            ' parameter, rather than the login name which were we using. This caused
            ' an error when the user and login names were not the same.
            'sSQL = "SELECT sid AS result FROM master..syslogins WHERE loginname = '" & Replace(.UserName, "'", "''") & "'"
            'Set rsTemp = rdoCon.OpenResultset(sSQL, rdOpenForwardOnly)
            '
            'If Not (rsTemp.BOF And rsTemp.EOF) Then
              'sSQL = "SELECT name AS result FROM sysusers WHERE sid = " & rsTemp!result
            rsTemp2.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
                
            If Not (rsTemp2.BOF And rsTemp2.EOF) Then
              ' JPD20030205 Fault 5022
              If objGroup.SystemPermissions.Item("P_MODULEACCESS_SYSTEMMANAGER").Allowed Or _
                objGroup.SystemPermissions.Item("P_MODULEACCESS_SECURITYMANAGER").Allowed Then
                sSQL = "sp_addrolemember 'db_owner', '" & Replace(rsTemp2!Result, "'", "''") & "'"
              Else
                sSQL = "sp_droprolemember 'db_owner', '" & Replace(rsTemp2!Result, "'", "''") & "'"
              End If
              gADOCon.Execute sSQL, , adExecuteNoRecords
            End If
              
            rsTemp2.Close
            Set rsTemp2 = Nothing
          End With
        Next
        Set objUser = Nothing
      End If
    End If
  Next objGroup
  
TidyUpAndExit:
  Set sSQLCommand = Nothing
  Set objGroup = Nothing
  Set objUser = Nothing
  ApplyChanges_DatabaseOwnership = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function ApplyChanges_DeleteGroups() As Boolean
  ' Apply any User Group (Role) deletions in the SQL Server database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim objGroup As SecurityGroup
  Dim rsUsersInGroup As New ADODB.Recordset
  Dim rsTempInfo As New ADODB.Recordset
  Dim rsInfo As New ADODB.Recordset
  Dim rsTemp As New ADODB.Recordset
  Dim strSQLUserName As String
  Dim sLoginName As String
  
  fOK = True
  
  For Each objGroup In gObjGroups
    With objGroup
    
      ' Check if the User Group (Role) has been deleted
      If .Changed And _
        .DeleteGroup Then
          
        ' Drop the existing child views for the deleted user group.
        sSQL = "SELECT childViewID FROM ASRSysChildViews2 WHERE role = '" & .Name & "'"
        rsTemp.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
        Do While Not rsTemp.EOF
          sSQL = "SELECT name FROM sysobjects WHERE name LIKE 'ASRSysCV" & Trim(Str(rsTemp!childViewID)) & "#%#" & Replace(.Name, " ", "_") & "' AND xtype = 'V'"
          rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
          
          With rsInfo
            Do While (Not .EOF)
              sSQL = "DROP VIEW " & !Name
              gADOCon.Execute sSQL, , adExecuteNoRecords
          
              .MoveNext
            Loop
          
            .Close
          End With
          Set rsInfo = Nothing
        
          rsTemp.MoveNext
        Loop
        
        rsTemp.Close
        
        ' Delete the
        sSQL = "DELETE FROM ASRSysChildViewParents2 WHERE childViewID IN (SELECT childViewID FROM ASRSysChildViews2 WHERE role = '" & .Name & "')"
        gADOCon.Execute sSQL, , adExecuteNoRecords
        
        sSQL = "DELETE FROM ASRSysChildViews2 WHERE role = '" & .Name & "'"
        gADOCon.Execute sSQL, , adExecuteNoRecords
  
        RemoveGroupAccessRecords (.Name)
        
        ' Delete any users that are still in the User Group (Role).
        ' Get the list of database owners.
                              
        'JPD 20030815 Fault 6244
        ' JPD20030206 Fault 5023
        'sSQL = "SELECT master..syslogins.loginname AS result" & _
          " FROM sysusers roles" & _
          " INNER JOIN sysusers users ON roles.uid = users.gid" & _
          " INNER JOIN master..syslogins ON users.sid = master..syslogins.sid" & _
          " WHERE roles.name = '" & Replace(.Name, "'", "''") & "'"
        'JPD 20040224 Fault 8132
        'sSQL = "SELECT users.name AS result" & _
          " FROM sysusers roles" & _
          " INNER JOIN sysusers users ON roles.uid = users.gid" & _
          " WHERE roles.name = '" & Replace(.Name, "'", "''") & "'" & _
          "   AND users.gid <> users.uid"
        sSQL = "SELECT usu.name AS result" & _
          " FROM sysusers usg" & _
          " INNER JOIN sysmembers mem ON usg.uid = mem.groupuid" & _
          " INNER JOIN sysusers usu ON mem.memberuid = usu.uid" & _
          " WHERE usg.name = '" & Replace(.Name, "'", "''") & "'"

        rsUsersInGroup.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
        Do While Not rsUsersInGroup.EOF
          If Not IsNull(rsUsersInGroup!Result) Then
            sLoginName = rsUsersInGroup!Result
            
            If LenB(sLoginName) > 0 Then
              strSQLUserName = LCase$(Replace(sLoginName, "'", "''"))
            
              'drop all temp tables owned by this user...
              sSQL = "SELECT sysobjects.name, sysobjects.type FROM sysobjects " & _
                     "JOIN sysusers ON sysusers.uid = sysobjects.uid " & _
                     "WHERE sysobjects.name LIKE 'ASRSysTemp%' " & _
                     "AND sysusers.name = '" & strSQLUserName & "'"
        
'********************************************************************************
'TM20020429 Fault 3690 - Drop any 'ASRSysTemp*' tables the user may own.

'          strSQLUserName = LCase(Replace(rsUsersInGroup!Users_in_group, "'", "''"))
'
'          'drop all temp tables owned by this user...
'          sSQL = "SELECT sysobjects.name, sysobjects.type FROM sysobjects " & _
'                 "JOIN sysusers ON sysusers.uid = sysobjects.uid " & _
'                 "WHERE sysobjects.name LIKE 'ASRSysTemp%' " & _
'                 "AND sysusers.name = '" & rsUsersInGroup!Users_in_group & "'"
              rsTempInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
              Do While Not rsTempInfo.EOF
                
                Select Case Trim(UCase(rsTempInfo.Fields(1).Value))
                Case "U": sSQL = "DROP TABLE "
                Case Else: sSQL = vbNullString
                End Select
      
                If sSQL <> vbNullString Then
                  sSQL = sSQL & "[" & strSQLUserName & "].[" & rsTempInfo.Fields(0).Value & "]"
                  gADOCon.Execute sSQL, , adExecuteNoRecords
                End If
      
                rsTempInfo.MoveNext
              Loop
      
              rsTempInfo.Close
        
              ' Drop the User from this User Group (Role).
              sSQL = "sp_droprolemember '" & .Name & "','" & strSQLUserName & "'"
              AuditGroup .Name, "User Deleted", strSQLUserName
              gADOCon.Execute sSQL, , adExecuteNoRecords

            End If
          End If
          
          rsUsersInGroup.MoveNext
        Loop
        rsUsersInGroup.Close
        
        ' Drop the role
        AuditGroup .Name, "User Group Deleted"
        sSQL = "sp_droprole '" & .Name & "'"
        gADOCon.Execute sSQL, , adExecuteNoRecords
            
        ' Delete the System Permissions for the User Group.
        sSQL = "DELETE FROM ASRSysGroupPermissions" & _
          " WHERE groupName = '" & .Name & "'"
        gADOCon.Execute sSQL, , adExecuteNoRecords
        
        gObjGroups.Remove .Name
      End If
    End With
  Next

TidyUpAndExit:
  Set rsTemp = Nothing
  Set rsUsersInGroup = Nothing
  Set rsTempInfo = Nothing
  
  Set objGroup = Nothing
  Set rsUsersInGroup = Nothing
  ApplyChanges_DeleteGroups = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function ApplyChanges_CleanUp() As Boolean
  ' Clean up the SQL Server database.
  ' ie. Ensure that the SQL Server 7.0 'public' User Group (Role) has all permission revoked
  ' so that the permission we've just applied take effect.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iNextIndex As Integer
  Dim lngPublicUID As Long
  Dim lngInfoSchemaUID As Long
  Dim sSQL As String
  Dim rsPermissions As New ADODB.Recordset
  Dim asSQL() As String

  fOK = True
  ReDim asSQL(0)
  
  ' Revoke all permissions on the 'public' group.
  
  ' Get the uid of the 'public' group.
  sSQL = "SELECT uid" & _
    " FROM sysusers" & _
    " WHERE name = 'public'"

  rsPermissions.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
  If Not (rsPermissions.EOF And rsPermissions.BOF) Then
    lngPublicUID = rsPermissions!uid
  End If
  rsPermissions.Close

  ' Get the uid of the 'INFORMATION_SCHEMA' group.
  sSQL = "SELECT uid" & _
    " FROM sysusers" & _
    " WHERE name = 'INFORMATION_SCHEMA'"
  rsPermissions.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
  If Not (rsPermissions.EOF And rsPermissions.BOF) Then
    lngInfoSchemaUID = rsPermissions!uid
  End If
  rsPermissions.Close
  
  sSQL = "SELECT sysobjects.name," & _
    " CASE" & _
    "  WHEN sysprotects.action = 26 THEN 'REFERENCES'" & _
    "  WHEN sysprotects.action = 193 THEN 'SELECT'" & _
    "  WHEN sysprotects.action = 195 THEN 'INSERT'" & _
    "  WHEN sysprotects.action = 196 THEN 'DELETE'" & _
    "  WHEN sysprotects.action = 197 THEN 'UPDATE'" & _
    "  WHEN sysprotects.action = 198 THEN 'CREATE TABLE'" & _
    "  WHEN sysprotects.action = 203 THEN 'CREATE DATABASE'" & _
    "  WHEN sysprotects.action = 207 THEN 'CREATE VIEW'" & _
    "  WHEN sysprotects.action = 222 THEN 'CREATE PROCEDURE'" & _
    "  WHEN sysprotects.action = 224 THEN 'EXECUTE'" & _
    "  WHEN sysprotects.action = 228 THEN 'BACKUP DATABASE'" & _
    "  WHEN sysprotects.action = 233 THEN 'CREATE DEFAULT'" & _
    "  WHEN sysprotects.action = 235 THEN 'BACKUP LOG'" & _
    "  WHEN sysprotects.action = 236 THEN 'CREATE RULE'" & _
    "  ELSE ''" & _
    " END As action" & _
    " FROM sysprotects" & _
    " INNER JOIN sysobjects ON sysprotects.ID = sysobjects.ID" & _
    " WHERE sysprotects.UID = " & Trim(Str(lngPublicUID)) & _
    " AND NOT sysprotects.grantor = " & Trim(Str(lngInfoSchemaUID))
  
  sSQL = sSQL & _
    " AND NOT sysobjects.name IN ('sysalternates'," & _
    "   'syscolumns'," & _
    "   'syscomments'," & _
    "   'sysconstraints'," & _
    "   'sysdepends'," & _
    "   'sysfilegroups'," & _
    "   'sysfiles'," & _
    "   'sysforeignkeys'," & _
    "   'sysfulltextcatalogs'," & _
    "   'sysindexes'," & _
    "   'sysindexkeys'," & _
    "   'sysmembers'," & _
    "   'sysobjects'," & _
    "   'syspermissions'," & _
    "   'sysprotects'," & _
    "   'sysreferences'," & _
    "   'syssegments'," & _
    "   'systypes'," & _
    "   'sysusers'," & _
    "   '.')"

  rsPermissions.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
  
  With rsPermissions
    Do While Not .EOF
      If Len(!Action) > 0 Then
        sSQL = "REVOKE " & !Action & " ON " & !Name & " FROM [public]"
        iNextIndex = UBound(asSQL) + 1
        ReDim Preserve asSQL(iNextIndex)
        asSQL(iNextIndex) = sSQL
      End If
      
      .MoveNext
    Loop
    
    .Close
  End With

  For iNextIndex = 1 To UBound(asSQL)
    gADOCon.Execute asSQL(iNextIndex), , adExecuteNoRecords
  Next iNextIndex

  ' NEWACCESS - needs to be updated as each report/utility is updated for the new access.
  sSQL = "DELETE FROM ASRSysBatchJobAccess" & _
    " WHERE groupName NOT IN" & _
    "   (SELECT name" & _
    "    FROM sysusers" & _
    "    WHERE gid = uid" & _
    "      AND uid <> 0)"
  gADOCon.Execute sSQL, , adExecuteNoRecords
  
  sSQL = "DELETE FROM ASRSysCalendarReportAccess" & _
    " WHERE groupName NOT IN" & _
    "   (SELECT name" & _
    "    FROM sysusers" & _
    "    WHERE gid = uid" & _
    "      AND uid <> 0)"
  gADOCon.Execute sSQL, , adExecuteNoRecords
  
  sSQL = "DELETE FROM ASRSysCrossTabAccess" & _
    " WHERE groupName NOT IN" & _
    "   (SELECT name" & _
    "    FROM sysusers" & _
    "    WHERE gid = uid" & _
    "      AND uid <> 0)"
  gADOCon.Execute sSQL, , adExecuteNoRecords
  
  sSQL = "DELETE FROM ASRSysCustomReportAccess" & _
    " WHERE groupName NOT IN" & _
    "   (SELECT name" & _
    "    FROM sysusers" & _
    "    WHERE gid = uid" & _
    "      AND uid <> 0)"
  gADOCon.Execute sSQL, , adExecuteNoRecords
  
  sSQL = "DELETE FROM ASRSysDataTransferAccess" & _
    " WHERE groupName NOT IN" & _
    "   (SELECT name" & _
    "    FROM sysusers" & _
    "    WHERE gid = uid" & _
    "      AND uid <> 0)"
  gADOCon.Execute sSQL, , adExecuteNoRecords
  
  sSQL = "DELETE FROM ASRSysExportAccess" & _
    " WHERE groupName NOT IN" & _
    "   (SELECT name" & _
    "    FROM sysusers" & _
    "    WHERE gid = uid" & _
    "      AND uid <> 0)"
  gADOCon.Execute sSQL, , adExecuteNoRecords
  
  sSQL = "DELETE FROM ASRSysGlobalAccess" & _
    " WHERE groupName NOT IN" & _
    "   (SELECT name" & _
    "    FROM sysusers" & _
    "    WHERE gid = uid" & _
    "      AND uid <> 0)"
  gADOCon.Execute sSQL, , adExecuteNoRecords
  
  sSQL = "DELETE FROM ASRSysImportAccess" & _
    " WHERE groupName NOT IN" & _
    "   (SELECT name" & _
    "    FROM sysusers" & _
    "    WHERE gid = uid" & _
    "      AND uid <> 0)"
  gADOCon.Execute sSQL, , adExecuteNoRecords
  
  sSQL = "DELETE FROM ASRSysMailMergeAccess" & _
    " WHERE groupName NOT IN" & _
    "   (SELECT name" & _
    "    FROM sysusers" & _
    "    WHERE gid = uid" & _
    "      AND uid <> 0)"
  gADOCon.Execute sSQL, , adExecuteNoRecords
  
  sSQL = "DELETE FROM ASRSysRecordProfileAccess" & _
    " WHERE groupName NOT IN" & _
    "   (SELECT name" & _
    "    FROM sysusers" & _
    "    WHERE gid = uid" & _
    "      AND uid <> 0)"
  gADOCon.Execute sSQL, , adExecuteNoRecords
  
  sSQL = "DELETE FROM ASRSysMatchReportAccess" & _
    " WHERE groupName NOT IN" & _
    "   (SELECT name" & _
    "    FROM sysusers" & _
    "    WHERE gid = uid" & _
    "      AND uid <> 0)"
  gADOCon.Execute sSQL, , adExecuteNoRecords
  
  'JPD 20071203 Faults 12580, 12670, 12671
  sSQL = "DELETE FROM ASRSysSSIHiddenGroups" & _
    " WHERE groupName NOT IN" & _
    "   (SELECT name" & _
    "    FROM sysusers" & _
    "    WHERE gid = uid" & _
    "      AND uid <> 0)"
  gADOCon.Execute sSQL, , adExecuteNoRecords

  ' JDM - 26/06/06 - Fault 10725 - Cleanup users that aren't on the server
  gADOCon.Execute "EXEC spASRTidyUpNonASRUsers", , adExecuteNoRecords


TidyUpAndExit:
  Set rsPermissions = Nothing
  
  ApplyChanges_CleanUp = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function PermittedChildView(psTableName As String, psGroupName As String) As Long
'  ' Return the ID of the child view on the given table that is appropriate for the given group (role).
'  ' Return 0 if no view is appropriate.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim iLoop As Integer
'  Dim iNextIndex As Integer
'  Dim lngChildViewID As Long
'  Dim lngParentViewID As Long
'  Dim sSQL As String
'  Dim sCode As String
'  Dim avParents() As Variant
'  Dim iParentJoinType As Integer
'  Dim iParentCount As Integer
'  Dim iOKViewCount As Integer
'
'  fOK = True
'  lngChildViewID = 0
'  iParentCount = 0
'
'  ' Create an array of the parents of the given table that are accessible by the given group.
'  ' Column 1 = parent type (UT = top-level table
'  '                         UV = view of a top-level table
'  '                         SV = system view)
'  ' Column2 = parent ID
'  ReDim avParents(2, 0)
'
'  ' Get the given table's parents.
'  sSQL = "SELECT parentTable.tableID, parentTable.tableName, parentTable.tableType" & _
'    " FROM ASRSysRelations" & _
'    " INNER JOIN ASRSysTables parentTable ON ASRSysRelations.parentID = parentTable.tableID" & _
'    " INNER JOIN ASRSysTables childTable ON ASRSysRelations.childID = childTable.tableID" & _
'    " WHERE childTable.tableName = '" & psTableName & "'"
'
'  Set rsParents = rdoCon.OpenResultset(sSQL)
'  With rsParents
'    ' Loop through the given table's parents, adding the permitted view of each to the array of parents.
'    Do While Not .EOF
'      If !TableType = tabParent Then
'        ' Parent is a top-level table.
'        If gObjGroups(psGroupName).Tables(!TableName).SelectPrivilege <> giPRIVILEGES_NONEGRANTED Then
'          ' The current group has permission to see all records in the parent table.
'          iParentCount = iParentCount + 1
'
'          iNextIndex = UBound(avParents, 2) + 1
'          ReDim Preserve avParents(2, iNextIndex)
'          avParents(1, iNextIndex) = "UT"
'          avParents(2, iNextIndex) = !TableID
'        Else
'          ' The current group does NOT have permission to see all records in the parent table.
'          ' Get the permitted views on the table.
'          iOKViewCount = 0
'
'          sSQL = "SELECT ASRSysViews.viewID, ASRSysViews.viewName" & _
'            " FROM ASRSysViews" & _
'            " WHERE ASRSysViews.viewTableID = " & Trim(Str(!TableID))
'          Set rsViews = rdoCon.OpenResultset(sSQL)
'          With rsViews
'            Do While Not .EOF
'              If gObjGroups(psGroupName).Views(!ViewName).SelectPrivilege <> giPRIVILEGES_NONEGRANTED Then
'                iOKViewCount = iOKViewCount + 1
'
'                iNextIndex = UBound(avParents, 2) + 1
'                ReDim Preserve avParents(2, iNextIndex)
'                avParents(1, iNextIndex) = "UV"
'                avParents(2, iNextIndex) = !ViewID
'              End If
'
'              .MoveNext
'            Loop
'
'            .Close
'          End With
'          Set rsViews = Nothing
'
'          If iOKViewCount > 0 Then
'            iParentCount = iParentCount + 1
'          End If
'        End If
'
'      ElseIf !TableType = tabChild Then
'        ' Parent is not a top-level table.
'        lngParentViewID = PermittedChildView(!TableName, psGroupName)
'
'        If lngParentViewID > 0 Then
'          iParentCount = iParentCount + 1
'
'          iNextIndex = UBound(avParents, 2) + 1
'          ReDim Preserve avParents(2, iNextIndex)
'          avParents(1, iNextIndex) = "SV"
'          avParents(2, iNextIndex) = lngParentViewID
'        End If
'      End If
'
'      .MoveNext
'    Loop
'
'    .Close
'  End With
'  Set rsParents = Nothing
'
'  iParentJoinType = 0
'  If iParentCount > 1 Then
'    ' More than 1 parent. Do we want the OR join child view, or the AND join child view ?
'    iParentJoinType = gObjGroups(psGroupName).Tables(psTableName).ParentJoinType
'  End If
'
'  If UBound(avParents, 2) > 0 Then
'    ' Get the child view permutation that is configured for the permitted set of parents.
'    For iLoop = 1 To UBound(avParents, 2)
'      sCode = sCode & _
'        " INNER JOIN ASRSysChildViewParents tmpTable_" & Trim(Str(iLoop)) & _
'        " ON (ASRSysChildViews.childViewID = tmpTable_" & Trim(Str(iLoop)) & ".childViewID" & _
'        " AND tmpTable_" & Trim(Str(iLoop)) & ".parentType = '" & avParents(1, iLoop) & "'" & _
'        " AND tmpTable_" & Trim(Str(iLoop)) & ".parentID = " & Trim(Str(avParents(2, iLoop))) & ")"
'    Next iLoop
'
'    sSQL = "SELECT ASRSysChildViews.childViewID" & _
'      " FROM ASRSysChildViews" & _
'      sCode & _
'      " INNER JOIN ASRSysTables ON ASRSysChildViews.tableID = ASRSysTables.tableID" & _
'      " INNER JOIN ASRSysChildViewParents parentCount" & _
'      " ON (ASRSysChildViews.childViewID = parentCount.childViewID)" & _
'      " GROUP BY ASRSysChildViews.childViewID, ASRSysTables.tableName, ASRSysChildViews.type" & _
'      " HAVING ASRSysTables.tableName = '" & psTableName & "'" & _
'      " AND " & IIf(iParentJoinType = 0, "(ASRSysChildViews.type = 0 OR ASRSysChildViews.type IS NULL)", "ASRSysChildViews.type = 1") & _
'      " AND COUNT(parentCount.childViewID) = " & Trim(Str(UBound(avParents, 2)))
'
'    Set rsChildView = rdoCon.OpenResultset(sSQL)
'    fOK = Not (rsChildView.BOF And rsChildView.EOF)
'    If fOK Then
'      lngChildViewID = rsChildView!childViewID
'    End If
'  End If
'
'TidyUpAndExit:
'  If fOK Then
'    PermittedChildView = lngChildViewID
'  Else
'    PermittedChildView = 0
'  End If
'
'  Exit Function
'
'ErrorTrap:
'  fOK = False
'  Resume TidyUpAndExit
'
End Function


Public Sub FlagChildrenChanged(ByRef pobjTableView As SecurityTable, psGroupName As String)
  ' Flag all of the children of the given table/view as changed.
  ' This is done as the permitted child views on the children need to be recalculated.
  Dim sSQL As String
  Dim rsInfo As New ADODB.Recordset
    
  ' Get the list of children of the given table/view.
  If (pobjTableView.TableType = tabParent) Or _
    (pobjTableView.TableType = tabChild) Then
    sSQL = "SELECT children.tableName" & _
      " FROM ASRSysRelations" & _
      " INNER JOIN ASRSysTables parent ON ASRSysRelations.parentID = parent.tableID" & _
      " INNER JOIN ASRSysTables children ON ASRSysRelations.childID = children.tableID" & _
      " WHERE parent.tableName = '" & pobjTableView.Name & "'"
  Else
    sSQL = "SELECT ASRSysTables.tableName" & _
      " FROM ASRSysRelations" & _
      " INNER JOIN ASRSysViews ON ASRSysRelations.parentID = ASRSysViews.viewTableID" & _
      " INNER JOIN ASRSysTables ON ASRSysRelations.childID = ASRSysTables.tableID" & _
      " WHERE ASRSysViews.viewName = '" & pobjTableView.Name & "'"
  End If
  
  rsInfo.Open sSQL, gADOCon, adOpenDynamic, adLockReadOnly, adCmdText

  Do While Not rsInfo.EOF
    gObjGroups(psGroupName).Tables(rsInfo.Fields(0).Value).Changed = True
    
    FlagChildrenChanged gObjGroups(psGroupName).Tables(CStr(rsInfo.Fields(0).Value)), psGroupName
    
    rsInfo.MoveNext
  Loop
  rsInfo.Close
  Set rsInfo = Nothing
  
End Sub


Public Function SetComboText(cboCombo As ComboBox, sText As String) As Boolean

  Dim lCount As Long
  
  With cboCombo
    For lCount = 1 To .ListCount
      'AE20071212 Fault #12704
      'If .List(lCount - 1) = sText Then
      If LCase(.List(lCount - 1)) = LCase(sText) Then
        .ListIndex = lCount - 1
        SetComboText = True
        Exit For
      End If
    Next
  End With

End Function


Private Function ApplyChanges_LogoutCheck() As Boolean

  Dim objGroup As SecurityGroup
  Dim objUser As SecurityUser
  Dim strGroupsToLogOut As String
  Dim strUsersToLogOut As String
  Dim frmViewUsers As frmViewCurrentUsers
  Dim blnCancelled As Boolean
  Dim astrWindowsUsers() As String
  Dim iCount As Integer

  strUsersToLogOut = vbNullString
  For Each objGroup In gObjGroups
     
    'MH20060810 Fault 11415
    'If objGroup.RequireLogout And objGroup.Users_Initialised = False Then
    If objGroup.RequireLogout Or gbShiftSave Then
      If objGroup.Users_Initialised = False Then
        InitialiseUsersCollection objGroup
      End If
    End If

    For Each objUser In objGroup.Users

      'MH20060810 Fault 11415
      'If objGroup.RequireLogout Or objUser.DeleteUser Then
      If objGroup.RequireLogout Or objUser.DeleteUser Or gbShiftSave Then
        If objUser.LoginType = iUSERTYPE_TRUSTEDGROUP Then

'          ' AE20080422 Fault #13090
'          If glngSQLVersion = 8 Then
'            astrWindowsUsers = GetUsersInWindowsGroup(objUser.UserName)
'            For icount = LBound(astrWindowsUsers) To UBound(astrWindowsUsers)
'              'strUsersToLogOut = _
'                IIf(strUsersToLogOut <> vbNullString, strUsersToLogOut & ", ", vbNullString) & _
'                "'" & Replace(astrWindowsUsers(icount), "'", "''") & "'"
'              strUsersToLogOut = strUsersToLogOut & astrWindowsUsers(icount) & vbCrLf
'            Next icount
'
'          Else
'            strGroupsToLogOut = strGroupsToLogOut & objUser.UserName & ","
'          End If

          ' AE20080829 Fault #13358
          If InStr(strGroupsToLogOut, objGroup.Name & ",") = 0 Then
            strGroupsToLogOut = strGroupsToLogOut & objGroup.Name & ","
          End If

        Else
          'strUsersToLogOut = _
            IIf(strUsersToLogOut <> vbNullString, strUsersToLogOut & ", ", vbNullString) & _
            "'" & Replace(objUser.UserName, "'", "''") & "'"
            strUsersToLogOut = strUsersToLogOut & objUser.UserName & vbCrLf
        End If
      End If

    Next
      
  Next

'  ' AE20080422 Fault #13090
'  If (glngSQLVersion >= 9) And (strGroupsToLogOut <> vbNullString) Then
'    strUsersToLogOut = strUsersToLogOut & GetCurrentUsersInWindowsGroups(strGroupsToLogOut)
'  End If

  If (strGroupsToLogOut <> vbNullString) Then
    strUsersToLogOut = strUsersToLogOut & GetCurrentUsersInGroups(strGroupsToLogOut)
  End If

  If strUsersToLogOut <> vbNullString Then
  
    Set frmViewUsers = New frmViewCurrentUsers
    With frmViewUsers
  
      'strUsersToLogOut = "AND loginame IN (''" & _
        IIf(strUsersToLogOut <> vbNullString, "," & strUsersToLogOut, vbNullString) & ")"
  
      If .OkayToSave(strUsersToLogOut) = False Then
  
        MsgBox "You have made changes which will affect user(s) who are currently logged into the system." & vbNewLine & _
               "You will need to ensure that all users are logged out and that you have locked the system " & _
               "before you can apply these changes.", vbInformation, "Saving Changes"
  
        .Enabled = True
        .Saving = True
        .Show vbModal
      End If
  
      If .ForciblyDisconnect Then
        AuditAccess "Forcibly Disconnect Users", "Security"
      End If
  
      ApplyChanges_LogoutCheck = Not .Cancelled
  
    End With
    Unload frmViewUsers
    Set frmViewUsers = Nothing
  Else
    ApplyChanges_LogoutCheck = True
  End If


End Function

Public Sub ControlsDisableAll(objCurrent As Object, Optional blnEnabled As Boolean = False)

  Dim ctl As Control

  'Not all controls have a backcolor !
  On Local Error Resume Next

  If TypeOf objCurrent Is Form Then
    For Each ctl In objCurrent
      
      EnableControl ctl, blnEnabled
    Next

  Else
    EnableControl objCurrent, blnEnabled
    For Each ctl In objCurrent.Parent
      If ctl.Container.Name = objCurrent.Name Then
        EnableControl ctl, blnEnabled
      End If
    Next

  End If

End Sub

Private Function EnableControl(ctl As Control, blnEnabled As Boolean)

  ' JPD20020920 Added ActiveBar to the list of controls that are not enabled/disabled, as it has no
  ' 'enabled' property. Instead use the 'EnableActiveBar' method.
  
  'If Left(ctl.Name, 3) = "opt" Then
  '  Stop
  'End If
  
  
  If TypeOf ctl Is MSComctlLib.TabStrip Or _
         TypeOf ctl Is ComctlLib.TabStrip Or _
         TypeOf ctl Is ActiveBar Or _
         TypeOf ctl Is SSTab Then
    'Stop

  ElseIf TypeOf ctl Is Frame Or _
         TypeOf ctl Is PictureBox Or _
         TypeOf ctl Is MSComctlLib.ListView Or _
         TypeOf ctl Is ComctlLib.ListView Or _
         TypeOf ctl Is ListBox Or _
         TypeOf ctl Is Label Then
         
    'Just make container controls and scroll-able controls look disabled...
    '(NOTE: Code must be placed in drag-drop events etc. to disable it)
    ctl.ForeColor = IIf(blnEnabled, vbWindowText, vbApplicationWorkspace)
    If (TypeOf ctl Is ListBox) Or _
        (TypeOf ctl Is MSComctlLib.ListView) Or _
        (TypeOf ctl Is ComctlLib.ListView) Then
        
      ctl.BackColor = IIf(blnEnabled, vbWindowBackground, vbButtonFace)    'SSDBGrid
    End If

    'TM20010822 Fault 2566
    'Controls of this type need to be enabled so that scroll/click functionality
    'is allowed.
    ctl.Enabled = True

  ElseIf TypeOf ctl Is SSDBGrid Then

    ctl.Enabled = blnEnabled
    ctl.BackColorEven = IIf(blnEnabled, vbWindowBackground, vbButtonFace)   'SSDBGrid
    ctl.BackColorOdd = IIf(blnEnabled, vbWindowBackground, vbButtonFace)   'SSDBGrid

  ElseIf TypeOf ctl Is CommandButton Then 'Or _
         TypeOf ctl Is SSCommand Then
    'Disable all CommandButtons except cancel...

    If ctl.Cancel = False Then
      ctl.Enabled = blnEnabled
    Else
      ctl.Enabled = True
    End If

  'ElseIf (TypeOf ctl Is SSCheck) Or (TypeOf ctl Is CheckBox) Or (TypeOf ctl Is OptionButton) Then
  ElseIf (TypeOf ctl Is CheckBox) Or (TypeOf ctl Is OptionButton) Then

    ctl.ForeColor = IIf(blnEnabled, vbWindowText, vbApplicationWorkspace)
    ctl.BackColor = vbButtonFace
    ctl.Enabled = blnEnabled

  'ElseIf (TypeOf ctl Is UpDown) Then
  '  ctl.Enabled = blnEnabled

  ElseIf (TypeOf ctl Is TextBox) Then
    ctl.ForeColor = IIf(blnEnabled, vbWindowText, vbApplicationWorkspace)
    ctl.BackColor = IIf(blnEnabled, vbWindowBackground, vbButtonFace)
    ctl.Locked = Not blnEnabled
    ctl.TabStop = blnEnabled
    ctl.Enabled = blnEnabled
  
  Else
    ctl.ForeColor = IIf(blnEnabled, vbWindowText, vbApplicationWorkspace)
    ctl.BackColor = IIf(blnEnabled, vbWindowBackground, vbButtonFace)
    ctl.Enabled = blnEnabled
  
  End If

End Function

Private Function GetDBVersion() As String

  Dim rsInfo As New ADODB.Recordset
  
  GetDBVersion = GetSystemSetting("Database", "Version", vbNullString)

  If GetDBVersion = vbNullString Then
  
    rsInfo.Open "SELECT SystemManagerVersion FROM ASRSysConfig", gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
  
    If Not rsInfo.BOF And Not rsInfo.EOF Then
      GetDBVersion = rsInfo.Fields(0).Value
    End If
  
    rsInfo.Close
    Set rsInfo = Nothing
  
  End If

End Function

Private Function ApplyChanges_UpdatePersonnelRecords() As Boolean
  ' Apply the given usernames to the login field on personnel records
  On Error GoTo 0 'ErrorTrap
  
  Dim bOK As Boolean
  Dim strSQL As String
  Dim objGroup As SecurityGroup
  Dim objUser As SecurityUser
  Dim rsTempInfo As New ADODB.Recordset
  Dim strSQLUserName As String
  
  ' Only run this routine if a login column has been defined
  If glngLoginColumnID = 0 Then
    'MH20011128 Fault 3223
    'Need to set this to true BEFORE this exit function!
    ApplyChanges_UpdatePersonnelRecords = True
    Exit Function
  End If
  
  ' JDM - 18/12/01 - Fault 3197 - Flag that we are currently saving login fields to Personnel records
  rsTempInfo.Open "SELECT @@spid AS CurrentSPID", gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
  SaveSystemSetting "Database", "UpdateLoginColumnSPID", rsTempInfo!CurrentSPID
  
  bOK = True
  
  For Each objGroup In gObjGroups
 
    ' If the Users have been initialised then go through them.
    If objGroup.Users_Initialised Then
    
      For Each objUser In objGroup.Users
        With objUser
            
          strSQLUserName = Replace(.UserName, "'", "''")
           
          ' 1st Stage - Apply new user logins
          If .Personnel_RecordID > 0 Then
            strSQL = "UPDATE " & modExpression.GetTableName(glngPersonnelTableID) _
              & " SET " & modExpression.GetColumnName(glngLoginColumnID) & " = '" & strSQLUserName & "'" _
              & " WHERE ID = " & .Personnel_RecordID
            gADOCon.Execute strSQL, , adExecuteNoRecords
          End If
          
          '2nd stage - Remove login fields where the user has been deleted
          If .DeleteUser Then
            strSQL = "UPDATE " & modExpression.GetTableName(glngPersonnelTableID) _
              & " SET " & modExpression.GetColumnName(glngLoginColumnID) & " = ''" _
              & " WHERE " & modExpression.GetColumnName(glngLoginColumnID) & " = '" & strSQLUserName & "'"
            gADOCon.Execute strSQL, , adExecuteNoRecords
          End If
          
        End With
      Next
      Set objUser = Nothing
      
    End If
  Next

TidyUpAndExit:
  
  ' Remove bypass on personnel trigger bypass
  SaveSystemSetting "Database", "UpdateLoginColumnSPID", 0

  Set objGroup = Nothing
  Set objUser = Nothing
  ApplyChanges_UpdatePersonnelRecords = bOK
  Exit Function

ErrorTrap:
  bOK = False
  Resume TidyUpAndExit

End Function

Private Function ApplyChanges_ForceChangePasswords() As Boolean
  
  ' Force users to change passwords at next login
  On Error GoTo ErrorTrap
  
  Dim bOK As Boolean
  Dim sSQL As String
  Dim rsInfo As New ADODB.Recordset
  Dim objGroup As SecurityGroup
  Dim objUser As SecurityUser
  
  bOK = True

  For Each objGroup In gObjGroups
    ' If the Users have been initialised then go through them.
    If objGroup.Users_Initialised Then
      For Each objUser In objGroup.Users
        With objUser
          
          ' If user is forced to change password
          If .ForcePasswordChange And glngSQLVersion < 9 Then
            ' we should now force the user to change on next login
            rsInfo.Open "Select * From ASRSysPasswords WHERE Username = '" & Replace(LCase$(.UserName), "'", "''") & "'", gADOCon, adOpenForwardOnly, adLockOptimistic
          
            If rsInfo.BOF And rsInfo.EOF Then
              rsInfo.AddNew
            'Else
            '  rsInfo.Update
            End If
            
            rsInfo!UserName = LCase$(.UserName)
            'JPD 20041117 Fault 9484
            rsInfo!LastChanged = Replace(Format(Now, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/")
            rsInfo!ForceChange = 1
            rsInfo.Update
           
            rsInfo.Close
          End If
        End With
      Next
      Set objUser = Nothing
      
    End If
  Next objGroup

TidyUpAndExit:
  Set objGroup = Nothing
  Set objUser = Nothing
  Set rsInfo = Nothing
  ApplyChanges_ForceChangePasswords = bOK
  Exit Function

ErrorTrap:
  bOK = False
  Resume TidyUpAndExit

End Function

Public Function sysUserName(sysUID As Long) As String

  Dim rsUserInfo As New ADODB.Recordset
  Dim sSQL As String
  
  On Error GoTo ErrorTrap

  sSQL = "SELECT sysusers.name " & _
          " FROM sysusers " & _
          " WHERE sysusers.uid = " & sysUID
          
  rsUserInfo.Open sSQL, gADOCon, adOpenStatic, adLockReadOnly, adCmdText
  
  With rsUserInfo
    If .RecordCount > 0 Then
      sysUserName = !Name
    Else
      sysUserName = vbNullString
    End If
    .Close
  End With
  
  Exit Function
  
ErrorTrap:
  sysUserName = vbNullString
  
End Function

Public Function ResetPrintArray(iIndex As Integer, boolValue As Boolean)
    gasPrintOptions(iIndex).PrintLPaneGROUPS = boolValue
    gasPrintOptions(iIndex).PrintLPaneGROUP = boolValue
    gasPrintOptions(iIndex).PrintLPaneUSERS = boolValue
    gasPrintOptions(iIndex).PrintLPaneTABLESVIEWS = boolValue
    gasPrintOptions(iIndex).PrintLPaneSYSTEM = boolValue
    gasPrintOptions(iIndex).PrintLPaneTABLE = boolValue
    '
    gasPrintOptions(iIndex).PrintRPaneGROUPS = boolValue
    gasPrintOptions(iIndex).PrintRPaneGROUP = boolValue
    gasPrintOptions(iIndex).PrintRPaneUSERS = boolValue
    gasPrintOptions(iIndex).PrintRPaneTABLESVIEWS = boolValue
    gasPrintOptions(iIndex).PrintRPaneSYSTEM = boolValue
    gasPrintOptions(iIndex).PrintRPaneTABLE = boolValue
End Function
    
Public Function CheckPlatform() As Boolean

  Dim sSQL As String, sMsg As String
  Dim lngSQLVersion As Double   'Changed from Long for SQL2008 R2
  
  Dim strLastSQLServerVersion As String
  Dim strLastDatabaseName As String
  Dim strLastServerName As String
  
  Dim rsSQLInfo As ADODB.Recordset
  
  ' Get the SQL Server version number.
  lngSQLVersion = 0
  sSQL = "master..xp_msver ProductVersion"
  Set rsSQLInfo = New ADODB.Recordset
  rsSQLInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsSQLInfo
    If Not (.BOF And .EOF) Then
      lngSQLVersion = Val(.Fields("character_value").Value)
    End If
    .Close
  End With
  Set rsSQLInfo = Nothing

  sMsg = vbNullString
  strLastSQLServerVersion = GetSystemSetting("Platform", "SQLServerVersion", 0)
  strLastDatabaseName = UCase$(GetSystemSetting("Platform", "DatabaseName", ""))
  strLastServerName = UCase$(GetSystemSetting("Platform", "ServerName", ""))
  If strLastServerName = "." Then strLastServerName = UCase$(UI.GetHostName)

  If GetServerName <> GetOldServerName Then
    sMsg = "The Microsoft SQL Server has been renamed but the operation is incomplete."
    
    MsgBox sMsg & vbCrLf & _
      "Please contact your System Administrator.", _
      vbOKOnly + vbExclamation, Application.Name
    
    CheckPlatform = False
    
    Exit Function
  Else
    If Val(strLastSQLServerVersion) <> lngSQLVersion Then
      sMsg = "The Microsoft SQL Version has been upgraded."
    ElseIf strLastServerName <> GetServerName() Then
      sMsg = "The database has moved to a different Microsoft SQL Server."
    ElseIf strLastDatabaseName <> GetDBName() Then
      sMsg = "The database name has changed."
  '  ElseIf Not FrameworkVersionOK Then
  '    sMsg = "The Microsoft .NET Framework version has changed on the server."
    End If
    
    If sMsg <> vbNullString Then
      MsgBox sMsg & vbCrLf & _
            "Please ask the System Administrator to update the database in the System Manager.", _
            vbOKOnly + vbExclamation, Application.Name
    
      CheckPlatform = False
      
      Exit Function
    End If
  End If

  CheckPlatform = True
    
End Function

Private Function EnableUDFFunctions() As Boolean

  Dim sSQL As String
  Dim rsUser As ADODB.Recordset
  
  sSQL = "exec master..xp_msver"
  
  Set rsUser = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  rsUser.MoveNext
  
  Select Case Val(rsUser(3))
    Case Is >= 8
      EnableUDFFunctions = True
    Case Else
      EnableUDFFunctions = False
  End Select
  
  rsUser.Close
  Set rsUser = Nothing
    
End Function


Public Function PrintPictureBox(Xcoordinate As Integer, Ycoordinate As Integer, fParsedBoolean)
Dim mlngBottom As Long

'NHRD19012006 Fault 10733 Replaced code below with this
If gasPrintOptions(1).PrintBlankVersion Then
  'If we are printing a blank version of the form then default to an empty check box.
  Printer.PaintPicture frmGroupMaint1.PicBlankCheckBox.Image, Xcoordinate, Ycoordinate
  Exit Function
End If

Select Case fParsedBoolean
    Case 0: Printer.PaintPicture frmGroupMaint1.PicBlankCheckBox.Image, Xcoordinate, Ycoordinate
    Case 1: Printer.PaintPicture frmGroupMaint1.PicTickedCheckBox.Image, Xcoordinate, Ycoordinate
    Case 2: Printer.PaintPicture frmGroupMaint1.PicVeryGreyCheckBox.Image, Xcoordinate, Ycoordinate
    Case True: Printer.PaintPicture frmGroupMaint1.PicTickedCheckBox.Image, Xcoordinate, Ycoordinate
    Case False: Printer.PaintPicture frmGroupMaint1.PicBlankCheckBox.Image, Xcoordinate, Ycoordinate
End Select

'MH20050812 Fault 10267
'    If gasPrintOptions(1).PrintBlankVersion Then fParsedBoolean = False
'    If fParsedBoolean Then
'        Printer.PaintPicture frmGroupMaint1.PicTickedCheckBox.Image, Xcoordinate, Ycoordinate
'    Else
'        Printer.PaintPicture frmGroupMaint1.PicBlankCheckBox.Image, Xcoordinate, Ycoordinate
'    End If
'    If gasPrintOptions(1).PrintBlankVersion Or Not fParsedBoolean Then
'        Printer.PaintPicture frmGroupMaint1.PicGreyCheckbox.Image, Xcoordinate, Ycoordinate
'    Else
'        Printer.PaintPicture frmGroupMaint1.PicTickedCheckBox.Image, Xcoordinate, Ycoordinate
'    End If
End Function

Private Function ApplyChanges_ViewMenuPermissions() As Boolean
  ' Apply the System permissions in the SQL Server database.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim sSQL As String
  Dim objGroup As SecurityGroup

  fOK = True

  For Each objGroup In gObjGroups
    With objGroup

      ' If the View menu permissions have been initialised then go through them.
      If .Initialised Then

        ' Delete all View menu Permission records for this user group.
        sSQL = "DELETE FROM ASRSysViewMenuPermissions" & _
          " WHERE groupName = '" & .Name & "'"
        gADOCon.Execute sSQL, , adExecuteNoRecords
        
        For iLoop = 1 To .Tables.Count

          ' Creating the View menu Permissions records for the current user group.
          sSQL = "INSERT INTO ASRSysViewMenuPermissions" & _
            " (GroupName, TableName, TableID, HideFromMenu)" & _
            " VALUES('" & .Name & "','" & .Tables.Item(iLoop).Name & "'," & .Tables.Item(iLoop).TableID & ", " _
            & IIf(.Tables.Item(iLoop).HideFromMenu = True, "1", "0") & ")"
          gADOCon.Execute sSQL, , adExecuteNoRecords
  
        Next iLoop
      End If
    End With
  Next objGroup

TidyUpAndExit:
  Set objGroup = Nothing
  ApplyChanges_ViewMenuPermissions = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Private Function ApplyChanges_TidyUpOrphans() As Boolean

  On Error GoTo ErrorTrap
  Dim bOK As Boolean
  Dim sSQL As String
  Dim iCount As Integer
  Dim strController As String
  Dim strDomain As String
  Dim astrDomains() As String
  Dim bFound As Boolean
   
  bOK = True

  If gbUserCanManageLogins And gbDeleteOrphanWindowsLogins Then
  
    astrDomains = GetWindowsDomains
  
    ' Go through domains individually in case there are network issues
    For iCount = LBound(astrDomains) + 1 To UBound(astrDomains)
    
      If astrDomains(iCount) = gobjNET.ComputerName Then
        bFound = True
      Else
        strController = gobjNET.GetPrimaryDomainController(astrDomains(iCount))
        bFound = (LenB(strController) <> 0)
      End If
      
      If bFound Then
        sSQL = "EXECUTE spASRDeleteInvalidLogins '" & astrDomains(iCount) & "'"
        gADOCon.Execute sSQL, , adExecuteNoRecords
      End If
      
    Next iCount
  End If

TidyUpAndExit:
  ApplyChanges_TidyUpOrphans = bOK
  Exit Function

ErrorTrap:
  bOK = False
  Resume TidyUpAndExit


End Function

Private Function ApplyChanges_ApplyRolesToLogins() As Boolean
  ' Create the new User Logins in the SQL Server database.
  On Error GoTo ErrorTrap
  
  Dim bIsSysManager As Boolean
  Dim fOK As Boolean
  Dim sSQL As String
  Dim objGroup As SecurityGroup
  
  fOK = True

  For Each objGroup In gObjGroups
    
    If objGroup.Initialised Then
      
      ' System metadata role
      If objGroup.SystemPermissions.Item("P_MODULEACCESS_SYSTEMMANAGER").Allowed Then
        sSQL = "EXEC sp_addrolemember @rolename = [ASRSysAdmin], @membername = [" & objGroup.Name & "]"
      Else
        sSQL = "EXEC sp_droprolemember @rolename = [ASRSysAdmin], @membername = [" & objGroup.Name & "]"
      End If
      gADOCon.Execute sSQL, , adExecuteNoRecords
      
      ' Workflow administer/view log role
      If objGroup.SystemPermissions.Item("P_WORKFLOW_ADMINISTER").Allowed Or objGroup.SystemPermissions.Item("P_WORKFLOW_VIEWLOG").Allowed Then
        sSQL = "EXEC sp_addrolemember @rolename = [ASRSysWorkflowAdmin], @membername = [" & objGroup.Name & "]"
      Else
        sSQL = "EXEC sp_droprolemember @rolename = [ASRSysWorkflowAdmin], @membername = [" & objGroup.Name & "]"
      End If
      gADOCon.Execute sSQL, , adExecuteNoRecords
               
    End If
    
  Next

TidyUpAndExit:
  ApplyChanges_ApplyRolesToLogins = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Private Function MarkAllGroupsAsChanged() As Boolean
  
  Dim iCount As Integer
  Dim sGroupName As String
  Dim objGroup As SecurityGroup
  Dim bOK As Boolean
  
  On Error GoTo ErrorTrap
  
  OutputCurrentProcess2 "", gObjGroups.Count
  
  For Each objGroup In gObjGroups
    sGroupName = objGroup.Name
    
    OutputCurrentProcess2 sGroupName
    gobjProgress.UpdateProgress2 False
    
    If Not gObjGroups(sGroupName).Initialised Then
      bOK = InitialiseGroup(gObjGroups(sGroupName), False)
    End If
  
    objGroup.Changed = True
  
  Next objGroup
  
TidyUpAndExit:
  MarkAllGroupsAsChanged = bOK
  Exit Function

ErrorTrap:
  bOK = False
  GoTo TidyUpAndExit

End Function


Private Function ApplyPostSaveProcessing() As Boolean

  On Error GoTo ErrorTrap

  Dim cmdPostProcess As New ADODB.Command
  Dim bOK As Boolean

  bOK = True

  With cmdPostProcess
    .CommandText = "spASRPostSystemSave"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon

    .Execute
  End With

  Set cmdPostProcess = Nothing

TidyUpAndExit:
  ApplyPostSaveProcessing = bOK
  Exit Function

ErrorTrap:
  bOK = False
  GoTo TidyUpAndExit

End Function

Public Function GetOldServerName() As String

  Dim sSQL As String
  Dim rsSQLInfo As ADODB.Recordset

  sSQL = "SELECT @@SERVERNAME"

  Set rsSQLInfo = New ADODB.Recordset
  rsSQLInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly

  With rsSQLInfo
    If Not (.BOF And .EOF) Then
      GetOldServerName = UCase$(.Fields(0).Value)
    End If
    .Close
  End With

  Set rsSQLInfo = Nothing
  
End Function


Public Function GetServerName() As String

  Dim sSQL As String
  Dim rsSQLInfo As ADODB.Recordset

  ' AJE20090114 Fault #13490
  'sSQL = "SELECT @@SERVERNAME"
  sSQL = "SELECT SERVERPROPERTY('servername') "
  Set rsSQLInfo = New ADODB.Recordset
  rsSQLInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly

  With rsSQLInfo
    If Not (.BOF And .EOF) Then
      GetServerName = UCase$(.Fields(0).Value)
    End If
    .Close
  End With

  Set rsSQLInfo = Nothing
  
End Function


Public Function GetDBName()

  Dim sSQL As String
  Dim rsSQLInfo As ADODB.Recordset
 
  sSQL = "SELECT DB_NAME()"
  
  Set rsSQLInfo = New ADODB.Recordset
  rsSQLInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsSQLInfo
    If Not (.BOF And .EOF) Then
      GetDBName = UCase$(.Fields(0).Value)
    End If
    .Close
  End With
  Set rsSQLInfo = Nothing

End Function

Public Function IsModuleEnabled(lngModuleCode As enum_Module) As Boolean
  IsModuleEnabled = (gobjLicence.Modules And lngModuleCode)
End Function

Public Function ApplyChanges_LoginMaintenance() As Boolean

On Error GoTo ErrorTrap:

  Dim bOK As Boolean
  Dim sSQL As String
  Dim sURL As String
  Dim sSelfServiceColumn As String
  Dim sLeavingDateColumn As String
  Dim sSecurityGroupColumn As String
  Dim sPersonnelTable As String
    
  bOK = True
  sURL = GetSystemSetting("Web", "SiteAddress", "")
  sPersonnelTable = GetTableName(glngPersonnelTableID)
  sSelfServiceColumn = GetColumnName(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_LOGINNAME))
  sLeavingDateColumn = GetColumnName(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_LEAVINGDATE))
  sSecurityGroupColumn = GetColumnName(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_SECURITYGROUP))
 
  sSQL = "/* --------------------------------------------------- */" & vbNewLine & _
        "/* Login Maintenance stored procedure.                 */" & vbNewLine & _
        "/* Automatically generated by the Security Manager.    */" & vbNewLine & _
        "/* --------------------------------------------------- */" & vbNewLine & _
        "ALTER PROCEDURE spASRGenerateSelfServiceLogins(@logins AS SelfServiceType READONLY)" & vbNewLine & _
        "WITH EXECUTE AS SELF" & vbNewLine & _
        "AS" & vbNewLine & _
        "BEGIN" & vbNewLine & vbNewLine & _
        "  SET NOCOUNT ON;" & vbNewLine & vbNewLine & _
        "  DECLARE @hResult int," & vbNewLine & _
        "      @sCode nvarchar(MAX) = ''," & vbNewLine & _
        "      @login nvarchar(255)," & vbNewLine & _
        "      @securityGroup nvarchar(255)," & vbNewLine & _
        "      @emailAddress nvarchar(255)," & vbNewLine & _
        "      @startDate datetime," & vbNewLine & _
        "      @leavingDate datetime," & vbNewLine & _
        "      @todaysDate datetime = DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE()))," & vbNewLine & _
        "      @knownAs nvarchar(255)," & vbNewLine & _
        "      @sendEmail bit = " & IIf(gbLoginMaintSendEmail, 1, 0) & "," & vbNewLine & _
        "      @emailContent nvarchar(MAX) = ''," & vbNewLine & _
        "      @initialPassword varchar(12) = 'Password!123';" & vbNewLine & vbNewLine

  If gbLoginMaintAutoAdd Then
  
    sSQL = sSQL & _
          "  DECLARE loginCursor CURSOR LOCAL FAST_FORWARD FOR SELECT [Login], [Email], [StartDate], [LeavingDate], [KnownAs], [SecurityGroup] FROM @logins" & vbNewLine & _
          "      WHERE [Login] <> '' AND [Email] <> '';" & vbNewLine & _
          "  OPEN loginCursor;" & vbNewLine & _
          "  FETCH NEXT FROM loginCursor INTO @login, @emailAddress, @startDate, @leavingDate, @knownAs, @securityGroup;" & vbNewLine & _
          "  WHILE @@FETCH_STATUS = 0" & vbNewLine & _
          "  BEGIN" & vbNewLine & vbNewLine
  
    If gbOverrideSecurityGroup Then
      sSQL = sSQL & _
          "      SET @securityGroup = '" & gstrLoginMaintAutoAddGroup & "';" & vbNewLine & vbNewLine
    End If
  
    sSQL = sSQL & _
          "      IF NOT EXISTS (SELECT loginname FROM sys.syslogins WHERE name = @login) AND EXISTS(SELECT Name FROM ASRSysGroups WHERE Name = @securityGroup)" & vbNewLine & _
          "      BEGIN" & vbNewLine & _
          "          IF @sendEmail = 1" & vbNewLine & _
          "              SET @initialPassword = LOWER(SUBSTRING(CONVERT(varchar(255), NEWID()), 1, 5)) + 'a!3' + LOWER(SUBSTRING(CONVERT(varchar(255), NEWID()), 1, 3));" & vbNewLine & vbNewLine & _
          "          SET @emailContent = 'Dear ' + @knownAs + ',' + CHAR(13) + 'Your login details for OpenHR Self-service are as follows:' + CHAR(13) + CHAR(13) +" & vbNewLine & _
          "              'Username : ' + @login + CHAR(13) +" & vbNewLine & _
          "              'Password : ' + @initialPassword + CHAR(13) + CHAR(13) +" & vbNewLine & _
          "              'Please follow the link provided below to access the website.' + CHAR(13) + CHAR(13) +" & vbNewLine & _
          "              '<" & sURL & ">' + CHAR(13) + CHAR(13) +" & vbNewLine & _
          "              'Note that you will be prompted to change your password the first time you login.'" & vbNewLine & vbNewLine & _
          "          SET @sCode = 'CREATE LOGIN [' + @login + '] WITH PASSWORD = ''' + @initialPassword + ''' MUST_CHANGE, CHECK_POLICY=ON, CHECK_EXPIRATION=ON; " & vbNewLine & _
          "              CREATE USER [' + @login + '] FOR LOGIN [' + @login + '];" & vbNewLine & _
          "              EXEC sp_addrolemember ''ASRSysGroup'', ''' + @login + ''';" & vbNewLine & _
          "              EXEC sp_addrolemember ''' + @securityGroup + ''', ''' + @login + '''';" & vbNewLine & vbNewLine & _
          "          EXEC sp_executeSQL @sCode;" & vbNewLine & vbNewLine
          
    sSQL = sSQL & _
          "        INSERT ASRSysAuditGroup(UserName, DateTimeStamp, GroupName, UserLogin, [Action])" & vbNewLine & _
          "            VALUES ('System', GETDATE(), @securityGroup, @login, 'User Added')" & vbNewLine & vbNewLine
          
    sSQL = sSQL & _
          "          IF @sendEmail = 1 AND (@leavingDate >= @todaysDate OR @leavingDate IS NULL)" & vbNewLine & _
          "          BEGIN" & vbNewLine & _
          "              IF @startDate < @todaysDate SET @startDate = @todaysDate;" & vbNewLine & _
          "              INSERT ASRSysEmailQueue(RecordDesc, ColumnValue, DateDue, UserName, [Immediate], RecalculateRecordDesc," & vbNewLine & _
          "                  RepTo, MsgText, WorkflowInstanceID, [Subject])" & vbNewLine & _
          "              VALUES ('', '', @startDate, 'OpenHR Web', 1, 0," & vbNewLine & _
          "                  @emailAddress, @emailContent, 0, 'Your Self-service login details');" & vbNewLine & vbNewLine & _
          "              EXECUTE dbo.spASREmailImmediate 'OpenHR Web';" & vbNewLine & _
          "          END" & vbNewLine & _
          "      END" & vbNewLine & _
          "      FETCH NEXT FROM loginCursor INTO @login, @emailAddress, @startDate, @leavingDate, @knownAs, @securityGroup;" & vbNewLine
    
      sSQL = sSQL & _
          "  END" & vbNewLine & _
          "CLOSE loginCursor;" & vbNewLine & _
          "DEALLOCATE loginCursor;"
  
  End If

  sSQL = sSQL & _
    "END" & vbNewLine

  gADOCon.Execute sSQL, , adExecuteNoRecords


  ' ----------------
  ' Removal of login stored procedure
  ' ----------------
  sSQL = "/* ---------------------------------------------------- */" & vbNewLine & _
        "/* Login Maintenance stored procedure.                 */" & vbNewLine & _
        "/* Automatically generated by the Security Manager.    */" & vbNewLine & _
        "/* --------------------------------------------------- */" & vbNewLine & _
        "ALTER PROCEDURE dbo.spASRDeleteExpiredSelfServiceLogins" & vbNewLine & _
        "WITH EXECUTE AS SELF" & vbNewLine & vbNewLine & _
        "AS" & vbNewLine & _
        "BEGIN" & vbNewLine & vbNewLine & _
        "    SET NOCOUNT ON;" & vbNewLine & vbNewLine & _
        "    DECLARE @yesterday datetime = DATEADD(dd, 0, DATEDIFF(dd, 0,  GETDATE())) - 1 ," & vbNewLine & _
        "        @login nvarchar(MAX)," & vbNewLine & _
        "        @sqlCode nvarchar(MAX) = '';" & vbNewLine & vbNewLine

  If gbLoginMaintDisableOnLeave Then
  
    sSQL = sSQL & _
        "    DECLARE loginCursor CURSOR LOCAL FAST_FORWARD FOR" & vbNewLine & _
        "        SELECT [" & sSelfServiceColumn & "] FROM [" & sPersonnelTable & "]" & vbNewLine & _
        "        WHERE [" & sLeavingDateColumn & "] <= @yesterday;" & vbNewLine & vbNewLine & _
        "    OPEN loginCursor;" & vbNewLine & _
        "    FETCH NEXT FROM loginCursor INTO @login;" & vbNewLine & _
        "    WHILE @@FETCH_STATUS = 0" & vbNewLine & _
        "    BEGIN" & vbNewLine & _
        "        IF EXISTS (SELECT * FROM sys.sysusers WHERE name = @login)" & vbNewLine & _
        "            EXECUTE ('DROP USER [' + @login + ']');" & vbNewLine & vbNewLine & _
        "        IF EXISTS (SELECT * FROM sys.syslogins WHERE name = @login)" & vbNewLine & _
        "        BEGIN" & vbNewLine & _
        "            EXECUTE ('DROP LOGIN [' + @login + ']');" & vbNewLine & _
        "            INSERT dbo.ASRSysAuditGroup(UserName, DateTimeStamp, GroupName, UserLogin, [Action])" & vbNewLine & _
        "                VALUES ('System', @yesterday + 1, '" & gstrLoginMaintAutoAddGroup & "', @login, 'User Deleted')" & vbNewLine & _
        "        END" & vbNewLine & _
        "        FETCH NEXT FROM loginCursor INTO @login;" & vbNewLine & _
        "   END" & vbNewLine & vbNewLine & _
        "CLOSE loginCursor;" & vbNewLine & _
        "DEALLOCATE loginCursor;" & vbNewLine
    
  End If

  sSQL = sSQL & _
    "END" & vbNewLine

  gADOCon.Execute sSQL, , adExecuteNoRecords


TidyUpAndExit:
  ApplyChanges_LoginMaintenance = bOK
  Exit Function

ErrorTrap:
  MsgBox ("Error Generating ApplyChanges_LoginMaintenance")
  bOK = False
  Resume TidyUpAndExit

End Function
