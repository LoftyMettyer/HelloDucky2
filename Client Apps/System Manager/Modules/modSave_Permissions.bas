Attribute VB_Name = "modSave_Permissions"
Option Explicit

Private gObjGroups As clsSecurityGroups


Public Function ReadPermissions(ByRef psErrMsg As String) As Boolean
  ' Create a collection of user groups and their table/view/column permissions
  ' Return TRUE if everything went okay.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fSystemManager As Boolean
  Dim fSecurityManager As Boolean
  Dim sSQL As String
  Dim sGroupName As String
  Dim rsGroups As New ADODB.Recordset
  Dim rsPermissions As New ADODB.Recordset
  Dim objGroup As clsSecurityGroup
  Dim lngAmountOfGroups As Long
  Dim objPerformance As SystemMgr.clsPerformance

  fOK = True
  Set gObjGroups = Nothing
  Set gObjGroups = New clsSecurityGroups
   
  Set objPerformance = New SystemMgr.clsPerformance
  objPerformance.ClearLogFile
   
  'MH20040112 Fault 5627
  'sSQL = "exec sp_ASRGetUserGroups"
  sSQL = "SELECT name FROM sysusers " & _
         "WHERE gid = uid AND gid > 0 " & _
         "AND not (name like 'ASRSys%') AND not (name like 'db[_]%')"
  
  ' Get the amount of records first
  rsGroups.Open sSQL, gADOCon, adOpenKeyset, adLockReadOnly
  lngAmountOfGroups = rsGroups.RecordCount
  rsGroups.Close
  
  OutputCurrentProcess2 vbNullString, lngAmountOfGroups + 1
  
  rsGroups.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsGroups
       
    If Not .EOF And Not .BOF Then
      While fOK And (Not .EOF)
        
        sGroupName = Trim(.Fields(0).value) 'Trim(!Name)
      
        OutputCurrentProcess2 sGroupName
        gobjProgress.UpdateProgress2
      
        ' Add the group to the groups collection
        Set objGroup = gObjGroups.Add(sGroupName)
        
        ' Check if the group is permitted use of the System or Security managers.
        fSystemManager = False
        fSecurityManager = False
        sSQL = "SELECT ASRSysGroupPermissions.permitted, ASRSysPermissionItems.itemKey" & _
          " FROM ASRSysPermissionItems" & _
          " INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
          " INNER JOIN ASRSysGroupPermissions ON ASRSysPermissionItems.itemID = ASRSysGroupPermissions.itemID" & _
          " WHERE (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER' OR ASRSysPermissionItems.itemkey = 'SECURITYMANAGER')" & _
          " AND ASRSysGroupPermissions.groupName = '" & sGroupName & "'"
        rsPermissions.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
        
        Do While (Not rsPermissions.EOF)
          
          If rsPermissions!permitted Then
            If rsPermissions!ItemKey = "SYSTEMMANAGER" Then
              fSystemManager = True
            Else
              fSecurityManager = True
            End If
          End If
          
          rsPermissions.MoveNext
        Loop
        rsPermissions.Close
        
        objGroup.SecurityManager = fSecurityManager
        objGroup.SystemManager = fSystemManager
        
        ' Initialise the user views collection.
        objPerformance.StartClock sGroupName
        fOK = SetupTablesCollection(objGroup)
        objPerformance.LogSummary
        
        fOK = fOK And Not gobjProgress.Cancelled

        .MoveNext

      Wend
    End If
  
    .Close
  End With
    
TidyUpAndExit:
  'If (Not fOK) And (Len(psErrMsg) = 0) Then
  If (Not fOK) And Not gobjProgress.Cancelled And _
    (Len(psErrMsg) = 0) Then
    OutputError "Error reading table permissions."
  End If
  Set rsPermissions = Nothing
  Set rsGroups = Nothing
  ReadPermissions = fOK
  Exit Function
  
ErrorTrap:
  'psErrMsg = "Error reading table/view permissions." & vbCr & vbCr & _
    Err.Description
  OutputError "Error reading table/view permissions."
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function SetupTablesCollection(pobjGroup As clsSecurityGroup) As Boolean
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
  Dim sRealSourceList As SystemMgr.cStringBuilder
  Dim sTableViewName As String
  Dim rsInfo1 As ADODB.Recordset
  Dim rsInfo2 As ADODB.Recordset
  Dim rsInfo3 As ADODB.Recordset
  Dim rsInfo4 As ADODB.Recordset
  Dim rsTables As ADODB.Recordset
  Dim rsViews As ADODB.Recordset
  Dim rsPermissions As ADODB.Recordset
  Dim objColumn As clsSecurityColumn
  Dim objColumns As clsSecurityColumns
  Dim objTableView As clsSecurityTable
  Dim avChildViews() As Variant
  Dim iTemp As Integer
  Dim fChildView As Boolean
  Dim sPermissionName As String
  Dim strObjectName As String
  Dim strColumnName As String
  Dim iAction As Integer
  Dim iSelect As Integer
  Dim iUpdate As Integer
  Dim iTableType As Integer

  Set sRealSourceList = New SystemMgr.cStringBuilder
  Set rsInfo1 = New ADODB.Recordset
  Set rsInfo2 = New ADODB.Recordset
  Set rsInfo3 = New ADODB.Recordset
  Set rsInfo4 = New ADODB.Recordset
  Set rsTables = New ADODB.Recordset
  Set rsViews = New ADODB.Recordset
  Set rsPermissions = New ADODB.Recordset

  fOK = True
  fSysSecManager = (pobjGroup.SecurityManager Or pobjGroup.SystemManager)
    
  ' Create an array of child view IDs and their associated table names.
  ' Column 1 - child view ID
  ' Column 2 - associated table name
  ' Column 3 - 0=OR, 1=AND
  sSQL = "SELECT ASRSysChildViews2.childViewID, ASRSysTables.tableName, ASRSysChildViews2.type" & _
    " FROM ASRSysChildViews2" & _
    " INNER JOIN ASRSysTables ON ASRSysChildViews2.tableID = ASRSysTables.tableID" & _
    " WHERE ASRSysChildViews2.role = '" & pobjGroup.Name & "'"
  rsInfo1.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText

  ReDim avChildViews(3, 100)
  lngNextIndex = -1
  If Not rsInfo1.EOF Then
    Do While Not rsInfo1.EOF
      lngNextIndex = lngNextIndex + 1
      If lngNextIndex > UBound(avChildViews, 2) Then ReDim Preserve avChildViews(3, lngNextIndex + 100)
      avChildViews(1, lngNextIndex) = rsInfo1(0).value
      avChildViews(2, lngNextIndex) = rsInfo1(1).value
      avChildViews(3, lngNextIndex) = IIf(IsNull(rsInfo1(2).value), 0, rsInfo1(2).value)
      rsInfo1.MoveNext
    Loop
    ReDim Preserve avChildViews(3, lngNextIndex)
  End If
  rsInfo1.Close
  
  ' Get the collection with items for each TABLE in the system.
  sSQL = "SELECT tableID, tableName, tableType FROM ASRSysTables"
  rsTables.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
  
  Do While Not rsTables.EOF
    Set objColumns = New clsSecurityColumns
    pobjGroup.Tables.Add objColumns, rsTables(1).value, rsTables(2).value
    Set objColumns = Nothing
    
    With pobjGroup.Tables(rsTables(1).value)
      .SelectPrivilege = IIf(fSysSecManager, giPRIVILEGES_ALLGRANTED, IIf(rsTables(2).value = iTabLookup, giPRIVILEGES_ALLGRANTED, giPRIVILEGES_NONEGRANTED))
      .UpdatePrivilege = IIf(fSysSecManager, giPRIVILEGES_ALLGRANTED, giPRIVILEGES_NONEGRANTED)
      .InsertPrivilege = IIf(fSysSecManager, True, False)
      .DeletePrivilege = IIf(fSysSecManager, True, False)
      .ParentJoinType = 0
    End With
    rsTables.MoveNext
  Loop
  rsTables.Close


  ' Initialise the collection with items for each VIEW in the system.
  sSQL = "SELECT ASRSysViews.viewName FROM ASRSysViews"
  rsViews.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
  Do While Not rsViews.EOF
    Set objColumns = New clsSecurityColumns
    
    sTableViewName = rsViews(0).value  ' ViewName
    pobjGroup.Views.Add objColumns, sTableViewName, 0
    Set objColumns = Nothing

    With pobjGroup.Views(sTableViewName)
      .SelectPrivilege = IIf(fSysSecManager, giPRIVILEGES_ALLGRANTED, giPRIVILEGES_NONEGRANTED)
      .UpdatePrivilege = IIf(fSysSecManager, giPRIVILEGES_ALLGRANTED, giPRIVILEGES_NONEGRANTED)
      .InsertPrivilege = IIf(fSysSecManager, True, False)
      .DeletePrivilege = IIf(fSysSecManager, True, False)
      .ParentJoinType = 0
    End With
   rsViews.MoveNext
  Loop
  rsViews.Close


  ' Get the permissions for each table or view.
  sRealSourceList.TheString = vbNullString
  sLastRealSource = vbNullString
   
  If Not fSysSecManager Then
    ' If the user is NOT a 'system manager' or 'security manager'
    ' read the table permissions from the server.
    sSQL = "exec sp_ASRAllTablePermissionsForGroup '" & pobjGroup.Name & "'"
    rsPermissions.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Do While Not rsPermissions.EOF
      Set objTableView = Nothing

      sPermissionName = UCase(rsPermissions.Fields(0).value)   ' Name
      iAction = rsPermissions.Fields(1).value                   ' Action

      If sLastRealSource <> sPermissionName Then
        sRealSourceList.Append IIf(sRealSourceList.Length <> 0, ", '", "'") & sPermissionName & "'"
        sLastRealSource = sPermissionName
      End If
      
      If (iAction = 195) Or (iAction = 196) Then
        fChildView = False

        If Left$(sPermissionName, 8) = "ASRSYSCV" Then
          fChildView = True
          ' Determine which table the child view is for.
          iTemp = InStr(sPermissionName, "#")
          lngChildViewID = val(Mid$(sPermissionName, 9, iTemp - 9))
        End If
        
        If fChildView Then
          For lngNextIndex = 0 To UBound(avChildViews, 2)
            If avChildViews(1, lngNextIndex) = lngChildViewID Then
              Set objTableView = pobjGroup.Tables(avChildViews(2, lngNextIndex))
              objTableView.ParentJoinType = avChildViews(3, lngNextIndex)
              Exit For
            End If
          Next lngNextIndex
        Else
          If pobjGroup.Tables.IsValid(sPermissionName) Then
            Set objTableView = pobjGroup.Tables(sPermissionName)
          Else
            Set objTableView = pobjGroup.Views(sPermissionName)
          End If
        End If
  
        If Not objTableView Is Nothing Then
          Select Case iAction
            Case 195 ' Insert permission.
              objTableView.InsertPrivilege = True
            Case 196 ' Delete permission.
              objTableView.DeletePrivilege = True
          End Select
        End If
      End If
      
      rsPermissions.MoveNext
    Loop
    rsPermissions.Close
  End If
      
  ' Get the list of all columns in all tables/views.
  sSQL = "SELECT ASRSysColumns.columnName," & _
    " ASRSysColumns.columnID," & _
    " ASRSysTables.tableName AS tableViewName," & _
    " ASRSysTables.tableType AS tableType" & _
    " FROM ASRSysColumns" & _
    " INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID" & _
    " WHERE ASRSysColumns.columnType <> " & Trim$(Str$(giCOLUMNTYPE_SYSTEM)) & _
    " AND ASRSysColumns.columnType <> " & Trim$(Str$(giCOLUMNTYPE_LINK)) & _
    " UNION" & _
    " SELECT ASRSysColumns.columnName," & _
    " ASRSysColumns.columnID," & _
    " ASRSysViews.viewName AS tableViewName," & _
    " 0 AS tableType" & _
    " FROM ASRSysColumns" & _
    " INNER JOIN ASRSysViews ON ASRSysColumns.tableID = ASRSysViews.viewTableID" & _
    " INNER JOIN ASRSysViewColumns ON (ASRSysColumns.columnID = ASRSysViewColumns.columnID AND ASRSysViews.viewID = ASRSysViewColumns.viewID)" & _
    " WHERE ASRSysViewColumns.inView = 1" & _
    " AND ASRSysColumns.columnType <> " & Trim$(Str$(giCOLUMNTYPE_SYSTEM)) & _
    " AND ASRSysColumns.columnType <> " & Trim$(Str$(giCOLUMNTYPE_LINK))
  rsInfo2.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText

  Do While Not rsInfo2.EOF
    
    strColumnName = rsInfo2.Fields(0).value
    iTableType = rsInfo2.Fields(3).value

    If iTableType <> iTabView Then
      Set objColumns = pobjGroup.Tables(rsInfo2.Fields(2).value).Columns
    Else
      Set objColumns = pobjGroup.Views(rsInfo2.Fields(2).value).Columns
    End If
    
    ' Add the column object to the collection.
    Set objColumn = objColumns.Add(UCase$(Trim$(strColumnName)))
    
    ' Set the security column properties
    objColumn.Name = strColumnName
    objColumn.ColumnID = rsInfo2.Fields(1).value
    objColumn.SelectPrivilege = fSysSecManager Or (iTableType = iTabLookup)
    objColumn.UpdatePrivilege = fSysSecManager
    
    ' Release the security column
    Set objColumn = Nothing
    Set objColumns = Nothing
    
    rsInfo2.MoveNext
    
  Loop
  rsInfo2.Close

  ' If the current user is not a system/security manager then read the column permissions from SQL.
  If (Not fSysSecManager) And (sRealSourceList.Length <> 0) Then
    ' Get the SQL group id of the current user.
    sSQL = "SELECT gid" & _
      " FROM sysusers" & _
      " WHERE name = '" & pobjGroup.Name & "'"
    rsInfo3.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    lngRoleID = rsInfo3.Fields(0).value
    rsInfo3.Close
        
    sSQL = "EXEC dbo.[spASRGetAllTableAndViewColumnPermissionsForGroup] " & lngRoleID
    rsInfo4.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsInfo4.EOF

      ' Get the current column's table/view name.
      Set objTableView = Nothing

      strObjectName = UCase(rsInfo4.Fields(0).value)
      strColumnName = rsInfo4.Fields(1).value
      iSelect = rsInfo4.Fields(2).value
      iUpdate = rsInfo4.Fields(3).value

      fChildView = False
      If Left$(strObjectName, 8) = "ASRSYSCV" Then
        fChildView = True
        ' Determine which table the child view is for.
        iTemp = InStr(strObjectName, "#")
        lngChildViewID = val(Mid$(strObjectName, 9, iTemp - 9))
      End If

      If fChildView Then
        For lngNextIndex = 0 To UBound(avChildViews, 2)
          If avChildViews(1, lngNextIndex) = lngChildViewID Then
            Set objTableView = pobjGroup.Tables(avChildViews(2, lngNextIndex))
            objTableView.ParentJoinType = avChildViews(3, lngNextIndex)
            Exit For
          End If
        Next lngNextIndex
      
      Else
        If pobjGroup.Tables.IsValid(strObjectName) Then
          Set objTableView = pobjGroup.Tables(strObjectName)
        Else
          Set objTableView = pobjGroup.Views(strObjectName)
        End If
      End If
      
      
      If Not objTableView Is Nothing Then
      
        If objTableView.Columns.IsValid(strColumnName) Then
          objTableView.Columns(strColumnName).SelectPrivilege = iSelect
          objTableView.Columns(strColumnName).UpdatePrivilege = iUpdate
        End If
      
'        If iAction = 193 Then
'          If objTableView.Columns.IsValid(strColumnName) Then
'            objTableView.Columns(strColumnName).SelectPrivilege = bPermission
'          End If
'        End If
'
'        If iAction = 197 Then
'          If objTableView.Columns.IsValid(strColumnName) Then
'            objTableView.Columns(strColumnName).UpdatePrivilege = bPermission
'          End If
'        End If
      End If

      rsInfo4.MoveNext
    Loop
    rsInfo4.Close

    ' Check if the table has SELECT/UPDATE ALL/SOME/NONE.
    For Each objTableView In pobjGroup.Tables.Collection
      fSelectAllPermission = True
      fSelectNonePermission = True
      fUpdateAllPermission = True
      fUpdateNonePermission = True
      
      For Each objColumn In objTableView.Columns.Collection
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
      
      objTableView.SelectPrivilege = IIf(fSelectAllPermission, giPRIVILEGES_ALLGRANTED, IIf(fSelectNonePermission, giPRIVILEGES_NONEGRANTED, giPRIVILEGES_SOMEGRANTED))
      objTableView.UpdatePrivilege = IIf(fUpdateAllPermission, giPRIVILEGES_ALLGRANTED, IIf(fUpdateNonePermission, giPRIVILEGES_NONEGRANTED, giPRIVILEGES_SOMEGRANTED))
    Next objTableView
  End If
  
TidyUpAndExit:
  Set rsInfo1 = Nothing
  Set rsInfo2 = Nothing
  Set rsInfo3 = Nothing
  Set rsInfo4 = Nothing
  Set rsTables = Nothing
  Set rsViews = Nothing
  Set objTableView = Nothing
  Set objColumn = Nothing
  Set rsPermissions = Nothing
    
  SetupTablesCollection = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  OutputError "Error Reading Tables Collection"
  Resume TidyUpAndExit

End Function


Public Function CreateViewValidationStoredProcedure(pLngCurrentTableID As Long, _
  psCurrentTableName As String, _
  piTableType As Integer, _
  pfNewTable As Integer, _
  pfRefreshDatabase As Boolean) As Boolean
  ' Create the view validation stored procedure for the given table (plngCurrentTableID).
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim fCreateSP As Boolean
  Dim sSQL As String
  Dim sSPName As String
  Dim sSPCode As String
  Dim objGroup As clsSecurityGroup
  Dim rsInfo As DAO.Recordset

  fOK = True
  
  ' Check if we need to create the stored procedure.
  ' NB. We only need to create the stored procedure if the table is new, or if views of
  ' the table have been added or deleted.
  
  fCreateSP = (piTableType = iTabParent)
  If fCreateSP And (Not pfRefreshDatabase) Then
    If Not pfNewTable Then
      ' Check if there are any new or deleted views on this table.
      sSQL = "SELECT COUNT(*) AS recCount" & _
        " FROM tmpViews" & _
        " WHERE viewTableID = " & Trim$(Str$(pLngCurrentTableID)) & _
        " AND (new = TRUE" & _
        " OR deleted = TRUE)"
      Set rsInfo = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
      fCreateSP = (rsInfo!reccount > 0)
      rsInfo.Close
    End If
  End If

  If fCreateSP Or Application.ChangedViewName Then
    For Each objGroup In gObjGroups.Collection
      With objGroup
      
      ' Create the stored procedure creation string if the table is a top level table.
        sSPName = "[" & gsVIEWVALIDATIONSPPREFIX & Trim$(Str$(pLngCurrentTableID)) & "_" & .Name & "]"
    
        DropProcedure sSPName
            
        sSPCode = "/* ----------------------------------------------------------------------------------------------- */" & vbNewLine & _
          "/* view validation stored procedure.                               */" & vbNewLine & _
          "/* Automatically generated by the System/Security Managers.   */" & vbNewLine & _
          "/* ----------------------------------------------------------------------------------------------- */" & vbNewLine & _
          "CREATE PROCEDURE dbo." & sSPName & vbNewLine & _
          "(" & vbNewLine & _
          "    @pfResult bit OUTPUT," & vbNewLine & _
          "    @piRecordID integer" & vbNewLine & _
          ")" & vbNewLine & _
          "AS" & vbNewLine & _
          "BEGIN" & vbNewLine & _
          "    DECLARE @iRecCount integer" & vbNewLine
          
        ' If the current group has permission on the whole table return 1.
        If Not .Tables.IsValid(psCurrentTableName) Then
          ' The table is not in the tables collection, so it must be new.
          ' ie. Permission must be granted.
          sSPCode = sSPCode & vbNewLine & _
            "    SET @pfResult = 1" & vbNewLine
        Else
          If .Tables.Item(psCurrentTableName).SelectPrivilege Then
            sSPCode = sSPCode & vbNewLine & _
              "    SET @pfResult = 1" & vbNewLine
          Else
            sSPCode = sSPCode & vbNewLine & _
              "    SET @pfResult = 0" & vbNewLine
          
            ' Loop through the views adding code for each permissable one.
            If Not (recViewEdit.BOF And recViewEdit.EOF) Then
              recViewEdit.MoveFirst
              Do While Not recViewEdit.EOF
                If (Not recViewEdit!Deleted) And (recViewEdit!ViewTableID = pLngCurrentTableID) Then
                  
                  If Not .Views.IsValid(recViewEdit!OriginalViewName) Then
                    ' The view is not in the views collection, so it must be new.
                    ' ie. Permission must be granted.
                    sSPCode = sSPCode & vbNewLine & _
                      "    IF @pfResult = 0" & vbNewLine & _
                      "    BEGIN" & vbNewLine & _
                      "        /* Check if the current user can see the record in the '" & recViewEdit!ViewName & "' view. */" & vbNewLine & _
                      "        SELECT @iRecCount = COUNT(id)" & vbNewLine & _
                      "        FROM " & recViewEdit!ViewName & vbNewLine & _
                      "        WHERE id = @piRecordID" & vbNewLine & vbNewLine & _
                      "        IF @iRecCount > 0 SET @pfResult = 1" & vbNewLine & _
                      "    END" & vbNewLine
                  Else
                    If .Views.Item(recViewEdit!OriginalViewName).SelectPrivilege Then
                      sSPCode = sSPCode & vbNewLine & _
                        "    IF @pfResult = 0" & vbNewLine & _
                        "    BEGIN" & vbNewLine & _
                        "        /* Check if the current user can see the record in the '" & recViewEdit!ViewName & "' view. */" & vbNewLine & _
                        "        SELECT @iRecCount = COUNT(id)" & vbNewLine & _
                        "        FROM " & recViewEdit!ViewName & vbNewLine & _
                        "        WHERE id = @piRecordID" & vbNewLine & vbNewLine & _
                        "        IF @iRecCount > 0 SET @pfResult = 1" & vbNewLine & _
                        "    END" & vbNewLine
                    End If
                  End If
                End If
              
                recViewEdit.MoveNext
              Loop
            End If
          End If
        End If
        
        sSPCode = sSPCode & vbNewLine & _
          "END"
      
        ' Create the stored procedure.
        gADOCon.Execute sSPCode, , adCmdText + adExecuteNoRecords
      End With
    Next objGroup
    Set objGroup = Nothing
  End If
  
TidyUpAndExit:
  Set rsInfo = Nothing
  CreateViewValidationStoredProcedure = fOK
  Exit Function

ErrorTrap:
  fOK = False
  OutputError "Error creating view validation stored procedure"
  Err = False
 
 Resume TidyUpAndExit

End Function



Public Function ApplyPermissions() As Boolean
  ' Grant permissions to the tables/views/columns.
  ' Return TRUE if everything passed off okay.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  ' Apply the Table and Table Column permissions.
  OutputCurrentProcess2 "Parent Tables", 3
  fOK = ApplyPermissions_NonChildTables
  fOK = fOK And Not gobjProgress.Cancelled
  gobjProgress.UpdateProgress2
           
  If fOK Then
    ' Apply the View and View Column permissions.
    OutputCurrentProcess2 "Views", 3
    fOK = ApplyPermissions_UserViews
    fOK = fOK And Not gobjProgress.Cancelled
    gobjProgress.UpdateProgress2
  End If
   
  If fOK Then
    ' Apply the new Child Table (and child view) permissions.
    OutputCurrentProcess2 "Child Tables", 3
    fOK = ApplyPermissions_ChildTables2
    gobjProgress.UpdateProgress2
  End If
  
TidyUpAndExit:
  ApplyPermissions = fOK
  Exit Function
  
ErrorTrap:
  OutputError "Error Applying Permissions"
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function ApplyPermissions_NonChildTables() As Boolean
  ' Apply the top-level and lookup Table and Table Column permissions to SQL Server database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim sGroupName As String
  Dim sTableName As String
  Dim objGroup As clsSecurityGroup
  Dim rsTables As DAO.Recordset
  Dim rsColumns As DAO.Recordset
  Dim lngOriginalTableID As Long
  Dim sOriginalTableName As String
  Dim strOriginalColumnName As String
  Dim strColumnName As String
  Dim sSelectGrant As String
  Dim sUpdateGrant As String
  
  Dim objSourcePermissions As clsSecurityTable
  
  
  fOK = True

  ' Get the set of top-level and lookup tables.
  sSQL = "SELECT tableName, tableID, OriginalTableName, changed, new," & _
    " CopySecurityTableID, CopySecurityTableName, tableType," & _
    " GrantRead, GrantNew, GrantEdit, GrantDelete" & _
    " FROM tmpTables" & _
    " WHERE (tableType = " & Trim$(Str$(iTabParent)) & _
    " OR tableType = " & Trim$(Str$(iTabLookup)) & ")" & _
    " AND deleted = FALSE"
  Set rsTables = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    
  Do While Not rsTables.EOF
    For Each objGroup In gObjGroups.Collection
      With objGroup
        sGroupName = "[" & .Name & "]"
      
        sTableName = rsTables.Fields(0).value

        If objGroup.SecurityManager Or _
          objGroup.SystemManager Then

          sSQL = "GRANT DELETE, INSERT, SELECT, UPDATE ON " & sTableName & " TO " & sGroupName
          gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
        Else

          Set objSourcePermissions = Nothing
          If (Not rsTables!New) Or (rsTables!copySecurityTableID > 0) Then
            ' Initialise the Table Column permissions command strings.
            sSelectGrant = vbNullString
            sUpdateGrant = vbNullString
  
            ' Get the set of non-system columns in the table.
            If rsTables!copySecurityTableID > 0 Then
              lngOriginalTableID = rsTables!copySecurityTableID
              sOriginalTableName = rsTables!copySecurityTableName
            Else
              lngOriginalTableID = rsTables!TableID
              sOriginalTableName = rsTables!OriginalTableName
            End If

            Set objSourcePermissions = .Tables.Item(sOriginalTableName)
          End If


          If Not objSourcePermissions Is Nothing Then
            
            With objSourcePermissions
   
              If .SelectPrivilege = giPRIVILEGES_SOMEGRANTED Or .UpdatePrivilege = giPRIVILEGES_SOMEGRANTED Then
   
                sSQL = "SELECT columnName, columnID, OriginalColumnName, new" & _
                  " FROM tmpColumns" & _
                  " WHERE tableID = " & Trim$(Str$(lngOriginalTableID)) & _
                  " AND deleted = FALSE" & _
                  " AND columnType <> " & Trim$(Str$(giCOLUMNTYPE_SYSTEM)) & _
                  " AND columnType <> " & Trim$(Str$(giCOLUMNTYPE_LINK))
                Set rsColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    
                Do While Not rsColumns.EOF

                  'MH20060712 Fault 11313 - Addendum to JDM's work below.  :O)
                  '"OriginalColumnName" is only used for existing columns hence moved into ELSE clause
                  ''''JDM - 22/06/2006 - Fault 11186 - Addendum to MH's work below. Moved column generation list from outside loop to inside
                  ''''strOriginalColumnName = UCase(Trim(rsColumns.Fields(2).Value))
                  
                  
                  strColumnName = UCase(Trim(rsColumns.Fields(0).value))
                
                  If rsColumns.Fields(3).value Then   ' Is new column
  
                    ' Build string of columns that are allowed
                    If .SelectPrivilege = giPRIVILEGES_SOMEGRANTED Then
                      sSelectGrant = sSelectGrant & IIf(LenB(sSelectGrant) <> 0, ",", vbNullString) & strColumnName
                    End If
  
                    If .UpdatePrivilege = giPRIVILEGES_SOMEGRANTED Then
                      sUpdateGrant = sUpdateGrant & IIf(LenB(sUpdateGrant) <> 0, ",", vbNullString) & strColumnName
                    End If
                  
                  Else
                    'MH20060712 Fault 11313
                    strOriginalColumnName = UCase(Trim(rsColumns.Fields(2).value))

                    ' Build string of columns that are revoked based on what they had before
                    If .Columns.Item(strOriginalColumnName).SelectPrivilege Or .TableType = iTabLookup Then
                      sSelectGrant = sSelectGrant & IIf(LenB(sSelectGrant) <> 0, ",", vbNullString) & strColumnName
                    End If
                      
                    If .Columns.Item(strOriginalColumnName).UpdatePrivilege Then
                      sUpdateGrant = sUpdateGrant & IIf(LenB(sUpdateGrant) <> 0, ",", vbNullString) & strColumnName
                    End If
                  End If
                
                  rsColumns.MoveNext
                Loop
                
                rsColumns.Close
                Set rsColumns = Nothing
              End If
            
              ' Delete permissions
              If .DeletePrivilege Then
                sSQL = "GRANT DELETE ON " & sTableName & " TO " & sGroupName
                gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
              End If
            
              ' Insert permissions
              If .InsertPrivilege Then
                sSQL = "GRANT INSERT ON " & sTableName & " TO " & sGroupName
                gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
              End If
            
              ' Select permissions
              If .SelectPrivilege = giPRIVILEGES_ALLGRANTED Or (.TableType = iTabLookup) Then
                sSQL = "GRANT SELECT ON " & sTableName & " TO " & sGroupName
                gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
              ElseIf .SelectPrivilege = giPRIVILEGES_SOMEGRANTED Then
                If LenB(sSelectGrant) <> 0 Then
                  
                  sSQL = "REVOKE SELECT ON " & sTableName & " TO " & sGroupName
                  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
                  
                  'MH20060620 Fault 11186
                  'sSQL = "GRANT SELECT(ID, " & sSelectGrant & ") ON " & sTableName & " TO " & sGroupName
                  sSQL = "GRANT SELECT(ID,TimeStamp," & sSelectGrant & ") ON " & sTableName & " TO " & sGroupName
                  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
                End If
              Else
                gADOCon.Execute "GRANT SELECT(id, TimeStamp) ON " & sTableName & " TO " & sGroupName
              End If
              
              ' Update permissions
              If .UpdatePrivilege = giPRIVILEGES_ALLGRANTED Then
                sSQL = "GRANT UPDATE ON " & sTableName & " TO " & sGroupName
                gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
              ElseIf .UpdatePrivilege = giPRIVILEGES_SOMEGRANTED Then
                If LenB(sUpdateGrant) <> 0 Then
                  
                  sSQL = "REVOKE UPDATE ON " & sTableName & " TO " & sGroupName
                  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
                  
                  'MH20060620 Fault 11186
                  'sSQL = "GRANT UPDATE(" & sUpdateGrant & ") ON " & sTableName & " TO " & sGroupName
                  sSQL = "GRANT UPDATE(ID,Timestamp," & sUpdateGrant & ") ON " & sTableName & " TO " & sGroupName
                  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
                End If
              End If
              
            End With
            
          Else
            ' New table, not having permissions copied from another table.
            ' Put user defined security settings on each group
            If rsTables!GrantDelete Then
              sSQL = "GRANT DELETE ON " & sTableName & " TO " & sGroupName
              gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
            End If
            
            If rsTables!GrantNew Then
              sSQL = "GRANT INSERT ON " & sTableName & " TO " & sGroupName
              gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
            End If
            
            If rsTables!GrantRead Or (rsTables!TableType = iTabLookup) Then
              sSQL = "GRANT SELECT ON " & sTableName & " TO " & sGroupName
              gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
            End If
            
            gADOCon.Execute "GRANT SELECT(id, TimeStamp) ON " & sTableName & " TO " & sGroupName
            
            If rsTables!GrantEdit Then
              sSQL = "GRANT UPDATE ON " & sTableName & " TO " & sGroupName
              gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
            End If

          End If
        End If
      End With
    Next objGroup
    Set objGroup = Nothing
    
    rsTables.MoveNext
  Loop
  rsTables.Close
  Set rsTables = Nothing
  
TidyUpAndExit:
  Set objGroup = Nothing
  ApplyPermissions_NonChildTables = fOK
  Exit Function

ErrorTrap:
  OutputError "Error Applying Permissions (Non-Child Tables)"
  fOK = False
  Resume TidyUpAndExit
  
End Function



Private Function ApplyPermissions_UserViews() As Boolean
  ' Apply the user-defined views permissions to SQL Server database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim sGroupName As String
  Dim sViewName As String
  Dim sOriginalViewName As String
'  Dim sSelectDeny As String
'  Dim sUpdateDeny As String
  Dim objGroup As clsSecurityGroup
  Dim rsViews As DAO.Recordset
  Dim rsColumns As DAO.Recordset
  Dim sPermissionTypes As String
  Dim strOriginalColumnName As String
  Dim strColumnName As String
  Dim sSelectGrant As String
  Dim sUpdateGrant As String
  
  fOK = True

  ' Get the set of user defined views.
  sSQL = "SELECT viewName, viewID, viewTableID, OriginalViewName, changed, new," & _
    " GrantRead, GrantNew, GrantEdit, GrantDelete" & _
    " FROM tmpViews" & _
    " WHERE deleted = FALSE"
  Set rsViews = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    
  Do While Not rsViews.EOF
    For Each objGroup In gObjGroups.Collection
      With objGroup
        sGroupName = "[" & .Name & "]"
        sViewName = rsViews!ViewName

        If objGroup.SecurityManager Or _
          objGroup.SystemManager Then
          sSQL = "GRANT DELETE, INSERT, SELECT, UPDATE ON " & sViewName & " TO " & sGroupName
          gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
        Else
          'NPG20080206 Fault 12874
          ' If Not rsViews!New Then
          If Not IsNull(rsViews!OriginalViewName) Then
            
            sOriginalViewName = rsViews!OriginalViewName

            ' Initialise the column permissions command strings.
            sSelectGrant = vbNullString
            sUpdateGrant = vbNullString

            With .Views.Item(sOriginalViewName)

              ' Get the set of non-system columns in the view.
              sSQL = "SELECT tmpColumns.columnName, tmpColumns.columnID, tmpColumns.OriginalColumnName, tmpColumns.new" & _
                " FROM tmpViewColumns, tmpColumns" & _
                " WHERE (tmpViewColumns.ColumnID = tmpColumns.ColumnID" & _
                " AND tmpColumns.columnType <> " & Trim$(Str$(giCOLUMNTYPE_SYSTEM)) & _
                " AND tmpColumns.columnType <> " & Trim$(Str$(giCOLUMNTYPE_LINK)) & _
                " AND tmpViewColumns.InView = TRUE" & _
                " AND tmpViewColumns.ViewID = " & Trim(Str(rsViews!ViewID)) & ")"
              Set rsColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  
              Do While Not rsColumns.EOF
              
                strOriginalColumnName = UCase(Trim(IIf(IsNull(rsColumns.Fields(2).value), "", rsColumns.Fields(2).value)))
                strColumnName = UCase(Trim(rsColumns.Fields(0).value))
              
                If Not .Columns.IsValid(strOriginalColumnName) Then
                  
                  ' New column
                  If .SelectPrivilege = giPRIVILEGES_SOMEGRANTED Then
                    sSelectGrant = sSelectGrant & IIf(LenB(sSelectGrant) <> 0, ",", vbNullString) & strColumnName
                  End If
  
                  If .UpdatePrivilege = giPRIVILEGES_SOMEGRANTED Then
                    sUpdateGrant = sUpdateGrant & IIf(LenB(sUpdateGrant) <> 0, ",", vbNullString) & strColumnName
                  End If
                Else
                
                  ' Existing column
                  If .Columns.Item(strOriginalColumnName).SelectPrivilege Then
                    sSelectGrant = sSelectGrant & IIf(LenB(sSelectGrant) <> 0, ",", vbNullString) & strColumnName
                  End If
  
                  If .Columns.Item(strOriginalColumnName).UpdatePrivilege Then
                    sUpdateGrant = sUpdateGrant & IIf(LenB(sUpdateGrant) <> 0, ",", vbNullString) & strColumnName
                  End If
                End If
  
                rsColumns.MoveNext
              Loop
              rsColumns.Close
              Set rsColumns = Nothing
          
              ' Delete permission
              If .DeletePrivilege Then
                sSQL = "GRANT DELETE ON " & sViewName & " TO " & sGroupName
                gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
              End If
                        
              ' Insert permission
              If .InsertPrivilege Then
                sSQL = "GRANT INSERT ON " & sViewName & " TO " & sGroupName
                gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
              End If
                          
              ' Select permissions
              If .SelectPrivilege = giPRIVILEGES_ALLGRANTED Then
                sSQL = "GRANT SELECT ON " & sViewName & " TO " & sGroupName
                gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
              ElseIf .SelectPrivilege = giPRIVILEGES_SOMEGRANTED Then
                If LenB(sSelectGrant) <> 0 Then
                  'MH20060620 Fault 11186
                  'sSQL = "GRANT SELECT(" & sSelectGrant & ") ON " & sViewName & " TO " & sGroupName
                  sSQL = "GRANT SELECT(ID,Timestamp," & sSelectGrant & ") ON " & sViewName & " TO " & sGroupName
                  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
                End If
              End If
                
              ' Update permissions
              If .UpdatePrivilege = giPRIVILEGES_ALLGRANTED Then
                sSQL = "GRANT UPDATE ON " & sViewName & " TO " & sGroupName
                gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
              ElseIf .UpdatePrivilege = giPRIVILEGES_SOMEGRANTED Then
                If LenB(sUpdateGrant) <> 0 Then
                  'MH20060620 Fault 11186
                  'sSQL = "GRANT UPDATE(" & sUpdateGrant & ") ON " & sViewName & " TO " & sGroupName
                  sSQL = "GRANT UPDATE(ID,Timestamp," & sUpdateGrant & ") ON " & sViewName & " TO " & sGroupName
                  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
                End If
              End If
                                      
            End With
                                      
          Else
            ' New view, not having permissions copied from another view.
            ' Put user defined security settings on each group
            If rsViews!GrantDelete Then
              sSQL = "GRANT DELETE ON " & sViewName & " TO " & sGroupName
              gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
            End If
            
            If rsViews!GrantNew Then
              sSQL = "GRANT INSERT ON " & sViewName & " TO " & sGroupName
              gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
            End If
            
            If rsViews!GrantRead Then
              sSQL = "GRANT SELECT ON " & sViewName & " TO " & sGroupName
              gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
            End If
            
            If rsViews!GrantEdit Then
              sSQL = "GRANT UPDATE ON " & sViewName & " TO " & sGroupName
              gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
            End If

          
          End If
          
        End If
      End With
    Next objGroup
    Set objGroup = Nothing
    
    rsViews.MoveNext
  Loop
  rsViews.Close
  Set rsViews = Nothing
  
TidyUpAndExit:
  Set objGroup = Nothing
  ApplyPermissions_UserViews = fOK
  Exit Function

ErrorTrap:
  OutputError "Error Applying Permissions (User Views)"
  fOK = False
  Resume TidyUpAndExit
  
End Function



Private Function ApplyPermissions_ChildTables2() As Boolean
  ' Apply the child table and permutated view permissions to SQL Server database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim sAllSQLCommands As String
  Dim sGroupName As String
  Dim sTableName As String
  'Dim sSelectDeny As String
  'Dim sUpdateDeny As String
  Dim sSelectGrant As String
  Dim sUpdateGrant As String
  Dim objGroup As clsSecurityGroup
  Dim rsTables As DAO.Recordset
  Dim rsTables2 As DAO.Recordset
  Dim rsColumns As DAO.Recordset
  Dim cmdChildView As ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim lngViewID As Long
  Dim sViewName As String
  Dim avChildTables() As Variant
  Dim iMaxRouteLength As Integer
  Dim rsChildren As DAO.Recordset
  Dim rsParents As DAO.Recordset
  Dim rsViews As DAO.Recordset
  Dim rsInfo1 As ADODB.Recordset
  Dim rsInfo2 As ADODB.Recordset
  Dim rsInfo3 As ADODB.Recordset
  Dim rsInfo4 As ADODB.Recordset
  Dim rsInfo5 As ADODB.Recordset
  Dim rsChildViews As ADODB.Recordset
  Dim iNextIndex As Integer
  Dim iLoop As Integer
  Dim iLoop1 As Integer
  Dim iLoop2 As Integer
  Dim iParentCount As Integer
  Dim avParents() As Variant
  Dim fTableOK As Boolean
  Dim iOKViewCount As Integer
  Dim fViewOK As Boolean
  Dim sTemp As String
  Dim iParentJoinType As Integer
  Dim lngParentViewID As Long
  Dim sCreatedChildViews As SystemMgr.cStringBuilder
  Dim lngLastParentID As Long
  Dim lngOriginalTableID As Long
  Dim sOriginalTableName As String
  Dim fTableReadable As Boolean
  Dim sRelatedChildTables As String
  Dim sSysSecRoles As String
  Dim sNonSysSecRoles As String
  Dim lngPreviousTimeOut As Long
  Dim sColumnName As String
  Dim sOriginalColumnName As String
  Dim strParentIDs As String
  
  Dim blnRemoveFix1082 As Boolean
  
  blnRemoveFix1082 = (GetSystemSetting("Remove Fix", "1082", "0") = "1")
  
  
  Set sCreatedChildViews = New SystemMgr.cStringBuilder
  Set rsInfo1 = New ADODB.Recordset
  Set rsChildViews = New ADODB.Recordset

  fOK = True
  sCreatedChildViews.TheString = "0"
  sRelatedChildTables = "0"
  sAllSQLCommands = vbNullString
  
  ' Drop all existing child views.
  sSQL = "SELECT name FROM sysobjects WHERE name LIKE 'ASRSysCV%' AND xtype = 'V'"
  rsInfo1.Open sSQL, gADOCon, adOpenDynamic, adLockReadOnly
    
  With rsInfo1
    Do While (Not .EOF)
      sSQL = "DROP VIEW " & .Fields(0).value
      gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
      .MoveNext
    Loop
  
    .Close
  End With
  Set rsInfo1 = Nothing
  
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
  ReDim avChildTables(2, 0)
  iMaxRouteLength = 0
  sSQL = "SELECT DISTINCT tmpRelations.childID, tmpTables.tableName" & _
    " FROM tmpRelations, tmpTables" & _
    " WHERE tmpRelations.childID = tmpTables.tableID"
  Set rsChildren = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  
  ReDim avChildTables(2, 100)
  iNextIndex = -1
  If Not rsChildren.EOF Then
    Do While Not rsChildren.EOF
      iNextIndex = iNextIndex + 1
      If iNextIndex > UBound(avChildTables, 2) Then ReDim Preserve avChildTables(2, iNextIndex + 100)
      avChildTables(1, iNextIndex) = rsChildren.Fields(0).value
      avChildTables(2, iNextIndex) = LongestRouteToTopLevel(rsChildren.Fields(0).value)
      
      sRelatedChildTables = sRelatedChildTables & "," & Trim(Str(rsChildren.Fields(0).value))
      
      If iMaxRouteLength < avChildTables(2, iNextIndex) Then
        iMaxRouteLength = CInt(avChildTables(2, iNextIndex))
      End If
      
      rsChildren.MoveNext
    Loop
    ReDim Preserve avChildTables(2, iNextIndex)
  End If
  rsChildren.Close
  Set rsChildren = Nothing
  
  
  ' Deny non-SysMgr and non-SecMgr users access to orphaned child tables.
  sSysSecRoles = vbNullString
  sNonSysSecRoles = vbNullString
  For Each objGroup In gObjGroups.Collection
    sGroupName = "[" & objGroup.Name & "]"
    
    If (objGroup.SecurityManager) Or (objGroup.SystemManager) Then
      sSysSecRoles = sSysSecRoles & IIf(LenB(sSysSecRoles) <> 0, ",", vbNullString) & sGroupName
    Else
      sNonSysSecRoles = sNonSysSecRoles & IIf(LenB(sNonSysSecRoles) <> 0, ",", vbNullString) & sGroupName
    End If
  Next objGroup
  
  ' JPD20021120 Fault 4793
  sSQL = "SELECT tmpTables.tableName" & _
    " FROM tmpTables" & _
    " WHERE tmpTables.tableType = " & Trim$(Str$(iTabChild)) & _
    " AND tmpTables.tableID NOT IN (" & sRelatedChildTables & ")" & _
    " AND tmpTables.deleted = FALSE"
  Set rsChildren = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  With rsChildren
    Do While (Not .EOF)
      If LenB(sSysSecRoles) <> 0 Then
        sSQL = "GRANT DELETE, INSERT, SELECT, UPDATE ON " & .Fields(0).value & " TO " & sSysSecRoles
        sAllSQLCommands = sAllSQLCommands & vbNewLine & sSQL
        gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
      End If
      .MoveNext
    Loop
  
    .Close
  End With
  Set rsChildren = Nothing
  
  ' For each child table (do those nearest to the top-level first).
  For iLoop1 = 1 To iMaxRouteLength
    ' For each table this distance from the top-level.
    For iLoop2 = 0 To UBound(avChildTables, 2)
      If CInt(avChildTables(2, iLoop2)) = iLoop1 Then
        
        ' Get the child table info.
        sSQL = "SELECT tableName, tableID, OriginalTableName, changed, new, GrantRead, GrantNew, GrantEdit, GrantDelete, CopySecurityTableID, CopySecurityTableName" & _
          " FROM tmpTables" & _
          " WHERE tableID = " & Trim$(Str$(avChildTables(1, iLoop2)))
        Set rsTables = daoDb.OpenRecordset(sSQL, _
          dbOpenForwardOnly, dbReadOnly)

        ' Set permissions differently if we are a copied table
        If rsTables!copySecurityTableID > 0 Then
          lngOriginalTableID = rsTables!copySecurityTableID
          sOriginalTableName = rsTables!copySecurityTableName
        Else
          lngOriginalTableID = rsTables!TableID
          sOriginalTableName = rsTables!OriginalTableName
        End If
        
        For Each objGroup In gObjGroups.Collection
          sGroupName = "[" & objGroup.Name & "]"
          sTableName = rsTables!TableName
          
          If (objGroup.SecurityManager) Or (objGroup.SystemManager) Then
            sSQL = "GRANT DELETE, INSERT, SELECT, UPDATE ON " & sTableName & " TO " & sGroupName
            gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
            
'''''          Else
'''''            sSQL = "DENY DELETE, INSERT, SELECT, UPDATE ON " & sTableName & " TO " & sGroupName
          End If

          If Not rsTables!New Or rsTables!copySecurityTableID > 0 Then

            ' NPG20080612 Fault 13198
            ' If the originaltablename is not in the objgroup collection, the original table was a lookup which was then copied and changed to a child
            If Not objGroup.Tables.IsValid(sOriginalTableName) Then
              ' Get the child table info.
              sSQL = "SELECT tableName, tableID, OriginalTableName, changed, new, GrantRead, GrantNew, GrantEdit, GrantDelete, CopySecurityTableID, CopySecurityTableName" & _
                " FROM tmpTables" & _
                " WHERE tableID = " & Trim$(Str$(lngOriginalTableID))
              Set rsTables2 = daoDb.OpenRecordset(sSQL, _
                dbOpenForwardOnly, dbReadOnly)
                
              fTableReadable = rsTables2!GrantRead
              
            Else
              fTableReadable = (objGroup.Tables.Item(sOriginalTableName).SelectPrivilege <> giPRIVILEGES_NONEGRANTED)
            End If
          Else
            ' Is a new table, or a copied one with specified permissions
            fTableReadable = rsTables!GrantRead
          End If
          
          'NPG20080509 Fault 12975
          ' If fTableReadable
          If fTableReadable Or (objGroup.SecurityManager) Or (objGroup.SystemManager) Then
              
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
            strParentIDs = ""
            sSQL = "SELECT tmpTables.tableID, tmpTables.tableName, tmpTables.originalTableName, tmpTables.tableType," & _
              " tmpTables.copySecurityTableID, tmpTables.copySecurityTableName, tmpTables.grantRead" & _
              " FROM tmpRelations" & _
              " INNER JOIN tmpTables ON tmpRelations.parentID = tmpTables.tableID" & _
              " WHERE tmpRelations.childID = " & Trim$(Str$(avChildTables(1, iLoop2)))
            
            Set rsParents = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
            Do While (Not rsParents.EOF)
              strParentIDs = strParentIDs & _
                IIf(LenB(strParentIDs) > 0, ",", "") & _
                "ID_" & CStr(rsParents!TableID)
              
              If rsParents!TableType = iTabParent Then
                ' Parent is a top-level table.
                If objGroup.Tables.IsValid(rsParents!OriginalTableName) Then
                  ' Table is in the tables collection.
                  fTableOK = (objGroup.Tables(rsParents!OriginalTableName).SelectPrivilege <> giPRIVILEGES_NONEGRANTED)
                Else
                  ' Table is NOT in the tables collection. Must be new,
                  ' so see if we're copying permissions from another table, or
                  ' if the default permissions are specified.
                  If rsParents!copySecurityTableID > 0 Then
                    fTableOK = (objGroup.Tables(rsParents!copySecurityTableName).SelectPrivilege <> giPRIVILEGES_NONEGRANTED)
                  Else
                    fTableOK = rsParents!GrantRead
                  End If
                End If
                    
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
                  sSQL = "SELECT tmpViews.viewID, tmpViews.viewName, tmpViews.originalViewName, tmpViews.grantRead" & _
                    " FROM tmpViews" & _
                    " WHERE tmpViews.viewTableID = " & Trim(Str(rsParents!TableID)) & _
                    " AND tmpViews.deleted = FALSE"
                  iOKViewCount = 0
                      
                  Set rsViews = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
                  Do While Not rsViews.EOF
                    If objGroup.Views.IsValid(rsViews!OriginalViewName) Then
                      ' View is in the views collection.
                      fViewOK = (objGroup.Views(rsViews!OriginalViewName).SelectPrivilege <> giPRIVILEGES_NONEGRANTED)
                    Else
                      ' View is NOT in the views collection. Must be new,
                      ' so get the default permissions.
                      fViewOK = rsViews!GrantRead
                    End If
                          
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
                  Set rsViews = Nothing
                    
                  If iOKViewCount > 0 Then
                    iParentCount = iParentCount + 1
                  End If
                End If
              ElseIf rsParents!TableType = iTabChild Then
                ' Parent is not a top-level table.
                lngParentViewID = 0
                sSQL = "SELECT childViewID FROM ASRSysChildViews2 WHERE role = '" & objGroup.Name & "' AND tableID = " & Trim(Str(rsParents!TableID))
                rsChildViews.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
                
                Do While (Not rsChildViews.EOF) And (lngParentViewID = 0)
                  lngParentViewID = rsChildViews!childViewID
                        
                  ' Check if it really exists.
                  Set rsInfo2 = New ADODB.Recordset
                  
                  sTemp = Left("ASRSysCV" & Trim$(Str$(lngParentViewID)) & "#" & Replace(rsParents!TableName, " ", "_") & "#" & Replace(objGroup.Name, " ", "_"), 255)
                  sSQL = "SELECT COUNT(Name) AS result FROM sysobjects WHERE name = '" & sTemp & "' AND xtype = 'V'"
                  rsInfo2.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
                  If rsInfo2.Fields(0).value = 0 Then
                    lngParentViewID = 0
                  End If
                  
                  rsInfo2.Close
                  Set rsInfo2 = Nothing
                        
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
                  avParents(4, iNextIndex) = Left("ASRSysCV" & Trim$(Str$(lngParentViewID)) & "#" & Replace(rsParents!TableName, " ", "_") & "#" & Replace(objGroup.Name, " ", "_"), 255)
                End If
              End If
                  
              rsParents.MoveNext
            Loop
              
            rsParents.Close
            Set rsParents = Nothing
                  
            If iParentCount > 0 Then
              iParentJoinType = 0
              
              ' More than 1 parent. Do we want the OR join child view, or the AND join child view ?
              If objGroup.Tables.IsValid(rsTables!OriginalTableName) Then
                iParentJoinType = objGroup.Tables(rsTables!OriginalTableName).ParentJoinType
              End If
  
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
                pmADO.value = rsTables!TableID
            
                Set pmADO = .CreateParameter("JoinType", adInteger, adParamInput)
                .Parameters.Append pmADO
                pmADO.value = iParentJoinType
                
                Set pmADO = .CreateParameter("Name", adVarChar, adParamInput, 256)
                .Parameters.Append pmADO
                pmADO.value = objGroup.Name
            
                .Execute
            
                lngViewID = IIf(IsNull(.Parameters(0).value), vbNullString, .Parameters(0).value)
              End With
              Set cmdChildView = Nothing
      
      
              sCreatedChildViews.Append "," & Trim$(Str$(lngViewID))
                  
              ' Delete the existing entries in the ASRSysChildViewParents2 table.
              sSQL = "DELETE FROM ASRSysChildViewParents2 WHERE childViewID = " & Trim$(Str$(lngViewID))
              gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
                
              For iNextIndex = 1 To UBound(avParents, 2)
                sSQL = "INSERT INTO ASRSysChildViewParents2" & _
                  " (childViewID, parentType, parentID, parentTableID)" & _
                  " VALUES (" & Trim$(Str$(lngViewID)) & ", " & _
                  "'" & avParents(1, iNextIndex) & "', " & _
                  Trim$(Str$(avParents(2, iNextIndex))) & ", " & _
                  Trim$(Str$(avParents(3, iNextIndex))) & ")"
                gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
              Next iNextIndex
              
              ' Create the view name.
              sViewName = Left$("ASRSysCV" & Trim$(Str$(lngViewID)) & "#" & Replace(sTableName, " ", "_") & "#" & Replace(objGroup.Name, " ", "_"), 255)
              
              
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
                  sTableName & ".ID_" & Trim$(Str$(avParents(3, 1))) & " IN (SELECT id FROM " & avParents(4, 1) & ")" & vbNewLine
  
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
                    sTableName & ".ID_" & Trim$(Str$(avParents(3, iLoop))) & " IN (SELECT id FROM " & avParents(4, iLoop) & ")" & vbNewLine
                Next iLoop
  
                sSQL = sSQL & "                )"
  
                gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords

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
                sAllSQLCommands = sAllSQLCommands & vbNewLine & sSQL
                gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
              End If
                
'''''              If Len(sNonSysSecRoles) > 0 Then
'''''                sSQL = "DENY DELETE, INSERT, SELECT, UPDATE ON " & sViewName & " TO " & sNonSysSecRoles
'''''                sAllSQLCommands = sAllSQLCommands & vbNewLine & sSQL
'''''                gADOCon.Execute sSQL, , adExecuteNoRecords
'''''              End If
                
              If (Not objGroup.SecurityManager) And (Not objGroup.SystemManager) Then
                ' Apply the configured permissions to the child view permutation.
                ' Initialise the Table Column permissions command strings.
                sSelectGrant = "ID,Timestamp," & strParentIDs
                sUpdateGrant = "ID,Timestamp," & strParentIDs

                'NPG20080206 SUGG S000586
                ' If (Not rsTables!New) Or (rsTables!copySecurityTableID > 0)Then
                'NPG20080613 Fault 13198
                If ((Not rsTables!New) Or (rsTables!copySecurityTableID > 0)) And objGroup.Tables.IsValid(sOriginalTableName) Then
                'If Trim(rsViews!OriginalViewName) <> vbNullString Then
    
                  With objGroup.Tables.Item(sOriginalTableName)
                    If .SelectPrivilege = giPRIVILEGES_SOMEGRANTED Or .UpdatePrivilege = giPRIVILEGES_SOMEGRANTED Then
                    
                      ' Get the set of non-system columns in the table.
                      sSQL = "SELECT columnName, OriginalColumnName, new" & _
                        " FROM tmpColumns" & _
                        " WHERE tableID = " & Trim$(Str$(lngOriginalTableID)) & _
                        " AND deleted = FALSE" & _
                        " AND columnType <> " & Trim$(Str$(giCOLUMNTYPE_SYSTEM)) & _
                        " AND columnType <> " & Trim$(Str$(giCOLUMNTYPE_LINK))
                      Set rsColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
                      
                      Do While Not rsColumns.EOF
                      
                        sColumnName = rsColumns.Fields(0).value
                        sOriginalColumnName = Trim(UCase(IIf(IsNull(rsColumns.Fields(1).value), "", rsColumns.Fields(1).value)))
                      
                        If rsColumns.Fields(2).value Then ' Is New
  
                          ' Build string of columns that are allowed
                          If .SelectPrivilege = giPRIVILEGES_SOMEGRANTED Then
                            sSelectGrant = sSelectGrant & IIf(LenB(sSelectGrant) <> 0, ",", vbNullString) & sColumnName
                          End If
        
                          If .UpdatePrivilege = giPRIVILEGES_SOMEGRANTED Then
                            sUpdateGrant = sUpdateGrant & IIf(LenB(sUpdateGrant) <> 0, ",", vbNullString) & sColumnName
                          End If
                        
                        Else
                        
                          ' Build string of columns that are revoked based on what they had before
                          If .Columns.Item(sOriginalColumnName).SelectPrivilege Or .TableType = iTabLookup Then
                            sSelectGrant = sSelectGrant & IIf(LenB(sSelectGrant) <> 0, ",", vbNullString) & sColumnName
                          End If
                            
                          If .Columns.Item(sOriginalColumnName).UpdatePrivilege Then
                            sUpdateGrant = sUpdateGrant & IIf(LenB(sUpdateGrant) <> 0, ",", vbNullString) & sColumnName
                          End If
                        End If
    
                        rsColumns.MoveNext
                      Loop
    
                      rsColumns.Close
                      Set rsColumns = Nothing
                    End If
                    
                    
                    ' Delete permissions
                    If .DeletePrivilege Then
                      sSQL = "GRANT DELETE ON " & sViewName & " TO " & sGroupName
                      gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
                    End If
                  
                    ' Insert permissions
                    If .InsertPrivilege Then
                      sSQL = "GRANT INSERT ON " & sViewName & " TO " & sGroupName
                      gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
                    End If
                  
                    ' Select permissions
                    If .SelectPrivilege = giPRIVILEGES_ALLGRANTED Or (.TableType = iTabLookup) Then
                      sSQL = "GRANT SELECT ON " & sViewName & " TO " & sGroupName
                      gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
                    ElseIf .SelectPrivilege = giPRIVILEGES_SOMEGRANTED Then
                      If LenB(sSelectGrant) <> 0 Then
                        'MH20060620 Fault 11186
                        'sSQL = "GRANT SELECT(" & sSelectGrant & ") ON " & sViewName & " TO " & sGroupName
                        sSQL = "GRANT SELECT(ID,Timestamp," & sSelectGrant & ") ON " & sViewName & " TO " & sGroupName
                        gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
                      End If
                    End If
                    
                    ' Update permissions
                    If .UpdatePrivilege = giPRIVILEGES_ALLGRANTED Then
                      sSQL = "GRANT UPDATE ON " & sViewName & " TO " & sGroupName
                      gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
                    ElseIf .UpdatePrivilege = giPRIVILEGES_SOMEGRANTED Then
                      sSQL = "SELECT columnName, OriginalColumnName, new" & _
                        " FROM tmpColumns" & _
                        " WHERE tableID = " & Trim$(Str$(lngOriginalTableID)) & _
                        " AND deleted = FALSE" & _
                        " AND columnType = " & Trim$(Str$(giCOLUMNTYPE_LINK))
                      Set rsColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

                      Do While Not rsColumns.EOF
                        sColumnName = rsColumns.Fields(0).value
                        sUpdateGrant = sUpdateGrant & IIf(LenB(sUpdateGrant) <> 0, ",", vbNullString) & sColumnName

                        rsColumns.MoveNext
                      Loop

                      rsColumns.Close
                      Set rsColumns = Nothing
                      
                      If LenB(sUpdateGrant) <> 0 Then
                        'MH20060620 Fault 11186
                        'sSQL = "GRANT UPDATE(" & sUpdateGrant & ") ON " & sViewName & " TO " & sGroupName
                        sSQL = "GRANT UPDATE(ID,Timestamp," & sUpdateGrant & ") ON " & sViewName & " TO " & sGroupName
                        gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
                      End If
                    End If
                  
                  End With
                Else
                  ' Is a new table, or a copied one with specified permissions
                  'NPG20080613 Fault 13198 - added the IIFs
                  ' If rsTables!GrantRead Then
                  'NPG20081003 Fault 13388 - If new table use applied permissions.
                  If (objGroup.Tables.IsValid(sOriginalTableName) Or rsTables!New) And rsTables2 Is Nothing Then
                      If rsTables!GrantRead Then
                        sSQL = "GRANT SELECT ON " & sViewName & " TO " & sGroupName
                        gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
                      End If
        
                      If rsTables!GrantEdit Then
                        sSQL = "GRANT UPDATE ON " & sViewName & " TO " & sGroupName
                        gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
                      End If
        
                      If rsTables!GrantNew Then
                        sSQL = "GRANT INSERT ON " & sViewName & " TO " & sGroupName
                        gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
                      End If
        
                      If rsTables!GrantDelete Then
                        sSQL = "GRANT DELETE ON " & sViewName & " TO " & sGroupName
                        gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
                      End If
                  Else
                      If rsTables2!GrantRead Then
                        sSQL = "GRANT SELECT ON " & sViewName & " TO " & sGroupName
                        gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
                      End If
        
                      If rsTables2!GrantEdit Then
                        sSQL = "GRANT UPDATE ON " & sViewName & " TO " & sGroupName
                        gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
                      End If
        
                      If rsTables2!GrantNew Then
                        sSQL = "GRANT INSERT ON " & sViewName & " TO " & sGroupName
                        gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
                      End If
        
                      If rsTables2!GrantDelete Then
                        sSQL = "GRANT DELETE ON " & sViewName & " TO " & sGroupName
                        gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
                      End If
                  
                  End If
                End If
              End If
            End If
          End If
        
        'NPG20080613 Fault 13198
        If Not rsTables2 Is Nothing Then
          rsTables2.Close
          Set rsTables2 = Nothing
        End If

        
        Next objGroup
        Set objGroup = Nothing

        rsTables.Close
        Set rsTables = Nothing
      End If
    Next iLoop2
  Next iLoop1
  
  ' Delete invalid records from ASRSysChildViews2 and ASRSysChildViewParents2
  sSQL = "DELETE FROM ASRSysChildViews2 WHERE childViewID NOT IN (" & sCreatedChildViews.ToString & ")"
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords

  sSQL = "DELETE FROM ASRSysChildViewParents2 WHERE childViewID NOT IN (" & sCreatedChildViews.ToString & ")"
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
  ' Drop all existing old style child views.
  Set rsInfo3 = New ADODB.Recordset
  sSQL = "SELECT name FROM sysobjects WHERE name LIKE 'ASRSysChildView[_]%' AND xtype = 'V'"
  rsInfo3.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly

  With rsInfo3
    Do While (Not .EOF)
      sSQL = "DROP VIEW " & .Fields(0).value
      gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords

      .MoveNext
    Loop

    .Close
  End With
  Set rsInfo3 = Nothing

  ' Drop the old style child view tables.
  Set rsInfo4 = New ADODB.Recordset
  sSQL = "SELECT COUNT(Name) AS result FROM sysobjects WHERE name = 'ASRSysChildViews' AND xtype = 'U'"
  rsInfo4.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  
  If rsInfo4!result > 0 Then
    sSQL = "DROP TABLE ASRSysChildViews"
    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  End If
  rsInfo4.Close
  Set rsInfo4 = Nothing
  
  Set rsInfo5 = New ADODB.Recordset
  sSQL = "SELECT COUNT(Name) AS result FROM sysobjects WHERE name = 'ASRSysChildViewParents' AND xtype = 'U'"
  rsInfo5.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  
  If rsInfo5!result > 0 Then
    sSQL = "DROP TABLE ASRSysChildViewParents"
    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  End If
  rsInfo5.Close
  Set rsInfo5 = Nothing
    
TidyUpAndExit:
  Set cmdChildView = Nothing
  Set objGroup = Nothing
  Set rsInfo1 = Nothing
  Set rsInfo2 = Nothing
  Set rsInfo3 = Nothing
  Set rsInfo4 = Nothing
  Set rsInfo5 = Nothing
  Set rsChildViews = Nothing
  
  ApplyPermissions_ChildTables2 = fOK
  Exit Function

ErrorTrap:
  OutputError "Error Applying Permissions (Child Tables)"
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function LongestRouteToTopLevel(plngTableID As Long) As Integer
  ' Return the given table's longest route to the top-level.
  ' This is used when creating child views.
  Dim iLongestRoute As Integer
  Dim iParentsLongestRoute As Integer
  Dim sSQL As String
  Dim rsParents As DAO.Recordset
  
  iLongestRoute = 0
  
  sSQL = "SELECT parentID" & _
    " FROM tmpRelations" & _
    " WHERE childID = " & Trim$(Str$(plngTableID))
  Set rsParents = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  With rsParents
    Do While (Not .EOF)
      iParentsLongestRoute = LongestRouteToTopLevel(.Fields(0).value)
      
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



Public Function ApplyDatabaseOwnership() As Boolean
  ' Apply any Database Ownership in the SQL Server database.
  On Error GoTo ErrorTrap

  Dim lngPreviousTimeOut As Long
  Dim fOK As Boolean
  Dim sSQL As String
  Dim objGroup As clsSecurityGroup
  Dim rsTemp1 As ADODB.Recordset
  Dim rsTemp2 As ADODB.Recordset
  Dim strXType As String
  Dim lngNextIndex As Long
  Dim astrCommands() As String

  ' Groups are only given write access to the tables in this array unless they have Payroll View Permissions
  Dim astrRevokedTables(7) As String
  
  astrRevokedTables(0) = "ASRSysAccordTransactionData"
  astrRevokedTables(1) = "ASRSysAccordTransactions"
  astrRevokedTables(2) = "ASRSysAccordTransactionWarnings"
  'MH20071107 Fault 5141
  astrRevokedTables(3) = "ASRSysAuditAccess"
  astrRevokedTables(4) = "ASRSysAuditCleardown"
  astrRevokedTables(5) = "ASRSysAuditGroup"
  astrRevokedTables(6) = "ASRSysAuditPermissions"
  astrRevokedTables(7) = "ASRSysAuditTrail"
  'NHRD Prototype Fusion Code  ' This can be deleted when it is confirmed its not needed
'  astrRevokedTables(0) = "ASRSysFusionTransactionData"
'  astrRevokedTables(1) = "ASRSysFusionTransactions"
'  astrRevokedTables(2) = "ASRSysFusionTransactionWarnings"

  Set rsTemp1 = New ADODB.Recordset
  Set rsTemp2 = New ADODB.Recordset

  fOK = True

  sSQL = "GRANT CREATE PROCEDURE TO [ASRSysGroup]"
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords

  sSQL = "GRANT CREATE TABLE TO [ASRSysGroup]"
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords

  If gbEnableUDFFunctions Then
    sSQL = "GRANT CREATE FUNCTION TO [ASRSysGroup]"
    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  End If
  
  ' Get all dbo owned tables/views/procedures etc and grant permission on them
  sSQL = "IF EXISTS (SELECT Name FROM dbo.sysobjects where id = object_id(N'[dbo].[#tmpProtects]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
      " DROP TABLE #tmpProtects" & vbNewLine & _
      " SELECT id INTO #tmpProtects FROM sysprotects" & _
      " INNER JOIN sysusers ON sysprotects.uid = sysusers.uid AND sysusers.name = 'ASRSysGroup'"
  gADOCon.Execute sSQL, , adExecuteNoRecords
  
  ReDim astrCommands(1000)
  lngNextIndex = -1
  
  sSQL = "SELECT DISTINCT sysobjects.name, sysobjects.xtype" & _
    " FROM sysobjects" & _
    " INNER JOIN sysusers ON sysobjects.uid = sysusers.uid" & _
    " LEFT JOIN #tmpProtects ON sysobjects.id = #tmpProtects.id" & _
    " WHERE (((sysobjects.xtype IN ('p','pc')) AND (sysobjects.name LIKE 'sp_asr%' OR sysobjects.name LIKE 'spasr%' OR sysobjects.name LIKE 'spstat%' OR sysobjects.name LIKE 'spsys%')" & _
    "    AND sysobjects.id NOT IN (SELECT id FROM #tmpProtects))" & _
    "    OR ((sysobjects.xtype IN ('fn','tf','if','fs')) " & _
    "AND (sysobjects.name LIKE 'udf_ASR%' OR sysobjects.name LIKE 'udfASR%' OR sysobjects.name LIKE 'udfsys%' OR sysobjects.name LIKE 'udfstat%')))" & _
    "    AND (sysusers.name = 'dbo')"
  
  rsTemp1.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  
  Do While Not rsTemp1.EOF

    strXType = UCase$(Trim$(rsTemp1.Fields(1).value))
    lngNextIndex = lngNextIndex + 1
    If lngNextIndex > UBound(astrCommands) Then ReDim Preserve astrCommands(lngNextIndex + 1000)

    If strXType = "P" Or strXType = "PC" Or strXType = "FN" Or strXType = "FS" Then
      astrCommands(lngNextIndex) = "GRANT EXEC ON [" & rsTemp1(0).value & "] TO [ASRSysGroup]"
    Else
      If strXType = "TF" Or strXType = "IF" Then
        astrCommands(lngNextIndex) = "GRANT SELECT ON [" & rsTemp1(0).value & "] TO [ASRSysGroup]"
      Else
        If InStr(1, Join(astrRevokedTables, ","), rsTemp1(0).value) Then
          astrCommands(lngNextIndex) = "REVOKE SELECT,INSERT,UPDATE,DELETE ON [" & rsTemp1(0).value & "] TO [ASRSysGroup]"
        Else
          astrCommands(lngNextIndex) = "GRANT SELECT,INSERT,UPDATE,DELETE ON [" & rsTemp1(0).value & "] TO [ASRSysGroup]"
        End If
      End If
    End If

    rsTemp1.MoveNext
  Loop
  rsTemp1.Close
  
  'MH20071107 Fault 5141
  'ReDim Preserve astrCommands(lngNextIndex)
  ReDim Preserve astrCommands(lngNextIndex + 2)
  astrCommands(lngNextIndex + 1) = "GRANT SELECT(DateTimeStamp, RecordID, ColumnID) ON ASRSysAuditTrail TO ASRSysGroup"
  astrCommands(lngNextIndex + 2) = "GRANT INSERT ON ASRSysAuditAccess TO ASRSysGroup"
  ' Merge all the procedures into a stored procedure and run that.
  'sSQL = "IF EXISTS (SELECT Name FROM sysobjects WHERE id = object_id('tmpAllocOwnership') AND sysstat & 0xf = 4) DROP PROCEDURE tmpAllocOwnership"
  'gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  DropProcedure "tmpAllocOwnership"

  sSQL = "CREATE PROCEDURE tmpAllocOwnership AS BEGIN " & vbNewLine _
    & "SET NOCOUNT ON" & vbNewLine & Join(astrCommands, vbNewLine) & vbNewLine & "END"
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords

  lngPreviousTimeOut = gADOCon.CommandTimeout
  gADOCon.CommandTimeout = 0
  gADOCon.Execute "tmpAllocOwnership", , adCmdStoredProc
  gADOCon.CommandTimeout = lngPreviousTimeOut

  ' Clear up our temporary stuff
  'sSQL = "DROP PROCEDURE tmpAllocOwnership"
  'gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  DropProcedure "tmpAllocOwnership"

  sSQL = "DROP TABLE #tmpProtects"
  gADOCon.Execute sSQL, , adExecuteNoRecords

  For Each objGroup In gObjGroups.Collection

    sSQL = "SELECT users.name AS result" & _
      " FROM sysusers roles" & _
      " INNER JOIN sysusers users ON roles.uid = users.gid" & _
      " WHERE roles.name = '" & Replace(objGroup.Name, "'", "''") & "'" & _
      " AND users.issqluser = 1"

    rsTemp2.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly

    If Not (rsTemp2.BOF And rsTemp2.EOF) Then
      If objGroup.SystemManager Or _
        objGroup.SecurityManager Then
        sSQL = "sp_addrolemember 'db_owner', '" & Replace(rsTemp2!result, "'", "''") & "'"
      Else
        sSQL = "sp_droprolemember 'db_owner', '" & Replace(rsTemp2!result, "'", "''") & "'"
      End If
      gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
      
    End If

    rsTemp2.Close
  Next objGroup

TidyUpAndExit:
  Set rsTemp1 = Nothing
  Set rsTemp2 = Nothing
  
  Set objGroup = Nothing
  ApplyDatabaseOwnership = fOK
  Exit Function

ErrorTrap:
  OutputError "Error Applying Database Permissions"
  fOK = False
  Resume TidyUpAndExit

End Function





