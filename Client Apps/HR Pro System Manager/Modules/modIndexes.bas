Attribute VB_Name = "modIndexes"
Option Explicit

Private Function CreateIndex(ByVal psTableName As String, pstrIndexName As String, pstrFields As String _
  , bClustered As Boolean, iFillFactor As Integer) As Boolean

  Dim sSQL As String
  Dim bOK As Boolean

  On Error GoTo ErrorTrap
  bOK = True

  ' On the base table
  If glngSQLVersion = 9 Then
    sSQL = "IF EXISTS" & _
      " (SELECT Name" & _
      " FROM sys.indexes WHERE object_id = object_id(N'tbuser_" & psTableName & "')" & _
      " AND name = N'" & pstrIndexName & "')" & _
      " DROP INDEX [" & pstrIndexName & "] ON tbuser_" & psTableName
  Else
    sSQL = "IF EXISTS" & _
      " (SELECT Name" & _
      " FROM sysindexes WHERE id = object_id(N'tbuser_" & psTableName & "')" & _
      " AND name = N'" & pstrIndexName & "')" & _
      " DROP INDEX [tbuser_" & psTableName & "].[" & pstrIndexName & "]"
  End If
  gADOCon.Execute sSQL, , adExecuteNoRecords

  sSQL = "CREATE " & IIf(bClustered, "CLUSTERED", "NONCLUSTERED") & _
    " INDEX [" & pstrIndexName & "] ON [tbuser_" & psTableName & "]" _
    & "(" & pstrFields & " Asc)" _
    & " WITH FILLFACTOR = " & iFillFactor
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords


  ' On its associated view
  If glngSQLVersion = 9 Then
    sSQL = "IF EXISTS" & _
      " (SELECT Name" & _
      " FROM sys.indexes WHERE object_id = object_id(N'" & psTableName & "')" & _
      " AND name = N'" & pstrIndexName & "')" & _
      " DROP INDEX [" & pstrIndexName & "] ON " & psTableName
  Else
    sSQL = "IF EXISTS" & _
      " (SELECT Name" & _
      " FROM sysindexes WHERE id = object_id(N'" & psTableName & "')" & _
      " AND name = N'" & pstrIndexName & "')" & _
      " DROP INDEX [" & psTableName & "].[" & pstrIndexName & "]"
  End If
  gADOCon.Execute sSQL, , adExecuteNoRecords

  sSQL = "CREATE " & IIf(bClustered, "CLUSTERED", "NONCLUSTERED") & _
    " INDEX [" & pstrIndexName & "] ON [" & psTableName & "]" _
    & "(" & pstrFields & " Asc)" _
    & " WITH FILLFACTOR = " & iFillFactor
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords



TidyUpAndExit:
  CreateIndex = bOK
  Exit Function
  
ErrorTrap:
  bOK = False
  GoTo TidyUpAndExit

End Function

Private Function CreatePrimaryIndex(ByVal psTableName As String, bClustered As Boolean) As Boolean

  Dim sSQL As String
  Dim bOK As Boolean
  Dim rstExisting As ADODB.Recordset

  On Error GoTo ErrorTrap
  bOK = True
  Set rstExisting = New ADODB.Recordset

  sSQL = "SELECT name FROM SysObjects WHERE xtype='PK'" & _
    " AND parent_obj = (SELECT ID FROM SysObjects WHERE xtype='U' AND Name = 'tbuser_" & psTableName & "')"
  rstExisting.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    
  If Not (rstExisting.EOF And rstExisting.BOF) Then
    sSQL = "ALTER TABLE [tbuser_" & psTableName & "] DROP CONSTRAINT [" & rstExisting.Fields(0).value & "]"
    gADOCon.Execute sSQL, , adExecuteNoRecords
  End If
    
  sSQL = "ALTER TABLE [tbuser_" & psTableName & "] ADD PRIMARY KEY " & IIf(bClustered, "CLUSTERED", "NONCLUSTERED") & " (ID)"
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords


TidyUpAndExit:
  CreatePrimaryIndex = bOK
  Exit Function
  
ErrorTrap:
  bOK = False
  GoTo TidyUpAndExit

End Function

Public Function CreateChildTableForeignKeys() As Boolean
  
  ' Generate an index for the relationships.
  On Error GoTo ErrorTrap
  
  Dim sTableName As String
  Dim bOK As Boolean
  bOK = True
  
  With recRelEdit
    If Not (.EOF And .BOF) Then
      .MoveFirst
      Do While Not .EOF
        sTableName = GetTableName(.Fields("ChildID").value)
        bOK = CreateIndex(sTableName, "FK_" & .Fields("ParentID").value, "ID_" & .Fields("ParentID").value, False, 80)
        .MoveNext
      Loop
    End If
  End With
  
TidyUpAndExit:
  CreateChildTableForeignKeys = bOK
  Exit Function
  
ErrorTrap:
  bOK = False
  OutputError "Error creating Child Table Foreign Keys"
  Resume TidyUpAndExit

End Function

Public Function CreateHierarchyIndexes() As Boolean

  Dim sSQL As String
  Dim bOK As Boolean
  Dim lngTempID As Long
  Dim sHierarchyTableName As String
  Dim sReportsToColumnName As String

  bOK = True

  ' Get the Hierarchy table ID and Name
  lngTempID = 0
  recModuleSetup.Index = "idxModuleParameter"
  recModuleSetup.Seek "=", gsMODULEKEY_HIERARCHY, gsPARAMETERKEY_HIERARCHYTABLE
  If Not recModuleSetup.NoMatch Then
    lngTempID = recModuleSetup!parametervalue
    recTabEdit.Index = "idxTableID"
    recTabEdit.Seek "=", lngTempID
    If Not recTabEdit.NoMatch Then
      sHierarchyTableName = recTabEdit!TableName
    Else
      sHierarchyTableName = vbNullString
    End If
  Else
    sHierarchyTableName = vbNullString
  End If

  ' Get the ReportsTo column variable
  lngTempID = 0
  recModuleSetup.Index = "idxModuleParameter"
  recModuleSetup.Seek "=", gsMODULEKEY_HIERARCHY, gsPARAMETERKEY_REPORTSTO
  If Not recModuleSetup.NoMatch Then
    lngTempID = recModuleSetup!parametervalue
    recColEdit.Index = "idxColumnID"
    recColEdit.Seek "=", lngTempID
    If Not recColEdit.NoMatch Then
      sReportsToColumnName = recColEdit!ColumnName
    Else
      sReportsToColumnName = vbNullString
    End If
  Else
    sReportsToColumnName = vbNullString
  End If

  If Len(sReportsToColumnName) > 0 And Len(sHierarchyTableName) > 0 Then
    bOK = CreateIndex(sHierarchyTableName, "IDX_Hierarchy_Reports_To_Column", sReportsToColumnName, False, 90)
  End If

TidyUpAndExit:
  CreateHierarchyIndexes = bOK
  Exit Function
  
ErrorTrap:
  bOK = False
  OutputError "Error creating Hierarchy foreign keys"
  Resume TidyUpAndExit

End Function



Public Function CreatePrimaryKeysForTables() As Boolean
  
  ' Generate an index for the relationships.
  On Error GoTo ErrorTrap
  
  Dim sTableName As String
  Dim bOK As Boolean
  bOK = True
  
  With recTabEdit
    If Not (.EOF And .BOF) Then
      .MoveFirst
      Do While Not .EOF
        ' AE20080303 Fault #12966
        If Not CBool(.Fields("Deleted").value) Then
          sTableName = .Fields("TableName").value
          bOK = CreatePrimaryIndex(sTableName, True)
        End If
        .MoveNext
      Loop
    End If
  End With
 
  
TidyUpAndExit:
  CreatePrimaryKeysForTables = bOK
  Exit Function
  
ErrorTrap:
  bOK = False
  OutputError "Error creating Child Table Foreign Keys"
  Resume TidyUpAndExit

End Function


' TO DO?

'Public Function GetCustomIndexes(ByRef astrIndexes() As String) As Boolean
'
'  Dim iLoop As Integer
'  Dim sSQL As String
'
'  sSQL = "SELECT MAX(I.Index_ID) FROM sys.indexes i" & _
'            " INNER JOIN sys.sysindexkeys k ON i.index_id = k.indid AND k.id = i.[Object_id]" & _
'            " WHERE OBJECT_NAME(i.[Object_ID]) = 'Personnel_Records'"
'
'
'  ReDim GetCustomIndexes(2, 1)
'  iLoop = 1
'  astrIndexes(iLoop, 0) = "indexname1"
'  astrIndexes(iLoop, 1) = "fieldss"
'  astrIndexes(iLoop, 2) = "asc,desc"
'  astrIndexes(iLoop, 3) = "cluster Y N"
'
'End Function
'
'Public Function GenerateCustomIndexes(ByRef astrIndexes() As String) As Boolean
'
'
'
'End Function

Public Function OtherBits() As Boolean

'global variable table - maybe lookup tables in general?

End Function

Public Function IndexPrimaryTableOrders() As Boolean

' needs debugging!

'DECLARE @CursorOuter as CURSOR
'DECLARE @CursorInner as CURSOR
'
'DECLARE @iOrderID as integer
'DECLARE @TableName nvarchar(1000)
'DECLARE @OrderName nvarchar(1000)
'DECLARE @ColumnName nvarchar(1000)
'DECLARE @Ascending integer
'DECLARE @IndexString nvarchar(4000)
'
'DECLARE @sGenerate nvarchar(4000)
'DECLARE @bGenerate bit
'
'SET @bGenerate = 0
'
'SET @CursorOuter = CURSOR LOCAL FAST_FORWARD READ_ONLY FOR SELECT o.[Name], o.[OrderID], t.[TableName] FROM ASRSysOrders o
'  INNER JOIN ASRSysTables t ON o.TableID = t.TableID
'  ORDER BY o.[TableID], o.[OrderID]
'
'OPEN @CursorOuter
'
'FETCH NEXT FROM @CursorOuter INTO @OrderName, @iOrderID, @TableName
'WHILE (@@fetch_status = 0)
'  BEGIN
'
'    SET @CursorInner = CURSOR LOCAL FAST_FORWARD READ_ONLY FOR
'      SELECT c.ColumnName, i.Ascending FROM ASRSysOrders o
'      INNER JOIN ASRSysOrderItems i ON o.OrderID = i.OrderID and i.Type = 'O'
'      INNER JOIN ASRSysColumns c ON i.ColumnID = c.ColumnID AND c.TableID = o.TableID
'      INNER JOIN ASRSysTables t ON t.TableID = o.TableID
'      WHERE o.[OrderID] = @iOrderID
'      ORDER BY o.[TableID], o.[OrderID], i.[Sequence]
'
'    OPEN @CursorInner
'    SET @IndexString = ''
'
'    FETCH NEXT FROM @CursorInner INTO @ColumnName, @Ascending
'    WHILE (@@fetch_status = 0)
'    BEGIN
'      IF LEN(@IndexString) = 0 SET @IndexString = @ColumnName
'        ELSE SET @IndexString = @IndexString + ',' + @ColumnName
'      FETCH NEXT FROM @CursorInner INTO @ColumnName, @Ascending
'    End
'
'    -- sql 2005
'    SET @sGenerate = 'IF EXISTS(SELECT Name FROM sys.indexes WHERE object_id = object_id(N''' + @TableName + ''')
'            AND name = ''IDXOrder_' + @OrderName + ''')
'      DROP INDEX [IDXOrder_' + @OrderName + '] ON ' + @TableName
'
'    -- sql 2000
'    SET @sGenerate = 'IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N''' + @TableName + ''')
'            AND name = ''IDXOrder_' + @OrderName + ''')
'      Drop Index ' + @TableName + '.[IDXOrder_' + @OrderName + ']'
'
'    Print 'Dropping --' + @TableName + '.' + @OrderName
'    EXECUTE sp_executesql @sGenerate
'
'    IF @bGenerate = 1
'    BEGIN
'      SET @sGenerate = 'CREATE NONCLUSTERED INDEX [IDXOrder_' + @OrderName + '] ON ['
'       + @TableName + '] (' + @IndexString + ')'
'      Print 'Generating --' + @TableName + '.' + @OrderName
'      EXECUTE sp_executesql @sGenerate
'    End
'
'    FETCH NEXT FROM @CursorOuter INTO @OrderName, @iOrderID, @TableName
'  End



End Function
