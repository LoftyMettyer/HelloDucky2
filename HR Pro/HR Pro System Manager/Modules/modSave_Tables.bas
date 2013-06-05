Attribute VB_Name = "modSave_Tables"
Option Explicit


'Private Function SaveTables(ByRef psErrMsg As String, pfRefreshDatabase As Boolean, pavOldColumns As Variant) As Boolean
'Private Function SaveTables(pfRefreshDatabase As Boolean, pavOldColumns As Variant) As Boolean
Public Function SaveTables(pfRefreshDatabase As Boolean, mfrmUse As frmUsage) ', pavOldColumns As Variant) As Boolean
  ' Save the new or modified Table definitions.
  On Error GoTo ErrorTrap
  
  Dim objTable As Table
  Dim fOK As Boolean
  Dim fCreateMaxIDStoredProcedure As Boolean
  Dim lngRecordCount As Long
  
  fOK = True
  fCreateMaxIDStoredProcedure = False
  
  With recTabEdit
    .Index = "idxTableID"
    If Not (.BOF And .EOF) Then
      .MoveFirst
      lngRecordCount = .RecordCount
    End If
    Do While fOK And Not .EOF
      
      'Do deleted ones first
      If !Deleted Then
        Set objTable = New Table
        objTable.TableID = !TableID
        Set mfrmUse = New frmUsage
        mfrmUse.ResetList
        If objTable.TableIsUsed(mfrmUse) Then
          gobjProgress.Visible = False
          Screen.MousePointer = vbNormal
          Select Case !TableType
            Case TableTypes.iTabParent
              mfrmUse.ShowMessage !TableName & " Table", "The table cannot be deleted as the table is used by the following:", UsageCheckObject.Table
            Case TableTypes.iTabChild
              mfrmUse.ShowMessage !TableName & " Child Table", "The table cannot be deleted as the table is used by the following:", UsageCheckObject.ChildTable
            Case TableTypes.iTabLookup
              mfrmUse.ShowMessage !TableName & " Lookup Table", "The table cannot be deleted as the table is used by the following:", UsageCheckObject.LookupTable
          End Select
          
          fOK = False
        End If
        UnLoad mfrmUse
        Set mfrmUse = Nothing
        
        gobjProgress.Visible = True
        
        If fOK Then
          OutputCurrentProcess2 "Deleting " & recTabEdit!TableName, lngRecordCount
          gobjProgress.UpdateProgress2
          fOK = TableDelete
          fCreateMaxIDStoredProcedure = True
        Else
          Exit Do
        End If
        
      End If

      fOK = fOK And Not gobjProgress.Cancelled
      .MoveNext
    Loop
  
  
    .Index = "idxTableID"
    If Not (.BOF And .EOF) Then
      .MoveFirst
      lngRecordCount = .RecordCount
    End If
    Do While fOK And Not .EOF
      
      'Now do new and changed ones
      If Not !Deleted Then
        If !New Then
          OutputCurrentProcess2 recTabEdit!TableName, lngRecordCount
          gobjProgress.UpdateProgress2
          fOK = TableNew
          fCreateMaxIDStoredProcedure = True
          
        ElseIf !Changed Or pfRefreshDatabase Then
          OutputCurrentProcess2 recTabEdit!TableName, lngRecordCount
          gobjProgress.UpdateProgress2
          fOK = TableSave(mfrmUse)
          fCreateMaxIDStoredProcedure = True
        End If
        
      End If

      fOK = fOK And Not gobjProgress.Cancelled
      .MoveNext
    Loop
  
  
  End With
  
  ' JPD20030313 Fault 5159
  If fOK And fCreateMaxIDStoredProcedure Then
    fOK = CreateMaxIDStoredProcedure
  End If

TidyUpAndExit:
  SaveTables = fOK
  Exit Function
  
ErrorTrap:
  OutputError "Error saving table definitions"
  fOK = False
  Resume TidyUpAndExit

End Function


Private Function TableDelete() As Boolean
  ' Delete the table on the server (and all of the associated records in the other tables)
  ' as defined by the current record in the local Tables table.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim sSQL As String
  Dim strTableID As String
  Dim sOriginalName As String

  fOK = True
  strTableID = CStr(recTabEdit!TableID)

  ' Delete table definition from Tables
  sSQL = "DELETE FROM ASRSysTables" & _
    " WHERE tableID=" & strTableID
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords

  ' Delete column control values for each of the columns in the table.
  sSQL = "DELETE FROM ASRSysColumnControlValues" & _
    " WHERE columnID IN (SELECT columnID FROM ASRSysColumns" & _
    " WHERE tableID=" & strTableID & ")"
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords

  ' Delete diary links for each of the columns in the table.
  sSQL = "DELETE FROM ASRSysDiaryLinks" & _
    " WHERE columnID IN (SELECT columnID FROM ASRSysColumns" & _
    " WHERE tableID=" & strTableID & ")"
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords

  ' Delete summary fields for the table.
  sSQL = "DELETE FROM ASRSysSummaryFields" & _
    " WHERE historyTableID=" & strTableID
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords

  ' Delete table validations for the table.
  sSQL = "DELETE FROM [ASRSysTableValidations]" & _
    " WHERE [TableID]=" & strTableID
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords

  ' Delete column definitions for this table from Columns
  sSQL = "DELETE FROM ASRSysColumns" & _
    " WHERE tableID=" & strTableID
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
  
  
  sSQL = "DELETE FROM ASRSysEmailLinksRecipients " & _
    " WHERE ASRSysEmailLinksRecipients.LinkID IN " & _
    "(SELECT LinkID FROM ASRSysEmailLinks " & _
    " WHERE ASRSysEmailLinks.TableID = " & strTableID & ")"
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
  sSQL = "DELETE FROM ASRSysEmailLinksColumns " & _
    " WHERE ASRSysEmailLinksColumns.LinkID IN " & _
    "(SELECT LinkID FROM ASRSysEmailLinks " & _
    " WHERE ASRSysEmailLinks.TableID = " & strTableID & ")"
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
  sSQL = "DELETE FROM ASRSysLinkContent " & _
    " WHERE ASRSysLinkContent.ContentID IN " & _
    "(SELECT SubjectContentID FROM ASRSysEmailLinks " & _
    " WHERE ASRSysEmailLinks.TableID = " & strTableID & ")"
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
  sSQL = "DELETE FROM ASRSysLinkContent " & _
    " WHERE ASRSysLinkContent.ContentID IN " & _
    "(SELECT BodyContentID FROM ASRSysEmailLinks " & _
    " WHERE ASRSysEmailLinks.TableID = " & strTableID & ")"
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
  sSQL = "DELETE FROM ASRSysEmailLinks" & _
    " WHERE ASRSysEmailLinks.TableID = " & strTableID
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
  
  
  sSQL = "DELETE FROM ASRSysOutlookLinks" & _
    " WHERE tableID=" & strTableID
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords

  If recTabEdit!Deleted Then
    ' NB. Deleting the ASRSysWorkflowTriggeredLinks record will NO LONGER trigger the deletion of related
    ' records in ASRSysWorkflowQueue and ASRSysWorkflowQueueColumns, so we need to do it manually here.
    ' NB again. ONly clear out the queue if the table has really been deleted, not if the tables been dropped and recreatred due
    ' to modifcation, version update or shift-save.
    gADOCon.Execute "DELETE FROM ASRSysWorkflowQueue" & _
      " WHERE linkID IN (SELECT linkID FROM ASRSysWorkflowTriggeredLinks WHERE tableID = " & strTableID & ")", , _
      adCmdText + adExecuteNoRecords
  End If
  
  sSQL = "DELETE FROM ASRSysWorkflowTriggeredLinks" & _
    " WHERE tableID=" & strTableID
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords

  ' Drop the table.
  sOriginalName = "tbuser_" & recTabEdit!OriginalTableName
  If Database.TableExists(sOriginalName) Then
    
    sSQL = "DROP VIEW " & recTabEdit!OriginalTableName
    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
    
    sSQL = "DROP TABLE " & sOriginalName
    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
  End If


TidyUpAndExit:
  TableDelete = fOK
  Exit Function

ErrorTrap:
  OutputError "Error deleting table"
  fOK = False
  Resume TidyUpAndExit

End Function

Private Function TableNew() As Boolean
  ' Create a new table on the server as defined by the current record in the local Tables table.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim iColumn As Integer
  Dim iDataType As Integer
  Dim lngTableID As Long
  Dim sSQL As String
  Dim sName As String
  Dim sTableName As String
  Dim sPhysicalTableName As String
  Dim sColCreate As String
  Dim rsColumns As ADODB.Recordset
  'Dim rsDiaryLinks As ADODB.Recordset
  'Dim rsEmailLinks As ADODB.Recordset
  'Dim rsEmailRecipients As ADODB.Recordset
  Dim rsControlValues As ADODB.Recordset
  Dim objTable As Table
  Dim objSummaryField As cSummaryField
  Dim objTableValidation As clsTableValidation
  Dim bEmbedded As Boolean
  Dim sTableCreate As HRProSystemMgr.cStringBuilder
  Dim sCreateView As HRProSystemMgr.cStringBuilder

  ' Check that the table has a default order defined.
  fOK = (recTabEdit!defaultOrderID > 0)

  Set sTableCreate = New HRProSystemMgr.cStringBuilder
  Set sCreateView = New HRProSystemMgr.cStringBuilder
  Set rsColumns = New ADODB.Recordset
  'Set rsDiaryLinks = New ADODB.Recordset
  
  Set rsControlValues = New ADODB.Recordset


  OpenDiaryRecordsets
  OpenEmailRecordsets


  If Not fOK Then
    MsgBox "A primary order must be defined for the '" & recTabEdit!TableName & "' table.", _
      vbCritical + vbOKOnly, App.Title
  Else
    lngTableID = recTabEdit!TableID
    sTableName = recTabEdit!TableName
    sPhysicalTableName = "tbuser_" & sTableName

    'MH20000728 Added Email
    sSQL = "INSERT INTO ASRSysTables (" & _
             "tableID, " & _
             "tableName, " & _
             "tableType, " & _
             "defaultOrderID, " & _
             "recordDescExprID, " & _
             "DefaultEmailID, " & _
             "AuditInsert, AuditDelete, " & _
             "ManualSummaryColumnBreaks, IsRemoteView) " & _
           "VALUES (" & _
             lngTableID & ", '" & _
             sTableName & "', " & _
             recTabEdit!TableType & ", " & _
             recTabEdit!defaultOrderID & "," & _
             recTabEdit!RecordDescExprID & "," & _
             IIf(IsNull(recTabEdit!DefaultEmailID), 0, recTabEdit!DefaultEmailID) & ", " & _
             IIf(recTabEdit!AuditInsert = True, 1, 0) & ", " & _
             IIf(recTabEdit!AuditDelete = True, 1, 0) & ", " & _
             IIf(recTabEdit!ManualSummaryColumnBreaks, 1, 0) & "," & IIf(recTabEdit!IsRemoteView, 1, 0) & ")"
    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords

    ' Add the Summary Field values.
    Set objTable = New Table
    objTable.TableID = lngTableID
    If objTable.ReadTable Then
    
      ' Commit the summary objects
      For Each objSummaryField In objTable.SummaryFields
        sSQL = "INSERT INTO ASRSysSummaryFields (ID, historyTableID, parentColumnID, sequence, startOfGroup,StartOfColumn)" & _
          " VALUES(" & objSummaryField.id & ", " & _
          objSummaryField.HistoryTableID & ", " & _
          objSummaryField.SummaryColumnID & ", " & _
          objSummaryField.Sequence & ", " & _
          IIf(objSummaryField.StartOfGroup, 1, 0) & ", " & _
          IIf(objSummaryField.StartOfColumn, 1, 0) & ")"
        gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
      Next objSummaryField
      Set objSummaryField = Nothing


      ' Commit the table validations
      For Each objTableValidation In objTable.TableValidations
        sSQL = "INSERT INTO [ASRSysTableValidations] ([ValidationID], [TableID], [Type]" & _
          ",[EventStartDateColumnID], [EventStartSessionColumnID], [EventEndDateColumnID], [EventEndSessionColumnID], [FilterID], [EventTypeColumnID] " & _
          ",[Severity], [Message])" & _
          " VALUES( " & objTableValidation.ValidationID & ", " & _
          objTableValidation.TableID & ", " & _
          objTableValidation.ValidationType & ", " & _
          objTableValidation.EventStartdateColumnID & ", " & _
          objTableValidation.EventStartSessionColumnID & ", " & _
          objTableValidation.EventEnddateColumnID & ", " & _
          objTableValidation.EventEndSessionColumnID & ", " & _
          objTableValidation.FilterID & ", " & _
          objTableValidation.EventTypeColumnID & ", " & _
          objTableValidation.Severity & ", " & _
          "'" & objTableValidation.Message & "')"
        gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
      Next
      Set objTableValidation = Nothing
     
      
    End If
    Set objTable = Nothing

    With recTabEdit
      .Index = "idxTableID"
      .Seek "=", lngTableID
    End With

    ' Open the server's column details table.
    rsColumns.Open "ASRSysColumns", gADOCon, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect
    rsControlValues.Open "ASRSysColumnControlValues", gADOCon, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect
    'rsDiaryLinks.Open "ASRSysDiaryLinks", gADOCon, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect
    'rsEmailLinks.Open "ASRSysEmailLinks", gADOCon, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    'rsEmailRecipients.Open "ASRSysEmailLinksRecipients", gADOCon, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect


    ' Add the column definitions for this table to the 'create table' SQL command string.
    recColEdit.Index = "idxName"
    recColEdit.Seek ">=", lngTableID

    If Not recColEdit.NoMatch Then
      ' For each column definition ...
      Do While (Not recColEdit.EOF) And fOK

        ' Stop looping when we've done all of the columns for the current table.
        If recColEdit!TableID <> lngTableID Then
          Exit Do
        End If

        ' Don't add deleted columns (obviously).
        If Not recColEdit!Deleted Then
          ' Add the column details to the server's ASRSysColumns table.
          With rsColumns
            .AddNew
            For iColumn = 0 To .Fields.Count - 1
              sName = .Fields(iColumn).Name

              'MH20011206
              'Both Balfour Beatty and Trade Team have had problems saving changes in
              'System Manager due to the column "msrepl_synctran_ts" which is added
              'to ASRSysColumns (and others?) for replication.  I have put this line
              'of code in as a temporary measure so that they can save changes but we
              'need to look into how this will effect the replication process.
              '(Trade Team were getting: "Incorrect type for Parameter").
              If Not (sName = "locked" Or sName = "lastupdatedby" Or sName = "lastupdated") Then
                If sName <> "msrepl_synctran_ts" Then
                  If Not IsNull(recColEdit.Fields(sName).value) Then
                    .Fields(iColumn).value = recColEdit.Fields(sName).value
                  End If
                End If
              End If

            Next iColumn
            .Update
          End With

          ' Add the columns control values.
          With recContValEdit
            If Not (.BOF And .EOF) Then
              .MoveFirst

              Do While Not .EOF
                If !ColumnID = recColEdit!ColumnID Then
                  rsControlValues.AddNew
                  rsControlValues!ColumnID = !ColumnID
                  rsControlValues!value = !value
                  If Not IsNull(!Sequence) Then
                    rsControlValues!Sequence = !Sequence
                  End If
                  rsControlValues.Update
                End If

                .MoveNext
              Loop
            End If
          End With


          SaveDiaryLinksForColumn recColEdit!ColumnID


          ' Add the column details to the SQL command string.
          iDataType = recColEdit.Fields("dataType").value
          bEmbedded = IIf(IsNull(recColEdit.Fields("OLEType").value), False, recColEdit.Fields("OLEType").value = 2)

          If ((iDataType = dtBINARY) Or (iDataType = dtVARBINARY) Or (iDataType = dtLONGVARBINARY)) And Not bEmbedded Then
            sColCreate = GetColCreateString(recColEdit!ColumnName, dtVARCHAR, 255, 0, False)
          ElseIf ((iDataType = dtBINARY) Or (iDataType = dtVARBINARY) Or (iDataType = dtLONGVARBINARY)) And bEmbedded Then
            sColCreate = GetColCreateString(recColEdit!ColumnName, dtLONGVARBINARY, 255, 0, False)
          ElseIf (iDataType = dtLONGVARCHAR) Then
            sColCreate = GetColCreateString(recColEdit!ColumnName, dtVARCHAR, 14, 0, recColEdit!MultiLine)
          Else
            sColCreate = GetColCreateString(recColEdit!ColumnName, iDataType, recColEdit!Size, recColEdit!Decimals, recColEdit!MultiLine)
          End If

          fOK = (LenB(sColCreate) <> 0)

          If fOK Then
            sTableCreate.Append IIf(sTableCreate.Length <> 0, ", ", vbNullString) & sColCreate
            sCreateView.Append IIf(sCreateView.Length <> 0, ", ", vbNullString) & recColEdit!ColumnName

            ' If this column is the record ID, then make it the primary key.
            If (recColEdit!ColumnType = giCOLUMNTYPE_SYSTEM) And (recColEdit!ColumnName = "ID") Then
              sTableCreate.Append " NOT NULL IDENTITY(1,1)"
            End If

            ' Check if default required.
            If LenB(Trim(recColEdit!DefaultValue)) <> 0 Then
              Select Case iDataType
                Case dtVARCHAR, dtLONGVARCHAR
                  'JPD 20041012 Fault 9293
                  'sSQL = sSQL & " DEFAULT '" & recColEdit!DefaultValue & "'"
                  sTableCreate.Append " DEFAULT '" & Replace(recColEdit!DefaultValue, "'", "''") & "'"
                Case dtINTEGER, dtNUMERIC
                  sTableCreate.Append " DEFAULT " & recColEdit!DefaultValue
                Case dtBIT
                  sTableCreate.Append " DEFAULT " & IIf(recColEdit!DefaultValue = "TRUE", "1", "0")
              End Select
            Else
              If iDataType = dtBIT Then
                sTableCreate.Append " DEFAULT 0"
              End If
            End If

          End If
        End If

        'Get next column definition
        recColEdit.MoveNext
      Loop
    End If

    ' Add a timestamp column.
    sTableCreate.Append IIf(sTableCreate.Length <> 0, ", ", vbNullString) & "TimeStamp"
    sCreateView.Append IIf(sCreateView.Length <> 0, ", ", vbNullString) & "TimeStamp"

    ' Complete the 'create table' SQL command string.
    sSQL = "CREATE TABLE dbo." & sPhysicalTableName & " (" & sTableCreate.ToString & ")"
    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords

    sSQL = "CREATE VIEW dbo." & sTableName & " WITH SCHEMABINDING AS SELECT " & sCreateView.ToString & " FROM dbo." & sPhysicalTableName
    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords

    ' Add an index
    sSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[" & sTableName & "]')" & _
          "AND name = N'IDX_ID')" & vbNewLine & _
          "CREATE UNIQUE CLUSTERED INDEX [IDX_ID] ON [dbo].[" & sTableName & "] ([ID] ASC)"
    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords

    ' Close recordsets.
    rsColumns.Close
    rsControlValues.Close
    'rsDiaryLinks.Close
    'rsEmailLinks.Close
    'rsEmailRecipients.Close

    If fOK Then
      fOK = CreateInsertStoredProcedure(lngTableID, sTableName)
    End If
    If fOK Then
      fOK = CreateIntInsertStoredProcedure(lngTableID, sTableName)
    End If

    If fOK Then
      fOK = CreateUpdateStoredProcedure(lngTableID, sTableName)
    End If
    If fOK Then
      fOK = CreateIntUpdateStoredProcedure(lngTableID, sTableName)
    End If

    If fOK Then
      fOK = CreateDeleteStoredProcedure(lngTableID, sTableName)
    End If

    If fOK Then
      fOK = CreateRecordAmendedStoredProcedure(lngTableID, sTableName)
    End If

    If fOK Then
      fOK = SaveEmailLinks(lngTableID)
    End If

    If fOK Then
      fOK = SaveOutlookLinks(lngTableID)
    End If

    If fOK Then
      fOK = SaveWorkflowLinks(lngTableID)
    End If

  End If

TidyUpAndExit:
  ' Disassociate object variables.
  Set rsColumns = Nothing
  Set rsControlValues = Nothing
  'Set rsDiaryLinks = Nothing
  
  CloseDiaryRecordsets
  CloseEmailRecordsets
  
  TableNew = fOK
  Exit Function

ErrorTrap:
  'MsgBox "Unable to create table " & sTableName & "." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, App.ProductName
  OutputError "Error creating table '" & sTableName & "'"
  fOK = False
  Resume TidyUpAndExit

End Function

'Private Function TableSave(pavOldColumns As Variant) As Boolean
Private Function TableSave(mfrmUse As frmUsage) As Boolean
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim iDataType As Integer
  Dim lngTableID As Long
  Dim dblMaxValue As Double
  Dim sSQL As String
  Dim sName As String
  Dim sTempName As String
  Dim sTableName As String
  Dim sPhysicalTableName As String
  Dim sOriginalTableName As String
  Dim rsTriggers As New ADODB.Recordset
  Dim rsTriggerDefn As New ADODB.Recordset
  Dim sTriggerDefn As String
  Dim asTriggers() As String
  Dim asTriggerErrors() As String
  Dim sMessage As String
  Dim lngRelocateTableID As Long
  Dim sRelocateColumnName As String
  Dim sColumnName As String
  Dim objColumn As Column
  Dim lngNextIdentitySeed As Long

  Dim asValueList() As String
  Dim asColumnList() As String
  Dim iColumnList As Integer
  Dim strErrorMessage As String
  Dim sAreaInCode As String

  'Dim sConnect As String
  'Dim blnReconnect As Boolean

  'sConnect = gADOCon.ConnectionString



  sAreaInCode = ""
  ' Get the current table ID and name.
  lngTableID = recTabEdit!TableID

  'MH20010321
  ' Extract data from this table into a temporary table.
  sTableName = recTabEdit!TableName
  sPhysicalTableName = "tbuser_" + LCase(sTableName)
  
  sTempName = Database.GetTempTableName("Tmp_" & sTableName)
  sOriginalTableName = "tbuser_" + recTabEdit!OriginalTableName
  fOK = Not (sTempName = vbNullString)

  If fOK Then
    sSQL = "SELECT * INTO " & sTempName & _
      " FROM " & sOriginalTableName
    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  End If

  ' JPD20030110 Fault 4162
  ' Read any custom triggers so they can be recreated after the table has been recreated.
  ReDim asTriggers(0)
  ReDim asTriggerErrors(0)

  If glngSQLVersion >= 8 Then
    ' Get any non-HR Pro generted triggers.
    sSQL = "SELECT triggerobjects.name" & _
      " FROM sysobjects tableobjects" & _
      " LEFT OUTER JOIN sysobjects triggerobjects ON tableobjects.id = triggerobjects.parent_obj" & _
      " WHERE tableobjects.name = '" & sTableName & "'" & _
      " AND triggerobjects.xtype = 'TR'" & _
      " AND triggerobjects.name <> 'INS_" & sTableName & "'" & _
      " AND triggerobjects.name <> 'UPD_" & sTableName & "'" & _
      " AND triggerobjects.name <> 'DEL_" & sTableName & "'"

    rsTriggers.Open sSQL, gADOCon, adOpenDynamic, adLockReadOnly

    ' Loop through the custom triggers.
    Do While Not rsTriggers.EOF
      sTriggerDefn = vbNullString

      ' Get the script for the custom trigger.
      sSQL = "sp_helptext '" & rsTriggers!Name & "'"

      rsTriggerDefn.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly

      If rsTriggerDefn.Fields.Count = 0 Then
        ' Trigger code could not be read. Might be encrypted.
        ReDim Preserve asTriggerErrors(UBound(asTriggerErrors) + 1)
        asTriggerErrors(UBound(asTriggerErrors)) = rsTriggers!Name
      Else
        Do While Not rsTriggerDefn.EOF
          sTriggerDefn = sTriggerDefn & rsTriggerDefn!Text

          rsTriggerDefn.MoveNext
        Loop
        rsTriggerDefn.Close

        ReDim Preserve asTriggers(UBound(asTriggers) + 1)
        asTriggers(UBound(asTriggers)) = sTriggerDefn
      End If

      rsTriggers.MoveNext
    Loop
    rsTriggers.Close
  End If

  If UBound(asTriggerErrors) > 0 Then
    sMessage = "The following custom trigger" & IIf(UBound(asTriggerErrors) = 1, vbNullString, "s") & _
      " could not be read, possibly due to encryption." & vbNewLine & _
      "Continuing to save changes will delete this trigger." & vbNewLine & vbNewLine

    For iLoop = 1 To UBound(asTriggerErrors)
      sMessage = sMessage & vbTab & asTriggerErrors(iLoop) & vbNewLine
    Next iLoop

    fOK = (OutputMessage(sMessage & vbNewLine & "Continue saving changes ?") = vbYes)
  End If

  ' Whats the next identity insert value
  lngNextIdentitySeed = GetNextIdentitySeed(sOriginalTableName)



  If fOK Then
    ' Drop the table.
    fOK = TableDelete
  End If

  If fOK Then
    ' Recreate the table.
    fOK = TableNew
  End If

  If fOK Then
    'Build list of columns with which to re-populate this table.
    ReDim asValueList(0)
    ReDim asColumnList(0)
    iColumnList = -1

    With recColEdit
      .Index = "idxName"
      .Seek ">=", lngTableID

      If Not .NoMatch Then
        Do While (Not .EOF)

          If !TableID <> lngTableID Then
            Exit Do
          End If

          ' Ignore deleted columns (obviously).
          If Not !Deleted Then

            ' JDM - Fault 10544 - 10/11/05 - Damn null values!
            sName = IIf(IsNull(.Fields("OriginalColumnName").value), .Fields("ColumnName").value, .Fields("OriginalColumnName").value)
            iDataType = IIf(IsNull(.Fields("OriginalDataType").value), .Fields("DataType").value, .Fields("OriginalDataType").value)
            sColumnName = !ColumnName

            If !New And (!DataType = dtBIT) Then
              ' Initialise new logic columns (required by SQL Server).
              iColumnList = iColumnList + 1
              If iColumnList > UBound(asColumnList) Then
                ReDim Preserve asColumnList(iColumnList + 100)
                ReDim Preserve asValueList(iColumnList + 100)
              End If
              asColumnList(iColumnList) = sColumnName
              asValueList(iColumnList) = "0"


            ElseIf Not !New Then
              ' Copy data from the old database structure to the new one.
              ' NB. Only try to copy data if the columns have compatible data types.
              ' Get the old column definition.

              iColumnList = iColumnList + 1
              If iColumnList > UBound(asColumnList) Then
                ReDim Preserve asColumnList(iColumnList + 100)
                ReDim Preserve asValueList(iColumnList + 100)
              End If

              If !DataType = iDataType Then
                asColumnList(iColumnList) = sColumnName

                Select Case !DataType
                  ' Convert character.
                  Case dtVARCHAR
                    asValueList(iColumnList) = "CONVERT(varchar(MAX)," & sName & ")"

                  Case dtLONGVARCHAR
                    asValueList(iColumnList) = "CONVERT(varchar(" & Trim(Str(14)) & ")," & sName & ")"

                  ' Convert numeric.
                  Case dtNUMERIC
                    ' Ensure that we don't try to copy any out of range data into the columns.
                    dblMaxValue = 10 ^ (!Size - !Decimals)
                    sSQL = "UPDATE " & sTempName & _
                      " SET " & sName & " = 0" & _
                      " WHERE " & sName & " >= " & Format(dblMaxValue, "0") & _
                      " OR " & sName & " <= -" & Format(dblMaxValue, "0")
                    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
                    asValueList(iColumnList) = "CONVERT(numeric(" & Trim(Str(!Size)) & "," & Trim(Str(!Decimals)) & "), " & sName & ")"

                  Case Else
                    asValueList(iColumnList) = sName

                End Select
              Else
                asColumnList(iColumnList) = sColumnName
                Select Case !DataType
                  ' Convert data into character if possible.
                  Case dtVARCHAR, dtLONGVARCHAR
                    If (iDataType = dtTIMESTAMP) Or _
                      (iDataType = dtINTEGER) Or _
                      (iDataType = dtNUMERIC) Or _
                      (iDataType = dtBIT) Then
                        asValueList(iColumnList) = "CONVERT(varchar(MAX), " & sName & ")"
                    Else
                      asValueList(iColumnList) = "''"
                    End If

                  ' Convert data into integer if possible.
                  Case dtINTEGER
                    If (iDataType = dtNUMERIC) Or (iDataType = dtBIT) Then
                      asValueList(iColumnList) = "CONVERT(int, " & sName & ")"
                    Else
                      asValueList(iColumnList) = "0"
                    End If

                  ' Convert data into numeric if possible.
                  Case dtNUMERIC
                    If (iDataType = dtINTEGER) Or (iDataType = dtBIT) Then
                      asValueList(iColumnList) = "CONVERT(numeric(" & Trim(Str(!Size)) & "," & Trim(Str(!Decimals)) & "), " & sName & ")"
                    Else
                      asValueList(iColumnList) = "0"
                    End If

                  ' Cannot convert any other datatype into bit, but we need to initialise it.
                  Case dtBIT
                    asValueList(iColumnList) = "0"

                  Case Else   'For example dates!!
                    asValueList(iColumnList) = "null"

                End Select
              End If

            End If

          Else
            'Column is deleted - check for usage.
            lngRelocateTableID = !TableID
            sRelocateColumnName = !ColumnName

            Set objColumn = New Column
            objColumn.ColumnID = !ColumnID
            Set mfrmUse = New frmUsage
            mfrmUse.ResetList
            If objColumn.ColumnIsUsed(mfrmUse) Then
              gobjProgress.Visible = False
              Screen.MousePointer = vbNormal
              mfrmUse.ShowMessage GetTableName(!TableID) & "." & !ColumnName & " Column", "The column cannot be deleted as the column is used by the following:", UsageCheckObject.Column
              fOK = False
            End If
            UnLoad mfrmUse
            Set mfrmUse = Nothing

            gobjProgress.Visible = True

            If Not fOK Then
              Exit Do
            End If

            'JPD 20040227 Fault 8163
            ' Need to relocate the record in the recColEdit recordset
            ' as the ColumnIsUsed function may have changed the current record.
            .Index = "idxName"
            .Seek "=", lngRelocateTableID, sRelocateColumnName, True
          End If

          .MoveNext
        Loop
      End If
    End With
  End If

  ' Trim the column and value list
  ReDim Preserve asColumnList(iColumnList)
  ReDim Preserve asValueList(iColumnList)

  If fOK Then
    ' Re-populate this table with saved data from temporary table.
    ' NB. We use the 'openResultSet' method to perform the 'SET IDENTITY_INSERT' operation
    ' now, instead of the 'execute' method which did not work under SQL Server 7.0.
    gADOCon.Execute _
        "SET IDENTITY_INSERT " & sPhysicalTableName & " ON" & vbNewLine & _
        "INSERT INTO " & sPhysicalTableName & " (" & Join(asColumnList, ",") & ")" & _
        " SELECT " & Join(asValueList, ",") & " FROM " & sTempName & vbNewLine & _
        "SET IDENTITY_INSERT " & sPhysicalTableName & " OFF", , adCmdText + adExecuteNoRecords
  End If

  ' Reseed the newly created table.
  SetNextIdentitySeed sPhysicalTableName, lngNextIdentitySeed

  ' JPD20030110 Fault 4162
  ' Recreate any custom triggers.
  sAreaInCode = "Recreating Custom Triggers"
  If fOK Then
    For iLoop = 1 To UBound(asTriggers)
      gADOCon.Execute asTriggers(iLoop), , adCmdText + adExecuteNoRecords
    Next iLoop
  End If

TidyUpAndExit:

  Set rsTriggerDefn = Nothing
  Set rsTriggers = Nothing

  ' Drop temporary tables.
  If Database.TableExists(sTempName) Then
    sSQL = "DROP TABLE " & sTempName
    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  End If
  TableSave = fOK
  Exit Function

ErrorTrap:
  strErrorMessage = "Error updating table '" & sTableName & "'"
  Select Case Err.Number
    Case -2147217900
      'This is a common error number so to make it idiosyncratic I've added sAreaInCode as a further check.
      If sAreaInCode = "Recreating Custom Triggers" Then
        strErrorMessage = sAreaInCode & " for '" & sTableName & "'" & vbNewLine & vbNewLine & _
            "You have changed a column or table name that is already referenced in a Custom Trigger." & vbNewLine & _
            "Change it back to the original value or amend the Custom Trigger that references it." & vbNewLine & vbNewLine & _
            "Error Message"
      End If
    Case -2147467259
      strErrorMessage = "Error updating table '" & sTableName & "'" & vbNewLine & vbNewLine & _
          "Unable to allocate required SQL Server resources. Please retry saving."
  End Select

  OutputError strErrorMessage
'  If blnReconnect Then
'    Reconnect sConnect
'  End If
  
  fOK = False
  Resume TidyUpAndExit
End Function

'Private Function Reconnect(sConnect As String)
'
'  'Try to re-establish database connection
'  On Local Error Resume Next
'
'  gADOCon.RollbackTrans
'  gADOCon.Close
'  Set gADOCon = Nothing
'
'  Set gADOCon = New ADODB.Connection
'  With gADOCon
'    .ConnectionString = sConnect
'    .Provider = "SQLOLEDB"
'    .CommandTimeout = 0
'    .ConnectionTimeout = 0
'    .CursorLocation = adUseServer
'    .Mode = adModeReadWrite
'    .Properties("Packet Size") = 32767
'    .Open
'  End With
'
'  gADOCon.BeginTrans
'
'End Function


Private Function CreateMaxIDStoredProcedure() As Boolean
  ' Create the Max ID stored procedure.
  ' JPD20030313 Fault 5159

  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim sSPCode As String
  
  Const sSPName = "dbo.spASRMaxID"
  
  fOK = True
  
  DropProcedure sSPName
  
  ' Create the stored procedure.
  sSPCode = "/* ------------------------------------------------------------------------------- */" & vbNewLine & _
    "/* HR Pro system stored procedure.                  */" & vbNewLine & _
    "/* Automatically generated by the System Manager.   */" & vbNewLine & _
    "/* ------------------------------------------------------------------------------- */" & vbNewLine & _
    "CREATE PROCEDURE " & sSPName & vbNewLine & _
    "(" & vbNewLine & _
    "    @piTableID integer,             /* Input variable to define the table ID. */" & vbNewLine & _
    "    @piMaxRecordID integer OUTPUT   /* Output variable to hold the max record ID. */" & vbNewLine & _
    ")" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine & _
    "    SET @piMaxRecordID = 0;" & vbNewLine & vbNewLine

  With recTabEdit
    .Index = "idxTableID"
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    Do While fOK And Not .EOF
      sSPCode = sSPCode & vbNewLine & _
        "    IF @piTableID = " & Trim(Str(recTabEdit!TableID)) & vbNewLine & _
        "    BEGIN" & vbNewLine & _
        "        SELECT @piMaxRecordID = MAX([id]) FROM [dbo].[tbuser_" & recTabEdit!TableName & "];" & vbNewLine & _
        "    END" & vbNewLine & vbNewLine

      .MoveNext
    Loop
  End With

  sSPCode = sSPCode & _
    "END"

  gADOCon.Execute sSPCode, , adCmdText + adExecuteNoRecords
  
TidyUpAndExit:
  CreateMaxIDStoredProcedure = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function GetColCreateString(ByVal psColumnName As String, ByVal plngDataType As Long, ByVal plngSize As Long _
      , ByVal piDecimals As Integer, ByVal bIsVarcharMax As Boolean) As String

  Select Case plngDataType
    Case dtVARCHAR
    If bIsVarcharMax Then
      GetColCreateString = "[" & Trim$(psColumnName) & "] [NVARCHAR] (MAX)"
    Else
      GetColCreateString = "[" & Trim$(psColumnName) & "] [VARCHAR] (" & plngSize & ")"
    End If

    Case dtLONGVARCHAR
      GetColCreateString = "[" & Trim$(psColumnName) & "] [VARCHAR] (14)"

    Case dtINTEGER
      GetColCreateString = "[" & Trim$(psColumnName) & "] [INT]"

    Case dtNUMERIC
      GetColCreateString = "[" & Trim$(psColumnName) & "] [NUMERIC] (" & plngSize & "," & piDecimals & ")"

    Case dtTIMESTAMP
      GetColCreateString = "[" & Trim$(psColumnName) & "] [DATETIME]"

    Case dtBIT
      GetColCreateString = "[" & Trim$(psColumnName) & "] [BIT]"

    Case dtVARBINARY, dtLONGVARBINARY
      GetColCreateString = "[" & Trim$(psColumnName) & "] [VARBINARY] (MAX)"

    Case dtGUID
      GetColCreateString = "[" & Trim$(psColumnName) & "] [UNIQUEIDENTIFIER]"

  End Select

End Function

Private Function CreateInsertStoredProcedure(plngTableID As Long, psTableName As String) As Boolean
  ' Create the insert stored procedure for the given table.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim sSPCode As String
  
  Const sSPName = "dbo.sp_ASRInsertNewRecord_"
  
  fOK = True
  
  DropProcedure sSPName & CStr(plngTableID)
  
  ' Create the stored procedure.
  sSPCode = "/* ------------------------------------------------------------------------------- */" & vbNewLine & _
    "/* HR Pro system stored procedure.                  */" & vbNewLine & _
    "/* Automatically generated by the System Manager.   */" & vbNewLine & _
    "/* ------------------------------------------------------------------------------- */" & vbNewLine & _
    "CREATE PROCEDURE " & sSPName & Trim$(Str$(plngTableID)) & vbNewLine & _
    "(" & vbNewLine & _
    "    @piNewRecordID integer OUTPUT,   /* Output variable to hold the new record ID. */" & vbNewLine & _
    "    @psInsertString nvarchar(MAX)    /* SQL Insert string to insert the new record. */" & vbNewLine & _
    ")" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine & _
    "    SET NOCOUNT ON;" & vbNewLine & vbNewLine & _
    "    /* Run the given SQL INSERT string. */" & vbNewLine & _
    "    EXECUTE (@psInsertString);" & vbNewLine & vbNewLine & _
    "    /* Get the ID of the inserted record." & vbNewLine & _
    "    NB. We do not use @@IDENTITY as the insertion that we have just performed may have triggered" & vbNewLine & _
    "    other insertions (eg. into the Audit Trail table. The @@IDENTITY variable would then be the last IDENTITY value" & vbNewLine & _
    "    entered in the Audit Trail table.*/" & vbNewLine & _
    "    SELECT @piNewRecordID = MAX(id) FROM " & psTableName & ";" & vbNewLine & vbNewLine & _
    "END"

  gADOCon.Execute sSPCode, , adCmdText + adExecuteNoRecords
  
TidyUpAndExit:
  CreateInsertStoredProcedure = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function CreateIntInsertStoredProcedure(plngTableID As Long, psTableName As String) As Boolean
  ' Create the insert stored procedure for the given table.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim sSPCode As New HRProSystemMgr.cStringBuilder
  
  Const sSPName = "dbo.spASRIntInsertNewRecord_"
  
  fOK = True
  
'  ' Drop the stored procedure if it already exists.
'  sSQL = "IF EXISTS" & _
'    " (SELECT Name" & _
'    "   FROM sysobjects" & _
'    "   WHERE id = object_id('" & sSPName & Trim$(Str$(plngTableID)) & "')" & _
'    "     AND sysstat & 0xf = 4)" & _
'    " DROP PROCEDURE " & sSPName & Trim$(Str$(plngTableID))
'  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  DropProcedure sSPName & CStr(plngTableID)
  
  ' Create the stored procedure.
  sSPCode.TheString = "/* ------------------------------------------------------------------------------- */" & vbNewLine & _
    "/* HR Pro Intranet Insert stored procedure.                 */" & vbNewLine & _
    "/* Automatically generated by the System Manager.   */" & vbNewLine & _
    "/* ------------------------------------------------------------------------------- */" & vbNewLine & _
    "CREATE PROCEDURE " & sSPName & Trim$(Str$(plngTableID)) & vbNewLine & _
    "(" & vbNewLine & _
    vbTab & "@piNewRecordID" & vbTab & "integer" & vbTab & "OUTPUT," & vbTab & "/* Output variable to hold the new record ID. */" & vbNewLine & _
    vbTab & "@psInsertDef" & vbTab & vbTab & "varchar(MAX)" & vbTab & vbTab & "/* SQL Insert string to insert the new record. */" & vbNewLine & _
    ")" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine
    
  sSPCode.Append _
    vbTab & "SET NOCOUNT ON;" & vbNewLine & vbNewLine & _
    vbTab & "DECLARE" & vbNewLine & _
    vbTab & vbTab & "@sTempString" & vbTab & vbTab & "varchar(MAX)," & vbNewLine & _
    vbTab & vbTab & "@sInsertString" & vbTab & vbTab & "nvarchar(MAX)," & vbNewLine & _
    vbTab & vbTab & "@iTemp" & vbTab & vbTab & "integer," & vbNewLine & _
    vbTab & vbTab & "@iCounter" & vbTab & vbTab & "integer," & vbNewLine & _
    vbTab & vbTab & "@iIndex1" & vbTab & vbTab & "integer," & vbNewLine & _
    vbTab & vbTab & "@iIndex2" & vbTab & vbTab & "integer," & vbNewLine & _
    vbTab & vbTab & "@iIndex3" & vbTab & vbTab & "integer," & vbNewLine & _
    vbTab & vbTab & "@sColumnID" & vbTab & vbTab & "varchar(255)," & vbNewLine & _
    vbTab & vbTab & "@sValue" & vbTab & vbTab & "varchar(MAX)," & vbNewLine
    
  sSPCode.Append _
    vbTab & vbTab & "@sColumnList" & vbTab & vbTab & "varchar(MAX)," & vbNewLine & _
    vbTab & vbTab & "@sValueList" & vbTab & vbTab & "varchar(MAX)," & vbNewLine & _
    vbTab & vbTab & "@iCopiedRecordID" & vbTab & "integer," & vbNewLine & _
    vbTab & vbTab & "@iDataType" & vbTab & vbTab & "integer," & vbNewLine & _
    vbTab & vbTab & "@sColumnName" & vbTab & vbTab & "varchar(255)," & vbNewLine & _
    vbTab & vbTab & "@sRealSource" & vbTab & vbTab & "sysname," & vbNewLine & _
    vbTab & vbTab & "@sMask" & vbTab & vbTab & "varchar(255)," & vbNewLine & _
    vbTab & vbTab & "@iOLEType" & vbTab & vbTab & "integer," & vbNewLine & _
    vbTab & vbTab & "@fCopyImageData" & vbTab & "bit;" & vbNewLine & vbNewLine & _
    vbTab & "SET @sColumnList = '';" & vbNewLine & _
    vbTab & "SET @sValueList = '';" & vbNewLine & vbNewLine

  sSPCode.Append _
    vbTab & "SET @iIndex1 = charindex(CHAR(9), @psInsertDef);" & vbNewLine & _
    vbTab & "SET @iIndex2 = charindex(CHAR(9), @psInsertDef, @iIndex1+1);" & vbNewLine & _
    vbTab & "SET @iIndex3 = charindex(CHAR(9), @psInsertDef, @iIndex2+1);" & vbNewLine & vbNewLine & _
    vbTab & "SET @sRealSource = replace(LEFT(@psInsertDef, @iIndex1-1), '''', '''''');" & vbNewLine & _
    vbTab & "SET @sValue = replace(SUBSTRING(@psInsertDef, @iIndex1+1, @iIndex2-@iIndex1-1), '''', '''''');" & vbNewLine & _
    vbTab & "SET @fCopyImageData = convert(bit, @sValue);" & vbNewLine & _
    vbTab & "SET @sValue = replace(SUBSTRING(@psInsertDef, @iIndex2+1, @iIndex3-@iIndex2-1), '''', '''''');" & vbNewLine & _
    vbTab & "SET @iCopiedRecordID = convert(integer, @sValue);" & vbNewLine & vbNewLine & _
    vbTab & "SET @psInsertDef = SUBSTRING(@psInsertDef, @iIndex3+1, LEN(@psInsertDef) - @iIndex3);" & vbNewLine & vbNewLine

  sSPCode.Append _
    vbTab & "SET @sColumnList = 'INSERT INTO ' + convert(varchar(255), @sRealSource) + ' (';" & vbNewLine & _
    vbTab & "SET @sValueList = '';" & vbNewLine & _
    vbTab & "SET @iCounter = 0;" & vbNewLine & vbNewLine

  sSPCode.Append _
    vbTab & "WHILE charindex(CHAR(9), @psInsertDef) > 0" & vbNewLine & _
    vbTab & "BEGIN" & vbNewLine & _
    vbTab & vbTab & "SET @iIndex1 = charindex(CHAR(9), @psInsertDef);" & vbNewLine & _
    vbTab & vbTab & "SET @iIndex2 = charindex(CHAR(9), @psInsertDef, @iIndex1+1);" & vbNewLine & vbNewLine & _
    vbTab & vbTab & "SET @sColumnID = replace(LEFT(@psInsertDef, @iIndex1-1), '''', '''''');" & vbNewLine & _
    vbTab & vbTab & "SET @sValue = replace(SUBSTRING(@psInsertDef, @iIndex1+1, @iIndex2-@iIndex1-1), '''', '''''');" & vbNewLine & vbNewLine & _
    vbTab & vbTab & "IF LEFT(@sColumnID, 3) = 'ID_'" & vbNewLine & _
    vbTab & vbTab & "BEGIN" & vbNewLine & _
    vbTab & vbTab & vbTab & "SET @sColumnName = @sColumnID;" & vbNewLine & _
    vbTab & vbTab & "END" & vbNewLine & _
    vbTab & vbTab & "ELSE" & vbNewLine & _
    vbTab & vbTab & "BEGIN" & vbNewLine & _
    vbTab & vbTab & vbTab & "SELECT @sColumnName = ASRSysColumns.columnName," & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & "@iDataType = ASRSysColumns.dataType," & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & "@sMask = ASRSysColumns.mask" & vbNewLine & _
    vbTab & vbTab & vbTab & "FROM ASRSysColumns" & vbNewLine & _
    vbTab & vbTab & vbTab & "WHERE ASRSysColumns.columnID = convert(integer, @sColumnID);" & vbNewLine & vbNewLine

  sSPCode.Append _
    vbTab & vbTab & vbTab & "-- Date" & vbNewLine & _
    vbTab & vbTab & vbTab & "IF (@iDataType = 11 AND @sValue <> 'null') SET @sValue = '''' + @sValue + '''';" & vbNewLine & vbNewLine & _
    vbTab & vbTab & vbTab & "-- Character" & vbNewLine & _
    vbTab & vbTab & vbTab & "IF (@iDataType = 12 AND (LEN(@sMask) = 0 OR @sValue <> 'null')) SET @sValue = '''' + @sValue + '''';" & vbNewLine & vbNewLine & _
    vbTab & vbTab & vbTab & "-- WorkingPattern" & vbNewLine & _
    vbTab & vbTab & vbTab & "IF (@iDataType = -1) SET @sValue = '''' + @sValue + '''';" & vbNewLine & vbNewLine & _
    vbTab & vbTab & vbTab & "-- Photo / OLE" & vbNewLine & _
    vbTab & vbTab & vbTab & "IF (@iDataType = -3 OR @iDataType = -4)" & vbNewLine & _
    vbTab & vbTab & vbTab & "BEGIN" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & "SET @iOLEType = convert(integer, LEFT(@sValue, 1));" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & "SET @sValue = SUBSTRING(@sValue, 2, LEN(@sValue) - 1);" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & "IF (@iOLEType < 2) SET @sValue = '''' + @sValue + '''';" & vbNewLine & _
    vbTab & vbTab & vbTab & "END" & vbNewLine & _
    vbTab & vbTab & "END" & vbNewLine & vbNewLine
    
  sSPCode.Append _
    vbTab & vbTab & "SET @sTempString =" & vbNewLine & _
    vbTab & vbTab & vbTab & "CASE" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & "WHEN @iCounter > 0 THEN ','" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & "ELSE ''" & vbNewLine & _
    vbTab & vbTab & vbTab & "END" & vbNewLine & _
    vbTab & vbTab & vbTab & "+ convert(varchar(255), @sColumnName);" & vbNewLine & vbNewLine & _
    vbTab & vbTab & "SET @sColumnList = @sColumnList + @sTempString;" & vbNewLine

  sSPCode.Append _
    vbTab & vbTab & "SET @sTempString =" & vbNewLine & _
    vbTab & vbTab & vbTab & "CASE" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & "WHEN @iCounter > 0 THEN ','" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & "ELSE ''" & vbNewLine & _
    vbTab & vbTab & vbTab & "END" & vbNewLine & _
    vbTab & vbTab & vbTab & "+ CASE" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & "WHEN @fCopyImageData = 1 THEN REPLACE(convert(varchar(MAX), @sValue), '''', '''''')" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & "ELSE convert(varchar(MAX), @sValue)" & vbNewLine & _
    vbTab & vbTab & vbTab & "END;" & vbNewLine & vbNewLine & _
    vbTab & vbTab & "SET @sValueList = @sValueList + @sTempString;" & vbNewLine & _
    vbTab & vbTab & "SET @iCounter = @iCounter + 1;" & vbNewLine & _
    vbTab & vbTab & "SET @psInsertDef = SUBSTRING(@psInsertDef, @iIndex2+1, LEN(@psInsertDef) - @iIndex2);" & vbNewLine & _
    vbTab & "END" & vbNewLine & vbNewLine
    
  'TM20070418 - Fault 12138.
  sSPCode.Append _
    vbTab & "IF @fCopyImageData = 1" & vbNewLine & _
    vbTab & "BEGIN" & vbNewLine & _
    vbTab & vbTab & "SET @sInsertString = @sColumnList + ')'" & vbNewLine & _
    vbTab & vbTab & vbTab & "+ ' EXECUTE(''SELECT ' + @sValueList" & vbNewLine & _
    vbTab & vbTab & vbTab & "+ ' FROM ' + convert(varchar(255), @sRealSource)" & vbNewLine & _
    vbTab & vbTab & vbTab & "+ ' WHERE id = ' + convert(varchar(255), @iCopiedRecordID) + ''')';" & vbNewLine & _
    vbTab & "END" & vbNewLine & _
    vbTab & "ELSE" & vbNewLine & _
    vbTab & "BEGIN" & vbNewLine & _
    vbTab & vbTab & "SET @sInsertString = @sColumnList + ')' + ' VALUES(' + @sValueList + ')';" & vbNewLine & _
    vbTab & "END" & vbNewLine & vbNewLine & _
    vbTab & "-- Run the constructed SQL INSERT string." & vbNewLine & _
    vbTab & "EXECUTE sp_executeSQL @sInsertString;" & vbNewLine & vbNewLine
    
  sSPCode.Append _
    vbTab & "/* Get the ID of the inserted record." & vbNewLine & _
    vbTab & "NB. We do not use @@IDENTITY as the insertion that we have just performed may have triggered" & vbNewLine & _
    vbTab & "other insertions (eg. into the Audit Trail table). The @@IDENTITY variable would then be the last IDENTITY value" & vbNewLine & _
    vbTab & "entered in the Audit Trail table.*/" & vbNewLine & _
    vbTab & "SELECT @piNewRecordID = MAX(id) FROM " & psTableName & ";" & vbNewLine & _
    "END"

  gADOCon.Execute sSPCode.ToString, , adCmdText + adExecuteNoRecords
  
TidyUpAndExit:
  CreateIntInsertStoredProcedure = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function CreateUpdateStoredProcedure(plngTableID As Long, psTableName As String) As Boolean
  ' Create the update stored procedure for the given table.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim sSPCode As New HRProSystemMgr.cStringBuilder
  
  Const sSPName = "dbo.sp_ASRUpdateRecord_"
  
  fOK = True
  
'  ' Drop the stored procedure if it already exists.
'  sSQL = "IF EXISTS" & _
'    " (SELECT Name" & _
'    "   FROM sysobjects" & _
'    "   WHERE id = object_id('" & sSPName & Trim$(Str$(plngTableID)) & "')" & _
'    "     AND sysstat & 0xf = 4)" & _
'    " DROP PROCEDURE " & sSPName & Trim$(Str$(plngTableID))
'  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  DropProcedure sSPName & CStr(plngTableID)
  
  ' Create the stored procedure.
  sSPCode.TheString = "CREATE PROCEDURE " & sSPName & Trim$(Str$(plngTableID)) & vbNewLine & _
    "(" & vbNewLine & _
    "    @piResult integer OUTPUT,   /* Output variable to hold the result. */" & vbNewLine & _
    "    @psUpdateString nvarchar(MAX),    /* SQL Update string to update the record. */" & vbNewLine & _
    "    @psRealSource sysname,    /* RealSource  being updated. */" & vbNewLine & _
    "    @piID integer,    /* ID the record being updated. */" & vbNewLine & _
    "    @piTimestamp integer    /* Original timestamp of the record being updated. */" & vbNewLine & _
    ")" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine & _
    "    SET NOCOUNT ON;" & vbNewLine & vbNewLine & _
    "    /* Run the given SQL UPDATE string. */" & vbNewLine & _
    "    /* Check if the given record has been deleted or changed first. */" & vbNewLine & _
    "    /* Return 0 if the record was OK to update. */" & vbNewLine & _
    "    /* Return 1 if the record has been amended AND is still in the given table/view. */" & vbNewLine & _
    "    /* Return 2 if the record has been amended AND is no longer in the given table/view. */" & vbNewLine & _
    "    /* Return 3 if the record has been deleted from the table. */" & vbNewLine

  sSPCode.Append _
    "    DECLARE @iCurrentTimestamp integer," & vbNewLine & _
    "        @sSQL nvarchar(MAX)," & vbNewLine & _
    "        @iCount integer;" & vbNewLine & _
    "    SET @piResult = 0;" & vbNewLine & vbNewLine & _
    "    /* Check that the record has not been updated by another user since it was last checked. */" & vbNewLine & _
    "    SELECT @iCurrentTimestamp = convert(integer, timestamp)" & vbNewLine & _
    "    FROM " & psTableName & vbNewLine & _
    "    WHERE id = @piID;" & vbNewLine & vbNewLine

  sSPCode.Append _
    "    IF @iCurrentTimestamp IS null" & vbNewLine & _
    "    BEGIN" & vbNewLine & _
    "        /* Record deleted. */" & vbNewLine & _
    "        SET @piResult = 3;" & vbNewLine & _
    "    END" & vbNewLine & _
    "    ELSE" & vbNewLine & _
    "    BEGIN" & vbNewLine & _
    "        IF (@iCurrentTimestamp <> @piTimestamp) AND (NOT @piTimestamp IS null)" & vbNewLine & _
    "        BEGIN" & vbNewLine & _
    "            /* Record changed. Check if it is in the given realsource. */" & vbNewLine & _
    "           SET @sSQL = 'SELECT @piResult = COUNT(id)' +" & vbNewLine & _
    "             ' FROM ' + @psRealSource +" & vbNewLine & _
    "             ' WHERE id = ' + convert(varchar(255), @piID);" & vbNewLine & _
    "           EXECUTE sp_executesql @sSQL, N'@piResult int OUTPUT', @iCount OUTPUT;" & vbNewLine & _
    "           IF @iCount > 0" & vbNewLine & _
    "           BEGIN" & vbNewLine & _
    "               SET @piResult = 1;" & vbNewLine & _
    "           END" & vbNewLine & _
    "           ELSE" & vbNewLine & _
    "           BEGIN" & vbNewLine & _
    "               SET @piResult = 2;" & vbNewLine & _
    "           END" & vbNewLine

  sSPCode.Append _
    "        END" & vbNewLine & _
    "        ELSE" & vbNewLine & _
    "        BEGIN" & vbNewLine & _
    "            -- Run the given SQL UPDATE string." & vbNewLine & _
    "            EXECUTE sp_executeSQL @psUpdateString;" & vbNewLine & _
    "        END" & vbNewLine & _
    "    END" & vbNewLine & vbNewLine & _
    "END"

  gADOCon.Execute sSPCode.ToString, , adCmdText + adExecuteNoRecords
  
TidyUpAndExit:
  CreateUpdateStoredProcedure = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function CreateIntUpdateStoredProcedure(plngTableID As Long, psTableName As String) As Boolean
  ' Create the INTRANET update stored procedure for the given table.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim sSPCode As New HRProSystemMgr.cStringBuilder
  
  Const sSPName = "dbo.spASRIntUpdateRecord_"
  
  fOK = True
  
'  ' Drop the stored procedure if it already exists.
'  sSQL = "IF EXISTS" & _
'    " (SELECT Name" & _
'    "   FROM sysobjects" & _
'    "   WHERE id = object_id('" & sSPName & Trim$(Str$(plngTableID)) & "')" & _
'    "     AND sysstat & 0xf = 4)" & _
'    " DROP PROCEDURE " & sSPName & Trim$(Str$(plngTableID))
'  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  DropProcedure sSPName & CStr(plngTableID)
  
  ' Create the stored procedure.
  sSPCode.TheString = "/* ------------------------------------------------------------------------------- */" & vbNewLine & _
    "/* HR Pro Intranet Update stored procedure.              */" & vbNewLine & _
    "/* Automatically generated by the System Manager.   */" & vbNewLine & _
    "/* ------------------------------------------------------------------------------- */" & vbNewLine & _
    "CREATE PROCEDURE " & sSPName & Trim$(Str$(plngTableID)) & vbNewLine & _
    "(" & vbNewLine & _
    vbTab & "@piResult" & vbTab & vbTab & "integer" & vbTab & "OUTPUT," & vbTab & "/* Output variable to hold the result. */" & vbNewLine & _
    vbTab & "@psUpdateDef varchar(MAX)," & vbTab & vbTab & "/* Update definition to update the record. */" & vbNewLine & _
    vbTab & "@psRealSource sysname," & vbTab & vbTab & "/* RealSource  being updated. */" & vbNewLine & _
    vbTab & "@piID" & vbTab & vbTab & "integer," & vbTab & vbTab & vbTab & "/* ID the record being updated. */" & vbNewLine & _
    vbTab & "@piTimestamp integer" & vbTab & vbTab & vbTab & "/* Original timestamp of the record being updated. */" & vbNewLine & _
    ")" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine & _
    vbTab & "/* Return 0 if the record was OK to update. */" & vbNewLine & _
    vbTab & "/* Return 1 if the record has been amended AND is still in the given table/view. */" & vbNewLine & _
    vbTab & "/* Return 2 if the record has been amended AND is no longer in the given table/view. */" & vbNewLine & _
    vbTab & "/* Return 3 if the record has been deleted from the table. */" & vbNewLine
    
  sSPCode.Append _
    vbTab & "SET NOCOUNT ON;" & vbNewLine & vbNewLine & _
    vbTab & "DECLARE" & vbNewLine & _
    vbTab & vbTab & "@iCurrentTimestamp" & vbTab & "integer," & vbNewLine & _
    vbTab & vbTab & "@sSQL" & vbTab & vbTab & vbTab & "nvarchar(MAX)," & vbNewLine & _
    vbTab & vbTab & "@iCount" & vbTab & vbTab & "integer," & vbNewLine & _
    vbTab & vbTab & "@sUpdateString" & vbTab & vbTab & "nvarchar(MAX)," & vbNewLine & _
    vbTab & vbTab & "@sTempString" & vbTab & vbTab & "varchar(MAX)," & vbNewLine & _
    vbTab & vbTab & "@iCounter" & vbTab & vbTab & vbTab & "integer," & vbNewLine & _
    vbTab & vbTab & "@iIndex1" & vbTab & vbTab & vbTab & "integer," & vbNewLine & _
    vbTab & vbTab & "@iIndex2" & vbTab & vbTab & vbTab & "integer," & vbNewLine & _
    vbTab & vbTab & "@sColumnID" & vbTab & vbTab & vbTab & "varchar(255)," & vbNewLine & _
    vbTab & vbTab & "@sValue" & vbTab & vbTab & "varchar(MAX)," & vbNewLine & _
    vbTab & vbTab & "@iDataType" & vbTab & vbTab & "integer," & vbNewLine & _
    vbTab & vbTab & "@sColumnName" & vbTab & vbTab & "varchar(255)," & vbNewLine & _
    vbTab & vbTab & "@sMask" & vbTab & vbTab & vbTab & "varchar(MAX)," & vbNewLine & _
    vbTab & vbTab & "@iOLEType" & vbTab & vbTab & "integer;" & vbNewLine & vbNewLine
        
  sSPCode.Append _
    vbTab & "-- Clean the input string parameters." & vbNewLine & _
    vbTab & "IF len(@psRealsource) > 0 SET @psRealsource = replace(@psRealsource, '''', '''''');" & vbNewLine & vbNewLine & _
    vbTab & "SET @piResult = 0;" & vbNewLine & _
    vbTab & "SET @sUpdateString = 'UPDATE ' + convert(varchar(255), @psRealSource) + ' SET ';" & vbNewLine & _
    vbTab & "SET @iCounter = 0;" & vbNewLine & vbNewLine & _
    vbTab & "-- Check that the record has not been updated by another user since it was last checked." & vbNewLine & _
    vbTab & "SELECT @iCurrentTimestamp = convert(integer, timestamp)" & vbNewLine & _
    vbTab & "FROM " & psTableName & vbNewLine & _
    vbTab & "WHERE id = @piID;" & vbNewLine & vbNewLine

  sSPCode.Append _
    vbTab & "IF @iCurrentTimestamp IS null" & vbNewLine & _
    vbTab & "BEGIN" & vbNewLine & _
    vbTab & vbTab & "-- Record deleted." & vbNewLine & _
    vbTab & vbTab & "SET @piResult = 3;" & vbNewLine & _
    vbTab & "END" & vbNewLine & _
    vbTab & "ELSE" & vbNewLine & _
    vbTab & "BEGIN" & vbNewLine & _
    vbTab & vbTab & "IF (@iCurrentTimestamp <> @piTimestamp) AND (NOT @piTimestamp IS null)" & vbNewLine & _
    vbTab & vbTab & "BEGIN" & vbNewLine & _
    vbTab & vbTab & vbTab & "-- Record changed. Check if it is in the given realsource." & vbNewLine & _
    vbTab & vbTab & vbTab & "SET @sSQL = 'SELECT @piResult = COUNT(id)' +" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & "' FROM ' + @psRealSource +" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & "' WHERE id = ' + convert(varchar(255), @piID)" & vbNewLine & _
    vbTab & vbTab & vbTab & "EXECUTE sp_executesql @sSQL, N'@piResult int OUTPUT', @iCount OUTPUT;" & vbNewLine & vbNewLine & _
    vbTab & vbTab & vbTab & "IF @iCount > 0" & vbNewLine & _
    vbTab & vbTab & vbTab & "BEGIN" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & "SET @piResult = 1;" & vbNewLine & _
    vbTab & vbTab & vbTab & "END" & vbNewLine & _
    vbTab & vbTab & vbTab & "ELSE" & vbNewLine & _
    vbTab & vbTab & vbTab & "BEGIN" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & "SET @piResult = 2;" & vbNewLine & _
    vbTab & vbTab & vbTab & "END" & vbNewLine

  sSPCode.Append _
    vbTab & vbTab & "END" & vbNewLine & _
    vbTab & vbTab & "ELSE" & vbNewLine & _
    vbTab & vbTab & "BEGIN" & vbNewLine & _
    vbTab & vbTab & vbTab & "WHILE charindex(CHAR(9), @psUpdateDef) > 0" & vbNewLine & _
    vbTab & vbTab & vbTab & "BEGIN" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & "SET @iIndex1 = charindex(CHAR(9), @psUpdateDef);" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & "SET @iIndex2 = charindex(CHAR(9), @psUpdateDef, @iIndex1+1);" & vbNewLine & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & "SET @sColumnID = replace(LEFT(@psUpdateDef, @iIndex1-1), '''', '''''');" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & "SET @sValue = replace(SUBSTRING(@psUpdateDef, @iIndex1+1, @iIndex2-@iIndex1-1), '''', '''''');" & vbNewLine & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & "IF LEFT(@sColumnID, 3) = 'ID_'" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & "BEGIN" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & vbTab & "SET @sColumnName = @sColumnID;" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & "END" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & "ELSE" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & "BEGIN" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & vbTab & "SELECT @sColumnName = ASRSysColumns.columnName," & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "@iDataType = ASRSysColumns.dataType," & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "@sMask = ASRSysColumns.mask" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & vbTab & "FROM ASRSysColumns" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & vbTab & "WHERE ASRSysColumns.columnID = convert(integer, @sColumnID);" & vbNewLine & vbNewLine

  sSPCode.Append _
    vbTab & vbTab & vbTab & vbTab & vbTab & "-- Date" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & vbTab & "IF (@iDataType = 11 AND @sValue <> 'null') SET @sValue = '''' + @sValue + '''';" & vbNewLine & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & vbTab & "-- Character" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & vbTab & "IF (@iDataType = 12 AND (LEN(@sMask) = 0 OR @sValue <> 'null')) SET @sValue = '''' + @sValue + '''';" & vbNewLine & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & vbTab & "-- WorkingPattern" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & vbTab & "IF (@iDataType = -1) SET @sValue = '''' + @sValue + '''';" & vbNewLine & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & vbTab & "-- Photo / OLE" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & vbTab & "IF (@iDataType = -3 OR @iDataType = -4)" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & vbTab & "BEGIN" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "SET @iOLEType = convert(integer, LEFT(@sValue, 1));" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "SET @sValue = SUBSTRING(@sValue, 2, LEN(@sValue) - 1);" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "IF (@iOLEType < 2) SET @sValue = '''' + @sValue + '''';" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & vbTab & "END" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & "END" & vbNewLine & vbNewLine
          
  sSPCode.Append _
    vbTab & vbTab & vbTab & vbTab & "SET @sTempString =" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & vbTab & "CASE" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "WHEN @iCounter > 0 THEN ','" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "ELSE ''" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & vbTab & "END" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & vbTab & "+ convert(varchar(255), @sColumnName) + ' = ' + convert(varchar(MAX), @sValue);" & vbNewLine & vbNewLine

  sSPCode.Append _
    vbTab & vbTab & vbTab & vbTab & "SET @sUpdateString = @sUpdateString + @sTempString;" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & "SET @iCounter = @iCounter + 1;" & vbNewLine & _
    vbTab & vbTab & vbTab & vbTab & "SET @psUpdateDef = SUBSTRING(@psUpdateDef, @iIndex2+1, LEN(@psUpdateDef) - @iIndex2);" & vbNewLine & _
    vbTab & vbTab & vbTab & "END" & vbNewLine & vbNewLine & _
    vbTab & vbTab & vbTab & "SET @sUpdateString = @sUpdateString + ' WHERE id = ' + convert(varchar(255), @piID);" & vbNewLine & vbNewLine
   
  sSPCode.Append _
    vbTab & vbTab & vbTab & "-- Run the constructed SQL UPDATE string." & vbNewLine & _
    vbTab & vbTab & vbTab & "EXEC sp_executeSQL @sUpdateString;" & vbNewLine & _
    vbTab & vbTab & "END" & vbNewLine & _
    vbTab & "END" & vbNewLine & _
    "END"

  gADOCon.Execute sSPCode.ToString, , adCmdText + adExecuteNoRecords
  
TidyUpAndExit:
  CreateIntUpdateStoredProcedure = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function CreateDeleteStoredProcedure(plngTableID As Long, psTableName As String) As Boolean
  ' Create the delete stored procedure for the given table.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim sSPCode As String
  
  Const sSPName = "dbo.sp_ASRDeleteRecord_"
  
  fOK = True
  
  DropProcedure sSPName & CStr(plngTableID)
  
  ' Create the stored procedure.
  sSPCode = "/* ------------------------------------------------------------------------------- */" & vbNewLine & _
    "/* HR Pro system stored procedure.                  */" & vbNewLine & _
    "/* Automatically generated by the System Manager.   */" & vbNewLine & _
    "/* ------------------------------------------------------------------------------- */" & vbNewLine & _
    "CREATE PROCEDURE " & sSPName & Trim$(Str$(plngTableID)) & vbNewLine & _
    "(" & vbNewLine & _
    "    @piResult integer OUTPUT,   /* Output variable to hold the result. */" & vbNewLine & _
    "    @psRealSource sysname,    /* RealSource being deleted from. */" & vbNewLine & _
    "    @piID integer    /* ID the record being deleted. */" & vbNewLine & _
    ")" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine & _
    "    SET NOCOUNT ON;" & vbNewLine & vbNewLine & _
    "    /*  Delete the given record. */" & vbNewLine & _
    "    /* Check if the given record has been deleted or changed first. */" & vbNewLine & _
    "    /* Return 0 if the record was OK to delete. */" & vbNewLine & _
    "    /* NOT USED HERE - Return 1 if the record has been amended AND is still in the given table/view. */" & vbNewLine & _
    "    /* Return 2 if the record has been amended AND is no longer in the given table/view. */" & vbNewLine & _
    "    /* Return 3 if the record has been deleted from the table. */" & vbNewLine
    
  sSPCode = sSPCode & _
    "    DECLARE @sSQL nvarchar(MAX)," & vbNewLine & _
    "        @iCount integer;" & vbNewLine & _
    "    SET @piResult = 0;" & vbNewLine & vbNewLine & _
    "    /* Check that the record has not been updated by another user since it was last checked. */" & vbNewLine & _
    "    SELECT @iCount = COUNT(id)" & vbNewLine & _
    "    FROM " & psTableName & vbNewLine & _
    "    WHERE id = @piID;" & vbNewLine & vbNewLine

  sSPCode = sSPCode & _
    "    IF @iCount = 0" & vbNewLine & _
    "    BEGIN" & vbNewLine & _
    "        /* Record deleted. */" & vbNewLine & _
    "        SET @piResult = 3;" & vbNewLine & _
    "    END" & vbNewLine & _
    "    ELSE" & vbNewLine & _
    "    BEGIN" & vbNewLine & _
    "        /* Check if the record is still in the given realsource. */" & vbNewLine & _
    "        SET @sSQL = 'SELECT @piResult = COUNT(id)' +" & vbNewLine & _
    "            ' FROM ' + @psRealSource +" & vbNewLine & _
    "            ' WHERE id = ' + convert(varchar(MAX), @piID);" & vbNewLine & _
    "        EXECUTE sp_executesql @sSQL, N'@piResult int OUTPUT', @iCount OUTPUT;" & vbNewLine

  sSPCode = sSPCode & _
    "        IF @iCount > 0" & vbNewLine & _
    "        BEGIN" & vbNewLine & _
    "            SET @sSQL = 'DELETE ' +" & vbNewLine & _
    "                ' FROM ' + @psRealSource +" & vbNewLine & _
    "                ' WHERE id = ' + convert(varchar(MAX), @piID);" & vbNewLine & _
    "            EXECUTE sp_executesql @sSQL;" & vbNewLine & _
    "        END" & vbNewLine & _
    "        ELSE" & vbNewLine & _
    "        BEGIN" & vbNewLine & _
    "            SET @piResult = 2;" & vbNewLine & _
    "        END" & vbNewLine & _
    "    END" & vbNewLine & vbNewLine & _
    "END"

  gADOCon.Execute sSPCode, , adCmdText + adExecuteNoRecords
  
TidyUpAndExit:
  CreateDeleteStoredProcedure = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function



Private Function CreateRecordAmendedStoredProcedure(plngTableID As Long, psTableName As String) As Boolean
  ' Create the update stored procedure for the given table.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim sSPCode As String
  
  Const sSPName = "dbo.sp_ASRRecordAmended_"
  
  fOK = True

  DropProcedure sSPName & CStr(plngTableID)
  
  ' Create the stored procedure.
  sSPCode = "/* ------------------------------------------------------------------------------- */" & vbNewLine & _
    "/* HR Pro system stored procedure.                  */" & vbNewLine & _
    "/* Automatically generated by the System Manager.   */" & vbNewLine & _
    "/* ------------------------------------------------------------------------------- */" & vbNewLine & _
    "CREATE PROCEDURE " & sSPName & Trim$(Str$(plngTableID)) & vbNewLine & _
    "(" & vbNewLine & _
    "    @piResult integer OUTPUT,   /* Output variable to hold the result. */" & vbNewLine & _
    "    @psRealSource sysname,    /* RealSource  being updated. */" & vbNewLine & _
    "    @piID integer,    /* ID the record being updated. */" & vbNewLine & _
    "    @piTimestamp integer    /* Original timestamp of the record being updated. */" & vbNewLine & _
    ")" & vbNewLine & _
    "AS" & vbNewLine & _
    "BEGIN" & vbNewLine & _
    "    /* Check if the given record has been deleted or changed by another user. */" & vbNewLine & _
    "    /* Return 0 if the record has NOT been amended. */" & vbNewLine & _
    "    /* Return 1 if the record has been amended AND is still in the given table/view. */" & vbNewLine & _
    "    /* Return 2 if the record has been amended AND is no longer in the given table/view. */" & vbNewLine & _
    "    /* Return 3 if the record has been deleted from the table. */" & vbNewLine & _
    "    SET NOCOUNT ON;" & vbNewLine

  sSPCode = sSPCode & _
    "    DECLARE @iCurrentTimestamp integer," & vbNewLine & _
    "        @sSQL nvarchar(MAX)," & vbNewLine & _
    "        @iCount integer;" & vbNewLine & _
    "    SET @piResult = 0;" & vbNewLine & vbNewLine & _
    "    /* Check that the record has not been updated by another user since it was last checked. */" & vbNewLine & _
    "    SELECT @iCurrentTimestamp = convert(integer, timestamp)" & vbNewLine & _
    "    FROM " & psTableName & vbNewLine & _
    "    WHERE id = @piID;" & vbNewLine & vbNewLine

  sSPCode = sSPCode & _
    "    IF @iCurrentTimestamp IS null" & vbNewLine & _
    "    BEGIN" & vbNewLine & _
    "        /* Record deleted. */" & vbNewLine & _
    "        SET @piResult = 3;" & vbNewLine & _
    "    END" & vbNewLine & _
    "    ELSE" & vbNewLine & _
    "    BEGIN" & vbNewLine & _
    "        IF @iCurrentTimestamp <> @piTimestamp" & vbNewLine & _
    "        BEGIN" & vbNewLine & _
    "            /* Record changed. Check if it is in the given realsource. */" & vbNewLine & _
    "           SET @sSQL = 'SELECT @piResult = COUNT(id)' +" & vbNewLine & _
    "             ' FROM ' + @psRealSource +" & vbNewLine & _
    "             ' WHERE id = ' + convert(varchar(255), @piID);" & vbNewLine & _
    "           EXECUTE sp_executesql @sSQL, N'@piResult int OUTPUT', @iCount OUTPUT;" & vbNewLine & _
    "           IF @iCount > 0" & vbNewLine & _
    "           BEGIN" & vbNewLine & _
    "               SET @piResult = 1;" & vbNewLine & _
    "           END" & vbNewLine & _
    "           ELSE" & vbNewLine & _
    "           BEGIN" & vbNewLine & _
    "               SET @piResult = 2;" & vbNewLine & _
    "           END" & vbNewLine

  sSPCode = sSPCode & _
    "        END" & vbNewLine & _
    "    END" & vbNewLine & _
    "END"

  gADOCon.Execute sSPCode, , adCmdText + adExecuteNoRecords
  
TidyUpAndExit:
  CreateRecordAmendedStoredProcedure = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function SetNextIdentitySeed(ByVal psTableName, ByVal mlngNewSeedValue As Long) As Boolean

  On Error GoTo ErrorTrap

  Dim sSQL As String

  If GetNextIdentitySeed(psTableName) <> mlngNewSeedValue Then
  sSQL = "DBCC CHECKIDENT ([" & psTableName & " ], RESEED, " & mlngNewSeedValue & ")"
  gADOCon.Execute sSQL, -1, adExecuteNoRecords
  End If
  
  SetNextIdentitySeed = True
  Exit Function
  
ErrorTrap:
  SetNextIdentitySeed = False

End Function

