Attribute VB_Name = "modSave_Accord"
Option Explicit

Public Function CreateAccordExpressionSPs(pfRefreshDatabase As Boolean) As Boolean

  On Error GoTo ErrorTrap

  Dim rsProcedures As dao.Recordset
  Dim rsMapData As dao.Recordset
  'Dim rsExistingProcedures As ADODB.Recordset

  Dim sSQL As String
  Dim strSPCode As String
  Dim bOK As Boolean
  Dim sProcedureName As String

  bOK = True
  'Set rsExistingProcedures = New ADODB.Recordset

  sSQL = "SELECT TransferFieldID, TransferTypeID FROM tmpAccordTransferFieldDefinitions WHERE ConvertData = true"
  Set rsProcedures = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

  Do While Not (rsProcedures.EOF)

    sProcedureName = "spASRAccordExpr_" & rsProcedures!TransferTypeID & "_" & rsProcedures!TransferFieldID
'    ' Drop any existing stored procedure with this name.
'    sSQL = "SELECT Name FROM sysObjects WHERE Type='P' AND Name='" & sProcedureName & "'"
'    rsExistingProcedures.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
'
'    If Not (rsExistingProcedures.BOF And rsExistingProcedures.EOF) Then
'      sSQL = "DROP PROCEDURE " & sProcedureName
'      gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'    End If
'    rsExistingProcedures.Close
    DropProcedure sProcedureName


    strSPCode = "/* ---------------------------------------------------------------- */" & vbNewLine _
            & "/* HR Pro Payroll module stored procedure.          */" & vbNewLine _
            & "/* Automatically generated by the System manager.   */" & vbNewLine _
            & "/* ---------------------------------------------------------------- */" & vbNewLine _
            & "CREATE PROCEDURE dbo." & sProcedureName & vbNewLine _
            & "(    @inputValue varchar(255)," & vbNewLine _
            & "     @result varchar(255) OUTPUT)" & vbNewLine _
            & "AS" & vbNewLine _
            & "BEGIN" & vbNewLine & vbNewLine _

    sSQL = "SELECT * FROM tmpAccordTransferFieldMappings WHERE FieldID = " & rsProcedures!TransferFieldID _
            & " AND TransferID = " & rsProcedures!TransferTypeID
    Set rsMapData = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    strSPCode = strSPCode & Space$(4) & "SET @result = @inputValue" & vbNewLine

    Do While Not rsMapData.EOF
      strSPCode = strSPCode & Space$(4) & "IF @inputValue = '" & rsMapData!HRProValue & "' SET @result = '" & rsMapData!AccordValue & "'" & vbNewLine
      rsMapData.MoveNext
    Loop

    strSPCode = strSPCode & "END"
  
    gADOCon.Execute strSPCode, , adCmdText + adExecuteNoRecords
    
    rsProcedures.MoveNext
  Loop

  rsProcedures.Close

TidyUpAndExit:
  Set rsProcedures = Nothing
  Set rsMapData = Nothing
  CreateAccordExpressionSPs = bOK
  Exit Function

ErrorTrap:
  OutputError "Error creating Payroll expression stored procedures"
  bOK = False
  Resume TidyUpAndExit

End Function

' Creates the triggers on the transfer table
Public Function CreateAccordTransferTriggers(pfRefreshDatabase As Boolean) As Boolean

  On Error GoTo ErrorTrap

  Dim rsAccordDetails As dao.Recordset
  Dim strUpdateTrigger As String
  Dim strDeleteTrigger As String
  Dim strDropTrigger As String
  Dim bOK As Boolean
  Dim sSQL As String
  
  bOK = True
   
  ' Drop existing trigger (Insert)
  strDropTrigger = "IF EXISTS" & _
    " (SELECT Name" & _
    "   FROM sysobjects" & _
    "   WHERE id = object_id('[INS_ASRSysAccordTransactions]')" & _
    "     AND objectproperty(id, N'IsTrigger') = 1)" & _
    " DROP TRIGGER [INS_ASRSysAccordTransactions]"
  gADOCon.Execute strDropTrigger, , adCmdText + adExecuteNoRecords
  
  ' Drop existing trigger (Update)
  strDropTrigger = "IF EXISTS" & _
    " (SELECT Name" & _
    "   FROM sysobjects" & _
    "   WHERE id = object_id('[UPD_ASRSysAccordTransactions]')" & _
    "     AND objectproperty(id, N'IsTrigger') = 1)" & _
    " DROP TRIGGER [UPD_ASRSysAccordTransactions]"
  gADOCon.Execute strDropTrigger, , adCmdText + adExecuteNoRecords
  
  ' Drop existing trigger (Delete)
  strDropTrigger = "IF EXISTS" & _
    " (SELECT Name" & _
    "   FROM sysobjects" & _
    "   WHERE id = object_id('[DEL_ASRSysAccordTransactions]')" & _
    "     AND objectproperty(id, N'IsTrigger') = 1)" & _
    " DROP TRIGGER [DEL_ASRSysAccordTransactions]"
  gADOCon.Execute strDropTrigger, , adCmdText + adExecuteNoRecords
  
  ' New triggers
  strUpdateTrigger = vbNullString
  strDeleteTrigger = vbNullString
  
  sSQL = "SELECT c.TableID, MIN(c.ColumnID) AS ColumnID, t.TransferTypeID, t.ASRBaseTableID FROM tmpColumns c" _
    & " INNER JOIN tmpAccordTransferTypes t ON t.ASRBaseTableID = c.TableID" _
    & " WHERE ColumnName NOT IN ('Timestamp', 'ID')" _
    & " GROUP BY c.TableID, t.TransferTypeID, t.ASRBaseTableID"
  Set rsAccordDetails = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  
  With rsAccordDetails
    Do While Not .EOF
      strUpdateTrigger = strUpdateTrigger _
        & Space$(8) & "IF @iTransferType = " & Trim(Str(!TransferTypeID)) & " UPDATE " _
        & GetTableName(!ASRBaseTableID) & " SET " _
        & GetColumnName(!ColumnID, False) & " =  " & GetColumnName(!ColumnID, False) & " WHERE ID = @recordID" & vbNewLine
      .MoveNext
    Loop
    .Close
  End With
  
  If LenB(strUpdateTrigger) <> 0 Then
    strUpdateTrigger = "/* ---------------------------------------------------------------- */" & vbNewLine _
        & "/* HR Pro Payroll module trigger.                   */" & vbNewLine _
        & "/* Automatically generated by the System manager.   */" & vbNewLine _
        & "/* ---------------------------------------------------------------- */" & vbNewLine _
        & "CREATE TRIGGER UPD_ASRSysAccordTransactions ON dbo.ASRSysAccordTransactions" & vbNewLine _
        & "FOR UPDATE" & vbNewLine _
        & "AS" & vbNewLine _
        & "BEGIN" & vbNewLine _
        & Space$(4) & "DECLARE @recordID int," & vbNewLine _
        & Space$(8) & "@iTransferType int," & vbNewLine _
        & Space$(8) & "@iStatus int," & vbNewLine _
        & Space$(8) & "@cursInsertedRecords cursor" & vbNewLine _
        & Space$(4) & "SET @cursInsertedRecords = CURSOR LOCAL FAST_FORWARD FOR SELECT inserted.HRProRecordID, inserted.TransferType, inserted.Status FROM inserted" & vbNewLine _
        & Space$(8) & "LEFT OUTER JOIN deleted ON inserted.TransactionID = deleted.TransactionID" & vbNewLine _
        & Space$(8) & "WHERE Deleted.Status <> inserted.Status AND inserted.Status IN (22,23)" & vbNewLine _
        & Space$(4) & "OPEN @cursInsertedRecords" & vbNewLine _
        & Space$(4) & "FETCH NEXT FROM @cursInsertedRecords INTO @recordID, @iTransferType, @iStatus" & vbNewLine _
        & Space$(4) & "WHILE (@@fetch_status = 0)" & vbNewLine _
        & Space$(4) & "BEGIN" & vbNewLine & strUpdateTrigger & vbNewLine _
        & Space$(8) & "FETCH NEXT FROM @cursInsertedRecords INTO @recordID, @iTransferType, @iStatus" & vbNewLine _
        & Space$(4) & "END" & vbNewLine _
        & "CLOSE @cursInsertedRecords" & vbNewLine _
        & "DEALLOCATE @cursInsertedRecords" & vbNewLine & vbNewLine _
        & "IF @@nestLevel = 1 EXECUTE spASRAccordPurgeTemp 1,0" & vbNewLine _
        & "END"
    gADOCon.Execute strUpdateTrigger, , adCmdText + adExecuteNoRecords
  End If
  
  
  strDeleteTrigger = "/* ---------------------------------------------------------------- */" & vbNewLine _
            & "/* HR Pro Payroll module trigger.                   */" & vbNewLine _
            & "/* Automatically generated by the System manager.   */" & vbNewLine _
            & "/* ---------------------------------------------------------------- */" & vbNewLine _
            & "CREATE TRIGGER DEL_ASRSysAccordTransactions ON dbo.ASRSysAccordTransactions" & vbNewLine _
            & "FOR DELETE" & vbNewLine _
            & "AS" & vbNewLine _
            & "BEGIN" & vbNewLine _
             & "   DELETE FROM ASRSysAccordTransactionData WHERE ASRSysAccordTransactionData.TransactionID IN (SELECT TransactionID FROM deleted)" & vbNewLine _
             & "   DELETE FROM ASRSysAccordTransactionWarnings WHERE ASRSysAccordTransactionWarnings.TransactionID IN (SELECT TransactionID FROM deleted)" & vbNewLine _
            & "END"
  gADOCon.Execute strDeleteTrigger, , adCmdText + adExecuteNoRecords

TidyUpAndExit:
'  Set rsAccordDetails = Nothing
  CreateAccordTransferTriggers = bOK
  Exit Function

ErrorTrap:
  OutputError "Error creating Payroll transfer triggers"
  bOK = False
  Resume TidyUpAndExit

End Function


Public Function CreateAccordTransferSPs(pfRefreshDatabase As Boolean) As Boolean
  
  ' Create the Payroll record export stored procedures.
  On Error GoTo ErrorTrap
  
  Dim bOK As Boolean
  Dim sSQL As String

  bOK = True
'  sSQL = "IF EXISTS " & _
'         "(SELECT Name FROM sysobjects " & _
'         "WHERE id = object_id('spASRAccordRunManualExport') " & _
'         "AND sysstat & 0xf = 4) " & _
'         "DROP PROCEDURE spASRAccordRunManualExport"
'  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  DropProcedure "spASRAccordRunManualExport"

TidyUpAndExit:
  CreateAccordTransferSPs = bOK
  Exit Function

ErrorTrap:
  OutputError "Error creating Payroll stored procedures"
  bOK = False
  Resume TidyUpAndExit

End Function

