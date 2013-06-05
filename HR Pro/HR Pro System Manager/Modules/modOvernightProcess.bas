Attribute VB_Name = "modOvernightProcess"
Option Explicit

Public glngOvernightJobTime As Long

Public Function CreateOvernightProcess(palngExpressions As Variant, pfRefreshDatabase As Boolean) As Boolean

  'Step 1 - Set flags
  'Step 2 - Update column calcs
  'Step 3 - Unset flags
  'Step 4 - Email/diary processing

  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim strScriptPath As String
  Dim strFileName As String
  
  If Application.ChangedOvernightJobSchedule Then
    'Make sure that we run the script for the overnight job Schedule
    strFileName = "HRProOvernightJob.sql"
    strScriptPath = App.Path & "\Update Scripts\" & strFileName
    
    RunOvernightScript strScriptPath
  End If
  
  fOK = OvernightJob2(palngExpressions)   'Always refresh Step 2 !!!

  If pfRefreshDatabase Then
    If fOK Then fOK = OvernightJob1       'Step 1
    If fOK Then fOK = OvernightJob3       'Step 3
    If fOK Then fOK = OvernightJob4       'Step 4
  End If

  ' Reindex and update stats job
  If fOK Then fOK = OvernightJob5         'Always refresh Step 5 !!!

TidyAndExit:
  If fOK = False Then
    OutputError "Error creating overnight process"
  End If
  CreateOvernightProcess = fOK
Exit Function


ErrorTrap:
  fOK = False
  Resume TidyAndExit

End Function

Private Function RunOvernightScript(pstrFileName As String) As Boolean

  Dim sUpdateScript As String
  Dim sReadString As String

  sUpdateScript = vbNullString

  Open pstrFileName For Input As #1
  Do While Not EOF(1)
    Line Input #1, sReadString
    sUpdateScript = sUpdateScript & sReadString & vbNewLine
  Loop
  Close #1

  If sUpdateScript <> vbNullString Then
    'Set rdoTempCon = rdoEnv.OpenConnection("", rdDriverNoPrompt, False, sConnect)
    gADOCon.Execute sUpdateScript, , adExecuteNoRecords
  End If

End Function

Private Function OvernightJob1() As Boolean

  Const strOvernightSP As String = "spASRSysOvernightStep1"
  Dim strSQL As String
  
  DropExistingJobStep (strOvernightSP)

  'strSQL = "CREATE PROCEDURE dbo." & strOvernightSP & " AS" & vbNewline & _
           "BEGIN" & vbNewline & _
           "/* This sets all of the flags prior to updating date dependant columns */" & vbNewline & _
           vbNewline & _
           "update ASRSysConfig set updatingDateDependentColumns = 1" & vbNewline & _
           vbNewline & _
           "exec sp_configure 'nested triggers', '0'" & vbNewline & _
           "reconfigure" & vbNewline & _
           "exec sp_dboption '" & Replace(Database.DatabaseName, "'", "''") & "', 'recursive triggers', 'FALSE'" & vbNewline & _
           vbNewline & _
           "END"

  'MH20020807 Fault 4209
  'Added "With Override" option and get current process database name

  'strSQL = "CREATE PROCEDURE dbo." & strOvernightSP & " AS" & vbNewLine & _
    "BEGIN" & vbNewLine & _
    "    DECLARE @sDBName nvarchar(MAX)" & vbNewLine & vbNewLine & _
    "    SELECT @sDBName = master..sysdatabases.name" & vbNewLine & _
    "    FROM master..sysdatabases" & vbNewLine & _
    "    INNER JOIN master..sysprocesses ON master..sysdatabases.dbid = master..sysprocesses.dbid" & vbNewLine & _
    "    WHERE master..sysprocesses.spid = @@spid" & vbNewLine & vbNewLine & _
    "    /* This sets all of the flags prior to updating date dependant columns */" & vbNewLine & _
    "    DELETE FROM ASRSYSSystemSettings WHERE [Section] = 'database' and [SettingKey] = 'updatingdatedependantcolumns'" & vbNewLine & vbNewLine & _
    "    INSERT ASRSYSSystemSettings([Section],[SettingKey],[SettingValue])" & vbNewLine & _
    "    VALUES('database','updatingdatedependantcolumns',1)" & vbNewLine & vbNewLine & _
    "    exec sp_configure 'nested triggers', '0'" & vbNewLine & _
    "    RECONFIGURE WITH OVERRIDE" & vbNewLine & vbNewLine & _
    "    exec sp_dboption @sDBName, 'recursive triggers', 'FALSE'" & vbNewLine & _
    "END"
    
'  strSQL = "CREATE PROCEDURE dbo." & strOvernightSP & " AS" & vbNewLine & _
'    "BEGIN" & vbNewLine & _
'    "    DECLARE @sDBName nvarchar(MAX)" & vbNewLine & vbNewLine & _
'    "    SELECT @sDBName = db_name()" & vbNewLine & vbNewLine & _
'    "    /* This sets all of the flags prior to updating date dependant columns */" & vbNewLine & _
'    "    DELETE FROM ASRSYSSystemSettings WHERE [Section] = 'database' and [SettingKey] = 'updatingdatedependantcolumns'" & vbNewLine & vbNewLine & _
'    "    INSERT ASRSYSSystemSettings([Section],[SettingKey],[SettingValue])" & vbNewLine & _
'    "    VALUES('database','updatingdatedependantcolumns',1)" & vbNewLine & vbNewLine & _
'    "    exec sp_configure 'nested triggers', '0'" & vbNewLine & _
'    "    RECONFIGURE WITH OVERRIDE" & vbNewLine & vbNewLine & _
'    "    exec sp_dboption @sDBName, 'recursive triggers', 'FALSE'" & vbNewLine & _
'    "END"
    
    'AE20080214 Fault #12726 - Nested Trigger check is now in the UPD_ triggers
  strSQL = "/* ------------------------------------------------------------------------------- */" & vbNewLine & _
    "/* HR Pro system stored procedure.                  */" & vbNewLine & _
    "/* Automatically generated by the System Manager.   */" & vbNewLine & _
    "/* ------------------------------------------------------------------------------- */" & vbNewLine & _
    "CREATE PROCEDURE dbo." & strOvernightSP & " AS" & vbNewLine & _
    "BEGIN" & vbNewLine & _
    "    DECLARE @sDBName nvarchar(MAX);" & vbNewLine & vbNewLine & _
    "    SELECT @sDBName = db_name()" & vbNewLine & vbNewLine & _
    "    /* This sets all of the flags prior to updating date dependant columns */" & vbNewLine & _
    "    DELETE FROM ASRSYSSystemSettings WHERE [Section] = 'database' and [SettingKey] = 'updatingdatedependantcolumns'" & vbNewLine & vbNewLine & _
    "    INSERT ASRSYSSystemSettings([Section],[SettingKey],[SettingValue])" & vbNewLine & _
    "    VALUES('database','updatingdatedependantcolumns',1)" & vbNewLine & vbNewLine & _
    "    exec sp_dboption @sDBName, 'recursive triggers', 'FALSE'" & vbNewLine & _
    "END"
    
  gADOCon.Execute strSQL, , adExecuteNoRecords

  OvernightJob1 = True

End Function

Private Function OvernightJob2(palngExpressions As Variant) As Boolean
  ' Create the scheduled job for refreshing date dependent columns.
  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim lngLastTable As Long
  Dim sSQL As String
  Dim sExprList As String
  Dim sJobName As String
  Dim rsInfo As New ADODB.Recordset
  Dim bArchive As Boolean
  Dim iArchivePeriod As Integer
  Dim iArchivePeriodType As Integer

  Const strOvernightSP As String = "spASRSysOvernightStep2"
  
  On Error GoTo ErrorTrap

  fOK = True
  sExprList = ""

  DropExistingJobStep (strOvernightSP)

  ' Get the expressions that use time dependent functions.
  sSQL = "SELECT ASRSysExpressions.exprID, ASRSysExpressions.parentComponentID" & _
    " FROM ASRSysExpressions" & _
    " INNER JOIN ASRSysExprComponents" & _
    "   ON ASRSysExpressions.exprID = ASRSysExprComponents.exprID" & _
    " INNER JOIN ASRSysFunctions" & _
    "   ON ASRSysExprComponents.functionID = ASRSysFunctions.functionID" & _
    " WHERE ASRSysFunctions.timeDependent = 1"

  'JPD 20051121 If columns that use special functions aren't being updated automatically
  ' make sure they are as part of the overnight job.
  ' Special function are Absence Duration, Absence Between Two Dates, and Working Days Between Two Dates
  If gbDisableSpecialFunctionAutoUpdate Then
    sSQL = sSQL & _
      "   OR ASRSysFunctions.functionID IN (30, 46, 47, 73)"
  End If
  
  rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    
  With rsInfo
    Do While Not .EOF
      iIndex = UBound(palngExpressions, 2) + 1
      ReDim Preserve palngExpressions(2, iIndex)
      palngExpressions(1, iIndex) = !ExprID
      palngExpressions(2, iIndex) = !ParentComponentID
      .MoveNext
    Loop
  End With
  rsInfo.Close
  
  iLoop = 1
  Do While iLoop <= UBound(palngExpressions, 2)
    If palngExpressions(2, iLoop) > 0 Then
      ' The expression is a sub-expression. Get the parent expression info.
      sSQL = "SELECT ASRSysExpressions.exprID, ASRSysExpressions.parentComponentID" & _
        " FROM ASRSysExpressions" & _
        " INNER JOIN ASRSysExprComponents" & _
        "   ON ASRSysExpressions.exprID = ASRSysExprComponents.exprID" & _
        " WHERE ASRSysExprComponents.componentID = " & Trim(Str(palngExpressions(2, iLoop)))
        
      rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
      With rsInfo
        If Not (.EOF And .BOF) Then
          iIndex = UBound(palngExpressions, 2) + 1
          ReDim Preserve palngExpressions(2, iIndex)
          palngExpressions(1, iIndex) = !ExprID
          palngExpressions(2, iIndex) = !ParentComponentID
        End If
      End With
      rsInfo.Close

    Else
      ' The expression is a top-level expression. Save its id and find any other
      ' expressions that use it.
      sExprList = sExprList & IIf(Len(sExprList) > 0, ",", "") & Trim(Str(palngExpressions(1, iLoop)))
      
      'JPD 20040220 Fault 8111
      sSQL = "SELECT ASRSysExpressions.exprID, ASRSysExpressions.parentComponentID" & _
        " FROM ASRSysExpressions" & _
        " INNER JOIN ASRSysExprComponents" & _
        "   ON ASRSysExpressions.exprID = ASRSysExprComponents.exprID" & _
        " WHERE (ASRSysExprComponents.calculationID = " & Trim(Str(palngExpressions(1, iLoop))) & ")" & _
        "   OR (ASRSysExprComponents.filterID = " & Trim(Str(palngExpressions(1, iLoop))) & ")" & _
        "   OR ((ASRSysExprComponents.fieldSelectionFilter = " & Trim(Str(palngExpressions(1, iLoop))) & ")" & _
        "     AND (ASRSysExprComponents.type = " & CStr(giCOMPONENT_FIELD) & "))"

      rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
      With rsInfo
        Do While Not .EOF
          iIndex = UBound(palngExpressions, 2) + 1
          ReDim Preserve palngExpressions(2, iIndex)
          palngExpressions(1, iIndex) = !ExprID
          palngExpressions(2, iIndex) = !ParentComponentID
          .MoveNext
        Loop
      End With
      rsInfo.Close
    End If
    
    iLoop = iLoop + 1
  Loop
  
  If Len(sExprList) > 0 Then
    ' We now have a list of all time-dependent expressions.
    ' Get the list of time dependent columns.
    lngLastTable = 0
    
    sSQL = "SELECT ASRSysColumns.columnName, ASRSysColumns.tableID, ASRSysTables.tableName" & _
      " FROM ASRSysColumns" & _
      " INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID" & _
      " WHERE ASRSysColumns.columnType = " & Trim(Str(giCOLUMNTYPE_CALCULATED)) & _
      " AND ASRSysColumns.calcExprID IN (" & sExprList & ")" & _
      " ORDER BY ASRSysColumns.tableID"
  
    rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    ' AE20080408 Fault #13072
    ' sSQL = "DELETE FROM ASRSysOvernightProgress" & vbNewLine & vbNewLine
    
    sSQL = "-- Create the progress table if it doesn't already exist" & vbNewLine & _
           "IF OBJECT_ID('ASRSysOvernightProgress', N'U') IS NOT NULL" & vbNewLine & _
           "    DELETE FROM ASRSysOvernightProgress" & vbNewLine & vbNewLine
    With rsInfo
      Do While Not .EOF
        If lngLastTable <> !TableID Then
          lngLastTable = !TableID
          
          ' AE20080213 Fault #12726 - changed to do in batches of 2000 IDs
          ' adding checkpoint to commit after each batch
          '
          ' JPD20020913 - reverted to the old method of updating all records
          ' in a table 'en masse', rather then record by record.
          ' This change was driven by the performance degradation reported by
          ' Islington.
'          sSQL = sSQL & _
'            " UPDATE " & Replace(!TableName, "'", "''") & _
'            " SET " & Replace(!ColumnName, "'", "''") & _
'            " = " & Replace(!ColumnName, "'", "''") & vbNewLine

          sSQL = sSQL & _
            "EXEC [spASRSysOvernightTableUpdate] " & _
              "'" & Replace(!TableName, "'", "''") & "', '" & Replace(!ColumnName, "'", "''") & "', 100" & vbNewLine

          'sSQL = sSQL & vbNewline & _
            "  --" & Replace(!TableName, "'", "''") & vbNewline & _
            "  DECLARE HRProCursor CURSOR" & vbNewline & _
            "  FOR select ID from " & Replace(!TableName, "'", "''") & vbNewline & vbNewline & _
            "  OPEN HRProCursor" & vbNewline & _
            "  FETCH NEXT FROM HRProCursor INTO @ID" & vbNewline & _
            "  WHILE @@FETCH_STATUS = 0" & vbNewline & _
            "  BEGIN" & vbNewline & _
            "    UPDATE " & Replace(!TableName, "'", "''") & _
                 " SET " & Replace(!ColumnName, "'", "''") & _
                 " = " & Replace(!ColumnName, "'", "''") & " WHERE ID = @ID" & vbNewline & _
            "    FETCH NEXT FROM HRProCursor INTO @ID" & vbNewline & _
            "  END" & vbNewline & vbNewline & _
            "  CLOSE HRProCursor" & vbNewline & _
            "  DEALLOCATE HRProCursor" & vbNewline & vbNewline
        End If
        
        .MoveNext
      Loop
    
    End With
    rsInfo.Close

    ' Payroll archiving functions
    If IsModuleEnabled(modAccordPayroll) Then

      ' Get the archive options
      With recModuleSetup
        .Seek "=", gsMODULEKEY_ACCORD, gsPARAMETERKEY_PURGEOPTION
        If .NoMatch Then
          .Seek "=", gsMODULEKEY_ACCORD, gsPARAMETERKEY_PURGEOPTION
          If .NoMatch Then
            bArchive = False
          Else
            bArchive = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
          End If
        Else
          bArchive = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
        End If
      
        ' Get the purge period.
        .Seek "=", gsMODULEKEY_ACCORD, gsPARAMETERKEY_PURGEOPTIONPERIOD
        If .NoMatch Then
          .Seek "=", gsMODULEKEY_ACCORD, gsPARAMETERKEY_PURGEOPTIONPERIOD
          If .NoMatch Then
            iArchivePeriod = 0
          Else
            iArchivePeriod = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
          End If
        Else
          iArchivePeriod = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
        End If

        ' Get the purge type.
        .Seek "=", gsMODULEKEY_ACCORD, gsPARAMETERKEY_PURGEOPTIONPERIODTYPE
        If .NoMatch Then
          .Seek "=", gsMODULEKEY_ACCORD, gsPARAMETERKEY_PURGEOPTIONPERIODTYPE
          If .NoMatch Then
            iArchivePeriodType = 0
          Else
            iArchivePeriodType = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
          End If
        Else
          iArchivePeriodType = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
        End If

      End With
 
      ' Do we need to add the purge trigger code?
      If bArchive Then
      
        sSQL = sSQL & vbNewLine & "-- Payroll archiving" & vbNewLine
      
        Select Case iArchivePeriodType
          ' Days
          Case 0
            sSQL = sSQL & "UPDATE ASRSysAccordTransactions SET Archived = 1 " _
              & "WHERE [CreatedDateTime] < DATEADD(dd,-" & Trim$(Str$(iArchivePeriod)) & ",getdate())" & vbNewLine
          
          ' Weeks
          Case 1
            sSQL = sSQL & "UPDATE ASRSysAccordTransactions SET Archived = 1 " _
              & "WHERE [CreatedDateTime] < DATEADD(wk,-" & Trim$(Str$(iArchivePeriod)) & ",getdate())" & vbNewLine
          
          'Months
          Case 2
            sSQL = sSQL & "UPDATE ASRSysAccordTransactions SET Archived = 1 " _
              & "WHERE [CreatedDateTime] < DATEADD(mm,-" & Trim$(Str$(iArchivePeriod)) & ",getdate())" & vbNewLine
          
          ' Years
          Case 3
            sSQL = sSQL & "UPDATE ASRSysAccordTransactions SET Archived = 1 " & vbNewLine _
              & "WHERE [CreatedDateTime] < DATEADD(yy,-" & Trim$(Str$(iArchivePeriod)) & ",getdate())" & vbNewLine
        
        End Select
          
      End If
    End If


    'JPD20011005 Fault 2899
    ' Put dummy code in the stored procedure if there is no generated code to go in it.
    ' SQL doesn't like having no code with a BEGIN-END block.
    If Len(sSQL) = 0 Then
      sSQL = "  DECLARE @iDummy integer" & vbNewLine & _
        "  SET @iDummy = 1"
    End If
    
    sSQL = "/* ------------------------------------------------------------------------------- */" & vbNewLine & _
      "/* HR Pro system stored procedure.                  */" & vbNewLine & _
      "/* Automatically generated by the System Manager.   */" & vbNewLine & _
      "/* ------------------------------------------------------------------------------- */" & vbNewLine & _
      "CREATE PROCEDURE [dbo].[" & strOvernightSP & "] AS" & vbNewLine & _
      "BEGIN" & vbNewLine & vbNewLine & _
      sSQL & vbNewLine & _
      "END"
    gADOCon.Execute sSQL, , adExecuteNoRecords

  End If


TidyUpAndExit:
  Set rsInfo = Nothing
  OvernightJob2 = fOK
  Exit Function

ErrorTrap:
  fOK = False
  'gobjProgress.Visible = False
  'MsgBox ODBC.FormatError(Err.Description), vbOKOnly + vbExclamation, Application.Name
  OutputError "Error creating calculated columns update job"
  Resume TidyUpAndExit

End Function


Private Function OvernightJob3() As Boolean

  Const strOvernightSP As String = "spASRSysOvernightStep3"
  Dim strSQL As String
  
  DropExistingJobStep (strOvernightSP)

  'strSQL = "CREATE PROCEDURE dbo." & strOvernightSP & " AS" & vbNewline & _
           "BEGIN" & vbNewline & _
           vbNewline & _
           "/* This unsets all of the flags after updating date dependant columns */" & vbNewline & _
           vbNewline & _
           "update ASRSysConfig set updatingDateDependentColumns = 0" & vbNewline & _
           vbNewline & _
           "exec sp_dboption '" & Replace(Database.DatabaseName, "'", "''") & "', 'recursive triggers', 'TRUE'" & vbNewline & _
           vbNewline & _
           "exec sp_configure 'nested triggers', '1'" & vbNewline & _
           "reconfigure" & vbNewline & _
           vbNewline & _
           "END"
  
  'MH20020807 Fault 4209
  'Added "With Override" option and get current process database name

  'strSQL = "    CREATE PROCEDURE dbo." & strOvernightSP & " AS" & vbNewLine & _
    "    BEGIN" & vbNewLine & _
    "    DECLARE @sDBName nvarchar(MAX)" & vbNewLine & vbNewLine & _
    "    SELECT @sDBName = master..sysdatabases.name" & vbNewLine & _
    "    FROM master..sysdatabases" & vbNewLine & _
    "    INNER JOIN master..sysprocesses ON master..sysdatabases.dbid = master..sysprocesses.dbid" & vbNewLine & _
    "    WHERE master..sysprocesses.spid = @@spid" & vbNewLine & vbNewLine & _
    "    /* This unsets all of the flags after updating date dependant columns */" & vbNewLine & _
    "    DELETE FROM ASRSYSSystemSettings WHERE [Section] = 'database' and [SettingKey] = 'updatingdatedependantcolumns'" & vbNewLine & vbNewLine & _
    "    INSERT ASRSYSSystemSettings([Section],[SettingKey],[SettingValue])" & vbNewLine & _
    "    VALUES('database','updatingdatedependantcolumns',0)" & vbNewLine & vbNewLine & _
    "    exec sp_dboption @sDBName, 'recursive triggers', 'TRUE'" & vbNewLine & vbNewLine & _
    "    exec sp_configure 'nested triggers', '1'" & vbNewLine & _
    "    RECONFIGURE WITH OVERRIDE" & vbNewLine & _
    "END"
  
'  strSQL = "    CREATE PROCEDURE dbo." & strOvernightSP & " AS" & vbNewLine & _
'    "    BEGIN" & vbNewLine & _
'    "    DECLARE @sDBName nvarchar(MAX)" & vbNewLine & vbNewLine & _
'    "    SELECT @sDBName = db_name()" & vbNewLine & vbNewLine & _
'    "    /* This unsets all of the flags after updating date dependant columns */" & vbNewLine & _
'    "    DELETE FROM ASRSYSSystemSettings WHERE [Section] = 'database' and [SettingKey] = 'updatingdatedependantcolumns'" & vbNewLine & vbNewLine & _
'    "    INSERT ASRSYSSystemSettings([Section],[SettingKey],[SettingValue])" & vbNewLine & _
'    "    VALUES('database','updatingdatedependantcolumns',0)" & vbNewLine & vbNewLine & _
'    "    exec sp_dboption @sDBName, 'recursive triggers', 'TRUE'" & vbNewLine & vbNewLine & _
'    "    exec sp_configure 'nested triggers', '1'" & vbNewLine & _
'    "    RECONFIGURE WITH OVERRIDE" & vbNewLine & _
'    "END"
  
    'AE20080214 Fault #12726 - Nested Trigger check is now in the UPD_ triggers
  strSQL = "/* ------------------------------------------------------------------------------- */" & vbNewLine & _
    "/* HR Pro system stored procedure.                  */" & vbNewLine & _
    "/* Automatically generated by the System Manager.   */" & vbNewLine & _
    "/* ------------------------------------------------------------------------------- */" & vbNewLine & _
    "CREATE PROCEDURE [dbo].[" & strOvernightSP & "] AS" & vbNewLine & _
    "BEGIN" & vbNewLine & _
    "    DECLARE @sDBName nvarchar(MAX)" & vbNewLine & vbNewLine & _
    "    SELECT @sDBName = db_name()" & vbNewLine & vbNewLine & _
    "    /* This unsets all of the flags after updating date dependant columns */" & vbNewLine & _
    "    DELETE FROM ASRSYSSystemSettings WHERE [Section] = 'database' and [SettingKey] = 'updatingdatedependantcolumns'" & vbNewLine & vbNewLine & _
    "    INSERT ASRSYSSystemSettings([Section],[SettingKey],[SettingValue])" & vbNewLine & _
    "    VALUES('database','updatingdatedependantcolumns',0)" & vbNewLine & vbNewLine & _
    "    EXEC sp_dboption @sDBName, 'recursive triggers', 'FALSE'" & vbNewLine & vbNewLine & _
    "END"
    
  gADOCon.Execute strSQL, , adExecuteNoRecords

  OvernightJob3 = True

End Function


Private Function OvernightJob4() As Boolean

  Const strOvernightSP As String = "spASRSysOvernightStep4"
  Dim strSQL As String
  
  DropExistingJobStep (strOvernightSP)

  strSQL = "/* ------------------------------------------------------------------------------- */" & vbNewLine & _
    "/* HR Pro system stored procedure.                  */" & vbNewLine & _
    "/* Automatically generated by the System Manager.   */" & vbNewLine & _
    "/* ------------------------------------------------------------------------------- */" & vbNewLine & _
  "CREATE PROCEDURE [dbo].[" & strOvernightSP & "] AS" & vbNewLine & _
    "BEGIN" & vbNewLine & _
    "    /* Overnight Email Processing */" & vbNewLine & _
    "    EXEC sp_ASRPurgeRecords 'EMAIL', 'ASRSysEmailQueue', 'DateDue'" & vbNewLine & _
    "    EXEC spASREmailRebuild" & vbNewLine & _
    "    EXEC spASREmailBatch" & vbNewLine & vbNewLine & _
    "    /* Overnight Diary Processing */" & vbNewLine & _
    "    EXEC sp_ASRDiaryPurge" & vbNewLine & vbNewLine & _
    IIf(Application.WorkflowModule, _
    "    /* Overnight Workflow Processing */" & vbNewLine & _
    "    EXEC spASRWorkflowRebuild" & vbNewLine & vbNewLine, vbNullString) & _
    "    /* Overnight Log Purging */" & vbNewLine & _
    "    EXEC sp_ASRAuditLogPurge" & vbNewLine & _
    "    EXEC sp_AsrEventLogPurge" & vbNewLine & vbNewLine & _
    "    /* Update Overnight Job Log */" & vbNewLine & _
    "    DELETE from ASRSysSystemSettings" & vbNewLine & _
    "    WHERE [Section] = 'overnight' and [SettingKey] = 'last completed'" & vbNewLine & _
    "    INSERT ASRSysSystemSettings([Section], [SettingKey], [SettingValue])" & vbNewLine & _
    "    VALUES('overnight', 'last completed', convert(varchar,getdate(),103)+' '+convert(varchar,getdate(),108))" & vbNewLine & _
    "END"

'  AE20071219 Fault #12731
'  If Application.WorkflowModule Then
'    strSQL = strSQL & _
'      "    EXEC spASRWorkflowRebuild" & vbNewLine
'  End If
  
'  strSQL = strSQL & _
'    "END"
  
  gADOCon.Execute strSQL, , adExecuteNoRecords

  OvernightJob4 = True

End Function

Private Function DropExistingJobStep(strOvernightSP As String) As Boolean

  Dim strSQL As String

  On Error GoTo ErrorTrap

  strSQL = "IF EXISTS" & _
    " (SELECT Name" & _
    "   FROM sysobjects" & _
    "   WHERE id = object_id('" & strOvernightSP & "')" & _
    "     AND sysstat & 0xf = 4)" & _
    " DROP PROCEDURE " & strOvernightSP
    
  gADOCon.Execute strSQL, , adExecuteNoRecords

Exit Function

ErrorTrap:
  OutputError "Error Removing overnight process"

End Function

' Automatic reindex job
Private Function OvernightJob5() As Boolean

  Const strOvernightSP As String = "spASRSysOvernightStep5"
  Dim strSQL As String
  
  DropExistingJobStep (strOvernightSP)

  strSQL = "/* ------------------------------------------------------------------------------- */" & vbNewLine & _
    "/* HR Pro system stored procedure.                  */" & vbNewLine & _
    "/* Automatically generated by the System Manager.   */" & vbNewLine & _
    "/* ------------------------------------------------------------------------------- */" & vbNewLine & _
    "CREATE PROCEDURE [dbo].[" & strOvernightSP & "] AS" & vbNewLine & _
    "BEGIN" & vbNewLine & _
    "    DECLARE @iCount int;" & vbNewLine & vbNewLine
    
'  If gbReorganiseIndexesInOvernightJob Then
  strSQL = strSQL & "    EXEC [dbo].[spASRGetCurrentUsers]" & vbNewLine & _
    "    IF @@ROWCOUNT = 0" & vbNewLine & _
    "    BEGIN" & vbNewLine & _
    "        -- Tidy up temporary objects" & vbNewLine & _
    "        EXEC sp_executeSQL spASRDropTempObjects;" & vbNewLine & vbNewLine & _
    "        -- Defragment indexes" & vbNewLine & _
    "        EXEC dbo.spASRDefragIndexes 100;" & vbNewLine & vbNewLine & _
    "        -- Update statistics" & vbNewLine & _
    "        EXEC sp_executeSQL spASRUpdateStatistics;" & vbNewLine & vbNewLine & _
    "        -- Optimise the record save for single record" & vbNewLine & _
    "        EXEC sp_executeSQL spadmin_optimiserecordsave;" & vbNewLine & _
    "    END" & vbNewLine
 ' End If
  
  strSQL = strSQL & "END"
  
  gADOCon.Execute strSQL, , adExecuteNoRecords
  
  OvernightJob5 = True

End Function


