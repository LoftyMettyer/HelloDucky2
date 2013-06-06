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
  sSQL = ""

  DropExistingJobStep (strOvernightSP)

  ' Get tables to update
  sSQL = "SELECT TableName FROM ASRSysTables"
  rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    
  sSQL = ""
  With rsInfo
    Do While Not .EOF
      
      sSQL = sSQL & _
        "EXEC [spASRSysOvernightTableUpdate] " & _
          "'tbuser_" & Replace(!TableName, "'", "''") & "', 'updflag', 100" & vbNewLine
      
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


